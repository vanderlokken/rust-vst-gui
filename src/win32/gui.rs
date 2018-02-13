use std::os::raw::c_void;
use std::ptr::{null, null_mut};
use std::rc::Rc;

use winapi::Interface;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::shared::winerror::*;
use winapi::shared::wtypes::*;
use winapi::um::combaseapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::oaidl::*;
use winapi::um::objidlbase::*;
use winapi::um::oleauto::*;
use winapi::um::winuser::*;

use lib::{JavascriptCallback, PluginGui};
use win32::client_site::*;
use win32::com_pointer::*;
use win32::ffi::*;
use win32::error::*;

struct Window {
    handle: HWND,
}

impl Window {
    // The "plugin_window" string in utf16.
    const CLASS_NAME: [u16; 14] = [
        0x0070, 0x006c, 0x0075, 0x0067, 0x0069, 0x006e, 0x005f, 0x0077,
        0x0069, 0x006e, 0x0064, 0x006f, 0x0077, 0x0000];

    pub fn new(parent: HWND) -> Window {
        Window::register_window_class();

        let handle = unsafe {
            const STYLE: DWORD = WS_CHILD | WS_VISIBLE;
            const STYLE_EXTENDED: DWORD = 0;

            CreateWindowExW(
                STYLE_EXTENDED,
                Window::CLASS_NAME.as_ptr(),
                null(), /*window_name*/
                STYLE,
                0, /*x*/
                0, /*y*/
                Window::default_size().0,
                Window::default_size().1,
                parent,
                null_mut(), /*menu*/
                GetModuleHandleW(null()),
                null_mut())
        };

        Window {
            handle: handle,
        }
    }

    fn size(&self) -> (i32, i32) {
        let mut rectangle =
            RECT {left: 0, top: 0, right: 0, bottom: 0};

        unsafe {
            GetWindowRect(self.handle, &mut rectangle);
        }

        let width = (rectangle.right - rectangle.left) as i32;
        let height = (rectangle.bottom - rectangle.top) as i32;

        (width, height)
    }

    fn default_size() -> (i32, i32) {
        unsafe {
            let width = GetSystemMetrics(SM_CXSCREEN) / 2;
            let height = GetSystemMetrics(SM_CYSCREEN) / 2;

            (width, height)
        }
    }

    fn register_window_class() {
        let class = WNDCLASSW {
            style: CS_DBLCLKS,
            lpfnWndProc: Some(Window::window_procedure),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: unsafe { GetModuleHandleW(null()) },
            hIcon: null_mut(),
            hCursor: unsafe { LoadCursorW(null_mut(), IDC_ARROW) },
            hbrBackground: null_mut(),
            lpszMenuName: null(),
            lpszClassName: Window::CLASS_NAME.as_ptr()
        };

        unsafe {
            RegisterClassW(&class);
        }
    }

    extern "system" fn window_procedure(
        handle: HWND, message: UINT, wparam: WPARAM, lparam: LPARAM) -> LRESULT
    {
        match message {
            WM_GETDLGCODE => {
                return DLGC_WANTALLKEYS;
            },
            _ => {}
        }
        unsafe {
            DefWindowProcW(handle, message, wparam, lparam)
        }
    }
}

struct WebBrowser {
    browser: ComPointer<IWebBrowser2>,
}

impl WebBrowser {
    fn new(
        window_handle: HWND,
        html_document: String,
        js_callback: Rc<JavascriptCallback>) -> Result<WebBrowser, RuntimeError>
    {
        unsafe {
            OleInitialize(null_mut());
        }

        let browser = WebBrowser {
            browser: WebBrowser::new_browser_com_object()?
        };

        browser.embed(window_handle, js_callback)?;

        // https://msdn.microsoft.com/library/aa752047
        browser.open_blank_page()?;
        browser.load_html_document(html_document)?;

        Ok(browser)
    }

    fn new_browser_com_object() ->
        Result<ComPointer<IWebBrowser2>, RuntimeError>
    {
        let mut web_browser = ComPointer::<IWebBrowser2>::new();

        let result = unsafe {
             CoCreateInstance(
                &CLSID_WebBrowser,
                null_mut(),
                CLSCTX_INPROC,
                &IWebBrowser2::uuidof(),
                web_browser.as_mut_ptr()
                    as *mut *mut IWebBrowser2
                    as *mut LPVOID)
        };

        if result == S_OK && web_browser.get().is_some() {
            Ok(web_browser)
        } else {
            Err(RuntimeError::new(
                "Couldn't get an instance of the 'IWebBrowser2' class"))
        }
    }

    fn browser(&self) -> &IWebBrowser2 {
        self.browser
            .get()
            .unwrap()
    }

    fn open_blank_page(&self) -> Result<(), RuntimeError> {
        use std::ffi::OsStr;
        use std::os::windows::ffi::OsStrExt;

        let url_buffer: Vec<u16> =
            OsStr::new("about:blank").encode_wide().collect();

        unsafe {
            let url = SysAllocStringLen(
                url_buffer.as_ptr(),
                url_buffer.len() as u32);

            if self.browser().Navigate(
                url,
                null_mut(),
                null_mut(),
                null_mut(),
                null_mut()) != S_OK
            {
                return Err(RuntimeError::new("Couldn't open blank page"))
            }

            SysFreeString(url);
        }

        Ok(())
    }

    fn load_html_document(
        &self, html_document: String) -> Result<(), RuntimeError>
    {
        let document_dispatch = {
            let mut result = ComPointer::<IDispatch>::new();

            // TODO: don't assume the document is ready
            let success = unsafe {
                self.browser().get_Document(result.as_mut_ptr()) == S_OK
            };
            if success {result} else {ComPointer::<IDispatch>::new()}
        };

        document_dispatch
            .get()
            .ok_or(RuntimeError::new(
                "Couldn't get an instance of the 'IDispatch' class"))?;

        let stream = ComPointer::<IStream>::from_raw(
            unsafe {
                SHCreateMemStream(
                    html_document.as_ptr(),
                    html_document.len() as u32)
            });

        stream
            .get()
            .ok_or(RuntimeError::new(
                "Couldn't get an instance of the 'IStream' class"))?;

        let success = document_dispatch
            .query_interface::<IPersistStreamInit>()
            .get()
            .map(|persist_stream| {
                unsafe {
                    persist_stream.InitNew() == S_OK &&
                    persist_stream.Load(stream.as_ptr()) == S_OK
                }
            })
            .ok_or(RuntimeError::new(
                "Couldn't get an instance of the 'IPersistStreamInit' class"))?;

        match success {
            true => Ok(()),
            false => Err(RuntimeError::new("Couldn't load an HTML document")),
        }
    }

    fn embed(
        &self,
        window_handle: HWND,
        js_callback: Rc<JavascriptCallback>) -> Result<(), RuntimeError>
    {
        let ole_object = self.browser.query_interface::<IOleObject>();

        ole_object
            .get()
            .ok_or(RuntimeError::new(
                "Couldn't get an instance of the 'IOleObject' class"))?;

        let ole_in_place_object =
            ole_object.query_interface::<IOleInPlaceObject>();

        ole_in_place_object
            .get()
            .ok_or(RuntimeError::new(
                "Couldn't get an instance of the 'IOleInPlaceObject' class"))?;

        let client_site = new_client_site(
            window_handle, ole_in_place_object, js_callback);

        let success = {
            let mut rectangle = RECT {left: 0, top: 0, right: 0, bottom: 0};

            unsafe {
                GetClientRect(window_handle, &mut rectangle);

                ole_object.get().unwrap().SetClientSite(
                    client_site.as_ptr()) == S_OK &&
                ole_object.get().unwrap().DoVerb(
                    OLEIVERB_INPLACEACTIVATE,
                    null_mut(),
                    client_site.as_ptr(),
                    0,
                    window_handle,
                    &rectangle) == S_OK &&
                self.browser().put_Width(rectangle.right) == S_OK &&
                self.browser().put_Height(rectangle.bottom) == S_OK &&
                self.browser().put_Visible(TRUE as VARIANT_BOOL) == S_OK
            }
        };

        match success {
            true => Ok(()),
            false => Err(RuntimeError::new("Couldn't reveal an HTML browser")),
        }
    }
}

struct Gui {
    html_document: String,
    js_callback: Rc<JavascriptCallback>,
    web_browser: Option<WebBrowser>,
    window: Option<Window>,
}

impl PluginGui for Gui {
    fn size(&self) -> (i32, i32) {
        match self.window {
            Some(ref window) => window.size(),
            None => (0, 0)
        }
    }

    fn position(&self) -> (i32, i32) {
        (0, 0)
    }

    fn close(&mut self) {
        self.web_browser = None;
        self.window = None;
    }

    fn open(&mut self, parent_handle: *mut c_void) {
        let window = Window::new(parent_handle as HWND);

        // TODO: display errors
        self.web_browser = WebBrowser::new(
            window.handle,
            self.html_document.clone(),
            self.js_callback.clone()
        ).ok();
        self.window = Some(window);
    }

    fn is_open(&mut self) -> bool {
        self.window.is_some()
    }
}

pub fn new_plugin_gui(
    html_document: String,
    js_callback: JavascriptCallback) -> Box<PluginGui>
{
    Box::new(
        Gui {
            html_document: html_document,
            js_callback: Rc::new(js_callback),
            web_browser: None,
            window: None,
        })
}
