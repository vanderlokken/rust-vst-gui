use std::error::Error;
use std::ffi::OsStr;
use std::mem::zeroed;
use std::os::raw::c_void;
use std::os::windows::ffi::OsStrExt;
use std::ptr::{null, null_mut};
use std::rc::Rc;

use winapi::Interface;
use winapi::shared::guiddef::*;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::shared::winerror::*;
use winapi::shared::wtypes::*;
use winapi::um::combaseapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::oaidl::*;
use winapi::um::oaidl::DISPID; // Required to eliminate ambiguity
use winapi::um::objidlbase::*;
use winapi::um::oleauto::*;
use winapi::um::winnt::*;
use winapi::um::winuser::*;

use lib::{JavascriptCallback, PluginGui};
use win32::client_site::*;
use win32::com_pointer::*;
use win32::ffi::*;

fn error(message: &str) -> Box<dyn Error> {
    From::from(message)
}

struct Window {
    handle: HWND,
}

impl Window {
    // The "plugin_window" string in utf16.
    const CLASS_NAME: [u16; 14] = [
        0x0070, 0x006c, 0x0075, 0x0067, 0x0069, 0x006e, 0x005f, 0x0077,
        0x0069, 0x006e, 0x0064, 0x006f, 0x0077, 0x0000];

    pub fn new(parent: HWND, size: Option<(i32, i32)>) -> Window {
        Window::register_window_class();

        let window_size = size.unwrap_or_else(|| Window::default_size());
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
                window_size.0,
                window_size.1,
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
        js_callback: Rc<JavascriptCallback>) ->
            Result<WebBrowser, Box<dyn Error>>
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
        Result<ComPointer<IWebBrowser2>, Box<dyn Error>>
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
            Err(error("Couldn't get an instance of the 'IWebBrowser2' class"))
        }
    }

    fn browser(&self) -> &IWebBrowser2 {
        self.browser
            .get()
            .unwrap()
    }

    fn open_blank_page(&self) -> Result<(), Box<dyn Error>> {
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
                return Err(error("Couldn't open a blank page"))
            }

            SysFreeString(url);
        }

        Ok(())
    }

    fn document_dispatch(&self) ->
        Result<ComPointer<IDispatch>, Box<dyn Error>>
    {
        let mut result = ComPointer::<IDispatch>::new();

        let success = unsafe {
            self.browser().get_Document(result.as_mut_ptr()) == S_OK &&
                result.get().is_some()
        };

        match success {
            true => Ok(result),
            false => Err(
                error(
                    "The 'IWebBrowser2::get_Document' method returned an \
                    error")),
        }
    }

    fn window_dispatch(&self) -> Result<ComPointer<IDispatch>, Box<dyn Error>> {
        let document_dispatch = self.document_dispatch()?;

        let window_dispatch = document_dispatch
            .query_interface::<IHTMLDocument2>()
            .get()
            .map(|document| {
                let mut window = ComPointer::<IHTMLWindow2>::new();

                unsafe {
                    if document.get_parentWindow(window.as_mut_ptr()) == S_OK {
                        window
                    } else {
                        ComPointer::<IHTMLWindow2>::new()
                    }
                }
            })
            .map(|window| {
                window.query_interface::<IDispatch>()
            })
            .unwrap_or(ComPointer::<IDispatch>::new());

        if window_dispatch.get().is_some() {
            Ok(window_dispatch)
        } else {
            Err(
                error(
                    "Couldn't get an instance of the 'IDispatch' class \
                    for the document window"))
        }
    }

    fn load_html_document(
        &self, html_document: String) -> Result<(), Box<dyn Error>>
    {
        // TODO: do not assume the document is ready
        let document_dispatch = self.document_dispatch()?;

        let stream = ComPointer::<IStream>::from_raw(
            unsafe {
                SHCreateMemStream(
                    html_document.as_ptr(),
                    html_document.len() as u32)
            });

        stream
            .get()
            .ok_or(error("Couldn't get an instance of the 'IStream' class"))?;

        let success = document_dispatch
            .query_interface::<IPersistStreamInit>()
            .get()
            .map(|persist_stream| {
                unsafe {
                    persist_stream.InitNew() == S_OK &&
                    persist_stream.Load(stream.as_ptr()) == S_OK
                }
            })
            .ok_or(
                error(
                    "Couldn't get an instance of the 'IPersistStreamInit' \
                    class"))?;

        match success {
            true => Ok(()),
            false => Err(error("Couldn't load an HTML document")),
        }
    }

    fn embed(
        &self,
        window_handle: HWND,
        js_callback: Rc<JavascriptCallback>) -> Result<(), Box<dyn Error>>
    {
        let ole_object = self.browser.query_interface::<IOleObject>();

        ole_object
            .get()
            .ok_or(
                error("Couldn't get an instance of the 'IOleObject' class"))?;

        let ole_in_place_object =
            ole_object.query_interface::<IOleInPlaceObject>();

        ole_in_place_object
            .get()
            .ok_or(
                error(
                    "Couldn't get an instance of the 'IOleInPlaceObject' \
                    class"))?;

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
            false => Err(error("Couldn't reveal an HTML browser")),
        }
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<dyn Error>> {
        let window_dispatch = self.window_dispatch()?;

        let argument_value: Vec<u16> = OsStr::new(javascript_code)
            .encode_wide()
            .collect();

        unsafe {
            let mut argument: VARIANT = zeroed();

            VariantInit(&mut argument);

            argument.n1.n2_mut().vt = VT_BSTR as u16;
            *argument.n1.n2_mut().n3.bstrVal_mut() = SysAllocStringLen(
                argument_value.as_ptr(),
                argument_value.len() as u32);

            let mut parameters = DISPPARAMS {
                rgvarg: &mut argument,
                rgdispidNamedArgs: null_mut(),
                cArgs: 1,
                cNamedArgs: 0,
            };

            // TODO: cache the 'window_dispatch' object and the method id

            let result = if window_dispatch
                .get()
                .unwrap()
                .Invoke(
                    WebBrowser::window_eval_method_id(&window_dispatch)?,
                    &IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    DISPATCH_METHOD,
                    &mut parameters,
                    null_mut(),
                    null_mut(),
                    null_mut()) == S_OK
            {
                Ok(())
            } else {
                Err(error("Execution of the Javascript code failed"))
            };

            VariantClear(&mut argument);
            result
        }
    }

    fn window_eval_method_id(window_dispatch: &ComPointer<IDispatch>) ->
        Result<DISPID, Box<dyn Error>>
    {
        assert!(window_dispatch.get().is_some());

        // The "eval" string in utf16.
        let mut method_name: Vec<u16> =
            vec![0x0065, 0x0076, 0x0061, 0x006c, 0x0000];
        let mut id: DISPID = 0;

        let result = unsafe {
            window_dispatch
                .get()
                .unwrap()
                .GetIDsOfNames(
                    &IID_NULL,
                    &mut method_name.as_mut_ptr(), 1, LOCALE_SYSTEM_DEFAULT,
                    &mut id)
        };

        if result == S_OK {
            Ok(id)
        } else {
            Err(error("Couldn't get an ID for the 'eval' method"))
        }
    }
}

struct Gui {
    html_document: String,
    js_callback: Rc<JavascriptCallback>,
    web_browser: Option<WebBrowser>,
    window: Option<Window>,
    window_size: Option<(i32, i32)>,
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

    fn open(&mut self, parent_handle: *mut c_void) -> bool {
        let window = Window::new(parent_handle as HWND, self.window_size);

        match WebBrowser::new(
            window.handle,
            self.html_document.clone(),
            self.js_callback.clone()) {
            Ok(browser) => {
                self.window = Some(window);
                self.web_browser = Some(browser);
                true
            },
            Err(_) => false // TODO: Display errors
        }
    }

    fn is_open(&mut self) -> bool {
        self.window.is_some()
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<dyn Error>> {
        if let Some(ref web_browser) = self.web_browser {
            web_browser.execute(javascript_code)
        } else {
            Err(error("The plugin window is closed"))
        }
    }
}

pub fn new_plugin_gui(
    html_document: String,
    js_callback: JavascriptCallback,
    window_size: Option<(i32, i32)>) -> Box<dyn PluginGui>
{
    Box::new(
        Gui {
            html_document: html_document,
            js_callback: Rc::new(js_callback),
            web_browser: None,
            window: None,
            window_size: window_size,
        })
}
