use std;
use std::error::Error;
use std::ffi::OsStr;
use std::iter::once;
use std::mem::transmute;
use std::os::raw::c_void;
use std::os::windows::ffi::OsStrExt;
use std::ptr::{null, null_mut};
use std::rc::Rc;

use winrt;
use winrt::{ComPtr, FastHString, HString, RtAsyncOperation};
use winrt::windows::foundation;
use winrt::windows::foundation::collections::IIterable;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::um::combaseapi::*;
use winapi::um::handleapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::synchapi::*;
use winapi::um::winuser::*;
use winapi::winrt::roapi::{RO_INIT_SINGLETHREADED, RoInitialize};

use lib::{JavascriptCallback, PluginGui};

use win32_modern::ffi::*;

fn error(message: &str) -> Box<Error> {
    From::from(message)
}

fn into_error(message: &str, error: &winrt::Error) -> Box<Error> {
    From::from(format!("{}: {:?}", message, error))
}

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
    web_view_control_process: ComPtr<WebViewControlProcess>,
    // web_view_async: ComPtr<foundation::IAsyncOperation<WebViewControl>>,
    web_view: ComPtr<WebViewControl>,
}

impl WebBrowser {
    fn new(
        window_handle: HWND,
        html_document: String,
        js_callback: Rc<JavascriptCallback>) -> Result<WebBrowser, Box<Error>>
    {
        let options: ComPtr<WebViewControlProcessOptions> =
            winrt::RtDefaultConstructible::new();

        let process = WebViewControlProcess::create_with_options(&*options)
            .map_err(|error| into_error(
                "Couldn't create a control process", &error))?;

        let handle = unsafe {transmute::<HWND, i64>(window_handle)};
        let bounds = WebBrowser::determine_control_bounds(window_handle);

        let web_view_async =
            process.create_web_view_control_async(handle, bounds)
                .map_err(|error| into_error(
                    "'create_web_view_control_async' has failed", &error))?;

        // We can't use the 'RtAsyncOperation::blocking_get' here because the
        // following trait bounds are not satisfied:
        // 'AsyncOperationCompletedHandler<WebViewControl>: ComIid'
        WebBrowser::blocking_wait(
            web_view_async.query_interface::<foundation::IAsyncInfo>()
                .ok_or(error("Couldn't get an IAsyncInfo class instance"))?);

        let browser = WebBrowser {
            web_view_control_process: process,
            web_view: web_view_async
                .get_results()
                .map_err(|error| into_error(
                    "Couldn't create a WebView control instance", &error))?
                .ok_or(error("Couldn't create a WebView control instance"))?
        };

        browser.web_view.navigate_to_string(&FastHString::new(&html_document))
            .map_err(|error| into_error(
                "Couldn't open an HTML document", &error))?;

        browser.web_view.add_script_notify(
            &foundation::TypedEventHandler::new(|_, _| Ok(())));

        browser.execute("alert('1');")?;

        Ok(browser)
    }

    fn determine_control_bounds(window_handle: HWND) -> foundation::Rect {
        let mut rectangle =
            RECT {left: 0, top: 0, right: 0, bottom: 0};

        unsafe {
            GetClientRect(window_handle, &mut rectangle);
        }

        foundation::Rect {
            X: 0.0,
            Y: 0.0,
            Width: (rectangle.right - rectangle.left) as f32,
            Height: (rectangle.bottom - rectangle.top) as f32
        }
    }

    fn blocking_wait(async: ComPtr<foundation::IAsyncInfo>) {
        let mut event = unsafe {
            CreateEventW(null_mut(), FALSE, FALSE, null())
        };

        while async
                .get_status()
                .unwrap_or(foundation::AsyncStatus::Error) ==
                    foundation::AsyncStatus::Started {
            unsafe {
                const TIMEOUT_MSEC: DWORD = 10;
                let mut index: DWORD = 0;

                // We use this function to enter the COM modal loop.
                CoWaitForMultipleHandles(
                    COWAIT_DISPATCH_CALLS,
                    TIMEOUT_MSEC,
                    1,
                    &mut event,
                    &mut index);
            }
        }

        unsafe {
            CloseHandle(event);
        }
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>> {
        // let arguments =
        //     foundation::PropertyValue::create_string_array(
        //         &[&FastHString::new(javascript_code)])
        //     .map_err(|error|
        //         into_error("Couldn't create arguments to 'eval'", &error))?
        //     .ok_or(error("Couldn't create arguments to 'eval'"))?
        //     .query_interface::<IIterable<HString>>();

        // if arguments.is_none() {
        //     return Err(error("Not an 'IIterable<HString>' instance"));
        // }

        // self.web_view
        //     .invoke_script_async(
        //         &FastHString::new("eval"), &*iterable.unwrap())
        //     .map_err(|error|
        //         into_error("Couldn't execute JS code", &error))?
        //     .blocking_get()
        //     .map_err(|error|
        //         into_error("JS code execution has failed", &error))?;

        Ok(())
    }
}

impl Drop for WebBrowser {
    fn drop(&mut self) {
        self.web_view.query_interface::<IWebViewControlSite>().map(
            |site| site.close());

        self.web_view_control_process.terminate().ok();
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

        let web_browser = WebBrowser::new(
            window.handle,
            self.html_document.clone(),
            self.js_callback.clone()
        );

        match web_browser {
            Ok(web_browser) => self.web_browser = Some(web_browser),
            Err(error) => {
                let message: Vec<u16> = OsStr::new(error.description())
                    .encode_wide()
                    .chain(once(0))
                    .collect();

                unsafe {
                    MessageBoxW(
                        null_mut(), message.as_ptr(), null_mut(), MB_ICONERROR);
                }
            }
        }

        self.window = Some(window);
    }

    fn is_open(&mut self) -> bool {
        self.window.is_some()
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>> {
        if let Some(ref web_browser) = self.web_browser {
            web_browser.execute(javascript_code)
        } else {
            Err(error("The plugin window is closed"))
        }
    }
}

pub fn new_plugin_gui(
    html_document: String,
    js_callback: JavascriptCallback) -> Box<PluginGui>
{
    // TODO: check the return value
    unsafe {RoInitialize(RO_INIT_SINGLETHREADED)};

    Box::new(
        Gui {
            html_document: html_document,
            js_callback: Rc::new(js_callback),
            web_browser: None,
            window: None,
        })
}
