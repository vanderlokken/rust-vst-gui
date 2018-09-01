use std::error::Error;
use std::os::raw::c_void;
use std::ptr::{null, null_mut};
use std::rc::Rc;

use winrt;
use winrt::ComPtr;
use winrt::windows::foundation;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::um::libloaderapi::*;
use winapi::um::winuser::*;

use lib::{JavascriptCallback, PluginGui};

use win32_modern::ffi::*;

fn error(message: &str) -> Box<Error> {
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

struct Gui {
    html_document: String,
    js_callback: Rc<JavascriptCallback>,
    // web_browser: Option<WebBrowser>,
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
        // self.web_browser = None;
        self.window = None;
    }

    fn open(&mut self, parent_handle: *mut c_void) {
        let window = Window::new(parent_handle as HWND);

        // TODO: display errors
        // self.web_browser = WebBrowser::new(
        //     window.handle,
        //     self.html_document.clone(),
        //     self.js_callback.clone()
        // ).ok();
        self.window = Some(window);
    }

    fn is_open(&mut self) -> bool {
        self.window.is_some()
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>> {
        // if let Some(ref web_browser) = self.web_browser {
        //     web_browser.execute(javascript_code)
        // } else {
        //     Err(error("The plugin window is closed"))
        // }

        Err(error("Not implemented"))
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
            // web_browser: None,
            window: None,
        })
}

fn some_stuff() {
    let web_view_control_process_options: ComPtr<WebViewControlProcessOptions> =
        winrt::RtDefaultConstructible::new();

    let web_view_control_process: ComPtr<WebViewControlProcess> =
        WebViewControlProcess::create_with_options(
            &*web_view_control_process_options).unwrap();

    let web_view_async = web_view_control_process.create_web_view_control_async(
        0,
        foundation::Rect {
            X: 0.0,
            Y: 0.0,
            Width: 0.0,
            Height: 0.0
        }).unwrap();

    web_view_async.
}
