use std::env::temp_dir;
use std::error::Error;
use std::ffi::OsStr;
use std::fs::File;
use std::iter::once;
use std::mem::transmute;
use std::os::raw::c_void;
use std::os::windows::ffi::OsStrExt;
use std::ptr::{null, null_mut};
use std::sync::Arc;

use winrt;
use winrt::{ComInterface, ComPtr, FastHString, HString};
use winrt::windows::foundation::{
    AsyncStatus,
    IAsyncInfo,
    IAsyncOperation,
    Rect,
    TypedEventHandler
};
use winrt::windows::foundation::collections::{IIterable, IVector};
use winrt::windows::storage::PathIO;
use winrt::windows::web::ui::interop::{
    IWebViewControlSite,
    WebViewControl,
    WebViewControlProcess,
    WebViewControlProcessOptions
};
use winrt::windows::web::ui::WebViewControlScriptNotifyEventArgs;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::um::combaseapi::*;
use winapi::um::handleapi::CloseHandle;
use winapi::um::libloaderapi::*;
use winapi::um::synchapi::CreateEventW;
use winapi::um::winuser::*;
use winapi::winrt::roapi::{RO_INIT_SINGLETHREADED, RoInitialize};

use lib::{JavascriptCallback, PluginGui};

fn error(message: &str) -> Box<Error> {
    From::from(message)
}

fn map_result<T>(
    result: winrt::Result<T>, method_name: &str) -> Result<T, Box<Error>>
{
    result.map_err(|error|
        From::from(format!("'{}' has failed: {:?}", method_name, error)))
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
    web_view: ComPtr<WebViewControl>,
    string_vector: ComPtr<IVector<HString>>,
}

impl WebBrowser {
    fn new(
        window_handle: HWND,
        html_document: String,
        js_callback: Arc<JavascriptCallback>) -> Result<WebBrowser, Box<Error>>
    {
        let options: ComPtr<WebViewControlProcessOptions> =
            winrt::RtDefaultConstructible::new();

        let process = map_result(
            WebViewControlProcess::create_with_options(&*options),
            "WebViewControlProcess::create_with_options")?;

        let handle = unsafe {transmute::<HWND, i64>(window_handle)};
        let bounds = WebBrowser::determine_control_bounds(window_handle);

        let web_view = WebBrowser::blocking_get(
            process.create_web_view_control_async(handle, bounds),
            "WebViewControlProcess::create_web_view_control_async")?.unwrap();

        map_result(
            web_view.navigate_to_string(&FastHString::new(&html_document)),
            "WebViewControl::navigate_to_string")?;

        let event_handler = move |
                _,
                arguments: *mut WebViewControlScriptNotifyEventArgs| {
            unsafe {(*arguments).get_value()}
                .map(|argument| js_callback(argument.to_string()))
                .map(|_| ())
        };

        map_result(
            web_view.add_script_notify(
                &TypedEventHandler::new(event_handler)),
            "WebViewControl::add_script_notify")?;

        let browser = WebBrowser {
            web_view_control_process: process,
            web_view: web_view,
            string_vector: WebBrowser::create_string_vector()?
        };

        browser.execute("document.write('1');")?;

        Ok(browser)
    }

    fn determine_control_bounds(window_handle: HWND) -> Rect {
        let mut rectangle =
            RECT {left: 0, top: 0, right: 0, bottom: 0};

        unsafe {
            GetClientRect(window_handle, &mut rectangle);
        }

        Rect {
            X: 0.0,
            Y: 0.0,
            Width: (rectangle.right - rectangle.left) as f32,
            Height: (rectangle.bottom - rectangle.top) as f32
        }
    }

    // We can't use the 'RtAsyncOperation::blocking_wait' method because it
    // works only when the multithreaded apartment model is used.
    fn blocking_wait<T: ComInterface>(async: &ComPtr<T>) {
        const TIMEOUT_IN_MILLISECONDS: DWORD = 10;

        let mut index: DWORD = 0;
        let mut event = unsafe {
            CreateEventW(null_mut(), FALSE, FALSE, null())
        };

        while async
                .query_interface::<IAsyncInfo>()
                .map(|async| async.get_status() == Ok(AsyncStatus::Started))
                .unwrap_or(false) {
            unsafe {
                // We use this function to enter the COM modal loop.
                CoWaitForMultipleHandles(
                    COWAIT_DISPATCH_CALLS,
                    TIMEOUT_IN_MILLISECONDS,
                    1,
                    &mut event,
                    &mut index);
            }
        }

        unsafe {
            CloseHandle(event);
        }
    }

    // We can't use the 'RtAsyncOperation::blocking_get' method because it
    // assumes multithreaded apartments.
    fn blocking_get<T: winrt::RtType>(
        async_result: winrt::Result<ComPtr<IAsyncOperation<T>>>,
        method: &str) -> Result<T::Out, Box<Error>>
    {
        let async = map_result(async_result, method)?;

        WebBrowser::blocking_wait(&async);

        Ok(map_result(async.get_results(), method)?)
    }

    // Rust WinRT bindings don't contain an 'IVector' implmentation. So we
    // obtain an instance of this class using the following method.
    fn create_string_vector() -> Result<ComPtr<IVector<HString>>, Box<Error>> {
        const FILENAME: &str = ".empty";

        let path = temp_dir()
            .join(FILENAME)
            .to_str()
            .unwrap_or(FILENAME)
            .to_string();

        let _file = File::create(path.clone())?;

        let async = map_result(
            PathIO::read_lines_async(&FastHString::new(&path)),
            "PathIO::read_lines_async")?;

        WebBrowser::blocking_wait(&async);

        Ok(map_result(
            async.get_results(), "IAsyncOperation::get_results")?.unwrap())
    }

    fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>> {
        let mut arguments = self.string_vector.clone();

        map_result(
            arguments.clear(),
            "IVector::clear")?;
        map_result(
            arguments.append(&FastHString::new(javascript_code)),
            "IVector::append")?;

        let async = map_result(
            self.web_view.invoke_script_async(
                &FastHString::new("eval"),
                &*arguments.query_interface::<IIterable<HString>>().unwrap()),
            "WebViewControl::invoke_script_async")?;

        WebBrowser::blocking_wait(&async);

        map_result(async.get_results(), "IAsyncOperation::get_results")
            .map(|_| ())
    }
}

impl Drop for WebBrowser {
    fn drop(&mut self) {
        self.web_view
            .query_interface::<IWebViewControlSite>()
            .map(|site| site.close());

        self.web_view_control_process
            .terminate()
            .ok();
    }
}

struct Gui {
    html_document: String,
    js_callback: Arc<JavascriptCallback>,
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
        // Note: the following condition cannot be ommitted since the
        // 'Drop::drop' method implementation for the 'WebBrowser' struct
        // can recursively call this method.
        if self.window.is_some() {
            self.window = None;
            self.web_browser = None;
        }
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
            js_callback: Arc::new(js_callback),
            web_browser: None,
            window: None,
        })
}
