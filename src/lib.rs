#[cfg(windows)]
#[macro_use]
extern crate memoffset;
#[cfg(windows)]
#[macro_use]
extern crate winapi;
extern crate vst;
extern crate web_view;

use std::error::Error;
use std::os::raw::c_void;

use web_view::*;

#[cfg(windows)]
mod win32;

mod lib {
    use std::error::Error;
    use std::os::raw::c_void;

    pub type JavascriptCallback = Box<dyn Fn(String) -> String>;

    pub trait PluginGui {
        fn size(&self) -> (i32, i32);
        fn position(&self) -> (i32, i32);
        fn close(&mut self);
        fn open(&mut self, parent_handle: *mut c_void) -> bool;
        fn is_open(&mut self) -> bool;
        fn execute(&self, javascript_code: &str) -> Result<(), Box<dyn Error>>;
    }
}

pub struct PluginGui {
    gui: Box<dyn lib::PluginGui>,
}

impl PluginGui {
    // Calls the Javascript 'eval' function with the specified argument.
    // This method always returns an error when the plugin window is closed.
    pub fn execute(&self, javascript_code: &str) -> Result<(), Box<dyn Error>> {
        self.gui.execute(javascript_code)
    }
}

impl vst::editor::Editor for PluginGui {
    fn size(&self) -> (i32, i32) {
        self.gui.size()
    }

    fn position(&self) -> (i32, i32) {
        self.gui.position()
    }

    fn close(&mut self) {
        self.gui.close()
    }

    fn open(&mut self, parent_handle: *mut c_void) -> bool {
        let html_content = "<html><body><h1>Hello, World!</h1></body></html>";
        web_view::builder()
            .title("This is a test")
            .content(Content::Html(html_content))
            .size(320, 480)
            .resizable(false)
            .debug(true)
            .user_data(())
            .invoke_handler(|_webview, _arg| Ok(()))
            .run()
            .unwrap();

        self.gui.open(parent_handle)
    }

    fn is_open(&mut self) -> bool {
        self.gui.is_open()
    }
}

pub use lib::JavascriptCallback;

pub fn new_plugin_gui(
    html_document: String,
    js_callback: JavascriptCallback,
    window_size: Option<(i32, i32)>) -> PluginGui
{
    #[cfg(windows)]
    {
        PluginGui {
            gui: win32::new_plugin_gui(html_document, js_callback, window_size)
        }
    }
}
