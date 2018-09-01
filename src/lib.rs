#[cfg(all(windows, not(feature = "modern")))]
#[macro_use]
extern crate memoffset;
#[cfg(windows)]
#[macro_use]
extern crate winapi;
#[cfg(all(windows, feature = "modern"))]
extern crate winrt;
extern crate vst;

use std::error::Error;
use std::os::raw::c_void;

#[cfg(all(windows, not(feature = "modern")))]
mod win32;
#[cfg(all(windows, feature = "modern"))]
mod win32_modern;

mod lib {
    use std::error::Error;
    use std::os::raw::c_void;

    pub type JavascriptCallback = Box<Fn(String) -> String>;

    pub trait PluginGui {
        fn size(&self) -> (i32, i32);
        fn position(&self) -> (i32, i32);
        fn close(&mut self);
        fn open(&mut self, parent_handle: *mut c_void);
        fn is_open(&mut self) -> bool;
        fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>>;
    }
}

pub struct PluginGui {
    gui: Box<lib::PluginGui>,
}

impl PluginGui {
    // Calls the Javascript 'eval' function with the specified argument.
    // This method always returns an error when the plugin window is closed.
    pub fn execute(&self, javascript_code: &str) -> Result<(), Box<Error>> {
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

    fn open(&mut self, parent_handle: *mut c_void) {
        self.gui.open(parent_handle)
    }

    fn is_open(&mut self) -> bool {
        self.gui.is_open()
    }
}

pub use lib::JavascriptCallback;

pub fn new_plugin_gui(
    html_document: String, js_callback: JavascriptCallback) -> PluginGui
{
    #[cfg(all(windows, not(feature = "modern")))]
    {
        PluginGui {gui: win32::new_plugin_gui(html_document, js_callback)}
    }

    #[cfg(all(windows, feature = "modern"))]
    {
        PluginGui {
            gui: win32_modern::new_plugin_gui(html_document, js_callback)}
    }
}
