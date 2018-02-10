#[cfg(windows)]
#[macro_use]
extern crate memoffset;
#[cfg(windows)]
#[macro_use]
extern crate winapi;
extern crate vst;

use std::os::raw::c_void;

#[cfg(windows)]
mod win32;

mod lib {
    use std::os::raw::c_void;

    pub type JavascriptCallback = Box<Fn(String) -> String>;

    pub trait PluginGui {
        fn size(&self) -> (i32, i32);
        fn position(&self) -> (i32, i32);
        fn close(&mut self);
        fn open(&mut self, parent_handle: *mut c_void);
        fn is_open(&mut self) -> bool;
    }
}

pub struct PluginGui {
    gui: Box<lib::PluginGui>,
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
    #[cfg(windows)]
    {
        PluginGui {gui: win32::new_plugin_gui(html_document, js_callback)}
    }
}
