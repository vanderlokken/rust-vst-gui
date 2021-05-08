extern crate vst;
extern crate web_view;

use std::error::Error;
use std::os::raw::c_void;

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
    js: JavascriptCallback,
    html: String,
}

impl PluginGui {
    // Calls the Javascript 'eval' function with the specified argument.
    // This method always returns an error when the plugin window is closed.
    pub fn execute(&mut self, _javascript_code: &str) -> Result<(), Box<dyn Error>> {
        // note: this is wrong...
        // self.gui.eval(javascript_code);
        Ok(())
    }
}

impl vst::editor::Editor for PluginGui {
    fn size(&self) -> (i32, i32) {
        (1920, 2048)
    }

    fn position(&self) -> (i32, i32) {
        (0,0)
    }

    fn close(&mut self) {
        
    }

    fn open(&mut self, _parent_handle: *mut c_void) -> bool {
        let h = self.html.clone();
        web_view::WebViewBuilder::new()
            .title("VST")
            .content(web_view::Content::Html(h))
            .size(1920, 2048)
            .resizable(true)
            .debug(true)
            .user_data(())
            .invoke_handler(|_webview, arg| {
                (self.js)(arg.to_string());
                Ok(())
            })
            .build()
            .unwrap()
            .run()
            .unwrap();

        true
    }

    fn is_open(&mut self) -> bool {
       true
    }
}

pub use lib::JavascriptCallback;

pub fn new_plugin_gui(
    html_document: String,
    js_callback: JavascriptCallback,
    _window_size: Option<(i32, i32)>) -> PluginGui
{
    PluginGui {
        js: js_callback,
        html: html_document,
    }
}
