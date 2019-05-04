// This module contains declarations which are missing from the 'winapi' crate.

use std::os::raw::*;

// Non-asterisk imports are required to eliminate ambiguity
use winapi::shared::guiddef::*;
use winapi::shared::minwindef::*;
use winapi::shared::windef::*;
use winapi::shared::windef::SIZE;
use winapi::shared::wtypes::*;
use winapi::shared::wtypesbase::*;
use winapi::um::oaidl::*;
use winapi::um::objidl::{IMoniker, IPersist,  IPersistVtbl};
use winapi::um::objidlbase::*;
use winapi::um::unknwnbase::*;
use winapi::um::winnt::*;
use winapi::um::winuser::*;

#[link(name = "ole32")]
extern "system" {
    pub fn OleInitialize(_: LPVOID) -> HRESULT;
    pub fn OleSetContainedObject(
        pUnknown: *mut IUnknown, fContained: BOOL) -> HRESULT;
}

#[link(name = "shlwapi")]
extern "system" {
    pub fn SHCreateMemStream(pInit: *const BYTE, cbInit: UINT) -> *mut IStream;
}

// We don't use these types so we don't need exact declarations.
pub type IDataObject = IUnknown;
pub type IDropTarget = IUnknown;
pub type IOleCommandTarget = IUnknown;
pub type IOleContainer = IUnknown;
pub type IOleInPlaceActiveObject = IUnknown;
pub type IHTMLWindow2 = IUnknown;

pub const OLEIVERB_INPLACEACTIVATE: LONG = -5;

RIDL!{
    #[uuid(0x00000112, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleObject(IOleObjectVtbl) : IUnknown(IUnknownVtbl) {
        fn SetClientSite(pClientSite: *mut IOleClientSite,) -> HRESULT,
        fn Unused_GetClientSite() -> HRESULT,
        fn Unused_SetHostNames() -> HRESULT,
        fn Unused_Close() -> HRESULT,
        fn Unused_SetMoniker() -> HRESULT,
        fn Unused_GetMoniker() -> HRESULT,
        fn Unused_InitFromData() -> HRESULT,
        fn Unused_GetClipboardData() -> HRESULT,
        fn DoVerb(
            iVerb: LONG,
            lpmsg: LPMSG,
            pActiveSite: *mut IOleClientSite,
            lindex: LONG,
            hwndParent: HWND,
            lprcPosRect: LPCRECT,) -> HRESULT,
        fn Unused_EnumVerbs() -> HRESULT,
        fn Unused_Update() -> HRESULT,
        fn Unused_IsUpToDate() -> HRESULT,
        fn Unused_GetUserClassID() -> HRESULT,
        fn Unused_GetUserType() -> HRESULT,
        fn Unused_SetExtent() -> HRESULT,
        fn Unused_GetExtent() -> HRESULT,
        fn Unused_Advise() -> HRESULT,
        fn Unused_Unadvise() -> HRESULT,
        fn Unused_EnumAdvise() -> HRESULT,
        fn Unused_GetMiscStatus() -> HRESULT,
        fn Unused_SetColorScheme() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x00000118, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleClientSite(IOleClientSiteVtbl) : IUnknown(IUnknownVtbl) {
        fn SaveObject() -> HRESULT,
        fn GetMoniker(
            dwAssign: DWORD,
            dwWhichMoniker: DWORD,
            ppmk: *mut *mut IMoniker,) -> HRESULT,
        fn GetContainer(ppContainer: *mut *mut IOleContainer,) -> HRESULT,
        fn ShowObject() -> HRESULT,
        fn OnShowWindow(fShow: BOOL,) -> HRESULT,
        fn RequestNewObjectLayout() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x00000114, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleWindow(IOleWindowVtbl) : IUnknown(IUnknownVtbl) {
        fn GetWindow(phwnd: *mut HWND,) -> HRESULT,
        fn ContextSensitiveHelp(fEnterMode: BOOL,) -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x00000115, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleInPlaceUIWindow(IOleInPlaceUIWindowVtbl) :
        IOleWindow(IOleWindowVtbl)
    {
        fn GetBorder(lprectBorder: LPRECT,) -> HRESULT,
        fn RequestBorderSpace(
            pborderwidths: LPCRECT,) -> HRESULT,
        fn SetBorderSpace(
            pborderwidths: LPCRECT,) -> HRESULT,
        fn SetActiveObject(
            pActiveObject: *mut IOleInPlaceActiveObject,
            pszObjName: LPCOLESTR,) -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x00000116, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleInPlaceFrame(IOleInPlaceFrameVtbl) :
        IOleInPlaceUIWindow(IOleInPlaceUIWindowVtbl)
    {
        fn InsertMenus(
            hmenuShared: HMENU,
            lpMenuWidths: LPVOID,) -> HRESULT,
        fn SetMenu(
            hmenuShared: HMENU,
            holemenu: HGLOBAL,
            hwndActiveObject: HWND,) -> HRESULT,
        fn RemoveMenus(
            hmenuShared: HMENU,) -> HRESULT,
        fn SetStatusText(
            pszStatusText: LPCOLESTR,) -> HRESULT,
        fn EnableModeless(
            fEnable: BOOL,) -> HRESULT,
        fn TranslateAccelerator(
            lpmsg: LPMSG,
            wID: WORD,) -> HRESULT,
    }
}

STRUCT!{
    struct OLEINPLACEFRAMEINFO {
        cb: UINT,
        fMDIApp: BOOL,
        hwndFrame: HWND,
        haccel: HACCEL,
        cAccelEntries: UINT,
    }
}

RIDL!{
    #[uuid(0x00000119, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleInPlaceSite(IOleInPlaceSiteVtbl) : IOleWindow(IOleWindowVtbl) {
        fn CanInPlaceActivate() -> HRESULT,
        fn OnInPlaceActivate() -> HRESULT,
        fn OnUIActivate() -> HRESULT,
        fn GetWindowContext(
            ppFrame: *mut *mut IOleInPlaceFrame,
            ppDoc: *mut *mut IOleInPlaceUIWindow,
            lprcPosRect: LPRECT,
            lprcClipRect: LPRECT,
            lpFrameInfo: *mut OLEINPLACEFRAMEINFO,) -> HRESULT,
        fn Scroll(scrollExtant: SIZE,) -> HRESULT,
        fn OnUIDeactivate(fUndoable: BOOL,) -> HRESULT,
        fn OnInPlaceDeactivate() -> HRESULT,
        fn DiscardUndoState() -> HRESULT,
        fn DeactivateAndUndo() -> HRESULT,
        fn OnPosRectChange(lprcPosRect: LPCRECT,) -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x00000113, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IOleInPlaceObject(IOleInPlaceObjectVtbl) : IOleWindow(IOleWindowVtbl) {
        fn InPlaceDeactivate() -> HRESULT,
        fn UIDeactivate() -> HRESULT,
        fn SetObjectRects(
            lprcPosRect: LPCRECT,
            lprcClipRect: LPCRECT,) -> HRESULT,
        fn ReactivateAndUndo() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0xeab22ac1, 0x30c1, 0x11cf, 0xa7, 0xeb, 0x00, 0x00, 0xc0, 0x5b, 0xae, 0x0b)]
    interface IWebBrowser(IWebBrowserVtbl) : IDispatch(IDispatchVtbl) {
        fn Unused_GoBack() -> HRESULT,
        fn Unused_GoForward() -> HRESULT,
        fn Unused_GoHome() -> HRESULT,
        fn Unused_GoSearch() -> HRESULT,
        fn Navigate(
            url: BSTR,
            flags: *mut VARIANT,
            targetFrameName: *mut VARIANT,
            postData: *mut VARIANT,
            headers: *mut VARIANT,) -> HRESULT,
        fn Unused_Refresh() -> HRESULT,
        fn Unused_Refresh2() -> HRESULT,
        fn Unused_Stop() -> HRESULT,
        fn Unused_get_Application() -> HRESULT,
        fn Unused_get_Parent() -> HRESULT,
        fn Unused_get_Container() -> HRESULT,
        fn get_Document(ppDisp: *mut *mut IDispatch,) -> HRESULT,
        fn Unused_get_TopLevelContainer() -> HRESULT,
        fn Unused_get_Type() -> HRESULT,
        fn Unused_get_Left() -> HRESULT,
        fn Unused_put_Left() -> HRESULT,
        fn Unused_get_Top() -> HRESULT,
        fn Unused_put_Top() -> HRESULT,
        fn Unused_get_Width() -> HRESULT,
        fn put_Width(width: c_long,) -> HRESULT,
        fn Unused_get_Height() -> HRESULT,
        fn put_Height(height: c_long,) -> HRESULT,
        fn Unused_get_LocationName() -> HRESULT,
        fn Unused_get_LocationURL() -> HRESULT,
        fn Unused_get_Busy() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x0002df05, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46)]
    interface IWebBrowserApp(IWebBrowserAppVtbl) : IWebBrowser(IWebBrowserVtbl) {
        fn Unused_Quit() -> HRESULT,
        fn Unused_ClientToWindow() -> HRESULT,
        fn Unused_PutProperty() -> HRESULT,
        fn Unused_GetProperty() -> HRESULT,
        fn Unused_get_Name() -> HRESULT,
        fn Unused_get_HWND() -> HRESULT,
        fn Unused_get_FullName() -> HRESULT,
        fn Unused_get_Path() -> HRESULT,
        fn Unused_get_Visible() -> HRESULT,
        fn put_Visible(value: VARIANT_BOOL,) -> HRESULT,
        fn Unused_get_StatusBar() -> HRESULT,
        fn Unused_put_StatusBar() -> HRESULT,
        fn Unused_get_StatusText() -> HRESULT,
        fn Unused_put_StatusText() -> HRESULT,
        fn Unused_get_ToolBar() -> HRESULT,
        fn Unused_put_ToolBar() -> HRESULT,
        fn Unused_get_MenuBar() -> HRESULT,
        fn Unused_put_MenuBar() -> HRESULT,
        fn Unused_get_FullScreen() -> HRESULT,
        fn Unused_put_FullScreen() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0xd30c1661, 0xcdaf, 0x11d0, 0x8a, 0x3e, 0x00, 0xc0, 0x4f, 0xc9, 0xe2, 0x6e)]
    interface IWebBrowser2(IWebBrowser2Vtbl) : IWebBrowserApp(IWebBrowserAppVtbl) {
        fn Unused_Navigate2() -> HRESULT,
        fn Unused_QueryStatusWB() -> HRESULT,
        fn Unused_ExecWB() -> HRESULT,
        fn Unused_ShowBrowserBar() -> HRESULT,
        fn Unused_get_ReadyState() -> HRESULT,
        fn Unused_get_Offline() -> HRESULT,
        fn Unused_put_Offline() -> HRESULT,
        fn Unused_get_Silent() -> HRESULT,
        fn Unused_put_Silent() -> HRESULT,
        fn Unused_get_RegisterAsBrowser() -> HRESULT,
        fn Unused_put_RegisterAsBrowser() -> HRESULT,
        fn Unused_get_RegisterAsDropTarget() -> HRESULT,
        fn Unused_put_RegisterAsDropTarget() -> HRESULT,
        fn Unused_get_TheaterMode() -> HRESULT,
        fn Unused_put_TheaterMode() -> HRESULT,
        fn Unused_get_AddressBar() -> HRESULT,
        fn Unused_put_AddressBar() -> HRESULT,
        fn Unused_get_Resizable() -> HRESULT,
        fn Unused_put_Resizable() -> HRESULT,
    }
}

DEFINE_GUID!{
    CLSID_WebBrowser,
    0x8856f961, 0x340a, 0x11d0, 0xa9, 0x6b, 0x00, 0xc0, 0x4f, 0xd7, 0x05, 0xa2}

RIDL!{
    #[uuid(0x7fd52380, 0x4e07, 0x101b, 0xae, 0x2d, 0x08, 0x00, 0x2b, 0x2e, 0xc7, 0x13)]
    interface IPersistStreamInit(IPersistStreamInitVtbl) : IPersist(IPersistVtbl) {
        fn IsDirty() -> HRESULT,
        fn Load(pStm: *mut IStream,) -> HRESULT,
        fn Save(pStm: *mut IStream, fClearDirty: BOOL,) -> HRESULT,
        fn GetSizeMax(pcbSize: *mut ULARGE_INTEGER,) -> HRESULT,
        fn InitNew() -> HRESULT,
    }
}

STRUCT!{
    struct DOCHOSTUIINFO {
        cbSize: c_ulong,
        dwFlags: DWORD,
        dwDoubleClick: DWORD,
        pchHostCss: *mut OLECHAR,
        pchHostNS: *mut OLECHAR,
    }
}

RIDL!{
    #[uuid(0xbd3f23c0, 0xd43e, 0x11cf, 0x89, 0x3b, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x1a)]
    interface IDocHostUIHandler(IDocHostUIHandlerVtbl) : IUnknown(IUnknownVtbl) {
        fn ShowContextMenu(
            dwID: DWORD,
            ppt: *mut POINT,
            pcmdtReserved: *mut IUnknown,
            pdispReserved: *mut IDispatch,) -> HRESULT,
        fn GetHostInfo(pInfo: *mut DOCHOSTUIINFO,) -> HRESULT,
        fn ShowUI(
            dwID: DWORD,
            pActiveObject: *mut IOleInPlaceActiveObject,
            pCommandTarget: *mut IOleCommandTarget,
            pFrame: *mut IOleInPlaceFrame,
            pDoc: *mut IOleInPlaceUIWindow,) -> HRESULT,
        fn HideUI() -> HRESULT,
        fn UpdateUI() -> HRESULT,
        fn EnableModeless(fEnable: BOOL,) -> HRESULT,
        fn OnDocWindowActivate(fActivate: BOOL,) -> HRESULT,
        fn OnFrameWindowActivate(fActivate: BOOL,) -> HRESULT,
        fn ResizeBorder(
            prcBorder: LPCRECT,
            pUIWindow: *mut IOleInPlaceUIWindow,
            fRameWindow: BOOL,) -> HRESULT,
        fn TranslateAccelerator(
            lpMsg: LPMSG,
            pguidCmdGroup: *const GUID,
            nCmdID: DWORD,) -> HRESULT,
        fn GetOptionKeyPath(
            pchKey: *mut LPOLESTR,
            dw: DWORD,) -> HRESULT,
        fn GetDropTarget(
            pDropTarget: *mut IDropTarget,
            ppDropTarget: *mut *mut IDropTarget,) -> HRESULT,
        fn GetExternal(
            ppDispatch: *mut *mut IDispatch,) -> HRESULT,
        fn TranslateUrl(
            dwTranslate: DWORD,
            pchURLIn: LPWSTR,
            ppchURLOut: *mut LPWSTR,) -> HRESULT,
        fn FilterDataObject(
            pDO: *mut IDataObject,
            ppDORet: *mut *mut IDataObject,) -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x626fc520, 0xa41e, 0x11cf, 0xa7, 0x31, 0x00, 0xa0, 0xc9, 0x08, 0x26, 0x37)]
    interface IHTMLDocument(IHTMLDocumentVtbl) : IDispatch(IDispatchVtbl) {
        fn Unused_get_Script() -> HRESULT,
    }
}

RIDL!{
    #[uuid(0x332c4425, 0x26cb, 0x11d0, 0xb4, 0x83, 0x00, 0xc0, 0x4f, 0xd9, 0x01, 0x19)]
    interface IHTMLDocument2(IHTMLDocument2Vtbl) : IHTMLDocument(IHTMLDocumentVtbl) {
        fn Unused_get_all() -> HRESULT,
        fn Unused_get_body() -> HRESULT,
        fn Unused_get_activeElement() -> HRESULT,
        fn Unused_get_images() -> HRESULT,
        fn Unused_get_applets() -> HRESULT,
        fn Unused_get_links() -> HRESULT,
        fn Unused_get_forms() -> HRESULT,
        fn Unused_get_anchors() -> HRESULT,
        fn Unused_put_title() -> HRESULT,
        fn Unused_get_title() -> HRESULT,
        fn Unused_get_scripts() -> HRESULT,
        fn Unused_put_designMode() -> HRESULT,
        fn Unused_get_designMode() -> HRESULT,
        fn Unused_get_selection() -> HRESULT,
        fn Unused_get_readyState() -> HRESULT,
        fn Unused_get_frames() -> HRESULT,
        fn Unused_get_embeds() -> HRESULT,
        fn Unused_get_plugins() -> HRESULT,
        fn Unused_put_alinkColor() -> HRESULT,
        fn Unused_get_alinkColor() -> HRESULT,
        fn Unused_put_bgColor() -> HRESULT,
        fn Unused_get_bgColor() -> HRESULT,
        fn Unused_put_fgColor() -> HRESULT,
        fn Unused_get_fgColor() -> HRESULT,
        fn Unused_put_linkColor() -> HRESULT,
        fn Unused_get_linkColor() -> HRESULT,
        fn Unused_put_vlinkColor() -> HRESULT,
        fn Unused_get_vlinkColor() -> HRESULT,
        fn Unused_get_referrer() -> HRESULT,
        fn Unused_get_location() -> HRESULT,
        fn Unused_get_lastModified() -> HRESULT,
        fn Unused_put_URL() -> HRESULT,
        fn Unused_get_URL() -> HRESULT,
        fn Unused_put_domain() -> HRESULT,
        fn Unused_get_domain() -> HRESULT,
        fn Unused_put_cookie() -> HRESULT,
        fn Unused_get_cookie() -> HRESULT,
        fn Unused_put_expando() -> HRESULT,
        fn Unused_get_expando() -> HRESULT,
        fn Unused_put_charset() -> HRESULT,
        fn Unused_get_charset() -> HRESULT,
        fn Unused_put_defaultCharset() -> HRESULT,
        fn Unused_get_defaultCharset() -> HRESULT,
        fn Unused_get_mimeType() -> HRESULT,
        fn Unused_get_fileSize() -> HRESULT,
        fn Unused_get_fileCreatedDate() -> HRESULT,
        fn Unused_get_fileModifiedDate() -> HRESULT,
        fn Unused_get_fileUpdatedDate() -> HRESULT,
        fn Unused_get_security() -> HRESULT,
        fn Unused_get_protocol() -> HRESULT,
        fn Unused_get_nameProp() -> HRESULT,
        fn Unused_write() -> HRESULT,
        fn Unused_writeln() -> HRESULT,
        fn Unused_open() -> HRESULT,
        fn Unused_close() -> HRESULT,
        fn Unused_clear() -> HRESULT,
        fn Unused_queryCommandSupported() -> HRESULT,
        fn Unused_queryCommandEnabled() -> HRESULT,
        fn Unused_queryCommandState() -> HRESULT,
        fn Unused_queryCommandIndeterm() -> HRESULT,
        fn Unused_queryCommandText() -> HRESULT,
        fn Unused_queryCommandValue() -> HRESULT,
        fn Unused_execCommand() -> HRESULT,
        fn Unused_execCommandShowHelp() -> HRESULT,
        fn Unused_createElement() -> HRESULT,
        fn Unused_put_onhelp() -> HRESULT,
        fn Unused_get_onhelp() -> HRESULT,
        fn Unused_put_onclick() -> HRESULT,
        fn Unused_get_onclick() -> HRESULT,
        fn Unused_put_ondblclick() -> HRESULT,
        fn Unused_get_ondblclick() -> HRESULT,
        fn Unused_put_onkeyup() -> HRESULT,
        fn Unused_get_onkeyup() -> HRESULT,
        fn Unused_put_onkeydown() -> HRESULT,
        fn Unused_get_onkeydown() -> HRESULT,
        fn Unused_put_onkeypress() -> HRESULT,
        fn Unused_get_onkeypress() -> HRESULT,
        fn Unused_put_onmouseup() -> HRESULT,
        fn Unused_get_onmouseup() -> HRESULT,
        fn Unused_put_onmousedown() -> HRESULT,
        fn Unused_get_onmousedown() -> HRESULT,
        fn Unused_put_onmousemove() -> HRESULT,
        fn Unused_get_onmousemove() -> HRESULT,
        fn Unused_put_onmouseout() -> HRESULT,
        fn Unused_get_onmouseout() -> HRESULT,
        fn Unused_put_onmouseover() -> HRESULT,
        fn Unused_get_onmouseover() -> HRESULT,
        fn Unused_put_onreadystatechange() -> HRESULT,
        fn Unused_get_onreadystatechange() -> HRESULT,
        fn Unused_put_onafterupdate() -> HRESULT,
        fn Unused_get_onafterupdate() -> HRESULT,
        fn Unused_put_onrowexit() -> HRESULT,
        fn Unused_get_onrowexit() -> HRESULT,
        fn Unused_put_onrowenter() -> HRESULT,
        fn Unused_get_onrowenter() -> HRESULT,
        fn Unused_put_ondragstart() -> HRESULT,
        fn Unused_get_ondragstart() -> HRESULT,
        fn Unused_put_onselectstart() -> HRESULT,
        fn Unused_get_onselectstart() -> HRESULT,
        fn Unused_elementFromPoint() -> HRESULT,
        fn get_parentWindow(p: *mut *mut IHTMLWindow2,) -> HRESULT,
        fn Unused_get_styleSheets() -> HRESULT,
        fn Unused_put_onbeforeupdate() -> HRESULT,
        fn Unused_get_onbeforeupdate() -> HRESULT,
        fn Unused_put_onerrorupdate() -> HRESULT,
        fn Unused_get_onerrorupdate() -> HRESULT,
        fn Unused_toString() -> HRESULT,
        fn Unused_createStyleSheet() -> HRESULT,
    }
}
