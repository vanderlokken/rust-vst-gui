use std::boxed::Box;
use std::ptr::null_mut;
use std::rc::Rc;

// Non-asterisk imports are required to eliminate ambiguity
use winapi::Interface;
use winapi::ctypes::*;
use winapi::shared::guiddef::*;
use winapi::shared::minwindef::*;
use winapi::shared::minwindef::ULONG;
use winapi::shared::ntdef::LPWSTR;
use winapi::shared::windef::*;
use winapi::shared::windef::SIZE;
use winapi::shared::winerror::*;
use winapi::shared::wtypesbase::*;
use winapi::shared::wtypes::VT_BSTR;
use winapi::um::oaidl::*;
use winapi::um::oaidl::DISPID;
use winapi::um::objidl::IMoniker;
use winapi::um::oleauto::*;
use winapi::um::unknwnbase::*;
use winapi::um::winbase::*;
use winapi::um::winnt::LCID;
use winapi::um::winuser::*;

use lib::JavascriptCallback;
use win32::com_pointer::ComPointer;
use win32::ffi::*;

#[repr(C, packed)]
struct ClientSite {
    ole_client_site: IOleClientSite,
    ole_in_place_site: IOleInPlaceSite,
    doc_host_ui_handler: IDocHostUIHandler,
    dispatch: IDispatch,
    ole_in_place_frame: ComPointer<IOleInPlaceFrame>,
    ole_in_place_object: ComPointer<IOleInPlaceObject>,
    reference_counter: ULONG,
    window: HWND,
    callback: Rc<JavascriptCallback>,
}

impl ClientSite {
    fn from_ole_in_place_site(
        instance: *mut IOleInPlaceSite) -> *mut ClientSite
    {
        ClientSite::from_member_and_offset(
            instance as *mut u8, offset_of!(ClientSite, ole_in_place_site))
    }

    fn from_doc_host_ui_handler(
        instance: *mut IDocHostUIHandler) -> *mut ClientSite
    {
        ClientSite::from_member_and_offset(
            instance as *mut u8, offset_of!(ClientSite, doc_host_ui_handler))
    }

    fn from_dispatch(
        instance: *mut IDispatch) -> *mut ClientSite
    {
        ClientSite::from_member_and_offset(
            instance as *mut u8, offset_of!(ClientSite, dispatch))
    }

    fn from_member_and_offset(
        member: *mut u8, offset: usize) -> *mut ClientSite
    {
        unsafe {
            member.offset(-(offset as isize)) as *mut ClientSite
        }
    }
}

const OLE_CLIENT_SITE_VTABLE: IOleClientSiteVtbl = IOleClientSiteVtbl {
    parent: IUnknownVtbl {
        AddRef:         IOleClientSite_AddRef,
        Release:        IOleClientSite_Release,
        QueryInterface: IOleClientSite_QueryInterface,
    },
    SaveObject:             IOleClientSite_SaveObject,
    GetMoniker:             IOleClientSite_GetMoniker,
    GetContainer:           IOleClientSite_GetContainer,
    ShowObject:             IOleClientSite_ShowObject,
    OnShowWindow:           IOleClientSite_OnShowWindow,
    RequestNewObjectLayout: IOleClientSite_RequestNewObjectLayout,
};

const OLE_IN_PLACE_SITE_VTABLE: IOleInPlaceSiteVtbl = IOleInPlaceSiteVtbl {
    parent: IOleWindowVtbl {
        parent: IUnknownVtbl {
            AddRef:         IOleInPlaceSite_AddRef,
            Release:        IOleInPlaceSite_Release,
            QueryInterface: IOleInPlaceSite_QueryInterface,
        },
        GetWindow:            IOleInPlaceSite_GetWindow,
        ContextSensitiveHelp: IOleInPlaceSite_ContextSensitiveHelp,
    },
    CanInPlaceActivate:  IOleInPlaceSite_CanInPlaceActivate,
    OnInPlaceActivate:   IOleInPlaceSite_OnInPlaceActivate,
    OnUIActivate:        IOleInPlaceSite_OnUIActivate,
    GetWindowContext:    IOleInPlaceSite_GetWindowContext,
    Scroll:              IOleInPlaceSite_Scroll,
    OnUIDeactivate:      IOleInPlaceSite_OnUIDeactivate,
    OnInPlaceDeactivate: IOleInPlaceSite_OnInPlaceDeactivate,
    DiscardUndoState:    IOleInPlaceSite_DiscardUndoState,
    DeactivateAndUndo:   IOleInPlaceSite_DeactivateAndUndo,
    OnPosRectChange:     IOleInPlaceSite_OnPosRectChange,
};

const DOC_HOST_UI_HANDLER_VTABLE: IDocHostUIHandlerVtbl =
    IDocHostUIHandlerVtbl {
        parent: IUnknownVtbl {
            AddRef:         IDocHostUIHandler_AddRef,
            Release:        IDocHostUIHandler_Release,
            QueryInterface: IDocHostUIHandler_QueryInterface,
        },
        ShowContextMenu:       IDocHostUIHandler_ShowContextMenu,
        GetHostInfo:           IDocHostUIHandler_GetHostInfo,
        ShowUI:                IDocHostUIHandler_ShowUI,
        HideUI:                IDocHostUIHandler_HideUI,
        UpdateUI:              IDocHostUIHandler_UpdateUI,
        EnableModeless:        IDocHostUIHandler_EnableModeless,
        OnDocWindowActivate:   IDocHostUIHandler_OnDocWindowActivate,
        OnFrameWindowActivate: IDocHostUIHandler_OnFrameWindowActivate,
        ResizeBorder:          IDocHostUIHandler_ResizeBorder,
        TranslateAccelerator:  IDocHostUIHandler_TranslateAccelerator,
        GetOptionKeyPath:      IDocHostUIHandler_GetOptionKeyPath,
        GetDropTarget:         IDocHostUIHandler_GetDropTarget,
        GetExternal:           IDocHostUIHandler_GetExternal,
        TranslateUrl:          IDocHostUIHandler_TranslateUrl,
        FilterDataObject:      IDocHostUIHandler_FilterDataObject,
    };

const DISPATCH_VTABLE: IDispatchVtbl = IDispatchVtbl {
    parent: IUnknownVtbl {
        AddRef:         IDispatch_AddRef,
        Release:        IDispatch_Release,
        QueryInterface: IDispatch_QueryInterface,
    },
    GetTypeInfoCount: IDispatch_GetTypeInfoCount,
    GetTypeInfo:      IDispatch_GetTypeInfo,
    GetIDsOfNames:    IDispatch_GetIDsOfNames,
    Invoke:           IDispatch_Invoke,
};

pub fn new_client_site(
    window: HWND,
    ole_in_place_object: ComPointer<IOleInPlaceObject>,
    callback: Rc<JavascriptCallback>) -> ComPointer<IOleClientSite>
{
    let client_site = Box::new(
        ClientSite {
            ole_client_site: IOleClientSite {
                lpVtbl: &OLE_CLIENT_SITE_VTABLE
            },
            ole_in_place_site: IOleInPlaceSite {
                lpVtbl: &OLE_IN_PLACE_SITE_VTABLE
            },
            doc_host_ui_handler: IDocHostUIHandler {
                lpVtbl: &DOC_HOST_UI_HANDLER_VTABLE
            },
            dispatch: IDispatch {
                lpVtbl: &DISPATCH_VTABLE,
            },
            ole_in_place_frame: new_in_place_frame(window),
            ole_in_place_object: ole_in_place_object,
            reference_counter: 1,
            window: window,
            callback: callback,
        });

    ComPointer::from_raw(Box::into_raw(client_site) as *mut IOleClientSite)
}

unsafe extern "system" fn IOleClientSite_AddRef(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = instance as *mut ClientSite;

    (*client_site).reference_counter += 1;
    (*client_site).reference_counter
}

unsafe extern "system" fn IOleClientSite_Release(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = instance as *mut ClientSite;

    let result = {
        (*client_site).reference_counter -= 1;
        (*client_site).reference_counter
    };

    assert!(result != ULONG::max_value());

    if result == 0 {
        Box::from_raw(client_site);
    }

    result
}

unsafe extern "system" fn IOleClientSite_QueryInterface(
    instance: *mut IUnknown,
    riid: REFIID,
    ppvObject: *mut *mut c_void) -> HRESULT
{
    let client_site = instance as *mut ClientSite;

    *ppvObject = if IsEqualGUID(&*riid, &IUnknown::uuidof()) {
        client_site as *mut c_void
    } else if IsEqualGUID(&*riid, &IOleClientSite::uuidof()) {
        client_site as *mut c_void
    } else if IsEqualGUID(&*riid, &IOleInPlaceSite::uuidof()) {
        &mut (*client_site).ole_in_place_site
            as *mut IOleInPlaceSite
            as *mut c_void
    } else if IsEqualGUID(&*riid, &IOleWindow::uuidof()) {
        &mut (*client_site).ole_in_place_site
            as *mut IOleInPlaceSite
            as *mut c_void
    } else if IsEqualGUID(&*riid, &IDocHostUIHandler::uuidof()) {
        &mut (*client_site).doc_host_ui_handler
            as *mut IDocHostUIHandler
            as *mut c_void
    } else if IsEqualGUID(&*riid, &IDispatch::uuidof()) {
        &mut (*client_site).dispatch
            as *mut IDispatch
            as *mut c_void
    } else {
        null_mut()
    };

    if *ppvObject != null_mut() {
        (*instance).AddRef();
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn IOleClientSite_SaveObject(
    _instance: *mut IOleClientSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleClientSite_GetMoniker(
    _instance: *mut IOleClientSite,
    _dwAssign: DWORD,
    _dwWhichMoniker: DWORD,
    ppmk: *mut *mut IMoniker,) -> HRESULT
{
    *ppmk = null_mut();
    E_NOTIMPL
}

unsafe extern "system" fn IOleClientSite_GetContainer(
    _instance: *mut IOleClientSite,
    ppContainer: *mut *mut IOleContainer) -> HRESULT
{
    *ppContainer = null_mut();
    E_NOINTERFACE
}

unsafe extern "system" fn IOleClientSite_ShowObject(
    _instance: *mut IOleClientSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleClientSite_OnShowWindow(
    _instance: *mut IOleClientSite, _fShow: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleClientSite_RequestNewObjectLayout(
    _instance: *mut IOleClientSite) -> HRESULT
{
    E_NOTIMPL
}

unsafe extern "system" fn IOleInPlaceSite_AddRef(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_ole_in_place_site(
        instance as *mut IOleInPlaceSite);

    (*client_site).ole_client_site.AddRef()
}

unsafe extern "system" fn IOleInPlaceSite_Release(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_ole_in_place_site(
        instance as *mut IOleInPlaceSite);

    (*client_site).ole_client_site.Release()
}

unsafe extern "system" fn IOleInPlaceSite_QueryInterface(
    instance: *mut IUnknown,
    riid: REFIID,
    ppvObject: *mut *mut c_void) -> HRESULT
{
    let client_site = ClientSite::from_ole_in_place_site(
        instance as *mut IOleInPlaceSite);

    (*client_site).ole_client_site.QueryInterface(riid, ppvObject)
}

unsafe extern "system" fn IOleInPlaceSite_GetWindow(
    instance: *mut IOleWindow,
    phwnd: *mut HWND) -> HRESULT
{
    let client_site = ClientSite::from_ole_in_place_site(
        instance as *mut IOleInPlaceSite);

    *phwnd = (*client_site).window;
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_ContextSensitiveHelp(
    _instance: *mut IOleWindow,
    _fEnterMode: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_CanInPlaceActivate(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_OnInPlaceActivate(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_OnUIActivate(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_GetWindowContext(
    instance: *mut IOleInPlaceSite,
    ppFrame: *mut *mut IOleInPlaceFrame,
    ppDoc: *mut *mut IOleInPlaceUIWindow,
    _lprcPosRect: LPRECT,
    _lprcClipRect: LPRECT,
    lpFrameInfo: *mut OLEINPLACEFRAMEINFO) -> HRESULT
{
    let client_site = ClientSite::from_ole_in_place_site(instance);

    *ppFrame = (*client_site).ole_in_place_frame.as_ptr();
    (**ppFrame).AddRef();

    *ppDoc = null_mut();

    // TODO: set rectangles

    (*lpFrameInfo).fMDIApp = FALSE;
    (*lpFrameInfo).hwndFrame = (*client_site).window;
    (*lpFrameInfo).haccel = null_mut();
    (*lpFrameInfo).cAccelEntries = 0;

    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_Scroll(
    _instance: *mut IOleInPlaceSite,
    _scrollExtant: SIZE) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_OnUIDeactivate(
    _instance: *mut IOleInPlaceSite,
    _fUndoable: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_OnInPlaceDeactivate(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_DiscardUndoState(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_DeactivateAndUndo(
    _instance: *mut IOleInPlaceSite) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceSite_OnPosRectChange(
    instance: *mut IOleInPlaceSite,
    lprcPosRect: LPCRECT) -> HRESULT
{
    let client_site = ClientSite::from_ole_in_place_site(instance);

    (*client_site)
        .ole_in_place_object
        .get()
        .map(|ole_in_place_object| {
            ole_in_place_object.SetObjectRects(
                lprcPosRect,
                lprcPosRect);
        });

    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_AddRef(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_doc_host_ui_handler(
        instance as *mut IDocHostUIHandler);

    (*client_site).ole_client_site.AddRef()
}

unsafe extern "system" fn IDocHostUIHandler_Release(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_doc_host_ui_handler(
        instance as *mut IDocHostUIHandler);

    (*client_site).ole_client_site.Release()
}

unsafe extern "system" fn IDocHostUIHandler_QueryInterface(
    instance: *mut IUnknown,
    riid: REFIID,
    ppvObject: *mut *mut c_void) -> HRESULT
{
    let client_site = ClientSite::from_doc_host_ui_handler(
        instance as *mut IDocHostUIHandler);

    (*client_site).ole_client_site.QueryInterface(riid, ppvObject)
}

unsafe extern "system" fn IDocHostUIHandler_ShowContextMenu(
    _instance: *mut IDocHostUIHandler,
    dwID: DWORD,
    _ppt: *mut POINT,
    _pcmdtReserved: *mut IUnknown,
    _pdispReserved: *mut IDispatch) -> HRESULT
{
    const CONTEXT_MENU_CONTROL: DWORD = 0x2;
    const CONTEXT_MENU_TEXTSELECT: DWORD = 0x4;
    const CONTEXT_MENU_VSCROLL: DWORD = 0x9;
    const CONTEXT_MENU_HSCROLL: DWORD = 0x10;

    match dwID {
        CONTEXT_MENU_CONTROL |
        CONTEXT_MENU_TEXTSELECT |
        CONTEXT_MENU_VSCROLL |
        CONTEXT_MENU_HSCROLL => S_FALSE,
        _ => S_OK
    }
}

unsafe extern "system" fn IDocHostUIHandler_GetHostInfo(
    _instance: *mut IDocHostUIHandler,
    pInfo: *mut DOCHOSTUIINFO) -> HRESULT
{
    const DOCHOSTUIFLAG_NO3DBORDER: DWORD = 0x00000004;
    const DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION: DWORD = 0x00010000;
    const DOCHOSTUIFLAG_THEME: DWORD = 0x00040000;
    const DOCHOSTUIFLAG_DPI_AWARE: DWORD = 0x40000000;

    (*pInfo).dwFlags =
        DOCHOSTUIFLAG_NO3DBORDER |
        DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION |
        DOCHOSTUIFLAG_THEME |
        DOCHOSTUIFLAG_DPI_AWARE;
    (*pInfo).dwDoubleClick = 0;

    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_ShowUI(
    _instance: *mut IDocHostUIHandler,
    _dwID: DWORD,
    _pActiveObject: *mut IOleInPlaceActiveObject,
    _pCommandTarget: *mut IOleCommandTarget,
    _pFrame: *mut IOleInPlaceFrame,
    _pDoc: *mut IOleInPlaceUIWindow) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_HideUI(
    _instance: *mut IDocHostUIHandler) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_UpdateUI(
    _instance: *mut IDocHostUIHandler) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_EnableModeless(
    _instance: *mut IDocHostUIHandler,
    _fEnable: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_OnDocWindowActivate(
    _instance: *mut IDocHostUIHandler,
    _fActivate: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_OnFrameWindowActivate(
    _instance: *mut IDocHostUIHandler,
    _fActivate: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_ResizeBorder(
    _instance: *mut IDocHostUIHandler,
    _prcBorder: LPCRECT,
    _pUIWindow: *mut IOleInPlaceUIWindow,
    _fRameWindow: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_TranslateAccelerator(
    _instance: *mut IDocHostUIHandler,
    _lpMsg: LPMSG,
    _pguidCmdGroup: *const GUID,
    _nCmdID: DWORD) -> HRESULT
{
    S_FALSE
}

unsafe extern "system" fn IDocHostUIHandler_GetOptionKeyPath(
    _instance: *mut IDocHostUIHandler,
    pchKey: *mut LPOLESTR,
    _dw: DWORD) -> HRESULT
{
    *pchKey = null_mut();
    S_FALSE
}

unsafe extern "system" fn IDocHostUIHandler_GetDropTarget(
    _instance: *mut IDocHostUIHandler,
    _pDropTarget: *mut IDropTarget,
    _ppDropTarget: *mut *mut IDropTarget) -> HRESULT
{
    E_NOTIMPL
}

unsafe extern "system" fn IDocHostUIHandler_GetExternal(
    instance: *mut IDocHostUIHandler,
    ppDispatch: *mut *mut IDispatch) -> HRESULT
{
    let client_site = ClientSite::from_doc_host_ui_handler(instance);

    (*client_site).ole_client_site.QueryInterface(
        &IDispatch::uuidof(), ppDispatch as *mut *mut c_void);
    S_OK
}

unsafe extern "system" fn IDocHostUIHandler_TranslateUrl(
    _instance: *mut IDocHostUIHandler,
    _dwTranslate: DWORD,
    _pchURLIn: LPWSTR,
    ppchURLOut: *mut LPWSTR) -> HRESULT
{
    *ppchURLOut = null_mut();
    S_FALSE
}

unsafe extern "system" fn IDocHostUIHandler_FilterDataObject(
    _instance: *mut IDocHostUIHandler,
    _pDO: *mut IDataObject,
    ppDORet: *mut *mut IDataObject) -> HRESULT
{
    *ppDORet = null_mut();
    S_FALSE
}

unsafe extern "system" fn IDispatch_AddRef(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_dispatch(
        instance as *mut IDispatch);

    (*client_site).ole_client_site.AddRef()
}

unsafe extern "system" fn IDispatch_Release(
    instance: *mut IUnknown) -> ULONG
{
    let client_site = ClientSite::from_dispatch(
        instance as *mut IDispatch);

    (*client_site).ole_client_site.Release()
}

unsafe extern "system" fn IDispatch_QueryInterface(
    instance: *mut IUnknown,
    riid: REFIID,
    ppvObject: *mut *mut c_void) -> HRESULT
{
    let client_site = ClientSite::from_dispatch(
        instance as *mut IDispatch);

    (*client_site).ole_client_site.QueryInterface(riid, ppvObject)
}

unsafe extern "system" fn IDispatch_GetTypeInfoCount(
    _instance: *mut IDispatch,
    pctinfo: *mut UINT) -> HRESULT
{
    *pctinfo = 0;
    S_OK
}

unsafe extern "system" fn IDispatch_GetTypeInfo(
    _instance: *mut IDispatch,
    _iTInfo: UINT,
    _lcid: LCID,
    ppTInfo: *mut *mut ITypeInfo) -> HRESULT
{
    *ppTInfo = null_mut();
    S_FALSE
}

unsafe extern "system" fn IDispatch_GetIDsOfNames(
    _instance: *mut IDispatch,
    _riid: REFIID,
    rgszNames: *mut LPOLESTR,
    cNames: UINT,
    _lcid: LCID,
    rgDispId: *mut DISPID) -> HRESULT
{
    for index in 0..cNames {
        *rgDispId.offset(index as isize) = DISPID_UNKNOWN;
    }

    // The "invoke" string encoded in utf16.
    const METHOD_NAME: [u16; 7] = [105, 110, 118, 111, 107, 101, 0];

    if cNames != 1 || lstrcmpW(*rgszNames, METHOD_NAME.as_ptr()) != 0 {
        return DISP_E_UNKNOWNNAME;
    }

    *rgDispId = 1;
    S_OK
}

unsafe extern "system" fn IDispatch_Invoke(
    instance: *mut IDispatch,
    dispIdMember: DISPID,
    _riid: REFIID,
    _lcid: LCID,
    wFlags: WORD,
    pDispParams: *mut DISPPARAMS,
    pVarResult: *mut VARIANT,
    _pExcepInfo: *mut EXCEPINFO,
    puArgErr: *mut UINT) -> HRESULT
{
    let client_site = ClientSite::from_dispatch(instance);

    if dispIdMember != 1 {
        return DISP_E_MEMBERNOTFOUND;
    }

    if wFlags != DISPATCH_METHOD &&
        wFlags != DISPATCH_METHOD | DISPATCH_PROPERTYGET
    {
        return DISP_E_BADPARAMCOUNT;
    }

    if (*pDispParams).cNamedArgs != 0 {
        return DISP_E_NONAMEDARGS;
    }

    if (*pDispParams).cArgs != 1 {
        return DISP_E_BADPARAMCOUNT;
    }

    if (*(*pDispParams).rgvarg).n1.n2().vt as u32 != VT_BSTR {
        *puArgErr = 0;
        return DISP_E_TYPEMISMATCH;
    }

    // This is ridiculous
    let argument = *(*(*pDispParams).rgvarg).n1.n2().n3.bstrVal();
    let argument_length = SysStringLen(argument);

    use std::ffi::{OsStr, OsString};
    use std::os::windows::ffi::{OsStrExt, OsStringExt};
    use std::slice;

    let argument_utf8 =
        OsString::from_wide(
            slice::from_raw_parts(
                argument,
                argument_length as usize))
        .into_string();

    if argument_utf8.is_err() {
        return DISP_E_OVERFLOW;
    }

    let result = ((*client_site).callback)(argument_utf8.unwrap());

    if pVarResult != null_mut() {
        VariantInit(pVarResult);

        let result_utf16: Vec<u16> =
            OsStr::new(&result).encode_wide().collect();

        (*pVarResult).n1.n2_mut().vt = VT_BSTR as u16;
        *(*pVarResult).n1.n2_mut().n3.bstrVal_mut() = SysAllocStringLen(
            result_utf16.as_ptr(),
            result_utf16.len() as u32);
    }

    S_OK
}

#[repr(C, packed)]
struct InPlaceFrame {
    ole_in_place_frame: IOleInPlaceFrame,
    reference_counter: ULONG,
    window: HWND,
}

const OLE_IN_PLACE_FRAME_VTABLE: IOleInPlaceFrameVtbl = IOleInPlaceFrameVtbl {
    parent: IOleInPlaceUIWindowVtbl {
        parent: IOleWindowVtbl {
            parent: IUnknownVtbl {
                AddRef:         IOleInPlaceFrame_AddRef,
                Release:        IOleInPlaceFrame_Release,
                QueryInterface: IOleInPlaceFrame_QueryInterface,
            },
            GetWindow:            IOleInPlaceFrame_GetWindow,
            ContextSensitiveHelp: IOleInPlaceFrame_ContextSensitiveHelp,
        },
        GetBorder:          IOleInPlaceFrame_GetBorder,
        RequestBorderSpace: IOleInPlaceFrame_RequestBorderSpace,
        SetBorderSpace:     IOleInPlaceFrame_SetBorderSpace,
        SetActiveObject:    IOleInPlaceFrame_SetActiveObject,
    },
    InsertMenus:          IOleInPlaceFrame_InsertMenus,
    SetMenu:              IOleInPlaceFrame_SetMenu,
    RemoveMenus:          IOleInPlaceFrame_RemoveMenus,
    SetStatusText:        IOleInPlaceFrame_SetStatusText,
    EnableModeless:       IOleInPlaceFrame_EnableModeless,
    TranslateAccelerator: IOleInPlaceFrame_TranslateAccelerator,
};

fn new_in_place_frame(window: HWND) -> ComPointer<IOleInPlaceFrame> {
    let in_place_frame = Box::new(
        InPlaceFrame {
            ole_in_place_frame: IOleInPlaceFrame {
                lpVtbl: &OLE_IN_PLACE_FRAME_VTABLE
            },
            reference_counter: 1,
            window: window,
        });

    ComPointer::from_raw(
        Box::into_raw(in_place_frame) as *mut IOleInPlaceFrame)
}

unsafe extern "system" fn IOleInPlaceFrame_AddRef(
    instance: *mut IUnknown) -> ULONG
{
    let in_place_frame = instance as *mut InPlaceFrame;

    (*in_place_frame).reference_counter += 1;
    (*in_place_frame).reference_counter
}

unsafe extern "system" fn IOleInPlaceFrame_Release(
    instance: *mut IUnknown) -> ULONG
{
    let in_place_frame = instance as *mut InPlaceFrame;

    let result = {
        (*in_place_frame).reference_counter -= 1;
        (*in_place_frame).reference_counter
    };

    assert!(result != ULONG::max_value());

    if result == 0 {
        Box::from_raw(in_place_frame);
    }

    result
}

unsafe extern "system" fn IOleInPlaceFrame_QueryInterface(
    instance: *mut IUnknown,
    riid: REFIID,
    ppvObject: *mut *mut c_void) -> HRESULT
{
    *ppvObject =
        if IsEqualGUID(&*riid, &IUnknown::uuidof()) ||
            IsEqualGUID(&*riid, &IOleWindow::uuidof()) ||
            IsEqualGUID(&*riid, &IOleInPlaceUIWindow::uuidof()) ||
            IsEqualGUID(&*riid, &IOleInPlaceFrame::uuidof())
        {
            instance as *mut c_void
        } else {
            null_mut()
        };

    if *ppvObject != null_mut() {
        (*instance).AddRef();
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn IOleInPlaceFrame_GetWindow(
    instance: *mut IOleWindow,
    phwnd: *mut HWND) -> HRESULT
{
    let in_place_frame = instance as *mut InPlaceFrame;

    *phwnd = (*in_place_frame).window;
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_ContextSensitiveHelp(
    _instance: *mut IOleWindow,
    _fEnterMode: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_GetBorder(
    _instance: *mut IOleInPlaceUIWindow,
    _lprectBorder: LPRECT) -> HRESULT
{
    INPLACE_E_NOTOOLSPACE
}

unsafe extern "system" fn IOleInPlaceFrame_RequestBorderSpace(
    _instance: *mut IOleInPlaceUIWindow,
    _pborderwidths: LPCRECT) -> HRESULT
{
    INPLACE_E_NOTOOLSPACE
}

unsafe extern "system" fn IOleInPlaceFrame_SetBorderSpace(
    _instance: *mut IOleInPlaceUIWindow,
    _pborderwidths: LPCRECT) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_SetActiveObject(
    _instance: *mut IOleInPlaceUIWindow,
    _pActiveObject: *mut IOleInPlaceActiveObject,
    _pszObjName: LPCOLESTR) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_InsertMenus(
    _instance: *mut IOleInPlaceFrame,
    _hmenuShared: HMENU,
    _lpMenuWidths: LPVOID) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_SetMenu(
    _instance: *mut IOleInPlaceFrame,
    _hmenuShared: HMENU,
    _holemenu: HGLOBAL,
    _hwndActiveObject: HWND) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_RemoveMenus(
    _instance: *mut IOleInPlaceFrame,
    _hmenuShared: HMENU) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_SetStatusText(
    _instance: *mut IOleInPlaceFrame,
    _pszStatusText: LPCOLESTR) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_EnableModeless(
    _instance: *mut IOleInPlaceFrame,
    _fEnable: BOOL) -> HRESULT
{
    S_OK
}

unsafe extern "system" fn IOleInPlaceFrame_TranslateAccelerator(
    _instance: *mut IOleInPlaceFrame,
    _lpmsg: LPMSG,
    _wID: WORD) -> HRESULT
{
    S_OK
}
