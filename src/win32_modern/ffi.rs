use std::ptr::null_mut;

use winapi::shared::winerror::S_OK;
use winrt::{
    ComIid,
    ComInterface,
    ComPtr,
    Error,
    Guid,
    HRESULT,
    IActivationFactory,
    IInspectable,
    IInspectableVtbl,
    Result,
    RtActivatable,
    RtClassInterface,
    RtInterface,
    RtNamedClass,
    RtType
};
use winrt::windows::foundation;

// Note: the following macro definitions are taken from the 'winrt' crate
// sources.

macro_rules! DEFINE_IID {
    (
        $name:ident, $l:expr, $w1:expr, $w2:expr, $b1:expr, $b2:expr, $b3:expr,
        $b4:expr, $b5:expr, $b6:expr, $b7:expr, $b8:expr
    ) => {
        const $name: &'static Guid = &Guid {
            Data1: $l,
            Data2: $w1,
            Data3: $w2,
            Data4: [$b1, $b2, $b3, $b4, $b5, $b6, $b7, $b8],
        };
    }
}

macro_rules! RT_INTERFACE {
    ($(#[$attr:meta])* interface $interface:ident<$t1:ident, $t2:ident> $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface<$t1,$t2> $($rest)*);
        unsafe impl<$t1: RtType, $t2: RtType> RtInterface for $interface<$t1,$t2> {}
        unsafe impl<$t1: RtType, $t2: RtType> RtClassInterface for $interface<$t1,$t2> {}
    };

    ($(#[$attr:meta])* interface $interface:ident<$t1:ident> $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface<$t1> $($rest)*);
        unsafe impl<$t1: RtType> RtInterface for $interface<$t1> {}
        unsafe impl<$t1: RtType> RtClassInterface for $interface<$t1> {}
    };

    ($(#[$attr:meta])* interface $interface:ident $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface $($rest)*);
        unsafe impl RtInterface for $interface {}
        unsafe impl RtClassInterface for $interface {}
    };

    ($(#[$attr:meta])* static interface $interface:ident<$t1:ident, $t2:ident> $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface<$t1,$t2> $($rest)*);
        unsafe impl<$t1: RtType, $t2: RtType> RtInterface for $interface<$t1,$t2> {}
    };

    ($(#[$attr:meta])* static interface $interface:ident<$t1:ident> $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface<$t1> $($rest)*);
        unsafe impl<$t1: RtType> RtInterface for $interface<$t1> {}
    };

    ($(#[$attr:meta])* static interface $interface:ident $($rest:tt)*) => {
        RT_INTERFACE!($(#[$attr])* basic $interface $($rest)*);
        unsafe impl RtInterface for $interface {}
    };

    // // version with no methods
    // ($(#[$attr:meta])* basic $interface:ident ($vtbl:ident) : $pinterface:ident ($pvtbl:ty) [$iid:ident]
    //     {}
    // ) => {
    //     #[repr(C)] #[allow(missing_copy_implementations)] #[doc(hidden)]
    //     pub struct $vtbl {
    //         pub parent: $pvtbl
    //     }
    //     $(#[$attr])* #[repr(C)] #[allow(missing_copy_implementations)]
    //     pub struct $interface {
    //         lpVtbl: *const $vtbl
    //     }
    //     impl ComIid for $interface {
    //         #[inline] fn iid() -> &'static Guid { &$iid }
    //     }
    //     impl ComInterface for $interface {
    //         type Vtbl = $vtbl;
    //     }
    //     impl RtType for $interface {
    //         type In = $interface;
    //         type Abi = *mut $interface;
    //         type Out = Option<Self::OutNonNull>;
    //         type OutNonNull = ComPtr<$interface>;

    //         #[doc(hidden)] #[inline] unsafe fn unwrap(v: &Self::In) -> Self::Abi { v as *const _ as *mut _ }
    //         #[doc(hidden)] #[inline] unsafe fn uninitialized() -> Self::Abi { ::std::ptr::null_mut() }
    //         #[doc(hidden)] #[inline] unsafe fn wrap(v: Self::Abi) -> Self::Out { ComPtr::wrap_optional(v) }
    //     }
    //     impl ::std::ops::Deref for $interface {
    //         type Target = $pinterface;
    //         #[inline]
    //         fn deref(&self) -> &$pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    //     impl ::std::ops::DerefMut for $interface {
    //         #[inline]
    //         fn deref_mut(&mut self) -> &mut $pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    // };

    // version with methods, but without generic parameters
    ($(#[$attr:meta])* basic $interface:ident ($vtbl:ident) : $pinterface:ident ($pvtbl:ty) [$iid:ident]
        {$(
            $(#[cfg($cond_attr:meta)])* fn $method:ident(&self $(,$p:ident : $t:ty)*) -> $rtr:ty
        ),+}
    ) => {
        #[repr(C)] #[allow(missing_copy_implementations)] #[doc(hidden)]
        pub struct $vtbl {
            pub parent: $pvtbl
            $(, $(#[cfg($cond_attr)])* pub $method: unsafe extern "system" fn(
                This: *mut $interface
                $(,$p: $t)*
            ) -> $rtr)+
        }
        $(#[$attr])* #[repr(C)] #[allow(missing_copy_implementations)]
        pub struct $interface {
            lpVtbl: *const $vtbl
        }
        impl ComIid for $interface {
            #[inline] fn iid() -> &'static Guid { &$iid }
        }
        impl ComInterface for $interface {
            type Vtbl = $vtbl;
        }
        impl RtType for $interface {
            type In = $interface;
            type Abi = *mut $interface;
            type Out = Option<Self::OutNonNull>;
            type OutNonNull = ComPtr<$interface>;

            #[doc(hidden)] #[inline] unsafe fn unwrap(v: &Self::In) -> Self::Abi { v as *const _ as *mut _ }
            #[doc(hidden)] #[inline] unsafe fn uninitialized() -> Self::Abi { ::std::ptr::null_mut() }
            #[doc(hidden)] #[inline] unsafe fn wrap(v: Self::Abi) -> Self::Out { ComPtr::wrap_optional(v) }
        }
        impl ::std::ops::Deref for $interface {
            type Target = $pinterface;
            #[inline]
            fn deref(&self) -> &$pinterface {
                unsafe { ::std::mem::transmute(self) }
            }
        }
        impl ::std::ops::DerefMut for $interface {
            #[inline]
            fn deref_mut(&mut self) -> &mut $pinterface {
                unsafe { ::std::mem::transmute(self) }
            }
        }
    };
    // // The $iid is actually not necessary, because it refers to the uninstantiated version of the interface,
    // // which is irrelevant at runtime (it is used to generate the IIDs of the parameterized interfaces).
    // ($(#[$attr:meta])* basic $interface:ident<$t1:ident> ($vtbl:ident) : $pinterface:ident ($pvtbl:ty) [$iid:ident]
    //     {$(
    //         $(#[cfg($cond_attr:meta)])* fn $method:ident(&self $(,$p:ident : $t:ty)*) -> $rtr:ty
    //     ),+}
    // ) => {
    //     #[repr(C)] #[allow(missing_copy_implementations)] #[doc(hidden)]
    //     pub struct $vtbl<$t1> where $t1: RtType {
    //         pub parent: $pvtbl
    //         $(, $(#[cfg($cond_attr)])* pub $method: unsafe extern "system" fn(
    //             This: *mut $interface<$t1>
    //             $(,$p: $t)*
    //         ) -> $rtr)+
    //     }
    //     $(#[$attr])* #[repr(C)] #[allow(missing_copy_implementations)]
    //     pub struct $interface<$t1> where $t1: RtType {
    //         lpVtbl: *const $vtbl<$t1>,
    //     }
    //     impl<$t1> ComInterface for $interface<$t1> where $t1: RtType {
    //         type Vtbl = $vtbl<$t1>;
    //     }
    //     impl<$t1> RtType for $interface<$t1> where $t1: RtType{
    //         type In = $interface<$t1>;
    //         type Abi = *mut $interface<$t1>;
    //         type Out = Option<Self::OutNonNull>;
    //         type OutNonNull = ComPtr<$interface<$t1>>;

    //         #[doc(hidden)] #[inline] unsafe fn unwrap(v: &Self::In) -> Self::Abi { v as *const _ as *mut _ }
    //         #[doc(hidden)] #[inline] unsafe fn uninitialized() -> Self::Abi { ::std::ptr::null_mut() }
    //         #[doc(hidden)] #[inline] unsafe fn wrap(v: Self::Abi) -> Self::Out { ComPtr::wrap_optional(v) }
    //     }
    //     impl<$t1> ::std::ops::Deref for $interface<$t1> where $t1: RtType {
    //         type Target = $pinterface;
    //         #[inline]
    //         fn deref(&self) -> &$pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    //     impl<$t1> ::std::ops::DerefMut for $interface<$t1> where $t1: RtType {
    //         #[inline]
    //         fn deref_mut(&mut self) -> &mut $pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    // };

    // ($(#[$attr:meta])* basic $interface:ident<$t1:ident, $t2:ident> ($vtbl:ident) : $pinterface:ident ($pvtbl:ty) [$iid:ident]
    //     {$(
    //         $(#[cfg($cond_attr:meta)])* fn $method:ident(&self $(,$p:ident : $t:ty)*) -> $rtr:ty
    //     ),+}
    // ) => {
    //     #[repr(C)] #[allow(missing_copy_implementations)] #[doc(hidden)]
    //     pub struct $vtbl<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         pub parent: $pvtbl
    //         $(, $(#[cfg($cond_attr)])* pub $method: unsafe extern "system" fn(
    //             This: *mut $interface<$t1, $t2>
    //             $(,$p: $t)*
    //         ) -> $rtr)+
    //     }
    //     $(#[$attr])* #[repr(C)] #[allow(missing_copy_implementations)]
    //     pub struct $interface<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         lpVtbl: *const $vtbl<$t1, $t2>,
    //     }
    //     impl<$t1, $t2> ComInterface for $interface<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         type Vtbl = $vtbl<$t1, $t2>;
    //     }
    //     impl<$t1, $t2> RtType for $interface<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         type In = $interface<$t1, $t2>;
    //         type Abi = *mut $interface<$t1, $t2>;
    //         type Out = Option<Self::OutNonNull>;
    //         type OutNonNull = ComPtr<$interface<$t1, $t2>>;

    //         #[doc(hidden)] #[inline] unsafe fn unwrap(v: &Self::In) -> Self::Abi { v as *const _ as *mut _ }
    //         #[doc(hidden)] #[inline] unsafe fn uninitialized() -> Self::Abi { ::std::ptr::null_mut() }
    //         #[doc(hidden)] #[inline] unsafe fn wrap(v: Self::Abi) -> Self::Out { ComPtr::wrap_optional(v) }
    //     }
    //     impl<$t1, $t2> ::std::ops::Deref for $interface<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         type Target = $pinterface;
    //         #[inline]
    //         fn deref(&self) -> &$pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    //     impl<$t1, $t2> ::std::ops::DerefMut for $interface<$t1, $t2> where $t1: RtType, $t2: RtType {
    //         #[inline]
    //         fn deref_mut(&mut self) -> &mut $pinterface {
    //             unsafe { ::std::mem::transmute(self) }
    //         }
    //     }
    // };
}

macro_rules! RT_CLASS {
    // {static class $cls:ident} => {
    //     pub struct $cls; // does not exist at runtime and has no ABI representation
    // };

    {class $cls:ident : $interface:ty} => {
        pub struct $cls($interface);
        unsafe impl RtInterface for $cls {}
        unsafe impl RtClassInterface for $cls {}
        impl ComInterface for $cls {
            type Vtbl = <$interface as ComInterface>::Vtbl;
        }
        impl ComIid for $cls {
            #[inline] fn iid() -> &'static Guid { <$interface as ComIid>::iid() }
        }
        impl RtType for $cls {
            type In = $cls;
            type Abi = *mut $cls;
            type Out = Option<Self::OutNonNull>;
            type OutNonNull = ComPtr<$cls>;

            #[doc(hidden)] #[inline] unsafe fn unwrap(v: &Self::In) -> Self::Abi { v as *const _ as *mut _ }
            #[doc(hidden)] #[inline] unsafe fn uninitialized() -> Self::Abi { ::std::ptr::null_mut() }
            #[doc(hidden)] #[inline] unsafe fn wrap(v: Self::Abi) -> Self::Out { ComPtr::wrap_optional(v) }
        }
        impl ::std::ops::Deref for $cls {
            type Target = $interface;
            #[inline]
            fn deref(&self) -> &$interface {
                &self.0
            }
        }
        impl ::std::ops::DerefMut for $cls {
            #[inline]
            fn deref_mut(&mut self) -> &mut $interface {
                &mut self.0
            }
        }
    };
}

macro_rules! DEFINE_CLSID {
    ($clsname:ident($id:expr) [$idname:ident]) => {
        const $idname: &'static [u16] = $id; // Full name of the class as null-terminated UTF16 string
        impl RtNamedClass for $clsname {
            #[inline]
            fn name() -> &'static [u16] { $idname }
        }
    }
}

#[inline]
pub fn err<T>(hr: HRESULT) -> Result<T> {
    Err(Error::from_hresult(hr))
}

DEFINE_IID!(IID_IWebViewControl, 1066537750, 48240, 19418, 145, 54, 201, 67, 112, 137, 159, 171);
RT_INTERFACE!{interface IWebViewControl(IWebViewControlVtbl): IInspectable(IInspectableVtbl) [IID_IWebViewControl] {
    fn Unused_get_Source(&self) -> HRESULT,
    fn Unused_put_Source(&self) -> HRESULT,
    fn Unused_get_DocumentTitle(&self) -> HRESULT,
    fn Unused_get_CanGoBack(&self) -> HRESULT,
    fn Unused_get_CanGoForward(&self) -> HRESULT,
    fn Unused_put_DefaultBackgroundColor(&self) -> HRESULT,
    fn Unused_get_DefaultBackgroundColor(&self) -> HRESULT,
    fn Unused_get_ContainsFullScreenElement(&self) -> HRESULT,
    fn Unused_get_Settings(&self) -> HRESULT,
    fn Unused_get_DeferredPermissionRequests(&self) -> HRESULT,
    fn Unused_GoForward(&self) -> HRESULT,
    fn Unused_GoBack(&self) -> HRESULT,
    fn Unused_Refresh(&self) -> HRESULT,
    fn Unused_Stop(&self) -> HRESULT,
    fn Unused_Navigate(&self) -> HRESULT,
    fn Unused_NavigateToString(&self) -> HRESULT,
    fn Unused_NavigateToLocalStreamUri(&self) -> HRESULT,
    fn Unused_NavigateWithHttpRequestMessage(&self) -> HRESULT,
    fn Unused_InvokeScriptAsync(&self) -> HRESULT,
    fn Unused_CapturePreviewToStreamAsync(&self) -> HRESULT,
    fn Unused_CaptureSelectedContentToDataPackageAsync(&self) -> HRESULT,
    fn Unused_BuildLocalStreamUri(&self) -> HRESULT,
    fn Unused_GetDeferredPermissionRequestById(&self) -> HRESULT,
    fn Unused_add_NavigationStarting(&self) -> HRESULT,
    fn Unused_remove_NavigationStarting(&self) -> HRESULT,
    fn Unused_add_ContentLoading(&self) -> HRESULT,
    fn Unused_remove_ContentLoading(&self) -> HRESULT,
    fn Unused_add_DOMContentLoaded(&self) -> HRESULT,
    fn Unused_remove_DOMContentLoaded(&self) -> HRESULT,
    fn Unused_add_NavigationCompleted(&self) -> HRESULT,
    fn Unused_remove_NavigationCompleted(&self) -> HRESULT,
    fn Unused_add_FrameNavigationStarting(&self) -> HRESULT,
    fn Unused_remove_FrameNavigationStarting(&self) -> HRESULT,
    fn Unused_add_FrameContentLoading(&self) -> HRESULT,
    fn Unused_remove_FrameContentLoading(&self) -> HRESULT,
    fn Unused_add_FrameDOMContentLoaded(&self) -> HRESULT,
    fn Unused_remove_FrameDOMContentLoaded(&self) -> HRESULT,
    fn Unused_add_FrameNavigationCompleted(&self) -> HRESULT,
    fn Unused_remove_FrameNavigationCompleted(&self) -> HRESULT,
    fn Unused_add_ScriptNotify(&self) -> HRESULT,
    fn Unused_remove_ScriptNotify(&self) -> HRESULT,
    fn Unused_add_LongRunningScriptDetected(&self) -> HRESULT,
    fn Unused_remove_LongRunningScriptDetected(&self) -> HRESULT,
    fn Unused_add_UnsafeContentWarningDisplaying(&self) -> HRESULT,
    fn Unused_remove_UnsafeContentWarningDisplaying(&self) -> HRESULT,
    fn Unused_add_UnviewableContentIdentified(&self) -> HRESULT,
    fn Unused_remove_UnviewableContentIdentified(&self) -> HRESULT,
    fn Unused_add_PermissionRequested(&self) -> HRESULT,
    fn Unused_remove_PermissionRequested(&self) -> HRESULT,
    fn Unused_add_UnsupportedUriSchemeIdentified(&self) -> HRESULT,
    fn Unused_remove_UnsupportedUriSchemeIdentified(&self) -> HRESULT,
    fn Unused_add_NewWindowRequested(&self) -> HRESULT,
    fn Unused_remove_NewWindowRequested(&self) -> HRESULT,
    fn Unused_add_ContainsFullScreenElementChanged(&self) -> HRESULT,
    fn Unused_remove_ContainsFullScreenElementChanged(&self) -> HRESULT,
    fn Unused_add_WebResourceRequested(&self) -> HRESULT,
    fn Unused_remove_WebResourceRequested(&self) -> HRESULT
}}
RT_CLASS!{class WebViewControl: IWebViewControl}

DEFINE_IID!(IID_IWebViewControlProcessFactory, 1203133689, 41682, 17724, 176, 151, 246, 119, 157, 75, 142, 2);
RT_INTERFACE!{static interface IWebViewControlProcessFactory(IWebViewControlProcessFactoryVtbl): IInspectable(IInspectableVtbl) [IID_IWebViewControlProcessFactory] {
    fn CreateWithOptions(&self, processOptions: *mut WebViewControlProcessOptions, out: *mut *mut WebViewControlProcess) -> HRESULT
}}
impl IWebViewControlProcessFactory {
    #[inline] pub fn create_with_options(&self, processOptions: &WebViewControlProcessOptions) -> Result<ComPtr<WebViewControlProcess>> { unsafe {
        let mut out = null_mut();
        let hr = ((*self.lpVtbl).CreateWithOptions)(self as *const _ as *mut _, processOptions as *const _ as *mut _, &mut out);
        if hr == S_OK { Ok(ComPtr::wrap(out)) } else { err(hr) }
    }}
}

DEFINE_IID!(IID_IWebViewControlProcess, 46605292, 39126, 16970, 182, 62, 198, 19, 108, 54, 160, 242);
RT_INTERFACE!{interface IWebViewControlProcess(IWebViewControlProcessVtbl): IInspectable(IInspectableVtbl) [IID_IWebViewControlProcess] {
    fn Unused_get_ProcessId(&self) -> HRESULT,
    fn Unused_get_EnterpriseId(&self) -> HRESULT,
    fn Unused_get_IsPrivateNetworkClientServerCapabilityEnabled(&self) -> HRESULT,
    fn CreateWebViewControlAsync(
        &self,
        hostWindowHandle: i64,
        bounds: foundation::Rect,
        out: *mut *mut foundation::IAsyncOperation<WebViewControl>
    ) -> HRESULT,
    fn Unused_GetWebViewControls(&self) -> HRESULT,
    fn Unused_Terminate(&self) -> HRESULT,
    fn Unused_add_ProcessExited(&self) -> HRESULT,
    fn Unused_remove_ProcessExited(&self) -> HRESULT
}}

impl IWebViewControlProcess {
    #[inline] pub fn create_web_view_control_async(&self, hostWindowHandle: i64, bounds: foundation::Rect) -> Result<ComPtr<foundation::IAsyncOperation<WebViewControl>>> { unsafe {
        let mut out = null_mut();
        let hr = ((*self.lpVtbl).CreateWebViewControlAsync)(self as *const _ as *mut _, hostWindowHandle, bounds, &mut out);
        if hr == S_OK { Ok(ComPtr::wrap(out)) } else { err(hr) }
    }}
}

RT_CLASS!{class WebViewControlProcess: IWebViewControlProcess}
impl RtActivatable<IWebViewControlProcessFactory> for WebViewControlProcess {}
impl RtActivatable<IActivationFactory> for WebViewControlProcess {}
impl WebViewControlProcess {
    #[inline] pub fn create_with_options(processOptions: &WebViewControlProcessOptions) -> Result<ComPtr<WebViewControlProcess>> {
        <Self as RtActivatable<IWebViewControlProcessFactory>>::get_activation_factory().create_with_options(processOptions)
    }
}
DEFINE_CLSID!(WebViewControlProcess(&[87,105,110,100,111,119,115,46,87,101,98,46,85,73,46,73,110,116,101,114,111,112,46,87,101,98,86,105,101,119,67,111,110,116,114,111,108,80,114,111,99,101,115,115,0]) [CLSID_WebViewControlProcess]);

DEFINE_IID!(IID_IWebViewControlProcessOptions, 483029671, 15318, 18470, 130, 97, 108, 129, 137, 80, 93, 137);
RT_INTERFACE!{interface IWebViewControlProcessOptions(IWebViewControlProcessOptionsVtbl): IInspectable(IInspectableVtbl) [IID_IWebViewControlProcessOptions] {
    fn Unused_put_EnterpriseId(&self) -> HRESULT,
    fn Unused_get_EnterpriseId(&self) -> HRESULT,
    fn Unused_put_PrivateNetworkClientServerCapability(&self) -> HRESULT,
    fn Unused_get_PrivateNetworkClientServerCapability(&self) -> HRESULT
}}

RT_CLASS!{class WebViewControlProcessOptions: IWebViewControlProcessOptions}
impl RtActivatable<IActivationFactory> for WebViewControlProcessOptions {}
DEFINE_CLSID!(WebViewControlProcessOptions(&[87,105,110,100,111,119,115,46,87,101,98,46,85,73,46,73,110,116,101,114,111,112,46,87,101,98,86,105,101,119,67,111,110,116,114,111,108,80,114,111,99,101,115,115,79,112,116,105,111,110,115,0]) [CLSID_WebViewControlProcessOptions]);
