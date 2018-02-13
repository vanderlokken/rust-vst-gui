use std::ptr::null_mut;

use winapi::Interface;
use winapi::shared::winerror::S_OK;
use winapi::shared::minwindef::LPVOID;
use winapi::um::unknwnbase::IUnknown;

pub struct ComPointer<T: Interface> {
    pointer: *mut T,
}

impl<T: Interface> ComPointer<T> {
    pub fn new() -> Self {
        ComPointer {
            pointer: null_mut(),
        }
    }

    pub fn from_raw(pointer: *mut T) -> Self {
        ComPointer {
            pointer: pointer,
        }
    }

    // Note: this method doesn't modify the reference counter
    pub fn as_ptr(&self) -> *mut T {
        self.pointer
    }

    // Note: this method doesn't modify the reference counter
    pub fn as_mut_ptr(&mut self) -> &mut *mut T {
        &mut self.pointer
    }

    pub fn get(&self) -> Option<&T> {
        if self.pointer != null_mut() {
            unsafe {
                Some(&*self.pointer)
            }
        } else {
            None
        }
    }

    pub fn query_interface<I: Interface>(&self) -> ComPointer<I> {
        let mut result = ComPointer::<I>::new();

        let success = if self.pointer != null_mut() {
            unsafe {
                let result_pointer = result.as_mut_ptr()
                    as *mut *mut I
                    as *mut LPVOID;

                (*(self.pointer as *mut IUnknown)).QueryInterface(
                    &I::uuidof(), result_pointer) == S_OK
            }
        } else {
            false
        };

        match success {
            true => result,
            false => ComPointer::<I>::new(),
        }
    }
}

impl<T: Interface> Drop for ComPointer<T> {
    fn drop(&mut self) {
        if self.pointer != null_mut() {
            unsafe {
                (*(self.pointer as *mut IUnknown)).Release();
            }
        }
    }
}
