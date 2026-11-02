//! A row argument walked as `PySequence_Fast` does it: exact lists and tuples
//! as they are, anything else materialized with `list(row)`.

use std::ffi::CStr;

use pyo3::ffi;
use pyo3::prelude::*;

/// A list or tuple that rows are read from by index.
pub struct FastSeq<'py> {
    seq: Bound<'py, PyAny>,
}

impl<'py> FastSeq<'py> {
    /// The fast sequence of `obj`; the error (a TypeError carrying `message`
    /// for a non-sequence) when it has none.
    pub fn new(obj: &Bound<'py, PyAny>, message: &CStr) -> PyResult<Self> {
        // SAFETY: `obj` is a live object (the GIL is held through `Bound`) and
        // `message` is a NUL-terminated string that outlives the call.
        let seq = unsafe { ffi::PySequence_Fast(obj.as_ptr(), message.as_ptr()) };
        // SAFETY: PySequence_Fast returns a new reference, or NULL with an
        // exception set, which is what from_owned_ptr_or_err expects.
        let seq = unsafe { Bound::from_owned_ptr_or_err(obj.py(), seq) }?;
        Ok(FastSeq { seq })
    }

    /// The current number of items. Python code can shrink a list between
    /// reads, so `get` checks the bound again.
    pub fn len(&self) -> usize {
        // SAFETY: `seq` is the exact list or tuple PySequence_Fast returned,
        // kept alive by `self`.
        unsafe { ffi::PySequence_Fast_GET_SIZE(self.seq.as_ptr()) as usize }
    }

    /// Item `i` as an owned reference, or None when the sequence no longer has it.
    /// Owned, so that Python code run while the item is in use cannot free it.
    pub fn get(&self, i: usize) -> Option<Bound<'py, PyAny>> {
        if i >= self.len() {
            return None;
        }
        // SAFETY: `seq` is an exact list or tuple kept alive by `self`, and
        // `i` is below its current size, checked just above with the GIL held.
        let item =
            unsafe { ffi::PySequence_Fast_GET_ITEM(self.seq.as_ptr(), i as ffi::Py_ssize_t) };
        // SAFETY: GET_ITEM returns a valid borrowed reference; from_borrowed_ptr
        // takes a new reference to it.
        Some(unsafe { Bound::from_borrowed_ptr(self.seq.py(), item) })
    }
}
