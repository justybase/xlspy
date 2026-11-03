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

#[cfg(test)]
mod tests {
    use super::FastSeq;
    use pyo3::exceptions::PyTypeError;
    use pyo3::prelude::*;
    use pyo3::types::{PyList, PyTuple};

    #[test]
    fn lists_and_tuples_are_read_in_place_and_iterables_materialized() {
        Python::initialize();
        Python::attach(|py| {
            let list = PyList::new(py, [1, 2, 3]).unwrap().into_any();
            let seq = FastSeq::new(&list, c"row").unwrap();
            assert!(seq.seq.is(&list), "an exact list is used as it is");
            assert_eq!(seq.len(), 3);
            assert_eq!(seq.get(2).unwrap().extract::<i64>().unwrap(), 3);
            assert!(seq.get(3).is_none());

            let tuple = PyTuple::new(py, ["a"]).unwrap().into_any();
            assert!(FastSeq::new(&tuple, c"row").unwrap().seq.is(&tuple));

            let generator = py.eval(c"(i * 10 for i in range(4))", None, None).unwrap();
            let seq = FastSeq::new(&generator, c"row").unwrap();
            assert_eq!(seq.len(), 4);
            assert_eq!(seq.get(3).unwrap().extract::<i64>().unwrap(), 30);
        });
    }

    #[test]
    fn a_non_iterable_raises_type_error_with_the_message() {
        Python::initialize();
        Python::attach(|py| {
            let five = 5i64.into_pyobject(py).unwrap().into_any();
            let err = FastSeq::new(&five, c"rows must be lists")
                .err()
                .expect("an int is not a sequence");
            assert!(err.is_instance_of::<PyTypeError>(py));
            assert!(err.to_string().contains("rows must be lists"));
        });
    }

    #[test]
    fn a_list_shrunk_after_creation_is_bounds_checked() {
        Python::initialize();
        Python::attach(|py| {
            let list = PyList::new(py, [1, 2, 3]).unwrap();
            let seq = FastSeq::new(list.as_any(), c"row").unwrap();
            list.del_item(2).unwrap();
            assert_eq!(seq.len(), 2);
            assert!(seq.get(2).is_none());
        });
    }
}
