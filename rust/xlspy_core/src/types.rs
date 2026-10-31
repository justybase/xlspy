//! xlspy classes registered from Python, and stdlib types looked up once.

use std::ptr;
use std::sync::atomic::{AtomicPtr, Ordering};

use pyo3::exceptions::PyTypeError;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyString, PyType};

/// Registered classes, read on every encode call: an atomic pointer instead
/// of a lock.  Each holds a strong reference that is never released, so a
/// class replaced while a call is using it stays alive.
static ERROR_TYPE: AtomicPtr<ffi::PyObject> = AtomicPtr::new(ptr::null_mut());
static STYLE_TYPE: AtomicPtr<ffi::PyObject> = AtomicPtr::new(ptr::null_mut());
static DECIMAL: PyOnceLock<Py<PyType>> = PyOnceLock::new();

fn registered<'py>(py: Python<'py>, slot: &AtomicPtr<ffi::PyObject>) -> Option<Bound<'py, PyType>> {
    let p = slot.load(Ordering::Acquire);
    if p.is_null() {
        return None;
    }
    // SAFETY: the slot owns a reference that is never released.
    Some(unsafe { Bound::from_borrowed_ptr(py, p).cast_into_unchecked() })
}

fn register(type_: &Bound<'_, PyAny>, slot: &AtomicPtr<ffi::PyObject>, what: &str) -> PyResult<()> {
    let type_ = type_
        .cast::<PyType>()
        .map_err(|_| PyTypeError::new_err(format!("{what} expects a class")))?;
    slot.store(type_.clone().into_ptr(), Ordering::Release);
    Ok(())
}

/// Register xlspy's CellError, the str subclass of error values.
#[pyfunction]
pub fn set_error_type(type_: &Bound<'_, PyAny>) -> PyResult<()> {
    register(type_, &ERROR_TYPE, "set_error_type")
}

/// Register xlspy.Style, the class of the second item of (value, Style) cells.
#[pyfunction]
pub fn set_style_type(type_: &Bound<'_, PyAny>) -> PyResult<()> {
    register(type_, &STYLE_TYPE, "set_style_type")
}

pub fn error_type(py: Python<'_>) -> Option<Bound<'_, PyType>> {
    registered(py, &ERROR_TYPE)
}

pub fn style_type(py: Python<'_>) -> Option<Bound<'_, PyType>> {
    registered(py, &STYLE_TYPE)
}

/// decimal.Decimal, or None if it cannot be imported.
pub fn decimal_type(py: Python<'_>) -> Option<&Bound<'_, PyType>> {
    DECIMAL.import(py, "decimal", "Decimal").ok()
}

/// `PyObject_TypeCheck(obj, type_)`: exact type or subclass, no `__instancecheck__`.
#[inline]
pub fn type_check(obj: &Bound<'_, PyAny>, type_: &Bound<'_, PyType>) -> bool {
    unsafe { pyo3::ffi::PyObject_TypeCheck(obj.as_ptr(), type_.as_type_ptr()) != 0 }
}

/// Kind of the second item of a 2-tuple cell: 1 format string, 2 Style,
/// 0 anything else (the tuple is then an ordinary value).
#[inline]
pub fn overlay_kind(item: &Bound<'_, PyAny>, style: Option<&Bound<'_, PyType>>) -> u8 {
    if item.is_instance_of::<PyString>() {
        1
    } else if style.is_some_and(|s| type_check(item, s)) {
        2
    } else {
        0
    }
}
