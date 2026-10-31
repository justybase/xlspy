//! Float text exactly as CPython writes and reads it.

use std::ffi::CStr;

use pyo3::ffi;
use pyo3::prelude::*;

/// Append `repr(v)`.  Whole numbers below 1e16 are formatted here
/// ("<digits>.0"); everything else goes through CPython's own routine, so the
/// text is identical to `repr()` (shortest round-trip digits).
pub fn write_repr(py: Python<'_>, out: &mut Vec<u8>, v: f64) -> PyResult<()> {
    if v > -1e16 && v < 1e16 && v == (v as i64) as f64 && !(v == 0.0 && v.is_sign_negative()) {
        let n = v as i64;
        if n < 0 {
            out.push(b'-');
        }
        write_u64(out, n.unsigned_abs());
        out.extend_from_slice(b".0");
        return Ok(());
    }
    // SAFETY: PyOS_double_to_string returns a PyMem-allocated C string (or
    // NULL with MemoryError set); the GIL is held by every caller.
    unsafe {
        let s = ffi::PyOS_double_to_string(
            v,
            b'r' as std::ffi::c_char,
            0,
            ffi::Py_DTSF_ADD_DOT_0,
            std::ptr::null_mut(),
        );
        if s.is_null() {
            return Err(PyErr::fetch(py)); // MemoryError
        }
        out.extend_from_slice(CStr::from_ptr(s).to_bytes());
        ffi::PyMem_Free(s.cast());
    }
    Ok(())
}

/// Append the decimal digits of `v`.
#[inline]
pub fn write_u32(out: &mut Vec<u8>, v: u32) {
    write_u64(out, u64::from(v));
}

#[inline]
fn write_u64(out: &mut Vec<u8>, v: u64) {
    let mut tmp = [0u8; 20];
    let mut i = tmp.len();
    let mut n = v;
    loop {
        i -= 1;
        tmp[i] = b'0' + (n % 10) as u8;
        n /= 10;
        if n == 0 {
            break;
        }
    }
    out.extend_from_slice(&tmp[i..]);
}

#[cfg(test)]
mod tests {
    use super::write_u32;

    #[test]
    fn u32_digits() {
        for v in [0u32, 7, 10, 3, 1234567890, u32::MAX] {
            let mut out = Vec::new();
            write_u32(&mut out, v);
            assert_eq!(out, v.to_string().into_bytes());
        }
    }
}
