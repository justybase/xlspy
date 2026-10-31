//! Access to the code points of a Python str without decoding it: CPython
//! strings may hold lone surrogates, which a Rust `String` cannot.

use pyo3::prelude::*;
use pyo3::types::PyString;
use pyo3::types::string::PyStringData;

/// Excel's per-cell text limit in UTF-16 code units, for XLSX and XLSB alike
/// (see xlspy/_text.py, which owns the error message).
pub const CELL_TEXT_LIMIT: usize = 32767;

/// The canonical code points of a str (its PEP 393 storage).
#[derive(Clone, Copy)]
pub enum CodePoints<'a> {
    Ucs1(&'a [u8]),
    Ucs2(&'a [u16]),
    Ucs4(&'a [u32]),
}

impl<'a> CodePoints<'a> {
    pub fn of(s: &'a Bound<'_, PyString>) -> PyResult<Self> {
        // SAFETY: the slice borrows `s`, which keeps the str alive; strs are
        // immutable.
        Ok(match unsafe { s.data()? } {
            PyStringData::Ucs1(d) => CodePoints::Ucs1(d),
            PyStringData::Ucs2(d) => CodePoints::Ucs2(d),
            PyStringData::Ucs4(d) => CodePoints::Ucs4(d),
        })
    }

    /// Code points before the first `\n` (the width of a cell's first line).
    pub fn first_line_len(&self) -> usize {
        match self {
            CodePoints::Ucs1(d) => d.iter().position(|&c| c == b'\n').unwrap_or(d.len()),
            CodePoints::Ucs2(d) => d.iter().position(|&c| c == 0x0A).unwrap_or(d.len()),
            CodePoints::Ucs4(d) => d.iter().position(|&c| c == 0x0A).unwrap_or(d.len()),
        }
    }

    /// Index of the first surrogate code point (0xD800..=0xDFFF), if any.
    pub fn first_surrogate(&self) -> Option<usize> {
        match self {
            CodePoints::Ucs1(_) => None,
            CodePoints::Ucs2(d) => d.iter().position(|&c| (0xD800..=0xDFFF).contains(&c)),
            CodePoints::Ucs4(d) => d.iter().position(|&c| (0xD800..=0xDFFF).contains(&c)),
        }
    }

    /// UTF-16 code units: non-BMP characters count two.
    pub fn utf16_len(&self) -> usize {
        match self {
            CodePoints::Ucs1(d) => d.len(),
            CodePoints::Ucs2(d) => d.len(),
            CodePoints::Ucs4(d) => d.len() + d.iter().filter(|&&c| c >= 0x10000).count(),
        }
    }

    /// Append the UTF-16LE encoding (surrogates already rejected by the caller).
    pub fn write_utf16le(&self, out: &mut Vec<u8>) {
        match self {
            CodePoints::Ucs1(d) => {
                let start = out.len();
                out.resize(start + d.len() * 2, 0);
                for (dst, &c) in out[start..].as_chunks_mut::<2>().0.iter_mut().zip(d.iter()) {
                    dst[0] = c;
                }
            }
            CodePoints::Ucs2(d) => {
                let start = out.len();
                out.resize(start + d.len() * 2, 0);
                for (dst, &c) in out[start..].as_chunks_mut::<2>().0.iter_mut().zip(d.iter()) {
                    dst.copy_from_slice(&c.to_le_bytes());
                }
            }
            CodePoints::Ucs4(d) => {
                out.reserve(d.len() * 2);
                for &c in *d {
                    if c < 0x10000 {
                        out.extend_from_slice(&(c as u16).to_le_bytes());
                    } else {
                        let c = c - 0x10000;
                        out.extend_from_slice(&(0xD800 | (c >> 10) as u16).to_le_bytes());
                        out.extend_from_slice(&(0xDC00 | (c & 0x3FF) as u16).to_le_bytes());
                    }
                }
            }
        }
    }
}

/// `bytes.decode("utf-16-le", "surrogatepass")`, as the C reader did.
pub fn decode_utf16le<'py>(py: Python<'py>, data: &[u8]) -> PyResult<Bound<'py, PyString>> {
    let mut byteorder: std::ffi::c_int = -1;
    unsafe {
        let obj = pyo3::ffi::PyUnicode_DecodeUTF16(
            data.as_ptr().cast(),
            data.len() as pyo3::ffi::Py_ssize_t,
            c"surrogatepass".as_ptr(),
            &mut byteorder,
        );
        Ok(Bound::from_owned_ptr_or_err(py, obj)?.cast_into_unchecked())
    }
}
