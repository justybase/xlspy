//! sharedStrings.bin: encoder and reader.

use pyo3::exceptions::{PyIndexError, PyTypeError, PyUnicodeEncodeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyList, PyString};

use crate::bytes::{ByteView, le_u32, read_varint, write_varint};
use crate::text::{CELL_TEXT_LIMIT, CodePoints, decode_utf16le};

/// `PyLong_AsUnsignedLongMask` truncated to 32 bits: the C extension's `I`
/// argument format, which wraps instead of checking the range.
pub fn u32_mask(obj: &Bound<'_, PyAny>) -> PyResult<u32> {
    // SAFETY: `obj` is live with the GIL held; the call accepts any object and
    // reports failure as c_ulong::MAX with an exception set, checked below.
    let v = unsafe { pyo3::ffi::PyLong_AsUnsignedLongMask(obj.as_ptr()) };
    if v == std::ffi::c_ulong::MAX
        && let Some(err) = PyErr::take(obj.py())
    {
        return Err(err);
    }
    Ok(v as u32)
}

/// The `UnicodeEncodeError` that `str.encode("utf-16-le")` raises at `pos`.
pub fn surrogate_error(s: &Bound<'_, PyString>, pos: usize) -> PyErr {
    let py = s.py();
    match py.get_type::<PyUnicodeEncodeError>().call1((
        "utf-16-le",
        s,
        pos,
        pos + 1,
        "surrogates not allowed",
    )) {
        Ok(exc) => PyErr::from_value(exc),
        Err(err) => err,
    }
}

pub fn text_limit_error(units: usize) -> PyErr {
    PyValueError::new_err(format!(
        "cell text exceeds Excel's {CELL_TEXT_LIMIT}-character limit ({units} UTF-16 code units)"
    ))
}

/// Encode the complete sharedStrings.bin part (book + SST records).
#[pyfunction]
pub fn encode_xlsb_sst<'py>(
    py: Python<'py>,
    ss_list: &Bound<'py, PyAny>,
    sst_all: &Bound<'py, PyAny>,
    sst_unique: &Bound<'py, PyAny>,
) -> PyResult<Bound<'py, PyBytes>> {
    let sst_all = u32_mask(sst_all)?;
    let sst_unique = u32_mask(sst_unique)?;
    let ss_list = ss_list
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("ss_list must be a list"))?;

    let mut out = Vec::with_capacity(14 + ss_list.len() * 16);
    out.extend_from_slice(&[0x9F, 0x01, 0x08]);
    out.extend_from_slice(&sst_all.to_le_bytes());
    out.extend_from_slice(&sst_unique.to_le_bytes());
    for item in ss_list.iter() {
        let s = item
            .cast::<PyString>()
            .map_err(|_| PyTypeError::new_err("SST entries must be str"))?;
        let cps = CodePoints::of(s)?;
        if let Some(pos) = cps.first_surrogate() {
            return Err(surrogate_error(s, pos));
        }
        let units = cps.utf16_len();
        if units > CELL_TEXT_LIMIT {
            return Err(text_limit_error(units));
        }
        out.push(0x13);
        write_varint(&mut out, 5 + 2 * units as u32);
        out.push(0x00);
        out.extend_from_slice(&(units as u32).to_le_bytes());
        cps.write_utf16le(&mut out);
    }
    out.extend_from_slice(&[0xA0, 0x01, 0x00]);
    Ok(PyBytes::new(py, &out))
}

/// Parse a sharedStrings.bin part into a list of strings.
///
/// Mirrors `_fill_bin_shared_strings` + `BiffReader`, quirks included: the
/// value of the first parsed record is discarded, BrtBeginSst and unknown
/// records are skipped, truncation stops silently, short BrtSSTItem or
/// BrtBeginSst records raise IndexError, and an overrunning string length
/// decodes the bytes present (odd length raises UnicodeDecodeError).
#[pyfunction]
pub fn read_xlsb_shared_strings<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
) -> PyResult<Bound<'py, PyList>> {
    let view = ByteView::new(data)?;
    let data = view.as_slice();
    let result = PyList::empty(py);
    let mut pos = 0usize;
    let mut first_record = true;
    while pos < data.len() {
        let Some(rec_id) = read_varint(data, &mut pos) else {
            break;
        };
        let Some(rec_len) = read_varint(data, &mut pos) else {
            break;
        };
        let rec_len = rec_len as usize;
        if rec_len > data.len() - pos {
            break;
        }
        let body = pos;
        pos += rec_len;
        let is_first = std::mem::replace(&mut first_record, false);

        if rec_id == 159 {
            if rec_len < 5 {
                return Err(PyIndexError::new_err("truncated BrtBeginSst record"));
            }
            continue;
        }
        if rec_id != 0x13 {
            continue;
        }
        if rec_len < 5 {
            return Err(PyIndexError::new_err("truncated BrtSSTItem record"));
        }
        let want = le_u32(data, body + 1).unwrap_or(0) as usize * 2;
        let have = want.min(rec_len - 5);
        let s = decode_utf16le(py, &data[body + 5..body + 5 + have])?;
        if !is_first {
            result.append(s)?;
        }
    }
    Ok(result)
}
