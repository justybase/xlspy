//! XLSX row encoder (`<row><c/>...</row>`), mirroring `_write_worksheet_xml`
//! cell by cell.  Rows it cannot encode exactly as the Python writer would
//! return None, and the Python row writer then handles (or rejects) them.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyBytes, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple,
};

use crate::cells::{
    Cell, Serial, Sst, Styles, date_style_xf, excel_serial, fmt_tuple_kind, fmt_xf_value, rollback,
    text_fits,
};
use crate::numfmt::{write_repr, write_u32};
use crate::sst::u32_mask;
use crate::types::{decimal_type, style_type};

fn is_hex(c: u8) -> bool {
    c.is_ascii_hexdigit()
}

/// XML-escape & < > over UTF-8 bytes (multi-byte units never collide).
/// False when the text needs an `_xHHHH_` escape (see xlspy/_text.py): a
/// control character other than tab and line feed, U+FFFE/U+FFFF, or literal
/// text of the form `_xHHHH_`; the Python writer then handles the row.
fn write_escaped(out: &mut Vec<u8>, s: &[u8]) -> bool {
    let mut run = 0;
    let n = s.len();
    for i in 0..n {
        let c = s[i];
        let entity: &[u8] = match c {
            b'&' => b"&amp;",
            b'<' => b"&lt;",
            b'>' => b"&gt;",
            _ => {
                if c < 0x20 {
                    if c != b'\t' && c != b'\n' {
                        return false;
                    }
                } else if c == b'_' {
                    if i + 6 < n
                        && s[i + 1] == b'x'
                        && s[i + 2..i + 6].iter().all(|&h| is_hex(h))
                        && s[i + 6] == b'_'
                    {
                        return false;
                    }
                } else if c == 0xEF
                    && i + 2 < n
                    && s[i + 1] == 0xBF
                    && (s[i + 2] == 0xBE || s[i + 2] == 0xBF)
                {
                    return false;
                }
                continue;
            }
        };
        out.extend_from_slice(&s[run..i]);
        out.extend_from_slice(entity);
        run = i + 1;
    }
    out.extend_from_slice(&s[run..]);
    true
}

fn style_attr(out: &mut Vec<u8>, style: u32) {
    if style != 0 {
        out.extend_from_slice(b" s=\"");
        write_u32(out, style);
        out.push(b'"');
    }
}

/// `<c><v>` or `<c s="N"><v>` for a numeric cell.
fn open_value(out: &mut Vec<u8>, style: u32) {
    out.extend_from_slice(b"<c");
    style_attr(out, style);
    out.extend_from_slice(b"><v>");
}

fn number(py: Python<'_>, out: &mut Vec<u8>, style: u32, v: f64) -> PyResult<Cell> {
    open_value(out, style);
    write_repr(py, out, v)?;
    out.extend_from_slice(b"</v></c>");
    Ok(Cell::Done)
}

struct Row<'a, 'py> {
    out: Vec<u8>,
    is_header: bool,
    sst: Option<Sst<'a, 'py>>,
}

impl Row<'_, '_> {
    /// String cell, mirroring `_write_string_cell`; text over the cell limit
    /// falls back so that the Python writer raises.
    fn string(&mut self, s: &Bound<'_, PyString>, style: u32) -> PyResult<Cell> {
        if !text_fits(s)? {
            return Ok(Cell::Fallback);
        }
        if let Some(sst) = self.sst.as_mut() {
            let idx = sst.index(s)?;
            self.out.extend_from_slice(b"<c t=\"s\"");
            self.style_or_header(style);
            self.out.extend_from_slice(b"><v>");
            write_u32(&mut self.out, idx);
            self.out.extend_from_slice(b"</v></c>");
            return Ok(Cell::Done);
        }
        let Ok(utf8) = s.to_str() else {
            return Ok(Cell::Fallback); // lone surrogates
        };
        let utf8 = utf8.as_bytes();
        let start = self.out.len();
        self.out.extend_from_slice(b"<c t=\"inlineStr\"");
        self.style_or_header(style);
        self.out.extend_from_slice(b"><is><t");
        if matches!(utf8.first(), Some(b' ' | b'\t')) {
            self.out.extend_from_slice(b" xml:space=\"preserve\"");
        }
        self.out.push(b'>');
        if !write_escaped(&mut self.out, utf8) {
            self.out.truncate(start);
            return Ok(Cell::Fallback);
        }
        self.out.extend_from_slice(b"</t></is></c>");
        Ok(Cell::Done)
    }

    fn style_or_header(&mut self, style: u32) {
        if style != 0 {
            style_attr(&mut self.out, style);
        } else if self.is_header {
            self.out.extend_from_slice(b" s=\"3\"");
        }
    }

    fn cell(
        &mut self,
        cell: &Bound<'_, PyAny>,
        style: u32,
        decimal: Option<&Bound<'_, pyo3::types::PyType>>,
    ) -> PyResult<Cell> {
        let py = cell.py();
        if cell.is_none() {
            self.out.extend_from_slice(b"<c");
            style_attr(&mut self.out, style);
            self.out.extend_from_slice(b"/>");
            return Ok(Cell::Done);
        }
        if let Ok(s) = cell.cast::<PyString>() {
            return self.string(s, style);
        }
        if let Ok(b) = cell.cast::<PyBool>() {
            self.out.extend_from_slice(b"<c t=\"b\"");
            style_attr(&mut self.out, style);
            self.out.extend_from_slice(if b.is_true() {
                b"><v>1</v></c>"
            } else {
                b"><v>0</v></c>"
            });
            return Ok(Cell::Done);
        }
        if cell.is_instance_of::<PyInt>() {
            // Ints are written as float(cell), like the Python path: beyond
            // int64 through float() (inf never raises for ints, but past float
            // range float() does, and the row falls back).
            let d = if cell.is_exact_instance_of::<PyInt>() {
                let v = unsafe { pyo3::ffi::PyLong_AsLongLong(cell.as_ptr()) };
                if v == -1 && PyErr::take(py).is_some() {
                    match py_float(cell) {
                        Some(d) => d,
                        None => return Ok(Cell::Fallback),
                    }
                } else {
                    v as f64
                }
            } else {
                match py_float(cell) {
                    Some(d) => d,
                    None => return Ok(Cell::Fallback),
                }
            };
            return number(py, &mut self.out, style, d);
        }
        if let Ok(f) = cell.cast::<PyFloat>() {
            let d;
            if cell.is_exact_instance_of::<PyFloat>() {
                d = f.value();
                if !d.is_finite() {
                    return Ok(Cell::Fallback); // the Python writer raises
                }
            } else {
                // Subclasses: route exactly as the Python writer's ==/!= tests.
                let Some(v) = py_float(cell) else {
                    return Ok(Cell::Fallback);
                };
                d = v;
                // PyObject_RichCompareBool, as C did: its identity shortcut makes
                // `cell != cell` false even for a NaN subclass instance.
                let inf = PyFloat::new(py, f64::INFINITY);
                let ninf = PyFloat::new(py, f64::NEG_INFINITY);
                for (other, op) in [
                    (cell.as_ptr(), pyo3::ffi::Py_NE),
                    (inf.as_ptr(), pyo3::ffi::Py_EQ),
                    (ninf.as_ptr(), pyo3::ffi::Py_EQ),
                ] {
                    match unsafe { pyo3::ffi::PyObject_RichCompareBool(cell.as_ptr(), other, op) } {
                        0 => {}
                        1 => return Ok(Cell::Fallback),
                        _ => {
                            PyErr::take(py);
                            return Ok(Cell::Fallback);
                        }
                    }
                }
            }
            return number(py, &mut self.out, style, d);
        }
        if decimal.is_some_and(|t| cell.is_instance(t).unwrap_or(false)) {
            return match py_float(cell) {
                Some(d) if d.is_finite() => number(py, &mut self.out, style, d),
                _ => Ok(Cell::Fallback),
            };
        }
        let is_dt = cell.is_instance_of::<PyDateTime>();
        if is_dt || cell.is_instance_of::<PyDate>() {
            return match excel_serial(cell, is_dt) {
                Serial::Fallback => Ok(Cell::Fallback),
                Serial::OutOfRange => match cell.str() {
                    // Year out of range: str(cell), like the Python path.
                    Ok(s) => self.string(&s, style),
                    Err(_) => Ok(Cell::Fallback),
                },
                Serial::Value(serial) => {
                    if style != 0 {
                        open_value(&mut self.out, style);
                    } else {
                        self.out.extend_from_slice(if is_dt {
                            b"<c s=\"2\"><v>"
                        } else {
                            b"<c s=\"1\"><v>"
                        });
                    }
                    write_repr(py, &mut self.out, serial)?;
                    self.out.extend_from_slice(b"</v></c>");
                    Ok(Cell::Done)
                }
            };
        }
        // Formula, Hyperlink and anything else: the Python writer.
        Ok(Cell::Fallback)
    }
}

/// `float(cell)`, or None when it raises.
fn py_float(cell: &Bound<'_, PyAny>) -> Option<f64> {
    let f = unsafe { pyo3::ffi::PyNumber_Float(cell.as_ptr()) };
    match unsafe { Bound::from_owned_ptr_or_err(cell.py(), f) } {
        Ok(f) => Some(unsafe { f.cast_into_unchecked::<PyFloat>() }.value()),
        Err(_) => None,
    }
}

/// Encode one XLSX sheet row.  Returns (row_xml, sst_uniq, sst_all), or None
/// when the Python row writer must handle the row; then the strings this
/// call added to the table are removed again.
#[pyfunction]
#[pyo3(signature = (row, ss_dict, ss_list, sst_uniq, sst_all, row_idx, use_shared_strings, style_ids=None, fmt_xf=None, row_attrs=None))]
#[allow(clippy::too_many_arguments)]
pub fn encode_xlsx_row<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
    ss_dict: &Bound<'py, PyAny>,
    ss_list: &Bound<'py, PyAny>,
    sst_uniq: &Bound<'py, PyAny>,
    sst_all: &Bound<'py, PyAny>,
    row_idx: &Bound<'py, PyAny>,
    use_shared_strings: &Bound<'py, PyAny>,
    style_ids: Option<&Bound<'py, PyAny>>,
    fmt_xf: Option<&Bound<'py, PyAny>>,
    row_attrs: Option<&Bound<'py, PyAny>>,
) -> PyResult<Option<(Bound<'py, PyBytes>, u32, u32)>> {
    let (uniq, all, row_idx) = (u32_mask(sst_uniq)?, u32_mask(sst_all)?, u32_mask(row_idx)?);
    let use_sst = use_shared_strings.is_truthy()?;
    let none = py.None().into_bound(py);
    let style_ids = style_ids.unwrap_or(&none);
    let attrs = match row_attrs {
        Some(a) if !a.is_none() => {
            // ASCII attributes built by the writer (ht, outlineLevel, hidden...).
            let a = a
                .cast::<PyString>()
                .map_err(|_| PyTypeError::new_err("row_attrs must be a str or None"))?;
            Some(a.to_str()?.to_owned())
        }
        _ => None,
    };
    let fmt_xf = match fmt_xf {
        Some(f) if !f.is_none() => Some(
            f.cast::<PyDict>()
                .map_err(|_| PyTypeError::new_err("fmt_xf must be a dict or None"))?,
        ),
        _ => None,
    };
    let (Ok(dict), Ok(list)) = (ss_dict.cast::<PyDict>(), ss_list.cast::<PyList>()) else {
        return Err(PyTypeError::new_err(
            "ss_dict must be a dict, ss_list a list",
        ));
    };
    let mut styles = Styles::new(style_ids, true)?;

    // PySequence_Fast: exact lists and tuples as they are, anything else as
    // list(row); a non-sequence row goes to the Python writer.
    let seq =
        unsafe { pyo3::ffi::PySequence_Fast(row.as_ptr(), c"row must be a sequence".as_ptr()) };
    let Ok(seq) = (unsafe { Bound::from_owned_ptr_or_err(py, seq) }) else {
        return Ok(None);
    };
    let n = unsafe { pyo3::ffi::PySequence_Fast_GET_SIZE(seq.as_ptr()) } as usize;
    let list_before = list.len();
    let style_type = style_type(py);
    let decimal = decimal_type(py);

    let mut row = Row {
        out: Vec::with_capacity(32 + n * 24),
        is_header: row_idx == 0,
        sst: use_sst.then_some(Sst {
            dict,
            list,
            uniq,
            all,
        }),
    };
    row.out.extend_from_slice(b"<row r=\"");
    write_u32(&mut row.out, row_idx.wrapping_add(1));
    row.out.push(b'"');
    if let Some(attrs) = &attrs {
        row.out.extend_from_slice(attrs.as_bytes());
    }
    row.out.push(b'>');

    let result = (|| -> PyResult<bool> {
        for i in 0..n {
            // SAFETY: `seq` is an exact list or tuple, kept alive, and i < its size.
            let item = unsafe {
                pyo3::ffi::PySequence_Fast_GET_ITEM(seq.as_ptr(), i as pyo3::ffi::Py_ssize_t)
            };
            let mut cell = unsafe { Bound::from_borrowed_ptr(py, item) };
            let mut style = styles.for_column(i)?;
            let mut explicit_fmt = false; // a (value, "format") cell keeps its format
            if !cell.is_none() && cell.is_instance_of::<PyTuple>() {
                let kind = fmt_tuple_kind(&cell, style_type.as_ref());
                if kind != 0 {
                    // The xf the Python writer recorded in fmt_xf for this format
                    // (and cell style); otherwise the Python path writes the row.
                    match fmt_xf_value(&cell, fmt_xf, &mut style)? {
                        Some(value) => cell = value,
                        None => return Ok(false),
                    }
                    explicit_fmt = kind == 1;
                }
            }
            if !explicit_fmt && !date_style_xf(&cell, fmt_xf, &mut style)? {
                return Ok(false);
            }
            if let Cell::Fallback = row.cell(&cell, style, decimal)? {
                return Ok(false);
            }
        }
        Ok(true)
    })();

    match result {
        Ok(true) => {
            row.out.extend_from_slice(b"</row>");
            let (uniq, all) = row.sst.as_ref().map_or((uniq, all), |s| (s.uniq, s.all));
            // UTF-8 by construction (ASCII markup and UTF-8 text from str), so the
            // Python writer can write the bytes to the sheet part as they are.
            Ok(Some((PyBytes::new(py, &row.out), uniq, all)))
        }
        other => {
            rollback(dict, list, list_before);
            other.map(|_| None)
        }
    }
}
