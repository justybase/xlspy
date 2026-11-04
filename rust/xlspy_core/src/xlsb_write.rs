//! XLSB row encoders (BrtRowHdr + cell records), shared-string and inline.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyBytes, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple, PyType,
};

use crate::bytes::write_varint;
use crate::cells::{
    Cell, Hooked, Serial, Sst, Styles, call_hook, date_style_xf, excel_serial, fmt_tuple_kind,
    fmt_xf_value, numpy_item, rollback, text_fits,
};
use crate::pyseq::FastSeq;
use crate::sst::u32_mask;
use crate::text::{CELL_TEXT_LIMIT, CodePoints};
use crate::types::{decimal_type, int_as_i64, style_type};

/// Where text cells go: the shared-strings table, or inline (BrtCellSt).
enum Strings<'a, 'py> {
    Shared(Sst<'a, 'py>),
    Inline,
}

struct RowBuf {
    out: Vec<u8>,
}

impl RowBuf {
    fn cell(&mut self, rec: [u8; 2], col: u32, style: u32) {
        self.out.extend_from_slice(&rec);
        self.out.extend_from_slice(&col.to_le_bytes());
        self.out.extend_from_slice(&style.to_le_bytes());
    }

    fn double(&mut self, col: u32, style: u32, value: f64) {
        self.cell([0x05, 0x10], col, style);
        self.out.extend_from_slice(&value.to_le_bytes());
    }
}

/// One text cell in the storage the caller chose.
fn text_cell(
    buf: &mut RowBuf,
    s: &Bound<'_, PyString>,
    strings: &mut Strings<'_, '_>,
    col: u32,
    row_idx: u32,
    style: u32,
) -> PyResult<Cell> {
    let style_ref = if style != 0 {
        style
    } else if row_idx == 0 {
        3
    } else {
        0
    };
    match strings {
        Strings::Shared(sst) => {
            if !text_fits(s)? {
                return Ok(Cell::Fallback);
            }
            let idx = sst.index(s)?;
            buf.cell([0x07, 0x0C], col, style_ref);
            buf.out.extend_from_slice(&idx.to_le_bytes());
        }
        Strings::Inline => {
            // The count is in UTF-16 code units, as Excel reads it.  Lone
            // surrogates and over-long text go to Python, which owns the error.
            let cps = CodePoints::of(s)?;
            if cps.first_surrogate().is_some() {
                return Ok(Cell::Fallback);
            }
            let units = cps.utf16_len();
            if units > CELL_TEXT_LIMIT {
                return Ok(Cell::Fallback);
            }
            buf.out.push(0x06);
            write_varint(&mut buf.out, 12 + 2 * units as u32);
            buf.out.extend_from_slice(&col.to_le_bytes());
            buf.out.extend_from_slice(&style_ref.to_le_bytes());
            buf.out.extend_from_slice(&(units as u32).to_le_bytes());
            cps.write_utf16le(&mut buf.out);
        }
    }
    Ok(Cell::Done)
}

/// str(cell) as a text cell: out-of-range dates, ints beyond float range and
/// Decimals whose float() raises, like the Python encoder.
fn str_cell(
    buf: &mut RowBuf,
    cell: &Bound<'_, PyAny>,
    strings: &mut Strings<'_, '_>,
    col: u32,
    row_idx: u32,
    style: u32,
) -> PyResult<Cell> {
    match cell.str() {
        Ok(s) => text_cell(buf, &s, strings, col, row_idx, style),
        Err(_) => Ok(Cell::Fallback), // the Python encoder calls str() again and raises
    }
}

/// How one value was encoded.
enum Value {
    Done,
    /// The row goes to the Python encoder (which may raise).
    Fallback,
    /// A type the encoder does not know: ask the writer's hook.
    Unknown,
}

/// Encode one cell value (None: a blank, written only with a style).
fn value_cell(
    buf: &mut RowBuf,
    cell: &Bound<'_, PyAny>,
    strings: &mut Strings<'_, '_>,
    col: u32,
    row_idx: u32,
    style: u32,
    decimal: Option<&Bound<'_, PyType>>,
) -> PyResult<Value> {
    let py = cell.py();
    let done = if cell.is_none() {
        if style != 0 {
            buf.cell([0x01, 0x08], col, style);
        }
        Cell::Done
    } else if let Ok(s) = cell.cast::<PyString>() {
        text_cell(buf, s, strings, col, row_idx, style)?
    } else if let Ok(b) = cell.cast::<PyBool>() {
        buf.cell([0x04, 0x09], col, style);
        buf.out.push(u8::from(b.is_true()));
        Cell::Done
    } else if cell.is_instance_of::<PyInt>() {
        match int_as_i64(cell) {
            // Beyond int64: a double, like float(cell) in Python; beyond
            // float range, str(cell).
            None => match cell.extract::<f64>() {
                Ok(f) => {
                    buf.double(col, style, f);
                    Cell::Done
                }
                Err(_) => str_cell(buf, cell, strings, col, row_idx, style)?,
            },
            Some(v) if (-536_870_912..=536_870_911).contains(&v) => {
                buf.cell([0x02, 0x0C], col, style);
                buf.out
                    .extend_from_slice(&(((v as u32) << 2) | 2).to_le_bytes());
                Cell::Done
            }
            Some(v) => {
                buf.double(col, style, v as f64);
                Cell::Done
            }
        }
    } else if let Ok(f) = cell.cast::<PyFloat>() {
        let v = f.value();
        if !v.is_finite() {
            return Ok(Value::Fallback); // the Python encoder raises
        }
        buf.double(col, style, v);
        Cell::Done
    } else if decimal.is_some_and(|d| cell.is_instance(d).unwrap_or(false)) {
        match cell
            .call_method0(pyo3::intern!(py, "__float__"))
            .and_then(|f| f.extract::<f64>())
        {
            Ok(v) if !v.is_finite() => return Ok(Value::Fallback),
            Ok(v) => {
                buf.double(col, style, v);
                Cell::Done
            }
            Err(_) => str_cell(buf, cell, strings, col, row_idx, style)?,
        }
    } else if cell.is_instance_of::<PyDateTime>() || cell.is_instance_of::<PyDate>() {
        let is_dt = cell.is_instance_of::<PyDateTime>();
        match excel_serial(cell, is_dt) {
            // Unreadable fields (pandas' NaT) or past exact microseconds.
            Serial::Fallback => return Ok(Value::Unknown),
            Serial::OutOfRange => str_cell(buf, cell, strings, col, row_idx, style)?,
            Serial::Value(serial) => {
                let default = if is_dt { 1 } else { 2 };
                buf.double(col, if style != 0 { style } else { default }, serial);
                Cell::Done
            }
        }
    } else {
        // Formula, Hyperlink, numpy scalars and other objects.
        return Ok(Value::Unknown);
    };
    Ok(match done {
        Cell::Done => Value::Done,
        Cell::Fallback => Value::Fallback,
    })
}

/// Encode one row: the bytes, or None when the Python encoder must write it.
#[allow(clippy::too_many_arguments)]
fn encode_row<'py>(
    row: &Bound<'py, PyAny>,
    strings: &mut Strings<'_, 'py>,
    row_idx: u32,
    style_ids: &Bound<'py, PyAny>,
    fmt_xf: &Bound<'py, PyAny>,
    row_header: &Bound<'py, PyAny>,
    hook: &Bound<'py, PyAny>,
) -> PyResult<Option<Vec<u8>>> {
    let py = row.py();
    // A list or tuple is read in place, any other sequence (a numpy row) as
    // list(row); a non-sequence goes to the Python encoder, as in XLSX.
    let Ok(row) = FastSeq::new(row, c"row must be a sequence") else {
        PyErr::take(py);
        return Ok(None);
    };
    let fmt_xf = if fmt_xf.is_none() {
        None
    } else {
        Some(
            fmt_xf
                .cast::<PyDict>()
                .map_err(|_| PyTypeError::new_err("fmt_xf must be a dict or None"))?,
        )
    };
    let header = if row_header.is_none() {
        None
    } else {
        match row_header.cast::<PyBytes>() {
            Ok(b) if b.as_bytes().len() <= 256 => Some(b.as_bytes()),
            _ => {
                return Err(PyTypeError::new_err(
                    "row_header must be a BrtRowHdr record (bytes) or None",
                ));
            }
        }
    };
    // Bad dense-list entries count as no style here (the XLSX encoder raises).
    let mut styles = Styles::new(style_ids, false)?;
    let style_type = style_type(py);
    let decimal = decimal_type(py);

    let row_len = row.len();
    let mut buf = RowBuf {
        out: Vec::with_capacity(32 + row_len * 16),
    };
    match header {
        // A custom height, hidden flag or outline level: the caller's record.
        Some(h) => buf.out.extend_from_slice(h),
        None => {
            let last_col = row_len.saturating_sub(1) as u32;
            buf.out.extend_from_slice(&[0x00, 0x19]);
            buf.out.extend_from_slice(&row_idx.to_le_bytes());
            buf.out.extend_from_slice(&[0; 4]);
            buf.out.extend_from_slice(&0x12Cu32.to_le_bytes());
            buf.out.extend_from_slice(&[0x00, 0x01, 0x00, 0x00, 0x00]);
            buf.out.extend_from_slice(&[0; 4]);
            buf.out.extend_from_slice(&last_col.to_le_bytes());
        }
    }

    for col_idx in 0..row_len {
        // Python code run below (hashing, conversions) can shrink a list row.
        let Some(mut cell) = row.get(col_idx) else {
            return Ok(None);
        };
        let mut style = styles.for_column(col_idx)?;
        let col = col_idx as u32;
        let mut explicit_fmt = false; // a (value, "format") cell keeps its format
        if cell.is_instance_of::<PyTuple>() {
            let kind = fmt_tuple_kind(&cell, style_type.as_ref());
            if kind != 0 {
                explicit_fmt = kind == 1;
                if !cell.is_exact_instance_of::<PyTuple>() {
                    return Ok(None);
                }
                match fmt_xf_value(&cell, fmt_xf, &mut style)? {
                    Some(value) => cell = value,
                    None => return Ok(None),
                }
            }
        }
        // The hook sees the style before a date's variant: a date it is asked
        // about (pandas' NaT) is no date to the Python encoder.
        let hook_style = style;
        if !explicit_fmt && !date_style_xf(&cell, fmt_xf, &mut style)? {
            return Ok(None);
        }
        match value_cell(&mut buf, &cell, strings, col, row_idx, style, decimal)? {
            Value::Done => {}
            Value::Fallback => return Ok(None),
            // A type the encoder does not know. A numpy scalar is the value it
            // holds, as the hook would have said; otherwise the writer's hook
            // encodes it (a formula), or gives the value to write in its place.
            Value::Unknown => {
                if let Some(item) = numpy_item(&cell) {
                    match value_cell(&mut buf, &item, strings, col, row_idx, hook_style, decimal)? {
                        Value::Done => continue,
                        Value::Fallback | Value::Unknown => return Ok(None),
                    }
                }
                match call_hook(hook, &cell, col, hook_style, row_idx)? {
                    Hooked::Fallback => return Ok(None),
                    Hooked::Raw(raw) => buf.out.extend_from_slice(raw.as_bytes()),
                    // Rich text is a shared string with a tuple for its key, written like
                    // any other string cell; inline strings have no table for it.
                    Hooked::SstKey(key) => match strings {
                        Strings::Shared(sst) => {
                            let idx = sst.index_key(&key)?;
                            let style_ref = if hook_style != 0 {
                                hook_style
                            } else if row_idx == 0 {
                                3
                            } else {
                                0
                            };
                            buf.cell([0x07, 0x0C], col, style_ref);
                            buf.out.extend_from_slice(&idx.to_le_bytes());
                        }
                        Strings::Inline => return Ok(None),
                    },
                    Hooked::Value(value, xf) => {
                        match value_cell(&mut buf, &value, strings, col, row_idx, xf, decimal)? {
                            Value::Done => {}
                            Value::Fallback | Value::Unknown => return Ok(None),
                        }
                    }
                }
            }
        }
    }
    Ok(Some(buf.out))
}

/// Encode a single XLSB row with shared strings.  Returns (bytes, sst_uniq,
/// sst_all), or None when the Python encoder must handle the row; then the
/// strings this call added to the table are removed again.
#[pyfunction]
#[pyo3(signature = (row, ss_dict, ss_list, sst_uniq, sst_all, row_idx, style_ids=None, fmt_xf=None, row_header=None, cell_hook=None))]
#[allow(clippy::too_many_arguments)]
pub fn encode_xlsb_row<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
    ss_dict: &Bound<'py, PyAny>,
    ss_list: &Bound<'py, PyAny>,
    sst_uniq: &Bound<'py, PyAny>,
    sst_all: &Bound<'py, PyAny>,
    row_idx: &Bound<'py, PyAny>,
    style_ids: Option<&Bound<'py, PyAny>>,
    fmt_xf: Option<&Bound<'py, PyAny>>,
    row_header: Option<&Bound<'py, PyAny>>,
    cell_hook: Option<&Bound<'py, PyAny>>,
) -> PyResult<Option<(Bound<'py, PyBytes>, u32, u32)>> {
    let (uniq, all, row_idx) = (u32_mask(sst_uniq)?, u32_mask(sst_all)?, u32_mask(row_idx)?);
    let list = ss_list
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("ss_list must be a list"))?;
    let list_before = list.len();
    let dict = ss_dict
        .cast::<PyDict>()
        .map_err(|_| PyTypeError::new_err("ss_dict must be a dict"))?;
    let none = py.None().into_bound(py);
    let mut strings = Strings::Shared(Sst {
        dict,
        list,
        uniq,
        all,
    });
    let result = encode_row(
        row,
        &mut strings,
        row_idx,
        style_ids.unwrap_or(&none),
        fmt_xf.unwrap_or(&none),
        row_header.unwrap_or(&none),
        cell_hook.unwrap_or(&none),
    );
    let Strings::Shared(sst) = strings else {
        unreachable!()
    };
    match result {
        Ok(Some(bytes)) => Ok(Some((PyBytes::new(py, &bytes), sst.uniq, sst.all))),
        other => {
            rollback(dict, list, list_before);
            other.map(|_| None)
        }
    }
}

/// Encode a single XLSB row with inline text cells.  Returns bytes, or None
/// to use the Python encoder.  Takes no shared-string state.
#[pyfunction]
#[pyo3(signature = (row, row_idx, style_ids=None, fmt_xf=None, row_header=None, cell_hook=None))]
pub fn encode_xlsb_row_inline<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
    row_idx: &Bound<'py, PyAny>,
    style_ids: Option<&Bound<'py, PyAny>>,
    fmt_xf: Option<&Bound<'py, PyAny>>,
    row_header: Option<&Bound<'py, PyAny>>,
    cell_hook: Option<&Bound<'py, PyAny>>,
) -> PyResult<Option<Bound<'py, PyBytes>>> {
    let row_idx = u32_mask(row_idx)?;
    let none = py.None().into_bound(py);
    let bytes = encode_row(
        row,
        &mut Strings::Inline,
        row_idx,
        style_ids.unwrap_or(&none),
        fmt_xf.unwrap_or(&none),
        row_header.unwrap_or(&none),
        cell_hook.unwrap_or(&none),
    )?;
    Ok(bytes.map(|b| PyBytes::new(py, &b)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_header_is_record_then_column_and_style() {
        let mut buf = RowBuf { out: Vec::new() };
        buf.cell([0x04, 0x09], 3, 7);
        assert_eq!(buf.out, [0x04, 0x09, 3, 0, 0, 0, 7, 0, 0, 0]);
    }

    #[test]
    fn double_is_a_brt_cell_real() {
        let mut buf = RowBuf { out: Vec::new() };
        buf.double(1, 0, 1.5);
        assert_eq!(&buf.out[..2], &[0x05, 0x10]);
        assert_eq!(&buf.out[2..10], &[1, 0, 0, 0, 0, 0, 0, 0]);
        assert_eq!(f64::from_le_bytes(buf.out[10..18].try_into().unwrap()), 1.5);
    }
}
