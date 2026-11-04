//! Rows of `XlsxUpdater` / `XlsbUpdater.replace_sheet_data`, cell for cell as
//! the updaters' Python code writes them (`_value_cell`, `_value_cell_bytes`).
//!
//! The common cells (text, numbers, booleans, dates) are written here; any
//! other cell goes to the updater's own Python function, one cell at a time,
//! so nothing a row writes is ever written twice and every error is raised
//! where the Python code raises it. Text goes through the updater's
//! shared-string index (`string_index`), which counts every use.

use pyo3::exceptions::{PyOSError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDate, PyDateAccess, PyDateTime, PyFloat, PyInt, PyString};
use pyo3::types::{PyTimeAccess, PyTuple};
use std::fs::OpenOptions;
use std::io::{BufReader, Read, Seek, SeekFrom, Write};

use crate::bytes::write_varint;
use crate::cells::text_fits;
use crate::dates::days_from_civil;
use crate::numfmt::{write_repr, write_u32};
use crate::pyseq::FastSeq;
use crate::text::{CELL_TEXT_LIMIT, CodePoints};
use crate::types::int_as_i64;

/// Excel's 1900 epoch (1899-12-30) and the one before 1900-03-01
/// (1899-12-31, for the fictitious 1900-02-29), as proleptic ordinals.
const EPOCH_1900: i64 = 693_594;
const EPOCH_1900_EARLY: i64 = 693_595;
/// 1904-01-01.
const EPOCH_1904: i64 = 695_056;
/// Python's `date.toordinal()` of 1970-01-01, the zero of `days_from_civil`.
const UNIX_ORDINAL: i64 = 719_163;

/// The cell's value: `(value, "format")` wrappers unwrapped, as `unwrap_cell`.
fn unwrapped<'py>(cell: &Bound<'py, PyAny>) -> PyResult<Bound<'py, PyAny>> {
    if let Ok(pair) = cell.cast::<PyTuple>()
        && pair.len() == 2
        && pair.get_item(1)?.is_instance_of::<PyString>()
    {
        return pair.get_item(0);
    }
    Ok(cell.clone())
}

/// `datetime_to_serial(value.replace(tzinfo=None), date_1904)` for an exact
/// date or datetime, bit for bit; None for anything else (the Python code).
fn serial(cell: &Bound<'_, PyAny>, date_1904: bool) -> Option<f64> {
    let is_dt = cell.is_exact_instance_of::<PyDateTime>();
    if !is_dt && !cell.is_exact_instance_of::<PyDate>() {
        return None;
    }
    // SAFETY: checked just above: an exact date or datetime (a date subclass).
    let date = unsafe { cell.cast_unchecked::<PyDate>() };
    let (y, mo, d) = (
        i64::from(date.get_year()),
        i64::from(date.get_month()),
        i64::from(date.get_day()),
    );
    let (mut hh, mut mi, mut ss, mut us) = (0i64, 0i64, 0i64, 0i64);
    if is_dt {
        // SAFETY: `is_dt` checked for an exact datetime.
        let dt = unsafe { cell.cast_unchecked::<PyDateTime>() };
        hh = i64::from(dt.get_hour());
        mi = i64::from(dt.get_minute());
        ss = i64::from(dt.get_second());
        us = i64::from(dt.get_microsecond());
    }
    let ordinal = days_from_civil(y, mo, d) + UNIX_ORDINAL;
    let epoch = if date_1904 {
        EPOCH_1904
    } else if (y, mo) < (1900, 3) {
        EPOCH_1900_EARLY
    } else {
        EPOCH_1900
    };
    let total_us = ((ordinal - epoch) * 86_400 + hh * 3600 + mi * 60 + ss) * 1_000_000 + us;
    // Within 2**53 the int division Python does is this float division.
    if !(-9_007_199_254_740_992..=9_007_199_254_740_992).contains(&total_us) {
        return None;
    }
    Some(total_us as f64 / 1_000_000.0 / 86_400.0)
}

/// The shortest digits that read back as `v`, without a trailing `.0`.
fn write_number(py: Python<'_>, out: &mut Vec<u8>, v: f64) -> PyResult<()> {
    write_repr(py, out, v)?;
    if out.ends_with(b".0") {
        out.truncate(out.len() - 2);
    }
    Ok(())
}

/// Column letters of a 0-based column (0: A).
fn write_letters(out: &mut Vec<u8>, column: u32) {
    let mut n = column + 1;
    let mut letters = [0u8; 4];
    let mut len = 0;
    while n > 0 {
        let rem = (n - 1) % 26;
        letters[len] = b'A' + rem as u8;
        len += 1;
        n = (n - 1) / 26;
    }
    out.extend(letters[..len].iter().rev());
}

/// The column's xf from a list of xfs (missing: 0).
fn column_style(styles: &[u32], column: usize) -> u32 {
    styles.get(column).copied().unwrap_or(0)
}

fn style_list(styles: &Bound<'_, PyAny>) -> PyResult<Vec<u32>> {
    styles
        .extract::<Vec<u32>>()
        .map_err(|_| PyTypeError::new_err("styles must be a list of xf ids"))
}

fn finite_float(value: &Bound<'_, PyAny>) -> Option<f64> {
    if value.is_exact_instance_of::<PyFloat>() {
        // SAFETY: an exact float.
        let v = unsafe { value.cast_unchecked::<PyFloat>() }.value();
        return v.is_finite().then_some(v);
    }
    None
}

/// One `<row>` of `XlsxUpdater`: `<{prefix}row r="n">` and its cells, as
/// `_row_xml` of `_value_cell` cells. `styles` are the columns' xfs (empty for
/// `style_fallback="general"`), `date_style` the xf for a date in a column
/// without one (0: none). `string_index(text)` is the shared-string index, or
/// None for inline text (then text cells go to `cell_xml`). `cell_xml(cell,
/// column)` writes any other cell.
#[pyfunction]
#[pyo3(signature = (row, row_number, styles, date_style, prefix, string_index, cell_xml, date_1904))]
#[allow(clippy::too_many_arguments)]
pub fn encode_update_xlsx_row<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
    row_number: u32,
    styles: &Bound<'py, PyAny>,
    date_style: u32,
    prefix: &str,
    string_index: &Bound<'py, PyAny>,
    cell_xml: &Bound<'py, PyAny>,
    date_1904: bool,
) -> PyResult<Bound<'py, PyBytes>> {
    let seq = FastSeq::new(row, c"row must be a sequence")?;
    let styles = style_list(styles)?;
    let prefix = prefix.as_bytes();
    let mut out = Vec::with_capacity(32 + seq.len() * 24);
    out.push(b'<');
    out.extend_from_slice(prefix);
    out.extend_from_slice(b"row r=\"");
    write_u32(&mut out, row_number);
    out.extend_from_slice(b"\">");
    let mut i = 0;
    while let Some(raw) = seq.get(i) {
        let column = i;
        i += 1;
        let value = unwrapped(&raw)?;
        if value.is_none() {
            continue;
        }
        let mut style = column_style(&styles, column);
        let open = |out: &mut Vec<u8>, style: u32| {
            out.push(b'<');
            out.extend_from_slice(prefix);
            out.extend_from_slice(b"c r=\"");
            write_letters(out, column as u32);
            write_u32(out, row_number);
            out.push(b'"');
            if style > 0 {
                out.extend_from_slice(b" s=\"");
                write_u32(out, style);
                out.push(b'"');
            }
        };
        let value_open = |out: &mut Vec<u8>| {
            out.extend_from_slice(b"><");
            out.extend_from_slice(prefix);
            out.extend_from_slice(b"v>");
        };
        let value_close = |out: &mut Vec<u8>| {
            out.extend_from_slice(b"</");
            out.extend_from_slice(prefix);
            out.extend_from_slice(b"v></");
            out.extend_from_slice(prefix);
            out.extend_from_slice(b"c>");
        };
        if let Ok(b) = value.cast_exact::<PyBool>() {
            open(&mut out, style);
            out.extend_from_slice(b" t=\"b\"");
            value_open(&mut out);
            out.push(if b.is_true() { b'1' } else { b'0' });
            value_close(&mut out);
            continue;
        }
        let number = if value.is_exact_instance_of::<PyInt>() {
            int_as_i64(&value).map(|v| v as f64)
        } else {
            finite_float(&value)
        };
        if let Some(v) = number {
            open(&mut out, style);
            value_open(&mut out);
            write_number(py, &mut out, v)?;
            value_close(&mut out);
            continue;
        }
        if let Some(v) = serial(&value, date_1904) {
            if style == 0 {
                style = date_style;
            }
            open(&mut out, style);
            value_open(&mut out);
            write_number(py, &mut out, v)?;
            value_close(&mut out);
            continue;
        }
        if !string_index.is_none()
            && let Ok(text) = value.cast_exact::<PyString>()
            && text_fits(text)?
        {
            let index: u32 = string_index.call1((text,))?.extract()?;
            open(&mut out, style);
            out.extend_from_slice(b" t=\"s\"");
            value_open(&mut out);
            write_u32(&mut out, index);
            value_close(&mut out);
            continue;
        }
        // Formulas, other types, inline text, over-long text (which raises).
        let xml = cell_xml.call1((&raw, column))?;
        let xml = xml
            .cast::<PyString>()
            .map_err(|_| PyTypeError::new_err("cell_xml must return a str"))?;
        out.extend_from_slice(xml.to_str()?.as_bytes());
    }
    out.extend_from_slice(b"</");
    out.extend_from_slice(prefix);
    out.extend_from_slice(b"row>");
    Ok(PyBytes::new(py, &out))
}

/// Append one BIFF12 record: id, size and payload.
fn record(out: &mut Vec<u8>, id: u8, payload: &[u8]) {
    out.push(id);
    write_varint(out, payload.len() as u32);
    out.extend_from_slice(payload);
}

fn cell_header(column: u32, style: u32) -> [u8; 8] {
    let mut header = [0u8; 8];
    header[..4].copy_from_slice(&column.to_le_bytes());
    header[4..].copy_from_slice(&style.to_le_bytes());
    header
}

/// One row of `XlsbUpdater`: its BrtRowHdr (row `row_number`, last column
/// `last_col`) and the records of its cells, as `_value_cell_bytes` writes
/// them. `styles` are the columns' xfs (empty for `style_fallback="general"`),
/// `date_style` the xf for a date in a column without one. `string_index` is
/// the shared-string index, or None for inline text; `cell_bytes(cell,
/// column)` writes any other cell. None cells get no record.
#[pyfunction]
#[pyo3(signature = (row, row_number, last_col, styles, date_style, string_index, cell_bytes, date_1904))]
#[allow(clippy::too_many_arguments)]
pub fn encode_update_xlsb_row<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
    row_number: u32,
    last_col: u32,
    styles: &Bound<'py, PyAny>,
    date_style: u32,
    string_index: &Bound<'py, PyAny>,
    cell_bytes: &Bound<'py, PyAny>,
    date_1904: bool,
) -> PyResult<Bound<'py, PyBytes>> {
    let seq = FastSeq::new(row, c"row must be a sequence")?;
    let styles = style_list(styles)?;
    let mut out = Vec::with_capacity(32 + seq.len() * 16);
    let mut header = [0u8; 25];
    header[..4].copy_from_slice(&row_number.to_le_bytes());
    header[8..12].copy_from_slice(&300u32.to_le_bytes());
    header[13] = 1;
    header[21..25].copy_from_slice(&last_col.to_le_bytes());
    record(&mut out, 0x00, &header);
    let mut payload = Vec::with_capacity(64);
    let mut i = 0;
    while let Some(raw) = seq.get(i) {
        let column = i as u32;
        i += 1;
        if raw.is_none() {
            continue;
        }
        let value = unwrapped(&raw)?;
        let style = column_style(&styles, column as usize);
        payload.clear();
        payload.extend_from_slice(&cell_header(column, style));
        if let Ok(b) = value.cast_exact::<PyBool>() {
            payload.push(u8::from(b.is_true()));
            record(&mut out, 0x04, &payload);
            continue;
        }
        if value.is_exact_instance_of::<PyInt>()
            && let Some(v) = int_as_i64(&value)
        {
            if (-536_870_912..=536_870_911).contains(&v) {
                payload.extend_from_slice(&((((v as i32) << 2) | 2) as u32).to_le_bytes());
                record(&mut out, 0x02, &payload);
            } else {
                payload.extend_from_slice(&(v as f64).to_le_bytes());
                record(&mut out, 0x05, &payload);
            }
            continue;
        }
        if let Some(v) = finite_float(&value) {
            payload.extend_from_slice(&v.to_le_bytes());
            record(&mut out, 0x05, &payload);
            continue;
        }
        if let Some(v) = serial(&value, date_1904) {
            let style = if style > 0 { style } else { date_style };
            payload.clear();
            payload.extend_from_slice(&cell_header(column, style));
            payload.extend_from_slice(&v.to_le_bytes());
            record(&mut out, 0x05, &payload);
            continue;
        }
        if let Ok(text) = value.cast_exact::<PyString>() {
            if string_index.is_none() {
                let points = CodePoints::of(text)?;
                let units = points.utf16_len();
                if points.first_surrogate().is_none() && units <= CELL_TEXT_LIMIT {
                    payload.extend_from_slice(&(units as u32).to_le_bytes());
                    points.write_utf16le(&mut payload);
                    record(&mut out, 0x06, &payload);
                    continue;
                }
            } else if text_fits(text)? {
                let index: u32 = string_index.call1((text,))?.extract()?;
                payload.extend_from_slice(&index.to_le_bytes());
                record(&mut out, 0x07, &payload);
                continue;
            }
        }
        // Formulas (refused), numpy and other types, over-long text (raises).
        let encoded = cell_bytes.call1((&raw, column))?;
        let encoded = encoded
            .cast::<PyBytes>()
            .map_err(|_| PyTypeError::new_err("cell_bytes must return bytes"))?;
        out.extend_from_slice(encoded.as_bytes());
    }
    Ok(PyBytes::new(py, &out))
}

/// One BIFF12 variable-length integer from `reader`; None at the end of the file.
fn read_vlq(reader: &mut impl Read) -> PyResult<Option<(u64, u64)>> {
    let (mut value, mut shift, mut read) = (0u64, 0u32, 0u64);
    let mut byte = [0u8; 1];
    loop {
        match reader.read(&mut byte) {
            Ok(0) if read == 0 => return Ok(None),
            Ok(0) => {
                return Err(PyValueError::new_err(
                    "Truncated BIFF12 variable-length integer",
                ));
            }
            Ok(_) => {}
            Err(e) => return Err(PyOSError::new_err(e.to_string())),
        }
        read += 1;
        value |= u64::from(byte[0] & 0x7F) << shift;
        if byte[0] & 0x80 == 0 {
            return Ok(Some((value, read)));
        }
        shift += 7;
        if shift > 63 {
            return Err(PyValueError::new_err(
                "BIFF12 variable-length integer is too long",
            ));
        }
    }
}

/// Set the last column of every BrtRowHdr in the rows file `path` (records the
/// streaming `XlsbUpdater` wrote with 0 there), reading it once in a stream.
#[pyfunction]
pub fn patch_xlsb_row_headers(path: std::path::PathBuf, last_col: u32) -> PyResult<()> {
    let io = |e: std::io::Error| PyOSError::new_err(e.to_string());
    let mut writer = OpenOptions::new().write(true).open(&path).map_err(io)?;
    let mut reader = BufReader::with_capacity(1 << 20, std::fs::File::open(&path).map_err(io)?);
    let mut offset = 0u64;
    loop {
        let Some((id, id_len)) = read_vlq(&mut reader)? else {
            return Ok(());
        };
        let Some((length, length_len)) = read_vlq(&mut reader)? else {
            return Err(PyValueError::new_err("Truncated BIFF12 record length"));
        };
        let data_start = offset + id_len + length_len;
        if id == 0 && length >= 25 {
            writer.seek(SeekFrom::Start(data_start + 21)).map_err(io)?;
            writer.write_all(&last_col.to_le_bytes()).map_err(io)?;
        }
        let skip =
            i64::try_from(length).map_err(|_| PyValueError::new_err("BIFF12 record too long"))?;
        reader.seek_relative(skip).map_err(io)?;
        offset = data_start + length;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn letters(column: u32) -> String {
        let mut out = Vec::new();
        write_letters(&mut out, column);
        String::from_utf8(out).unwrap()
    }

    #[test]
    fn columns_are_named_as_excel_names_them() {
        assert_eq!(letters(0), "A");
        assert_eq!(letters(25), "Z");
        assert_eq!(letters(26), "AA");
        assert_eq!(letters(701), "ZZ");
        assert_eq!(letters(702), "AAA");
        assert_eq!(letters(16_383), "XFD");
    }

    #[test]
    fn epochs_are_the_ordinals_python_gives() {
        // date(1899, 12, 30), date(1899, 12, 31) and date(1904, 1, 1).toordinal()
        assert_eq!(days_from_civil(1899, 12, 30) + UNIX_ORDINAL, EPOCH_1900);
        assert_eq!(
            days_from_civil(1899, 12, 31) + UNIX_ORDINAL,
            EPOCH_1900_EARLY
        );
        assert_eq!(days_from_civil(1904, 1, 1) + UNIX_ORDINAL, EPOCH_1904);
    }
}
