//! Streaming XLSB worksheet reader: an iterator of rows (lists of values).
//!
//! Observable behaviour mirrors the Python reference (`BiffReader` + the
//! row builder in reader.py), quirks included: the layout (first column and
//! width) is frozen from the first row with cells, missing rows are yielded
//! as empty rows, cells before any Row record form a phantom row, and
//! undersized records are ignored.  A record cut short by the end of the data
//! raises InvalidFileError.

use pyo3::exceptions::{PyOverflowError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyBytes, PyDict, PyList, PySet, PyString, PyTuple};
use pyo3::{PyTraverseError, PyVisit};

use crate::bytes::{ByteView, le_f64, le_u32, read_rk, read_varint};
use crate::dates::xlsb_date_value;
use crate::model::{self, CellClass};
use crate::text::decode_utf16le;
use crate::types::{error_type, invalid_file_error};

const XLSB_MAX_ROW_INDEX: u32 = 1_048_575;

/// One record header: id, body start and the claimed body length.
pub struct Record {
    pub id: u64,
    pub body: usize,
    pub len: usize,
}

impl Record {
    /// The record's body in `data` (the buffer `next_record` read it from).
    pub fn bytes<'a>(&self, data: &'a [u8]) -> &'a [u8] {
        &data[self.body..self.body + self.len]
    }
}

/// A record cut short: its header or body runs past the end of the data.
#[derive(Debug, PartialEq, Eq)]
pub struct Truncated;

impl From<Truncated> for PyErr {
    fn from(_: Truncated) -> PyErr {
        invalid_file_error("XLSB record is cut short".to_string())
    }
}

/// Next record at `*pos` (advanced past its body).  Callers stop at the end
/// of the data; a header or body cut short there is an error, so a damaged
/// sheet is reported instead of read in part.
#[inline]
pub fn next_record(data: &[u8], pos: &mut usize) -> Result<Record, Truncated> {
    let id = read_varint(data, pos).ok_or(Truncated)?;
    let len = read_varint(data, pos).ok_or(Truncated)? as usize;
    let body = *pos;
    if len > data.len() - body {
        return Err(Truncated);
    }
    *pos = body + len;
    Ok(Record { id, body, len })
}

/// Row index of a BrtRowHdr body (the record's bytes only), like the
/// reference: a body shorter than 4 bytes and row 0xFFFFFFFF both read as -1.
/// A row past Excel's last raises ValueError.
pub fn row_index(body: &[u8]) -> PyResult<i64> {
    match le_u32(body, 0) {
        None | Some(u32::MAX) => Ok(-1),
        Some(raw) if raw > XLSB_MAX_ROW_INDEX => Err(invalid_file_error(format!(
            "XLSB row record has row index {raw}, past the last row ({XLSB_MAX_ROW_INDEX})"
        ))),
        Some(raw) => Ok(i64::from(raw)),
    }
}

pub fn error_text(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#VALUE!",
    }
}

/// One flag per cell style (xf index): 1 when its number format is a date
/// format.  Built once per sheet, so the cell loop does no lookups.
pub fn date_xf_flags(styles: &Bound<'_, PyList>, date_num_fmts: &Bound<'_, PyAny>) -> Vec<u8> {
    styles
        .iter()
        .map(|style| {
            let Ok(style) = style.cast::<PyDict>() else {
                return 0;
            };
            // PyDict_GetItemString semantics: lookup errors count as a miss.
            let Ok(Some(nf)) = style.get_item("numFmtId") else {
                return 0;
            };
            // PySet_Contains; errors (a non-set argument too) count as a miss.
            // SAFETY: both objects are live with the GIL held; PySet_Contains
            // reports a non-set argument or a hashing failure as -1 with an
            // exception set, which the last arm clears.
            match unsafe { pyo3::ffi::PySet_Contains(date_num_fmts.as_ptr(), nf.as_ptr()) } {
                1 => 1,
                0 => 0,
                _ => {
                    PyErr::take(date_num_fmts.py());
                    0
                }
            }
        })
        .collect()
}

/// Number cell value: an int for whole finite doubles, else a float.
pub fn number(py: Python<'_>, value: f64) -> PyResult<Bound<'_, PyAny>> {
    // No epsilon test here: it would turn tiny non-zero doubles into zero.
    if value.is_finite() && value.trunc() == value {
        // SAFETY: PyLong_FromDouble takes any double and returns a new reference,
        // or NULL with an exception set (not reached: `value` is finite).
        unsafe { Bound::from_owned_ptr_or_err(py, pyo3::ffi::PyLong_FromDouble(value)) }
    } else {
        Ok(pyo3::types::PyFloat::new(py, value).into_any())
    }
}

/// Date-formatted cell value: the datetime, or the number when the serial is
/// outside the datetime range (NaN, infinite, past 9999), as Excel shows ####.
fn date_or_number(py: Python<'_>, value: f64, date_1904: bool) -> PyResult<Bound<'_, PyAny>> {
    match xlsb_date_value(py, value, date_1904) {
        Ok(moment) => Ok(moment),
        Err(e)
            if e.is_instance_of::<PyValueError>(py) || e.is_instance_of::<PyOverflowError>(py) =>
        {
            number(py, value)
        }
        Err(e) => Err(e),
    }
}

/// Context of the cell parser: what a sheet's cells are decoded against.
pub struct CellContext<'a, 'py> {
    pub shared_strings: &'a Bound<'py, PyList>,
    pub ss_len: usize,
    pub date_xf: &'a [u8],
    pub date_1904: bool,
}

impl<'py> CellContext<'_, 'py> {
    fn is_date_xf(&self, xf: u32) -> bool {
        xf > 0 && (xf as usize) < self.date_xf.len() && self.date_xf[xf as usize] != 0
    }

    /// Value of a cell record (ids 0x01..=0x0B) whose body lies inside
    /// `data`, or None when the reference ignores the record.  Errors are
    /// fatal, as in the reference (e.g. dates out of range).
    pub fn value(&self, data: &[u8], rec: &Record, xf: u32) -> PyResult<Option<Bound<'py, PyAny>>> {
        let py = self.shared_strings.py();
        let body = rec.body;
        Ok(Some(match rec.id {
            0x01 => py.None().into_bound(py),
            // BrtCellError / BrtFmlaError: an error value, not a blank cell; shorter, not a cell.
            0x03 | 0x0B => {
                if rec.len < 9 {
                    return Ok(None);
                }
                let text = PyString::new(py, &format!("error:{}", error_text(data[body + 8])));
                match error_type(py) {
                    Some(t) => t.call1((text,))?,
                    None => text.into_any(),
                }
            }
            0x02 => {
                if rec.len < 12 {
                    return Ok(None);
                }
                let v = read_rk(data, body + 8);
                if self.is_date_xf(xf) {
                    date_or_number(py, v, self.date_1904)?
                } else {
                    number(py, v)?
                }
            }
            // Only byte value 1 reads as True.
            0x04 | 0x0A => {
                if rec.len < 9 {
                    return Ok(None);
                }
                pyo3::types::PyBool::new(py, data[body + 8] == 1)
                    .to_owned()
                    .into_any()
            }
            0x05 | 0x09 => {
                if rec.len < 16 {
                    return Ok(None);
                }
                let v = le_f64(data, body + 8);
                if self.is_date_xf(xf) {
                    date_or_number(py, v, self.date_1904)?
                } else {
                    number(py, v)?
                }
            }
            // Records shorter than 12 bytes, or whose string overruns the
            // record, are ignored (never clobbering an earlier duplicate).
            0x06 | 0x08 => {
                if rec.len < 12 {
                    return Ok(None);
                }
                let chars = le_u32(data, body + 8).unwrap_or(0) as usize;
                if chars * 2 + 12 > rec.len {
                    return Ok(None);
                }
                decode_utf16le(py, &data[body + 12..body + 12 + chars * 2])?.into_any()
            }
            0x07 => {
                if rec.len < 12 {
                    return Ok(None);
                }
                let idx = le_u32(data, body + 8).unwrap_or(u32::MAX) as usize;
                if idx < self.ss_len {
                    match self.shared_strings.get_item(idx) {
                        Ok(s) => s,
                        Err(_) => py.None().into_bound(py),
                    }
                } else {
                    py.None().into_bound(py)
                }
            }
            _ => py.None().into_bound(py),
        }))
    }
}

/// Whether a cell record counts as a cell: the row reader ignores undersized
/// bodies and overrunning strings, so such a record starts no row either.
fn counts_for_layout(data: &[u8], rec: &Record) -> bool {
    if rec.len < 8 || rec.body.saturating_add(rec.len) > data.len() {
        return false;
    }
    match rec.id {
        0x02 | 0x07 => rec.len >= 12,
        0x03 | 0x04 | 0x0A | 0x0B => rec.len >= 9,
        0x05 | 0x09 => rec.len >= 16,
        0x06 | 0x08 => {
            rec.len >= 12
                && match le_u32(data, rec.body + 8) {
                    Some(chars) => (chars as usize) * 2 + 12 <= rec.len,
                    None => false,
                }
        }
        _ => true,
    }
}

#[pyclass(module = "xlspy._rs_core")]
pub struct XlsbReader {
    view: ByteView,
    pos: usize,
    done: bool,
    shared_strings: Py<PyList>,
    ss_len: usize,
    date_xf: Vec<u8>,
    cur_cells: Vec<Option<Py<PyAny>>>,
    cur_row_num: i64,
    gap_rows_remaining: i64,
    /// The width of the first row with cells, which later rows are padded to, as in
    /// `_xlsx_read.rows_in_place`.
    first_width: Option<usize>,
    is_initialized: bool,
    seen_row_record: bool,
    date_1904: bool,
}

impl XlsbReader {
    /// The row read so far, as long as the first row or its own last cell.
    fn take_row(&mut self, py: Python<'_>) -> PyResult<Py<PyList>> {
        let cells = std::mem::take(&mut self.cur_cells);
        let width = self.first_width.unwrap_or(0).max(cells.len());
        if width > 0 {
            self.first_width.get_or_insert(width);
        }
        let mut cells = cells.into_iter();
        let items = (0..width).map(|_| cells.next().flatten().unwrap_or_else(|| py.None()));
        Ok(PyList::new(py, items)?.unbind())
    }

    /// A row with no record: as long as the first row, which it may come before.
    fn empty_row(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        let width = self.first_width.unwrap_or(0);
        Ok(PyList::new(py, (0..width).map(|_| py.None()))?.unbind())
    }

    fn process_cell(&mut self, py: Python<'_>, rec: &Record) -> PyResult<()> {
        let data = self.view.as_slice();
        // The reference rejects every cell record shorter than the 8-byte
        // column/style header, then any whose body overruns the stream.
        if rec.len < 8 {
            return Ok(());
        }
        let (Some(col), Some(xf)) = (le_u32(data, rec.body), le_u32(data, rec.body + 4)) else {
            return Ok(());
        };
        if rec.body.saturating_add(rec.len) > data.len() {
            return Ok(());
        }
        let col = i64::from(col);
        let shared_strings = self.shared_strings.bind(py);
        let ctx = CellContext {
            shared_strings,
            ss_len: self.ss_len,
            date_xf: &self.date_xf,
            date_1904: self.date_1904,
        };
        let Some(value) = ctx.value(data, rec, xf & 0xFF_FFFF)? else {
            return Ok(());
        };
        // A cell past column XFD is corrupt, as in load_workbook and the
        // Python reader (BiffReader).
        if col >= 16384 {
            return Err(invalid_file_error(
                "XLSB column exceeds Excel limits".to_string(),
            ));
        }

        let idx = col as usize;
        if idx >= self.cur_cells.len() {
            let width = self.first_width.unwrap_or(0).max(idx + 1);
            self.cur_cells
                .reserve(width.saturating_sub(self.cur_cells.len()));
            self.cur_cells.resize_with(idx + 1, || None);
        }
        self.cur_cells[idx] = Some(value.unbind());
        Ok(())
    }

    fn advance(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyList>>> {
        if self.done {
            return Ok(None);
        }
        // On the first call, skip metadata to the first Row record or the first
        // cell record (cells before any Row record form a phantom row).
        if !self.is_initialized {
            let data_len = self.view.as_slice().len();
            while self.pos < data_len {
                let start = self.pos;
                let rec = next_record(self.view.as_slice(), &mut self.pos)?;
                if rec.id == 0x00 {
                    self.cur_row_num = row_index(rec.bytes(self.view.as_slice()))?;
                    // The rows before the first one are there too, and empty.
                    self.gap_rows_remaining = self.cur_row_num;
                    self.seen_row_record = true;
                    self.is_initialized = true;
                    break;
                } else if (0x01..=0x0B).contains(&rec.id)
                    && counts_for_layout(self.view.as_slice(), &rec)
                {
                    // A record the cell reader ignores starts no row either.
                    self.pos = start; // let the loop below consume this cell
                    self.is_initialized = true;
                    break;
                } else if rec.id == 0x92 {
                    break; // BrtEndSheetData: no rows (records after it are not cells)
                }
            }
            if !self.is_initialized {
                self.done = true;
                return Ok(None);
            }
        }
        if self.gap_rows_remaining > 0 {
            self.gap_rows_remaining -= 1;
            return self.empty_row(py).map(Some);
        }
        while self.pos < self.view.as_slice().len() {
            let rec = next_record(self.view.as_slice(), &mut self.pos)?;
            match rec.id {
                0x00 => {
                    let row = self.take_row(py)?;
                    let new_row = row_index(rec.bytes(self.view.as_slice()))?;
                    // No gap rows before the first Row record was seen.
                    self.gap_rows_remaining = if self.seen_row_record {
                        new_row - self.cur_row_num - 1
                    } else {
                        0
                    };
                    self.cur_row_num = new_row;
                    self.seen_row_record = true;
                    return Ok(Some(row));
                }
                0x01..=0x0B => self.process_cell(py, &rec)?,
                0x92 => break,
                _ => {}
            }
        }
        self.done = true;
        if !self.cur_cells.is_empty() || self.cur_row_num >= 0 {
            return self.take_row(py).map(Some);
        }
        Ok(None)
    }
}

#[pymethods]
impl XlsbReader {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyList>>> {
        let result = self.advance(py);
        if result.is_err() {
            self.done = true; // an error ends the iteration, like the reference
        }
        result
    }

    // The collector must see what the reader holds: a shared-string list or a
    // cell that refers back to the iterator would otherwise keep both alive.
    // (The buffer's source is bytes-like and cannot take part in a cycle.)
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.shared_strings)?;
        for cell in self.cur_cells.iter().flatten() {
            visit.call(cell)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        self.done = true;
        self.cur_cells.clear();
        self.ss_len = 0;
        Python::attach(|py| self.shared_strings = PyList::empty(py).unbind());
    }
}

/// Read an XLSB worksheet from raw bytes.  Returns an iterator of rows, each
/// cell at its own column and each row at its own place, as for XLSX.
#[pyfunction]
#[pyo3(signature = (data, shared_strings, styles_list, date_num_fmts, date_1904=None))]
pub fn read_xlsb_worksheet<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
    shared_strings: &Bound<'py, PyAny>,
    styles_list: &Bound<'py, PyAny>,
    date_num_fmts: &Bound<'py, PyAny>,
    date_1904: Option<&Bound<'py, PyAny>>,
) -> PyResult<Bound<'py, XlsbReader>> {
    let view = ByteView::new(data)?;
    let shared_strings = shared_strings
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("shared_strings must be a list"))?;
    let styles_list = styles_list
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("styles_list must be a list"))?;
    let date_num_fmts = date_num_fmts
        .cast::<PySet>()
        .map_err(|_| PyTypeError::new_err("date_num_fmts must be a set"))?;
    let date_1904 = match date_1904 {
        Some(flag) => flag.is_truthy()?,
        None => false,
    };
    let reader = XlsbReader {
        view,
        pos: 0,
        done: false,
        shared_strings: shared_strings.clone().unbind(),
        ss_len: shared_strings.len(),
        date_xf: date_xf_flags(styles_list, date_num_fmts.as_any()),
        cur_cells: Vec::new(),
        cur_row_num: -1,
        gap_rows_remaining: 0,
        first_width: None,
        is_initialized: false,
        seen_row_record: false,
        date_1904,
    };
    Bound::new(py, reader)
}

/// Load an XLSB sheet's data into Cell objects of the workbook model.
///
/// Returns (max_row, max_column, formulas, row_headers, formula_groups,
/// sheet_data_start, sheet_data_end).  Formula cells are returned as
/// (row, kind, record body) for the caller to decode; row headers with
/// custom attributes as (row, body); BrtShrFmla/BrtArrFmla records as
/// (record id, body).
#[pyfunction]
#[allow(clippy::too_many_arguments)]
pub fn xlsb_load_cells<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
    shared_strings: &Bound<'py, PyAny>,
    styles_list: &Bound<'py, PyAny>,
    date_num_fmts: &Bound<'py, PyAny>,
    date_1904: &Bound<'py, PyAny>,
    cell_type: &Bound<'py, PyAny>,
    descriptors: &Bound<'py, PyAny>,
    sheet: &Bound<'py, PyAny>,
    cells_dict: &Bound<'py, PyAny>,
) -> PyResult<Bound<'py, PyTuple>> {
    let view = ByteView::new(data)?;
    let date_1904 = date_1904.is_truthy()?;
    let (true, Ok(cells_dict), Ok(shared_strings), Ok(styles_list)) = (
        cell_type.is_instance_of::<pyo3::types::PyType>(),
        cells_dict.cast::<PyDict>(),
        shared_strings.cast::<PyList>(),
        styles_list.cast::<PyList>(),
    ) else {
        return Err(PyTypeError::new_err(
            "xlsb_load_cells: unexpected arguments",
        ));
    };
    let cell_class = CellClass::new(cell_type, descriptors)?;
    let (formulas, headers, groups) = (PyList::empty(py), PyList::empty(py), PyList::empty(py));
    let intern = |s: &str| PyString::intern(py, s).into_any();
    let (type_n, type_s, type_b, type_e) = (intern("n"), intern("s"), intern("b"), intern("e"));
    let none = py.None().into_bound(py);
    let mut columns: Vec<Option<Bound<'py, PyAny>>> = vec![None; 16384];
    let date_xf = date_xf_flags(styles_list, date_num_fmts);
    let ctx = CellContext {
        shared_strings,
        ss_len: shared_strings.len(),
        date_xf: &date_xf,
        date_1904,
    };

    let data = view.as_slice();
    let mut pos = 0usize;
    let (mut data_start, mut data_end) = (-1i64, -1i64);
    let (mut current_row, mut previous_column) = (-1i64, -1i64);
    let (mut max_row, mut max_column) = (0i64, 0i64);
    let mut row: Option<(Bound<'py, PyAny>, Bound<'py, PyDict>)> = None; // (row key, row dict)

    while pos < data.len() {
        let header_start = pos;
        let Record {
            id: rec_id,
            body,
            len: rec_len,
        } = next_record(data, &mut pos)?;
        if data_start < 0 {
            if rec_id == 0x91 {
                data_start = pos as i64;
            }
            continue;
        }
        match rec_id {
            0x92 => {
                data_end = header_start as i64;
                break;
            }
            0x00 => {
                // A row record too short for its index is ignored (its cells
                // stay in the row before); row 0xFFFFFFFF is no row (its cells
                // are skipped), and one past the last row is refused, as in
                // the row readers.
                if rec_len < 4 {
                    continue;
                }
                current_row = row_index(&data[body..body + rec_len])?;
                previous_column = -1;
                row = None;
                if current_row < 0 {
                    continue;
                }
                // Custom height, hidden flag or outline level.
                if rec_len > 11 && data[body + 11] & 0x7F != 0 {
                    headers.append((
                        current_row + 1,
                        PyBytes::new(py, &data[body..body + rec_len]),
                    ))?;
                }
                continue;
            }
            0x1AA | 0x1AB => {
                groups.append((rec_id, PyBytes::new(py, &data[body..body + rec_len])))?;
                continue;
            }
            _ => {}
        }
        if !(0x01..=0x12).contains(&rec_id) || current_row < 0 {
            continue;
        }
        // Compact records (0x0C..0x12) omit the column: they follow the
        // previous cell.  Viewed 4 bytes earlier they parse like full ones.
        let (kind, column, cell_body, cell_len) = if rec_id >= 0x0C {
            if rec_len < 4 || body < 4 {
                continue;
            }
            (rec_id - 0x0B, previous_column + 1, body - 4, rec_len + 4)
        } else {
            if rec_len < 8 {
                continue;
            }
            (
                rec_id,
                i64::from(le_u32(data, body).unwrap_or(0)),
                body,
                rec_len,
            )
        };
        let xf = le_u32(data, cell_body + 4).unwrap_or(0) & 0xFF_FFFF;
        if !(0..16384).contains(&column) {
            return Err(invalid_file_error(
                "XLSB column exceeds Excel limits".to_string(),
            ));
        }
        previous_column = column;

        if kind >= 0x08 {
            // Formulas are rare: the caller decodes them, from the full record
            // so that compact cells get their column back.
            let mut payload = Vec::with_capacity(cell_len);
            payload.extend_from_slice(&[(column & 0xFF) as u8, ((column >> 8) & 0xFF) as u8, 0, 0]);
            payload.extend_from_slice(&data[cell_body + 4..cell_body + cell_len]);
            formulas.append((current_row + 1, kind, PyBytes::new(py, &payload)))?;
            continue;
        }
        let (value, data_type) = if kind == 0x03 {
            if cell_len < 9 {
                continue;
            }
            (
                PyString::new(py, error_text(data[cell_body + 8])).into_any(),
                &type_e,
            )
        } else {
            let rec = Record {
                id: kind,
                body: cell_body,
                len: cell_len,
            };
            let Some(value) = ctx.value(data, &rec, xf)? else {
                continue;
            };
            let data_type = if value.is_none() {
                if xf == 0 {
                    continue; // only a styled blank is a cell
                }
                &type_n
            } else if value.is_instance_of::<PyString>() {
                &type_s
            } else if value.is_instance_of::<PyBool>() {
                &type_b
            } else {
                &type_n
            };
            (value, data_type)
        };

        if row.is_none() {
            let key = (current_row + 1).into_pyobject(py)?.into_any();
            let dict = match cells_dict.get_item(&key)? {
                Some(d) => model::as_row_dict(d)?,
                None => {
                    let d = PyDict::new(py);
                    cells_dict.set_item(&key, &d)?;
                    d
                }
            };
            row = Some((key, dict));
        }
        let (row_key, row_dict) = row.as_ref().expect("the row dict was set just above");
        let col_obj = match &mut columns[column as usize] {
            Some(c) => &*c,
            slot @ None => &*slot.insert((column + 1).into_pyobject(py)?.into_any()),
        };
        let style = xf.into_pyobject(py)?.into_any();
        let cell = cell_class.new_cell(
            py,
            [
                Some(sheet),
                Some(row_key),
                Some(col_obj),
                Some(&value),
                Some(&none),
                Some(data_type),
                Some(&style),
                Some(&none),
                Some(&none),
                Some(&none),
            ],
        )?;
        let stored = row_dict.set_item(col_obj, &cell);
        if !value.is_none() || xf != 0 || !data_type.is(&type_n) {
            max_row = max_row.max(current_row + 1);
            max_column = max_column.max(column + 1);
        }
        stored?;
    }
    if data_start >= 0 && data_end < 0 {
        data_end = data.len() as i64;
    }
    PyTuple::new(
        py,
        [
            max_row.into_pyobject(py)?.into_any(),
            max_column.into_pyobject(py)?.into_any(),
            formulas.into_any(),
            headers.into_any(),
            groups.into_any(),
            data_start.into_pyobject(py)?.into_any(),
            data_end.into_pyobject(py)?.into_any(),
        ],
    )
}

/// A range record's `r1, r2, c1, c2`.
fn range_of(data: &[u8], rec: &Record) -> Option<[u32; 4]> {
    let body = rec.bytes(data);
    if body.len() < 16 {
        return None;
    }
    Some([
        le_u32(body, 0)?,
        le_u32(body, 4)?,
        le_u32(body, 8)?,
        le_u32(body, 12)?,
    ])
}

/// One pass over an XLSB worksheet for `XlsbUpdater.update_cells`, the twin of
/// `_scan_sheet` in xlspy/_xlsb_edit.py: `(begin start, begin end, end start,
/// [(row record start, row)], merged ranges, formula group ranges, (dimension
/// start, its end, its range) or None, (first column, first row, last column,
/// last row) of the cells or None)`.
///
/// None for anything unusual (a record cut short, a body too short for what it
/// holds, no sheet data): Python then decides, and its error is the one raised.
#[pyfunction]
pub fn xlsb_sheet_scan<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
) -> PyResult<Option<Bound<'py, PyTuple>>> {
    let view = ByteView::new(data)?;
    let data = view.as_slice();
    let mut pos = 0usize;
    let mut begin: Option<(usize, usize)> = None;
    let mut end: Option<usize> = None;
    let mut rows: Vec<(usize, u32)> = Vec::new();
    let mut merges: Vec<[u32; 4]> = Vec::new();
    let mut groups: Vec<[u32; 4]> = Vec::new();
    let mut dimension: Option<(usize, usize, [u32; 4])> = None;
    let mut current_row: Option<u32> = None;
    let mut previous_column = -1i64;
    let mut bounds: Option<[i64; 4]> = None;
    while pos < data.len() {
        let start = pos;
        let Ok(rec) = next_record(data, &mut pos) else {
            return Ok(None);
        };
        let in_data = begin.is_some() && end.is_none();
        match rec.id {
            0x00 if in_data => {
                let Some(row) = le_u32(rec.bytes(data), 0) else {
                    return Ok(None);
                };
                rows.push((start, row));
                current_row = Some(row);
                previous_column = -1;
            }
            0x01..=0x12 if in_data => {
                let Some(row) = current_row else { continue };
                let column = if rec.id >= 0x0C {
                    previous_column + 1
                } else {
                    let Some(column) = le_u32(rec.bytes(data), 0) else {
                        return Ok(None);
                    };
                    i64::from(column)
                };
                previous_column = column;
                if column < 16384 {
                    let row = i64::from(row);
                    bounds = Some(match bounds {
                        None => [column, row, column, row],
                        Some([c1, r1, c2, r2]) => {
                            [c1.min(column), r1.min(row), c2.max(column), r2.max(row)]
                        }
                    });
                }
            }
            0x91 if begin.is_none() => begin = Some((start, pos)),
            0x92 if end.is_none() => end = Some(start),
            0x94 if dimension.is_none() => {
                if rec.len != 16 {
                    return Ok(None);
                }
                let Some(range) = range_of(data, &rec) else {
                    return Ok(None);
                };
                dimension = Some((start, pos, range));
            }
            0xB0 => {
                let Some(range) = range_of(data, &rec) else {
                    return Ok(None);
                };
                merges.push(range);
            }
            0x1AA..=0x1AC => {
                let Some(range) = range_of(data, &rec) else {
                    return Ok(None);
                };
                groups.push(range);
            }
            _ => {}
        }
    }
    let (Some((begin_start, begin_end)), Some(end_start)) = (begin, end) else {
        return Ok(None);
    };
    let rows = PyList::new(py, rows)?;
    let merges = PyList::new(py, merges.iter().map(|r| (r[0], r[1], r[2], r[3])))?;
    let groups = PyList::new(py, groups.iter().map(|r| (r[0], r[1], r[2], r[3])))?;
    let dimension = dimension.map(|(s, e, r)| (s, e, (r[0], r[1], r[2], r[3])));
    let bounds = bounds.map(|b| (b[0], b[1], b[2], b[3]));
    let result = (
        begin_start,
        begin_end,
        end_start,
        rows,
        merges,
        groups,
        dimension,
        bounds,
    );
    result.into_pyobject(py).map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A record: varint id, varint length, body.
    fn record(id: u8, body: &[u8]) -> Vec<u8> {
        let mut out = vec![id, body.len() as u8];
        out.extend_from_slice(body);
        out
    }

    #[test]
    fn next_record_walks_records_and_reports_a_truncated_header() {
        let mut data = record(0x00, &[1, 0, 0, 0]);
        data.extend(record(0x07, &[0; 12]));
        data.push(0x80); // a header cut inside its id varint
        let mut pos = 0;
        let first = next_record(&data, &mut pos).unwrap();
        assert_eq!((first.id, first.body, first.len), (0x00, 2, 4));
        let second = next_record(&data, &mut pos).unwrap();
        assert_eq!((second.id, second.len), (0x07, 12));
        assert_eq!(pos, data.len() - 1);
        assert_eq!(next_record(&data, &mut pos).err(), Some(Truncated));
    }

    #[test]
    fn next_record_reports_a_body_that_overruns_the_data() {
        let data = [0x07, 20, 0, 0];
        let mut pos = 0;
        assert_eq!(next_record(&data, &mut pos).err(), Some(Truncated));
        // A header whose length varint is missing is cut short too.
        let mut pos = 0;
        assert_eq!(next_record(&[0x07], &mut pos).err(), Some(Truncated));
    }

    #[test]
    fn row_index_of_short_and_sentinel_bodies_is_minus_one() {
        assert_eq!(row_index(&[1, 0]).unwrap(), -1);
        assert_eq!(row_index(&[0xFF; 4]).unwrap(), -1);
        assert_eq!(row_index(&[5, 0, 0, 0]).unwrap(), 5);
        assert_eq!(
            row_index(&XLSB_MAX_ROW_INDEX.to_le_bytes()).unwrap(),
            i64::from(XLSB_MAX_ROW_INDEX)
        );
    }

    #[test]
    fn a_short_row_record_does_not_read_the_next_record_as_its_index() {
        // A 2-byte BrtRowHdr followed by a record whose bytes would make a
        // row index past the last row if read as part of it.
        let mut data = record(0x00, &[1, 0]);
        data.extend(record(0x26, &[]));
        data.extend(record(0x07, &[0; 12]));
        let mut pos = 0;
        let rec = next_record(&data, &mut pos).unwrap();
        assert_eq!(rec.bytes(&data), &[1, 0]);
        assert_eq!(row_index(rec.bytes(&data)).unwrap(), -1);
    }

    #[test]
    fn only_records_the_cell_reader_reads_count_for_the_layout() {
        let counts = |id: u8, body: &[u8]| {
            let data = record(id, body);
            let mut pos = 0;
            let rec = next_record(&data, &mut pos).unwrap();
            counts_for_layout(&data, &rec)
        };
        // Blank needs the header alone; error, bool and the other kinds their value.
        assert!(counts(0x01, &[0; 8]));
        assert!(!counts(0x01, &[0; 7]));
        assert!(counts(0x03, &[0; 9]));
        assert!(!counts(0x03, &[0; 8]));
        assert!(counts(0x0B, &[0; 9]));
        assert!(!counts(0x0B, &[0; 8]));
        assert!(counts(0x02, &[0; 12]));
        assert!(!counts(0x02, &[0; 11]));
        assert!(counts(0x04, &[0; 9]));
        assert!(!counts(0x04, &[0; 8]));
        assert!(counts(0x05, &[0; 16]));
        assert!(!counts(0x05, &[0; 15]));
        assert!(counts(0x07, &[0; 12]));
        assert!(!counts(0x07, &[0; 11]));
        // A string fits its record, or the record is ignored.
        let mut text = vec![0u8; 12];
        text[8] = 2;
        assert!(!counts(0x06, &text));
        text.extend_from_slice(&[b'a', 0, b'b', 0]);
        assert!(counts(0x06, &text));
    }

    #[test]
    fn error_codes_map_to_excel_text() {
        assert_eq!(error_text(0x07), "#DIV/0!");
        assert_eq!(error_text(0x2A), "#N/A");
        assert_eq!(error_text(0x99), "#VALUE!");
    }

    #[test]
    fn layout_counts_only_complete_cell_records() {
        let check = |id: u8, body: &[u8]| {
            let data = record(id, body);
            let mut pos = 0;
            let rec = next_record(&data, &mut pos).unwrap();
            counts_for_layout(&data, &rec)
        };
        assert!(!check(0x05, &[0; 7])); // under the 8-byte column/style header
        assert!(check(0x05, &[0; 16])); // BrtCellReal
        assert!(!check(0x05, &[0; 15]));
        assert!(check(0x04, &[0; 9])); // BrtCellBool
        let mut text = vec![0; 8];
        text.extend_from_slice(&2u32.to_le_bytes());
        text.extend_from_slice(&[b'a', 0, b'b', 0]);
        assert!(check(0x06, &text)); // BrtCellSt, two UTF-16 units
        text.truncate(text.len() - 1);
        assert!(!check(0x06, &text)); // the string overruns the record
        assert!(check(0x01, &[0; 8])); // BrtCellBlank
    }
}
