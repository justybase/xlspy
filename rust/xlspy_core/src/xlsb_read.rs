//! Streaming XLSB worksheet reader: an iterator of rows (lists of values).
//!
//! Observable behaviour mirrors the Python reference (`BiffReader` + the
//! row builder in reader.py), quirks included: the layout (first column and
//! width) is frozen from the first row with cells, missing rows are yielded
//! as empty rows, cells before any Row record form a phantom row, and
//! undersized or overrunning records are ignored.

use pyo3::exceptions::{PyMemoryError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PySet, PyString};

use crate::bytes::{ByteView, le_f64, le_u32, read_rk, read_varint};
use crate::dates::xlsb_date_value;
use crate::text::decode_utf16le;
use crate::types::error_type;

const XLSB_MAX_ROW_INDEX: u32 = 1_048_575;
/// Bound of the dense per-row cell array.  The format maximum is 16384
/// columns; wider nonstandard layouts are left to the Python reference.
pub const MAX_COL_SPAN: i64 = 1_048_576;

/// One record header: id, body start and the claimed body length.
pub struct Record {
    pub id: u64,
    pub body: usize,
    pub len: usize,
}

/// Next record at `*pos` (advanced past its body), or None at a truncated
/// header.  The body may overrun the data; callers check.
#[inline]
pub fn next_record(data: &[u8], pos: &mut usize) -> Option<Record> {
    let id = read_varint(data, pos)?;
    let len = read_varint(data, pos)? as usize;
    let body = *pos;
    *pos = body.saturating_add(len);
    Some(Record { id, body, len })
}

/// Row index of a BrtRowHdr body, like the reference: a short body and row
/// 0xFFFFFFFF both read as -1.  A row past Excel's last raises ValueError.
pub fn row_index(data: &[u8], pos: usize) -> PyResult<i64> {
    match le_u32(data, pos) {
        None | Some(u32::MAX) => Ok(-1),
        Some(raw) if raw > XLSB_MAX_ROW_INDEX => Err(PyValueError::new_err(format!(
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
pub fn date_xf_flags(styles: &Bound<'_, PyList>, date_num_fmts: &Bound<'_, PySet>) -> Vec<u8> {
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
            u8::from(date_num_fmts.contains(nf).unwrap_or(false))
        })
        .collect()
}

/// Number cell value: an int for whole finite doubles, else a float.
pub fn number(py: Python<'_>, value: f64) -> PyResult<Bound<'_, PyAny>> {
    // No epsilon test here: it would turn tiny non-zero doubles into zero.
    if value.is_finite() && value.trunc() == value {
        unsafe { Bound::from_owned_ptr_or_err(py, pyo3::ffi::PyLong_FromDouble(value)) }
    } else {
        Ok(pyo3::types::PyFloat::new(py, value).into_any())
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
            // BrtCellError / BrtFmlaError: an error value, not a blank cell.
            0x03 | 0x0B => {
                if rec.len < 9 {
                    py.None().into_bound(py)
                } else {
                    let text = PyString::new(py, &format!("error:{}", error_text(data[body + 8])));
                    match error_type(py) {
                        Some(t) => t.call1((text,))?,
                        None => text.into_any(),
                    }
                }
            }
            0x02 => {
                if rec.len < 12 {
                    return Ok(None);
                }
                let v = read_rk(data, body + 8);
                if self.is_date_xf(xf) {
                    xlsb_date_value(py, v, self.date_1904)?
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
                    xlsb_date_value(py, v, self.date_1904)?
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

/// Whether a cell record counts as a cell for the layout: the reference
/// ignores undersized bodies and overrunning strings, and so must the prescan.
fn counts_for_layout(data: &[u8], rec: &Record) -> bool {
    if rec.len < 8 || rec.body.saturating_add(rec.len) > data.len() {
        return false;
    }
    match rec.id {
        0x02 | 0x07 => rec.len >= 12,
        0x04 | 0x0A => rec.len >= 9,
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
    /// Layout frozen from the first row with cells; -1 means unset, and the
    /// first streamed cell then freezes `first_col`, like the reference.
    first_col: i64,
    cols_per_row: i64,
    is_initialized: bool,
    seen_row_record: bool,
    date_1904: bool,
}

impl XlsbReader {
    /// Determine the column layout of the first row that contains a cell,
    /// so empty leading rows get the data width.  Returns false when the
    /// layout is too wide for the dense row array (the caller uses Python).
    fn prescan(&mut self) -> bool {
        let data = self.view.as_slice();
        let mut pos = 0usize;
        let mut cur_row = -1i64;
        let mut span_row = -2i64; // row of the first row with cells (-2: none)
        let (mut first_c, mut last_c) = (-1i64, -1i64);
        while pos < data.len() {
            let Some(rec) = next_record(data, &mut pos) else {
                break;
            };
            match rec.id {
                0x00 => {
                    if span_row != -2 && cur_row != span_row {
                        break;
                    }
                    // A bad row index raises later, when iteration reaches it.
                    let Ok(row) = row_index(data, rec.body) else {
                        break;
                    };
                    cur_row = row;
                }
                0x01..=0x0B => {
                    if !counts_for_layout(data, &rec) {
                        continue;
                    }
                    if span_row == -2 {
                        span_row = cur_row;
                    }
                    if cur_row == span_row
                        && let Some(col) = le_u32(data, rec.body)
                    {
                        let col = i64::from(col);
                        if first_c < 0 {
                            first_c = col;
                        }
                        last_c = col;
                    }
                }
                0x92 => break,
                _ => {}
            }
        }
        if span_row != -2 && first_c >= 0 {
            self.first_col = first_c;
            // May be <= 0 for unordered first rows: the reference then yields
            // empty rows, which build_row reproduces.
            self.cols_per_row = last_c - first_c + 1;
            if self.cols_per_row > MAX_COL_SPAN {
                return false;
            }
        }
        true
    }

    fn take_row(&mut self, py: Python<'_>) -> PyResult<Py<PyList>> {
        let cells = std::mem::take(&mut self.cur_cells);
        let out_cols = if self.cols_per_row == -1 {
            cells.len().max(1) as i64
        } else {
            self.cols_per_row.max(0)
        };
        if out_cols > MAX_COL_SPAN {
            return Err(PyMemoryError::new_err(()));
        }
        let mut cells = cells.into_iter();
        let items = (0..out_cols).map(|_| cells.next().flatten().unwrap_or_else(|| py.None()));
        Ok(PyList::new(py, items)?.unbind())
    }

    fn empty_row(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        let cols = if self.cols_per_row == -1 {
            1
        } else {
            self.cols_per_row.max(0)
        };
        if cols > MAX_COL_SPAN {
            return Err(PyMemoryError::new_err(()));
        }
        Ok(PyList::new(py, (0..cols).map(|_| py.None()))?.unbind())
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

        if self.cols_per_row != -1 {
            if self.cols_per_row <= 0
                || self.cols_per_row > MAX_COL_SPAN
                || col < self.first_col
                || col - self.first_col >= self.cols_per_row
            {
                // Valid later rows may be wider than the first: keep their cells.
                if self.cols_per_row > 0
                    && self.cols_per_row <= MAX_COL_SPAN
                    && col >= self.first_col
                    && col < 16384
                    && col - self.first_col >= self.cols_per_row
                {
                    self.cols_per_row = col - self.first_col + 1;
                } else {
                    return Ok(());
                }
            }
        } else {
            // Unset layout: freeze on the first stored cell; absurd spans are
            // corrupt-input noise the reference never shows.
            if self.first_col < 0 {
                self.first_col = col;
            }
            if col < self.first_col || col - self.first_col > MAX_COL_SPAN {
                return Ok(());
            }
        }
        let idx = (col - self.first_col) as usize;
        if idx >= self.cur_cells.len() {
            if self.cols_per_row > 0 && self.cols_per_row <= MAX_COL_SPAN {
                self.cur_cells
                    .reserve((self.cols_per_row as usize).saturating_sub(self.cur_cells.len()));
            }
            self.cur_cells.resize_with(idx + 1, || None);
        }
        self.cur_cells[idx] = Some(value.unbind());
        Ok(())
    }

    fn advance(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyList>>> {
        if self.done {
            return Ok(None);
        }
        if self.gap_rows_remaining > 0 {
            self.gap_rows_remaining -= 1;
            return self.empty_row(py).map(Some);
        }
        // On the first call, skip metadata to the first Row record or the first
        // cell record (cells before any Row record form a phantom row).
        if !self.is_initialized {
            let data_len = self.view.as_slice().len();
            while self.pos < data_len {
                let start = self.pos;
                let Some(rec) = next_record(self.view.as_slice(), &mut self.pos) else {
                    self.done = true;
                    return Ok(None);
                };
                if rec.id == 0x00 {
                    self.cur_row_num = row_index(self.view.as_slice(), rec.body)?;
                    self.seen_row_record = true;
                    self.is_initialized = true;
                    break;
                } else if (0x01..=0x0B).contains(&rec.id) {
                    self.pos = start; // let the loop below consume this cell
                    self.is_initialized = true;
                    break;
                }
            }
            if !self.is_initialized {
                self.done = true;
                return Ok(None);
            }
        }
        while self.pos < self.view.as_slice().len() {
            let Some(rec) = next_record(self.view.as_slice(), &mut self.pos) else {
                break;
            };
            match rec.id {
                0x00 => {
                    let row = self.take_row(py)?;
                    let new_row = row_index(self.view.as_slice(), rec.body)?;
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
        if !self.cur_cells.is_empty() {
            return self.take_row(py).map(Some);
        }
        if self.cur_row_num >= 0 {
            return self.empty_row(py).map(Some);
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
}

/// Read an XLSB worksheet from raw bytes.  Returns an iterator of rows, or
/// None when the layout is too wide (the Python reader then handles it).
#[pyfunction]
#[pyo3(signature = (data, shared_strings, styles_list, date_num_fmts, date_1904=None))]
pub fn read_xlsb_worksheet<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
    shared_strings: &Bound<'py, PyAny>,
    styles_list: &Bound<'py, PyAny>,
    date_num_fmts: &Bound<'py, PyAny>,
    date_1904: Option<&Bound<'py, PyAny>>,
) -> PyResult<Option<Bound<'py, XlsbReader>>> {
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
    let mut reader = XlsbReader {
        view,
        pos: 0,
        done: false,
        shared_strings: shared_strings.clone().unbind(),
        ss_len: shared_strings.len(),
        date_xf: date_xf_flags(styles_list, date_num_fmts),
        cur_cells: Vec::new(),
        cur_row_num: -1,
        gap_rows_remaining: 0,
        first_col: -1,
        cols_per_row: -1,
        is_initialized: false,
        seen_row_record: false,
        date_1904,
    };
    if !reader.prescan() {
        return Ok(None);
    }
    Ok(Some(Bound::new(py, reader)?))
}
