//! Streaming XLSX worksheet scanner, mirroring `_xlsx_read._fast_row` over
//! the raw worksheet part (and, in model mode, `workbook/_fast_sheet.fast_row`).
//!
//! The part is read in bounded chunks.  A row the scanner cannot reproduce
//! exactly is yielded as its raw bytes (`<row ...>...</row>`): the caller
//! parses it with ElementTree and calls `resume(row_number)`.  Rows are
//! `(row_number, [(column, value), ...])` with 1-based rows and 0-based
//! columns; in model mode cells go straight into the sheet's cells dict and
//! only the fallback rows are yielded.

use std::ffi::CStr;

use memchr::memmem;
use pyo3::exceptions::{PyIndexError, PyOverflowError, PyTypeError, PyValueError};
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyBool, PyBytes, PyDict, PyFloat, PyList, PyString, PyTime, PyTuple, PyType};

use crate::bytes::ByteView;
use crate::dates::serial_datetime;
use crate::model::{self, CellClass};
use crate::types::error_type;

const CHUNK: usize = 1 << 18;
/// 26 + 26^2 + 26^3 column keys, plus slot 0 meaning "not cached".
const COLUMN_KEYS: usize = 26 * 26 * 26 + 26 * 26 + 26 + 1;
const MAX_COLUMNS: i64 = 16384;

static DATETIME: PyOnceLock<Py<PyType>> = PyOnceLock::new();
static TIMEDELTA: PyOnceLock<Py<PyType>> = PyOnceLock::new();

/// What a scanner step gives: a value, or "let ElementTree read this row".
type Scan<T> = PyResult<Option<T>>;

#[inline]
fn word_byte(c: u8) -> bool {
    c.is_ascii_alphanumeric() || c == b'_'
}

#[inline]
fn space_byte(c: u8) -> bool {
    matches!(c, b' ' | b'\t' | b'\n' | b'\r' | 0x0C | 0x0B)
}

#[inline]
fn find(hay: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    if from > hay.len() {
        return None;
    }
    if needle.len() == 1 {
        return memchr::memchr(needle[0], &hay[from..]).map(|i| i + from);
    }
    memmem::find(&hay[from..], needle).map(|i| i + from)
}

#[inline]
fn contains(hay: &[u8], c: u8) -> bool {
    memchr::memchr(c, hay).is_some()
}

/// `bytes.decode("utf-8", errors)` through CPython, for its exact errors.
fn decode_utf8<'py>(
    py: Python<'py>,
    b: &[u8],
    errors: Option<&CStr>,
) -> PyResult<Bound<'py, PyAny>> {
    unsafe {
        let obj = ffi::PyUnicode_DecodeUTF8(
            b.as_ptr().cast(),
            b.len() as ffi::Py_ssize_t,
            errors.map_or(std::ptr::null(), |e| e.as_ptr()),
        );
        Bound::from_owned_ptr_or_err(py, obj)
    }
}

/// Text of the scanned part, which is valid UTF-8 (validated as it is read).
fn text<'py>(py: Python<'py>, b: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    match std::str::from_utf8(b) {
        Ok(s) => Ok(PyString::new(py, s).into_any()),
        Err(_) => decode_utf8(py, b, None),
    }
}

fn hex(c: u8) -> u32 {
    match c {
        b'0'..=b'9' => u32::from(c - b'0'),
        b'A'..=b'F' => u32::from(c - b'A' + 10),
        _ => u32::from(c.wrapping_sub(b'a').wrapping_add(10)),
    }
}

/// `xlspy._text.xstring_unescape`: every `_xHHHH_` becomes that code point.
fn unescape<'py>(py: Python<'py>, p: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    if memmem::find(p, b"_x").is_none() {
        return text(py, p);
    }
    let mut out = Vec::with_capacity(p.len());
    let (mut i, mut run) = (0usize, 0usize);
    while i + 7 <= p.len() {
        if p[i] == b'_'
            && p[i + 1] == b'x'
            && p[i + 2..i + 6].iter().all(u8::is_ascii_hexdigit)
            && p[i + 6] == b'_'
        {
            let cp =
                (hex(p[i + 2]) << 12) | (hex(p[i + 3]) << 8) | (hex(p[i + 4]) << 4) | hex(p[i + 5]);
            out.extend_from_slice(&p[run..i]);
            // Three-byte form for every cp >= 0x800, surrogates included
            // (decoded with surrogatepass below).
            if cp < 0x80 {
                out.push(cp as u8);
            } else if cp < 0x800 {
                out.extend_from_slice(&[(0xC0 | (cp >> 6)) as u8, (0x80 | (cp & 0x3F)) as u8]);
            } else {
                out.extend_from_slice(&[
                    (0xE0 | (cp >> 12)) as u8,
                    (0x80 | ((cp >> 6) & 0x3F)) as u8,
                    (0x80 | (cp & 0x3F)) as u8,
                ]);
            }
            i += 7;
            run = i;
        } else {
            i += 1;
        }
    }
    out.extend_from_slice(&p[run..]);
    if out.is_empty() {
        return Ok(PyString::new(py, "").into_any());
    }
    decode_utf8(py, &out, Some(c"surrogatepass"))
}

/// `[+-]? (digits [. digits?] | . digits) ([eE] [+-]? digits)?`
fn plain_float(p: &[u8]) -> bool {
    let n = p.len();
    let (mut i, mut digits) = (0usize, 0usize);
    if i < n && (p[i] == b'-' || p[i] == b'+') {
        i += 1;
    }
    while i < n && p[i].is_ascii_digit() {
        i += 1;
        digits += 1;
    }
    if i < n && p[i] == b'.' {
        i += 1;
        while i < n && p[i].is_ascii_digit() {
            i += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return false;
    }
    if i < n && (p[i] == b'e' || p[i] == b'E') {
        i += 1;
        if i < n && (p[i] == b'-' || p[i] == b'+') {
            i += 1;
        }
        let start = i;
        while i < n && p[i].is_ascii_digit() {
            i += 1;
        }
        if i == start {
            return false;
        }
    }
    i == n
}

/// int(text): plain values from a small integer; anything else (spaces,
/// underscores, huge numbers) goes through int() itself.
fn py_int<'py>(py: Python<'py>, p: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    let neg = p.first() == Some(&b'-');
    let digits = &p[usize::from(neg)..];
    if !digits.is_empty() && digits.len() <= 18 && digits.iter().all(u8::is_ascii_digit) {
        let v = digits
            .iter()
            .fold(0i64, |v, &d| v * 10 + i64::from(d - b'0'));
        return Ok((if neg { -v } else { v }).into_pyobject(py)?.into_any());
    }
    let t = text(py, p)?;
    unsafe { Bound::from_owned_ptr_or_err(py, ffi::PyLong_FromUnicodeObject(t.as_ptr(), 10)) }
}

/// float(text): plain decimal text through the routine float() uses; the
/// rest through float() itself, so nan/inf/underscores and errors match.
fn py_float_value(py: Python<'_>, p: &[u8]) -> PyResult<f64> {
    if plain_float(p) {
        let mut buf = Vec::with_capacity(p.len() + 1);
        buf.extend_from_slice(p);
        buf.push(0);
        let mut end: *mut std::ffi::c_char = std::ptr::null_mut();
        let v = unsafe {
            ffi::PyOS_string_to_double(buf.as_ptr().cast(), &mut end, std::ptr::null_mut())
        };
        let consumed = end as usize - buf.as_ptr() as usize;
        if PyErr::take(py).is_none() && consumed == p.len() {
            return Ok(v);
        }
    }
    let t = text(py, p)?;
    let f = unsafe { Bound::from_owned_ptr_or_err(py, ffi::PyFloat_FromString(t.as_ptr()))? };
    Ok(unsafe { f.cast_into_unchecked::<PyFloat>() }.value())
}

fn py_float<'py>(py: Python<'py>, p: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    Ok(PyFloat::new(py, py_float_value(py, p)?).into_any())
}

fn py_number<'py>(py: Python<'py>, p: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    if p.iter().any(|&c| c == b'.' || c == b'e' || c == b'E') {
        py_float(py, p)
    } else {
        py_int(py, p)
    }
}

fn fromisoformat<'py>(py: Python<'py>, v: &[u8]) -> PyResult<Bound<'py, PyAny>> {
    let t = text(py, v)?;
    DATETIME
        .import(py, "datetime", "datetime")?
        .call_method1(pyo3::intern!(py, "fromisoformat"), (t,))
}

/// The attributes and content of one `<c>` element.
#[derive(Default)]
struct RawCell<'a> {
    s: Option<&'a [u8]>,
    t: Option<&'a [u8]>,
    v: Option<&'a [u8]>,
    inl: Option<&'a [u8]>,
}

impl RawCell<'_> {
    fn t_is(&self, name: &[u8]) -> bool {
        self.t == Some(name)
    }

    fn v_or_empty(&self) -> &[u8] {
        self.v.unwrap_or(b"")
    }

    fn has_value(&self) -> bool {
        self.v.is_some_and(|v| !v.is_empty())
    }
}

/// Model-mode state (workbook/_fast_sheet.py).
struct Model {
    data_only: bool,
    rich_text: bool,
    cell: CellClass,
    sheet: Py<PyAny>,
    cells_dict: Py<PyDict>,
    column_numbers: Vec<Option<Py<PyAny>>>,
    row_cells: Vec<Py<PyAny>>,
    type_n: Py<PyString>,
    type_s: Py<PyString>,
    type_b: Py<PyString>,
    type_e: Py<PyString>,
    type_d: Py<PyString>,
    max_row: i64,
    max_column: i64,
    row_max_column: i64,
    current_row: i64,
}

/// Scanner state apart from the byte buffer, so parsing can borrow the
/// buffer while updating the state.
struct State {
    shared_strings: Py<PyList>,
    ss_len: usize,
    /// Date flags (ExcelReader) or date kinds 1..3 (model) per style.
    is_date_style: Vec<u8>,
    date_1904: bool,
    /// Column letters already seen, as in `_FastSheet.columns`: refs whose
    /// letters were cached are accepted without re-validating the digits.
    columns: Vec<i64>,
    next_row: i64,
    next_row_valid: bool,
    model: Option<Model>,
}

impl State {
    /// True for a date style; errors as the Python path raises them.
    fn date_style(&self, py: Python<'_>, s: Option<&[u8]>) -> PyResult<bool> {
        let style_len = self.is_date_style.len() as i64;
        let mut style: i64 = 0;
        if let Some(s) = s {
            match py_int(py, s)?.extract::<i64>() {
                Ok(v) => style = v,
                Err(e) if e.is_instance_of::<PyOverflowError>(py) => return Ok(false),
                Err(e) => return Err(e),
            }
        }
        if style < 0 {
            style += style_len;
        }
        if style < 0 {
            // Python's `style < len(...)` passes for any negative index, so a
            // too-negative one raises IndexError from the list access.
            return Err(PyIndexError::new_err("list index out of range"));
        }
        Ok(style < style_len && self.is_date_style[style as usize] != 0)
    }

    /// One cell value, following `_fast_row` + `_fast_value` exactly.
    fn value<'py>(&self, py: Python<'py>, c: &RawCell<'_>) -> Scan<Bound<'py, PyAny>> {
        // Common case: an untyped, non-empty value with no date style.
        if c.has_value() && c.t.is_none() && !self.date_style(py, c.s)? {
            return py_number(py, c.v_or_empty()).map(Some);
        }
        if c.t_is(b"inlineStr") {
            let Some(inl) = c.inl else {
                return Ok(Some(PyString::new(py, "").into_any()));
            };
            if contains(inl, b'&') || contains(inl, b'\r') {
                return Ok(None);
            }
            let n = inl.len();
            if n < 7 || inl[0] != b'<' || inl[1] != b't' {
                return Ok(None);
            }
            let mut q = 2;
            if inl[q] == b' ' {
                while q < n && inl[q] != b'>' {
                    q += 1;
                }
                if q >= n {
                    return Ok(None);
                }
            }
            if inl[q] != b'>' {
                return Ok(None);
            }
            q += 1;
            let start = q;
            while q < n && inl[q] != b'<' {
                q += 1;
            }
            if q + 4 != n || &inl[q..] != b"</t>" {
                return Ok(None);
            }
            return unescape(py, &inl[start..q]).map(Some);
        }
        if c.t_is(b"s") {
            let mut index: Option<i64> = Some(0);
            if c.has_value() {
                match py_int(py, c.v_or_empty())?.extract::<i64>() {
                    Ok(v) => index = Some(v),
                    Err(e) if e.is_instance_of::<PyOverflowError>(py) => index = None,
                    Err(e) => return Err(e),
                }
            }
            if let Some(i) = index
                && i >= 0
                && (i as usize) < self.ss_len
            {
                return Ok(Some(self.shared_strings.bind(py).get_item(i as usize)?));
            }
            return text(py, c.v_or_empty()).map(Some);
        }
        if c.t_is(b"b") {
            return Ok(Some(
                PyBool::new(py, c.v == Some(b"1")).to_owned().into_any(),
            ));
        }
        if c.t_is(b"e") {
            let t = PyString::new(
                py,
                &format!(
                    "error:{}",
                    text(py, c.v_or_empty())?.cast::<PyString>()?.to_str()?
                ),
            );
            return Ok(Some(match error_type(py) {
                Some(e) => e.call1((t,))?,
                None => t.into_any(),
            }));
        }
        if c.t_is(b"str") {
            if c.has_value() && (contains(c.v_or_empty(), b'&') || contains(c.v_or_empty(), b'\r'))
            {
                return Ok(None);
            }
            return unescape(py, c.v_or_empty()).map(Some);
        }
        if c.t_is(b"d") {
            if !c.has_value() {
                return Ok(Some(py.None().into_bound(py)));
            }
            return fromisoformat(py, c.v_or_empty()).map(Some);
        }
        // Default numeric path, including t="n" and unknown types.
        if !c.has_value() {
            return Ok(Some(py.None().into_bound(py)));
        }
        let v = c.v_or_empty();
        if self.date_style(py, c.s)? {
            let serial = match py_float_value(py, v) {
                Ok(f) => f,
                Err(e) if e.is_instance_of::<PyValueError>(py) => return text(py, v).map(Some),
                Err(e) => return Err(e),
            };
            return match serial_datetime(py, serial, self.date_1904) {
                Ok(moment) => Ok(Some(moment)),
                Err(e)
                    if e.is_instance_of::<PyValueError>(py)
                        || e.is_instance_of::<PyOverflowError>(py) =>
                {
                    text(py, v).map(Some)
                }
                Err(e) => Err(e),
            };
        }
        py_number(py, v).map(Some)
    }

    /// A numeric value under a date format, as workbook/_reader._number_value:
    /// kind 1 a datetime, 2 a time of day, 3 a timedelta.  None when Python
    /// would keep the number (the XML parser then decides).
    fn model_date<'py>(&self, py: Python<'py>, serial: f64, kind: u8) -> Scan<Bound<'py, PyAny>> {
        match kind {
            1 => match serial_datetime(py, serial, self.date_1904) {
                Ok(v) => Ok(Some(v)),
                Err(e)
                    if e.is_instance_of::<PyValueError>(py)
                        || e.is_instance_of::<PyOverflowError>(py) =>
                {
                    Ok(None)
                }
                Err(e) => Err(e),
            },
            2 => {
                let seconds = (serial - serial.floor()) * 86400.0;
                let whole = seconds as i64;
                let micros = ((seconds - whole as f64) * 1_000_000.0).round_ties_even() as i64;
                if !serial.is_finite() || !(0..1_000_000).contains(&micros) {
                    return Ok(None);
                }
                Ok(Some(
                    PyTime::new(
                        py,
                        ((whole / 3600) % 24) as u8,
                        ((whole % 3600) / 60) as u8,
                        (whole % 60) as u8,
                        micros as u32,
                        None,
                    )?
                    .into_any(),
                ))
            }
            _ => {
                if !serial.is_finite() || serial.abs() > 999_999_999.0 {
                    return Ok(None);
                }
                let kwargs = PyDict::new(py);
                kwargs.set_item("days", serial)?;
                match TIMEDELTA
                    .import(py, "datetime", "timedelta")?
                    .call((), Some(&kwargs))
                {
                    Ok(v) => Ok(Some(v)),
                    Err(e) if e.is_instance_of::<PyOverflowError>(py) => Ok(None),
                    Err(e) => Err(e),
                }
            }
        }
    }

    /// One cell in model mode, mirroring `fast_row`: added to the row's cells
    /// (or skipped as empty), None to fall back.
    fn model_cell(&mut self, py: Python<'_>, col: i64, c: &RawCell<'_>) -> Scan<()> {
        let mut style: i64 = 0;
        if let Some(s) = c.s.filter(|s| !s.is_empty()) {
            if s.len() > 9 || !s.iter().all(u8::is_ascii_digit) {
                return Ok(None);
            }
            style = s.iter().fold(0, |v, &d| v * 10 + i64::from(d - b'0'));
        }
        let m = self.model.as_ref().unwrap();
        let typed = matches!(c.t, Some(b"s" | b"b" | b"e" | b"d" | b"str" | b"inlineStr"));
        let (value, data_type) = if !typed {
            let value = if !c.has_value() {
                py.None().into_bound(py)
            } else {
                let serial = match py_float_value(py, c.v_or_empty()) {
                    Ok(f) => f,
                    Err(e) if e.is_instance_of::<PyValueError>(py) => return Ok(None),
                    Err(e) => return Err(e),
                };
                let flags = &self.is_date_style;
                let kind = if (style as usize) < flags.len() {
                    flags[style as usize]
                } else {
                    flags.first().copied().unwrap_or(0)
                };
                if kind != 0 {
                    match self.model_date(py, serial, kind)? {
                        Some(v) => v,
                        None => return Ok(None),
                    }
                } else if serial.is_finite() && serial == serial.floor() {
                    unsafe { Bound::from_owned_ptr_or_err(py, ffi::PyLong_FromDouble(serial))? }
                } else {
                    PyFloat::new(py, serial).into_any()
                }
            };
            (value, &m.type_n)
        } else if c.t_is(b"inlineStr") {
            if m.rich_text {
                return Ok(None);
            }
            let value = if c.inl.is_none() {
                py.None().into_bound(py)
            } else {
                match self.value(py, c)? {
                    Some(v) => v,
                    None => return Ok(None),
                }
            };
            (value, &m.type_s)
        } else if c.t_is(b"str") || c.t_is(b"b") {
            let Some(value) = self.value(py, c)? else {
                return Ok(None);
            };
            (value, if c.t_is(b"str") { &m.type_s } else { &m.type_b })
        } else if c.t_is(b"s") {
            let mut value = None;
            if c.has_value() {
                let v = c.v_or_empty();
                if v.len() > 9 || !v.iter().all(u8::is_ascii_digit) {
                    return Ok(None);
                }
                let index = v
                    .iter()
                    .fold(0usize, |i, &d| i * 10 + usize::from(d - b'0'));
                if index < self.ss_len {
                    value = Some(self.shared_strings.bind(py).get_item(index)?);
                }
            }
            let value = match value {
                Some(v) => v,
                None => text(py, c.v_or_empty())?,
            };
            (value, &m.type_s)
        } else {
            // e, d: the text itself (d parsed when present).
            let is_e = c.t_is(b"e");
            if c.has_value() && contains(c.v_or_empty(), b'&') {
                return Ok(None);
            }
            let value = if !is_e && c.has_value() {
                match fromisoformat(py, c.v_or_empty()) {
                    Ok(v) => v,
                    Err(e) if e.is_instance_of::<PyValueError>(py) => return Ok(None),
                    Err(e) => return Err(e),
                }
            } else if !is_e {
                py.None().into_bound(py)
            } else {
                text(py, c.v_or_empty())?
            };
            (value, if is_e { &m.type_e } else { &m.type_d })
        };
        if value.is_none() && style == 0 {
            return Ok(Some(()));
        }
        let data_type = data_type.clone_ref(py).into_bound(py).into_any();
        let m = self.model.as_mut().unwrap();
        let column = match &m.column_numbers[col as usize] {
            Some(c) => c.bind(py).clone(),
            None => {
                let c = (col + 1).into_pyobject(py)?.into_any();
                m.column_numbers[col as usize] = Some(c.clone().unbind());
                c
            }
        };
        let cell = m.cell.alloc(py)?;
        let none = py.None().into_bound(py);
        m.cell.set(&cell, model::SHEET, m.sheet.bind(py));
        m.cell.set(&cell, model::COLUMN, &column);
        m.cell.set(&cell, model::VALUE, &value);
        m.cell.set(&cell, model::FORMULA, &none);
        m.cell.set(&cell, model::DATA_TYPE, &data_type);
        m.cell
            .set(&cell, model::STYLE_ID, &style.into_pyobject(py)?.into_any());
        m.cell.set(&cell, model::COMMENT, &none);
        m.cell.set(&cell, model::HYPERLINK, &none);
        m.cell.set(&cell, model::NAMED_STYLE, &none);
        m.cell.set(
            &cell,
            model::DIRTY,
            &PyBool::new(py, false).to_owned().into_any(),
        );
        m.row_cells.push(cell.unbind());
        m.row_max_column = m.row_max_column.max(col + 1);
        Ok(Some(()))
    }

    /// One `<c ...>` element at `body[*pos]`; the position moves past it.
    fn parse_cell<'py>(
        &mut self,
        py: Python<'py>,
        body: &[u8],
        pos: &mut usize,
        next_col: &mut i64,
        cells: Option<&Bound<'py, PyList>>,
    ) -> Scan<()> {
        let n = body.len();
        let mut i = *pos + 2;
        let mut c = RawCell::default();
        let mut r: Option<&[u8]> = None;
        while i < n && body[i] == b' ' {
            let mark = i;
            i += 1;
            let name = i;
            while i < n && word_byte(body[i]) {
                i += 1;
            }
            let name = &body[name..i];
            if name.is_empty() || i + 1 >= n || body[i] != b'=' || body[i + 1] != b'"' {
                // Not an attribute pair: the terminator check below decides,
                // like the regex's trailing \s* (`<c />`, `<c r="A1" >`).
                i = mark;
                break;
            }
            i += 2;
            let start = i;
            while i < n && body[i] != b'"' {
                i += 1;
            }
            if i >= n {
                return Ok(None);
            }
            let value = &body[start..i];
            i += 1;
            match name {
                b"r" => r = Some(value),
                b"s" => c.s = Some(value),
                b"t" => c.t = Some(value),
                _ => {}
            }
        }
        while i < n && space_byte(body[i]) {
            i += 1;
        }
        if i >= n {
            return Ok(None);
        }
        if body[i] == b'/' {
            if i + 1 >= n || body[i + 1] != b'>' {
                return Ok(None);
            }
            i += 2;
        } else if body[i] == b'>' {
            i += 1;
            if i + 2 < n && body[i] == b'<' && body[i + 1] == b'f' && !word_byte(body[i + 2]) {
                let Some(gt) = find(body, b">", i + 2) else {
                    return Ok(None);
                };
                if body[gt - 1] == b'/' {
                    i = gt + 1;
                } else {
                    let Some(close) = find(body, b"</f>", gt + 1) else {
                        return Ok(None);
                    };
                    if find(body, b"<", gt + 1).is_some_and(|lt| lt < close) {
                        return Ok(None);
                    }
                    i = close + 4;
                }
            }
            if body[i..].starts_with(b"<v>") {
                let start = i + 3;
                let end = start + memchr::memchr(b'<', &body[start..]).unwrap_or(n - start);
                if !body[end..].starts_with(b"</v>") {
                    return Ok(None);
                }
                c.v = Some(&body[start..end]);
                i = end + 4;
            } else if body[i..].starts_with(b"<is>") {
                let Some(close) = find(body, b"</is>", i + 4) else {
                    return Ok(None);
                };
                c.inl = Some(&body[i + 4..close]);
                i = close + 5;
            }
            if !body[i..].starts_with(b"</c>") {
                return Ok(None);
            }
            i += 4;
        } else {
            return Ok(None);
        }

        let is_model = self.model.is_some();
        let col = match r.filter(|r| !r.is_empty()) {
            None => *next_col,
            Some(r) => {
                let digits = r.iter().rev().take_while(|c| c.is_ascii_digit()).count();
                let letters = r.len() - digits;
                if !(1..=3).contains(&letters) || !r[..letters].iter().all(u8::is_ascii_uppercase) {
                    return Ok(None);
                }
                let key = r[..letters]
                    .iter()
                    .fold(0usize, |k, &ch| k * 26 + usize::from(ch - b'A' + 1));
                let col = if self.columns[key] != 0 {
                    self.columns[key] - 1
                } else {
                    if digits == 0 {
                        return Ok(None);
                    }
                    self.columns[key] = key as i64;
                    key as i64 - 1
                };
                if let Some(m) = &self.model {
                    // The model keeps the row of the reference: a cell whose
                    // reference names another row goes to the XML parser.
                    if digits == 0 || digits > 7 || col >= MAX_COLUMNS {
                        return Ok(None);
                    }
                    let ref_row = r[letters..]
                        .iter()
                        .fold(0i64, |v, &d| v * 10 + i64::from(d - b'0'));
                    if ref_row != m.current_row {
                        return Ok(None);
                    }
                }
                col
            }
        };
        *next_col = col + 1;
        if is_model {
            if col >= MAX_COLUMNS {
                return Ok(None);
            }
            let ok = self.model_cell(py, col, &c)?;
            if ok.is_some() {
                *pos = i;
            }
            return Ok(ok);
        }
        let Some(value) = self.value(py, &c)? else {
            return Ok(None);
        };
        cells.unwrap().append(PyTuple::new(
            py,
            [col.into_pyobject(py)?.into_any(), value],
        )?)?;
        *pos = i;
        Ok(Some(()))
    }

    /// The cells of a row body, parsed until the first fallback.
    fn parse_cells<'py>(
        &mut self,
        py: Python<'py>,
        body: &[u8],
        cells: Option<&Bound<'py, PyList>>,
    ) -> Scan<()> {
        let mut next_col = 0i64;
        let mut i = 0usize;
        while i < body.len() {
            let Some(lt) = memchr::memchr(b'<', &body[i..]) else {
                break;
            };
            i += lt;
            if body.get(i + 1) == Some(&b'c') {
                if self
                    .parse_cell(py, body, &mut i, &mut next_col, cells)?
                    .is_none()
                {
                    return Ok(None);
                }
            } else {
                i += 1;
            }
        }
        Ok(Some(()))
    }

    /// A row in model mode: its cells go into the cells dict only once the
    /// whole row parsed, so a fallback row leaves nothing behind.  Rows with
    /// formulas (unless data_only) or with height/visibility/outline
    /// attributes are left to the XML parser.
    fn model_row(
        &mut self,
        py: Python<'_>,
        row_number: i64,
        attrs: &[u8],
        body: Option<&[u8]>,
    ) -> Scan<()> {
        if !(1..=1_048_576).contains(&row_number) {
            return Ok(None);
        }
        if [&b" ht="[..], b" hidden=", b" outlineLevel=", b" collapsed="]
            .iter()
            .any(|a| memmem::find(attrs, a).is_some())
        {
            return Ok(None);
        }
        let m = self.model.as_mut().unwrap();
        if body.is_some_and(|b| !m.data_only && memmem::find(b, b"<f").is_some()) {
            return Ok(None);
        }
        m.current_row = row_number;
        m.row_max_column = 0;
        let parsed = match body {
            Some(b) => self.parse_cells(py, b, None),
            None => Ok(Some(())),
        };
        let m = self.model.as_mut().unwrap();
        let cells = std::mem::take(&mut m.row_cells);
        let parsed = parsed.and_then(|ok| {
            if ok.is_some() && !cells.is_empty() {
                let key = row_number.into_pyobject(py)?.into_any();
                let cells_dict = m.cells_dict.bind(py);
                let row_dict = match cells_dict.get_item(&key)? {
                    Some(d) => d,
                    None => {
                        let d = PyDict::new(py).into_any();
                        cells_dict.set_item(&key, &d)?;
                        d
                    }
                };
                for cell in &cells {
                    let cell = cell.bind(py);
                    m.cell.set(cell, model::ROW, &key);
                    let column = m
                        .cell
                        .get(cell, model::COLUMN)
                        .unwrap_or_else(|| py.None().into_bound(py));
                    row_dict.set_item(column, cell)?;
                }
            }
            Ok(ok)
        });
        let ok = parsed?;
        self.next_row = row_number + 1;
        self.next_row_valid = true;
        if ok.is_none() {
            return Ok(None);
        }
        m.max_row = m.max_row.max(row_number);
        m.max_column = m.max_column.max(m.row_max_column);
        Ok(Some(()))
    }

    /// One complete row: `Some(row tuple)` (None in model mode), or the
    /// outer None to fall back.
    fn parse_row<'py>(
        &mut self,
        py: Python<'py>,
        attrs: &[u8],
        body: Option<&[u8]>,
    ) -> Scan<Option<Bound<'py, PyAny>>> {
        if body.is_some_and(has_prefix) {
            return Ok(None);
        }
        let mut row_number: i64 = -1;
        let space_r = memmem::find(attrs, b" r=").is_some();
        let mut i = 0;
        while let Some(at) = find(attrs, b" r=\"", i) {
            let mut j = at + 4;
            let start = j;
            while j < attrs.len() && attrs[j].is_ascii_digit() {
                j += 1;
            }
            let digits = j - start;
            if digits > 0 && j < attrs.len() && attrs[j] == b'"' {
                if digits > 18 {
                    return Ok(None);
                }
                let value = attrs[start..j]
                    .iter()
                    .fold(0i64, |v, &d| v * 10 + i64::from(d - b'0'));
                if value > i64::MAX - 1 {
                    return Ok(None);
                }
                row_number = value;
                break;
            }
            i = at + 1;
        }
        if row_number < 0 {
            if space_r || !self.next_row_valid {
                return Ok(None);
            }
            row_number = self.next_row;
        }
        if self.model.is_some() {
            return Ok(self.model_row(py, row_number, attrs, body)?.map(|()| None));
        }
        let cells = PyList::empty(py);
        let parsed = self.parse_cells(py, body.unwrap_or(b""), Some(&cells));
        self.next_row = row_number + 1;
        self.next_row_valid = true;
        if parsed?.is_none() {
            return Ok(None);
        }
        Ok(Some(Some(
            PyTuple::new(
                py,
                [row_number.into_pyobject(py)?.into_any(), cells.into_any()],
            )?
            .into_any(),
        )))
    }
}

/// Mirrors `_PREFIXED`: `</?[A-Za-z_][\w.-]*:`.
fn has_prefix(p: &[u8]) -> bool {
    let n = p.len();
    let mut i = 0;
    while i + 1 < n {
        let Some(lt) = memchr::memchr(b'<', &p[i..n - 1]) else {
            return false;
        };
        i += lt;
        let mut j = i + 1;
        if j < n && p[j] == b'/' {
            j += 1;
        }
        if j < n && (p[j].is_ascii_alphabetic() || p[j] == b'_') {
            j += 1;
            while j < n && (word_byte(p[j]) || p[j] == b'.' || p[j] == b'-' || p[j] >= 0x80) {
                j += 1;
            }
            if j < n && p[j] == b':' {
                return true;
            }
        }
        i += 1;
    }
    false
}

/// Where the next complete `<row>` element lies in the buffer.
struct RowSpan {
    start: usize,
    end: usize,
    attrs: (usize, usize),
    body: Option<(usize, usize)>,
}

fn next_row(buf: &[u8], from: usize) -> Option<RowSpan> {
    let len = buf.len();
    let mut i = from;
    loop {
        i = find(buf, b"<", i)?;
        if i + 4 > len {
            return None; // more bytes needed to test <row
        }
        if &buf[i..i + 4] != b"<row" {
            i += 1;
            continue;
        }
        if i + 4 == len {
            return None;
        }
        let boundary = buf[i + 4];
        if word_byte(boundary) || boundary >= 0x80 {
            i += 1;
            continue;
        }
        let gt = find(buf, b">", i + 4)?;
        if gt > i + 4 && buf[gt - 1] == b'/' {
            return Some(RowSpan {
                start: i,
                end: gt + 1,
                attrs: (i + 4, gt - 1),
                body: None,
            });
        }
        let close = find(buf, b"</row>", gt + 1)?;
        return Some(RowSpan {
            start: i,
            end: close + 6,
            attrs: (i + 4, gt),
            body: Some((gt + 1, close)),
        });
    }
}

#[pyclass(module = "xlspy._rs_core")]
pub struct XlsxReader {
    stream: Py<PyAny>,
    buf: Vec<u8>,
    /// First byte not yet scanned, and bytes known to be valid UTF-8.
    pos: usize,
    validated: usize,
    eof: bool,
    done: bool,
    state: State,
}

impl XlsxReader {
    /// Check the bytes not yet validated.  A truncated character at the very
    /// end is not an error: the next chunk may complete it.
    fn validate(&mut self, py: Python<'_>, final_check: bool) -> PyResult<()> {
        let rest = &self.buf[self.validated..];
        match std::str::from_utf8(rest) {
            Ok(_) => {
                self.validated = self.buf.len();
                Ok(())
            }
            Err(e) if e.error_len().is_none() && !final_check => {
                self.validated += e.valid_up_to();
                Ok(())
            }
            // CPython's decoder, for the exact exception it raises.
            Err(_) => match decode_utf8(py, rest, None) {
                Ok(_) => {
                    self.validated = self.buf.len();
                    Ok(())
                }
                Err(e) => Err(e),
            },
        }
    }

    fn fill(&mut self, py: Python<'_>) -> PyResult<()> {
        if self.eof {
            return Ok(());
        }
        if self.pos > 0 {
            self.buf.drain(..self.pos);
            self.validated = self.validated.saturating_sub(self.pos);
            self.pos = 0;
        }
        let chunk = self
            .stream
            .bind(py)
            .call_method1(pyo3::intern!(py, "read"), (CHUNK,))?;
        let view = ByteView::new(&chunk)?;
        let data = view.as_slice();
        if data.is_empty() {
            self.eof = true;
            return Ok(());
        }
        self.buf.extend_from_slice(data);
        self.validate(py, false)
    }

    fn advance<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyAny>>> {
        if self.done {
            return Ok(None);
        }
        loop {
            if let Some(span) = next_row(&self.buf, self.pos) {
                let attrs = &self.buf[span.attrs.0..span.attrs.1];
                let body = span.body.map(|(a, b)| &self.buf[a..b]);
                let row = self.state.parse_row(py, attrs, body);
                self.pos = span.end;
                match row? {
                    None => {
                        return Ok(Some(
                            PyBytes::new(py, &self.buf[span.start..span.end]).into_any(),
                        ));
                    }
                    Some(None) => continue, // model mode: loaded into the cells dict
                    Some(Some(row)) => return Ok(Some(row)),
                }
            }
            if self.eof || find(&self.buf, b"</sheetData>", self.pos).is_some() {
                self.done = true;
                // Once iteration stops no chunk can complete a truncated
                // character, so the Python reader's final decode would raise.
                self.validate(py, true)?;
                return Ok(None);
            }
            self.fill(py)?;
        }
    }
}

#[pymethods]
impl XlsxReader {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyAny>>> {
        let result = self.advance(py);
        if result.is_err() {
            self.done = true;
        }
        result
    }

    /// resume(row_number): continue after a fallback row handled by the caller.
    fn resume(&mut self, py: Python<'_>, row_number: &Bound<'_, PyAny>) -> PyResult<()> {
        match row_number.extract::<i64>() {
            Ok(v) if (0..i64::MAX).contains(&v) => {
                self.state.next_row = v + 1;
                self.state.next_row_valid = true;
            }
            Ok(_) => self.state.next_row_valid = false,
            Err(e) if e.is_instance_of::<PyOverflowError>(py) => self.state.next_row_valid = false,
            Err(e) => return Err(e),
        }
        Ok(())
    }

    /// bounds(): (max_row, max_column) of the rows loaded in model mode.
    fn bounds(&self) -> (i64, i64) {
        self.state
            .model
            .as_ref()
            .map_or((0, 0), |m| (m.max_row, m.max_column))
    }

    /// tail(): the bytes read after the last row, once iteration stopped.
    fn tail<'py>(&self, py: Python<'py>) -> Bound<'py, PyBytes> {
        PyBytes::new(py, &self.buf[self.pos.min(self.buf.len())..])
    }
}

/// Scan an XLSX worksheet from a stream.  Returns an iterator of rows.
///
/// Args: (stream, head, shared_strings, is_date_style, date_1904, model)
/// where model is None or (Cell, slot descriptors, sheet, cells dict,
/// data_only, rich_text).
#[pyfunction]
#[pyo3(signature = (stream, head, shared_strings, is_date_style, date_1904=None, model=None))]
pub fn read_xlsx_worksheet<'py>(
    py: Python<'py>,
    stream: &Bound<'py, PyAny>,
    head: &Bound<'py, PyAny>,
    shared_strings: &Bound<'py, PyAny>,
    is_date_style: &Bound<'py, PyAny>,
    date_1904: Option<&Bound<'py, PyAny>>,
    model: Option<&Bound<'py, PyAny>>,
) -> PyResult<Bound<'py, XlsxReader>> {
    let date_1904 = match date_1904 {
        Some(flag) => flag.is_truthy()?,
        None => false,
    };
    let model = match model.filter(|m| !m.is_none()) {
        None => None,
        Some(m) => {
            let m = match m.cast::<PyTuple>() {
                Ok(t) if t.len() == 6 => t,
                _ => {
                    return Err(PyTypeError::new_err(
                        "model must be (Cell, slots, sheet, cells, data_only, rich_text)",
                    ));
                }
            };
            let (data_only, rich_text) = (m.get_item(4)?.is_truthy()?, m.get_item(5)?.is_truthy()?);
            let (cell_type, cells_dict) = (m.get_item(0)?, m.get_item(3)?);
            let (true, Ok(cells_dict)) = (
                cell_type.is_instance_of::<PyType>(),
                cells_dict.cast::<PyDict>(),
            ) else {
                return Err(PyTypeError::new_err(
                    "model needs the Cell type and a cells dict",
                ));
            };
            Some((
                cell_type.clone(),
                m.get_item(1)?,
                m.get_item(2)?,
                cells_dict.clone(),
                data_only,
                rich_text,
            ))
        }
    };
    let head = head
        .cast::<PyBytes>()
        .map_err(|_| PyTypeError::new_err("head must be bytes"))?;
    let shared_strings = shared_strings
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("shared_strings must be a list"))?;
    let styles = is_date_style
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("is_date_style must be a list"))?;
    if !stream.hasattr("read").unwrap_or(false) {
        return Err(PyTypeError::new_err("stream must have a read method"));
    }
    let model = match model {
        None => None,
        Some((cell_type, descriptors, sheet, cells_dict, data_only, rich_text)) => {
            let intern = |s: &str| PyString::intern(py, s).unbind();
            Some(Model {
                data_only,
                rich_text,
                cell: CellClass::new(&cell_type, &descriptors)?,
                sheet: sheet.unbind(),
                cells_dict: cells_dict.unbind(),
                column_numbers: (0..MAX_COLUMNS).map(|_| None).collect(),
                row_cells: Vec::new(),
                type_n: intern("n"),
                type_s: intern("s"),
                type_b: intern("b"),
                type_e: intern("e"),
                type_d: intern("d"),
                max_row: 0,
                max_column: 0,
                row_max_column: 0,
                current_row: 0,
            })
        }
    };
    // Flags (bool) for ExcelReader, date kinds (int 0..3) for the model.
    let mut flags = Vec::with_capacity(styles.len());
    for item in styles.iter() {
        let kind =
            if item.is_instance_of::<pyo3::types::PyInt>() && !item.is_instance_of::<PyBool>() {
                let k: i64 = item.extract()?;
                if (0..=3).contains(&k) { k as u8 } else { 0 }
            } else {
                u8::from(item.is_truthy()?)
            };
        flags.push(kind);
    }
    let mut reader = XlsxReader {
        stream: stream.clone().unbind(),
        buf: head.as_bytes().to_vec(),
        pos: 0,
        validated: 0,
        eof: false,
        done: false,
        state: State {
            ss_len: shared_strings.len(),
            shared_strings: shared_strings.clone().unbind(),
            is_date_style: flags,
            date_1904,
            columns: vec![0; COLUMN_KEYS],
            next_row: 1,
            next_row_valid: true,
            model,
        },
    };
    reader.validate(py, false)?;
    Bound::new(py, reader)
}
