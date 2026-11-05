//! XLSX row encoder (`<row><c/>...</row>`), mirroring `_write_worksheet_xml`
//! cell by cell.  Rows it cannot encode exactly as the Python writer would
//! return None, and the Python row writer then handles (or rejects) them.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyBytes, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple,
};

use crate::cells::{
    Cell, Hooked, Serial, Sst, Styles, call_hook, date_style_xf, excel_serial, fmt_tuple_kind,
    fmt_xf_value, rollback, text_fits,
};
use crate::numfmt::{write_repr, write_u32};
use crate::pyseq::FastSeq;
use crate::sst::u32_mask;
use crate::types::{
    decimal_type, formula_type, hyperlink_type, int_as_i64, style_type, type_check,
};

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

/// XML-escape & < > only, as `saxutils.escape` does for formula text. The
/// Python writer does not apply the `_xHHHH_` escapes there, so neither does this.
fn write_plain_escaped(out: &mut Vec<u8>, s: &[u8]) {
    let mut run = 0;
    for (i, &c) in s.iter().enumerate() {
        let entity: &[u8] = match c {
            b'&' => b"&amp;",
            b'<' => b"&lt;",
            b'>' => b"&gt;",
            _ => continue,
        };
        out.extend_from_slice(&s[run..i]);
        out.extend_from_slice(entity);
        run = i + 1;
    }
    out.extend_from_slice(&s[run..]);
}

/// Whether `<t>` needs xml:space="preserve", as `updater_utils.xml_space_attribute`
/// decides it: Excel drops the spaces at the ends of a text without it.
fn needs_preserve(text: &str) -> bool {
    let is_space = |c: char| c.is_whitespace() || ('\u{1c}'..='\u{1f}').contains(&c);
    text.starts_with(is_space) || text.ends_with(is_space) || text.contains(['\t', '\n', '\r'])
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
    /// Style id of the built-in link style (0: the Python writer handles links).
    link_xf: u32,
    /// (column, Hyperlink) for the link cells written, in column order.
    links: Vec<(u32, Py<PyAny>)>,
    /// The writer's cell hook (see `call_hook`), or None.
    hook: Bound<'py, PyAny>,
    row_idx: u32,
    /// The current cell's style before a date's variant (see `date_style_xf`).
    hook_style: u32,
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
        let Ok(text) = s.to_str() else {
            return Ok(Cell::Fallback); // lone surrogates
        };
        let utf8 = text.as_bytes();
        let start = self.out.len();
        self.out.extend_from_slice(b"<c t=\"inlineStr\"");
        self.style_or_header(style);
        self.out.extend_from_slice(b"><is><t");
        if needs_preserve(text) {
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

    /// Hyperlink cell, as the Python writer writes a link with no style of its
    /// own: the display text in the built-in link style, and the link kept for
    /// the sheet's <hyperlinks>. Anything else goes to the Python writer.
    fn hyperlink(
        &mut self,
        cell: &Bound<'_, PyAny>,
        style: u32,
        col: u32,
        decimal: Option<&Bound<'_, pyo3::types::PyType>>,
        may_hook: bool,
    ) -> PyResult<Cell> {
        if style != 0 && may_hook {
            // A link in a styled cell: the hook knows the link variant of the
            // style (and keeps the link) once the Python writer has made it.
            return self.hooked(cell, style, col, decimal, true);
        }
        if style != 0 || self.link_xf == 0 {
            return Ok(Cell::Fallback);
        }
        let text = cell.getattr("display_text")?;
        let Ok(text) = text.cast::<PyString>() else {
            return Ok(Cell::Fallback);
        };
        let link_xf = self.link_xf;
        match self.string(text, link_xf)? {
            Cell::Done => {
                self.links.push((col, cell.clone().unbind()));
                Ok(Cell::Done)
            }
            Cell::Fallback => Ok(Cell::Fallback),
        }
    }

    fn cell(
        &mut self,
        cell: &Bound<'_, PyAny>,
        style: u32,
        col: u32,
        decimal: Option<&Bound<'_, pyo3::types::PyType>>,
    ) -> PyResult<Cell> {
        self.value(cell, style, col, decimal, true)
    }

    /// A cell of a type the encoder does not know, through the writer's hook.
    fn hooked(
        &mut self,
        cell: &Bound<'_, PyAny>,
        style: u32,
        col: u32,
        decimal: Option<&Bound<'_, pyo3::types::PyType>>,
        is_link: bool,
    ) -> PyResult<Cell> {
        // A link's style is its own; anything else asks in the cell's style
        // before a date's variant (see `hook_style`).
        let style = if is_link { style } else { self.hook_style };
        match call_hook(&self.hook, cell, col, style, self.row_idx)? {
            Hooked::Fallback | Hooked::SstKey(_) => Ok(Cell::Fallback),
            Hooked::Raw(raw) => {
                self.out.extend_from_slice(raw.as_bytes());
                Ok(Cell::Done)
            }
            Hooked::Value(value, xf) => self.value(&value, xf, col, decimal, false),
        }
    }

    /// Encode one value; `may_hook` false for a value the hook returned.
    fn value(
        &mut self,
        cell: &Bound<'_, PyAny>,
        style: u32,
        col: u32,
        decimal: Option<&Bound<'_, pyo3::types::PyType>>,
        may_hook: bool,
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
            let exact = if cell.is_exact_instance_of::<PyInt>() {
                int_as_i64(cell)
            } else {
                None
            };
            let d = match exact {
                Some(v) => v as f64,
                None => match int_float(cell) {
                    IntFloat::Value(d) => d,
                    // Past the float range: the digits, as text (as XlsbWriter).
                    IntFloat::Overflow => {
                        return match cell.str() {
                            Ok(s) => self.string(&s, style),
                            Err(_) => Ok(Cell::Fallback),
                        };
                    }
                    IntFloat::Error => return Ok(Cell::Fallback),
                },
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
                    // SAFETY: `cell`, `inf` and `ninf` are live for the loop; the
                    // call returns -1 with an exception set on failure (last arm).
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
                // Unreadable fields (pandas' NaT) or past exact microseconds.
                Serial::Fallback if may_hook => self.hooked(cell, style, col, decimal, false),
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
        if formula_type(py).is_some_and(|t| type_check(cell, &t)) {
            return self.formula(cell, style);
        }
        if hyperlink_type(py).is_some_and(|t| type_check(cell, &t)) {
            return self.hyperlink(cell, style, col, decimal, may_hook);
        }
        // Anything else (numpy scalars, times, rich text...): the hook, which
        // gives the value the Python writer would write, or the Python writer.
        if may_hook {
            return self.hooked(cell, style, col, decimal, false);
        }
        Ok(Cell::Fallback)
    }

    /// Formula cell, mirroring `_write_formula_cell`: the cell's own style only,
    /// then the formula text and the cached value. Array formulas and values
    /// the native code does not reproduce exactly fall back, and nothing of a
    /// fallen-back cell is kept.
    fn formula(&mut self, cell: &Bound<'_, PyAny>, style: u32) -> PyResult<Cell> {
        let start = self.out.len();
        if self.formula_xml(cell, style)? {
            Ok(Cell::Done)
        } else {
            self.out.truncate(start);
            Ok(Cell::Fallback)
        }
    }

    /// False when the cell must go to the Python writer (the caller drops any
    /// partial output).
    fn formula_xml(&mut self, cell: &Bound<'_, PyAny>, style: u32) -> PyResult<bool> {
        let py = cell.py();
        if !cell.getattr("array_ref")?.is_none() {
            return Ok(false);
        }
        // A dynamic-array formula needs the anchor's cm attribute and the
        // metadata part: the Python writer writes the row.
        if !cell.getattr("spill_ref")?.is_none() {
            return Ok(false);
        }
        let text_obj = cell.getattr("text")?;
        let Ok(text) = text_obj.cast::<PyString>() else {
            return Ok(false);
        };
        let Ok(text) = text.to_str() else {
            return Ok(false); // lone surrogates
        };
        let value = cell.getattr("value")?;
        self.out.extend_from_slice(b"<c");
        style_attr(&mut self.out, style);
        if value.is_none() {
            self.out.extend_from_slice(b"><f>");
            write_plain_escaped(&mut self.out, text.as_bytes());
            self.out.extend_from_slice(b"</f></c>");
            return Ok(true);
        }
        if let Ok(s) = value.cast::<PyString>() {
            // t="str" with the value escaped as a string cell's text is.
            if !text_fits(s)? {
                return Ok(false); // over the cell limit: the Python writer raises
            }
            let Ok(v) = s.to_str() else {
                return Ok(false);
            };
            self.out.extend_from_slice(b" t=\"str\"><f>");
            write_plain_escaped(&mut self.out, text.as_bytes());
            self.out.extend_from_slice(b"</f><v>");
            if !write_escaped(&mut self.out, v.as_bytes()) {
                return Ok(false);
            }
            self.out.extend_from_slice(b"</v></c>");
            return Ok(true);
        }
        self.out.extend_from_slice(b"><f>");
        write_plain_escaped(&mut self.out, text.as_bytes());
        self.out.extend_from_slice(b"</f><v>");
        if value.is_exact_instance_of::<PyInt>() {
            // _format_number: str(int) of the value, here within int64.
            let Some(v) = int_as_i64(&value) else {
                return Ok(false);
            };
            self.out.extend_from_slice(v.to_string().as_bytes());
        } else if let Ok(f) = value.cast::<PyFloat>() {
            // Exact floats only: a subclass may print differently from float.
            if !value.is_exact_instance_of::<PyFloat>() {
                return Ok(false);
            }
            let d = f.value();
            if !d.is_finite() {
                return Ok(false);
            }
            if d.fract() == 0.0 {
                // _format_number: str(int(d)), exact below 2**53.
                if d.abs() >= 9_007_199_254_740_992.0 {
                    return Ok(false);
                }
                self.out
                    .extend_from_slice((d as i64).to_string().as_bytes());
            } else {
                write_repr(py, &mut self.out, d)?;
            }
        } else {
            // Decimal (converted by the Python writer), bool and other types.
            return Ok(false);
        }
        self.out.extend_from_slice(b"</v></c>");
        Ok(true)
    }
}

/// `float(cell)`, or None when it raises.
fn py_float(cell: &Bound<'_, PyAny>) -> Option<f64> {
    // SAFETY: `cell` is live with the GIL held.
    let f = unsafe { pyo3::ffi::PyNumber_Float(cell.as_ptr()) };
    // SAFETY: PyNumber_Float returns a new reference or NULL with an exception set.
    match unsafe { Bound::from_owned_ptr_or_err(cell.py(), f) } {
        // SAFETY: on success PyNumber_Float returns a float (or float subclass) instance.
        Ok(f) => Some(unsafe { f.cast_into_unchecked::<PyFloat>() }.value()),
        Err(_) => None,
    }
}

/// `float(cell)` of an int: its value, OverflowError (past the float range,
/// written as text), or any other error (the Python writer raises it).
enum IntFloat {
    Value(f64),
    Overflow,
    Error,
}

fn int_float(cell: &Bound<'_, PyAny>) -> IntFloat {
    // SAFETY: `cell` is live with the GIL held.
    let f = unsafe { pyo3::ffi::PyNumber_Float(cell.as_ptr()) };
    // SAFETY: PyNumber_Float returns a new reference or NULL with an exception set.
    match unsafe { Bound::from_owned_ptr_or_err(cell.py(), f) } {
        // SAFETY: on success PyNumber_Float returns a float (or float subclass) instance.
        Ok(f) => IntFloat::Value(unsafe { f.cast_into_unchecked::<PyFloat>() }.value()),
        Err(e) if e.is_instance_of::<pyo3::exceptions::PyOverflowError>(cell.py()) => {
            IntFloat::Overflow
        }
        Err(_) => IntFloat::Error,
    }
}

/// (row_xml, sst_uniq, sst_all, hyperlinks) as returned to the Python writer.
type EncodedRow<'py> = (Bound<'py, PyBytes>, u32, u32, Bound<'py, PyList>);

/// Encode one XLSX sheet row.  Returns (row_xml, sst_uniq, sst_all, hyperlinks), or None
/// when the Python row writer must handle the row; then the strings this
/// call added to the table are removed again.
///
/// ``write_r`` False leaves the ``r`` attribute off the ``<row>``. The row then
/// takes the number after the previous row, which is how the writer keeps dense
/// rows small; the caller sets it only for a row that follows a gap.
#[pyfunction]
#[pyo3(signature = (row, ss_dict, ss_list, sst_uniq, sst_all, row_idx, use_shared_strings, style_ids=None, fmt_xf=None, row_attrs=None, link_xf=None, write_r=true, cell_hook=None))]
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
    link_xf: Option<&Bound<'py, PyAny>>,
    write_r: bool,
    cell_hook: Option<&Bound<'py, PyAny>>,
) -> PyResult<Option<EncodedRow<'py>>> {
    let link_xf = match link_xf {
        Some(x) if !x.is_none() => x
            .extract::<u32>()
            .map_err(|_| PyTypeError::new_err("link_xf must be an int or None"))?,
        _ => 0,
    };
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

    // A non-sequence row goes to the Python writer.
    let Ok(seq) = FastSeq::new(row, c"row must be a sequence") else {
        return Ok(None);
    };
    let n = seq.len();
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
        link_xf,
        links: Vec::new(),
        hook: cell_hook
            .cloned()
            .unwrap_or_else(|| py.None().into_bound(py)),
        row_idx,
        hook_style: 0,
    };
    row.out.extend_from_slice(b"<row");
    if write_r {
        row.out.extend_from_slice(b" r=\"");
        write_u32(&mut row.out, row_idx.wrapping_add(1));
        row.out.push(b'"');
    }
    if let Some(attrs) = &attrs {
        row.out.extend_from_slice(attrs.as_bytes());
    }
    row.out.push(b'>');

    let result = (|| -> PyResult<bool> {
        for i in 0..n {
            // The body below can run Python code (hashing a format key, conversions)
            // that shrinks a list row, so the bound is checked before every read.
            let Some(mut cell) = seq.get(i) else {
                return Ok(false);
            };
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
            // The hook sees the style before a date's variant: a date it is
            // asked about (pandas' NaT) is no date to the Python writer.
            row.hook_style = style;
            if !explicit_fmt && !date_style_xf(&cell, fmt_xf, &mut style)? {
                return Ok(false);
            }
            if let Cell::Fallback = row.cell(&cell, style, i as u32, decimal)? {
                return Ok(false);
            }
        }
        Ok(true)
    })();

    match result {
        Ok(true) => {
            row.out.extend_from_slice(b"</row>");
            let (uniq, all) = row.sst.as_ref().map_or((uniq, all), |s| (s.uniq, s.all));
            // (row, column, Hyperlink) for the sheet's <hyperlinks>, as the Python writer records them.
            let links = PyList::empty(py);
            for (col, link) in row.links {
                links.append((row_idx, col, link.bind(py)))?;
            }
            // UTF-8 by construction (ASCII markup and UTF-8 text from str), so the
            // Python writer can write the bytes to the sheet part as they are.
            Ok(Some((PyBytes::new(py, &row.out), uniq, all, links)))
        }
        other => {
            rollback(dict, list, list_before);
            other.map(|_| None)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn escaped(s: &str) -> Option<String> {
        let mut out = Vec::new();
        write_escaped(&mut out, s.as_bytes()).then(|| String::from_utf8(out).unwrap())
    }

    #[test]
    fn markup_characters_become_entities() {
        assert_eq!(escaped("a<b & c>d").as_deref(), Some("a&lt;b &amp; c&gt;d"));
        assert_eq!(
            escaped("tab\tand\nnewline").as_deref(),
            Some("tab\tand\nnewline")
        );
        assert_eq!(escaped("zażółć").as_deref(), Some("zażółć"));
    }

    #[test]
    fn text_needing_an_xhhhh_escape_is_left_to_python() {
        assert_eq!(escaped("bell\u{7}"), None);
        assert_eq!(escaped("cr\r"), None);
        assert_eq!(escaped("_x0041_"), None);
        assert_eq!(escaped("\u{FFFE}"), None);
        assert_eq!(escaped("\u{FFFF}"), None);
        // Not of the _xHHHH_ form, so written as is.
        assert_eq!(escaped("_x004G_").as_deref(), Some("_x004G_"));
        assert_eq!(escaped("_x0041").as_deref(), Some("_x0041"));
    }

    #[test]
    fn formula_text_escapes_markup_only() {
        let mut out = Vec::new();
        write_plain_escaped(&mut out, b"IF(A1<>\"_x0041_\",1,2)&\x07");
        assert_eq!(out, b"IF(A1&lt;&gt;\"_x0041_\",1,2)&amp;\x07");
    }

    #[test]
    fn value_cell_opens_with_its_style() {
        let mut out = Vec::new();
        open_value(&mut out, 0);
        open_value(&mut out, 12);
        assert_eq!(out, b"<c><v><c s=\"12\"><v>");
    }
}
