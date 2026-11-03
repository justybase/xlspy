//! Cell helpers shared by the XLSB and XLSX row encoders: the shared-string
//! table, `(value, format)` / `(value, Style)` cells, date styles and serials.

use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{
    PyDate, PyDateAccess, PyDateTime, PyDict, PyInt, PyList, PyString, PyTimeAccess, PyTuple,
    PyType,
};

use crate::dates::days_from_civil;
use crate::text::CELL_TEXT_LIMIT;
use crate::types::overlay_kind;

/// Result of encoding one cell: written, or the whole row goes to Python.
pub enum Cell {
    Done,
    Fallback,
}

/// True when the str fits in a cell (at most 32767 UTF-16 units).  Each code
/// point is one or two units, so only text longer than half the limit needs
/// counting, and only 4-byte-kind strings hold two-unit characters.
pub fn text_fits(s: &Bound<'_, PyString>) -> PyResult<bool> {
    let len = s.len()?;
    if len <= CELL_TEXT_LIMIT / 2 {
        return Ok(true);
    }
    if len > CELL_TEXT_LIMIT {
        return Ok(false);
    }
    Ok(crate::text::CodePoints::of(s)?.utf16_len() <= CELL_TEXT_LIMIT)
}

/// The shared-string table of one encode call: the caller's dict and list,
/// and the counters, which are only handed back when the row is encoded.
pub struct Sst<'a, 'py> {
    pub dict: &'a Bound<'py, PyDict>,
    pub list: &'a Bound<'py, PyList>,
    pub uniq: u32,
    pub all: u32,
}

impl Sst<'_, '_> {
    /// Index of `s`, added to the table when new (list first, so a failed
    /// insert leaves only a list entry, which the rollback removes).
    pub fn index(&mut self, s: &Bound<'_, PyString>) -> PyResult<u32> {
        let idx = match self.dict.get_item(s)? {
            // The writer stores the indices it hands out; anything else is a
            // corrupted table, which must not reach the file as index 0xFFFFFFFF.
            Some(idx) => idx.extract::<u32>()?,
            None => {
                let idx = self.uniq;
                self.list.append(s)?;
                self.dict.set_item(s, idx)?;
                self.uniq += 1;
                idx
            }
        };
        self.all += 1;
        Ok(idx)
    }
}

/// Remove the entries a failed or declined row added after `len_before`, so
/// the caller's counters stay valid.  Entries already there are never touched.
pub fn rollback(dict: &Bound<'_, PyDict>, list: &Bound<'_, PyList>, len_before: usize) {
    let n = list.len();
    if n <= len_before {
        return;
    }
    for i in len_before..n {
        if let Ok(key) = list.get_item(i) {
            let _ = dict.del_item(key); // a key never inserted: nothing to undo
        }
    }
    let _ = list.del_slice(len_before, usize::MAX);
}

/// Exact `_is_formatted_cell`: a tuple of length 2 with a str or Style second
/// item.  Returns its overlay kind (0 when the cell is not formatted).
pub fn fmt_tuple_kind(cell: &Bound<'_, PyAny>, style: Option<&Bound<'_, PyType>>) -> u8 {
    if !cell.is_instance_of::<PyTuple>() {
        return 0;
    }
    match cell.len() {
        Ok(2) => {}
        _ => return 0,
    }
    match cell.get_item(1) {
        Ok(fmt) => overlay_kind(&fmt, style),
        Err(_) => 0,
    }
}

fn get_xf<'py>(fmt_xf: &Bound<'py, PyDict>, key: impl IntoPyObject<'py>) -> PyResult<Option<u32>> {
    match fmt_xf.get_item(key)? {
        Some(xf) => Ok(Some(as_unsigned_long(&xf)?)),
        None => Ok(None),
    }
}

/// A style or xf id: `PyLong_AsUnsignedLong` on a 32-bit `unsigned long`,
/// as the C extension behaved on Windows.  Negative values and values past
/// `u32::MAX` raise OverflowError on every platform (C on 64-bit Linux
/// silently truncated the latter).
pub fn as_unsigned_long(obj: &Bound<'_, PyAny>) -> PyResult<u32> {
    obj.extract::<u32>()
}

/// Resolve a (value, format) or (value, Style) cell through `fmt_xf`.  The
/// key is the format string, or the Style's id(), for an unstyled cell and
/// (style id, that key) for a styled one.  On success `style` becomes the xf
/// and the cell's value is returned; None (row fallback) when the key is
/// missing, the value is None, or a nested tuple.
pub fn fmt_xf_value<'py>(
    cell: &Bound<'py, PyAny>,
    fmt_xf: Option<&Bound<'py, PyDict>>,
    style: &mut u32,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    let py = cell.py();
    let Some(fmt_xf) = fmt_xf else {
        return Ok(None);
    };
    let Ok(fmt) = cell.get_item(1) else {
        return Ok(None);
    };
    let fmt = if fmt.is_instance_of::<PyString>() {
        fmt
    } else {
        (fmt.as_ptr() as usize).into_pyobject(py)?.into_any()
    };
    let xf = if *style != 0 {
        get_xf(
            fmt_xf,
            PyTuple::new(py, [style.into_pyobject(py)?.into_any(), fmt])?,
        )?
    } else {
        get_xf(fmt_xf, fmt)?
    };
    let Some(xf) = xf else { return Ok(None) };
    let Ok(inner) = cell.get_item(0) else {
        return Ok(None);
    };
    if inner.is_none() || inner.is_instance_of::<PyTuple>() {
        return Ok(None);
    }
    *style = xf;
    Ok(Some(inner))
}

/// Resolve the xf of a styled date or datetime cell through `fmt_xf`: the
/// Python writer records the variant with the default date format under
/// ("date" | "datetime", style xf).  False: row fallback.
pub fn date_style_xf(
    cell: &Bound<'_, PyAny>,
    fmt_xf: Option<&Bound<'_, PyDict>>,
    style: &mut u32,
) -> PyResult<bool> {
    if *style == 0 || !cell.is_instance_of::<PyDate>() {
        return Ok(true);
    }
    let kind = if cell.is_instance_of::<PyDateTime>() {
        "datetime"
    } else {
        "date"
    };
    let Some(fmt_xf) = fmt_xf else {
        return Ok(false);
    };
    match get_xf(fmt_xf, (kind, *style))? {
        Some(xf) => {
            *style = xf;
            Ok(true)
        }
        None => Ok(false),
    }
}

pub enum Serial {
    Value(f64),
    /// Year outside 1900..=9999: the writers store str(cell).
    OutOfRange,
    /// Unreadable fields, or beyond exact-double microseconds: row fallback.
    Fallback,
}

fn long_attr(cell: &Bound<'_, PyAny>, name: &str) -> Option<i64> {
    cell.getattr(name).ok()?.extract::<i64>().ok()
}

/// Excel serial bit-identical to `datetime_to_serial`, which computes
/// `(moment - epoch).total_seconds() / 86400.0` with exact integer
/// microseconds.
pub fn excel_serial(cell: &Bound<'_, PyAny>, is_datetime: bool) -> Serial {
    let (mut hh, mut mi, mut ss, mut us) = (0i64, 0i64, 0i64, 0i64);
    let exact = if is_datetime {
        cell.is_exact_instance_of::<PyDateTime>()
    } else {
        cell.is_exact_instance_of::<PyDate>()
    };
    let (y, mo, d);
    if exact {
        // SAFETY: `exact` checked that `cell` is exactly a date or datetime,
        // and datetime is a date subclass.
        let date = unsafe { cell.cast_unchecked::<PyDate>() };
        y = i64::from(date.get_year());
        mo = i64::from(date.get_month());
        d = i64::from(date.get_day());
        if is_datetime {
            // SAFETY: with `is_datetime`, `exact` checked for an exact datetime.
            let dt = unsafe { cell.cast_unchecked::<PyDateTime>() };
            hh = i64::from(dt.get_hour());
            mi = i64::from(dt.get_minute());
            ss = i64::from(dt.get_second());
            us = i64::from(dt.get_microsecond());
        }
    } else {
        let (Some(a), Some(b), Some(c)) = (
            long_attr(cell, "year"),
            long_attr(cell, "month"),
            long_attr(cell, "day"),
        ) else {
            return Serial::Fallback;
        };
        (y, mo, d) = (a, b, c);
        if is_datetime {
            let (Some(a), Some(b), Some(c), Some(e)) = (
                long_attr(cell, "hour"),
                long_attr(cell, "minute"),
                long_attr(cell, "second"),
                long_attr(cell, "microsecond"),
            ) else {
                return Serial::Fallback;
            };
            (hh, mi, ss, us) = (a, b, c, e);
        }
    }
    if !(1900..=9999).contains(&y) {
        return Serial::OutOfRange;
    }
    if !(1..=12).contains(&mo) || !(1..=31).contains(&d) {
        return Serial::Fallback;
    }
    if is_datetime
        && (!(0..=23).contains(&hh)
            || !(0..=59).contains(&mi)
            || !(0..=61).contains(&ss)
            || !(0..=999_999).contains(&us))
    {
        return Serial::Fallback;
    }
    let ord = days_from_civil(y, mo, d) + 719_163;
    let epoch_ord = if y == 1900 && mo < 3 {
        693_595
    } else {
        693_594
    };
    let whole = (ord - epoch_ord) * 86_400 + hh * 3600 + mi * 60 + ss;
    let total_us = whole * 1_000_000 + us;
    if !(-9_007_199_254_740_992..=9_007_199_254_740_992).contains(&total_us) {
        return Serial::Fallback;
    }
    Serial::Value(total_us as f64 / 1_000_000.0 / 86_400.0)
}

/// Per-cell style source: one style for the row, a dense list, or sparse
/// (column, style_id) pairs in column order.
pub struct Styles<'a, 'py> {
    ids: &'a Bound<'py, PyAny>,
    row_style: u32,
    len: usize,
    pair_index: usize,
    /// A bad dense-list entry raises (XLSX) instead of meaning no style (XLSB).
    strict: bool,
}

impl<'a, 'py> Styles<'a, 'py> {
    pub fn new(ids: &'a Bound<'py, PyAny>, strict: bool) -> PyResult<Self> {
        if !(ids.is_none()
            || ids.is_instance_of::<PyList>()
            || ids.is_instance_of::<PyTuple>()
            || ids.is_instance_of::<PyInt>())
        {
            return Err(PyTypeError::new_err(
                "style_ids must be an integer, list, tuple, or None",
            ));
        }
        let row_style = if ids.is_instance_of::<PyInt>() {
            as_unsigned_long(ids)?
        } else {
            0
        };
        // Only lists and tuples have a length here; len(None) would raise, and
        // an exception per row is expensive.
        let len = match (ids.cast::<PyList>(), ids.cast::<PyTuple>()) {
            (Ok(list), _) => list.len(),
            (_, Ok(tuple)) => tuple.len(),
            _ => 0,
        };
        Ok(Styles {
            ids,
            row_style,
            len,
            pair_index: 0,
            strict,
        })
    }

    pub fn for_column(&mut self, col: usize) -> PyResult<u32> {
        if self.ids.is_none() {
            return Ok(0);
        }
        if self.row_style != 0 {
            return Ok(self.row_style);
        }
        if let Ok(list) = self.ids.cast::<PyList>() {
            if col < self.len
                && let Ok(item) = list.get_item(col)
                && !item.is_none()
            {
                return match as_unsigned_long(&item) {
                    Ok(v) => Ok(v),
                    Err(err) if self.strict => Err(err),
                    Err(_) => Ok(0),
                };
            }
            return Ok(0);
        }
        let Ok(pairs) = self.ids.cast::<PyTuple>() else {
            return Ok(0);
        };
        while self.pair_index < self.len {
            let pair = pairs.get_item(self.pair_index)?;
            let pair = match pair.cast::<PyTuple>() {
                Ok(p) if p.len() == 2 => p.clone(),
                _ => {
                    return Err(PyTypeError::new_err(
                        "sparse style entries must be (column, style_id) pairs",
                    ));
                }
            };
            let style_col: i64 = pair.get_item(0)?.extract()?;
            if style_col < col as i64 {
                self.pair_index += 1;
                continue;
            }
            if style_col == col as i64 {
                let v = as_unsigned_long(&pair.get_item(1)?)?;
                self.pair_index += 1;
                return Ok(v);
            }
            break;
        }
        Ok(0)
    }
}
