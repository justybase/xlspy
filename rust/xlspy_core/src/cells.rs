//! Cell helpers shared by the XLSB and XLSX row encoders: the shared-string
//! table, `(value, format)` / `(value, Style)` cells, date styles and serials.

use pyo3::prelude::*;
use pyo3::types::{
    PyDate, PyDateAccess, PyDateTime, PyDict, PyList, PyString, PyTimeAccess, PyTuple, PyType,
};

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
            Some(idx) => idx.extract::<i64>().unwrap_or(-1) as u32,
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

/// Howard Hinnant's days_from_civil (days since 1970-01-01).
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let yy = if m <= 2 { y - 1 } else { y };
    let era = if yy >= 0 { yy } else { yy - 399 } / 400;
    let yoe = yy - era * 400;
    let mp = if m > 2 { m - 3 } else { m + 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
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
        let date = unsafe { cell.cast_unchecked::<PyDate>() };
        y = i64::from(date.get_year());
        mo = i64::from(date.get_month());
        d = i64::from(date.get_day());
        if is_datetime {
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

#[cfg(test)]
mod tests {
    use super::days_from_civil;

    #[test]
    fn civil_days() {
        assert_eq!(days_from_civil(1970, 1, 1), 0);
        assert_eq!(days_from_civil(2000, 3, 1), 11_017);
        assert_eq!(days_from_civil(1899, 12, 30) + 719_163, 693_594);
    }
}
