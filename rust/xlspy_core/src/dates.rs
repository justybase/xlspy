//! Excel serial -> datetime, as `excel_dates.serial_to_datetime`.

use pyo3::exceptions::{PyOverflowError, PyValueError};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyDateTime, PyDict, PyType};

static TIMEDELTA: PyOnceLock<Py<PyType>> = PyOnceLock::new();

/// Howard Hinnant's days_from_civil (days since 1970-01-01).
pub fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let yy = if m <= 2 { y - 1 } else { y };
    let era = if yy >= 0 { yy } else { yy - 399 } / 400;
    let yoe = yy - era * 400;
    let mp = if m > 2 { m - 3 } else { m + 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// Inverse of [`days_from_civil`]: (year, month, day).
pub fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let yy = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { yy + 1 } else { yy }, m, d)
}

const DAY_US: i64 = 86_400_000_000;

/// `serial_to_datetime`: datetime + timedelta(days=serial), with CPython's
/// trick of keeping the fractional day separate.  Raises what Python raises.
pub fn serial_datetime(py: Python<'_>, serial: f64, date_1904: bool) -> PyResult<Bound<'_, PyAny>> {
    let mut s = serial;
    if !date_1904 && (1.0..60.0).contains(&s) {
        s += 1.0;
    }
    if !s.is_finite() {
        return Err(PyValueError::new_err("cannot convert float NaN to integer"));
    }
    let whole = s.floor();
    if !(-999_999_999.0..=999_999_999.0).contains(&whole) {
        return Err(PyOverflowError::new_err("days out of range for timedelta"));
    }
    let mut us = ((s - whole) * 86_400_000_000.0).round_ties_even() as i64;
    let mut days = whole as i64;
    if us >= DAY_US {
        days += 1;
        us -= DAY_US;
    } else if us < 0 {
        days -= 1;
        us += DAY_US;
    }
    let epoch = if date_1904 {
        days_from_civil(1904, 1, 1)
    } else {
        days_from_civil(1899, 12, 30)
    } + 719_163;
    let ordinal = epoch + days;
    if ordinal < 1 || ordinal > days_from_civil(9999, 12, 31) + 719_163 {
        return Err(PyOverflowError::new_err("date value out of range"));
    }
    let (y, m, d) = civil_from_days(ordinal - 719_163);
    let secs = us / 1_000_000;
    Ok(PyDateTime::new(
        py,
        y as i32,
        m as u8,
        d as u8,
        (secs / 3600) as u8,
        ((secs / 60) % 60) as u8,
        (secs % 60) as u8,
        (us % 1_000_000) as u32,
        None,
    )?
    .into_any())
}

/// The reference conversion for serials outside the fast range, so NaN,
/// infinities and out-of-range values raise the same exception and message.
fn serial_datetime_slow(
    py: Python<'_>,
    serial: f64,
    date_1904: bool,
) -> PyResult<Bound<'_, PyAny>> {
    let epoch = if date_1904 {
        PyDateTime::new(py, 1904, 1, 1, 0, 0, 0, 0, None)?
    } else {
        PyDateTime::new(py, 1899, 12, 30, 0, 0, 0, 0, None)?
    };
    let kwargs = PyDict::new(py);
    kwargs.set_item("days", serial)?;
    let td = TIMEDELTA
        .import(py, "datetime", "timedelta")?
        .call((), Some(&kwargs))?;
    epoch.add(td)
}

/// XLSB date cell: the value of serial_to_datetime().
pub fn xlsb_date_value(py: Python<'_>, serial: f64, date_1904: bool) -> PyResult<Bound<'_, PyAny>> {
    if serial.abs() < 999_999_999.0 {
        serial_datetime(py, serial, date_1904)
    } else {
        serial_datetime_slow(py, serial, date_1904)
    }
}

#[cfg(test)]
mod tests {
    use super::{civil_from_days, days_from_civil};

    #[test]
    fn civil_roundtrip() {
        for z in [-719_162i64, -1, 0, 1, 11_017, 2_932_896] {
            let (y, m, d) = civil_from_days(z);
            assert_eq!(days_from_civil(y, m, d), z);
        }
        assert_eq!(civil_from_days(0), (1970, 1, 1));
    }
}
