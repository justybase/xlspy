//! Column-width estimate and the formatted-cell pre-scan of the writers.

use crate::pyseq::FastSeq;
use pyo3::exceptions::{PySystemError, PyTypeError};
use pyo3::prelude::*;
use pyo3::types::{PyDateTime, PyFloat, PyList, PyTuple, PyType};

use crate::text::CodePoints;
use crate::types::{decimal_type, overlay_kind, style_type};

/// `cell[0]` of a `(value, format)` / `(value, Style)` tuple, `Some(cell)` for
/// any other cell, `None` when the cell is skipped (errors are swallowed).
fn unwrap_overlay<'py>(
    cell: Bound<'py, PyAny>,
    style: Option<&Bound<'py, PyType>>,
) -> Option<Bound<'py, PyAny>> {
    if !cell.is_instance_of::<PyTuple>() {
        return Some(cell);
    }
    if cell.len().ok()? != 2 {
        return Some(cell);
    }
    let fmt = cell.get_item(1).ok()?;
    if overlay_kind(&fmt, style) == 0 {
        return Some(cell);
    }
    cell.get_item(0).ok()
}

/// A numpy scalar of kind int, unsigned, bool, or float16/float32, as `.item()`
/// gives it (the Python value the writers store); any other cell as it is.
/// numpy is not imported: the class's module and the scalar's dtype identify
/// it.  The Python twin is xlspy/_scalars.py.
fn numpy_item<'py>(cell: Bound<'py, PyAny>) -> Bound<'py, PyAny> {
    let py = cell.py();
    let is_numpy = cell
        .get_type()
        .as_any()
        .getattr(pyo3::intern!(py, "__module__"))
        .and_then(|m| m.extract::<String>())
        .is_ok_and(|m| m == "numpy");
    if !is_numpy {
        return cell;
    }
    let converts = (|| -> PyResult<bool> {
        if cell.getattr(pyo3::intern!(py, "shape"))?.len()? != 0 {
            return Ok(false);
        }
        let dtype = cell.getattr(pyo3::intern!(py, "dtype"))?;
        let kind: String = dtype.getattr(pyo3::intern!(py, "kind"))?.extract()?;
        let itemsize: usize = dtype.getattr(pyo3::intern!(py, "itemsize"))?.extract()?;
        Ok(match kind.as_str() {
            "i" | "u" | "b" => true,
            "f" => itemsize < 8,
            _ => false,
        })
    })()
    .unwrap_or(false);
    if !converts {
        return cell;
    }
    cell.call_method0(pyo3::intern!(py, "item"))
        .unwrap_or_else(|_| cell.clone())
}

fn cell_width(cell: &Bound<'_, PyAny>) -> Option<f64> {
    if cell.is_instance_of::<PyDateTime>() {
        return Some(18.0);
    }
    let s = cell.str().ok()?;
    let len = CodePoints::of(&s).ok()?.first_line_len();
    Some(1.25 * len as f64 + 4.0)
}

/// Calculate column widths from row data.
///
/// Mirrors the Python fallback: unwrap (value, format) tuples, measure
/// Decimal as float(cell) and numpy scalars as their item(), datetimes as 18,
/// anything else by the code points of the first line of str(cell); cells that
/// raise are skipped.
#[pyfunction]
pub fn calc_column_widths<'py>(
    py: Python<'py>,
    rows: &Bound<'py, PyAny>,
    max_cols: i32,
) -> PyResult<Bound<'py, PyList>> {
    let rows = rows
        .cast::<PyList>()
        .map_err(|_| PyTypeError::new_err("rows must be a list"))?;
    if max_cols < 0 {
        return Err(PySystemError::new_err("bad argument to internal function"));
    }
    let style = style_type(py);
    let decimal = decimal_type(py);
    let nrows = rows.len();
    let mut widths = Vec::with_capacity(max_cols as usize);
    for i in 0..max_cols as usize {
        let mut max_width = 0.0f64;
        for r in 0..nrows {
            let Ok(row) = rows.get_item(r) else { continue };
            let Ok(row) = row.cast::<PyList>() else {
                continue;
            };
            let Ok(cell) = row.get_item(i) else { continue };
            if cell.is_none() {
                continue;
            }
            let Some(mut cell) = unwrap_overlay(cell, style.as_ref()) else {
                continue;
            };
            if cell.is_none() {
                continue;
            }
            if decimal.is_some_and(|d| cell.is_instance(d).unwrap_or(false)) {
                let Ok(f) = cell.extract::<f64>() else {
                    continue;
                };
                cell = PyFloat::new(py, f).into_any();
            }
            cell = numpy_item(cell);
            if let Some(w) = cell_width(&cell) {
                max_width = max_width.max(w);
            }
        }
        if max_width == 0.0 {
            max_width = 4.0;
        }
        widths.push(max_width.min(255.0));
    }
    PyList::new(py, widths)
}

/// Return True if a row contains any (value, format) or (value, Style) cell,
/// or None when the row cannot be scanned here (the Python check then raises
/// the same errors).
#[pyfunction]
pub fn xlsb_row_has_formatted<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
) -> PyResult<Option<bool>> {
    let Ok(seq) = FastSeq::new(row, c"row must be a sequence") else {
        return Ok(None);
    };
    let n = seq.len();
    let mut style = None; // looked up at the first tuple: most rows have none
    for i in 0..n {
        // Python code runs in this loop (len, get_item, overlay_kind) and can
        // shrink a list row, so the bound is checked before every read.
        let Some(cell) = seq.get(i) else {
            return Ok(None);
        };
        if cell.is_none() || !cell.is_instance_of::<PyTuple>() {
            continue;
        }
        let Ok(len) = cell.len() else { return Ok(None) };
        if len != 2 {
            continue;
        }
        let Ok(fmt) = cell.get_item(1) else {
            return Ok(None);
        };
        let style = style.get_or_insert_with(|| style_type(py));
        if overlay_kind(&fmt, style.as_ref()) != 0 {
            return Ok(Some(true));
        }
    }
    Ok(Some(false))
}
