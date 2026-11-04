//! The cells of one DataFrame row for the writers (`_dataframe_rows.dataframe_to_rows`).
//!
//! The plain Python values a frame yields (str, int, bool, float, None, pandas'
//! NA and NaT) are decided here; every other cell goes to the Python normaliser
//! one cell at a time, so the result is what the Python loop gives.

use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyList, PyString, PyTuple};

use crate::pyseq::FastSeq;

/// The row's cells, normalised: a missing value becomes `na_rep`, a cell in a column
/// with a format becomes `(value, format)`. `missing` holds the objects that are
/// missing by identity (`pd.NA`, `pd.NaT`), `timestamp` is `pd.Timestamp` (or None),
/// whose values become `datetime` through `to_pydatetime()`; `slow(cell, format)` gives
/// the finished cell for any other type. None when the row is no sequence (the caller
/// converts it).
#[pyfunction]
pub fn convert_frame_row<'py>(
    row: &Bound<'py, PyAny>,
    formats: &Bound<'py, PyList>,
    na_rep: &Bound<'py, PyAny>,
    missing: &Bound<'py, PyTuple>,
    timestamp: &Bound<'py, PyAny>,
    slow: &Bound<'py, PyAny>,
) -> PyResult<Option<Bound<'py, PyList>>> {
    let py = row.py();
    let Ok(row) = FastSeq::new(row, c"row must be a sequence") else {
        PyErr::take(py);
        return Ok(None);
    };
    let n = row.len().min(formats.len());
    let out = PyList::empty(py);
    for i in 0..n {
        let Some(cell) = row.get(i) else {
            return Ok(None);
        };
        let format = formats.get_item(i)?;
        let value = if cell.is_none() {
            None
        } else if cell.is_exact_instance_of::<PyString>()
            || cell.is_exact_instance_of::<PyBool>()
            || cell.is_exact_instance_of::<PyInt>()
        {
            Some(cell)
        } else if let Ok(f) = cell.cast_exact::<PyFloat>() {
            let v = f.value();
            #[allow(clippy::eq_op)]
            if v != v { None } else { Some(cell) }
        } else if missing.iter().any(|m| m.is(&cell)) {
            None
        } else if cell.get_type().is(timestamp) {
            // pd.Timestamp: what `_normalize_cell` does for a datetime.
            match cell.call_method0(pyo3::intern!(py, "to_pydatetime")) {
                Ok(moment) => Some(moment),
                Err(_) => {
                    out.append(slow.call1((cell, format))?)?;
                    continue;
                }
            }
        } else {
            out.append(slow.call1((cell, format))?)?;
            continue;
        };
        match value {
            None => out.append(na_rep)?,
            Some(value) if format.is_none() => out.append(value)?,
            Some(value) => out.append(PyTuple::new(py, [value, format])?)?,
        }
    }
    Ok(Some(out))
}
