//! The cells of one DataFrame row for the writers (`_dataframe_rows.dataframe_to_rows`).
//!
//! The plain Python values a frame yields (str, int, bool, float, None, pandas'
//! NA and NaT) are decided here; every other cell goes to the Python normaliser
//! one cell at a time, so the result is what the Python loop gives.

use pyo3::exceptions::PyTypeError;
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
    // A TypeError means "not a sequence": the caller converts it.  Any other
    // error (a custom iterable's __iter__, a MemoryError) must not be hidden.
    let row = match FastSeq::new(row, c"row must be a sequence") {
        Ok(row) => row,
        Err(error) if error.is_instance_of::<PyTypeError>(py) => return Ok(None),
        Err(error) => return Err(error),
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

#[cfg(test)]
mod tests {
    use super::*;
    use pyo3::types::{PyList, PyTuple};

    #[test]
    fn plain_values_and_nan_are_normalised() {
        Python::initialize();
        Python::attach(|py| {
            let row = PyList::empty(py);
            row.append("a").unwrap();
            row.append(1i64).unwrap();
            row.append(f64::NAN).unwrap();
            let formats = PyList::empty(py);
            for _ in 0..3 {
                formats.append(py.None()).unwrap();
            }
            let na_rep = "NA".into_pyobject(py).unwrap().into_any();
            let missing = PyTuple::empty(py);
            let nothing = py.None().into_bound(py);
            let result = convert_frame_row(
                row.as_any(),
                &formats,
                &na_rep,
                &missing,
                &nothing,
                &nothing,
            )
            .unwrap()
            .unwrap();
            assert_eq!(
                result.get_item(0).unwrap().extract::<String>().unwrap(),
                "a"
            );
            assert_eq!(result.get_item(1).unwrap().extract::<i64>().unwrap(), 1);
            assert_eq!(
                result.get_item(2).unwrap().extract::<String>().unwrap(),
                "NA"
            );
        });
    }

    #[test]
    fn a_non_sequence_is_left_to_the_caller() {
        Python::initialize();
        Python::attach(|py| {
            let row = 5i64.into_pyobject(py).unwrap().into_any();
            let formats = PyList::empty(py);
            let na_rep = py.None().into_bound(py);
            let missing = PyTuple::empty(py);
            assert!(
                convert_frame_row(row.as_any(), &formats, &na_rep, &missing, &na_rep, &na_rep)
                    .unwrap()
                    .is_none()
            );
        });
    }

    #[test]
    fn an_iterators_own_error_is_not_hidden() {
        Python::initialize();
        Python::attach(|py| {
            let globals = pyo3::types::PyDict::new(py);
            py.run(
                c"class Broken:\n    def __iter__(self):\n        raise RuntimeError('boom')\n",
                Some(&globals),
                None,
            )
            .unwrap();
            let row = globals
                .get_item("Broken")
                .unwrap()
                .unwrap()
                .call0()
                .unwrap();
            let formats = PyList::empty(py);
            let na_rep = py.None().into_bound(py);
            let missing = PyTuple::empty(py);
            let error =
                convert_frame_row(&row, &formats, &na_rep, &missing, &na_rep, &na_rep).unwrap_err();
            assert!(error.is_instance_of::<pyo3::exceptions::PyRuntimeError>(py));
        });
    }
}
