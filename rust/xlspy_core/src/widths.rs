//! The formatted-cell pre-scan of the XLSB writer.

use crate::pyseq::FastSeq;
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::PyTuple;

use crate::types::{overlay_kind, style_type};

/// Return True if a row contains any (value, format) or (value, Style) cell,
/// or None when the row cannot be scanned here (the Python check then raises
/// the same errors).
#[pyfunction]
pub fn xlsb_row_has_formatted<'py>(
    py: Python<'py>,
    row: &Bound<'py, PyAny>,
) -> PyResult<Option<bool>> {
    // A TypeError means "not a sequence": the Python check handles it.  Any
    // other error from a custom iterable must propagate.
    let seq = match FastSeq::new(row, c"row must be a sequence") {
        Ok(seq) => seq,
        Err(error) if error.is_instance_of::<PyTypeError>(py) => return Ok(None),
        Err(error) => return Err(error),
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

#[cfg(test)]
mod tests {
    use super::*;
    use pyo3::types::{PyList, PyTuple};

    #[test]
    fn formatted_cells_are_detected() {
        Python::initialize();
        Python::attach(|py| {
            let plain = PyList::new(py, [1i64, 2]).unwrap().into_any();
            assert_eq!(xlsb_row_has_formatted(py, &plain).unwrap(), Some(false));

            let formatted = PyList::empty(py);
            formatted.append(1i64).unwrap();
            let pair = PyTuple::new(
                py,
                [
                    2i64.into_pyobject(py).unwrap().into_any(),
                    "0.00".into_pyobject(py).unwrap().into_any(),
                ],
            )
            .unwrap();
            formatted.append(pair).unwrap();
            assert_eq!(
                xlsb_row_has_formatted(py, formatted.as_any()).unwrap(),
                Some(true)
            );

            let wrong_tuple = PyList::empty(py);
            wrong_tuple
                .append(PyTuple::new(py, [1i64, 2i64, 3i64]).unwrap())
                .unwrap();
            assert_eq!(
                xlsb_row_has_formatted(py, wrong_tuple.as_any()).unwrap(),
                Some(false)
            );

            let not_a_row = 5i64.into_pyobject(py).unwrap().into_any();
            assert_eq!(xlsb_row_has_formatted(py, &not_a_row).unwrap(), None);
        });
    }
}
