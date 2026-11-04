//! The formatted-cell pre-scan of the XLSB writer.

use crate::pyseq::FastSeq;
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
