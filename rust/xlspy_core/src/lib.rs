//! Native accelerator for xlspy (`xlspy._rs_core`), a PyO3 port of the
//! former C extension that keeps its behaviour.
//!
//! `xlspy/_accel.py` wraps every entry point; each one returns None where the
//! pure-Python implementation must take over.

use pyo3::prelude::*;

mod bytes;
mod cells;
mod dates;
mod model;
mod numfmt;
mod sst;
mod text;
mod types;
mod widths;
mod xlsb_read;
mod xlsb_write;
mod xlsx_read;
mod xlsx_write;
mod xlsx_sst;
mod xml_text;

#[pymodule(gil_used = true)]
fn _rs_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__doc__", "Rust-optimized functions for xlspy")?;
    m.add_function(wrap_pyfunction!(types::set_error_type, m)?)?;
    m.add_function(wrap_pyfunction!(types::set_style_type, m)?)?;
    m.add_function(wrap_pyfunction!(types::set_formula_type, m)?)?;
    m.add_function(wrap_pyfunction!(types::set_invalid_file_type, m)?)?;
    m.add_function(wrap_pyfunction!(sst::encode_xlsb_sst, m)?)?;
    m.add_function(wrap_pyfunction!(sst::read_xlsb_shared_strings, m)?)?;
    m.add_function(wrap_pyfunction!(xlsx_sst::read_xlsx_shared_strings, m)?)?;
    m.add_function(wrap_pyfunction!(widths::calc_column_widths, m)?)?;
    m.add_function(wrap_pyfunction!(widths::xlsb_row_has_formatted, m)?)?;
    m.add_function(wrap_pyfunction!(xlsb_write::encode_xlsb_row, m)?)?;
    m.add_function(wrap_pyfunction!(xlsb_write::encode_xlsb_row_inline, m)?)?;
    m.add_function(wrap_pyfunction!(xlsx_write::encode_xlsx_row, m)?)?;
    m.add_function(wrap_pyfunction!(xlsb_read::read_xlsb_worksheet, m)?)?;
    m.add_class::<xlsb_read::XlsbReader>()?;
    m.add_function(wrap_pyfunction!(xlsb_read::xlsb_load_cells, m)?)?;
    m.add_function(wrap_pyfunction!(xlsx_read::read_xlsx_worksheet, m)?)?;
    m.add_class::<xlsx_read::XlsxReader>()?;
    Ok(())
}
