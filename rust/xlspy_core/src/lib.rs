//! Native accelerator for xlspy: a PyO3 port of `c_ext/_c_core.c`.
//!
//! Each Python entry point mirrors the C function of the same name; the
//! selection between the two lives in `xlspy/_accel.py`.

use pyo3::prelude::*;

mod bytes;
mod cells;
mod sst;
mod text;
mod types;
mod widths;
mod xlsb_write;

#[pymodule(gil_used = true)]
fn _rs_core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__doc__", "Rust-optimized functions for xlspy")?;
    m.add_function(wrap_pyfunction!(types::set_error_type, m)?)?;
    m.add_function(wrap_pyfunction!(types::set_style_type, m)?)?;
    m.add_function(wrap_pyfunction!(sst::encode_xlsb_sst, m)?)?;
    m.add_function(wrap_pyfunction!(sst::read_xlsb_shared_strings, m)?)?;
    m.add_function(wrap_pyfunction!(widths::calc_column_widths, m)?)?;
    m.add_function(wrap_pyfunction!(widths::xlsb_row_has_formatted, m)?)?;
    m.add_function(wrap_pyfunction!(xlsb_write::encode_xlsb_row, m)?)?;
    m.add_function(wrap_pyfunction!(xlsb_write::encode_xlsb_row_inline, m)?)?;
    Ok(())
}
