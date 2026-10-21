# Examples

Start with the [quick start](../README.md#quick-start) or the
[usage guide](../docs/usage.md) for a walkthrough of the API.

Runnable scripts demonstrating `xlspy`. Each script inserts `../src` on
`sys.path`, so you can run them straight from a clone without installing the
package. Scripts that create workbooks write them to the current directory
(the pandas/polars scripts accept an optional output directory argument).

```bash
python examples/simple_usage.py
```

## Writing rows (no extra dependencies)

| Script | What it shows |
|--------|---------------|
| [`simple_usage.py`](simple_usage.py) | Basic `XlsbWriter` and `XlsxWriter` usage: multiple sheets, mixed Python types. |
| [`simple_usage_xlsx.py`](simple_usage_xlsx.py) | Minimal `XlsxWriter` example. |
| [`format_demo.py`](format_demo.py) | All cell number/date/datetime formats from `xlspy.F`. |
| [`update_existing.py`](update_existing.py) | Updating the data region of existing `.xlsx` / `.xlsm` / `.xlsb` workbooks. |

## Reports and large exports (no extra dependencies)

| Script | What it shows |
|--------|---------------|
| [`xlsx_report_features.py`](xlsx_report_features.py) | Formulas, a table, dropdown validation, hyperlinks, an optional logo (`--logo`), print settings, and horizontal bar and scatter charts in one report. |
| [`xlsx_comments_and_templates.py`](xlsx_comments_and_templates.py) | Classic notes, sparse template edits, named ranges, sheet protection and reading formula text; `--out reviewed.xlsb` writes XLSB. |
| [`xlsx_grouping.py`](xlsx_grouping.py) | Two-level row and column groups (`row_groups`, `column_groups`) with region subtotals and quarter columns; `--out grouping.xlsb` writes XLSB. |
| [`xlsx_streaming_export.py`](xlsx_streaming_export.py) | Streams rows from a DB-API cursor (sqlite3 here) into XLSX or XLSB with `useSharedStrings=False` and reports the peak Python allocations. |

## DataFrames (optional pandas / polars)

These require an extra install:

```bash
pip install xlspy[pandas]   # or: pip install xlspy[polars]
```

| Script | What it shows |
|--------|---------------|
| [`dataframe_to_xlsx.py`](dataframe_to_xlsx.py) | One-liner `to_xlsx(df, path)`. |
| [`dataframe_to_xlsb.py`](dataframe_to_xlsb.py) | One-liner `to_xlsb(df, path)`. |
| [`dataframe_writer_methods.py`](dataframe_writer_methods.py) | `add_sheet` + `write_dataframe` + `write_sheet` across several sheets (including a hidden one). |
| [`dataframe_formats_and_index.py`](dataframe_formats_and_index.py) | `include_index=True`, per-column `formats={...}` and `na_rep`. |
| [`dataframe_roundtrip.py`](dataframe_roundtrip.py) | Write a DataFrame, read it back with `ExcelReader`, rebuild it. |
| [`polars_to_xlsx_xlsb.py`](polars_to_xlsx_xlsb.py) | The same API with a polars DataFrame, including durations and nulls. |

## Benchmarks

See [benchmark results and methodology](../docs/benchmarks.md) for the recorded
comparisons and the environment used to measure them.

| Script | What it shows |
|--------|---------------|
| [`benchmark.py`](benchmark.py) | Timing/size comparison against `xlsxwriter`. |
| [`performance_test.py`](performance_test.py) | Larger write/read benchmark against `xlsxwriter` and `openpyxl`. |
| [`write_benchmark.py`](write_benchmark.py) | Reproducible write benchmark: isolated processes, medians, peak memory, file size, JSON/CSV output; xlspy (C and Python), `openpyxl` and `xlsxwriter`. |
