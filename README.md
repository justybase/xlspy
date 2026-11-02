![xlspy — Excel reports, made with Python](docs/xlspy-banner.svg)

# Excel reports. Python simplicity.

**Turn Python data into Excel workbooks — from everyday reports to large exports.**

Read and write **XLSB** and **XLSX**, bring your pandas or polars DataFrames,
and refresh existing **XLSM** templates while preserving their macros.
Style the result with charts, tables, formatting, and cell notes — all from Python.

[Get started](#quick-start) · [Usage guide](docs/usage.md) · [Examples](examples/README.md) · [Benchmarks](docs/benchmarks.md)

## Built for speed

**2.5 million cells in under a second, as XLSB.**

| 50,000 × 50 cells, Windows | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Write | **0.81 s** | **1.27 s** | 10.17 s | 17.06 s |
| Read | **0.14 s** | **0.73 s** | — | 10.75 s |

The native (Rust) module writes 2.5–3.7× and reads XLSB 46× faster than
xlspy's pure-Python fallback. Medians of three runs on an AMD Ryzen 7 7840HS
with Python 3.14 (2026-11-02); xlspy with inline strings, xlsxwriter with
`constant_memory`, openpyxl with `write_only` and `read_only`.
[See the full comparisons, hardware, and methodology →](docs/benchmarks.md)

## Made for the work behind the workbook

| Your workflow | What xlspy brings |
| :--- | :--- |
| **Export large datasets** | XLSB and XLSX output from a generator or database cursor, with bounded memory via `use_shared_strings=False`, and native (Rust) acceleration for supported read and write paths. A failed export publishes no partial workbook ([failure handling](docs/guide/writing.md#failure-handling)). |
| **Deliver polished reports** | Cell styles, number formats, formulas, dropdown validation, hyperlinks, images, print setup, charts (column, bar, line, area, pie, doughnut, scatter, stacked), tables, and conditional formatting. Both XLSX and XLSB support them; see the [feature support table](docs/guide/objects.md#feature-support). |
| **Work with DataFrames** | Export pandas or polars data directly to either format. |
| **Refresh a familiar template** | Replace data in XLSB, XLSX, or XLSM while keeping workbook structure, styles, and unrelated sheets. |
| **Connect a data pipeline** | Read rows as a generator and stream updates from a database cursor. |

The core library has **no required Python package dependencies**. Add pandas or
polars support when you need it. Microsoft Excel is not required for the core
read and write APIs.

## Quick start

Requires **Python 3.11 or newer** (tested on 3.11–3.15).

```bash
pip install xlspy
```

Wheels include the native (Rust) extension. Installing from the source
distribution builds it, which needs Rust 1.88 or newer
([rustup](https://rustup.rs)). Set `XLSPY_ACCEL=python` to run the pure-Python
code instead.

Create a report with a styled header and formatted values:

```python
from xlspy import F, Style, XlsxWriter

with XlsxWriter("sales.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet(
        [
            ["Product", "Revenue"],
            ["Coffee", (12500, F.THOUSANDS_SEP)],
            ["Tea", (8400, F.THOUSANDS_SEP)],
        ],
        header_style=Style(bold=True, fill="#183C52", font_color="#FFFFFF"),
    )
```

For binary Excel output, use `XlsbWriter` and a `.xlsb` filename with the same
API. Explore [formatting, charts, and tables](docs/guide/styling.md#cell-styles-colors-fonts-alignment)
when your report needs more detail.

## Bring your data. Keep your workflow.

**Already using pandas or polars?** Export with `to_xlsx(df, "report.xlsx")`
or `to_xlsb(df, "report.xlsb")`.
[See DataFrame examples →](docs/guide/dataframes.md#writing-pandas--polars-dataframes-optional)

**Have an Excel template?** Refresh its data with `XlsxUpdater`, `XlsbUpdater`,
or `XlsmUpdater`. XLSM updates preserve the existing VBA project.
[Update a workbook →](docs/guide/updating.md#updating-an-existing-workbook)

**Need to read the results?** Use `ExcelReader` for XLSB and XLSX, with the same
API for both formats. It streams values without styles and is the fastest way in.
For random access, formatting and edits, `load_workbook()` loads either format
into an openpyxl-like model; styles are resolved only when you read them, so
there is no flag to set.
[Choose a read mode →](docs/guide/reading.md#read-modes-which-one-to-use)

## Explore further

| Resource | What you will find |
| :--- | :--- |
| [Usage guide](docs/usage.md) | Detailed examples, API options, and format support limits. |
| [API reference](docs/api/index.md) | Every public class and function, generated from the docstrings. |
| [Runnable examples](examples/README.md) | Scripts for reports, DataFrames, and workbook updates. |
| [Benchmarks](docs/benchmarks.md) | Read and write comparisons, dataset details, and reproduction commands. |
| [Validation](docs/validation.md) | Windows and WSL2 test results, compatibility checks, and known limitations. |
| [Test suite](tests/README.md) | How to run tests and measure coverage. |
| [Contributing](CONTRIBUTING.md) | Development setup, building the Rust module, checks and commit style. |
| [Architecture](docs/architecture.md) | Which module does what, the layers, and the native module. |
| [Changelog](CHANGELOG.md) | Releases, improvements, and fixes. |

[MIT licensed](LICENSE) · [PyPI](https://pypi.org/project/xlspy/) · [Source and issues](https://github.com/justybase/xlspy)
