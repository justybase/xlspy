![xlspy — Excel reports, made with Python](https://raw.githubusercontent.com/justybase/xlspy/main/docs/xlspy-banner.svg)

# Excel reports. Python simplicity.

**Turn Python data into Excel workbooks — from everyday reports to large exports.**

Read and write **XLSB** and **XLSX**, bring your pandas or polars DataFrames,
and refresh existing **XLSM** templates while preserving their macros.
Style the result with charts, tables, formatting, and cell notes — all from Python.

[Get started](#quick-start) · [Usage guide](https://github.com/justybase/xlspy/blob/main/docs/usage.md) · [Examples](https://github.com/justybase/xlspy/blob/main/examples/README.md) · [Benchmarks](https://github.com/justybase/xlspy/blob/main/docs/benchmarks.md)

## Built for speed

**2.5 million cells in under a second, as XLSB.**

Times are medians of three runs. WSL2 and Mac were measured with xlspy 0.14.0
on Python 3.14 on 2026-11-05. The Windows reference row is from xlspy 0.12.0
on Python 3.15.2 (2026-11-04), so cross-platform differences are indicative,
not a controlled hardware comparison.

50,000 × 50 cells:

| Windows 11, AMD Ryzen 7 7840HS, Python 3.15.2 | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Write | **0.67 s** | **1.16 s** | 11.71 s | 20.39 s |
| Read | **0.22 s** | **0.98 s** | — | 14.35 s |

| WSL2, Ubuntu, AMD Ryzen 7 7840HS, Python 3.14.4 | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Write | **0.65 s** | **1.00 s** | 8.54 s | 14.21 s |
| Read | **0.16 s** | **0.59 s** | — | 9.14 s |

| macOS 27.2, Apple M4, Python 3.14.10 | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Write | **0.44 s** | **0.71 s** | 5.98 s | 9.71 s |
| Read | **0.07 s** | **0.44 s** | — | 6.72 s |

The native (Rust) module writes XLSB with inline strings 4.2–4.4× faster than
xlspy's pure-Python fallback and reads XLSB 30–56× faster. The WSL2 and Mac
writer comparisons use inline strings for xlspy, `constant_memory` for
xlsxwriter, and `write_only`/`read_only` for openpyxl. Read timings are medians
of three end-to-end runs with a warm page cache. Raw results and methodology:
[benchmarks](https://github.com/justybase/xlspy/blob/main/docs/benchmarks.md#release-candidate-0140-2026-11-05).
[See the full comparisons, hardware, and methodology →](https://github.com/justybase/xlspy/blob/main/docs/benchmarks.md)

## Made for the work behind the workbook

| Your workflow | What xlspy brings |
| :--- | :--- |
| **Export large datasets** | XLSB and XLSX output from a generator or database cursor, with bounded memory via `use_shared_strings=False`, and native (Rust) acceleration for supported read and write paths. A failed export publishes no partial workbook ([failure handling](https://github.com/justybase/xlspy/blob/main/docs/guide/writing.md#failure-handling)). |
| **Deliver polished reports** | Cell styles, number formats, formulas, dropdown validation, hyperlinks, images, print setup, charts (column, bar, line, area, pie, doughnut, scatter, stacked — the chart types listed in the feature table), tables, and conditional formatting. Both XLSX and XLSB support them; see the [feature support table](https://github.com/justybase/xlspy/blob/main/docs/guide/objects.md#feature-support). |
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
API. Explore [formatting, charts, and tables](https://github.com/justybase/xlspy/blob/main/docs/guide/styling.md#cell-styles-colors-fonts-alignment)
when your report needs more detail.

## Bring your data. Keep your workflow.

**Already using pandas or polars?** Export with `to_xlsx(df, "report.xlsx")`
or `to_xlsb(df, "report.xlsb")`.
[See DataFrame examples →](https://github.com/justybase/xlspy/blob/main/docs/guide/dataframes.md#writing-pandas--polars-dataframes-optional)

**Have an Excel template?** Refresh its data with `XlsxUpdater`, `XlsbUpdater`,
or `XlsmUpdater`. XLSM updates preserve the existing VBA project.
[Update a workbook →](https://github.com/justybase/xlspy/blob/main/docs/guide/updating.md#updating-an-existing-workbook)

**Need to read the results?** Use `ExcelReader` for XLSB and XLSX, with the same
API for both formats. It streams values without styles and is the fastest way in.
For random access, formatting and edits, `load_workbook()` loads either format
into an openpyxl-like model; styles are resolved only when you read them, so
there is no flag to set.
[Choose a read mode →](https://github.com/justybase/xlspy/blob/main/docs/guide/reading.md#read-modes-which-one-to-use)

## Explore further

| Resource | What you will find |
| :--- | :--- |
| [Usage guide](https://github.com/justybase/xlspy/blob/main/docs/usage.md) | Detailed examples, API options, and format support limits. |
| [API reference](https://github.com/justybase/xlspy/blob/main/docs/api/index.md) | Every public class and function, generated from the docstrings. |
| [Runnable examples](https://github.com/justybase/xlspy/blob/main/examples/README.md) | Scripts for reports, DataFrames, and workbook updates. |
| [Benchmarks](https://github.com/justybase/xlspy/blob/main/docs/benchmarks.md) | Read and write comparisons, dataset details, and reproduction commands. |
| [Validation](https://github.com/justybase/xlspy/blob/main/docs/validation.md) | Windows and WSL2 test results, compatibility checks, and known limitations. |
| [Test suite](https://github.com/justybase/xlspy/blob/main/tests/README.md) | How to run tests and measure coverage. |
| [Contributing](https://github.com/justybase/xlspy/blob/main/CONTRIBUTING.md) | Development setup, building the Rust module, checks and commit style. |
| [Architecture](https://github.com/justybase/xlspy/blob/main/docs/architecture.md) | Which module does what, the layers, and the native module. |
| [Changelog](https://github.com/justybase/xlspy/blob/main/CHANGELOG.md) | Releases, improvements, and fixes. |

[MIT licensed](https://github.com/justybase/xlspy/blob/main/LICENSE) · [PyPI](https://pypi.org/project/xlspy/) · [Source and issues](https://github.com/justybase/xlspy)
