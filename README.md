![xlspy — Excel reports, made with Python](docs/xlspy-banner.svg)

# Excel reports. Python simplicity.

**Turn Python data into Excel workbooks — from everyday reports to large exports.**

Read and write **XLSB** and **XLSX**, bring your pandas or polars DataFrames,
and refresh existing **XLSM** templates while preserving their macros.
Style the result with charts, tables, and formatting — all from Python.

[Get started](#quick-start) · [Usage guide](docs/usage.md) · [Examples](examples/README.md) · [Benchmarks](docs/benchmarks.md)

## Built for speed

**2.5 million cells. Ready in seconds.**

| XLSB with C acceleration | Windows | WSL2 |
| :--- | ---: | ---: |
| Write | **1.46 s** | **1.17 s** |
| Read | **0.60 s** | **0.43 s** |

C acceleration delivered **over 10× faster XLSB reads** than xlspy's
pure-Python fallback in both environments. Figures are medians of three runs
on a 50,000 × 50 dataset (Windows: Python 3.14.6; WSL2: Python 3.14.4).
[See the full comparisons, hardware, and methodology →](docs/benchmarks.md)

## Made for the work behind the workbook

| Your workflow | What xlspy brings |
| :--- | :--- |
| **Export large datasets** | XLSB and XLSX output, with C acceleration for supported read and write paths. |
| **Deliver polished reports** | Cell styles, number formats, charts, tables, and conditional formatting. |
| **Work with DataFrames** | Export pandas or polars data directly to either format. |
| **Refresh a familiar template** | Replace data in XLSB, XLSX, or XLSM while keeping workbook structure, styles, and unrelated sheets. |
| **Connect a data pipeline** | Read rows as a generator and stream updates from a database cursor. |

The core library has **no required Python package dependencies**. Add pandas or
polars support when you need it. Microsoft Excel is not required for the core
read and write APIs.

## Quick start

Requires **Python 3.12 or newer**.

```bash
pip install xlspy
```

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
API. Explore [formatting, charts, and tables](docs/usage.md#cell-styles-colors-fonts-alignment)
when your report needs more detail.

## Bring your data. Keep your workflow.

**Already using pandas or polars?** Export with `to_xlsx(df, "report.xlsx")`
or `to_xlsb(df, "report.xlsb")`.
[See DataFrame examples →](docs/usage.md#writing-pandas--polars-dataframes-optional)

**Have an Excel template?** Refresh its data with `XlsxUpdater`, `XlsbUpdater`,
or `XlsmUpdater`. XLSM updates preserve the existing VBA project.
[Update a workbook →](docs/usage.md#updating-an-existing-workbook)

**Need to read the results?** Use `ExcelReader` for XLSB and XLSX, with the same
API for both formats.
[Read a workbook →](docs/usage.md#reading-xlsb-and-xlsx-files)

## Explore further

| Resource | What you will find |
| :--- | :--- |
| [Usage guide](docs/usage.md) | Detailed examples, API options, and format support limits. |
| [Runnable examples](examples/README.md) | Scripts for reports, DataFrames, and workbook updates. |
| [Benchmarks](docs/benchmarks.md) | Read and write comparisons, dataset details, and reproduction commands. |
| [Validation](docs/validation.md) | Windows and WSL2 test results, compatibility checks, and known limitations. |
| [Test suite](tests/README.md) | How to run tests and measure coverage. |
| [Changelog](CHANGELOG.md) | Releases, improvements, and fixes. |

[MIT licensed](LICENSE) · [PyPI](https://pypi.org/project/xlspy/) · [Source and issues](https://github.com/justybase/xlspy)
