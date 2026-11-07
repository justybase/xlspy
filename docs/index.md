---
description: Read and write XLSB and XLSX files, export pandas and polars DataFrames, and refresh XLSM templates from Python.
---

<div class="xlspy-hero" markdown>

# Excel reports. Python simplicity.

<p class="xlspy-tagline">
Turn Python data into Excel workbooks — from everyday reports to large exports.
Read and write <strong>XLSB</strong> and <strong>XLSX</strong>, export DataFrames,
and refresh existing <strong>XLSM</strong> templates while preserving their macros.
</p>

<div class="xlspy-badges">
  <a href="https://pypi.org/project/xlspy/"><img src="https://img.shields.io/pypi/v/xlspy?color=2f9e8f&label=PyPI" alt="PyPI"></a>
  <a href="https://pypi.org/project/xlspy/"><img src="https://img.shields.io/pypi/pyversions/xlspy?color=2f9e8f" alt="Python versions"></a>
  <a href="https://github.com/justybase/xlspy/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-2f9e8f" alt="MIT license"></a>
  <a href="https://github.com/justybase/xlspy/actions/workflows/ci.yml"><img src="https://github.com/justybase/xlspy/actions/workflows/ci.yml/badge.svg?branch=main" alt="CI"></a>
</div>

<div class="xlspy-actions" markdown>
[Start the tutorial](learn/index.md){ .md-button .md-button--primary }
[API reference](api/index.md){ .md-button }
</div>

<p class="xlspy-install"><code>pip install xlspy</code></p>

</div>

## Why xlspy

<div class="xlspy-cards" markdown>

<div class="xlspy-card" markdown>
### Fast by default
2.5 million cells in under a second as XLSB. The native Rust extension accelerates reads, writes and `load_workbook()`.
</div>

<div class="xlspy-card" markdown>
### XLSB and XLSX
One API for both formats: `XlsxWriter` and `XlsbWriter` take the same rows, styles, charts and tables.
</div>

<div class="xlspy-card" markdown>
### DataFrames included
Export pandas or polars directly with `to_xlsx()` / `to_xlsb()`, or stream rows from a database cursor.
</div>

<div class="xlspy-card" markdown>
### Templates preserved
Refresh data in XLSM templates without touching the VBA project, drawings or unrelated sheets.
</div>

<div class="xlspy-card" markdown>
### No required dependencies
The core library has no required Python packages. Pandas and polars are optional extras.
</div>

<div class="xlspy-card" markdown>
### Excel not required
Create and read workbooks on Linux, Windows or macOS. Microsoft Excel is not needed for the core APIs.
</div>

</div>

## Your first report

```python
from xlspy import F, Style, XlsxWriter

rows = [
    ["Product", "Units", "Revenue"],
    ["Coffee", 120, (12500, F.THOUSANDS_SEP)],
    ["Tea", 90, (8400, F.THOUSANDS_SEP)],
    ["Juice", 45, (5100, F.THOUSANDS_SEP)],
]

with XlsxWriter("sales.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet(
        rows,
        header_style=Style(bold=True, fill="#183C52", font_color="#FFFFFF"),
        freeze_panes="A2",
    )
```

For binary output, change `XlsxWriter` to `XlsbWriter` and the filename to `.xlsb`.
The rest of the code stays the same.

[Build this report step by step →](learn/first-workbook.md){ .md-button }

## Performance at a glance

50,000 × 50 cells (2.5 million cells), Windows 11, AMD Ryzen 7 7840HS:

| | xlspy XLSB | xlspy XLSX | xlsxwriter | openpyxl |
| :--- | ---: | ---: | ---: | ---: |
| Write | **0.67 s** | **1.16 s** | 11.71 s | 20.39 s |
| Read | **0.22 s** | **0.98 s** | — | 14.35 s |

[Full benchmarks and methodology →](performance.md)

## Start here

<div class="xlspy-cards" markdown>

<div class="xlspy-card" markdown>
### [Learn xlspy](learn/index.md)
A ten-step tutorial: install, write your first workbook, style it, add formulas and charts, read it back.
</div>

<div class="xlspy-card" markdown>
### [How-to guides](guide/index.md)
Task-focused pages: choose a format, stream large exports, update templates, add pivots, comments and protection.
</div>

<div class="xlspy-card" markdown>
### [API reference](api/index.md)
Every public class and function, generated from the docstrings, with signatures and examples.
</div>

<div class="xlspy-card" markdown>
### [Explanation](explanation/index.md)
Architecture, performance, validation and the openpyxl parity matrix — the reasoning behind the library.
</div>

</div>

## Works with your stack

| Workflow | Entry point |
| :--- | :--- |
| pandas or polars | `to_xlsx(df, "report.xlsx")`, `to_xlsb(df, "report.xlsb")` |
| Existing Excel template | `XlsxUpdater`, `XlsmUpdater`, `XlsbUpdater` |
| Database cursor | `writer.write_sheet(cursor_rows(cursor))` |
| Random access and edits | `load_workbook(path)` |
| Values only, fastest | `ExcelReader(path).get_rows(sheet)` |

[Examples](https://github.com/justybase/xlspy/blob/main/examples/README.md) · [Changelog](https://github.com/justybase/xlspy/blob/main/CHANGELOG.md) · [GitHub](https://github.com/justybase/xlspy)
