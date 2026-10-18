![xlspy — Fast spreadsheet I/O for Python](docs/xlspy-banner.svg)

# Python XLSB/XLSX/XLSM Reader & Writer

A Python library for reading and writing XLSB and XLSX files efficiently, with
existing-workbook updates for macro-enabled XLSM packages.

## Installation

```bash
pip install xlspy
```

## Usage

### Basic Example

```python
from xlspy import XlsbWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.utcnow()]
]

# Initialize writer with a specific compression level
with XlsbWriter("output.xlsb", compressionLevel=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

### XlsxWriter Example

```python
from xlspy import XlsxWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.utcnow()]
]

# Initialize writer with a specific compression level
with XlsxWriter("output.xlsx", compressionLevel=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

### Cell Formatting (Number, Date, DateTime)

`xlspy` supports custom cell formatting for numbers, dates, and datetimes in both XLSB and XLSX output. Pass a `(value, format_string)` tuple to apply a format to a specific cell.

```python
from xlspy import XlsxWriter, F  # or XlsbWriter
import datetime

with XlsxWriter("formatted.xlsx") as writer:
    writer.add_sheet("Formats")
    writer.write_sheet([
        ["Description", "Value"],
        ["Thousands separator", (100000, F.THOUSANDS_SEP)],
        ["Currency PLN",       (100000, F.CURRENCY_PLN)],
        ["Currency EUR",       (100000, F.CURRENCY_EUR)],
        ["Percentage",         (100000, F.PERCENTAGE)],
        ["Scientific",         (100000, F.SCIENTIFIC)],
        ["Two decimals",       (100000, F.TWO_DECIMALS)],
        ["Text",               (100000, F.TEXT)],
        ["Leading zeros",      (100000, F.LEADING_ZEROS)],
        ["Short date",         (datetime.date(2026,6,1), F.DATE_SHORT)],
        ["Long date",          (datetime.date(2026,6,1), F.DATE_LONG)],
        ["ISO date",           (datetime.date(2026,6,1), F.DATE_ISO)],
        ["Month + year",       (datetime.date(2026,6,1), F.DATE_MONTH_YEAR)],
        ["Weekday + date",     (datetime.date(2026,6,1), F.DATE_WEEKDAY)],
        ["Short datetime",     (datetime.datetime(2026,6,1,14,34), F.DATETIME_SHORT)],
        ["Time only",          (datetime.datetime(2026,6,1,14,34), F.TIME_HH_MM)],
        ["12h time",           (datetime.datetime(2026,6,1,14,34), F.TIME_12H)],
        ["ISO datetime",       (datetime.datetime(2026,6,1,14,34), F.DATETIME_ISO)],
    ])
```

#### Available Format Constants (`xlspy.F`)

| Number | Date | DateTime |
|--------|------|----------|
| `F.THOUSANDS_SEP` — `#,##0` | `F.DATE_SHORT` — `dd.mm.yyyy` | `F.DATETIME_SHORT` — `dd.mm.yyyy hh:mm` |
| `F.CURRENCY_PLN` — `#,##0.00 "zł"` | `F.DATE_LONG` — `d mmmm yyyy` | `F.DATETIME_LONG` — `d mmmm yyyy hh:mm:ss` |
| `F.CURRENCY_EUR` — `#,##0.00 €` | `F.DATE_DAY_MONTH_YEAR` — `dd-mm-yyyy` | `F.TIME_HH_MM` — `hh:mm` |
| `F.PERCENTAGE` — `0%` | `F.DATE_ISO` — `yyyy-mm-dd` | `F.TIME_HH_MM_SS` — `hh:mm:ss` |
| `F.SCIENTIFIC` — `0.00E+00` | `F.DATE_MONTH_YEAR` — `mmmm yyyy` | `F.TIME_12H` — `h:mm AM/PM` |
| `F.TWO_DECIMALS` — `#,##0.00` | `F.DATE_WEEKDAY` — `dddd, d mmmm yyyy` | `F.DATETIME_24H` — `dd.mm.yyyy hh:mm:ss` |
| `F.TEXT` — `@` | `F.DATE_DAY_MONTH` — `d mmmm` | `F.DATETIME_ISO` — `yyyy-mm-dd"T"hh:mm:ss` |
| `F.LEADING_ZEROS` — `000000000` | `F.DATE_YEAR_ONLY` — `yyyy` | `F.TIME_MS` — `hh:mm:ss.000` |

You can also use custom format strings directly:

```python
writer.write_sheet([
    ["Custom", (1234.56, '#,##0.00 "USD"')],
    ["Date",   (datetime.date(2026,6,1), 'dd.mm.yyyy')],
])
```

The formatting works transparently on both `XlsxWriter` and `XlsbWriter`.

### Cell Styles (colors, fonts, alignment)

`Style` can be used on the header or on arbitrary A1 cell ranges. Styles are
deduplicated in the workbook, so cells with the same style share one Excel
style record.

```python
from xlspy import F, Style, XlsxWriter  # XlsbWriter supports the flat fields below

header = Style(fill="#008000", font_color="#FFFFFF", bold=True)
warning = Style(fill="FFF2CC", italic=True, horizontal="center", wrap=True)

with XlsxWriter("report.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_sheet(
        [
            ["Product", "Quantity", "Comment"],
            ["Apples", (1200, F.THOUSANDS_SEP), "Restock soon"],
        ],
        header_style=header,
        styles={"B2:B2": warning, "C2:C2": warning},
    )
```

Flat `Style` properties: `fill`, `font_color`, `bold`, `italic`, `underline`,
`number_format`, `horizontal`, `vertical`, and `wrap`. Colors are six-digit RGB
strings with or without `#`. Use `header_style` to style all cells in the first
row (including numeric cells); use `styles` to provide an ordered mapping of A1
cell/range references to styles. When ranges overlap, later ranges override
only the properties they specify. A cell's existing `(value, number_format)`
tuple remains supported and its number format takes precedence over a range
style's `number_format`.

`write_dataframe()` accepts the same `header_style` and `styles` arguments.
Unstyled rows continue to use the XLSB C acceleration path. The C encoder also
accepts compact row style IDs, avoiding the old tuple-based fallback for
styled rows.

#### Structured styles (openpyxl-like)

For full Excel styling, `Style` also accepts openpyxl-style objects. Each
sub-property left as `None` falls back to the corresponding flat shorthand,
and the structured object wins on conflict:

```python
from xlspy import (
    Alignment, Border, Color, Font, GradientFill, PatternFill, Side, Style,
)

style = Style(
    font=Font(name="Arial", size=12, bold=True, color=Color(theme=1),
              scheme="major", shadow=True),
    fill=GradientFill("linear", degree=90, stop=["#4472C4", "#FFFFFF"]),
    border=Border(left=Side("thin", "000000"), bottom=Side("double")),
    alignment=Alignment(horizontal="center", wrap_text=True, reading_order=2),
    number_format="#,##0.00",
)
```

- `fill` accepts a color string (solid fill), a `PatternFill` (`fill_type`,
  `start_color`, `end_color`) or a `GradientFill` (`linear`/`path` with `Stop`
  objects or plain colors spread evenly between position 0 and 1). When both
  `fill` and `pattern` are given, `pattern` wins.
- `Color` takes exactly one of `rgb`, `indexed`, `theme` (optionally with
  `tint`) or `auto=True`.
- `Font` covers `name`, `size`, `bold`, `italic`, `strike`, `underline`,
  `vertAlign`, `color`, `charset`, `family`, `scheme`, `outline`, `shadow`,
  `condense`, and `extend`.
- `Border` covers all nine sides (`left`, `right`, `top`, `bottom`,
  `diagonal`, `vertical`, `horizontal`, `start`, `end`) plus
  `diagonal_direction`/`diagonal_up`/`diagonal_down` and `outline`.
- `Alignment` covers `horizontal`, `vertical`, `wrap_text`, `shrink_to_fit`,
  `indent`, `relative_indent`, `text_rotation`, `justify_last_line`, and
  `reading_order`.
- `Protection(locked=..., hidden=...)` for cell locking.

#### Conditional formatting

`write_sheet(..., conditional_formats=[...])` accepts `CellIsRule`,
`FormulaRule`, `ColorScaleRule`, `DataBarRule`, and `IconSetRule`. Rules are
grouped by range; priorities are assigned automatically or taken from the
rule's `priority`. A `DifferentialStyle` (Excel `dxf`) can style the matched
cells with a font, number format, fill, alignment, border, and protection:

XLSB supports `ColorScaleRule`, `DataBarRule`, and `IconSetRule`; XLSB
`CellIsRule` and `FormulaRule` currently raise `NotImplementedError`.

```python
from xlspy import (
    CellIsRule, DifferentialStyle, Font, IconSetRule, PatternFill,
)

formats = [
    CellIsRule("B2:B5", "lessThan", ["5"], dxf=DifferentialStyle(
        font=Font(bold=True, color="9C0006"),
        number_format="0.00",
        fill=PatternFill("solid", start_color="FFC7CE"),
    )),
    IconSetRule("C2:C5", "5Arrows", [0, 20, 40, 60, 80], show_value=True),
]
```

#### Tables

`write_sheet(..., tables=[ExcelTable("Sales", "A1:C5", columns=["Qty", "Price", "Total"])])`
adds an Excel ListObject table to XLSX and XLSB workbooks.

#### Charts

`write_sheet(..., charts=[...])` adds charts (XLSX and XLSB):

```python
from xlspy import Chart, ChartSeries

writer.write_sheet(
    [["Q", "A"], ["Q1", 10], ["Q2", 30]],
    charts=[
        Chart("bar", [ChartSeries("B2:B3", title_cell="B1")],
              title="Sales", anchor="D2"),
        Chart("line", [ChartSeries("B2:B3", title="Series 1")],
              title="Trend", anchor="D18"),
    ],
)
```

`Chart(chart_type="bar"|"line"|"pie", series=[...], title=None, anchor="E2",
width=15.0, height=7.5, style=None)` anchors the chart at a cell; dimensions
are in centimeters (openpyxl defaults: 15 × 7.5). Each `ChartSeries` takes an
A1 `values` range on the same sheet, an optional legend label (`title` literal
or `title_cell` reference) and, except for pie charts, optional `categories`.
Both writers emit the same DrawingML parts (`drawings/drawingN.xml`,
`charts/chartN.xml`); the XLSX writer links them with a `<drawing>` sheet
element, the XLSB writer with a `BrtDrawing` record. XLSB chart parts also
mirror real Excel files: cached values inside every reference, text
categories via `strRef`, and an explicit `varyColors` flag.

**XLSB parity:** `XlsbWriter` supports the same `Style` object model
(`Font`, `PatternFill`, `GradientFill`, `Border`, `Alignment`, `Protection`)
as well as `merges`, `column_widths`, `hidden_columns`, `row_heights`,
`hidden_rows`, `freeze_panes`, and `auto_filter`. A few properties have no
BIFF12 representation and raise `NotImplementedError` instead of being
silently ignored: `relative_indent`, the `vertical`/`horizontal`/`start`/`end`
  border sides, and `outline=False`.

### Writing pandas / polars DataFrames (optional)

`xlspy` can write a DataFrame directly to `.xlsx` or `.xlsb`. **pandas and
polars are optional** — `xlspy` has no required dependencies and imports them
lazily only when you actually pass a DataFrame:

```bash
pip install xlspy[pandas]   # pandas support
pip install xlspy[polars]   # polars support
```

The quickest way is the one-liner `to_xlsx` / `to_xlsb`:

```python
import pandas as pd
from xlspy import to_xlsx, to_xlsb

df = pd.DataFrame({
    "name": ["Alice", "Bob"],
    "age": [25, 30],
    "joined": pd.to_datetime(["2024-01-15", "2023-06-01"]),
})

to_xlsx(df, "people.xlsx")
to_xlsb(df, "people.xlsb")
```

The same works for polars, with an identical API:

```python
import polars as pl
from xlspy import to_xlsb

df = pl.DataFrame({"name": ["Alice", "Bob"], "age": [25, 30]})
to_xlsb(df, "people.xlsb")
```

For full control, use the writer method, which behaves exactly like
`write_sheet` and writes into the last sheet added with `add_sheet`:

```python
import pandas as pd
from xlspy import XlsxWriter, F

df = pd.DataFrame({
    "joined": pd.to_datetime(["2024-01-15", "2023-06-01"]),
    "age": [25, 30],
})

with XlsxWriter("report.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_dataframe(
        df,
        include_index=True,
        na_rep="n/a",
        formats={"joined": F.DATE_ISO, "age": F.THOUSANDS_SEP},
    )
    writer.add_sheet("Notes")
    writer.write_sheet([["Generated by xlspy"]])
```

#### Options

| Parameter | Applies to | Meaning |
|-----------|-----------|---------|
| `include_index` | pandas only | Prepend the index as leading column(s); all MultiIndex levels are written. polars has no index and raises `ValueError`. |
| `na_rep` | both | Text written for missing cells (`None` / `NaN` / `NaT` / `NA`). Default is `None`, i.e. an empty cell. |
| `formats` | both | `{column_name: format_string}` mapping, using the constants from `xlspy.F`. |

Values are converted to the types the writers understand:

| DataFrame value | Written as |
|-----------------|-----------|
| `int` / `float` / `str` / `bool` / `Decimal` | native cell |
| `pandas.Timestamp`, `datetime.datetime` | date/time cell |
| `pandas.Timedelta`, `datetime.timedelta` | day fraction (Excel duration; apply `F.TIME_*` for display) |
| `datetime.time` (incl. `polars.Time`) | day fraction (apply `F.TIME_*` for display) |
| `numpy` scalars (`int64`, `float32`, `bool_`, ...) | native Python value |
| `None`, `NaN`, `NaT`, `NA` | empty cell (or `na_rep`) |
| MultiIndex columns | flattened by joining the levels with a space |

See [`examples/dataframe_*.py`](examples/) and
[`examples/polars_to_xlsx_xlsb.py`](examples/polars_to_xlsx_xlsb.py) for full,
runnable scripts, indexed in [`examples/README.md`](examples/README.md).

### Updating an Existing Workbook

`XlsxUpdater` (also for macro-enabled `.xlsm` files) and `XlsbUpdater` replace
the data region of an existing workbook
while keeping the workbook structure, styles, drawings and unrelated sheets.
They are useful when a workbook contains pivot tables: the worksheet source
range and pivot-cache refresh metadata are updated together.

```python
from xlspy import XlsxUpdater, XlsbUpdater

rows = [
    ["Alice", 42],
    ["Bob", 37],
]

updater = XlsxUpdater("template.xlsx")
print(updater.get_sheet_names())
updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
updater.save("result.xlsx")       # use save() to replace the input atomically

# The XLSB API is identical:
binary_updater = XlsbUpdater("template.xlsb")
binary_updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
binary_updater.save("result.xlsb")

# XLSM uses the same XML updater.  The VBA project is preserved byte-for-byte.
macro_updater = XlsxUpdater("template.xlsm")
macro_updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
macro_updater.save("result.xlsm")
```

`XlsmUpdater` is an explicit alias of `XlsxUpdater` for `.xlsm` inputs.  The
updater does not parse, execute or rewrite VBA; `xl/vbaProject.bin` and other
unrelated macro parts are copied as opaque ZIP members.

For a database cursor or another one-pass source, use the streaming variant.
Rows are written through bounded-memory temporary files, so the complete
result set is not kept in Python memory. Multiple sheets may be streamed on
the same updater before `save()`:

```python
def rows_from_cursor(cursor):
    while True:
        row = cursor.fetchone()
        if row is None:
            break
        yield list(row)

updater = XlsxUpdater("template.xlsx")  # also accepts .xlsm; use XlsbUpdater for .xlsb
updater.replace_sheet_data_stream(
    "Data",
    rows_from_cursor(cursor),
    headers=["Name", "Amount"],
)
updater.save("result.xlsx")              # or result.xlsb
```

The streaming API requires `save()` for its memory-bounded write path.
`to_bytes()` necessarily creates the complete archive in memory because it
returns the complete file as `bytes`. The original `replace_sheet_data()` API
is unchanged and remains available for in-memory inputs.

Rows may be any iterable, including a generator. Only trailing rows consisting
entirely of `None` are removed; blank rows in the middle are preserved. By
default, the updater inherits the dominant style of each existing column. Use
`style_fallback="general"` to write new cells with the General style. Existing
writer-style `(value, format_string)` tuples are accepted for compatibility,
but updating a workbook does not create new styles.

When a pivot cache is present, keep the source column schema (column order and
count) unchanged; changing the schema requires rebuilding the pivot cache and
is outside the updater's scope.

On Windows, the generated workbook can be checked with the installed Excel
COM server (the script opens files read-only using Excel's normal loader):

```powershell
powershell -ExecutionPolicy Bypass -File tools\validate_excel_com.ps1 result.xlsx result.xlsb
```

### Reading XLSB and XLSX Files

Reading files is done via the `ExcelReader` class, which automatically detects the format.

```python
from xlspy import ExcelReader

with ExcelReader("input.xlsx") as reader:  # or .xlsb
    names = reader.get_sheet_names()
    print(f"Sheets: {names}")

    for sheet_name in names:
        rows = reader.read_all(sheet_name)
        for row in rows:
            print(row)

# Generator usage (memory efficient for large files):
with ExcelReader("large_file.xlsb") as reader:
    for row in reader.get_rows("Sheet1"):
        print(row)
```

### Streaming from a Database (Netezza)

This example shows how to stream data directly from a database query into an XLSB file using `nzpy-extended`. This is highly memory-efficient as it doesn't load the entire dataset into memory.

First, ensure you have `nzpy-extended` installed:
```bash
pip install nzpy-extended
```

Then, you can use a generator function to feed data to `XlsbWriter`.

```python
import os
from typing import Generator
from xlspy import XlsbWriter

# --- Configuration ---
NZ_CONFIG = {
    "host": os.environ.get("NZ_DEV_HOST", "your_host"),
    "port": int(os.environ.get("NZ_DEV_PORT", "5480")),
    "database": os.environ.get("NZ_DEV_DB", "your_db"),
    "user": os.environ.get("NZ_DEV_USER", "your_user"),
    "password": os.environ.get("NZ_DEV_PASSWORD", "your_password"),
}
QUERY = "SELECT * FROM YourTable"
OUTPUT_FILENAME = "db_output.xlsb"


def row_generator(cursor) -> Generator[list, None, None]:
    """Yields column headers first, then each data row."""
    headers = [column[0] for column in cursor.description]
    yield headers
    while row := cursor.fetchone():
        yield list(row)


# --- Main Execution ---
try:
    import nzpy_extended.sync as nzpy

    with nzpy.connect(**NZ_CONFIG) as conn:
        cursor = conn.cursor()
        cursor.execute(QUERY)

        with XlsbWriter(OUTPUT_FILENAME) as writer:
            writer.add_sheet("Database Export")
            writer.write_sheet(row_generator(cursor))
            writer.add_sheet("SQL Query", hidden=True)
            writer.write_sheet([["SQL"], [QUERY]])

    print(f"Successfully created '{OUTPUT_FILENAME}'")

except Exception as e:
    print(f"An unexpected error occurred: {e}")
```

## Performance

`xlspy` is designed for high performance. Since version 0.1.0, the library includes a **C extension** (`_c_core`) that accelerates XLSB read and write as well as the plain XLSX write path. The C extension is **enabled by default** (compiled automatically on install). Set `XLSPY_DISABLE_C_EXT=1` to force the pure Python fallback.

All benchmarks: **50000 × 50** dataset (2.5M cells); times below are the **median of three runs**. Tests performed on **Windows** (AMD64, AMD Ryzen 7 7840HS, Python 3.14.6; xlspy 0.7.0, openpyxl 3.1.5, xlsxwriter 3.2.9).

### Write

| Library | Format | Time | Size |
|---------|--------|------|------|
| **xlspy (C_EXT)** | XLSB | **1.46 s** | 7.25 MB |
| xlspy (Python) | XLSB | 3.21 s | 7.25 MB |
| xlspy | XLSX | 4.56 s | 6.43 MB |
| [xlsxwriter](https://pypi.org/project/xlsxwriter/) | XLSX | 9.85 s | 11.57 MB |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 17.77 s | 11.71 MB |

### Read

| Library | Format | Time | Notes |
|---------|--------|------|-------|
| **xlspy (C_EXT)** | XLSB | **0.60 s** | default, compiled C |
| xlspy | XLSX | 4.96 s | uses expat XML parser (C) |
| xlspy (Python) | XLSB | 6.33 s | pure Python fallback |
| [openpyxl](https://pypi.org/project/openpyxl/) | XLSX | 11.06 s | read-only mode |


### Analysis

The **10.6× read speedup** comes from two factors:
- **~60–70%** — native C compilation, no interpreter overhead per record
- **~30–40%** — algorithm simplification: flat array indexed by `col − first_col` instead of `Dict[int, Any]`, no `isinstance` per cell, no `BiffReader.read_worksheet()` method call per record

Run the benchmarks yourself with `examples/performance_test.py`.

### Correctness validation

Validation on the same Windows/Python 3.14.6 environment:

- C-extension test run: **268 passed, 2 skipped**; the three direct Excel COM
  tests were run separately and passed. The skips were the optional real-data
  dependency and the opt-in COM updater test.
- Pure-Python fallback test run (`XLSPY_DISABLE_C_EXT=1`): **242 passed,
  28 skipped**; C-only tests are expected to skip in this mode.
- Cross-validation against **openpyxl** and **xlsxwriter**: **11 passed**.
- Excel COM validation: direct XLSB/XLSX open/readback tests **3 passed**;
  PowerShell validation of updated XLSB/XLSX files **1 passed**.
- Additional differential gates compared C and Python writer payloads and
  reader grids, including seeded mixed-value cases, a 50000 × 50 workbook,
  malformed BIFF records, date-system edge cases, and shared-string truncation
  and corruption cases.

The legacy in-process COM tests pass but emit `0x80010108` shutdown diagnostics
from Excel/pywin32 on this host; the separate PowerShell COM validation passed
without that diagnostic.

### Local test run (WSL2)

Run on **2026-10-18** in WSL2 (Linux x86_64, Python 3.14.4, pytest 9.1.1):

- With the C extension built: **257 passed, 13 skipped**.
- Pure-Python fallback (without the C extension): **239 passed, 31 skipped**.
- The skips were for optional `nzpy_extended`/`psutil` dependencies and
  Windows Excel COM tests, which cannot run in this WSL2 environment.

These results are for WSL2/Linux; the CI workflow separately includes native
Windows runners in its operating-system matrix.

## Repository

<https://github.com/justybase/xlspy/>
