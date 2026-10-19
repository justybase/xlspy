# Using xlspy

[Home](../README.md) · [Runnable examples](../examples/README.md) · [Benchmarks](benchmarks.md)

Read, write, and update Excel workbooks from Python. This guide covers the
APIs and options behind the [quick start](../README.md#quick-start).

- [Write XLSB](#basic-example) or [XLSX](#xlsxwriter-example)
- [Format cells](#cell-formatting-number-date-datetime) and [style reports](#cell-styles-colors-fonts-alignment)
- [Export pandas and polars DataFrames](#writing-pandas--polars-dataframes-optional)
- [Update existing workbooks](#updating-an-existing-workbook)
- [Read XLSB and XLSX files](#reading-xlsb-and-xlsx-files)
- [Stream from a database](#streaming-from-a-database-netezza)

Run commands and examples from the repository root unless stated otherwise.

## Basic Example

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

## XlsxWriter Example

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

## Cell Formatting (Number, Date, DateTime)

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

### Available Format Constants (`xlspy.F`)

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

## Cell Styles (colors, fonts, alignment)

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

### Structured styles (openpyxl-like)

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

### Conditional formatting

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

### Tables

`write_sheet(..., tables=[ExcelTable("Sales", "A1:C5", columns=["Qty", "Price", "Total"])])`
adds an Excel ListObject table to XLSX and XLSB workbooks.

Excel opens a table only when every cell of its header row holds text equal to
that column's name (case aside) and the names are unique. Without `columns`, the
names are taken from the header row of the data; with `columns`, they must match
it. A header cell that is empty or not text, a name that differs from its cell,
or a repeated name raises `ValueError`, and the workbook is not written. Table
names that read as cell references (`T1`, `AB12`, `R1C1`, `C`) are refused when
the `ExcelTable` is created. A table with `header_row=False` has no header cells
and gets the names `Column1`, `Column2`, and so on.

### Charts

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
`hidden_rows`, `row_groups`, `column_groups`, `freeze_panes`, and
`auto_filter`. A few properties have no BIFF12 representation and raise
`NotImplementedError` instead of being silently ignored: `relative_indent`,
the `vertical`/`horizontal`/`start`/`end` border sides, and `outline=False`.

## Row and column groups (outline)

`row_groups` and `column_groups` give Excel's grouping: the `+`/`−` buttons
beside the rows or above the columns, and the level buttons (`1`, `2`, `3`) in
the corner. Both writers support them, as do `write_dataframe`, `to_xlsx`, and
`to_xlsb`.

```python
from xlspy import Group, XlsxWriter

with XlsxWriter("grouped.xlsx") as writer:
    writer.add_sheet("Sales")
    writer.write_sheet(
        rows,
        # Rows 3-8 form one group; rows 3-4 and 6-7 are nested inside it.
        row_groups=[Group(2, 7), Group(2, 3), Group(5, 6, collapsed=True)],
        # Columns D-K, with D-F and H-J nested inside.
        column_groups=[Group("D", "K"), Group("D", "F"), Group("H", "J")],
    )
```

- `Group(first, last, collapsed=False)` covers rows or columns `first` to
  `last`, both included. Rows are 0-based ints. Columns are letters or 0-based
  ints. A `(first, last)` or `(first, last, collapsed)` tuple works as well.
- The level comes from nesting: a group inside another is one level deeper.
  Groups may nest or sit side by side, but a partial overlap raises
  `ValueError`, as does a duplicate group or more than seven levels (Excel's
  limit).
- The summary row is the row just below a group, and the summary column is the
  column just to the right of it (Excel's default). Write your totals there.
- `collapsed=True` hides the group's rows or columns and marks its summary row
  or column, so Excel shows the `+` button. The summary row must be among the
  rows you write. Hidden rows and columns from `hidden_rows` and
  `hidden_columns` still apply.
- Grouped rows are written by the Python row encoder rather than the C one,
  like rows with a custom height. Sheets without groups are unchanged and keep
  the fast path.
- Updaters (`XlsxUpdater`, `XlsbUpdater`) do not create or change groups.

[`xlsx_grouping.py`](../examples/xlsx_grouping.py) builds a two-level report
with region subtotals and quarter columns (pass `--out grouping.xlsb` for XLSB).

## Formulas, validation, hyperlinks, images, and printing (XLSX)

These features are written by `XlsxWriter`. `XlsbWriter` refuses them with a
`TypeError` (or a `ValueError` for chart options), so an unsupported option is
never dropped silently. See [feature support](#feature-support) for the full
matrix.

### Formulas

Wrap formula text in `Formula`. A plain string is always text, so a string
that starts with `=` is never turned into a formula, and exporting untrusted
data cannot inject one.

```python
from xlspy import Formula, XlsxWriter

with XlsxWriter("report.xlsx") as writer:
    writer.add_sheet("Report")
    writer.write_sheet([
        ["Revenue", "Cost", "Profit"],
        [1000, 700, Formula("A2-B2")],
        [2000, 900, Formula("A3-B3", value=1100)],   # cached result for readers that do not recalculate
        ["Total", None, Formula("SUM(C2:C3)")],
        [Formula("XLOOKUP(\"Pear\",A2:A3,C2:C3)"), None, None],
    ])
```

- `Formula(text, value=..., array_ref=...)`: the text may start with `=`; `value`
  (number or string) is written as the cached result; `array_ref="C2:C5"` writes a
  legacy array (CSE) formula over that range.
- Excel recalculates formulas on open. Cells without a cached value read as
  `None` in data-only readers until Excel saves the file.
- Functions added after Excel 2007 (`XLOOKUP`, `IFS`, `STDEV.S`, `CONCAT`, ...)
  are written with the `_xlfn.` prefix the file format requires; otherwise Excel
  shows `#NAME?`. Quoted sheet names and string literals are left as written.
- Formulas work in styled cells, in table body rows, and with number formats
  (`(Formula(...), F.THOUSANDS_SEP)`). A pandas or polars column of `Formula`
  objects is written as formulas too.
- Not supported: dynamic-array spill formulas (`FILTER`, `UNIQUE`, `SORT`,
  `SEQUENCE` returning arrays), which need cell metadata xlspy does not write, and
  `LET`/`LAMBDA`-family functions, which need `_xlpm.` parameter names. These
  raise a `ValueError` where detectable. Wrap array functions so they return one
  value (for example `COUNTA(UNIQUE(A:A))`).
- xlspy does not evaluate formulas. Its syntax check only verifies balanced
  parentheses and quotes, and length; Excel reports any other error on open.

### Data validation

```python
from xlspy import DataValidation

DataValidation("list", "B2:B1000", formula1='"New,Active,Closed"')           # literal dropdown
DataValidation("list", "C2:C1000", formula1="Lists!$A$2:$A$9")               # dropdown from a range
DataValidation("whole", ["D2:D50", "F2"], formula1=1, formula2=100,
               error="Enter a whole number from 1 to 100.")                  # operator defaults to between
DataValidation("decimal", "E2:E50", formula1=0, operator="greater_than")
DataValidation("date", "G2:G50", formula1=datetime.date(2026, 1, 1), operator="greater_than_or_equal")
DataValidation("text_length", "H2:H50", formula1=20, operator="less_than_or_equal")
DataValidation("time", "I2:I50", formula1=datetime.time(8), formula2=datetime.time(17), operator="between")
DataValidation("custom", "J2:J50", formula1="ISNUMBER(J2)")
```

Pass the rules as `data_validations=[...]` to `write_sheet` or `write_dataframe`.
Options: `allow_blank`, `show_dropdown` (list rules only; `False` hides the arrow),
`show_input_message` with `prompt_title` and `prompt`, `show_error_message` with
`error_title`, `error`, and `error_style` (`"stop"`, `"warning"`, `"information"`).
Invalid combinations (an operator on a list rule, a missing `formula2`, a prompt
without `show_input_message`, a literal list over 255 characters) raise a
`ValueError` when the rule is created.

### Hyperlinks

A `Hyperlink` is the cell's value. Its text is the label, which defaults to the
target.

```python
from xlspy import Hyperlink

writer.write_sheet([
    ["Report", Hyperlink("https://example.com/report", text="Open report", tooltip="Q3")],
    ["Sales", Hyperlink("#'Sales Data'!B2", text="Go to sales")],    # cell in this workbook
    ["Contact", Hyperlink("mailto:team@example.com")],
])
```

Allowed targets are `http`, `https`, `mailto` and `ftp` URLs, and internal
`#Sheet!A1` locations. Other schemes (such as `javascript:` or `file:`) and
relative file paths raise `ValueError`. An internal link to a sheet that is not in
the workbook fails when the workbook is saved. Linked cells use a built-in
hyperlink look unless the cell has its own style.

### Images

```python
from xlspy import Image

writer.write_sheet(rows, images=[
    Image("company_logo.png", anchor="A1", width=180, height=60, description="Company logo"),
    Image("photo.jpg", anchor="H2"),          # size taken from the file
])
```

PNG and JPEG only, embedded byte-for-byte; no re-encoding or Pillow dependency.
`width` and `height` are pixels. Give one and the other keeps the aspect ratio;
give neither and the file's own size is used. Pictures and charts can share a
sheet. Identical files are stored once.

### Printing

```python
from xlspy import PageMargins, PrintSettings

PrintSettings(
    orientation="landscape",              # or "portrait"
    paper_size="A4",                      # names (A4, letter, legal, A3, ...) or OOXML codes
    margins=PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75),
    print_area="A1:H100",                 # several ranges: "A1:H100,J1:K5"
    repeat_rows="1:2",                    # rows repeated on each page
    repeat_columns="A:B",
    fit_to_width=1, fit_to_height=0,      # fit to page; 0 = as many pages as needed
    # scale=80,                           # or a fixed percentage (10-400), not with fit_to_*
    center_horizontally=True,
    header="&LSales report&RPage &P of &N",   # raw Excel header/footer codes, max 255 characters
    footer="&CDraft && confidential",         # a literal ampersand is written as &&
    row_breaks=[25, 50],                  # page break after row 25 and row 50 (1-based)
    col_breaks=[4],
)
```

### Charts

The original chart options still work: `Chart("bar", ...)` is a column chart,
`"line"` and `"pie"` are unchanged. New in XLSX:

- `"barh"`: horizontal bars. The category axis is vertical.
- `"scatter"`: XY charts. Series take `x_values` (a numeric range); markers are
  shown by default and `line=True` joins the points.
- `x_axis` (the category or X axis) and `y_axis` (the value axis) are
  `ChartAxis(title=, minimum=, maximum=, major_unit=, number_format=, gridlines=, reverse=, visible=)`.
  In a horizontal bar chart `x_axis` still means the category axis, drawn vertically.
- `legend`: `ChartLegend("bottom")` (positions `right`, `left`, `top`,
  `bottom`, `top_right`, `overlay=True`), or `False` to hide it.
- `show_data_labels=True` or `data_labels=DataLabels(number_format=, position=, show_category=, show_series_name=)`.
- `ChartSeries(color="#C00000", marker=ChartMarker("diamond", 8), line=True)`:
  the series colour fills bars and sets the line or marker colour.
- `title_size=16` (points) for the chart title.

```python
from xlspy import Chart, ChartAxis, ChartLegend, ChartMarker, ChartSeries

Chart(
    "scatter",
    [ChartSeries("D2:D20", x_values="B2:B20", title="Revenue", color="#C00000",
                 marker=ChartMarker("circle", 7))],
    title="Revenue against units", anchor="H2",
    x_axis=ChartAxis(title="Units sold", minimum=0),
    y_axis=ChartAxis(title="Revenue", minimum=0, number_format="#,##0"),
    legend=ChartLegend("bottom"),
    show_data_labels=True,
)
```

Pie charts have no axes. Options that a chart type cannot use (a marker on a bar
chart, a colour on a pie) are rejected with a `ValueError`.

### Streaming large exports

`write_sheet` consumes its input exactly once, row by row:

- **XLSX** encodes rows into a temporary spool (kept in memory up to
  `spool_max_bytes`, 8 MiB by default, then on disk). The sheet is assembled once
  the last row is known, so the dimension and autofilter are exact. Writes are
  bounded by the spool setting and the 100-row width sample, not by the row count.
- **XLSB** writes each row to the archive as it is produced.
- **Shared strings** (the default) keep one table of every distinct string in
  memory until save. For high-cardinality exports pass `useSharedStrings=False`:
  text is then stored inline in each cell, so memory no longer grows with the
  number of distinct strings. Files are somewhat larger. The C extension encodes
  inline XLSB rows too; rows it cannot handle (formatted cells, or text it
  refuses) go to the Python encoder (see the benchmarks).

```python
with XlsbWriter("sales.xlsb", useSharedStrings=False) as writer:
    writer.add_sheet("Sales")
    writer.write_sheet(cursor_rows(cursor))      # a generator over a DB cursor
```

Examples: [`xlsx_streaming_export.py`](../examples/xlsx_streaming_export.py) (DB-API cursor, XLSX or XLSB).

### Cell text

- **Length limit.** Excel stores at most 32,767 characters in a cell, counted in
  UTF-16 code units: a character outside the Basic Multilingual Plane (such as
  most emoji) counts as two, and a combining accent counts as one. Every writer
  path (XLSX and XLSB, shared and inline strings, styled cells, DataFrames,
  hyperlink text and cached formula results) raises `ValueError` for a longer
  value. Nothing is truncated, and the workbook is not published (see
  *Failure handling* below). Excel 16.0 refuses to open XLSX and inline-XLSB
  files whose cells exceed the limit, which is why it is enforced.
- **Invalid Unicode.** Text with a lone surrogate (for example `"\ud800"`) raises
  `UnicodeEncodeError`.
- **Characters XML cannot hold.** In XLSX, control characters, carriage returns
  and U+FFFE/U+FFFF are written in OOXML's `_xHHHH_` form (`"\x07"` becomes
  `_x0007_`), and literal text that already looks like `_xHHHH_` has its
  underscore escaped as `_x005F_`. Excel and xlspy's reader turn these back into
  the original text. Some other readers (openpyxl among them) show the escapes
  as written. XLSB stores text as UTF-16 and needs no escaping.

### Failure handling

A workbook reaches its target only after it has been written completely. If the
input raises, a value is refused, or anything fails while saving, the exception
leaves the `with` block and the target is left as it was:

- **A file path** is written to a temporary file in the same directory (named
  `.<name>.<random>.tmp`) and moved into place with `os.replace` when the
  workbook is complete. A failed write creates no file and keeps an existing
  file at that path unchanged. The final step is a rename, not a forced write to
  disk: xlspy does not call `fsync`.
- **A file object** (`io.BytesIO`, an open file, a socket stream) is written to
  a temporary spool, in memory up to `spool_max_bytes` (8 MiB for `XlsbWriter`)
  and on disk beyond that, and copied into the stream on success. A failed write
  leaves the stream untouched. The copy itself cannot be undone: if the stream
  raises part-way through it (a full disk, for example), the error reaches the
  caller and the stream holds only part of the workbook, which must be
  discarded. The workbook is appended at the stream's current position.

After a failure the writer refuses further `write_sheet` and `save` calls.
`save()` writes the workbook parts; the file is published when the writer
closes, so a writer that is closed without a successful save publishes
nothing. The context manager calls both.

### Combined example

[`xlsx_report_features.py`](../examples/xlsx_report_features.py) combines formulas,
a table, styles, validation, hyperlinks, a logo, print settings, and two charts in one
report.

### Feature support

| Feature | XLSX | XLSB |
|---|---|---|
| Formulas (`Formula`), cached values, array formulas | Yes | No (`TypeError`) |
| Data validation (all seven types, prompts, errors) | Yes | No (`TypeError`) |
| Hyperlinks (external, internal, tooltip, styled) | Yes | No (`TypeError`) |
| Images (PNG, JPEG) | Yes | No (`TypeError`) |
| Print settings, print area, titles, breaks, headers | Yes | No (`TypeError`) |
| Row and column groups (`row_groups`, `column_groups`) | Yes | Yes |
| Column, bar, line, pie charts (original options) | Yes | Yes |
| Horizontal bar, scatter, axes, legend, labels, markers, colours | Yes | No (`ValueError`) |
| Streaming: one-pass input, spooled rows | Yes | Yes (rows written as produced) |
| Inline strings (`useSharedStrings=False`) | Yes | Yes (C row encoder; Python fallback) |
| Chart caches from streamed rows | n/a (no caches) | Yes (only chart cells are kept) |

## Writing pandas / polars DataFrames (optional)

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

### Options

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

See [`examples/dataframe_*.py`](../examples/) and
[`examples/polars_to_xlsx_xlsb.py`](../examples/polars_to_xlsx_xlsb.py) for full,
runnable scripts, indexed in [`examples/README.md`](../examples/README.md).

## Updating an Existing Workbook

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

## Reading XLSB and XLSX Files

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

## Streaming from a Database (Netezza)

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
