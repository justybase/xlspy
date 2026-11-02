# Formatting and layout

[Guide](../usage.md) · [Home](../../README.md) · [Examples](../../examples/README.md)

## Cell Formatting (Number, Date, DateTime)

`xlspy` supports custom cell formatting for numbers, dates, and datetimes in both XLSB and XLSX output. Pass a `(value, format_string)` tuple to apply a format to a specific cell.

Without a format, both writers show a `date` as `yyyy-mm-dd` and a `datetime` as
`yyyy-mm-dd hh:mm`, whatever the language of the Excel that opens the file.

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

### Open ranges and per-cell styles

A `styles` key can also be open-ended, so a streamed sheet can be styled
without knowing its length: whole columns (`"B:D"`), columns from a row down
(`"B2:D"`), or whole rows (`"1:1"`). Only the cells of the rows you write are
styled; an open range adds no empty rows.

For a style that changes from cell to cell, put it in the row as
`(value, Style)`, just like `(value, number_format)`. The cell's `Style` goes on
top of the header and range styles, overriding only the properties it sets:

```python
ok = Style(fill="E2EFDA")
late = Style(fill="FCE4D6", bold=True)

with XlsbWriter("status.xlsb") as writer:          # XlsxWriter works the same way
    writer.add_sheet("Orders")
    writer.write_sheet(
        ([order.id, (order.status, late if order.late else ok), order.amount]
         for order in orders),
        styles={"C:C": Style(number_format="#,##0.00")},
    )
```

`(value, Style)` cells produce the same file as one `styles` entry per cell
(`{"B2": late, "B3": ok, ...}`), but much faster: Python does not have to
build and parse an A1 key per cell. The workbook model uses them when it
saves styled cells.

### Performance of styled writes

Styles cost little on top of a plain write. The native encoders write header,
range, `(value, format)` and `(value, Style)` cells, and rows with a height,
hidden flag or outline level. Each distinct style combination is resolved
once and then reused. Medians from `tools/bench_styles.py` (10,000 rows × 12
columns, Python 3.14, Windows, native extension on, seconds):

| Scenario | XLSX 0.8.1 | XLSX now | XLSB 0.8.1 | XLSB now |
|---|---:|---:|---:|---:|
| no styles | 0.045 | 0.048 | 0.027 | 0.028 |
| `header_style` | 0.104 | 0.050 | 0.028 | 0.028 |
| header + `"B2:C10000"` range | 0.122 | 0.055 | 0.046 | 0.035 |
| four overlapping ranges | 0.438 | 0.060 | 0.354 | 0.036 |
| `(value, format)` cells | 0.126 | 0.048 | 0.091 | 0.029 |
| `(value, format)` + range | 0.387 | 0.060 | 0.420 | 0.039 |
| one `styles` entry per cell (60,000) | 21.0 | 0.183 | 20.4 | 0.171 |
| the same with `(value, Style)` cells | — | 0.054 | — | 0.032 |
| range + row heights and groups | 0.146 | 0.066 | 0.145 | 0.046 |
| workbook model, per-cell styles | 20.6 | 0.112 | 21.1 | 0.092 |

Practical advice:

- Prefer one range per column or block (`"B:B"`, `"B2:D"`) over many small
  ranges; for styles that vary per cell use `(value, Style)`.
- Reuse `Style` objects instead of creating an equal one per cell: identical
  objects are recognised without comparing their contents.
- Run `python tools/bench_styles.py --help` to measure your own shapes;
  `--compare before.json after.json` prints the change per scenario.

### Column widths

Columns you do not size with `column_widths` get an automatic width from the
first 100 rows, close to Excel's own *AutoFit* but never narrower:

- Text is measured per character with widths taken from Excel (Calibri 11):
  `W`, `M`, `@` and capitals take more room than `i` or `l`, and Chinese,
  Japanese or Korean characters about two digits each.
- Numbers and dates are measured as Excel shows them: `(1234567.891, "#,##0.00")`
  as `1,234,567.89`, General numbers with at most 11 characters, and dates
  with their format. Month and day names, `TRUE`/`FALSE` and separators
  depend on the reader's language, so they get room for long names.
- Bold text, a `Font(size=...)` from the cell's style, and wrapped text (the
  longest line) are taken into account.
- Header cells under a filter or table drop-down button get two extra
  characters for the button.
- A column is never narrower than Excel's default width (8.43).

Values beyond the first 100 rows are not measured; give `column_widths` for
columns whose longest values come later. `tools/score_column_widths.py`
compares the estimate with Excel's AutoFit widths recorded by
`tools/calibrate_column_widths.py`.

#### Fixed widths

`column_widths` sets the width of chosen columns and skips the automatic
measurement for them. It is a mapping from a column to a width in Excel's
character units (the value shown in *Column Width*):

```python
from xlspy import XlsbWriter, XlsxWriter  # both take the same parameter

with XlsxWriter("report.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_sheet(
        rows,
        column_widths={"A": 8, "B": 30, "E": 60},   # column letters
        # column_widths={0: 8, 1: 30, 4: 60},       # or 0-based indexes
    )
```

- Keys are column letters (`"A"`, `"AB"`) or 0-based integers. Any other key
  type raises `TypeError`.
- Widths are numbers from `0.5` to `255`. Other values raise `ValueError`.
- Columns that are not listed keep the automatic width described above.
- `XlsbWriter.write_sheet()`, `write_dataframe()` and the `to_xlsx()` and
  `to_xlsb()` helpers accept the same mapping.

### Grid lines

`show_grid_lines=False` hides the sheet's grid lines, like Excel's
*View → Gridlines* checkbox or openpyxl's `ws.sheet_view.showGridLines = False`.
Both writers and `write_dataframe()` accept it; the default `True` writes the
same file as before. The setting belongs to each sheet:

```python
with XlsxWriter("report.xlsx") as writer:       # XlsbWriter works the same way
    writer.add_sheet("Summary")
    writer.write_sheet(summary_rows, show_grid_lines=False)
    writer.add_sheet("Data")
    writer.write_sheet(data_rows)                # grid lines shown
```

The [workbook model](model.md#openpyxl-like-workbook-model) saves
`ws.sheet_view.showGridLines = False` the same way, for new XLSX and XLSB
workbooks as well as loaded ones.

### Tab colour

`tab_color` colours the sheet's tab: an `"RRGGBB"` string (with or without
`#`) or a `Color`, including theme colours. Both writers and
`write_dataframe()` accept it, and the workbook model saves
`ws.sheet_properties.tabColor` for new workbooks:

```python
writer.add_sheet("Summary")
writer.write_sheet(summary_rows, tab_color="#C00000")
writer.add_sheet("Data")
writer.write_sheet(data_rows, tab_color=Color(theme=4, tint=0.4))
```

### Zoom

`zoom=85` (10-400 percent) sets the sheet's zoom in both writers and
`write_dataframe()`; the workbook model saves `ws.sheet_view.zoomScale`.

### Headings, zeros and right-to-left

`show_row_col_headers=False` hides the row numbers and column letters,
`show_zero_values=False` shows cells holding zero as empty, and
`right_to_left=True` puts column A on the right. All three are bools in both
writers and `write_dataframe()`, and the workbook model saves
`ws.sheet_view.showRowColHeaders`, `showZeros` and `rightToLeft`:

```python
writer.write_sheet(rows, show_grid_lines=False, show_row_col_headers=False)  # a clean dashboard
```

### Headers of several rows

`header_rows=2` (or more) declares a header of several rows, for example group
titles above the column names. The automatic filter then starts at the last
header row, so its buttons sit on the column names, and `header_style` covers
every header row (both writers):

```python
writer.write_sheet(
    [
        ["Customer", None, "Sales", None],
        ["Name", "Code", "Q1", "Q2"],
        ["North", "N", 120, 140],
    ],
    header_rows=2,
    header_style=Style(bold=True, fill="#1F4E78", font_color="#FFFFFF"),
    merges=["A1:B1", "C1:D1"],
    freeze_panes="A3",
)
```

Here the filter covers `A2:D3`. `styles` ranges still overlay the header style,
and the column widths leave room for the filter buttons on row 2. Without data
rows below the header no filter is added. The default `header_rows=1` writes the
same file as before.

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

`XlsbWriter` writes the same rules. Its `DifferentialStyle` takes the font
name, size, colour, bold, italic, underline and strike; pattern fills; the
top/bottom/left/right borders; horizontal alignment and wrapping; protection
and a number format. The other options (`GradientFill`, `vertAlign`, diagonal
borders, vertical alignment and so on) raise `TypeError` in XLSB, as does a
`CellIsRule` with a text operator such as `containsText` (use `TextRule`).

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
or `title_cell` reference) and optional `categories` (the point labels: axis
labels, or the slice names in a pie chart's legend).
Both writers emit the same DrawingML parts (`drawings/drawingN.xml`,
`charts/chartN.xml`) and take every chart option below; the XLSX writer links
them with a `<drawing>` sheet element, the XLSB writer with a `BrtDrawing`
record. XLSB chart parts also mirror real Excel files: cached values inside
every reference (a `Formula` cell contributes its cached `value`, an empty
cell is a gap), text categories via `strRef`, and an explicit `varyColors`
flag.

**XLSB parity:** `XlsbWriter` supports the same `Style` object model
(`Font`, `PatternFill`, `GradientFill`, `Border`, `Alignment`, `Protection`)
as well as `merges`, `column_widths`, `hidden_columns`, `row_heights`,
`hidden_rows`, `row_groups`, `column_groups`, `freeze_panes`,
`show_grid_lines`, and `auto_filter`. A few properties have no BIFF12 representation and raise
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
- Grouped rows are written by the Python row encoder rather than the native one,
  like rows with a custom height. Sheets without groups are unchanged and keep
  the fast path.
- Updaters (`XlsxUpdater`, `XlsbUpdater`) do not create or change groups.

[`xlsx_grouping.py`](../../examples/xlsx_grouping.py) builds a two-level report
with region subtotals and quarter columns (pass `--out grouping.xlsb` for XLSB).
