# Formulas, charts and other objects

[Guide](../usage.md) · [Home](../../README.md) · [Examples](../../examples/README.md)

## Formulas, validation, hyperlinks, images, and printing

`XlsxWriter` and `XlsbWriter` take the same parameters for these features
(`data_validations=`, `images=`, `print_settings=`). XLSB formulas use the
BIFF12 grammar described in [XLSB formulas](#xlsb-formulas-and-the-bulk-data-path);
see [feature support](#feature-support) for the full matrix.

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

**What clicking does.** The whole cell is the hot area, and an internal link
keeps the reader inside the workbook: clicking `#'Sales Data'!B2` moves the
selection to B2 of *Sales Data*, which is how you navigate between sheets of a
report. A web or mail target opens the browser or the mail client instead.
Excel may ask for `Ctrl` first: *File > Options > Advanced* and clear
**Use CTRL+Click to follow hyperlink** to follow links with a plain click. To
select a linked cell without following it, press and hold the mouse button until
the pointer changes from the pointing hand to the selection cross and then
release, or use the arrow keys.

**Reading links back.** `load_workbook()` restores them as `cell.hyperlink`
(the label stays in `cell.value`); the read-only streaming readers do not
report hyperlinks, and changing a link in a loaded workbook raises
`NotImplementedError` at save time.

**XLSB** takes `Hyperlink` cells too, with the same targets, tooltip and link look
as XLSX. External links are stored in the sheet's relationships and internal
ones as location records. The `HYPERLINK` function is also available in XLSB when
you want the target to come from a formula:

```python
from xlspy import Formula

writer.write_sheet([
    ["Sales", Formula('HYPERLINK("#Sales!B2","Go to sales")', value="Go to sales")],
    # a sheet name with spaces: HYPERLINK("#'Sales Data'!B2","Go to sales")
])
```

Excel evaluates it on open and jumps to the target on click just like a
`Hyperlink` cell. Differences: no tooltip and no link styling (style the cell
yourself). In XLSB a cell without `value=` keeps a zero placeholder for
readers that do not recalculate. Use `,` as the argument separator and English
function names: that is the invariant form the file formats store, whatever
separator your Excel displays.

### Images

```python
from xlspy import Image

writer.write_sheet(rows, images=[
    Image("company_logo.png", anchor="A1", width=180, height=60, description="Company logo"),
    Image("photo.jpg", anchor="H2"),          # size taken from the file
])
```

The first argument is a file path, the image bytes, or a binary file object
such as `io.BytesIO` (read from its current position), so a chart rendered in
memory needs no temporary file:

```python
buffer = io.BytesIO()
figure.savefig(buffer, format="png")      # e.g. matplotlib
buffer.seek(0)
writer.write_sheet(rows, images=[Image(buffer, anchor="H2")])
```

PNG, JPEG, GIF and BMP, embedded byte-for-byte; no re-encoding or Pillow
dependency. `width` and `height` are pixels. Give one and the other keeps the aspect ratio;
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
`"line"` and `"pie"` are unchanged. The supported chart types are `"bar"`, `"barh"`,
`"line"`, `"area"`, `"pie"`, `"doughnut"`, `"scatter"` and `"radar"`. The other
names in `CHART_TYPES` (`area3d`, `bar3d`, `bubble`, `line3d`, `pie3d`,
`projected_pie`, `stock`, `surface` and `surface3d`) raise `ValueError` when the
chart is created. A radar series takes no trendline or error bars (`ValueError`).
New in XLSX:

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
- `"area"`: an area chart, and `"doughnut"`: a pie with a hole (one series,
  no axes, slices coloured from the theme).
- `grouping="stacked"` or `"percent_stacked"` stacks the series of a `"bar"`,
  `"barh"`, `"line"` or `"area"` chart (each category as 100% for
  `"percent_stacked"`).
- Combined charts: `ChartSeries(chart_type="line")` draws that series as a line (or
  `"bar"`, `"area"`) in a `"bar"`, `"line"` or `"area"` chart, and
  `secondary_axis=True` reads it against a second value axis on the right
  (`Chart(y2_axis=ChartAxis(title=...))`).
- `ChartSeries(trendline=Trendline("linear"|"exponential"|"logarithmic"|"polynomial"|"power"|"moving_average",
  order=, period=, forward=, backward=, display_equation=, display_r_squared=))` and
  `ChartSeries(error_bars=ErrorBars("fixed"|"percentage"|"std_dev"|"std_err", value, direction=, axis=, end_cap=))`.
- Loaded charts (XLSX and XLSB) keep their XML; `LoadedChart.title`, `x_axis_title`,
  `y_axis_title` and `legend_position` can be changed.

```python
Chart("bar", [ChartSeries("B2:B13", categories="A2:A13", title_cell="B1"),
              ChartSeries("C2:C13", categories="A2:A13", title_cell="C1")],
      title="Sales by region", grouping="stacked")
```

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

### Chart formatting

Charts also take the chart-area and plot-area styling and the gap and hole sizes:

- `gap_width=0..500` (percent) sets the space between bar groups; bar and horizontal bar charts only.
- `hole_size=10..90` (percent) sets the hole of a doughnut; doughnut charts only.
- `fill`, `border` and `plot_fill` take colours (`"#RRGGBB"`); `fill` and `border` style the chart area,
  `plot_fill` the plot area.
- `font_size` (points) and `font_color` set the chart text.

```python
Chart("doughnut", [ChartSeries("L22:L24", categories="K22:K24")],
      anchor="F9", width=8, height=7, hole_size=60, fill="#FFFFFF", border="#D9D9D9", font_size=9)
```

These work in XLSX and XLSB.

### Sparklines

A sparkline is a tiny chart inside one cell. `SparklineGroup(data, location, kind=)` draws one
sparkline per row or column of `location`, reading the matching row or column of `data`:

```python
from xlspy import SparklineGroup

writer.write_sheet(rows, sparklines=[
    SparklineGroup(data="B2:D4", location="E2:E4", kind="line", color="#2E75B6", markers=True,
                   last_point=True),
    SparklineGroup(data="B2:D2", location="F2", kind="column", color="#70AD47"),
])
```

- `kind`: `"line"` (default), `"column"` or `"win_loss"`.
- `color`, `negative_color`, `marker_color`, `high_color`, `low_color`, `first_color`, `last_color`,
  `axis_color` are `"#RRGGBB"` colours.
- `markers`, `high_point`, `low_point`, `first_point`, `last_point` and `negative_points` are booleans.
- `empty_cells`: `"gap"` (default), `"zero"` or `"connect"`. `show_hidden`, `show_axis` and
  `right_to_left` are booleans. `axis_min` and `axis_max` take a number (a fixed end of the
  vertical axis), `"group"` (one scale for the whole group) or `"individual"` (the default).

Both writers support sparklines; `data` may name another sheet (`"'Sales data'!B2:M2"`).
Renaming the sheet rewrites the references in a loaded XLSX workbook.
The library does not shift sparkline or slicer ranges when rows or columns are inserted or deleted
in a loaded workbook; update those ranges yourself if the layout changes.

### Table slicers

A slicer filters a table by one of its columns. The table must have a header row and be on the
same sheet as the slicer:

```python
from xlspy import ExcelTable, Slicer

writer.write_sheet(rows, tables=[ExcelTable("Sprzedaz", "A1:H200")],
                   slicers=[Slicer(table="Sprzedaz", column="Region", anchor="K4",
                                   width=144, height=180, caption="Region")])
```

`Slicer(table, column, anchor="H2", width=144, height=180, caption=None, name=None, columns=1)`.
`column` is the header text; `name` and `caption` default to it. The slicer name must be unique
in the workbook, and so must the cache name, `Slicer_<column>`, so one column name can carry only
one slicer per workbook. Validation runs before any rows are written when the table lists its
`columns`; when the names come from the header row, an unknown column fails the sheet once that
row has streamed.

The slicer's effect is on rows, not on formulas. A plain `SUM` or `SUMIFS` ignores the filter.
To make totals react, add a helper column with `SUBTOTAL(103, A2)` (1 when the row is visible
under the table's filter) and sum with `SUMIFS(..., Visible, 1)`. The dashboard example does this
for its KPI cards.

Both writers support slicers; XLSB stores them in binary slicer and cache parts, as Excel does.
A loaded workbook keeps its slicers when it is saved. Removing a sheet that holds table slicers
from a loaded workbook is refused, because the slicer caches would be left behind. Pivot-table
slicers and timelines are not supported.

### Shapes and KPI cards

`Shape` adds a rectangle, rounded rectangle, ellipse or text box to a sheet. It works in XLSX and
XLSB:

```python
from xlspy import Shape

writer.write_sheet(rows, shapes=[
    Shape(anchor="A5", width=196, height=60, kind="round", fill="#F2F7FC", line="#BDD7EE",
          text="Sprzedaż", font_size=20, bold=True, align="center", valign="center",
          link_cell="T5", name="Karta sprzedaży"),
])
```

- `anchor` is the top-left cell; `width` and `height` are pixels.
- `kind`: `"rect"`, `"round"`, `"ellipse"` or `"textbox"`.
- `text`, `font_size`, `font_color`, `bold`, `align` (`"left"`, `"center"`, `"right"`) and
  `valign` (`"top"`, `"center"`, `"bottom"`) style the text. `fill` and `line` are colours.
- `link_cell` puts a cell's value into the shape's first paragraph and keeps it updated when the cell
  changes. Use it for KPI cards that show a formula result.

Shapes and charts can share a sheet, and their drawing ids are kept unique.

### Rich text

`CellRichText` puts several fonts in one cell (both writers, the workbook
model and XLSX updates). `XlsbWriter` stores it in the shared-string table, as
Excel does, also with `use_shared_strings=False`. The classes follow openpyxl:

```python
from xlspy import CellRichText, InlineFont, TextBlock

total = CellRichText(
    "Total: ",
    TextBlock(InlineFont(b=True, color="C00000", sz=14), "1 234"),
    " EUR",
)
writer.write_sheet([["Summary"], [total]])
```

`InlineFont` takes `b`, `i`, `u`, `strike`, `sz`, `color`, `rFont`,
`vertAlign`, `family`, `charset` and `scheme`. `str(total)` is the plain text,
which is also what `ExcelReader` and `load_workbook()` return unless
`load_workbook(..., rich_text=True)` is used.

### More conditional formatting

Besides `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule` and
`IconSetRule`, both writers have Excel's other built-in rules. Each takes the
range first and a `dxf=DifferentialStyle(...)` for the look:

| Rule | Excel's menu entry |
|---|---|
| `TopBottomRule(sqref, rank=10, percent=False, bottom=False)` | Top/Bottom Rules |
| `AboveAverageRule(sqref, below=False, equal_average=False, std_dev=None)` | Above/Below Average |
| `TextRule(sqref, "contains" \| "notContains" \| "beginsWith" \| "endsWith", text)` | Text that Contains |
| `DuplicateRule(sqref, unique=False)` | Duplicate Values |
| `TimePeriodRule(sqref, "today" \| "last7Days" \| "thisMonth" \| ...)` | A Date Occurring |
| `BlanksRule(sqref, "blanks" \| "no_blanks" \| "errors" \| "no_errors")` | Blanks / Errors |

In XLSB the rules take the same `DifferentialStyle` options as `CellIsRule`
(see [Conditional formatting](styling.md#conditional-formatting)).

Pie and doughnut charts have no axes. Options that a chart type cannot use (a
marker on a bar chart, a colour on a pie, `grouping` on a pie) are rejected
with a `ValueError`.

### Document properties

`set_properties()` fills in what Excel shows under *File → Info* (both writers):

```python
import datetime

with XlsxWriter("report.xlsx") as writer:          # XlsbWriter works the same way
    writer.set_properties(
        title="Sales Q3", subject="Sales", author="Finance team",
        manager="CFO", company="ACME", category="Reports",
        keywords="sales, q3", comments="Generated nightly", status="Final",
        created=datetime.datetime(2024, 10, 1, tzinfo=datetime.timezone.utc),
        custom={"Project": "Alpha", "Revision": 3, "Approved": True},
    )
    writer.add_sheet("Data")
    writer.write_sheet(rows)
```

Every field is optional. `created` defaults to the time of the call; a naive
datetime is taken as UTC. `custom` holds str, int, float, bool, date or
datetime values, shown under *Properties → Advanced Properties → Custom*.
Without `set_properties()` no `docProps` parts are written, so repeated writes
of the same data stay byte-for-byte identical. The workbook model saves
`wb.properties` (title, creator, subject, keywords, ...) the same way.

### Combined example

[`xlsx_report_features.py`](../../examples/xlsx_report_features.py) combines formulas,
a table, styles, validation, hyperlinks, a logo, print settings, and two charts in one
report.

### Feature support

| Feature | XLSX | XLSB |
|---|---|---|
| Formulas (`Formula`), cached values, legacy array formulas | Yes | Yes (supported BIFF12 grammar below) |
| Data validation (all seven types, prompts, errors) | Yes | Yes |
| Hyperlinks (external, internal, tooltip, styled) | Yes | Yes ([Hyperlinks](#hyperlinks)) |
| Images (PNG, JPEG, GIF, BMP; files or in-memory bytes) | Yes | Yes |
| Print settings, print area, titles, breaks, headers | Yes | Yes |
| Row and column groups (`row_groups`, `column_groups`) | Yes | Yes |
| Classic cell notes (`Comment`) | Yes | Yes |
| User-defined names, sheet/workbook protection | Yes | Yes |
| Formula text reading (`data_only=False`) | Yes | Yes (supported BIFF12 tokens below) |
| Sparse template edits (`update_cells`) | Yes (XLSX/XLSM updater) | Yes |
| Column, bar, line, pie charts (original options) | Yes | Yes |
| Horizontal bar, scatter, area, doughnut, stacked, axes, legend, labels, markers, colours | Yes | Yes |
| Document properties (`set_properties`, custom properties) | Yes | Yes |
| Rich text cells (`CellRichText`) | Yes | Yes |
| Top/bottom, average, text, duplicate, date and blank rules | Yes | Yes |
| Conditional format styles (`DifferentialStyle`) | Yes | Yes, except gradient fills, diagonal borders, vertical alignment and rare font options (`TypeError`) |
| Sheet zoom (`zoom`) | Yes | Yes |
| Streaming: one-pass input, spooled rows | Yes | Yes (rows written as produced) |
| Inline strings (`use_shared_strings=False`) | Yes | Yes (native row encoder; Python fallback) |
| Chart caches from streamed rows | n/a (no caches) | Yes (only chart cells are kept) |
| Chart formatting (`gap_width`, `hole_size`, `fill`, `border`, `plot_fill`, `font_size`, `font_color`) | Yes | Yes |
| Shapes and text boxes (`Shape`, `link_cell` KPI cards) | Yes | Yes |
| Sparklines (`SparklineGroup`) | Yes | Yes |
| Table slicers (`Slicer`) | Yes | Yes |
| Headings, zero values, right-to-left (`show_row_col_headers`, `show_zero_values`, `right_to_left`) | Yes | Yes |
| Pivot tables (`PivotTable`, compact layout) | Yes | Yes (no dates, booleans or NaN in the fields) |
| Pivot-table slicers, timelines | No | No |

## Pivot tables (XLSX and XLSB)

`PivotTable` creates a pivot table over a table written in the same `write_sheet`
call. Pass the pivots with `pivots=`; `XlsxWriter` and `XlsbWriter` take the same
arguments:

```python
from xlspy import ExcelTable, PivotTable, PivotValue, XlsxWriter

rows = [
    ["Region", "Product", "Quarter", "Sales", "Units"],
    ["North", "Coffee", 1, 1200, 40],
    ["North", "Tea", 1, 800, 50],
    ["South", "Coffee", 2, 1600, 58],
    # ...
]

with XlsxWriter("pivot.xlsx") as writer:
    sales = ExcelTable("SalesTable", "A1:E9")
    writer.add_sheet("Data")
    writer.write_sheet(
        rows,
        tables=[sales],
        pivots=[
            PivotTable(
                source=sales,
                rows="Region",
                columns="Quarter",
                values=[PivotValue("Sales", "sum", number_format="#,##0")],
                location="A12",          # below the data, on the same sheet
            ),
        ],
    )
```

To put a pivot on another sheet, name it with `sheet=` (or `location="Summary!A3"`) in
the call that writes the source table. That sheet must be added after the source
sheet. `XlsxWriter` also accepts a sheet that is never written; `XlsbWriter` needs it
written with `write_sheet` (an empty list is enough).

- **Layout:** one `rows` field; `columns` takes at most one field; `filters` add
  page fields above the report; `values` takes one or more `PivotValue(field,
  function, name=None, number_format=None)` with `function` one of `"sum"`,
  `"count"`, `"average"`, `"min"`, `"max"`.
- **Values:** `sum`, `average`, `min` and `max` skip blank and text cells; `count`
  counts the non-blank cells.
- **Output:** the report cells are written by the writer, from the rows as they
  stream past. The cache holds no records and refreshes when the file opens, so
  Excel rebuilds the layout and shows its own (localised) labels.
- **Column widths:** when Excel refreshes a pivot on open it can resize the columns the pivot
  covers. By default xlspy turns Excel's "Autofit column widths on update" off, so a pivot on the
  same sheet as its data keeps the widths the sheet was given. `PivotTable(..., autofit_columns=True)`
  turns it on, as Excel does for new pivots. The option is written for XLSX and XLSB alike.
- **XLSB:** the same pivots as binary parts (the records were mapped from files
  Excel saved). The fields a pivot reads cannot hold dates, booleans or NaN: a
  `ValueError` names the field, and XLSX takes those.
- **Not included:** layouts beyond the compact form, and changes to pivot tables
  already in a loaded workbook (XLSX or XLSB), which are kept as they are. Removing a
  loaded sheet that holds a pivot table is refused.

## Comments and template metadata (XLSX/XLSM/XLSB)

Classic Excel notes are available through `Comment`. These are the notes with
a cell indicator, rather than Microsoft 365 threaded conversations.

```python
from xlspy import Comment, DefinedName, SheetProtection, WorkbookProtection, XlsxWriter

with XlsxWriter("review.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_sheet(
        [["Amount"], [120]],
        comments={"A2": Comment("Confirm this amount", "Jan", width=240, height=100)},
        sheet_protection=SheetProtection(password="review"),
    )
    writer.set_defined_name(DefinedName("Amount", "'Data'!$A$2"))
    writer.set_workbook_protection(WorkbookProtection(password="review"))
```

For binary workbooks use `XlsbWriter("review.xlsb")`; the same metadata API
is available. `comments` is also accepted by `write_dataframe()`, `to_xlsx()`
and `to_xlsb()`. Addresses
are worksheet coordinates, including the header row. Notes may be attached to
empty cells; a note in a merged range must use its top-left cell. Text and
author are required strings; width and height are positive pixel dimensions.
Comment text has the same 32,767 UTF-16 unit limit as cell text.

Read and edit an existing workbook without replacing its entire data region:

```python
from xlspy import Comment, ExcelReader, Formula, XlsxUpdater

updater = XlsxUpdater("review.xlsx")  # XlsmUpdater and XlsbUpdater have the same API
try:
    updater.update_cells("Data", {"A2": 125, "B2": Formula("A2*2", value=250)})
    updater.set_comments("Data", {
        "A2": Comment("Confirmed", "Ola"),
        "C3": None,  # remove this note if present
    })
    updater.save("reviewed.xlsx")
finally:
    updater.close()

with ExcelReader("reviewed.xlsx") as reader:
    notes = reader.get_comments("Data")  # {"A2": Comment(...)}
    names = reader.get_defined_names()
    formulas = reader.read_all("Data", data_only=False)
```

`set_comments()` patches only the supplied addresses. `get_comments()` is
available on both reader and updater and guarantees text and author; returned
dimensions use the defaults (144 × 79 pixels). Untouched note XML/binary records, including
rich formatting, and unrelated VML controls are preserved during updates.
Existing threaded conversations are preserved; creating/editing them, or adding
a classic note to a cell with a threaded comment, is unsupported. Ambiguous
or unsupported VML relationships/shapes cause an error before changes are staged.
Deleting the last note removes unused comment parts and drawing references;
VML parts containing controls are retained.

`set_hyperlinks(sheet_name, {"A2": Hyperlink(...), "B5": None})` sets or
removes the links of the given XLSX cells and keeps the others; a link whose
range starts at a given cell is replaced whole. Only the link changes, not the
cell's value or style, and a removed link's relationship goes with it.

`update_cells(sheet_name, cells)` accepts a mapping from single-cell addresses
to strings, finite numbers, booleans, dates/datetimes, `Formula`, or `None`.
Existing styles and notes stay attached to the same addresses. Dates are stored
as Excel serials and use the target cell's existing number format. `None` clears
the value while retaining formatting and metadata. A `(value, format_string)`
wrapper uses its value, as in the existing updater. Strings are written inline;
other cells and shared-string indexes retain their contents. The method expands
worksheet dimensions, but does not resize tables, filters or pivot sources.
It materializes the selected worksheet XML or binary part; use `replace_sheet_data_stream()`
for bounded-memory replacement of large data regions.

Formula edits request recalculation in Excel and remove the obsolete calculation
chain. This library does not calculate results; `Formula.value` supplies an
optional cached result. Point edits into shared, array or data-table formula
groups are refused, and `update_cells()` accepts only scalar new formulas.
Edits into merged ranges must use the top-left cell. Each new updater operation
stages all affected parts atomically; `save()` publishes them using a temporary file.
Ordinary `replace_sheet_data()` and its streaming variant keep notes at their
original addresses, including notes outside a newly shortened data region.

`DefinedName(name, formula, scope=None)` represents an Excel expression, constant
or range without evaluating it. `scope=None` creates a global name; a worksheet
title creates a local name. Writer and updater expose `set_defined_name()`,
`get_defined_names()` and `remove_defined_name(name, scope=None)`. Names replace
existing names case-insensitively within their scope. Built-in `_xlnm.*` names
remain managed by the print/filter APIs and are excluded from `get_defined_names()`.
Local worksheet titles must match the workbook; the writer checks them on save.
XLSB names retain their ordinal slots so existing binary formula references
remain valid. Removing a name replaces its slot with a hidden internal
`_xlspy_deleted_*` name whose value is `#NAME?`; these slots are not listed.

`SheetProtection(enabled=True, select_locked_cells=True,
select_unlocked_cells=True, password=None)` controls selection and enables cell
protection. Use `Style(protection=Protection(locked=False))` for editable cells.
`WorkbookProtection(lock_structure=True, password=None)` protects workbook
structure. The updater exposes `set_sheet_protection(sheet_name, protection)`
and `set_workbook_protection(protection)`; pass `enabled=False` or
`lock_structure=False` to disable protection. Writer accepts `sheet_protection`
on its sheet/DataFrame methods and `set_workbook_protection()` for the workbook.
Passwords use Excel's legacy hash, accept up to 15 ASCII characters, and do not
encrypt the file.

These additions support XLSX/XLSB creation and XLSX/XLSM/XLSB updates without
additional runtime dependencies. See the runnable
[`comments and templates example`](../../examples/xlsx_comments_and_templates.py)
(`--out reviewed.xlsb` selects binary output).

### XLSB formulas and the bulk data path

`Formula` in XLSB supports arithmetic, comparisons, concatenation, constants,
cell/range references (including absolute, whole-row/column and sheet references),
defined names, legacy array formulas and the built-in functions in the BIFF12
function table. Supported newer functions such as `CONCAT` and `XLOOKUP` use
hidden `_xlfn.*` names. Excel performs calculation; a missing `Formula.value`
is initially stored as a numeric zero placeholder in XLSB.

`HYPERLINK` is one of those built-in functions, so a binary workbook can also get
a clickable link from a formula, `Formula('HYPERLINK("#Other!A1","Go")')`. A
`Hyperlink` cell works in XLSB as well. See [Hyperlinks](#hyperlinks) for the
differences and an example.

Structured table references, external-workbook references, inline array
constants and unsupported BIFF12 formula tokens raise an explicit error.
`LET` and the `LAMBDA` family remain unsupported. These limits also apply
when reading formula text; reading cached values does not parse formula tokens.
Point edits into grouped formulas are refused. Sparse edits discard the
modified sheet's obsolete binary index and request recalculation.

Ordinary exports and reads with `data_only=True` retain their native row encoder
and reader, with the existing Python fallback. They do not parse formula
tokens or look up comments for each cell. The binary formula context is
created only when formulas or names are used. Incremental binary reading
is available through `get_rows(..., streaming=True)`; `data_only=False`
selects the formula-aware incremental reader automatically.
