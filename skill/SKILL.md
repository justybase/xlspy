---
name: xlspy
description: Read, write and update Excel XLSX, XLSM and XLSB workbooks from Python with xlspy. Use when creating styled reports, exporting pandas or polars DataFrames, streaming large exports, adding formulas, charts, tables, pivots, validation or comments, reading cell values quickly, or refreshing existing templates while preserving styles and VBA macros.
---

# xlspy for AI agents

xlspy is a Python library for Excel workbooks. It writes and reads **XLSX**,
**XLSM** and **XLSB**, has no required runtime dependencies, and does not need
Microsoft Excel installed.

This folder is the compact, agent-oriented variant of the documentation:

- [reference.md](reference.md) — signatures, options and value types.
- [recipes.md](recipes.md) — complete, runnable task recipes.

The published human documentation is at
<https://justybase.github.io/xlspy/>.

## Install and verify

```bash
pip install xlspy
python -c "import xlspy; print(xlspy.__version__)"
```

Optional DataFrame extras:

```bash
pip install "xlspy[pandas]"    # or "xlspy[polars]"
```

Python 3.11+ is required. The wheels include a Rust extension; set
`XLSPY_ACCEL=python` to force the slower pure-Python fallback.

## Choose the right entry point

| Task | Use |
| :--- | :--- |
| Write a new workbook from rows | `XlsxWriter` or `XlsbWriter` + `write_sheet()` |
| Export a pandas/polars DataFrame | `to_xlsx(df, path)` / `to_xlsb(df, path)` |
| Write a DataFrame with styles and extra sheets | `writer.write_dataframe(df, ...)` |
| Stream millions of rows | `write_sheet(generator())`, prefer XLSB |
| Read values only, fastest | `ExcelReader(path).get_rows(sheet)` |
| Read values plus coordinates | `load_workbook(path, read_only=True)` |
| Edit cells, styles or structure | `load_workbook(path)` and `Workbook.save()` |
| Replace a template's data region | `XlsxUpdater` / `XlsmUpdater` / `XlsbUpdater` |
| Keep a template's macros and formatting | the matching updater, never a new writer |

## Choose the format

| Need | Choose |
| :--- | :--- |
| Widest compatibility, web services, third-party readers | **XLSX** |
| Fastest large export, smallest file, Excel-only consumers | **XLSB** |
| Preserve an existing VBA project | Update the existing **XLSM** with `XlsmUpdater` |
| A new macro-enabled workbook from scratch | Not supported — start from a template |

`XlsxWriter` and `XlsbWriter` expose the same row, style, chart, table and
option API. Switching format is normally just the class and the extension.

## Minimal write

```python
from xlspy import F, Style, XlsxWriter

rows = [
    ["Product", "Revenue"],
    ["Coffee", (12500, F.THOUSANDS_SEP)],
    ["Tea", (8400, F.THOUSANDS_SEP)],
]

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet(
        rows,
        header_style=Style(bold=True, fill="183C52", font_color="FFFFFF"),
        freeze_panes="A2",
    )
```

The `with` block stages the package and publishes it on success. If the body
raises, the target file is left untouched. For binary output use `XlsbWriter`
and `report.xlsb`.

## Cell value model

A row is an iterable of cells. A cell is one of:

| Python value | Written as |
| :--- | :--- |
| `str`, `int`, `float`, `bool`, `Decimal` | text, number, boolean |
| `datetime.date`, `datetime.datetime`, `datetime.time`, `datetime.timedelta` | Excel date/time serial |
| `None` | empty cell |
| `Formula("SUM(A2:A4)", value=100)` | formula, optionally with a cached result |
| `Hyperlink("https://…", text="Open")` | linked cell |
| `CellRichText(...)` | rich text runs |
| `(value, "0.00")` | value with a number format |
| `(value, Style(...))` | value with a full cell style |

Rules that prevent most bugs:

- A plain string starting with `=` is **text**, never a formula. This is an
  injection guard. Wrap formula text in `Formula`.
- `Formula` takes the expression **without** a leading `=`.
- `NaN`, `inf` and `-inf` raise `ValueError` in `write_sheet()`. Use `na_rep`
  with DataFrames or replace the value first.
- Excel text is limited to 32,767 UTF-16 code units; lone surrogates raise
  `UnicodeEncodeError`.

## Sheet options at a glance

Pass them as keywords to `write_sheet()` / `write_dataframe()`, or build one
reusable `SheetOptions(...)` and pass `options=`:

| Option | Meaning |
| :--- | :--- |
| `header_style` | `Style` for the first row |
| `styles` | `{ "A1:C1": Style(...), "B:B": ... }`, open-ended ranges allowed |
| `merges` | `["A1:C1", ...]` |
| `tables` | list of `ExcelTable` |
| `charts` | list of `Chart` |
| `pivots` | list of `PivotTable` |
| `data_validations` | list of `DataValidation` |
| `conditional_formats` | list of rule objects |
| `images`, `shapes`, `sparklines`, `slicers` | object lists |
| `comments` | `{ "A2": Comment("text", "Author") }` |
| `column_widths` | `{"A": 18, 3: 30}` or `None` to disable autofit |
| `row_heights` | `{"1": 24}` |
| `hidden_rows`, `hidden_columns` | indexes or letters |
| `row_groups`, `column_groups` | list of `Group(first, last)` |
| `freeze_panes` | `"A2"`, `"B2"` |
| `auto_filter` | `True`, or `auto_filter_ref="A1:C10"` |
| `print_settings` | `PrintSettings(...)` |
| `sheet_protection` | `SheetProtection(enabled=True)` |
| `zoom`, `tab_color`, `show_grid_lines`, `show_row_col_headers` | view settings |
| `ignored_errors`, `sort_state` | advanced Excel metadata |

`write_sheet()` can be called only once per sheet; a second call raises
`ValueError`. Add another sheet with `add_sheet()`.

## Styles

Two layers, freely mixed:

```python
from xlspy import Alignment, Border, F, Font, PatternFill, Side, Style

header = Style(
    font=Font(bold=True, color="FFFFFF"),
    fill=PatternFill("solid", start_color="1F4E78"),
    alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    border=Border(bottom=Side("thin", "BF9000")),
)
money = Style(number_format=F.THOUSANDS_SEP)
```

Flat fields: `fill`, `font_color`, `bold`, `italic`, `underline`,
`number_format`, `horizontal`, `vertical`, `wrap`.

Unset fields are `None` and mean "leave unchanged", which makes styles
composable when a header style and a range style overlap.

Common `F` constants: `THOUSANDS_SEP`, `TWO_DECIMALS`, `CURRENCY_EUR`,
`CURRENCY_PLN`, `PERCENTAGE`, `TEXT`, `DATE_ISO`, `DATE_SHORT`, `DATETIME_SHORT`,
`TIME_HH_MM`, `SCIENTIFIC`, `LEADING_ZEROS`.

## Read modes

| Mode | Builds | Styles | Best for |
| :--- | :--- | :--- | :--- |
| `ExcelReader(path).get_rows(sheet)` | streamed values | no | ETL, bulk extraction |
| `load_workbook(path, read_only=True)` | read-only sheet objects | no | coordinates, `iter_rows()` |
| `load_workbook(path)` | full editable model | lazy | edits, formatting, saving |

`ExcelReader` returns cached values by default (`data_only=True`);
`load_workbook` follows openpyxl and returns formula text by default
(`data_only=False`). Pass the flag explicitly when it matters.

## Update or rebuild

Use an **updater** when the file already exists and you want to keep styles,
sheets, drawings, pivot caches and macros. Use a **writer** when you build a new
report from data.

```python
from xlspy import Formula, XlsxUpdater

updater = XlsxUpdater("template.xlsx")
updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
updater.update_cells("Data", {"B4": Formula("SUM(B2:B3)", value=90)})
updater.save("result.xlsx")
```

`XlsmUpdater` is an alias of `XlsxUpdater`; the VBA project is copied
byte-for-byte. `XlsbUpdater.replace_sheet_data()` refuses formulas — write the
values with it and use `update_cells()` for formulas.

## Objects

| Object | Key arguments |
| :--- | :--- |
| `ExcelTable` | `display_name`, `ref`, `columns`, `totals_row`, `style_name` |
| `PivotTable` | `source`, `rows`, `columns`, `values=[PivotValue(...)]`, `location`, `sheet` |
| `Chart` | `chart_type`, `series=[ChartSeries(...)]`, `title`, `anchor`, axes, legend |
| `DataValidation` | `type`, `sqref`, `formula1`, `formula2`, `operator` |
| `Hyperlink` | `target`, `text`, `tooltip` |
| `Comment` | `text`, `author`, `width`, `height` |
| `Image` | `path`, `anchor`, `width`, `height`, `description` |
| `Shape` | `anchor`, `width`, `height`, `kind`, `text`, `fill` |
| `SparklineGroup` | `data`, `location`, `kind`, `color` |
| `Slicer` | `table`, `column`, `anchor`, `caption` |
| `PrintSettings` | `orientation`, `paper_size`, `print_area`, `repeat_rows`, `fit_to_width` |
| `Group` | `first`, `last`, `collapsed` |

Supported chart types: `bar`, `barh`, `line`, `area`, `pie`, `doughnut`,
`scatter`, `radar`.

## Pitfalls that break files

- A table name may not look like a cell reference (`T1`, `AB12`), and its
  header cells must contain the column names.
- A workbook is one format; you cannot put an XLSB sheet inside an XLSX file.
- XLSB formulas use a binary grammar. External references and inline array
  constants are refused with `UnsupportedFeatureError`; write XLSX instead.
- A hyperlink target must be `http`, `https`, `mailto`, `ftp` or an existing
  internal `#Sheet!A1`. Relative paths and unknown sheets raise `ValueError`.
- With shared strings (the default), memory grows with the number of distinct
  strings. For high-cardinality exports pass `use_shared_strings=False`.
- XLSX rows spool to a temporary file after `spool_max_bytes`; XLSB encodes rows
  as they are produced.
- Loading and then saving a Strict Open XML file raises `NotImplementedError`.
  Read it with `ExcelReader` or `read_only=True` instead.
- XLSB cannot store a few openpyxl-only style properties; those raise
  `NotImplementedError` rather than being silently dropped.
- A failed write publishes nothing, so a missing output file means the body
  raised. Read the original exception.

## Errors

| Error | Meaning |
| :--- | :--- |
| `InvalidFileError` | Not a readable XLSX/XLSM/XLSB package; also a `ValueError` and `BadZipFile` |
| `UnsupportedFeatureError` | The requested feature cannot be represented in this format; also `NotImplementedError` |
| `ValueError` | Invalid argument or value, refused before publishing |
| `TypeError` | Wrong type or unknown keyword argument |
| `UnicodeEncodeError` | Text contains a lone surrogate |

Do not catch `UnsupportedFeatureError` and continue: the alternative is a file
Excel repairs or loses data.

## Verify generated files

At minimum, read the result back:

```python
from xlspy import ExcelReader

with ExcelReader("report.xlsx") as reader:
    print(reader.get_sheet_names())
    print(reader.read_all(reader.get_sheet_names()[0]))
```

For layout, formulas and charts, open the file in Excel or LibreOffice once.
Reading values back proves the package is valid; it does not prove the chart or
pivot renders as intended.
