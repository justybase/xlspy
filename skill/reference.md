# xlspy API reference (compact)

Every public name below is importable from `xlspy` unless noted. Full generated
documentation: <https://justybase.github.io/xlspy/api/>.

## Writers

### `XlsxWriter(file, compression_level=3, use_shared_strings=True, spool_max_bytes=1_048_576)`

Writes an XLSX package. Use as a context manager; the file is published on a
clean exit.

```python
from xlspy import XlsxWriter

with XlsxWriter("report.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_sheet(rows, **sheet_options)
```

- `file`: path or binary file object (`BytesIO`).
- `compression_level`: 0–9.
- `use_shared_strings`: `False` stores text inline, bounding memory for
  high-cardinality data.
- `spool_max_bytes`: rows kept in memory before spilling to disk.

### `XlsbWriter(file, compression_level=3, use_shared_strings=True)`

Same API, binary XLSB output. Rows are encoded as they are produced.

### Writer methods

| Method | Notes |
| :--- | :--- |
| `add_sheet(name, hidden=False)` | Adds a worksheet; writes go to the last added sheet. |
| `write_sheet(rows, *, options=None, **sheet_options)` | Consumes rows once. One call per sheet. |
| `write_dataframe(df, include_index=False, na_rep=None, formats=None, **sheet_options)` | pandas/polars export. polars rejects `include_index=True`. |
| `save()` / context exit | Publishes the package. |

`SheetOptions` bundles every sheet option for reuse:

```python
from xlspy import SheetOptions

opts = SheetOptions(freeze_panes="A2", auto_filter=True, zoom=90)
writer.write_sheet(rows, options=opts)
```

Keyword arguments override fields of `options`.

### Sheet options

`header_style`, `styles`, `merges`, `conditional_formats`, `tables`,
`column_widths`, `autofit_rows` (default 1000; `None` disables), `autofit_padding`,
`hidden_columns`, `row_heights`, `hidden_rows`, `row_groups`, `column_groups`,
`freeze_panes`, `auto_filter`, `auto_filter_ref`, `charts`, `data_validations`,
`images`, `print_settings`, `comments`, `sheet_protection`, `show_grid_lines`,
`header_rows`, `tab_color`, `zoom`, `ignored_errors`, `sparklines`, `slicers`,
`pivots`, `shapes`, `show_row_col_headers`, `show_zero_values`, `right_to_left`,
`sort_state`.

`styles` accepts open-ended A1 ranges: `"B:B"`, `"C2:C"`, `"1:1"`.

## Values

| Type | Constructor |
| :--- | :--- |
| Formula | `Formula(text, value=None, array_ref=None, spill_ref=None)` |
| Array formula | `ArrayFormula(ref, text)` |
| Data table | `DataTableFormula(ref, ca=False, dt2D=False, dtr=False, r1=None, r2=None, del1=False, del2=False)` |
| Hyperlink | `Hyperlink(target, text=None, tooltip=None)` |
| Rich text | `CellRichText(*parts)`, `TextBlock(InlineFont(...), text)` |
| Number format | `(value, "format string")` |
| Cell style | `(value, Style(...))` |

`Formula.text` has **no** leading `=`. `value` is the cached result.
`array_ref` writes a legacy CSE formula; `spill_ref` writes a dynamic array.
`ArrayFormula`/`DataTableFormula` are the openpyxl objects; XLSX takes both,
XLSB takes `ArrayFormula` only.

`InlineFont` uses openpyxl names: `rFont`, `b`, `i`, `u`, `strike`, `sz`,
`color`, `vertAlign`, `family`, `charset`, `scheme`.

## Styles

### `Style(fill=None, font_color=None, bold=None, italic=None, underline=None, number_format=None, horizontal=None, vertical=None, wrap=None, font=None, pattern=None, border=None, alignment=None, protection=None)`

Flat fields are shorthands; structured fields win per property. Unset fields
stay `None`, so styles compose.

### Structured attributes

| Class | Fields |
| :--- | :--- |
| `Font` | `name`, `size`, `bold`, `italic`, `strike`, `underline`, `vertAlign`, `color`, `charset`, `family`, `scheme`, `outline`, `shadow`, `condense`, `extend` |
| `PatternFill` | `fill_type`, `start_color`, `end_color` |
| `GradientFill` | `fill_type`, `degree`, `left`, `right`, `top`, `bottom`, `stop` |
| `Stop` | `color`, `position` |
| `Side` | `style`, `color` |
| `Border` | `left`, `right`, `top`, `bottom`, `diagonal`, `diagonal_direction`, `vertical`, `horizontal`, `start`, `end`, `diagonal_up`, `diagonal_down`, `outline` |
| `Alignment` | `horizontal`, `vertical`, `wrap_text`, `shrink_to_fit`, `indent`, `text_rotation`, `relative_indent`, `justify_last_line`, `reading_order` |
| `Protection` | `locked`, `hidden` |
| `Color` | `rgb`, `indexed`, `theme`, `tint`, `auto` |

Colors accept `"RRGGBB"`, `"#RRGGBB"`, `"AARRGGBB"`, or `Color(...)`.
`DifferentialStyle(font, fill, border, number_format, alignment, protection)`
is the conditional-formatting style; its `fill` must be a `PatternFill` or
`GradientFill`.

### `F` constants

`THOUSANDS_SEP`, `TWO_DECIMALS`, `CURRENCY_EUR`, `CURRENCY_PLN`, `PERCENTAGE`,
`SCIENTIFIC`, `TEXT`, `LEADING_ZEROS`, `DATE_SHORT`, `DATE_LONG`,
`DATE_DAY_MONTH_YEAR`, `DATE_ISO`, `DATE_MONTH_YEAR`, `DATE_WEEKDAY`,
`DATE_DAY_MONTH`, `DATE_YEAR_ONLY`, `DATETIME_SHORT`, `DATETIME_LONG`,
`TIME_HH_MM`, `TIME_HH_MM_SS`, `DATETIME_ISO`, `TIME_MS`, `SHORT_DATE`,
`LONG_DATE`, `ISO_DATE`, `SHORT_DATETIME`, `LONG_DATETIME`, `ISO_DATETIME`.

## Conditional formats

| Rule | Signature |
| :--- | :--- |
| `CellIsRule` | `(sqref, operator, formula, dxf=None, stop_if_true=None, priority=None)` |
| `FormulaRule` | `(sqref, formula, dxf=None, stop_if_true=None, priority=None)` |
| `ColorScaleRule` | `(sqref, minimum=None, midpoint=None, maximum=None, colors=(...), priority=None)` |
| `DataBarRule` | `(sqref, minimum=None, maximum=None, color=..., show_value=None, min_length=None, max_length=None, priority=None)` |
| `IconSetRule` | `(sqref, icon_style, values, cfvo_type="percent", show_value=None, percent=None, reverse=None, stop_if_true=None, priority=None)` |
| `TopBottomRule`, `AboveAverageRule`, `TextRule`, `DuplicateRule`, `TimePeriodRule`, `BlanksRule` | see the API page |

`CellIsRule.operator` is camelCase: `between`, `notBetween`, `equal`,
`notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`,
`lessThanOrEqual`, `containsText`, `notContains`, `beginsWith`, `endsWith`.

## Tables, pivots and objects

| Class | Signature highlights |
| :--- | :--- |
| `ExcelTable` | `(display_name, ref, columns=None, header_row=True, totals_row=False, style_name="TableStyleMedium9", show_first_column=False, show_last_column=False, show_row_stripes=True, show_column_stripes=False, totals_row_functions=None)` |
| `PivotTable` | `(source, rows, values, location="A3", columns=(), filters=(), sheet=None, name=None, style="PivotStyleLight16", data_caption="Values", autofit_columns=False)` |
| `PivotValue` | `(field, function="sum", name=None, number_format=None)`; functions: `sum`, `count`, `average`, `min`, `max` |
| `DataValidation` | `(type, sqref, formula1=None, formula2=None, operator=None, allow_blank=True, show_dropdown=True, show_input_message=False, prompt_title=None, prompt=None, show_error_message=True, error_title=None, error=None, error_style="stop", ime_mode=None)` |
| `Comment` | `(text, author, width=144, height=79)` |
| `Image` | `(path, anchor="A1", width=None, height=None, description=None)` |
| `Shape` | `(anchor, width=200, height=80, kind="rect", text=None, fill=None, line=None, font_size=11, font_color="#000000", bold=False, align="left", valign="top", link_cell=None, name=None)` |
| `SparklineGroup` | `(data, location, kind="line", color=..., markers=False, high_point=False, low_point=False, first_point=False, last_point=False, negative_points=False, ...)` |
| `Slicer` | `(table, column, anchor="H2", width=144, height=180, caption=None, name=None, columns=1)` |
| `Group` | `(first, last, collapsed=False)` |
| `PrintSettings` | `(orientation=None, paper_size=None, margins=None, print_area=None, repeat_rows=None, repeat_columns=None, fit_to_width=None, fit_to_height=None, scale=None, center_horizontally=False, center_vertically=False, header=None, footer=None, row_breaks=None, col_breaks=None)` |
| `PageMargins` | `(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3)` inches |
| `SheetProtection` | `(enabled=True, select_locked_cells=True, select_unlocked_cells=True, password=None, hashed_password=None)` |

`DataValidation.type`: `list`, `whole`, `decimal`, `date`, `time`,
`text_length`, `custom`.
`DataValidation.operator`: `between`, `not_between`, `equal`, `not_equal`,
`greater_than`, `less_than`, `greater_than_or_equal`, `less_than_or_equal`.
`Shape.kind`: `rect`, `round`, `ellipse`, `textbox`.

### Charts

```python
Chart(
    chart_type, series, title=None, anchor="E2", width=15.0, height=7.5,
    style=None, x_axis=None, y_axis=None, y2_axis=None, legend=None,
    data_labels=None, show_data_labels=False, title_size=None, grouping=None,
    gap_width=None, hole_size=None, fill=None, border=None, plot_fill=None,
    font_size=None, font_color=None,
)
```

`ChartSeries(values, title=None, title_cell=None, categories=None,
x_values=None, color=None, marker=None, line=None, chart_type=None,
secondary_axis=False, trendline=None, error_bars=None)`.
`title` and `title_cell` are mutually exclusive.

`ChartAxis(title=None, minimum=None, maximum=None, major_unit=None,
number_format=None, gridlines=None, reverse=False, visible=True)`.
`ChartLegend(position="right", overlay=False)`.
`ChartMarker(symbol="circle", size=7)`.
`DataLabels(number_format=None, position=None, show_category=False,
show_series_name=False)`.
`Trendline(type="linear", order=None, period=None, forward=None,
backward=None, display_equation=False, display_r_squared=False, name=None,
color=None)`.
`ErrorBars(kind=..., value=None, direction=..., axis=..., end_cap=...,
color=None)`.

Chart types: `bar`, `barh`, `line`, `area`, `pie`, `doughnut`, `scatter`,
`radar`. Grouping (`stacked`, `percent_stacked`) applies to bar, barh, line and
area. A series may set `chart_type` to bar/line/area for a combined chart and
`secondary_axis=True` for the right-hand axis.

## DataFrames

| Function | Signature |
| :--- | :--- |
| `to_xlsx(df, path, sheet_name="Sheet1", *, include_index=False, na_rep=None, formats=None, header_style=None, styles=None, merges=None, conditional_formats=None, tables=None, column_widths=None, autofit_rows=1000, autofit_padding=0.0, hidden_columns=None, row_heights=None, hidden_rows=None, row_groups=None, column_groups=None, freeze_panes=None, auto_filter=None, charts=None, comments=None, sheet_protection=None, **writer_kwargs)` | one-liner |
| `to_xlsb(...)` | same, XLSB output |

Both return `path`. pandas supports `include_index` (including MultiIndex);
polars raises on `include_index=True`.

Type mapping: timestamps become date/time cells, timedeltas become Excel
durations, numpy scalars are unwrapped, and `None`/`NaN`/`NaT`/`NA` become empty
cells or `na_rep`.

## Reader

### `ExcelReader(path=None, use_memory_stream_for_xlsb=True, errors="legacy")`

| Method | Returns |
| :--- | :--- |
| `open(path)` | opens a file |
| `get_sheet_names()` / `get_chartsheet_names()` | names |
| `get_sheet_index(name)` | zero-based index |
| `get_rows(sheet, *, data_only=True, streaming=False)` | generator of rows |
| `read_all(sheet, *, data_only=True, streaming=False)` | list of rows |
| `get_comments(sheet)` | `{address: Comment}` |
| `get_defined_names()` | defined names |
| `read_part(path)` | raw package part bytes |
| `close()` | releases the archive |

`errors="legacy"` returns `"error:#DIV/0!"`; `"openpyxl"` returns `"#DIV/0!"`.

## Workbook model

### `load_workbook(filename, *, read_only=False, data_only=False, keep_vba=False, rich_text=False, keep_links=True)`

Returns a `Workbook`. `read_only=True` returns sheets that stream rows and cannot
be edited or saved.

### `Workbook(write_only=False, iso_dates=False)`

| Member | Notes |
| :--- | :--- |
| `active`, `sheetnames`, `worksheets`, `chartsheets` | workbook structure |
| `create_sheet(title=None, index=None)`, `remove(sheet)`, `move_sheet(...)` | edit structure |
| `wb["Sheet"]` / `ws["A1"]` | worksheet and cell access |
| `ws.cell(row, column, value=None)`, `ws.append(iterable)` | write cells |
| `ws.iter_rows(...)`, `ws.iter_cols(...)` | read cells; `values_only=True` |
| `ws.merge_cells(...)`, `unmerge_cells(...)` | merged ranges |
| `ws.insert_rows/cols`, `delete_rows/cols` | structural edits |
| `ws.add_chart`, `add_image`, `add_table`, `add_data_validation` | objects |
| `workbook.save(filename)` | XLSX/XLSM/XLSB chosen by extension |
| `workbook.close()` | release resources |

Cell values use the same types as the writers. Styles are openpyxl-like:
`cell.font`, `cell.fill`, `cell.border`, `cell.alignment`,
`cell.number_format`, `cell.protection`.

## Updaters

`XlsxUpdater(path)` — XLSX. `XlsmUpdater` is an alias that keeps the VBA project.
`XlsbUpdater(path)` — XLSB.

| Method | Notes |
| :--- | :--- |
| `get_sheet_names()` | worksheet names |
| `replace_sheet_data(sheet, rows, *, headers=None, style_fallback="inherit")` | replaces data below the header; keeps column styles. `style_fallback` is `"inherit"` or `"general"` |
| `replace_sheet_data_stream(...)` | same, bounded-memory streaming |
| `update_cells(sheet, {address: value})` | sparse edit; values, `Formula`, `None` |
| `set_comments(sheet, {address: Comment \| None})` | replace/clear notes |
| `get_comments(sheet)` | notes |
| `set_hyperlinks(sheet, {...})` | links |
| `set_defined_name(...)`, `set_defined_names(...)`, `remove_defined_name(...)` | names |
| `set_sheet_protection(...)`, `set_workbook_protection(...)` | protection |
| `save(output_path=None)`, `to_bytes()` | publish |
| `close()` | release resources |

XLSB-only structural helpers: `add_sheet`, `remove_sheet`, `rename_sheet`,
`reorder_sheets`, `set_sheet_state`, `set_active_sheet`, `set_merged_cells`,
`set_data_validations`, `set_conditional_formats`.

XLSB `replace_sheet_data()` rejects formulas; use `update_cells()` for them.
`to_bytes()` necessarily builds the whole archive in memory.

## Errors

| Exception | Also | Raised when |
| :--- | :--- | :--- |
| `XlspyError` | `Exception` | base class |
| `InvalidFileError` | `ValueError`, `zipfile.BadZipFile` | unreadable or corrupt package, old `.xls`, encrypted file |
| `UnsupportedFeatureError` | `NotImplementedError` | feature cannot be represented (e.g. XLSB formula token, Strict Open XML edit) |
| `ValueError` | — | invalid argument or value; refused before publishing |
| `TypeError` | — | wrong type or unknown keyword argument |
| `UnicodeEncodeError` | — | lone surrogate in text |

## Environment

- Python 3.11+.
- `XLSPY_ACCEL=python` forces the pure-Python fallback.
- `xlspy._accel.HAVE_NATIVE` and `xlspy._accel.ACCEL_BACKEND` report the active
  backend.
