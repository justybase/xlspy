# The workbook model

[Guide](../usage.md) · [Home](../../README.md) · [Examples](../../examples/README.md)

## openpyxl-like workbook model

`xlspy.Workbook` and `xlspy.load_workbook()` provide random cell access with an
API modelled on openpyxl. The fast `XlsxWriter`, `XlsbWriter` and `ExcelReader`
paths are unchanged and do not import this model; use it when you need
coordinates, styles, merges or in-place edits rather than bulk throughput.

```python
from xlspy import Comment, Font, Workbook, load_workbook

workbook = Workbook()
sheet = workbook.active
sheet.title = "Data"
sheet["A1"] = "Amount"
sheet["A2"] = 120
sheet["B2"] = "=A2*2"            # stored as a formula
sheet["A1"].font = Font(bold=True)
sheet["A2"].comment = Comment("Confirm", "Ann")
sheet.merge_cells("D1:E2")
sheet.column_dimensions["A"].width = 18
sheet.freeze_panes = "A2"
workbook.save("report.xlsx")      # .xlsb uses XlsbWriter

workbook = load_workbook("report.xlsx")
workbook["Data"]["A2"] = 125
workbook.save("report.xlsx")
workbook.close()
```

Supported model features:

- `Workbook`: `active`, `sheetnames`, `worksheets`, `create_sheet(title,
  index)`, `remove()`, `copy_worksheet()`, `properties`, `defined_names`
  (read), `save()` and `close()`. `Workbook(write_only=True)` creates
  append-only sheets.
- `Worksheet`: `cell()`, `ws["A1"]`, ranges such as `ws["A1:B3"]`, `append()`,
  `iter_rows()`/`iter_cols()` (with `values_only`), `rows`, `values`,
  `merge_cells()`/`unmerge_cells()`, `column_dimensions`/`row_dimensions`
  (width/height, hidden, outline level), `freeze_panes`, `auto_filter.ref`,
  `sheet_state`, `sheet_properties.tabColor` and `sheet_view` options
  (a new XLSX/XLSB workbook saves `sheet_view.showGridLines`; a loaded one
  also keeps the other view settings).
- `Cell`: `value` (numbers, text, booleans, dates/times, `Formula` or text
  starting with `=`, `CellRichText`), `font`, `fill`, `border`, `alignment`,
  `protection`, `number_format`, `comment` and `hyperlink`.
- Rows and columns: `insert_rows()`, `delete_rows()`, `insert_cols()`,
  `delete_cols()` and `move_range(range, rows, cols, translate=False)`, also
  in loaded workbooks.
- Objects: `add_chart()`, `add_image()`, `add_table()`,
  `add_data_validation()` and `conditional_formatting.add(range, rule)`, in
  new and loaded workbooks.
- Printing (new and loaded XLSX and XLSB workbooks): `print_title_rows`, `print_title_cols`,
  `print_area`, `page_setup`, `page_margins`, `print_options`,
  `sheet_properties.pageSetUpPr.fitToPage`, `oddHeader`/`oddFooter`, and
  `sheet_view.zoomScale` (XLSX and XLSB).

#### Moving from openpyxl

Most report scripts only change their imports. The differences:

| openpyxl | xlspy |
|---|---|
| `from openpyxl.chart import BarChart` ... `ws.add_chart(chart, "E2")` | `ws.add_chart(Chart("bar", [ChartSeries("B2:B9")]), "E2")`: xlspy's own chart, image, table, validation and rule classes; passing openpyxl objects raises `TypeError` naming the xlspy class. |
| `CellIsRule(operator=..., formula=...)` | xlspy rules take the range first: `CellIsRule("B2:B9", "greaterThan", ["10"], dxf=...)`; `conditional_formatting.add(range, rule)` replaces it with its own range. |
| `insert_rows()` keeps row heights in place | Heights and widths move with their rows and columns, as in Excel. Cell contents, styles, comments and hyperlinks move as in openpyxl; merged ranges and formulas do not (`move_range(translate=True)` shifts the references of moved formulas). |
| `CellRichText`, `TextBlock`, `InlineFont` from `openpyxl.cell.rich_text` / `.text` | The same names from `xlspy`; `load_workbook(..., rich_text=True)` returns them. |
| Named styles | Register with `add_named_style()` and assign by name with `cell.style = "Name"`; saved in new and loaded XLSX/XLSB workbooks. |

The first row of a model sheet is written as stored: unlike `XlsxWriter`, the
model does not make first-row text bold.

Saving a loaded XLSX/XLSM workbook patches only what changed. Changed cells,
styles, notes, sheet metadata, added/removed/renamed/reordered sheets and
`styles.xml` are rewritten; drawings, charts, pivots, VBA projects and other
parts are copied unchanged. When sheets are removed or reordered,
worksheet-local defined names (such as print areas) keep pointing to the same
sheet, names local to a removed sheet are dropped, and the calculation chain is
removed so that Excel rebuilds it. Rewritten workbook and style parts keep
Excel's namespace prefixes, which `mc:Ignorable` refers to.

Charts, images, tables, data validation, conditional formats and print setup
added to a loaded sheet are saved next to the file's own objects, which stay
as they are: a new rule comes after the sheet's rules (its priority follows
theirs), a new picture joins the sheet's drawing, a new table gets a free id
and must not share its name with another table in the workbook or overlap a
table on its sheet (`ValueError`). Print settings change only what was set:
setting the orientation keeps the sheet's paper size, and setting the footer
keeps its header. The file's print setup (print area and titles, page setup,
fit to page, centring, header and footer) is loaded into the model, so
clearing a setting (`ws.print_area = None`, default margins, an empty footer)
removes it from the file. Each object is saved once; saving again writes only
the objects added since.

Hyperlinks set, removed or moved (by `insert_rows`, `delete_rows` and
`move_range`) on a loaded sheet are saved; the sheet's other links stay as
they are. Links that `Hyperlink` cannot represent, such as relative file
paths, are kept in the file but do not move with their cells.

Like openpyxl, `load_workbook()` returns formula text by default (`data_only=False`).

#### Loaded XLSX and XLSB workbooks

`load_workbook()` reads XLSB into the same model as XLSX. Saving a loaded XLSB
patches the binary parts in place, as `XlsxUpdater` does for XLSX: only the
records of what changed are re-encoded, and everything else keeps its bytes. An
edit that a loaded XLSB cannot save raises `UnsupportedFeatureError` (a
`NotImplementedError`) naming it when you call `save()`; nothing is dropped
silently.

| Feature | XLSX read | XLSX save | XLSB read | XLSB save |
|---|---|---|---|---|
| Values, formulas, comments, inserted and deleted rows | yes | yes | yes | yes |
| Cell styles (font, fill, border, alignment, protection, number format) | yes | yes | yes | yes ¹ |
| Merged cells, hyperlinks | yes | yes | yes | yes |
| Data validation, defined names | yes | yes | yes | yes |
| Column widths, row heights, hidden rows and columns, outline | yes | yes | yes | yes |
| Freeze panes, sheet view (grid lines, zoom), tab colour | yes | yes | yes | yes ² |
| Sheet protection | yes | yes | yes | yes ³ |
| Print setup (area, titles, page setup, margins, header/footer, breaks) | yes | yes | yes | yes |
| Conditional formatting | yes | yes | yes ⁴ | yes ⁴ |
| Tables | yes | yes | yes | yes ⁵ |
| Charts, images | yes | add, edit ⁶ | yes | add, remove |
| Auto filter range | yes | yes | yes | yes ⁷ |
| Document properties | yes | yes | yes | yes |
| Add, remove, rename and reorder sheets | yes | yes | yes | yes |

¹ Styles that XLSB cannot store are refused as they are for new XLSB files:
`Alignment.relative_indent` and the `vertical`, `horizontal`, `start`, `end`
and `outline=False` border settings raise `NotImplementedError`; font names
longer than 31 characters and number formats longer than 255 raise
`ValueError`. The existing style sheet is extended, never reordered, and is
copied byte for byte when no style changed.

² A sheet saved with several window views cannot have its view edited.

³ Protection whose password the file stores as a hash (written by Excel)
cannot have its password changed; in XLSB its permission flags can be changed
and the hash is kept, so Excel still accepts the old password. A password set
through xlspy can be changed as before.

⁴ Every rule type the writers accept is loaded into the model, including colour
scales, data bars and icon sets, in both XLSX and XLSB. In XLSB, an unchanged
rule keeps its Excel 2010 extension (negative bar colours, axes, gradient
settings). Changing only its priority also keeps the extension. Changing a
field represented by the opaque extension, or changing the rule's range, raises
`UnsupportedFeatureError` rather than silently dropping those settings. Removing
the rule removes its extension.

⁵ Tables can be added, removed and changed in a loaded XLSB: the range, the name,
the style (name and banding), the filter range and the column names. Only the
records of what changed are re-encoded, so the rest of the table's part (totals,
column formulas, formats, filter criteria) and the parts of the tables you did not
touch keep their bytes. Not saved: a table with another number of columns (the
range's width must stay) raises `UnsupportedFeatureError`, and so does renaming or
removing a table in a workbook that holds pivot tables or slicers, which would be
left pointing at the old table, and removing a table that a formula refers to. A
binary workbook names a table in a formula by its id, so renaming one is safe. A new
formula can refer to a table added in the same save, as well as to the tables already there;
a table or column renamed in the same save is found by its new name.

⁶ The title, axis titles, legend position, series titles and colours of a
loaded chart can be changed; the rest of its XML is kept.

⁷ A custom auto filter (criteria other than a value list) is kept when the
sheet is saved unchanged; changing it raises `UnsupportedFeatureError`, and
removing it is allowed. A filter's sort state is saved in loaded and new
workbooks: the range, direction, case sensitivity, sort method and
value/cell-colour/font-colour/icon conditions (icon needs `iconSet` and `iconId`;
`3Stars`/`3Triangles`/`5Boxes` use `BrtBeginSortCond14`). A worksheet-level sort
state (Data > Sort without a filter) is not part of the model.

Where a loaded XLSB refuses an edit, `XlsbUpdater` may support it, or the
workbook can be built with `Workbook()` and saved as `.xlsb`.

### Read-only mode

`load_workbook(path, read_only=True)` streams XLSX, XLSM and XLSB worksheets
without building the cell model, following openpyxl's read-only semantics:

- Rows and columns use worksheet coordinates. `iter_rows()` starts at row 1 and
  column 1 unless `min_row`/`min_col` are given; rows missing from the file are
  yielded as empty rows.
- Bounds come from the stored worksheet dimension (`min_row`, `min_column`,
  `max_row`, `max_column`, `dimensions`). Without one, `max_row` and
  `max_column` are `None` and each row is as wide as its last cell;
  `reset_dimensions()` ignores a wrong stored dimension and
  `calculate_dimension(force=True)` scans the sheet.
- `iter_rows()` yields `ReadOnlyCell` objects (`value`, `row`, `column`,
  `coordinate`) and an empty-cell filler; `values_only=True` yields values.
  `ws["B3"]`, `ws.cell(row, column)` and ranges scan the sheet, as in openpyxl.
- Read-only sheets cannot be modified or saved.

This differs from `ExcelReader.get_rows()`, which returns compact rows: the
XLSX reader does not synthesize missing rows, and the XLSB reader starts at the
first column and row of the first row that contains data.
