# Reading workbooks

[Guide](../usage.md) · [Home](../../README.md) · [Examples](../../examples/README.md)

## Reading XLSB and XLSX Files

For XLSX/XLSM, `get_rows()` parses worksheet XML incrementally. It does not load
the complete worksheet or retain previously returned rows. Shared strings and
styles are still loaded separately; shared-formula definitions are retained when
reading formula text. `read_all()` explicitly collects all rows in memory.

XLSX and XLSB return rows of the same shape, by every read mode. Each value is at its
own position: `row[2]` is column C, and rows absent from the file are there as empty
rows. The first row with cells fixes the width that later rows are padded to, rows
before it are `[]`, and a wider row keeps all its cells without widening the ones after it.

The default `data_only=True` returns stored values, including formula caches.
`get_rows(sheet_name, data_only=False)` and `read_all(..., data_only=False)` return
formula strings beginning with `=`. Shared formulas are expanded to each cell's
relative references; legacy array formulas appear at their source cell. Formula
text reading is supported for XLSX/XLSM and supported XLSB tokens. Inline rich text is returned as
plain text with all displayed fragments and spaces retained.

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

`get_sheet_names()` lists every sheet in workbook order, chartsheets included. A
chartsheet has no cells, so `get_rows` yields nothing for it;
`get_chartsheet_names()` names them.

### Read modes: which one to use

Every read mode is fast by default; nothing has to be switched on or off. The
modes differ in what they build, so pick the lightest one that does the job:

| Mode | Builds | Styles | Use it for |
|---|---|---|---|
| `ExcelReader(path).get_rows()` | Nothing; rows of values, streamed | No | Bulk data extraction, ETL, DataFrames. Fastest, constant memory. |
| `load_workbook(path, read_only=True)` | Nothing; openpyxl-style read-only sheets | No | openpyxl read-only code: worksheet coordinates, `iter_rows()`, `ReadOnlyCell`. |
| `load_workbook(path)` | The full editable model | Yes, lazily | Random access, formatting, sheet metadata, objects, edits and save. |

The same three modes work for XLSX, XLSM and XLSB. `data_only` selects stored
values (`True`) or formula text (`False`) in all of them; the defaults follow
the API each mirrors (`ExcelReader` returns values, `load_workbook()` returns
formula text, as openpyxl does).

`load_workbook()` keeps only the style index of each cell while it loads. The
`Font`, `Fill`, `Border`, `Alignment`, `Protection` and number format objects
are built the first time a cell's style is read (`cell.font`,
`cell.number_format`, ...), and identical styles share one object. A workbook
whose styles are never inspected therefore loads at nearly the same speed as an
unstyled one; there is no separate "fast" or "styled" flag. When you only need
values, `ExcelReader` is still several times faster than `load_workbook()`,
because it does not create a `Cell` for each value (see
[benchmarks](../benchmarks.md)).

Strict Open XML workbooks (Excel's "Strict Open XML Spreadsheet" format,
ISO/IEC 29500 Strict namespaces) are readable through `ExcelReader`, including
values, formula text, notes and defined names, and through
`load_workbook(..., read_only=True)`. Updating them is not supported:
`XlsxUpdater` and `load_workbook()` raise `NotImplementedError`, because edits
would mix Transitional markup into a Strict package. Save the file as a
regular `.xlsx` to edit it.
