---
description: "Refresh data in XLSX, XLSM and XLSB templates while preserving styles, sheets and macros."
---

# Updating existing workbooks

[Guide](../usage.md) · [Home](../index.md) · [Examples](https://github.com/justybase/xlspy/blob/main/examples/README.md)

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

# An updater is a context manager: leaving the block releases its files.
# Changes are kept only by a save() made inside the block.
with XlsxUpdater("template.xlsx") as updater:
    updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
    updater.save("result.xlsx")
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

`replace_sheet_data()` and its streaming variant keep each table's header row
valid. If a header cell would no longer hold the table's column name (compared
without regard to case), or the new data does not reach the header row, the call
raises `UnsupportedFeatureError` and writes nothing: Excel repairs such a table.
The table's range is not changed. Formula cells are written as formulas in XLSX,
with `Formula.value` as the cached result. In XLSB, `replace_sheet_data()` raises
`UnsupportedFeatureError` for a formula; use `update_cells()` for formulas there.

Besides values, both updaters can replace sheet options. `update_cells()` takes
`styles={address: XF index}`; an address given only there keeps its value (or
becomes a styled blank). `set_merged_cells()` replaces the merged ranges and
`set_data_validations()` replaces the validation rules; pass an empty iterable
to remove them. For example:

```python
from xlspy import DataValidation

updater.update_cells("Data", {"B2": 42}, styles={"B2": 7})
updater.set_merged_cells("Data", ["A1:C1"])
updater.set_data_validations(
    "Data", [DataValidation("list", "B2:B20", formula1='"New,Active,Closed"')]
)
```

Both updaters also expose the workbook's sheet metadata. `get_sheet_states()`
returns `"visible"`, `"hidden"` or `"veryHidden"` by name; `get_active_sheet()`
and `set_active_sheet()` read and set the tab Excel opens; `set_sheet_state()`
changes visibility; and `reorder_sheets(["First", "Second", ...])` moves the
tabs. `add_sheet(name, index=None, *, state="visible")` creates an empty
worksheet and returns its part path, while `remove_sheet(name)` and
`rename_sheet(old, new)` update the package and every reachable reference.
Chartsheets count as sheets. Reordering carries names local to a sheet and the
active tab with it, and activates the first visible sheet when the active one
becomes hidden.

```python
updater.set_sheet_state("Notes", "hidden")
updater.add_sheet("Scratch", 1)
updater.rename_sheet("Data", "Current")
updater.reorder_sheets(["Summary", "Current", "Notes"])
updater.set_active_sheet(0)
```

Removing a sheet is refused when it is the last one or holds pivot tables or
table slicers; those caches are workbook-level and rebuilding them is outside
the updater's scope. Removing a worksheet also removes the comments, drawings,
charts and tables that only it referred to.

When a pivot cache is present, keep the source column schema (column order and
count) unchanged; changing the schema requires rebuilding the pivot cache and
is outside the updater's scope.

On Windows, the generated workbook can be checked with the installed Excel
COM server (the script opens files read-only using Excel's normal loader):

```powershell
powershell -ExecutionPolicy Bypass -File tools\validate_excel_com.ps1 result.xlsx result.xlsb
```
