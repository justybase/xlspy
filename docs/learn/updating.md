---
description: Replace a workbook's data region while keeping styles, sheets, drawings and macros.
---

# Updating a template

Updaters replace the data region of an existing workbook. They keep the workbook
structure, styles, drawings and unrelated sheets. This is the workflow for
templates with pivot tables, formatting, formulas and VBA.

| File | Class |
| :--- | :--- |
| `.xlsx` | `XlsxUpdater` |
| `.xlsm` | `XlsxUpdater` (also exported as `XlsmUpdater`) |
| `.xlsb` | `XlsbUpdater` |

## Replace the data in an XLSX template

<!-- runnable -->
```python
from xlspy import Formula, XlsxUpdater, XlsxWriter

# A stand-in for a real template.
with XlsxWriter("template.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        [["Name", "Amount"], ["Alice", 42], ["Bob", 37]],
        freeze_panes="A2",
    )

updater = XlsxUpdater("template.xlsx")
updater.replace_sheet_data(
    "Data",
    [["Alice", 50], ["Bob", 40]],
    headers=["Name", "Amount"],
)
updater.update_cells("Data", {"B4": Formula("SUM(B2:B3)", value=90)})
updater.save("result.xlsx")

print("wrote result.xlsx")
```

`replace_sheet_data()` replaces the rows below the header, keeps the dominant
style of each existing column and only removes trailing rows that are entirely
empty. Blank rows in the middle are preserved.

## Update individual cells

`update_cells()` is a sparse edit: it patches only the addresses you pass and
leaves everything else, including shared strings and styles, intact.

```python
updater.update_cells(
    "Data",
    {
        "A2": "Alicia",
        "B2": 55,
        "B4": Formula("SUM(B2:B3)", value=95),
        "C2": None,          # clear the value, keep formatting
    },
)
```

Values may be strings, finite numbers, booleans, dates/datetimes, `Formula`, or
`None`. Existing styles and comments stay attached to the same addresses.

## XLSM templates

`XlsmUpdater` is an alias of `XlsxUpdater`. The VBA project is copied
byte-for-byte; xlspy does not parse or rewrite macros.

```python
from xlspy import XlsmUpdater

updater = XlsmUpdater("template.xlsm")
updater.replace_sheet_data("Data", rows, headers=["Name", "Amount"])
updater.save("result.xlsm")
```

## XLSB templates

<!-- runnable -->
```python
from xlspy import XlsbUpdater, XlsbWriter

with XlsbWriter("template.xlsb") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet([["Name", "Amount"], ["Alice", 42], ["Bob", 37]])

updater = XlsbUpdater("template.xlsb")
updater.replace_sheet_data(
    "Data",
    [["Alice", 50], ["Bob", 40]],
    headers=["Name", "Amount"],
)
updater.save("result.xlsb")
```

In XLSB, `replace_sheet_data()` refuses formulas; use `update_cells()` for them.
The same applies to array and dynamic-array formulas in XLSB.

## Streaming a large update

For a database cursor or another one-pass source, use
`replace_sheet_data_stream()`. It writes through bounded-memory temporary files.

```python
updater.replace_sheet_data_stream(
    "Data",
    rows_from_cursor(cursor),
    headers=["Name", "Amount"],
)
updater.save("result.xlsx")
```

The streaming path requires `save()`; `to_bytes()` necessarily builds the whole
archive in memory.

!!! warning "Keep the pivot source schema"
    If a pivot cache reads the sheet, keep the source column order and count
    unchanged. Changing the schema requires rebuilding the pivot cache, which is
    outside the updater's scope.

[Next: Streaming a large export →](streaming.md){ .md-button .md-button--primary }
