---
description: Switch an xlspy report from XLSX to XLSB without changing the rest of the code.
---

# XLSB in one line

XLSB is Excel's binary workbook format. It reads and writes faster than XLSX and
produces smaller files, which makes it a good default for large exports.

## The same report as XLSB

Change the writer class and the file extension. Nothing else changes.

<!-- runnable -->
```python
from xlspy import F, Style, XlsbWriter

rows = [
    ["Product", "Units", "Revenue"],
    ["Coffee", 120, (12500, F.THOUSANDS_SEP)],
    ["Tea", 90, (8400, F.THOUSANDS_SEP)],
    ["Juice", 45, (5100, F.THOUSANDS_SEP)],
]

with XlsbWriter("sales.xlsb") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet(
        rows,
        header_style=Style(bold=True, fill="#183C52", font_color="#FFFFFF"),
        freeze_panes="A2",
    )

print("wrote sales.xlsb")
```

## When to use XLSB

| Use XLSB when | Use XLSX when |
| :--- | :--- |
| You export large datasets | You exchange files with many tools |
| You control the readers | A web service or library reads the file |
| File size and speed matter | Maximum compatibility matters |
| You want the native fast path | You need the widest ecosystem support |

Both writers support the same high-level features: styles, number formats,
formulas, tables, charts, validation, images, print settings and pivots.

!!! note "XLSB formula grammar"
    XLSB formulas use a binary grammar. The common functions and operators are
    supported; a formula the writer cannot encode raises an error instead of
    writing a broken file. If you need an unusual formula, write XLSX or compute
    the value in Python. See
    [XLSB formulas](../guide/objects.md#xlsb-formulas-and-the-bulk-data-path).

## Read it back

<!-- runnable -->
```python
from xlspy import ExcelReader, XlsbWriter

with XlsbWriter("sales.xlsb") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet([["Product", "Revenue"], ["Coffee", 12500]])

with ExcelReader("sales.xlsb") as reader:
    for row in reader.get_rows("Sales"):
        print(row)
```

`ExcelReader` detects the format from the file, so the same code reads XLSX and
XLSB.

[Next: Styles and number formats →](styles.md){ .md-button .md-button--primary }
