---
description: Write millions of rows from a generator or database cursor with bounded memory.
---

# Streaming a large export

`write_sheet()` consumes its input exactly once, row by row. You can pass a
generator, a database cursor wrapper, or any iterable of rows.

## A generator

<!-- runnable -->
```python
from xlspy import XlsbWriter

def rows(count: int):
    for i in range(1, count + 1):
        yield [i, f"row-{i}", i * 1.5]

with XlsbWriter("large.xlsb", use_shared_strings=False) as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        rows(1000),
        header_style=None,
        column_widths={"A": 10, "B": 24, "C": 14},
        freeze_panes="A2",
    )

print("wrote large.xlsb")
```

`use_shared_strings=False` stores text inline in each cell. Memory no longer grows
with the number of distinct strings; the file is somewhat larger.

## From a database cursor

```python
def cursor_rows(cursor):
    while True:
        row = cursor.fetchone()
        if row is None:
            break
        yield list(row)

with XlsbWriter("export.xlsb", use_shared_strings=False) as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(cursor_rows(cursor))
```

The same pattern works for XLSX with `XlsxWriter`. XLSX encodes rows into a
temporary spool (1 MiB in memory by default, then disk), so memory is bounded by
the spool setting, not the row count.

## What stays in memory

| Format | Behaviour |
| :--- | :--- |
| XLSB | writes each row to the archive as it is produced |
| XLSX | spools rows, then assembles the sheet once the last row is known |
| Shared strings (default) | one table of every distinct string until save |
| Inline strings (`use_shared_strings=False`) | text in each cell, no growing table |

If you need a smaller XLSX spool, pass `spool_max_bytes`:

```python
with XlsxWriter("export.xlsx", spool_max_bytes=64 * 1024) as workbook:
    ...
```

## Failure handling

A workbook is published only after it has been written completely. If the input
raises or a value is refused, the target file is left as it was. A path is
written to a temporary file and moved into place with `os.replace`; a file object
is written to a temporary spool and copied on success.

!!! warning "Copying into a stream cannot be undone"
    If the stream itself fails while the finished workbook is copied in, the
    stream holds part of the workbook and must be discarded.

[Next: Next steps →](next-steps.md){ .md-button .md-button--primary }
