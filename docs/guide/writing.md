# Writing workbooks

[Guide](../usage.md) · [Home](../../README.md) · [Examples](../../examples/README.md)

## Basic Example

```python
from xlspy import XlsbWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.now(datetime.UTC)]
]

# Initialize writer with a specific compression level
with XlsbWriter("output.xlsb", compression_level=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

## XlsxWriter Example

```python
from xlspy import XlsxWriter
import datetime
from decimal import Decimal

data = [
    ["Name", "Age", "City", "info"],
    [-123, 2147483647, 2147483648, 2147483999],
    ["x", "y", "z", datetime.datetime.today()],
    ["Alice", 25, "New York", datetime.date.today()],
    ["Bob", 30, "London", Decimal(3.14)],
    ["Charlie", 35, "Paris", datetime.datetime.now()],
    [True, False, None, datetime.datetime.now(datetime.UTC)]
]

# Initialize writer with a specific compression level
with XlsxWriter("output.xlsx", compression_level=6) as writer:
    # Add a visible sheet
    writer.add_sheet("Visible Sheet")
    writer.write_sheet(data)

    # Add a hidden sheet
    writer.add_sheet("Hidden Sheet", hidden=True)
    writer.write_sheet([["This sheet is hidden."]])
```

## Sheet options

`write_sheet` takes its sheet settings (zoom, freeze panes, tab colour, grid lines,
tables, and so on) in two equivalent ways. Keyword arguments are the short form:

```python
writer.write_sheet(rows, zoom=80, freeze_panes="A2")
```

A `SheetOptions` object is the same settings bundled once and reused:

```python
from xlspy import SheetOptions

report = SheetOptions(zoom=80, freeze_panes="A2", tab_color="#1F4E78")
writer.write_sheet(rows, options=report)
writer.write_sheet(other_rows, options=report, zoom=100)   # a keyword overrides the object
```

`SheetOptions` is frozen. An unknown name raises `TypeError`. `XlsbWriter` does not
accept `ignored_errors`; it raises `TypeError` as well.

## Streaming large exports

`write_sheet` consumes its input exactly once, row by row:

- **XLSX** encodes rows into a temporary spool (kept in memory up to
  `spool_max_bytes`, 1 MiB by default, then on disk). The sheet is assembled once
  the last row is known, so the dimension and autofilter are exact. Writes are
  bounded by the spool setting and the 100-row width sample, not by the row count.
- **XLSB** writes each row to the archive as it is produced.
- **Shared strings** (the default) keep one table of every distinct string in
  memory until save. For high-cardinality exports pass `use_shared_strings=False`:
  text is then stored inline in each cell, so memory no longer grows with the
  number of distinct strings. Files are somewhat larger. The native extension encodes
  inline XLSB rows too; rows it cannot handle (formatted cells, or text it
  refuses) go to the Python encoder (see the benchmarks).

```python
with XlsbWriter("sales.xlsb", use_shared_strings=False) as writer:
    writer.add_sheet("Sales")
    writer.write_sheet(cursor_rows(cursor))      # a generator over a DB cursor
```

Examples: [`xlsx_streaming_export.py`](../../examples/xlsx_streaming_export.py) (DB-API cursor, XLSX or XLSB).

## Cell text

- **Length limit.** Excel stores at most 32,767 characters in a cell, counted in
  UTF-16 code units: a character outside the Basic Multilingual Plane (such as
  most emoji) counts as two, and a combining accent counts as one. Every writer
  path (XLSX and XLSB, shared and inline strings, styled cells, DataFrames,
  hyperlink text and cached formula results) raises `ValueError` for a longer
  value. Nothing is truncated, and the workbook is not published (see
  *Failure handling* below). Excel 16.0 refuses to open XLSX and inline-XLSB
  files whose cells exceed the limit, which is why it is enforced.
- **Invalid Unicode.** Text with a lone surrogate (for example `"\ud800"`) raises
  `UnicodeEncodeError`.
- **Characters XML cannot hold.** In XLSX, control characters, carriage returns
  and U+FFFE/U+FFFF are written in OOXML's `_xHHHH_` form (`"\x07"` becomes
  `_x0007_`), and literal text that already looks like `_xHHHH_` has its
  underscore escaped as `_x005F_`. Excel and xlspy's reader turn these back into
  the original text. Some other readers (openpyxl among them) show the escapes
  as written. XLSB stores text as UTF-16 and needs no escaping.
- **Numbers.** `NaN` and infinite numbers raise `ValueError` in every writer and
  in the updaters. XLSX used to write them as text.

## Failure handling

A workbook reaches its target only after it has been written completely. If the
input raises, a value is refused, or anything fails while saving, the exception
leaves the `with` block and the target is left as it was:

- **A file path** is written to a temporary file in the same directory (named
  `.<name>.<random>.tmp`) and moved into place with `os.replace` when the
  workbook is complete. A failed write creates no file and keeps an existing
  file at that path unchanged. The final step is a rename, not a forced write to
  disk: xlspy does not call `fsync`.
- **A file object** (`io.BytesIO`, an open file, a socket stream) is written to
  a temporary spool, in memory up to `spool_max_bytes` (8 MiB for `XlsbWriter`)
  and on disk beyond that, and copied into the stream on success. A failed write
  leaves the stream untouched. The copy itself cannot be undone: if the stream
  raises part-way through it (a full disk, for example), the error reaches the
  caller and the stream holds only part of the workbook, which must be
  discarded. The workbook is appended at the stream's current position.

After a failure the writer refuses further `write_sheet` and `save` calls.
`save()` writes the workbook parts; the file is published when the writer
closes, so a writer that is closed without a successful save publishes
nothing. The context manager calls both.

## Errors

Every error xlspy raises on purpose derives from `XlspyError`, so
`except XlspyError` catches them all:

- `InvalidFileError` is also a `ValueError` and a `zipfile.BadZipFile`. It is
  raised when the file is not a readable XLSX, XLSM or XLSB package, or when the
  native reader finds data it cannot decode.
- `UnsupportedFeatureError` is also a `NotImplementedError`. It is raised when the
  workbook uses, or the call asks for, a feature xlspy does not implement yet,
  for example a formula in `XlsbUpdater.replace_sheet_data()` (use `update_cells()`).

Other refused arguments raise the built-in `ValueError` or `TypeError`. For
example: a second `write_sheet()` on the same sheet, two tables whose names match
ignoring case, a sheet name with control characters, and a chart type that is not
supported. The refusal happens before the workbook is published.


## Streaming from a database

`write_sheet` takes any iterable of rows, so a generator over a DB-API cursor writes a query result of any size without holding it in memory. [`examples/xlsx_streaming_export.py`](../../examples/xlsx_streaming_export.py) does it with sqlite3; [`examples/netezza_export.py`](../../examples/netezza_export.py) with Netezza (`nzpy-extended`).
