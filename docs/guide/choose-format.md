---
description: Compare XLSX and XLSB in xlspy and choose the right format for reports, exports and pipelines.
---

# Choose XLSX or XLSB

Both formats are first-class in xlspy. The API is the same; the trade-offs are
compatibility, speed and file size.

## Short answer

| If you need… | Choose |
| :--- | :--- |
| Maximum compatibility with other tools | **XLSX** |
| The fastest large export | **XLSB** |
| The smallest file | **XLSB** |
| A file a web service or JavaScript library will read | **XLSX** |
| A file for Excel users on Windows or Mac | Either; XLSB is smaller |
| Macros preserved from a template | Update the existing `.xlsm` with `XlsmUpdater` |

## Detailed comparison

| | XLSX | XLSB |
| :--- | :--- | :--- |
| Format | ZIP of XML parts | ZIP of binary parts |
| Excel support | Excel 2007+ | Excel 2007+ |
| Third-party readers | very wide | narrower |
| Read speed in xlspy | fast | faster |
| Write speed in xlspy | fast | faster |
| File size | larger | smaller |
| Human-inspectable | yes (XML) | no (binary) |
| Formula grammar | OOXML text | BIFF12 binary |
| Best for | interchange, reports | pipelines, large exports |

## The same code

```python
from xlspy import XlsbWriter, XlsxWriter

for writer, path in [(XlsxWriter, "report.xlsx"), (XlsbWriter, "report.xlsb")]:
    with writer(path) as workbook:
        workbook.add_sheet("Data")
        workbook.write_sheet([["Name", "Amount"], ["Alice", 42]])
```

`XlsxWriter` and `XlsbWriter` accept the same rows and most of the same options:
styles, number formats, formulas, tables, charts, validation, images, print
settings, comments and pivots.

## Formula differences

XLSX stores formula text. XLSB stores a binary token stream. The common functions
and operators are supported, but a few constructs are not:

- A lambda called in place is written to XLSX; XLSB needs a `LET` binding.
- External-workbook references and inline array constants are rejected in XLSB.
- An unsupported token raises `UnsupportedFeatureError` instead of writing a
  broken file.

If a formula is the only blocker, write XLSX or calculate the value in Python and
write the result.

## Reading

`ExcelReader` detects the format from the file, so the read code does not change:

```python
from xlspy import ExcelReader

for path in ("report.xlsx", "report.xlsb"):
    with ExcelReader(path) as reader:
        print(path, reader.get_sheet_names())
```

`load_workbook()` reads both formats into the same model.

## Mixed workbooks

A workbook is one format. You cannot put an XLSB sheet inside an XLSX file. If
you need both, write two files from the same rows.

[Back to the how-to overview](index.md){ .md-button }
