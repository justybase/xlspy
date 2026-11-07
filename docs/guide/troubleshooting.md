---
description: Common xlspy errors and limitations, with the fix for each.
---

# Troubleshooting

Start with the exception type. xlspy raises a small, deliberate set of errors.

## `InvalidFileError`

The file is not a readable XLSX, XLSM or XLSB package, or the native reader found
data it cannot decode. `InvalidFileError` is also a `ValueError` and a
`zipfile.BadZipFile`.

Common causes:

- The file is an old `.xls` (BIFF8) workbook, not `.xlsx` or `.xlsb`.
- The file was truncated or renamed with the wrong extension.
- The package is encrypted or password-protected.
- A ZIP entry is corrupt.

Re-save the file in Excel as `.xlsx` or `.xlsb` and try again.

## `UnsupportedFeatureError`

The workbook uses, or the call asks for, a feature xlspy does not implement.
`UnsupportedFeatureError` is also a `NotImplementedError`.

Examples:

- A formula in `XlsbUpdater.replace_sheet_data()`; use `update_cells()`.
- An XLSB conditional-format edit that would invalidate an opaque Excel 2010
  extension.
- An XLSB edit that a loaded workbook cannot save; `XlsbUpdater` may support it,
  or build the workbook with `Workbook()` and save it as `.xlsb`.

The message names the feature. Do not catch and ignore it: the alternative is a
file Excel repairs or silently loses data.

## `ValueError`

The argument is understood but invalid. Typical cases:

- A cell string longer than 32,767 UTF-16 code units.
- `NaN` or infinity in a numeric cell.
- A table header that does not hold the column name.
- A sheet name with control characters.
- A chart option the chart type cannot use.
- A second `write_sheet()` on the same sheet.
- Two tables whose names match ignoring case.

The refusal happens before the workbook is published, so the target file is
unchanged.

## `UnicodeEncodeError`

Text contains a lone surrogate such as `"\ud800"`. Clean the string before
writing; Excel's UTF-16 cannot represent it.

## A formula shows `#NAME?`

A function added after Excel 2007 (`XLOOKUP`, `IFS`, `CONCAT`, `STDEV.S`, …)
needs the `_xlfn.` prefix in the file. xlspy adds it for known future functions.
An unknown function may be written as given; use the invariant English name and
comma separators.

## Excel repairs the file on open

Check these first:

- A table header cell is empty, not text, or differs from the column name.
- A table name looks like a cell reference (`T1`, `AB12`).
- A cell string exceeds 32,767 UTF-16 units.
- A formula was written to XLSB that the binary grammar cannot encode.
- A hyperlink points to a sheet that is not in the workbook.

Rebuild the smallest failing case and inspect the exception; xlspy refuses
invalid input instead of writing a file that needs repair.

## The output file was not created

The writer publishes only after a complete save. If the `with` block raises, the
target is left as it was. Check the original exception, not the missing file.

## The output file is locked

On Windows, Excel keeps a file open exclusively. Close the workbook in Excel
before overwriting it.

## Macros disappeared

Use `XlsmUpdater` or `XlsbUpdater` on the existing macro-enabled file. Creating a
new `.xlsm` from scratch is not supported. Do not save an `.xlsm` as `.xlsx` and
expect the VBA project to remain.

## Strict Open XML

Strict `.xlsx` files are readable through `ExcelReader` and
`load_workbook(..., read_only=True)`. Updating them and `load_workbook()` raise
`NotImplementedError` because edits would mix Transitional markup into a Strict
package. Save the file as a regular `.xlsx` to edit it.

## Still stuck?

Open an issue with the smallest script and the full traceback:
[github.com/justybase/xlspy/issues](https://github.com/justybase/xlspy/issues).

[Back to the how-to overview](index.md){ .md-button }
