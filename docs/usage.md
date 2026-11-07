# Using xlspy

[Home](index.md) · [Runnable examples](https://github.com/justybase/xlspy/blob/main/examples/README.md) · [Benchmarks](benchmarks.md)

Read, write, and update Excel workbooks from Python. This guide covers the
APIs and options behind the [first workbook](learn/first-workbook.md).

The guide is split by task:

- [Writing workbooks](guide/writing.md): XlsbWriter and XlsxWriter: sheets, options, large exports, text rules, failures and errors.
- [Formatting and layout](guide/styling.md): Number and date formats, cell styles, column widths, sheet view, headers, conditional formats, tables and row/column groups.
- [Formulas, charts and other objects](guide/objects.md): Formulas, validation, hyperlinks, images, printing, charts, sparklines, slicers, shapes, rich text, document properties, pivot tables, comments and template metadata.
- [pandas and polars DataFrames](guide/dataframes.md): to_xlsx, to_xlsb and write_dataframe.
- [Updating existing workbooks](guide/updating.md): XlsxUpdater, XlsmUpdater and XlsbUpdater.
- [Reading workbooks](guide/reading.md): ExcelReader and the read modes.
- [The workbook model](guide/model.md): load_workbook and Workbook: the openpyxl-like model, loaded XLSX and XLSB workbooks, read-only mode.

The [API reference](api/index.md) lists every public class and function with its arguments.

Run commands and examples from the repository root unless stated otherwise.
