---
description: "Task-focused xlspy guides: choose a format, style reports, add objects, update templates, read workbooks and troubleshoot."
---

# How-to guides

These pages assume you know the basics from the [tutorial](../learn/index.md) and
want to solve a specific task.

<div class="xlspy-cards" markdown>

<div class="xlspy-card" markdown>
### [Choose XLSX or XLSB](choose-format.md)
Compatibility, speed, file size and reader support.
</div>

<div class="xlspy-card" markdown>
### [Writing workbooks](writing.md)
Sheet options, large exports, cell text rules, failures and errors.
</div>

<div class="xlspy-card" markdown>
### [Formatting and layout](styling.md)
Number formats, styles, column widths, headers, tables, conditional formats, groups.
</div>

<div class="xlspy-card" markdown>
### [Formulas, charts and objects](objects.md)
Formulas, validation, links, images, printing, charts, sparklines, slicers, pivots.
</div>

<div class="xlspy-card" markdown>
### [DataFrames](dataframes.md)
`to_xlsx`, `to_xlsb` and `write_dataframe`.
</div>

<div class="xlspy-card" markdown>
### [Updating workbooks](updating.md)
`XlsxUpdater`, `XlsmUpdater` and `XlsbUpdater`.
</div>

<div class="xlspy-card" markdown>
### [Reading workbooks](reading.md)
`ExcelReader` and the read modes.
</div>

<div class="xlspy-card" markdown>
### [Workbook model](model.md)
`load_workbook()` and the openpyxl-like model.
</div>

<div class="xlspy-card" markdown>
### [Troubleshooting](troubleshooting.md)
Common errors and what to do about them.
</div>

</div>

## Pick by goal

| Goal | Start here |
| :--- | :--- |
| Export a query result | [Streaming](writing.md#streaming-large-exports) |
| Make a polished report | [Formatting and layout](styling.md) |
| Add a chart or table | [Formulas, charts and objects](objects.md) |
| Refresh a template | [Updating workbooks](updating.md) |
| Read values fast | [Reading workbooks](reading.md) |
| Edit an existing file | [Workbook model](model.md) |
| Use pandas or polars | [DataFrames](dataframes.md) |
