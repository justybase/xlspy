---
description: "A ten-step xlspy tutorial: install, write, style, chart, read, update and stream Excel workbooks."
---

# Learn xlspy

This tutorial takes you from an empty Python file to a polished, automated Excel
report. Each lesson is short, runnable, and builds on the previous one.

<div class="xlspy-cards" markdown>

<div class="xlspy-card" markdown>
### 1. [Installation](installation.md)
Install xlspy, verify the native extension, and pick the optional extras.
</div>

<div class="xlspy-card" markdown>
### 2. [Your first workbook](first-workbook.md)
Write a styled XLSX report with headers, number formats and frozen panes.
</div>

<div class="xlspy-card" markdown>
### 3. [XLSB in one line](xlsb.md)
Switch the same report to binary XLSB without changing the API.
</div>

<div class="xlspy-card" markdown>
### 4. [Styles and number formats](styles.md)
Apply colors, borders, alignment, column widths and per-cell styles.
</div>

<div class="xlspy-card" markdown>
### 5. [Formulas, tables and charts](objects.md)
Add formulas, an Excel table, data validation and a chart.
</div>

<div class="xlspy-card" markdown>
### 6. [DataFrames](dataframes.md)
Export pandas and polars data, with formats and missing values.
</div>

<div class="xlspy-card" markdown>
### 7. [Reading a workbook](reading.md)
Stream values with `ExcelReader` or load an editable model.
</div>

<div class="xlspy-card" markdown>
### 8. [Updating a template](updating.md)
Replace a template's data while keeping its styles, sheets and macros.
</div>

<div class="xlspy-card" markdown>
### 9. [Streaming a large export](streaming.md)
Write millions of rows from a generator with bounded memory.
</div>

<div class="xlspy-card" markdown>
### 10. [Next steps](next-steps.md)
Where to go after the tutorial: how-to guides, reference and examples.
</div>

</div>

## Before you start

You need Python 3.11 or newer. Microsoft Excel is not required. If you want to
run the DataFrame lesson, install pandas or polars as well.

```bash
python --version
pip install xlspy
```

Every lesson runs from a normal Python script. Commands are shown for a shell;
use PowerShell on Windows if you prefer.

!!! tip "Follow along in a scratch directory"
    Create an empty folder, save each lesson as `lesson.py`, and run it with
    `python lesson.py`. The files the lesson creates appear next to it.

[Start with installation →](installation.md){ .md-button .md-button--primary }
