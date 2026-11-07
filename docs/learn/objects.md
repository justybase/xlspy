---
description: Add formulas, an Excel table, data validation and a chart to an xlspy report.
---

# Formulas, tables and charts

A report becomes useful when it calculates, filters and explains. This lesson
adds a formula column, an Excel table, a dropdown and a chart.

## The report

<!-- runnable -->
```python
from xlspy import (
    Chart,
    ChartSeries,
    DataValidation,
    ExcelTable,
    Formula,
    Style,
    XlsxWriter,
)

rows = [
    ["Region", "Q1", "Q2", "Total"],
    ["North", 120, 140, Formula("B2+C2")],
    ["South", 90, 110, Formula("B3+C3")],
    ["East", 75, 95, Formula("B4+C4")],
]

sales = ExcelTable("Sales", "A1:D4")

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Report")
    workbook.write_sheet(
        rows,
        header_style=Style(bold=True, fill="#1F4E78", font_color="#FFFFFF"),
        tables=[sales],
        data_validations=[
            DataValidation(
                "list",
                "A2:A4",
                formula1='"North,South,East"',
                error="Choose a region from the list.",
            )
        ],
        charts=[
            Chart(
                "bar",
                [ChartSeries("D2:D4", categories="A2:A4", title_cell="D1")],
                title="Sales by region",
                anchor="F2",
            )
        ],
        freeze_panes="A2",
    )
```

## Formulas

Wrap formula text in `Formula`. A plain string is always text, so exporting
untrusted data cannot inject a formula.

```python
Formula("B2+C2")
Formula("SUM(C2:C4)")
Formula("XLOOKUP(\"Pear\",A2:A4,B2:B4)")
Formula("B2+C2", value=260)   # cached result for readers that do not recalculate
```

- `value=` writes a cached result. Excel recalculates on open anyway.
- `array_ref="C2:C5"` writes a legacy array formula.
- `spill_ref="A1:B3"` writes a dynamic-array formula.
- xlspy does not evaluate formulas. It only checks balanced parentheses and quotes.

!!! warning "XLSB formula limits"
    XLSB uses a binary formula grammar. Common functions and operators are
    supported; an unsupported token raises `UnsupportedFeatureError` instead of
    writing a file Excel cannot open. Write XLSX or compute the value in Python
    when you hit the limit.

## Tables

`ExcelTable(name, ref)` turns a range into an Excel ListObject. The header row
must contain the column names.

```python
sales = ExcelTable("Sales", "A1:D4")
writer.write_sheet(rows, tables=[sales])
```

- Table names must be unique in the workbook.
- A table name that looks like a cell reference (`T1`, `AB12`) is refused.
- `totals_row=True` treats the last row as a totals row.
- `style_name="TableStyleMedium9"` changes the built-in table style.

## Data validation

`DataValidation(type, sqref, ...)` adds a dropdown or a rule:

```python
DataValidation("list", "A2:A4", formula1='"North,South,East"')
DataValidation("whole", "D2:D50", formula1=1, formula2=100)
DataValidation("date", "G2:G50", formula1=datetime.date(2026, 1, 1), operator="greater_than")
```

Invalid combinations raise `ValueError` when the rule is created.

## Charts

`Chart(chart_type, series, ...)` anchors a chart at a cell. The common types are
`"bar"`, `"barh"`, `"line"`, `"area"`, `"pie"`, `"doughnut"`, `"scatter"` and
`"radar"`.

```python
Chart(
    "line",
    [ChartSeries("D2:D4", categories="A2:A4", title="Total")],
    title="Trend",
    anchor="F18",
    width=12,
    height=6,
)
```

Series can have their own color, marker, trendline and error bars. Axes, legends
and data labels are configured with `ChartAxis`, `ChartLegend` and `DataLabels`.

## Hyperlinks

A `Hyperlink` is the cell's value:

```python
from xlspy import Hyperlink

rows = [
    ["Report", Hyperlink("https://example.com/report", text="Open report")],
    ["Sales", Hyperlink("#'Report'!A1", text="Back to top")],
]
```

Allowed targets are `http`, `https`, `mailto`, `ftp` and internal `#Sheet!A1`
locations. Other schemes and relative file paths raise `ValueError`.

[Next: pandas and polars DataFrames →](dataframes.md){ .md-button .md-button--primary }
