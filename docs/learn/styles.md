---
description: Apply number formats, colors, borders, alignment, column widths and per-cell styles with xlspy.
---

# Styles and number formats

`xlspy` has two styling layers:

- **Flat `Style` fields** such as `bold`, `fill`, `font_color`, `number_format`,
  `horizontal`, `vertical` and `wrap`.
- **openpyxl-style objects** such as `Font`, `PatternFill`, `Border`, `Alignment`
  and `Protection`, for full Excel formatting.

Both work in `XlsxWriter` and `XlsbWriter`.

## A styled report

<!-- runnable -->
```python
from xlspy import (
    Alignment,
    Border,
    F,
    Font,
    PatternFill,
    Side,
    Style,
    XlsxWriter,
)

header = Style(
    font=Font(bold=True, color="FFFFFF"),
    fill=PatternFill("solid", start_color="1F4E78"),
    alignment=Alignment(horizontal="center", vertical="center"),
)
money = Style(number_format=F.THOUSANDS_SEP)
warning = Style(fill="FFF2CC", border=Border(bottom=Side("thin", "BF9000")))

rows = [
    ["Product", "Status", "Revenue"],
    ["Coffee", "OK", 12500],
    ["Tea", ("Low", warning), 8400],
    ["Juice", "OK", 5100],
]

with XlsxWriter("styled.xlsx") as workbook:
    workbook.add_sheet("Report")
    workbook.write_sheet(
        rows,
        header_style=header,
        styles={"B2:B4": Style(horizontal="center"), "C2:C4": money},
        column_widths={"A": 18, "B": 12, "C": 14},
        freeze_panes="A2",
        auto_filter=True,
    )
```

## Number and date formats

Pass a `(value, format_string)` tuple to format one cell, or use the constants in
`xlspy.F`.

| Constant | Excel format |
| :--- | :--- |
| `F.THOUSANDS_SEP` | `#,##0` |
| `F.CURRENCY_EUR` | `#,##0.00 €` |
| `F.PERCENTAGE` | `0%` |
| `F.DATE_ISO` | `yyyy-mm-dd` |
| `F.DATETIME_SHORT` | `dd.mm.yyyy hh:mm` |
| `F.TIME_HH_MM` | `hh:mm` |

You can also pass your own format string:

```python
rows = [["Total", (1234.56, '#,##0.00 "USD"')]]
```

## Styles by range

`styles` takes A1 references. Ranges can be open ended, so you do not need to
know how many rows you will write:

```python
writer.write_sheet(
    rows,
    styles={
        "B:B": Style(horizontal="center"),      # whole column
        "C2:C": Style(number_format="#,##0.00"), # column C, from row 2 down
        "1:1": header,                           # whole first row
    },
)
```

When ranges overlap, later ranges override only the properties they set.

## Per-cell styles

For a style that changes from cell to cell, put a `Style` in the row:

<!-- runnable -->
```python
from xlspy import Style, XlsxWriter

ok = Style(fill="E2EFDA")
late = Style(fill="FCE4D6", bold=True)

rows = [
    ["Order", "Status"],
    [1001, "OK"],
    [1002, ("Late", late)],
    [1003, ("OK", ok)],
]

with XlsxWriter("status.xlsx") as workbook:
    workbook.add_sheet("Orders")
    workbook.write_sheet(rows, freeze_panes="A2")
```

`(value, Style)` cells are faster than building one A1 `styles` entry per cell,
especially for large sheets.

## Column widths and layout

`column_widths` accepts column letters or 0-based indexes. Widths are in Excel's
character units (`0.5`–`255`).

```python
writer.write_sheet(
    rows,
    column_widths={"A": 18, "B": 30, 4: 60},
    freeze_panes="A2",
    zoom=90,
    show_grid_lines=False,
    tab_color="#1F4E78",
)
```

For columns you do not set, xlspy measures the first 1000 rows and approximates
Excel's AutoFit. Change the sample size or add padding:

```python
writer.write_sheet(rows, autofit_rows=5000, autofit_padding=1.5)
```

## Structured styles

When you need a border on one side, a gradient fill, or a specific font, use the
structured objects:

```python
from xlspy import Alignment, Border, Color, Font, GradientFill, Side, Style

style = Style(
    font=Font(name="Arial", size=12, bold=True, color=Color(theme=1)),
    fill=GradientFill("linear", degree=90, stop=["#4472C4", "#FFFFFF"]),
    border=Border(left=Side("thin", "000000"), bottom=Side("double")),
    alignment=Alignment(horizontal="center", wrap_text=True),
    number_format="#,##0.00",
)
```

!!! tip "XLSB limits"
    XLSB cannot store a few openpyxl-only properties. `relative_indent`, the
    `vertical`/`horizontal`/`start`/`end` border sides and `outline=False` raise
    `NotImplementedError` instead of being silently dropped.

[Next: Formulas, tables and charts →](objects.md){ .md-button .md-button--primary }
