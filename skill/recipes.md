# xlspy recipes

Complete, runnable snippets. Each recipe writes into the current directory.

## 1. Styled XLSX report

```python
from xlspy import Alignment, Border, F, Font, PatternFill, Side, Style, XlsxWriter

header = Style(
    font=Font(bold=True, color="FFFFFF"),
    fill=PatternFill("solid", start_color="1F4E78"),
    alignment=Alignment(horizontal="center", vertical="center"),
    border=Border(bottom=Side("thin", "BF9000")),
)
money = Style(number_format=F.THOUSANDS_SEP)
warning = Style(fill="FFF2CC", bold=True)

rows = [
    ["Product", "Status", "Revenue"],
    ["Coffee", "OK", 12500],
    ["Tea", ("Low", warning), 8400],
    ["Juice", "OK", 5100],
]

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet(
        rows,
        header_style=header,
        styles={"B2:B4": Style(horizontal="center"), "C2:C4": money},
        column_widths={"A": 18, "B": 12, "C": 14},
        freeze_panes="A2",
        auto_filter=True,
    )
```

## 2. Export pandas or polars

```python
import pandas as pd

from xlspy import F, Style, XlsxWriter

df = pd.DataFrame(
    {
        "product": ["Coffee", "Tea", "Juice"],
        "revenue": [12500, 8400, 5100],
        "joined": pd.to_datetime(["2024-01-15", "2023-06-01", "2022-11-20"]),
        "note": ["ok", None, "check"],
    }
)

with XlsxWriter("people.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_dataframe(
        df,
        na_rep="n/a",
        formats={"revenue": F.THOUSANDS_SEP, "joined": F.DATE_ISO},
        header_style=Style(bold=True, fill="1F4E78", font_color="FFFFFF"),
        include_index=True,
        freeze_panes="B2",
    )
```

One-liners for a single default sheet:

```python
import pandas as pd

from xlspy import to_xlsb, to_xlsx

df = pd.DataFrame({"name": ["Alice", "Bob"], "age": [25, 30]})
to_xlsx(df, "people.xlsx")
to_xlsb(df, "people.xlsb")
```

## 3. Stream a large XLSB export

```python
from xlspy import XlsbWriter


def rows_from_cursor(cursor):
    while (row := cursor.fetchone()) is not None:
        yield list(row)


def generated_rows(count):
    for i in range(count):
        yield [i, f"row-{i}", i * 1.5]


with XlsbWriter("large.xlsb", use_shared_strings=False) as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        generated_rows(1_000_000),
        column_widths={"A": 10, "B": 24, "C": 14},
        freeze_panes="A2",
    )
```

`use_shared_strings=False` prevents the shared-string table from growing with
distinct text. For XLSX, lower `spool_max_bytes` if the temp spool is large.

## 4. Update a template and keep its formatting

```python
from xlspy import Formula, XlsxUpdater, XlsxWriter

# A stand-in for a real template.
with XlsxWriter("template.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        [["Name", "Amount"], ["Alice", 42], ["Bob", 37]],
        freeze_panes="A2",
    )

updater = XlsxUpdater("template.xlsx")
updater.replace_sheet_data(
    "Data",
    [["Alice", 50], ["Bob", 40]],
    headers=["Name", "Amount"],
)
updater.update_cells(
    "Data",
    {
        "B4": Formula("SUM(B2:B3)", value=90),
        "C2": None,          # clear the value, keep the style
    },
)
updater.save("result.xlsx")
```

For `.xlsm`, use `XlsmUpdater` (an alias of `XlsxUpdater`); the VBA project is
copied byte-for-byte. For `.xlsb`, use `XlsbUpdater`; its
`replace_sheet_data()` refuses formulas, so write formulas with `update_cells()`.

Streaming update:

```python
from xlspy import XlsxUpdater, XlsxWriter

with XlsxWriter("stream-template.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet([["Name", "Amount"]])


def rows():
    yield ["Alice", 50]
    yield ["Bob", 40]


updater = XlsxUpdater("stream-template.xlsx")
updater.replace_sheet_data_stream("Data", rows(), headers=["Name", "Amount"])
updater.save("stream-result.xlsx")
```

## 5. Read values fast

```python
from xlspy import ExcelReader, XlsxWriter

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet([["Product", "Revenue"], ["Coffee", 12500], ["Tea", 8400]])

with ExcelReader("report.xlsx") as reader:
    print(reader.get_sheet_names())

    for row in reader.get_rows("Sales"):
        print(row)

    formulas = reader.read_all("Sales", data_only=False)
    values = reader.read_all("Sales", data_only=True)
```

Use `load_workbook(path, read_only=True)` when you need coordinates or
`iter_rows()` but still want streaming. Use the full `load_workbook(path)` only
when you need styles or edits.

## 6. Edit a workbook with the openpyxl-like model

```python
from xlspy import Font, XlsxWriter, load_workbook

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet([["Product", "Revenue"], ["Coffee", 12500]])

workbook = load_workbook("report.xlsx")
sheet = workbook["Sales"]

print(sheet["A1"].value)
sheet["B2"] = 130
sheet["A1"].font = Font(bold=True)
sheet.freeze_panes = "A2"
sheet.merge_cells("D1:F1")
sheet["D1"] = "Updated"

workbook.save("report-updated.xlsx")
workbook.close()
```

Saving chooses the format from the extension, so `workbook.save("copy.xlsb")`
writes XLSB.

## 7. Table, chart, pivot and slicer

```python
from xlspy import (
    Chart,
    ChartSeries,
    ExcelTable,
    PivotTable,
    PivotValue,
    Slicer,
    Style,
    XlsxWriter,
)

rows = [
    ["Region", "Quarter", "Sales", "Units"],
    ["North", "Q1", 1300, 44],
    ["South", "Q1", 1600, 58],
    ["North", "Q2", 1400, 47],
    ["South", "Q2", 1700, 61],
]
sales = ExcelTable("SalesTable", "A1:D5")

with XlsxWriter("dashboard.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        rows,
        tables=[sales],
        charts=[
            Chart(
                "bar",
                [ChartSeries("C2:C5", categories="A2:A5", title="Sales")],
                title="Sales by region",
                anchor="F2",
            )
        ],
        slicers=[Slicer("SalesTable", "Region", anchor="F20", caption="Region")],
        pivots=[
            PivotTable(
                sales,
                rows="Region",
                columns="Quarter",
                values=[PivotValue("Sales", "sum", number_format="#,##0")],
                location="A3",
                sheet="Summary",
            )
        ],
        header_style=Style(bold=True, fill="1F4E78", font_color="FFFFFF"),
        freeze_panes="A2",
    )
    workbook.add_sheet("Summary")
    workbook.write_sheet([])
```

A pivot target sheet must be added and written after the source sheet. A slicer
must be in the same `write_sheet()` call as its table.

## 8. Validation, notes, links and print setup

```python
import datetime

from xlspy import (
    Comment,
    DataValidation,
    Hyperlink,
    PageMargins,
    PrintSettings,
    XlsxWriter,
)

rows = [
    ["Region", "Amount", "Date", "Link"],
    ["North", 10, datetime.date(2026, 1, 5),
     Hyperlink("https://example.com/report", text="Open")],
]

with XlsxWriter("operations.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        rows,
        data_validations=[
            DataValidation("list", "A2:A50", formula1='"North,South,East"'),
            DataValidation("whole", "B2:B50", formula1=1, formula2=100),
            DataValidation(
                "date", "C2:C50",
                formula1=datetime.date(2026, 1, 1),
                operator="greater_than",
            ),
        ],
        comments={"B2": Comment("Confirm this amount", "Jan", 240, 100)},
        print_settings=PrintSettings(
            orientation="landscape",
            paper_size="A4",
            print_area="A1:D50",
            repeat_rows="1:1",
            fit_to_width=1,
            margins=PageMargins(left=0.4, right=0.4),
            footer="&LReport&RPage &P of &N",
        ),
    )
```

## 9. Conditional formatting and rich text

```python
from xlspy import (
    CellIsRule,
    CellRichText,
    DifferentialStyle,
    Font,
    InlineFont,
    PatternFill,
    TextBlock,
    XlsxWriter,
)

bad = DifferentialStyle(
    font=Font(color="9C0006"),
    fill=PatternFill("solid", start_color="FFC7CE"),
)

rows = [
    ["Metric", "Value"],
    ["Revenue", 12500],
    ["Churn", 0.03],
    [CellRichText("Note: ", TextBlock(InlineFont(b=True), "reviewed"))],
]

with XlsxWriter("quality.xlsx") as workbook:
    workbook.add_sheet("Data")
    workbook.write_sheet(
        rows,
        conditional_formats=[
            CellIsRule("B2:B3", "greaterThan", ["10000"], dxf=bad),
        ],
    )
```

## 10. Verify what you wrote

```python
from xlspy import ExcelReader, XlsxWriter

with XlsxWriter("report.xlsx") as workbook:
    workbook.add_sheet("Sales")
    workbook.write_sheet([["Product", "Revenue"], ["Coffee", 12500]])

with ExcelReader("report.xlsx") as reader:
    sheet = reader.get_sheet_names()[0]
    rows = reader.read_all(sheet)
    assert rows, "the sheet is empty"
    print(sheet, rows[:5])
```

The read-back proves the package is valid. Open the file in Excel or
LibreOffice once to confirm layout, charts and pivots.
