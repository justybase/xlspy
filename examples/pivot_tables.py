"""Pivot tables written with XlsxWriter.

A pivot reads a table written in the same ``write_sheet`` call. The report is
rendered from the rows as they stream past and refreshes when Excel opens the file.
"""

import sys

from xlspy import ExcelTable, PivotTable, PivotValue, XlsxWriter

ROWS = [
    ["Region", "Product", "Quarter", "Sales", "Units"],
    ["North", "Coffee", 1, 1200, 40],
    ["North", "Tea", 1, 800, 50],
    ["South", "Coffee", 1, 1500, 55],
    ["South", "Tea", 2, 700, 30],
    ["East", "Coffee", 2, 900, 33],
    ["East", "Tea", 2, 1100, 60],
    ["North", "Coffee", 2, 1300, 44],
    ["South", "Coffee", 2, 1600, 58],
]


def main(path: str = "pivot_tables.xlsx") -> None:
    sales = ExcelTable("SalesTable", "A1:E9")
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            ROWS,
            tables=[sales],
            pivots=[
                # Sales by region and quarter, on the Summary sheet.
                PivotTable(
                    source=sales,
                    rows="Region",
                    columns="Quarter",
                    values=[PivotValue("Sales", "sum", number_format="#,##0")],
                    location="A3",
                    sheet="Summary",
                ),
                # Average units by product, below the data on this sheet.
                PivotTable(
                    source=sales,
                    rows="Product",
                    values=[PivotValue("Units", "average", name="Avg units", number_format="0.0")],
                    location="A12",
                ),
            ],
        )
        writer.add_sheet("Summary")
        writer.write_sheet([])
    print(f"wrote {path}")


if __name__ == "__main__":
    main(*sys.argv[1:])
