"""Pivot tables written with XlsxWriter or XlsbWriter.

A pivot reads a table written in the same ``write_sheet`` call. The report is
rendered from the rows as they stream past and refreshes when Excel opens the file.

    py examples/pivot_tables.py pivot_tables.xlsx
    py examples/pivot_tables.py pivot_tables.xlsb
"""

import sys

from xlspy import ExcelTable, PivotTable, PivotValue, XlsbWriter, XlsxWriter

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
    writer_class = XlsbWriter if path.lower().endswith(".xlsb") else XlsxWriter
    sales = ExcelTable("SalesTable", "A1:E9")
    with writer_class(path) as writer:
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
