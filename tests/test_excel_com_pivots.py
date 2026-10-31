"""Opt-in checks in Microsoft Excel (COM): a pivot table written by XlsxWriter
opens in Excel with the range and the numbers the writer pre-renders. Set
XLSPY_RUN_EXCEL_COM=1 to run.

Labels are not compared: Excel localises them ("Row Labels" becomes
"Etykiety wierszy" in Polish) when it refreshes the pivot on open.
"""

import gc
import io
import os
import time

import openpyxl
import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)

from xlspy import ExcelTable, PivotTable, PivotValue, XlsxWriter  # noqa: E402

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


def _table_pivot(**kwargs):
    return PivotTable(source="SalesTable", location="A12", **kwargs)


CASES = {
    "row_one_value": (
        {"rows": "Region", "values": "Sales"},
        "Data", "A12:B16",
    ),
    "row_several_values": (
        {"rows": "Region", "values": ["Sales", "Units"]},
        "Data", "A12:C16",
    ),
    "row_column_one_value": (
        {"rows": "Region", "columns": "Quarter", "values": "Sales"},
        "Data", "A12:D17",
    ),
    "row_column_several_values": (
        {"rows": "Region", "columns": "Quarter", "values": ["Sales", "Units"]},
        "Data", "A12:G18",
    ),
    # TableRange2 takes in the page field above the report.
    "row_average_with_filter": (
        {"rows": "Region", "values": [PivotValue("Units", "average")], "filters": "Product"},
        "Data", "A10:B16",
    ),
}


@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    yield app
    app.Quit()
    # Collect the COM wrappers while Excel can still answer; left for interpreter
    # shutdown they crash the process.
    del app
    gc.collect()


def _open(excel, path):
    """Open read-only, retrying the calls Excel rejects while it is busy."""
    import pywintypes

    rejected = -2147418111  # RPC_E_CALL_REJECTED
    for attempt in range(10):
        try:
            return excel.Workbooks.Open(str(path), 0, True)
        except pywintypes.com_error as error:
            if error.hresult != rejected or attempt == 9:
                raise
            time.sleep(0.5)


def _write(path, pivot_kwargs):
    out = io.BytesIO()
    with XlsxWriter(out) as w:
        w.add_sheet("Data")
        w.write_sheet(
            ROWS,
            tables=[ExcelTable("SalesTable", "A1:E9")],
            pivots=[_table_pivot(**pivot_kwargs)],
        )
    path.write_bytes(out.getvalue())


@pytest.mark.parametrize("name", sorted(CASES))
def test_excel_lays_out_the_pivot_as_written(tmp_path, excel, name):
    pivot_kwargs, sheet, ref = CASES[name]
    path = tmp_path / f"{name}.xlsx"
    _write(path, pivot_kwargs)

    ours = openpyxl.load_workbook(path)[sheet][ref]
    expected = [[cell.value for cell in row] for row in ours]

    workbook = _open(excel, path)
    try:
        pivot_table = workbook.Worksheets(sheet).PivotTables(1)
        target = pivot_table.TableRange2
        # pywin32 exposes Address as a property ("$A$12:$B$16" with the default arguments).
        assert target.Address.replace("$", "") == ref
        actual = target.Value2
    finally:
        workbook.Close(False)

    for row_index, row in enumerate(expected):
        for column_index, value in enumerate(row):
            excel_value = actual[row_index][column_index]
            where = f"{name}: row {row_index + 1}, column {column_index + 1}"
            if isinstance(value, (int, float)) or isinstance(excel_value, (int, float)):
                assert isinstance(value, (int, float)), f"{where}: Excel shows {excel_value!r}"
                assert isinstance(excel_value, (int, float)), f"{where}: Excel shows {excel_value!r}"
                assert excel_value == pytest.approx(value), where
