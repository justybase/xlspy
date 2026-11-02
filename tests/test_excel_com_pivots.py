"""Opt-in checks in Microsoft Excel (COM): a pivot table written by XlsxWriter
or XlsbWriter opens in Excel, refreshes, and shows the range and the numbers
the writer pre-renders. Set XLSPY_RUN_EXCEL_COM=1 to run.

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
pytestmark = [pytest.mark.com, pytestmark]

from xlspy import (  # noqa: E402
    ExcelTable, PivotTable, PivotValue, XlsbWriter, XlsxWriter, load_workbook,
)

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
    kwargs.setdefault("location", "A12")
    return PivotTable(source="SalesTable", **kwargs)


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
    "row_column_with_filter": (
        {"rows": "Region", "columns": "Quarter", "values": "Sales", "filters": "Product"},
        "Data", "A10:D17",
    ),
    "text_column_field": (
        {"rows": "Region", "columns": "Product", "values": "Sales"},
        "Data", "A12:D17",
    ),
    "count": (
        {"rows": "Region", "values": [PivotValue("Sales", "count")]},
        "Data", "A12:B16",
    ),
    "min": (
        {"rows": "Region", "values": [PivotValue("Sales", "min")]},
        "Data", "A12:B16",
    ),
    "max": (
        {"rows": "Region", "values": [PivotValue("Sales", "max")]},
        "Data", "A12:B16",
    ),
    "other_sheet": (
        {"rows": "Region", "values": "Sales", "location": "A3", "sheet": "Summary"},
        "Summary", "A3:B7",
    ),
    "other_sheet_row_column_several_values": (
        {"rows": "Region", "columns": "Quarter", "values": ["Sales", "Units"],
         "location": "B2", "sheet": "Summary"},
        "Summary", "B2:H8",
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


def _retry(call):
    """Run ``call``, retrying while Excel rejects calls because it is busy."""
    import pywintypes

    rejected = -2147418111  # RPC_E_CALL_REJECTED
    for attempt in range(20):
        try:
            return call()
        except pywintypes.com_error as error:
            if error.hresult != rejected or attempt == 19:
                raise
            time.sleep(0.5)
    return None


def _open(excel, path):
    """Open read-only, retrying the calls Excel rejects while it is busy."""
    return _retry(lambda: excel.Workbooks.Open(str(path), 0, True))


WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


def _write(path, pivot_kwargs, writer, *, summary=False):
    out = io.BytesIO()
    with writer(out) as w:
        w.add_sheet("Data")
        w.write_sheet(
            ROWS,
            tables=[ExcelTable("SalesTable", "A1:E9")],
            pivots=[_table_pivot(**pivot_kwargs)],
        )
        if summary:
            w.add_sheet("Summary")
            w.write_sheet([])
    path.write_bytes(out.getvalue())


def _written_cells(path, sheet, ref):
    """The cells the writer pre-rendered for ``ref``."""
    if path.suffix == ".xlsb":
        book = load_workbook(path)
        try:
            return [[cell.value for cell in row] for row in book[sheet][ref]]
        finally:
            book.close()
    return [[cell.value for cell in row] for row in openpyxl.load_workbook(path)[sheet][ref]]


def _excel_range(excel, path, sheet):
    """The first pivot's ``TableRange2`` address and values after Excel refreshes it.

    The refresh is explicit: when it fails Excel keeps the cells the writer
    rendered, so reading them alone would pass for a pivot Excel cannot build.
    """
    workbook = _open(excel, path)
    try:
        pivot_table = _retry(lambda: workbook.Worksheets(sheet).PivotTables(1))
        _retry(lambda: pivot_table.RefreshTable())
        target = _retry(lambda: pivot_table.TableRange2)
        # pywin32 exposes Address as a property ("$A$12:$B$16" with the default arguments).
        return _retry(lambda: target.Address).replace("$", ""), _retry(lambda: target.Value2)
    finally:
        _retry(lambda: workbook.Close(False))


@pytest.mark.parametrize(("writer", "extension"), WRITERS)
def test_pivot_still_refreshes_after_the_workbook_is_loaded_and_saved(
    tmp_path, excel, writer, extension,
):
    source = tmp_path / f"source.{extension}"
    _write(source, {"rows": "Region", "columns": "Quarter", "values": "Sales"}, writer)
    book = load_workbook(source)
    try:
        book["Data"]["H1"] = "added after loading"
        copy = tmp_path / f"copy.{extension}"
        book.save(copy)
    finally:
        book.close()

    address, actual = _excel_range(excel, copy, "Data")
    assert address == "A12:D17"
    assert actual[-1][-1] == 9100


@pytest.mark.parametrize(("writer", "extension"), WRITERS)
@pytest.mark.parametrize("name", sorted(CASES))
def test_excel_lays_out_the_pivot_as_written(tmp_path, excel, name, writer, extension):
    pivot_kwargs, sheet, ref = CASES[name]
    path = tmp_path / f"{name}.{extension}"
    _write(path, pivot_kwargs, writer, summary=sheet == "Summary")
    expected = _written_cells(path, sheet, ref)

    address, actual = _excel_range(excel, path, sheet)
    assert address == ref

    for row_index, row in enumerate(expected):
        for column_index, value in enumerate(row):
            excel_value = actual[row_index][column_index]
            where = f"{name}: row {row_index + 1}, column {column_index + 1}"
            if isinstance(value, (int, float)) or isinstance(excel_value, (int, float)):
                assert isinstance(value, (int, float)), f"{where}: Excel shows {excel_value!r}"
                assert isinstance(excel_value, (int, float)), f"{where}: Excel shows {excel_value!r}"
                assert excel_value == pytest.approx(value), where
