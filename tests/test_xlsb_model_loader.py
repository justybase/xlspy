"""The C loader of XLSB sheet data builds the same cells as the Python one."""

import datetime

import pytest

from xlspy import Font, Workbook, load_workbook
from xlspy.workbook import _xlsb_reader

pytestmark = pytest.mark.skipif(
    _xlsb_reader.xlsb_load_cells is None, reason="native extension not built",
)


def snapshot(workbook):
    return {
        sheet.title: sorted(
            (cell.row, cell.column, repr(cell.value), cell.data_type, cell._style_id)
            for row in sheet._cells.values() for cell in row.values()
        ) + [("max", sheet.max_row, sheet.max_column)]
        for sheet in workbook.worksheets
    }


def both_loaders(path, monkeypatch, **options):
    with_c = load_workbook(path, **options)
    try:
        expected = snapshot(with_c)
    finally:
        with_c.close()
    monkeypatch.setattr(_xlsb_reader, "xlsb_load_cells", None)
    python = load_workbook(path, **options)
    try:
        return expected, snapshot(python)
    finally:
        python.close()


def test_native_and_python_loaders_match(tmp_path, monkeypatch):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["text", 1, 2.5, True, datetime.datetime(2024, 5, 6, 7, 8), None, "=A1&\"x\""])
    sheet.append(["=SUM(B1:C1)", -3, 1e300, False, datetime.date(1900, 1, 1), "", "#N/A"])
    sheet["A4"].font = Font(bold=True)  # a styled blank cell
    sheet["Z10"] = "far"
    sheet.row_dimensions[3].height = 30
    other = workbook.create_sheet("Other")
    other["B2"] = "=Sheet!B1*2"
    path = tmp_path / "types.xlsb"
    workbook.save(path)
    workbook.close()

    for data_only in (False, True):
        expected, actual = both_loaders(path, monkeypatch, data_only=data_only)
        assert actual == expected
        monkeypatch.undo()


def test_loaded_cells_have_openpyxl_data_types(tmp_path):
    workbook = Workbook()
    workbook.active.append(["x", 1, True, "=B1+1"])
    path = tmp_path / "kinds.xlsb"
    workbook.save(path)
    workbook.close()
    loaded = load_workbook(path)
    try:
        assert [cell.data_type for cell in loaded.active[1]] == ["s", "n", "b", "f"]
    finally:
        loaded.close()


@pytest.mark.com
@pytest.mark.com
def test_native_loader_on_excel_written_formulas(tmp_path, monkeypatch):
    client = pytest.importorskip("win32com.client")
    excel = client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Range("A1:A20").Formula = "=ROW()*2"  # a shared formula
        sheet.Range("B1:B3").FormulaArray = "=A1:A3*10"  # an array formula
        sheet.Range("C1").Formula = "=1/0"
        sheet.Range("D1:D5").Value = 7
        path = tmp_path / "excel.xlsb"
        book.SaveAs(str(path), 50)
        book.Close(False)
    finally:
        excel.Quit()
    expected, actual = both_loaders(path, monkeypatch)
    assert actual == expected
