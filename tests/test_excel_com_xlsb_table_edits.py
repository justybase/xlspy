"""Excel reads the tables of a loaded XLSB that xlspy changed, and keeps what xlspy did not touch."""

import pytest

from xlspy import ExcelTable, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50
XL_TOTALS_SUM = 1


def _excel_file(excel, path):
    """Two tables written by Excel; the first has a totals row, a calculated column and no filter buttons."""
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("A1:C1").Value = ("name", "qty", "other")
        for row in range(2, 7):
            sheet.Range(f"A{row}:C{row}").Value = (f"r{row}", row, row * 2)
        first = sheet.ListObjects.Add(1, sheet.Range("A1:C6"), None, 1)
        first.Name = "First"
        first.ShowAutoFilterDropDown = False
        first.ListColumns(3).DataBodyRange.Formula = "=[@qty]*2"
        first.ShowTotals = True
        first.ListColumns(2).TotalsCalculation = XL_TOTALS_SUM
        sheet.Range("E1:F1").Value = ("code", "n")
        for row in range(2, 5):
            sheet.Range(f"E{row}:F{row}").Value = (f"c{row}", row)
        second = sheet.ListObjects.Add(1, sheet.Range("E1:F4"), None, 1)
        second.Name = "Second"
        book.SaveAs(str(path), XL_XLSB)
    finally:
        book.Close(False)


def test_excel_opens_the_edited_tables_and_keeps_their_totals_and_formulas(tmp_path, excel, open_workbook):
    source = tmp_path / "excel.xlsb"
    _excel_file(excel, source)
    workbook = load_workbook(source)
    sheet = workbook.active
    assert [t.displayName for t in sheet.tables] == ["First", "Second"]
    sheet.tables["Second"].ref = "E1:F6"
    sheet["E5"], sheet["F5"], sheet["E6"], sheet["F6"] = "c5", 5, "c6", 6
    sheet.tables["Second"].tableStyleInfo.name = "TableStyleDark3"
    sheet.tables["First"].displayName = sheet.tables["First"].name = "Sales"
    sheet.add_table(ExcelTable("Added", "H1:I2"))
    sheet["H1"], sheet["I1"], sheet["H2"], sheet["I2"] = "k", "v", "x", 1
    out = tmp_path / "out.xlsb"
    workbook.save(out)
    workbook.close()

    book = open_workbook(out)
    data = book.Worksheets("Data")
    assert data.ListObjects.Count == 3
    sales = data.ListObjects("Sales")
    assert sales.Range.Address == "$A$1:$C$7"
    assert sales.ShowTotals and sales.ListColumns(2).TotalsCalculation == XL_TOTALS_SUM
    assert sales.ListColumns(3).DataBodyRange.Cells(1, 1).Formula == "=[@qty]*2"
    assert not sales.ShowAutoFilterDropDown  # a record the edit did not touch
    second = data.ListObjects("Second")
    assert second.Range.Address == "$E$1:$F$6"
    assert second.TableStyle.Name == "TableStyleDark3"
    assert data.ListObjects("Added").Range.Address == "$H$1:$I$2"


def test_excel_opens_a_file_whose_tables_were_removed(tmp_path, excel, open_workbook):
    source = tmp_path / "excel.xlsb"
    _excel_file(excel, source)
    workbook = load_workbook(source)
    workbook.active.tables.remove("Second")
    out = tmp_path / "out.xlsb"
    workbook.save(out)
    workbook.close()
    data = open_workbook(out).Worksheets("Data")
    assert data.ListObjects.Count == 1 and data.ListObjects(1).Name == "First"
    assert data.Range("A2").Value == "r2"
    assert data.ListObjects(1).ShowTotals  # the formulas of the table still calculate
    assert data.Range("B7").Value == sum(range(2, 7))
