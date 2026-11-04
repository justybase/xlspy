"""Excel reads the structured references xlspy writes into an XLSB and calculates them."""

import pytest

from xlspy import ExcelTable, Workbook, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50


def _sales(path):
    """Table Tbl (A1:B3) on Data; D1 sums its qty, C2 doubles the qty of its row; Calc sums it too."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for offset, row in enumerate([["name", "qty"], ["a", 2], ["b", 3]]):
        for column, value in enumerate(row):
            sheet.cell(row=offset + 1, column=column + 1, value=value)
    sheet.add_table(ExcelTable("Tbl", "A1:B3"))
    sheet["D1"] = "=SUM(Tbl[qty])"
    sheet["C2"] = "=Tbl[@qty]*2"
    workbook.create_sheet("Calc")["A1"] = "=SUM(Tbl[qty])+1"
    workbook.save(path)
    workbook.close()


def test_excel_calculates_the_references_xlspy_wrote_to_a_new_workbook(tmp_path, excel, open_workbook):
    path = tmp_path / "new.xlsb"
    _sales(path)
    opened = open_workbook(path)
    data = opened.Worksheets("Data")
    calc = opened.Worksheets("Calc")
    assert data.Range("D1").Value2 == 5
    assert data.Range("C2").Value2 == 4
    assert calc.Range("A1").Value2 == 6
    assert "Tbl[qty]" in data.Range("D1").Formula
    assert "Tbl[qty]" in calc.Range("A1").Formula


def test_excel_opens_a_table_with_a_numeric_totals_row(tmp_path, excel, open_workbook):
    path = tmp_path / "totals.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value)
    sheet.add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
    workbook.save(path)
    workbook.close()

    data = open_workbook(path).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.Range.Address == "$A$1:$B$4"
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$4:$B$4"
    assert data.Range("A4").Value2 == "Total"
    assert data.Range("B4").Value2 == 5


def test_excel_calculates_the_references_xlspy_added_to_a_loaded_workbook(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    _sales(source)
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        sheet["E1"] = "=SUM(Tbl[qty])*10"
        sheet["E2"] = "=Tbl[[#Totals],[qty]]"
        out = tmp_path / "out.xlsb"
        workbook.save(out)
    finally:
        workbook.close()

    opened = open_workbook(out)
    data = opened.Worksheets("Data")
    assert data.Range("E1").Value2 == 50
    assert data.Range("D1").Value2 == 5
    assert data.Range("C2").Value2 == 4
    # Excel has no totals row here, so the reference is an error, not a value.
    assert data.Range("E2").Text.startswith("#")


def test_excel_refuses_a_defined_name_that_refers_to_a_structured_reference(excel):
    com_error = pytest.importorskip("pywintypes").com_error
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Range("A1:B3").Value = (("name", "qty"), ("a", 2), ("b", 3))
        sheet.ListObjects.Add(1, sheet.Range("A1:B3"), None, 1).Name = "Tbl"
        # The same call with a plain range is accepted, so the refusal is about the structured reference.
        book.Names.Add("Plain", f"='{sheet.Name}'!$B$2:$B$3")
        with pytest.raises(com_error):
            book.Names.Add("Qty", "=Tbl[qty]")
    finally:
        book.Close(False)


def test_excel_calculates_a_formula_that_refers_to_a_table_added_in_the_same_loaded_session(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    _sales(source)
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        sheet["F1"], sheet["G1"], sheet["F2"], sheet["G2"] = "name", "qty", "x", 10
        sheet.add_table(ExcelTable("Added", "F1:G2"))
        sheet["H1"] = "=SUM(Added[qty])*2"
        out = tmp_path / "out.xlsb"
        workbook.save(out)
    finally:
        workbook.close()

    opened = open_workbook(out)
    assert opened.Worksheets("Data").Range("H1").Value2 == 20
