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


# Row keywords Excel accepts, each counted by ROWS over a 3-row table (header row + 3 data rows, no totals).
KEYWORD_CASES = [
    ("=ROWS(Tbl[[#Data],[qty]])", 3),
    ("=ROWS(Tbl[#Data])", 3),
    ("=ROWS(Tbl[[#Data],[#Data],[qty]])", 3),
    ("=ROWS(Tbl[[#Headers],[qty]])", 1),
    ("=ROWS(Tbl[[#Headers],[#Data],[qty]])", 4),
    ("=ROWS(Tbl[[#All],[qty]])", 4),
    ("=ROWS(Tbl[#All])", 4),
    ("=ROWS(Tbl[[#All],[#All]])", 4),
    ("=ROWS(Tbl[[#Data],[qty]:[z]])", 3),
]


def test_excel_reads_every_row_keyword_combination_xlspy_writes(tmp_path, excel, open_workbook):
    path = tmp_path / "keywords.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty", "z"], ["a", 1, 2], ["b", 2, 3], ["c", 3, 4]], start=1):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value)
    sheet.add_table(ExcelTable("Tbl", "A1:C4"))
    for index, (formula, _) in enumerate(KEYWORD_CASES, start=1):
        sheet.cell(row=index, column=5, value=formula)
    workbook.save(path)
    workbook.close()

    data = open_workbook(path).Worksheets("Data")
    for index, (formula, expected) in enumerate(KEYWORD_CASES, start=1):
        assert data.Cells(index, 5).Value2 == expected, formula


def test_excel_calculates_a_formula_that_uses_a_table_and_column_renamed_in_the_same_save(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    _sales(source)
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        table = sheet.tables["Tbl"]
        table.tableColumns[1].name = "amount"
        sheet["B1"] = "amount"  # the header cell follows the column name
        table.displayName = table.name = "Sales"
        sheet["E1"] = "=SUM(Sales[amount])*10"
        out = tmp_path / "out.xlsb"
        workbook.save(out)
    finally:
        workbook.close()

    data = open_workbook(out).Worksheets("Data")
    assert data.Range("E1").Value2 == 50
    assert "Sales[amount]" in data.Range("E1").Formula


def test_excel_calculates_a_formula_on_an_earlier_sheet_that_uses_a_table_added_to_a_later_sheet(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    _sales(source)
    workbook = load_workbook(source)
    try:
        calc = workbook["Calc"]
        calc["C1"], calc["D1"], calc["C2"], calc["D2"], calc["C3"], calc["D3"] = "name", "qty", "x", 10, "y", 20
        calc.add_table(ExcelTable("Later", "C1:D3"))
        workbook["Data"]["E2"] = "=SUM(Later[qty])"
        out = tmp_path / "out.xlsb"
        workbook.save(out)
    finally:
        workbook.close()

    opened = open_workbook(out)
    assert opened.Worksheets("Data").Range("E2").Value2 == 30
