"""Excel opens the totals rows xlspy writes in XLSB and XLSX."""

import pytest

from xlspy import ExcelTable, Formula, Workbook, load_workbook

pytestmark = pytest.mark.com

XL_XLSB = 50


def _xlsb_totals(path, value):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", value]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
    workbook.save(path)
    workbook.close()


def _assert_xlsb_totals(path, open_workbook, expected, *, has_formula=False):
    data = open_workbook(path).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.Range.Address == "$A$1:$B$4"
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$4:$B$4"
    assert data.Range("A4").Value2 == "Total"
    assert data.Range("B4").Value2 == expected
    assert bool(data.Range("B4").HasFormula) is has_formula


@pytest.mark.parametrize(
    ("value", "expected", "has_formula"),
    [
        (5, 5, False),
        ("5", "5", False),
        (None, None, False),
        (Formula("SUM(B2:B3)", 5), 5, True),
    ],
)
def test_excel_opens_an_xlsb_totals_row(tmp_path, excel, open_workbook, value, expected, has_formula):
    path = tmp_path / "totals.xlsb"
    _xlsb_totals(path, value)
    _assert_xlsb_totals(path, open_workbook, expected, has_formula=has_formula)


@pytest.mark.parametrize("value", [5, Formula("SUM(B2:B3)", 5)])
def test_excel_saves_and_reopens_an_xlsb_totals_row(tmp_path, excel, open_workbook, value):
    path = tmp_path / "totals.xlsb"
    _xlsb_totals(path, value)
    saved = tmp_path / "saved.xlsb"
    book = excel.Workbooks.Open(str(path), 0, False)
    try:
        book.SaveAs(str(saved), XL_XLSB)
    finally:
        book.Close(False)
    _assert_xlsb_totals(saved, open_workbook, 5, has_formula=isinstance(value, Formula))


def test_excel_opens_an_xlsb_totals_table_added_to_a_loaded_workbook(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    workbook.save(source)
    workbook.close()

    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        workbook["Data"].add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
        workbook.save(out)
    finally:
        workbook.close()
    _assert_xlsb_totals(out, open_workbook, 5)


def test_excel_opens_a_loaded_xlsb_totals_table_after_its_range_moves(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    _xlsb_totals(source, 5)
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        sheet["A5"], sheet["B5"] = "Sum", 10
        sheet.tables["Tbl"].ref = "A1:B5"
        workbook.save(out)
    finally:
        workbook.close()

    data = open_workbook(out).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.Range.Address == "$A$1:$B$5"
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$5:$B$5"
    assert data.Range("A5").Value2 == "Sum"
    assert data.Range("B5").Value2 == 10


def test_excel_opens_a_loaded_xlsb_after_totals_are_enabled(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B3"))
    workbook.save(source)
    workbook.close()

    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        sheet["A3"], sheet["B3"] = "Total", 5
        sheet.tables["Tbl"].totalsRowCount = 1
        workbook.save(out)
    finally:
        workbook.close()

    data = open_workbook(out).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$3:$B$3"
    assert data.Range("A3").Value2 == "Total"
    assert data.Range("B3").Value2 == 5


def test_excel_opens_a_headerless_xlsb_table_with_totals(tmp_path, excel, open_workbook):
    path = tmp_path / "headerless.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B3", header_row=False, totals_row=True))
    workbook.save(path)
    workbook.close()

    data = open_workbook(path).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.Range.Address == "$A$1:$B$3"
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$3:$B$3"
    assert data.Range("B3").Value2 == 5


def test_excel_sees_the_totals_row_in_xlsx(tmp_path, excel, open_workbook):
    path = tmp_path / "totals.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty"], ["a", 2], ["b", 3], ["Total", 5]], start=1):
        for column, cell in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=cell)
    sheet.add_table(ExcelTable("Tbl", "A1:B4", totals_row=True))
    workbook.save(path)
    workbook.close()

    data = open_workbook(path).Worksheets("Data")
    table = data.ListObjects("Tbl")
    assert table.Range.Address == "$A$1:$B$4"
    assert table.ShowTotals
    assert table.TotalsRowRange.Address == "$A$4:$B$4"
    assert data.Range("B4").Value2 == 5
