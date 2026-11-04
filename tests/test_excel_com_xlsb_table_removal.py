"""An XLSB saved after removing a table opens in Excel with the other table."""

import pytest

from xlspy import ExcelTable, Workbook, load_workbook

pytestmark = pytest.mark.com


def test_excel_opens_an_xlsb_after_a_table_is_removed(tmp_path, excel, open_workbook):
    source = tmp_path / "source.xlsb"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row, values in enumerate([["name", "qty", "", "code", "n"], ["a", 2, None, "x", 10], ["b", 3, None, "y", 20]], start=1):
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value)
    sheet.add_table(ExcelTable("Tbl", "A1:B3"))
    sheet.add_table(ExcelTable("Other", "D1:E3"))
    workbook.save(source)
    workbook.close()

    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        workbook["Data"].tables.remove("Other")
        workbook.save(out)
    finally:
        workbook.close()

    data = open_workbook(out).Worksheets("Data")
    names = [data.ListObjects(index).Name for index in range(1, data.ListObjects.Count + 1)]
    assert names == ["Tbl"]
    assert data.Range("A2").Value2 == "a"
    assert data.Range("D2").Value2 == "x"
