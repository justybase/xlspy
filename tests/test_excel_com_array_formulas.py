"""Excel checks array and data-table formulas written by xlspy, and xlspy reads Excel's.

The data table is opened by Excel (which reports it as a TABLE array formula), then
saved again by Excel; the re-saved file is read back through ``load_workbook``.
"""

from __future__ import annotations

import pytest

from xlspy import ArrayFormula, DataTableFormula, XlsbWriter, XlsxWriter, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]
XL_FILE_XLSX = 51


def _array_workbook(path, writer_class):
    with writer_class(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([
            ["input", "array"],
            [1, ArrayFormula("B2:B4", "A2:A4*2")],
            [2, None],
            [3, None],
        ])


@pytest.mark.parametrize("writer_class, suffix", WRITERS)
def test_excel_calculates_an_array_formula(tmp_path, open_workbook, writer_class, suffix):
    path = tmp_path / f"array.{suffix}"
    _array_workbook(path, writer_class)

    sheet = open_workbook(path).Worksheets("S")
    sheet.Application.Calculate()
    assert sheet.Range("B2").HasArray
    assert sheet.Range("B2").FormulaArray == "=A2:A4*2"
    assert [sheet.Cells(row, 2).Value for row in (2, 3, 4)] == [2.0, 4.0, 6.0]


def _data_table_workbook(path):
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([
            [None, 1, 2, None],
            [10, DataTableFormula("B2:C4", dt2D=True, r1="D1", r2="D2"), None, 5],
            [20, None, None, None],
            [30, None, None, None],
        ])


def test_excel_opens_a_data_table_formula(tmp_path, open_workbook):
    path = tmp_path / "table.xlsx"
    _data_table_workbook(path)

    sheet = open_workbook(path).Worksheets("S")
    sheet.Application.Calculate()
    assert sheet.Range("B2").HasArray
    # Excel localises the TABLE function name but keeps the input cells.
    formula = sheet.Range("B2").FormulaArray
    assert "D1" in formula and "D2" in formula
    assert all(
        sheet.Cells(row, column).HasArray
        for row in (2, 3, 4) for column in (2, 3)
    )


def test_xlspy_reads_a_data_table_resaved_by_excel(tmp_path, excel):
    source = tmp_path / "table.xlsx"
    resaved = tmp_path / "excel_table.xlsx"
    _data_table_workbook(source)

    book = excel.Workbooks.Open(str(source), 0, False)
    try:
        book.Application.Calculate()
        book.SaveAs(str(resaved), XL_FILE_XLSX)
    finally:
        book.Close(False)

    loaded = load_workbook(resaved)
    try:
        value = loaded["S"]["B2"].value
        assert isinstance(value, DataTableFormula)
        assert value.ref == "B2:C4"
        assert value.dt2D is True
        assert (value.r1, value.r2) == ("D1", "D2")
    finally:
        loaded.close()


@pytest.mark.parametrize("suffix, file_format", [("xlsx", 51), ("xlsb", 50)])
def test_xlspy_reads_an_array_formula_authored_by_excel(tmp_path, excel, suffix, file_format):
    path = tmp_path / f"excel_array.{suffix}"
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "S"
        sheet.Range("A1").Value = "input"
        sheet.Range("B1").Value = "array"
        for row, value in enumerate((1, 2, 3), start=2):
            sheet.Cells(row, 1).Value = value
        sheet.Range("B2:B4").FormulaArray = "=A2:A4*2"
        book.SaveAs(str(path), file_format)
    finally:
        book.Close(False)

    loaded = load_workbook(path)
    try:
        cell = loaded["S"]["B2"]
        if suffix == "xlsx":
            assert isinstance(cell.value, ArrayFormula)
            assert cell.value.ref == "B2:B4"
            assert loaded["S"].array_formulae == {"B2": "B2:B4"}
        else:
            # Loading XLSB array groups into ArrayFormula objects is the next step.
            assert cell.value == "=A2:A4*2"
    finally:
        loaded.close()
