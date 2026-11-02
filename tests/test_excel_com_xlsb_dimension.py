"""Excel as the oracle for BrtWsDim (PR #1): the dimension XlsbWriter writes must be the one Excel
writes when it resaves the same workbook. Excel itself ignores the record on load (UsedRange is
recomputed), which is why the all-zero placeholder of xlspy before 0.11 went unnoticed."""

import pytest

from test_xlsb_interop import DIMENSION_CASES, _write, sheet_dimension, write_case

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50


def _excel_resave(excel, path, target):
    book = excel.Workbooks.Open(str(path), 0, True)
    try:
        book.SaveAs(str(target), XL_XLSB)
    finally:
        book.Close(False)


@pytest.mark.parametrize("case", sorted(DIMENSION_CASES))
def test_dimension_matches_excel(tmp_path, excel, case):
    path, resaved = tmp_path / f"{case}.xlsb", tmp_path / f"{case}_excel.xlsb"
    write_case(path, case)
    _excel_resave(excel, path, resaved)
    # The table in test_xlsb_interop.py is Excel's own answer.
    assert sheet_dimension(resaved) == DIMENSION_CASES[case][2]
    assert sheet_dimension(path) == sheet_dimension(resaved)


def test_spooled_sheet_opens_in_excel(tmp_path, excel, open_workbook):
    # Larger than the row buffer: the head is written after the compressed rows.
    path, resaved = tmp_path / "big.xlsb", tmp_path / "big_excel.xlsb"
    _write(path, ([None, f"row {i}", i, i * 1.5] for i in range(30000)), hidden_columns=["C"])
    book = open_workbook(path)
    sheet = book.Worksheets("Data")
    assert sheet.UsedRange.Address == "$B$1:$D$30000"
    assert sheet.Range("B30000").Value == "row 29999"
    assert sheet.Range("D30000").Value == pytest.approx(44998.5)
    assert sheet.Columns("C").Hidden
    book.SaveAs(str(resaved), XL_XLSB)
    assert sheet_dimension(path) == sheet_dimension(resaved) == (0, 29999, 1, 3)
