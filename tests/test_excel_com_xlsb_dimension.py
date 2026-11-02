"""Excel as the oracle for BrtWsDim (PR #1): the dimension XlsbWriter writes must match the one Excel
writes when it resaves the same workbook. Excel itself ignores the record on load (UsedRange is
recomputed), which is why the all-zero placeholder went unnoticed."""

import pytest

from test_xlsb_interop import DIMENSION_CASES, _write, sheet_dimension

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
    rows, options, expected = DIMENSION_CASES[case]
    path, resaved = tmp_path / f"{case}.xlsb", tmp_path / f"{case}_excel.xlsb"
    _write(path, rows, **options)
    _excel_resave(excel, path, resaved)
    # The table's expectations are Excel's own; keep them honest.
    assert sheet_dimension(resaved) == expected
    ours = sheet_dimension(path)
    if ours == (0, 0, 0, 0) != expected:
        # Known defect: drop this branch once the writer records the used range.
        pytest.xfail("XlsbWriter writes an all-zero BrtWsDim (PR #1)")
    assert ours == expected
