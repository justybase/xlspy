"""Excel reads colour scales, data bars and icon sets written into XLSX files."""

import pytest

from xlspy import Workbook

from test_excel_com_xlsb_cf import _check, _fill, _rules

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)


def test_new_xlsx_conditional_formats(tmp_path, open_workbook):
    workbook = Workbook()
    _fill(workbook.active)
    for sqref, rule in _rules():
        workbook.active.conditional_formatting.add(sqref, rule)
    workbook.save(tmp_path / "new.xlsx")
    workbook.close()
    _check(open_workbook(tmp_path / "new.xlsx"))
