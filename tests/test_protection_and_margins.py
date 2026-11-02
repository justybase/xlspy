"""The public SheetProtection and PageMargins applied to the workbook model."""
from __future__ import annotations

import pytest

import xlspy
from xlspy._metadata import password_hash
from xlspy.workbook import Workbook, load_workbook
from xlspy.workbook import WorksheetPageMargins as ModelPageMargins
from xlspy.workbook import WorksheetProtection as ModelSheetProtection


def test_public_sheet_protection_saves_and_reloads(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.protection = xlspy.SheetProtection(password="abc")

    # The model holds its own type, never the public class.
    assert type(ws.protection) is ModelSheetProtection
    assert ws.protection.sheet is True
    assert ws.protection.objects is True

    path = tmp_path / "protected.xlsx"
    wb.save(path)

    loaded = load_workbook(path)
    protection = loaded.active.protection
    assert protection.sheet is True
    assert protection.password == password_hash("abc")


def test_public_sheet_protection_flags_are_inverted_for_model(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    # Public flags mean "selection allowed"; the model's OOXML flags mean "locked".
    ws.protection = xlspy.SheetProtection(select_locked_cells=False, select_unlocked_cells=True)
    assert ws.protection.selectLockedCells is True
    assert ws.protection.selectUnlockedCells is False


def test_disabled_public_protection_does_not_protect(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.protection = xlspy.SheetProtection(enabled=False)
    assert ws.protection.sheet is False
    wb.save(tmp_path / "unprotected.xlsx")
    assert load_workbook(tmp_path / "unprotected.xlsx").active.protection.sheet is False


def test_protection_rejects_unknown_types():
    ws = Workbook().active
    with pytest.raises(TypeError):
        ws.protection = "locked"


def test_public_page_margins_become_model_margins(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws.page_margins = xlspy.PageMargins(left=1.5, top=2.0)

    assert type(ws.page_margins) is ModelPageMargins
    # The model object is mutable, as in openpyxl.
    ws.page_margins.right = 1.25

    path = tmp_path / "margins.xlsx"
    wb.save(path)
    margins = load_workbook(path).active.page_margins
    assert margins.left == pytest.approx(1.5)
    assert margins.top == pytest.approx(2.0)
    assert margins.right == pytest.approx(1.25)
