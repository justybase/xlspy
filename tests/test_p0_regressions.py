"""Regression tests for the P0 review items.

- Assigning the public ``xlspy.SheetProtection`` / ``xlspy.PageMargins`` to a
  worksheet used to save with an AttributeError (protection) or keep the frozen
  public object (margins).
- Saving a model whose cells are far apart must not build rows as wide as the
  sheet.
- A package whose parts add up to more than the package limit is refused.
"""

import time
import zipfile

import pytest

import xlspy
from xlspy import _untrusted
from xlspy._metadata import password_hash
from xlspy.workbook import Workbook, load_workbook
from xlspy.workbook._model import PageMargins as ModelPageMargins
from xlspy.workbook._model import SheetProtection as ModelSheetProtection


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


def test_widely_separated_cells_save_quickly(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "start"
    ws["XFD1"] = "far right"
    ws["A20000"] = "far down"

    path = tmp_path / "sparse.xlsx"
    started = time.perf_counter()
    wb.save(path)
    elapsed = time.perf_counter() - started

    # The old dense loop built 20000 rows of 16384 entries each (about 330 M
    # iterations); the bounded loop finishes in well under a second.
    assert elapsed < 20

    loaded = load_workbook(path).active
    assert loaded["A1"].value == "start"
    assert loaded["XFD1"].value == "far right"
    assert loaded["A20000"].value == "far down"


def test_styled_empty_cell_far_right_keeps_its_style(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws["D1"].font = xlspy.Font(bold=True)
    path = tmp_path / "styled.xlsx"
    wb.save(path)
    assert load_workbook(path).active["D1"].font.bold is True


def test_package_total_size_is_limited(tmp_path, monkeypatch):
    path = tmp_path / "many.zip"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("a.xml", b"x" * 1000)
        archive.writestr("b.xml", b"y" * 1000)

    monkeypatch.setattr(_untrusted, "MAX_PACKAGE_BYTES", 1500)
    with pytest.raises(ValueError, match="package declares"):
        _untrusted.SafeZipFile(path, "r")

    monkeypatch.setattr(_untrusted, "MAX_PACKAGE_BYTES", 4000)
    with _untrusted.SafeZipFile(path, "r") as archive:
        assert archive.read("a.xml") == b"x" * 1000
