"""Opt-in validation of notes, template edits and protection in real Excel."""

import os
import gc

import pytest

from xlspy import Comment, DefinedName, Formula, SheetProtection, WorkbookProtection, XlsxUpdater, XlsxWriter

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1 to validate in Excel"),
]


@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    yield app
    gc.collect()
    app.Quit()
    del app
    gc.collect()


@pytest.fixture
def open_workbook(excel):
    opened = []
    def open_file(path):
        workbook = excel.Workbooks.Open(str(path), 0, True)
        opened.append(workbook)
        return workbook
    yield open_file
    while opened:
        workbook = opened.pop()
        try:
            workbook.Close(False)
        finally:
            del workbook
    gc.collect()


def test_notes_dimensions_names_and_protection_in_excel(tmp_path, open_workbook):
    path = tmp_path / "metadata.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1], [2]], comments={"A1": Comment("Zażółć <&>\nline", "Jan", 300, 120), "C8": Comment("literal _x0041_", "Ola")},
                           sheet_protection=SheetProtection(password="abc"))
        writer.set_defined_name(DefinedName("Amount", "'Data'!$A$2"))
        writer.set_workbook_protection(WorkbookProtection(password="abc"))
    book = open_workbook(path)
    sheet = book.Worksheets("Data")
    assert sheet.Range("A1").Comment.Text() == "Zażółć <&>\nline"
    assert sheet.Range("C8").Comment.Text() == "literal _x0041_"
    assert sheet.Range("A1").Comment.Author == "Jan"
    assert sheet.Range("A1").Comment.Shape.Width == pytest.approx(225, abs=1)
    assert sheet.Range("A1").Comment.Shape.Height == pytest.approx(90, abs=1)
    assert sheet.ProtectContents
    assert book.ProtectStructure
    assert book.Names("Amount").RefersToRange.Value == 2


def test_edit_notes_preserves_excel_button(tmp_path, excel, open_workbook):
    path = tmp_path / "button.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[10]], comments={"A1": Comment("old", "Jan")})
    book = excel.Workbooks.Open(str(path))
    try:
        sheet = book.Worksheets("Data")
        button = sheet.Buttons().Add(20, 20, 80, 24)
        button.Name = "PreservedButton"
        button.Caption = "Click"
        book.Save()
    finally:
        book.Close(False)
    updater = XlsxUpdater(path)
    try:
        updater.set_comments("Data", {"A1": None, "C3": Comment("new", "Ola", 200, 100)})
        updater.update_cells("Data", {"A1": 15, "B1": Formula("A1*2", 30)})
        updater.save()
    finally:
        updater.close()
    book = open_workbook(path)
    sheet = book.Worksheets("Data")
    assert sheet.Buttons().Count == 1
    assert sheet.Buttons("PreservedButton").Caption == "Click"
    assert sheet.Range("A1").Comment is None
    assert sheet.Range("C3").Comment.Text() == "new"
    assert sheet.Range("B1").Value == 30
