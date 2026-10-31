"""New .xlsm/.xltx/.xltm workbooks and keep_vba, as in openpyxl."""

import io
import os
import zipfile

import pytest

from xlspy import Workbook, load_workbook

MAIN = {
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}


def _vba_source(tmp_path):
    """A .xlsm package with a (stand-in) VBA project."""
    workbook = Workbook()
    workbook.active["A1"] = "macro book"
    workbook.save(tmp_path / "plain.xlsx")
    workbook.close()
    path = tmp_path / "macros.xlsm"
    with zipfile.ZipFile(tmp_path / "plain.xlsx") as source, zipfile.ZipFile(path, "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item))
        target.writestr("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0 stand-in VBA project")
    return path


@pytest.mark.parametrize("extension", sorted(MAIN))
def test_new_workbook_kinds(tmp_path, extension):
    workbook = Workbook()
    workbook.active["A1"] = 1
    path = tmp_path / f"book.{extension}"
    workbook.save(path)
    workbook.close()
    with zipfile.ZipFile(path) as archive:
        types = archive.read("[Content_Types].xml").decode()
        assert f'ContentType="{MAIN[extension]}"' in types
        assert "xl/vbaProject.bin" not in archive.namelist()
    reloaded = load_workbook(path)
    try:
        assert reloaded.active["A1"].value == 1
    finally:
        reloaded.close()


def test_keep_vba_moves_the_project_into_a_new_workbook(tmp_path):
    source = load_workbook(_vba_source(tmp_path), keep_vba=True)
    try:
        assert isinstance(source.vba_archive, zipfile.ZipFile)
        workbook = Workbook()
        workbook.active["A1"] = "new"
        workbook.vba_archive = source.vba_archive
        stream = io.BytesIO()
        workbook.save(tmp_path / "new.xlsm")
        workbook.save(stream)  # a stream is written as .xlsx: no extension to go by
    finally:
        source.close()
    with zipfile.ZipFile(tmp_path / "new.xlsm") as archive:
        assert archive.read("xl/vbaProject.bin").endswith(b"stand-in VBA project")
        rels = archive.read("xl/_rels/workbook.xml.rels").decode()
        assert "relationships/vbaProject" in rels and 'Target="vbaProject.bin"' in rels
        assert "application/vnd.ms-office.vbaProject" in archive.read("[Content_Types].xml").decode()


def test_loaded_workbook_has_no_vba_archive_by_default(tmp_path):
    workbook = load_workbook(_vba_source(tmp_path))
    try:
        assert workbook.vba_archive is None
    finally:
        workbook.close()


@pytest.mark.com
@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1")
def test_excel_runs_the_moved_macro(tmp_path):
    import pywintypes
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Add()
        try:
            module = book.VBProject.VBComponents.Add(1)
        except pywintypes.com_error:
            book.Close(False)
            pytest.skip("Excel does not trust access to the VBA project model")
        module.CodeModule.AddFromString('Function Answer()\n    Answer = 42\nEnd Function\n')
        book.SaveAs(str(tmp_path / "excel.xlsm"), 52)  # xlOpenXMLWorkbookMacroEnabled
        book.Close(False)

        source = load_workbook(tmp_path / "excel.xlsm", keep_vba=True)
        try:
            workbook = Workbook()
            workbook.active["A1"] = "new"
            workbook.vba_archive = source.vba_archive
            workbook.save(tmp_path / "moved.xlsm")
            workbook.close()
        finally:
            source.close()

        book = excel.Workbooks.Open(str(tmp_path / "moved.xlsm"))
        try:
            assert book.HasVBProject
            assert excel.Run("'moved.xlsm'!Answer") == 42
            assert book.Worksheets(1).Range("A1").Value == "new"
        finally:
            book.Close(False)
    finally:
        excel.Quit()


@pytest.mark.com
@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1")
def test_excel_opens_new_templates_and_macro_books(tmp_path):
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    def opened(path):
        """How Excel takes the file: its format and whether it opened a copy."""
        book = excel.Workbooks.Open(str(path))
        try:
            return book.FileFormat, book.Name == path.name, book.Worksheets(1).Range("A1").Value
        finally:
            book.Close(False)

    try:
        # xlOpenXMLWorkbookMacroEnabled, xlOpenXMLTemplateMacroEnabled, xlOpenXMLTemplate
        for extension, file_format in (("xlsm", 52), ("xltm", 53), ("xltx", 54)):
            native = tmp_path / f"excel.{extension}"
            book = excel.Workbooks.Add()
            book.Worksheets(1).Range("A1").Value = extension
            book.SaveAs(str(native), file_format)
            book.Close(False)
            path = tmp_path / f"book.{extension}"
            workbook = Workbook()
            workbook.active["A1"] = extension
            workbook.save(path)
            workbook.close()
            # Excel opens a copy of a template: ours must behave as its own.
            assert opened(path) == opened(native)
            assert opened(path)[2] == extension
    finally:
        excel.Quit()
