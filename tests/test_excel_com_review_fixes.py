"""Opt-in Excel COM checks for header/footer VML, sheet removal, read-only
coordinates, XLSB formulas/arrays and Strict Open XML reading."""

import gc
import os
import re
import struct
import zipfile
import zlib

import pytest

from xlspy import (
    Comment, DefinedName, ExcelReader, Font, Formula, Group, XlsbUpdater, XlsbWriter,
    XlsxUpdater, XlsxWriter, load_workbook,
)

pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1 to validate in Excel"),
]

XL_XLSB, XL_STRICT = 50, 61


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


def _png(path):
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    raw = b"".join(b"\0" + b"\xff\0\0" * 4 for _ in range(4))
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 4, 8, 2, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    return path


def _save_excel_book(excel, path, build, file_format=None):
    book = excel.Workbooks.Add()
    try:
        build(book)
        if file_format is None:
            book.SaveAs(str(path))
        else:
            book.SaveAs(str(path), file_format)
    finally:
        book.Close(False)
        del book
        gc.collect()


def _vml_parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name.endswith(".vml")}


@pytest.mark.parametrize("extension, updater_class, file_format", [
    ("xlsx", XlsxUpdater, None), ("xlsb", XlsbUpdater, XL_XLSB),
])
def test_notes_keep_excel_header_picture(tmp_path, excel, open_workbook, extension, updater_class, file_format):
    picture = _png(tmp_path / "logo.png")
    path = tmp_path / f"header.{extension}"

    def build(book):
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("A1").Value = 10
        sheet.PageSetup.LeftHeaderPicture.Filename = str(picture)
        sheet.PageSetup.LeftHeader = "&G"

    _save_excel_book(excel, path, build, file_format)
    header_vml = _vml_parts(path)
    assert len(header_vml) == 1
    noted, cleared = tmp_path / f"noted.{extension}", tmp_path / f"cleared.{extension}"
    for source, target, value in ((path, noted, Comment("checked", "Ann")), (noted, cleared, None)):
        updater = updater_class(source)
        try:
            updater.set_comments("Data", {"B2": value})
            updater.save(target)
        finally:
            updater.close()

    noted_vml = _vml_parts(noted)
    assert len(noted_vml) == 2
    assert all(noted_vml[name] == data for name, data in header_vml.items())
    book = open_workbook(noted)
    sheet = book.Worksheets("Data")
    assert sheet.Comments.Count == 1
    assert sheet.Range("B2").Comment.Text() == "checked"
    assert sheet.PageSetup.LeftHeader == "&G"
    assert sheet.Range("A1").Value == 10

    assert _vml_parts(cleared) == header_vml
    book = open_workbook(cleared)
    sheet = book.Worksheets("Data")
    assert sheet.Comments.Count == 0
    assert sheet.PageSetup.LeftHeader == "&G"


def test_removing_sheet_keeps_local_names_and_print_area(tmp_path, excel, open_workbook):
    path = tmp_path / "three.xlsx"

    def build(book):
        while book.Worksheets.Count < 3:
            book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        for index, name in enumerate(("First", "Second", "Third"), 1):
            book.Worksheets(index).Name = name
            book.Worksheets(index).Range("A1").Value = index
        book.Worksheets("Second").Range("B1").Formula = "=First!A1+1"
        book.Worksheets("Third").Range("B1").Formula = "=A1*2"
        book.Worksheets("Third").PageSetup.PrintArea = "$A$1:$B$2"
        book.Worksheets("Third").Names.Add("LocalC", "=Third!$A$1")
        book.Names.Add("GlobalB", "=Second!$A$1")

    _save_excel_book(excel, path, build)
    workbook = load_workbook(path)
    workbook.remove(workbook["First"])
    target = tmp_path / "two.xlsx"
    workbook.save(target)
    workbook.close()

    with zipfile.ZipFile(target) as archive:
        names = set(archive.namelist())
        rels = archive.read("xl/_rels/workbook.xml.rels").decode()
    assert "xl/calcChain.xml" not in names
    for part in re.findall(r'Target="([^"]+)"', rels):
        part = part.lstrip("/")
        assert (part if part.startswith("xl/") else "xl/" + part) in names

    book = open_workbook(target)
    assert [sheet.Name for sheet in book.Worksheets] == ["Second", "Third"]
    sheet_c = book.Worksheets("Third")
    assert sheet_c.PageSetup.PrintArea == "$A$1:$B$2"
    assert sheet_c.Names("LocalC").RefersToRange.Value == 3
    assert sheet_c.Range("B1").Value == 6
    assert book.Names("GlobalB").RefersToRange.Value == 2


def test_style_edit_keeps_cells_without_references(tmp_path, excel, open_workbook):
    source, stripped = tmp_path / "source.xlsx", tmp_path / "stripped.xlsx"
    with XlsxWriter(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[1, 2, 3], ["a", "b", "c"]])
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(stripped, "w") as output:
        for name in original.namelist():
            data = original.read(name)
            if name == "xl/worksheets/sheet1.xml":
                data = re.sub(rb'(<c)\s+r="[A-Z]+\d+"', rb"\1", data)
            output.writestr(name, data)
    workbook = load_workbook(stripped)
    workbook["S"]["B1"].font = Font(bold=True)
    target = tmp_path / "styled.xlsx"
    workbook.save(target)
    workbook.close()
    book = open_workbook(target)
    sheet = book.Worksheets("S")
    assert sheet.Range("A1:C2").Value == ((1, 2, 3), ("a", "b", "c"))
    assert sheet.Range("B1").Font.Bold
    assert not sheet.Range("C1").Font.Bold


def test_binary_streaming_and_read_only_match_excel(tmp_path, excel):
    path = tmp_path / "sparse.xlsb"

    def build(book):
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("B3").Value = "head"
        sheet.Range("C3").Value = 1
        sheet.Range("C4").Formula = "=C3*10"
        sheet.Range("E8").Value = 5

    _save_excel_book(excel, path, build, XL_XLSB)
    with ExcelReader(str(path)) as reader:
        bulk = reader.read_all("Data")
        assert reader.read_all("Data", streaming=True) == bulk
        formulas = reader.read_all("Data", data_only=False)
    # The bulk shape starts at the first data cell (B3).
    assert bulk[0][:2] == ["head", 1] and bulk[1][1] == 10 and bulk[5][3] == 5
    assert formulas[1][1] == "=C3*10"
    assert [len(row) for row in formulas] == [len(row) for row in bulk]

    workbook = load_workbook(path, read_only=True)
    try:
        sheet = workbook["Data"]
        assert sheet.dimensions == "B3:E8"
        values = list(sheet.values)
        assert len(values) == 8 and values[2][1:3] == ("head", 1)
        # As openpyxl, load_workbook defaults to data_only=False.
        assert values[3][2] == "=C3*10" and values[7][4] == 5
        assert sheet["E8"].value == 5 and sheet["A1"].value is None
    finally:
        workbook.close()


def test_read_only_xlsx_coordinates_match_excel(tmp_path, excel):
    path = tmp_path / "sparse.xlsx"

    def build(book):
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("C2").Value = "x"
        sheet.Range("A6").Value = 6
        sheet.Range("D9").Value = 9

    _save_excel_book(excel, path, build)
    workbook = load_workbook(path, read_only=True)
    try:
        sheet = workbook["Data"]
        assert (sheet.min_row, sheet.min_column, sheet.max_row, sheet.max_column) == (2, 1, 9, 4)
        rows = list(sheet.iter_rows(values_only=True))
        assert len(rows) == 9 and rows[0] == (None,) * 4
        assert rows[1][2] == "x" and rows[5][0] == 6 and rows[8][3] == 9
        cells = list(sheet.iter_rows(min_row=6, max_row=6))
        assert cells[0][0].coordinate == "A6" and cells[0][0].value == 6
    finally:
        workbook.close()


def test_binary_local_name_array_dates_and_row_options(tmp_path, open_workbook):
    import datetime
    path = tmp_path / "formulas.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [[Formula("Data!Rate*2"), 3],
             [None, None, 1, Formula("C2:C5*2", array_ref="D2:D5"), None,
              Formula("B1:C1", array_ref="F2:G2"), datetime.date(2020, 1, 1)]],
            row_heights={3: 30}, hidden_rows=[4], row_groups=[Group(2, 3)],
        )
        writer.set_defined_name(DefinedName("Rate", "Data!$B$1", "Data"))
    book = open_workbook(path)
    book.Application.CalculateFull()
    sheet = book.Worksheets("Data")
    assert sheet.Range("A1").Value == 6
    assert sheet.Range("D2").HasArray
    assert sheet.Range("D2").CurrentArray.Address == "$D$2:$D$5"
    assert sheet.Range("D2").Value == 2
    assert sheet.Range("F2").CurrentArray.Address == "$F$2:$G$2"
    assert sheet.Range("F2").Value == 3
    assert sheet.Rows(4).RowHeight == pytest.approx(30)
    assert sheet.Rows(5).Hidden
    assert sheet.Rows(3).OutlineLevel == 2 and sheet.Rows(4).OutlineLevel == 2


def test_xlsx_auto_filter_ref_in_excel(tmp_path, open_workbook):
    path = tmp_path / "filter.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a", "b", "c"], [1, 2, 3], [4, 5, 6]], auto_filter_ref="A1:B3")
    book = open_workbook(path)
    sheet = book.Worksheets("Data")
    assert sheet.AutoFilterMode
    assert sheet.AutoFilter.Range.Address == "$A$1:$B$3"


def test_strict_open_xml_reading(tmp_path, excel):
    path = tmp_path / "strict.xlsx"

    def build(book):
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("A1").Value = "name"
        sheet.Range("B1").Value = 2.5
        sheet.Range("C1").Formula = "=B1*2"
        sheet.Range("D1").Value2 = 45356  # 2024-03-05 as a serial avoids COM time zones
        sheet.Range("D1").NumberFormat = "yyyy-mm-dd"
        sheet.Range("B4").Value = 7
        sheet.Range("A1").AddComment("strict note")
        book.Names.Add("Amount", "=Data!$B$1")

    _save_excel_book(excel, path, build, XL_STRICT)
    with zipfile.ZipFile(path) as archive:
        assert b"http://purl.oclc.org/ooxml/spreadsheetml/main" in archive.read("xl/workbook.xml")
    import datetime
    with ExcelReader(str(path)) as reader:
        assert reader.get_sheet_names() == ["Data"]
        rows = reader.read_all("Data")
        assert rows[0] == ["name", 2.5, 5, datetime.datetime(2024, 3, 5)]
        assert rows[1][1] == 7
        assert reader.read_all("Data", data_only=False)[0][2] == "=B1*2"
        assert reader.get_comments("Data")["A1"].text == "strict note"
        assert reader.get_defined_names() == [DefinedName("Amount", "Data!$B$1")]
    workbook = load_workbook(path, read_only=True)
    try:
        assert list(workbook["Data"].values)[3][1] == 7
    finally:
        workbook.close()
    with pytest.raises(NotImplementedError, match="Strict"):
        XlsxUpdater(path)
    with pytest.raises(NotImplementedError, match="Strict"):
        load_workbook(path)


@pytest.mark.parametrize("edit", ["unchanged", "values", "font", "add_sheet", "rename"])
def test_load_workbook_edits_of_excel_files_open_in_excel(tmp_path, excel, open_workbook, edit):
    path = tmp_path / "excel.xlsx"

    def build(book):
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        sheet.Range("A1").Value = "name"
        sheet.Range("B1").Value = 2
        sheet.Range("C1").Formula = "=B1*2"
        sheet.Range("A1").Font.Italic = True

    _save_excel_book(excel, path, build)
    workbook = load_workbook(path)
    sheet = workbook["Data"]
    if edit == "values":
        sheet["B1"] = 5
        sheet["D4"] = "new"
    elif edit == "font":
        sheet["B1"].font = Font(bold=True)
    elif edit == "add_sheet":
        workbook.create_sheet("Extra", 0)["A1"] = "x"
    elif edit == "rename":
        sheet.title = "Renamed"
    target = tmp_path / f"{edit}.xlsx"
    workbook.save(target)
    workbook.close()

    book = open_workbook(target)
    book.Application.CalculateFull()
    names = [item.Name for item in book.Worksheets]
    result = book.Worksheets("Renamed" if edit == "rename" else "Data")
    assert result.Range("A1").Value == "name" and result.Range("A1").Font.Italic
    if edit == "values":
        assert result.Range("C1").Value == 10 and result.Range("D4").Value == "new"
    else:
        assert result.Range("C1").Value == 4
    if edit == "font":
        assert result.Range("B1").Font.Bold
    if edit == "add_sheet":
        assert names == ["Extra", "Data"] and book.Worksheets("Extra").Range("A1").Value == "x"


def test_local_names_with_chartsheet_in_excel(tmp_path, excel, open_workbook):
    path = tmp_path / "chartsheet.xlsx"

    def build(book):
        data = book.Worksheets(1)
        data.Name = "Data"
        data.Range("A1:A3").Value = ((1,), (2,), (3,))
        # pywin32 ignores Before=/After= keywords here, so pass positions.
        other = book.Worksheets.Add(None, data)
        other.Name = "Other"
        chart = book.Charts.Add()
        chart.SetSourceData(data.Range("A1:A3"))
        chart.Name = "Graph"
        chart.Move(book.Sheets(1))  # chartsheet first: localSheetId counts it
        other.Range("A1").Value = 10
        other.Names.Add("LocalOther", "=Other!$A$1")
        data.Names.Add("LocalData", "=Data!$A$2")

    _save_excel_book(excel, path, build)
    with zipfile.ZipFile(path) as archive:
        order = re.findall(r'<sheet name="([^"]+)"', archive.read("xl/workbook.xml").decode())
    assert order == ["Graph", "Data", "Other"]
    workbook = load_workbook(path)
    scopes = {name.name: name.scope for name in workbook.defined_names}
    assert scopes == {"LocalData": "Data", "LocalOther": "Other"}
    workbook.remove(workbook["Data"])
    target = tmp_path / "edited.xlsx"
    workbook.save(target)
    workbook.close()

    book = open_workbook(target)
    assert [sheet.Name for sheet in book.Sheets] == ["Graph", "Other"]
    assert book.Worksheets("Other").Names("LocalOther").RefersToRange.Value == 10
    assert book.Charts.Count == 1
