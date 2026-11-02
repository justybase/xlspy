"""Excel opens an Excel-written XLSB after layout and sheet-list edits by the model."""


import pytest

from xlspy import load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

XL_XLSB = 50


@pytest.fixture
def native(tmp_path, excel):
    """A workbook saved by Excel itself, so the edits patch Excel's own records."""
    book = excel.Workbooks.Add()
    while book.Worksheets.Count < 3:
        book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
    for index, name in enumerate(("Data", "Lists", "Report"), 1):
        book.Worksheets(index).Name = name
    data = book.Worksheets("Data")
    for row in range(1, 11):
        data.Cells(row, 1).Value = f"name{row}"
        data.Cells(row, 2).Value = row
    book.Worksheets("Report").Range("A1").Formula = "=SUM(Data!B1:B10)"
    book.Worksheets("Report").Range("A2").Formula = "=Lists!A1"
    book.Names.Add("Total", "=Data!$B$1:$B$10")
    path = tmp_path / "native.xlsb"
    book.SaveAs(str(path), XL_XLSB)
    book.Close(False)
    return path


def test_layout_edits_in_excel(native, tmp_path, open_workbook):
    workbook = load_workbook(native)
    try:
        sheet = workbook["Data"]
        sheet.column_dimensions["A"].width = 30
        sheet.column_dimensions["C"].hidden = True
        sheet.row_dimensions[2].height = 33
        sheet.row_dimensions[3].hidden = True
        sheet.freeze_panes = "B2"
        sheet.sheet_view.zoomScale = 150
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.tabColor = "FF00AA00"
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.paperSize = 9
        sheet.page_margins.left = 1.25
        sheet.oddHeader.center.text = "Report &P"
        sheet.print_area = "A1:B5"
        sheet.print_title_rows = "1:1"
        sheet.auto_filter.ref = "A1:B10"
        sheet.protection.sheet = True
        sheet.protection.formatColumns = False
        workbook.save(tmp_path / "edited.xlsb")
    finally:
        workbook.close()

    book = open_workbook(tmp_path / "edited.xlsb")
    sheet = book.Worksheets("Data")
    assert sheet.Columns("A").ColumnWidth == pytest.approx(30, abs=0.8)
    assert sheet.Columns("C").Hidden
    assert sheet.Rows(2).RowHeight == pytest.approx(33)
    assert sheet.Rows(3).Hidden
    assert sheet.Tab.Color == 0x00AA00
    assert sheet.ProtectContents
    assert sheet.Protection.AllowFormattingColumns
    assert not sheet.Protection.AllowFormattingCells
    assert sheet.AutoFilterMode
    assert sheet.AutoFilter.Range.Address == "$A$1:$B$10"
    assert sheet.Range("A2").Value == "name2"
    assert book.Worksheets("Data").Names("Print_Area").RefersTo == "=Data!$A$1:$B$5"
    assert book.Worksheets("Data").Names("Print_Titles").RefersTo == "=Data!$1:$1"
    book.Worksheets("Data").Activate()
    window = book.Windows(1)
    assert window.FreezePanes and window.SplitRow == 1 and window.SplitColumn == 1
    assert window.Zoom == 150
    assert not window.DisplayGridlines


def test_sheet_list_edits_in_excel(native, tmp_path, open_workbook):
    workbook = load_workbook(native)
    try:
        workbook["Lists"].title = "Choices"
        workbook.remove(workbook["Data"])
        new = workbook.create_sheet("New", 0)
        new["A1"] = "fresh"
        new["B1"] = "=Choices!A1"
        workbook["Report"].sheet_state = "hidden"
        workbook.active = 0
        workbook.properties.title = "From xlspy"
        workbook.save(tmp_path / "sheets.xlsb")
    finally:
        workbook.close()

    book = open_workbook(tmp_path / "sheets.xlsb")
    assert [sheet.Name for sheet in book.Worksheets] == ["New", "Choices", "Report"]
    assert book.ActiveSheet.Name == "New"
    assert book.Worksheets("Report").Visible == 0  # xlSheetHidden
    assert book.Worksheets("New").Range("A1").Value == "fresh"
    assert book.Worksheets("New").Range("B1").Formula == "=Choices!A1"
    report = book.Worksheets("Report")
    assert report.Range("A2").Formula == "=Choices!A1"
    assert "#REF!" in report.Range("A1").Formula
    assert book.BuiltinDocumentProperties("Title").Value == "From xlspy"
