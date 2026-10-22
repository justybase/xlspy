"""Opt-in checks in Microsoft Excel (COM) of openpyxl-style workbook model
scripts: objects, print setup, zoom, row/column moves, rich text and the
extra conditional formatting rules. Set XLSPY_RUN_EXCEL_COM=1 to run.
"""

import datetime
import io
import os
import struct
import zlib

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)

from xlspy import (  # noqa: E402
    AboveAverageRule, BlanksRule, CellIsRule, CellRichText, Chart, ChartSeries, Comment,
    DataValidation, DifferentialStyle, DuplicateRule, ExcelTable, Font, Image, InlineFont,
    TextBlock, TextRule, TimePeriodRule, TopBottomRule, Workbook, XlsbWriter, XlsxWriter,
    load_workbook,
)

BOLD = DifferentialStyle(font=Font(bold=True))
XL_LANDSCAPE = 2
XL_VALIDATE_WHOLE_NUMBER = 1


def _bgr(rgb: str) -> int:
    return int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)


def _png(width, height):
    raw = b"".join(b"\x00" + b"\x10\x80\x10" * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    yield app
    app.Quit()


@pytest.fixture
def open_workbook(excel):
    opened = []

    def _open(path):
        workbook = excel.Workbooks.Open(str(path))
        opened.append(workbook)
        return workbook

    yield _open
    for workbook in opened:
        workbook.Close(False)


def _report():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet.append(["Region", "Sales"])
    for row in (["North", 10], ["South", 30], ["East", 20]):
        sheet.append(row)
    return workbook, sheet


def _model_report():
    workbook, sheet = _report()
    sheet.add_chart(Chart("bar", [ChartSeries("B2:B4", categories="A2:A4")]), "D2")
    sheet.add_image(Image(io.BytesIO(_png(40, 20))), "D20")
    sheet.add_table(ExcelTable(display_name="Sales", ref="A1:B4"))
    sheet.add_data_validation(DataValidation(type="whole", sqref="B2:B4", operator="between",
                                             formula1="0", formula2="100"))
    sheet.conditional_formatting.add("B2:B4", CellIsRule("A1", "greaterThan", ["15"], dxf=BOLD))
    sheet.print_title_rows = "1:1"
    sheet.print_area = "A1:B4"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToHeight = 0
    sheet.oddHeader.center.text = "Report"
    sheet.sheet_view.zoomScale = 85
    return workbook


def test_model_objects_and_zoom_in_excel(tmp_path, open_workbook):
    path = tmp_path / "model.xlsx"
    _model_report().save(path)

    book = open_workbook(path)
    ws = book.Worksheets("Report")
    assert ws.ChartObjects().Count == 1
    assert ws.Shapes.Count == 2  # the chart and the picture
    assert ws.ListObjects.Count == 1 and ws.ListObjects(1).Name == "Sales"
    assert ws.Range("B2").Validation.Type == XL_VALIDATE_WHOLE_NUMBER
    assert ws.Range("B2:B4").FormatConditions.Count == 1
    assert ws.Names("Print_Area").RefersTo == "=Report!$A$1:$B$4"
    ws.Activate()
    assert book.Windows(1).Zoom == 85


@pytest.mark.printer
def test_model_print_setup_in_excel(tmp_path, open_workbook):
    path = tmp_path / "model.xlsx"
    _model_report().save(path)

    ws = open_workbook(path).Worksheets("Report")
    setup = ws.PageSetup
    assert setup.PrintTitleRows == "$1:$1"
    assert setup.PrintArea == "$A$1:$B$4"
    # PaperSize needs an installed printer that knows A4, so the XML tests
    # check it (test_model_openpyxl_compat.py).
    assert setup.Orientation == XL_LANDSCAPE
    assert setup.FitToPagesWide == 1 and setup.FitToPagesTall is False
    assert setup.CenterHeader == "Report"


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_writer_zoom_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"zoom.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1, 2]], zoom=150)
    book = open_workbook(path)
    assert book.Windows(1).Zoom == 150


def test_insert_and_delete_in_a_loaded_workbook_in_excel(tmp_path, open_workbook):
    workbook, sheet = _report()
    sheet["A2"].font = Font(bold=True)
    sheet["A3"].comment = Comment("south", "me")
    sheet.row_dimensions[3].height = 30
    workbook.save(tmp_path / "base.xlsx")
    loaded = load_workbook(tmp_path / "base.xlsx")
    ws = loaded["Report"]
    ws.insert_rows(2, 2)
    ws.insert_cols(1)
    ws["B2"] = "inserted"
    ws.delete_rows(3)
    loaded.save(tmp_path / "moved.xlsx")
    loaded.close()

    xl = open_workbook(tmp_path / "moved.xlsx").Worksheets("Report")
    assert [xl.Range(f"B{row}").Value for row in range(1, 6)] == ["Region", "inserted", "North", "South", "East"]
    assert [xl.Range(f"C{row}").Value for row in range(1, 6)] == ["Sales", None, 10, 30, 20]
    assert xl.Range("A1").Value is None
    assert xl.Range("B3").Font.Bold and not xl.Range("B2").Font.Bold
    assert xl.Range("B4").Comment.Text() == "south"
    assert xl.Rows(4).RowHeight == 30 and xl.Rows(3).RowHeight != 30


def test_rich_text_in_excel(tmp_path, open_workbook):
    rich = CellRichText("Total: ", TextBlock(InlineFont(b=True, color="C00000", sz=14), "1234"), " EUR")
    path = tmp_path / "rich.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["head"], [rich]])
    workbook = Workbook()
    workbook.active["A1"] = rich
    workbook.save(tmp_path / "model_rich.xlsx")
    for file in (path, tmp_path / "model_rich.xlsx"):
        sheet = open_workbook(file).Worksheets(1)
        cell = sheet.Range("A2" if file == path else "A1")
        assert cell.Value == "Total: 1234 EUR"
        plain, bold = cell.GetCharacters(1, 7).Font, cell.GetCharacters(8, 4).Font
        assert not plain.Bold and bold.Bold
        assert bold.Color == _bgr("C00000") and bold.Size == 14


# XlFormatConditionType values Excel reports for each rule.
CF_CASES = [
    (TopBottomRule("A2:A9", rank=2, dxf=BOLD), 5),
    (AboveAverageRule("A2:A9", dxf=BOLD), 12),
    (TextRule("B2:B9", "contains", "ab", dxf=BOLD), 9),
    (DuplicateRule("B2:B9", dxf=BOLD), 8),
    (TimePeriodRule("C2:C9", "today", dxf=BOLD), 11),
    (BlanksRule("D2:D9", dxf=BOLD), 10),
    (BlanksRule("D2:D9", "errors", dxf=BOLD), 16),
]


def test_extra_conditional_formatting_in_excel(tmp_path, open_workbook):
    today = datetime.date.today()
    rows = [["n", "text", "date", "blank"]] + [
        [n, text, today - datetime.timedelta(days=n), None]
        for n, text in zip(range(8), ["ab", "xab", "cd", "ab", "ef", "gh", "ab", "zz"])
    ]
    path = tmp_path / "cf.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, conditional_formats=[rule for rule, _ in CF_CASES])
    sheet = open_workbook(path).Worksheets("S")
    types = {}
    for column in "ABCD":
        conditions = sheet.Range(f"{column}2:{column}9").FormatConditions
        types[column] = [conditions(i).Type for i in range(1, conditions.Count + 1)]
    assert types == {"A": [5, 12], "B": [9, 8], "C": [11], "D": [10, 16]}
    # Excel applies them: the top 2 of 0..7 are 6 and 7, "ab" occurs in B2..B8.
    assert sheet.Range("A9").DisplayFormat.Font.Bold and not sheet.Range("A3").DisplayFormat.Font.Bold
    assert sheet.Range("B3").DisplayFormat.Font.Bold and not sheet.Range("B4").DisplayFormat.Font.Bold
    assert sheet.Range("C2").DisplayFormat.Font.Bold and not sheet.Range("C3").DisplayFormat.Font.Bold
