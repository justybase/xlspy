"""Opt-in checks of written workbooks in the installed Microsoft Excel (COM).

Set XLSPY_RUN_EXCEL_COM=1 to run. Excel opens each file and the tests read
back what Excel itself sees: computed formulas, freeze panes, tables, charts,
pictures, print settings, row and column groups, and cell text at the length
limit and with characters that need escaping. Only the installed Excel version
is covered.
"""

import io
import os

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)

from xlspy import ExcelTable, Group, XlsbWriter, XlsxWriter  # noqa: E402

from test_xlsx_integration import _png, build_report  # noqa: E402

XL_VALIDATE_WHOLE_NUMBER = 1
XL_VALIDATE_LIST = 3
XL_LANDSCAPE = 2
MSO_PICTURE = 13
SPECIAL = ["bell\x07", "crlf\r\nend", "lf\nonly", "tab\tx", "lit _x0041_ text", "_x005F_",
           "ctl\x01\x1f", "ff￾", "a&b<c>d"]
AT_LIMIT = ["a" * 32767, "\U0001D11E" * 16383 + "a", "ż" * 32767, "é" * 16383]


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


def _excel_len(sheet, row, col):
    """Excel's own LEN() of a cell (counts UTF-16 code units)."""
    probe = sheet.Cells(40, 40)
    probe.Formula = f"=LEN({sheet.Cells(row, col).Address})"
    return probe.Value


def test_combined_report_in_excel(tmp_path, open_workbook):
    path = build_report(tmp_path / "report.xlsx", _png(tmp_path / "logo.png"))
    workbook = open_workbook(path)
    assert [workbook.Worksheets(i).Name for i in range(1, 4)] == ["Summary", "Q1 Sales", "Data"]
    sheet = workbook.Worksheets("Summary")

    assert sheet.Range("D3").Value == 1500           # B3*C3 = 120 * 12.5
    assert sheet.Range("B7").Value == 545            # SUM(B3:B6)
    assert sheet.Range("D7").Value == pytest.approx(3207)
    assert sheet.Range("B9").Value == 31             # SUM('Q1 Sales'!B2:B4)
    assert sheet.Range("B10").Value == 46023         # Data!A2, 2026-01-01 as a serial

    assert sheet.Range("A1").MergeArea.Address == "$A$1:$F$1"
    assert sheet.Hyperlinks.Count == 4
    assert sheet.Range("F3").Hyperlinks(1).Address == "https://example.com/coffee"
    assert sheet.Range("E3").Validation.Type == XL_VALIDATE_LIST
    assert sheet.Range("B3").Validation.Type == XL_VALIDATE_WHOLE_NUMBER
    assert sheet.ListObjects.Count == 1 and sheet.ListObjects(1).Name == "ProductSales"
    assert sheet.Range("D3:D6").FormatConditions.Count >= 1
    assert sheet.ChartObjects().Count == 3
    assert sum(1 for i in range(1, sheet.Shapes.Count + 1) if sheet.Shapes(i).Type == MSO_PICTURE) == 1

    assert sheet.Names("Print_Area").RefersTo == "=Summary!$A$1:$F$7"

    sheet.Activate()
    window = workbook.Windows(1)
    assert window.FreezePanes
    assert (window.SplitRow, window.SplitColumn) == (2, 0)
    assert workbook.Worksheets("Q1 Sales").AutoFilterMode


@pytest.mark.printer
def test_combined_report_page_setup_in_excel(tmp_path, open_workbook):
    path = build_report(tmp_path / "report.xlsx", _png(tmp_path / "logo.png"))
    setup = open_workbook(path).Worksheets("Summary").PageSetup
    assert setup.PrintArea == "$A$1:$F$7"
    assert setup.PrintTitleRows == "$2:$2"
    assert setup.Orientation == XL_LANDSCAPE
    assert setup.PaperSize == 9  # xlPaperA4


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_freeze_panes_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"frozen.{ext}"
    rows = [[f"c{j}" for j in range(6)]] + [[i * j for j in range(6)] for i in range(1, 30)]
    with writer_cls(path) as writer:
        writer.add_sheet("Both")
        writer.write_sheet(rows, freeze_panes="B3")
        writer.add_sheet("Columns")
        writer.write_sheet(rows, freeze_panes="C1")
        writer.add_sheet("Default")
        writer.write_sheet(rows)  # the header row is frozen by default
    workbook = open_workbook(path)
    expected = {"Both": (2, 1), "Columns": (0, 2), "Default": (1, 0)}
    for name, split in expected.items():
        workbook.Worksheets(name).Activate()
        window = workbook.Windows(1)
        assert window.FreezePanes, name
        assert (window.SplitRow, window.SplitColumn) == split, name


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_tables_open_with_names_from_their_headers(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"tables.{ext}"
    headers = ["Product", "lit _x0041_", "two\nlines", "a&b"]
    with writer_cls(path) as writer:
        writer.add_sheet("Top")
        writer.write_sheet([headers, ["Tea", 1, 2, 3], ["Cocoa", 4, 5, 6]],
                           tables=[ExcelTable("TopTable", "A1:D3")])
        writer.add_sheet("Late")
        rows = [[f"filler {i}", i] for i in range(150)] + [["Name", "Qty"], ["a", 1], ["b", 2]]
        writer.write_sheet(rows, tables=[ExcelTable("LateTable", "A151:B153", columns=["NAME", "qty"])])
    workbook = open_workbook(path)
    top = workbook.Worksheets("Top").ListObjects(1)
    assert [top.ListColumns(i).Name for i in range(1, 5)] == headers
    late = workbook.Worksheets("Late").ListObjects(1)
    assert late.Range.Address == "$A$151:$B$153"
    assert [late.ListColumns(i).Name for i in (1, 2)] == ["NAME", "qty"]


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_file_object_output_opens_in_excel(tmp_path, open_workbook, writer_cls, ext):
    buffer = io.BytesIO()
    with writer_cls(buffer) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["Name", "Qty"]] + [[f"item {i}", i] for i in range(1, 501)])
    path = tmp_path / f"from_buffer.{ext}"
    path.write_bytes(buffer.getvalue())
    sheet = open_workbook(path).Worksheets("Data")
    assert sheet.Range("A501").Value == "item 500"
    assert sheet.Range("B501").Value == 500


@pytest.mark.parametrize("writer_cls,ext,shared", [
    (XlsxWriter, "xlsx", True), (XlsxWriter, "xlsx", False),
    (XlsbWriter, "xlsb", True), (XlsbWriter, "xlsb", False),
], ids=["xlsx-shared", "xlsx-inline", "xlsb-shared", "xlsb-inline"])
def test_cell_text_reads_back_exactly_in_excel(tmp_path, open_workbook, writer_cls, ext, shared):
    path = tmp_path / f"text.{ext}"
    with writer_cls(path, useSharedStrings=shared) as writer:
        writer.add_sheet("T")
        writer.write_sheet([["header"], SPECIAL, AT_LIMIT])
    sheet = open_workbook(path).Worksheets("T")
    assert [sheet.Cells(2, c + 1).Value for c in range(len(SPECIAL))] == SPECIAL
    assert [sheet.Cells(3, c + 1).Value for c in range(len(AT_LIMIT))] == AT_LIMIT
    # Excel counts UTF-16 units: the astral column is 16,384 code points but 32,767 units.
    assert [_excel_len(sheet, 3, c + 1) for c in range(len(AT_LIMIT))] == [32767, 32767, 32767, 32766]


@pytest.mark.parametrize("writer_cls, ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_row_and_column_groups_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"groups.{ext}"
    rows = [[f"r{r}c{c}" for c in range(10)] for r in range(12)]
    with writer_cls(path) as writer:
        writer.add_sheet("G")
        writer.write_sheet(
            rows,
            row_groups=[Group(2, 7), Group(2, 3), Group(5, 6, collapsed=True)],
            column_groups=[Group("C", "F"), Group("D", "E", collapsed=True)],
        )

    sheet = open_workbook(path).Worksheets("G")
    # Excel counts outline levels from 1 (ungrouped), so our level n is n + 1.
    assert [int(sheet.Rows(r).OutlineLevel) for r in range(1, 10)] == [1, 1, 3, 3, 2, 3, 3, 2, 1]
    assert [r for r in range(1, 13) if sheet.Rows(r).Hidden] == [6, 7]
    assert [int(sheet.Columns(c).OutlineLevel) for c in range(1, 8)] == [1, 1, 2, 3, 3, 2, 1]
    assert [c for c in range(1, 11) if sheet.Columns(c).Hidden] == [4, 5]
    assert sheet.Outline.SummaryRow == 1             # xlSummaryBelow
    assert sheet.Outline.SummaryColumn == -4152      # xlSummaryOnRight
    # Expanding the collapsed groups from Excel shows their rows and columns again.
    sheet.Rows(8).ShowDetail = True
    sheet.Columns(6).ShowDetail = True
    assert not any(sheet.Rows(r).Hidden for r in range(1, 13))
    assert not any(sheet.Columns(c).Hidden for c in range(1, 11))
