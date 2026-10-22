"""Opt-in checks in Microsoft Excel (COM): sheet tab colours and a filter on
the last row of a two-row header. Set XLSPY_RUN_EXCEL_COM=1 to run.
"""

import os

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)

from xlspy import Color, Style, XlsbWriter, XlsxWriter  # noqa: E402

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]
HEADER = Style(fill="#1F4E78", font_color="#FFFFFF")
ROWS = [
    ["Region", None, "Sales", None],
    ["Name", "Code", "Q1", "Q2"],
    ["North", "N", 1, 2],
    ["South", "S", 3, 4],
    ["East", "E", 5, 6],
]


def _bgr(rgb: str) -> int:
    return int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)


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


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_tab_color_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"tabs.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Red")
        writer.write_sheet(ROWS, tab_color="#C00000")
        writer.add_sheet("Theme")
        writer.write_sheet(ROWS, tab_color=Color(theme=5))
        writer.add_sheet("Plain")
        writer.write_sheet(ROWS)
    workbook = open_workbook(path)
    assert workbook.Worksheets("Red").Tab.Color == _bgr("C00000")
    # Theme colour 5 is Accent 2: xlThemeColorAccent2 = 6.
    assert workbook.Worksheets("Theme").Tab.ThemeColor == 6
    assert workbook.Worksheets("Plain").Tab.Color is False  # no colour


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_filter_on_a_two_row_header_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"header.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS, header_rows=2, header_style=HEADER,
                           merges=["A1:B1", "C1:D1"], freeze_panes="A3")
    sheet = open_workbook(path).Worksheets("Data")
    assert sheet.AutoFilterMode
    assert sheet.AutoFilter.Range.Address == "$A$2:$D$5"
    for ref in ("A1", "C1", "A2", "D2"):
        cell = sheet.Range(ref)
        assert cell.Interior.Color == _bgr("1F4E78"), ref
        assert cell.Font.Bold and cell.Font.Color == _bgr("FFFFFF"), ref
    assert not sheet.Range("A3").Font.Bold
    # Filtering acts on the data under the second header row.
    sheet.Range("A2:D5").AutoFilter(1, "South")
    visible = [row for row in range(3, 6) if not sheet.Rows(row).Hidden]
    assert visible == [4]
