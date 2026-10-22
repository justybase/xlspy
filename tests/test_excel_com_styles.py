"""Opt-in checks in Microsoft Excel (COM) of styled sheets and grid lines.

Set XLSPY_RUN_EXCEL_COM=1 to run. The sheets cover the style paths the
writers take: overlapping ranges merged per cell, ``(value, format)`` cells
inside styled ranges, one style range per cell, the Workbook API's per-cell
styles, and ``show_grid_lines``. Excel opens each file and the tests read
back the formatting it applies.
"""

import datetime
import os

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)

from xlspy import (  # noqa: E402
    Alignment, Border, Font, PatternFill, Side, Style, Workbook, XlsbWriter, XlsxWriter,
)

XL_EDGE_LEFT = 7
XL_CONTINUOUS = 1
XL_LINE_NONE = -4142
XL_CENTER = -4108
XL_GENERAL = 1
WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]

HEADER = Style(fill="1F4E78", font_color="FFFFFF")
BOXED = Style(border=Border(left=Side(style="thin")))
MONEY = Style(number_format="#,##0.00")
BOLD = Style(font=Font(bold=True))
CENTERED = Style(alignment=Alignment(horizontal="center"))
ZEBRA = Style(pattern=PatternFill(fill_type="solid", start_color="DDEBF7"))


def _bgr(rgb: str) -> int:
    """Excel's ``Color`` value (0xBBGGRR) of an RRGGBB string."""
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


def _rows(n, cols=4):
    return [[f"r{r}c{c}" if c == 0 else r * 10 + c for c in range(cols)] for r in range(n)]


def _has_left_border(cell):
    return cell.Borders(XL_EDGE_LEFT).LineStyle == XL_CONTINUOUS


XL_DECIMAL_SEPARATOR = 3
XL_THOUSANDS_SEPARATOR = 4


def _number_format(cell, code):
    """Whether ``cell`` has format ``code``. Excel reports formats with the
    locale's separators (``# ##0,00`` in Poland), so ``code`` is translated;
    ``General`` is compared with an untouched cell's format."""
    if code == "General":
        return cell.NumberFormat == cell.Worksheet.Range("Z1000").NumberFormat
    # Late-bound COM returns International as one tuple (1-based in VBA).
    international = cell.Application.International
    decimal = international[XL_DECIMAL_SEPARATOR - 1]
    thousands = international[XL_THOUSANDS_SEPARATOR - 1]
    local = "".join({",": thousands, ".": decimal}.get(char, char) for char in code)
    return cell.NumberFormat in (code, local)


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_overlapping_ranges_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"overlap.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_rows(25), header_style=HEADER, styles={
            "A1:D20": BOXED,
            "B2:C20": MONEY,
            "C2:C10": BOLD,
            "A5:D6": CENTERED,
        })
    sheet = open_workbook(path).Worksheets("Data")

    header = sheet.Range("A1")
    assert header.Interior.Color == _bgr("1F4E78")
    assert header.Font.Bold and header.Font.Color == _bgr("FFFFFF")
    assert _has_left_border(header)

    assert _number_format(sheet.Range("C3"), "#,##0.00")
    assert sheet.Range("C3").Font.Bold and _has_left_border(sheet.Range("C3"))
    assert _number_format(sheet.Range("C12"), "#,##0.00")
    assert not sheet.Range("C12").Font.Bold
    assert sheet.Range("B5").HorizontalAlignment == XL_CENTER
    assert _number_format(sheet.Range("B5"), "#,##0.00")
    assert sheet.Range("D5").HorizontalAlignment == XL_CENTER
    assert sheet.Range("D7").HorizontalAlignment == XL_GENERAL
    assert _has_left_border(sheet.Range("D20"))
    assert sheet.Range("D21").Borders(XL_EDGE_LEFT).LineStyle == XL_LINE_NONE
    assert _number_format(sheet.Range("B21"), "General")
    assert sheet.Range("C3").Value == 22


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_formats_inside_styled_ranges_in_excel(tmp_path, open_workbook, writer_cls, ext):
    rows = [["Name", "Amount", "Share"]]
    rows += [[f"n{r}", (r * 1.25, "0.000"), (r / 100, "0%")] for r in range(1, 40)]
    path = tmp_path / f"formats.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        # Rows 2..20 are boxed; the rest carry only their own formats.
        writer.write_sheet(rows, header_style=HEADER, styles={"A2:C20": BOXED, "C2:C30": BOLD})
    sheet = open_workbook(path).Worksheets("Data")

    for address, number_format, boxed, bold in [
        ("B2", "0.000", True, False),
        ("C2", "0%", True, True),
        ("B25", "0.000", False, False),
        ("C25", "0%", False, True),
        ("C35", "0%", False, False),
    ]:
        cell = sheet.Range(address)
        assert _number_format(cell, number_format), address
        assert _has_left_border(cell) is boxed, address
        assert bool(cell.Font.Bold) is bold, address
    assert sheet.Range("B3").Value == pytest.approx(2.5)
    assert sheet.Range("A1").Font.Bold


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_one_style_range_per_cell_in_excel(tmp_path, open_workbook, writer_cls, ext):
    fills = ["DDEBF7", "FCE4D6", "E2EFDA"]
    palette = [Style(fill=color) for color in fills]
    styles = {}
    for r in range(300):
        for c, letter in enumerate("ABCD"):
            if (r + c) % 2 == 0:
                styles[f"{letter}{r + 1}"] = palette[(r + c) % 3]
    path = tmp_path / f"cells.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_rows(300), styles=styles)
    sheet = open_workbook(path).Worksheets("Data")

    for r in (0, 1, 2, 149, 298, 299):
        for c, letter in enumerate("ABCD"):
            cell = sheet.Range(f"{letter}{r + 1}")
            if (r + c) % 2 == 0:
                assert cell.Interior.Color == _bgr(fills[(r + c) % 3]), cell.Address
            else:
                assert cell.Interior.ColorIndex == -4142, cell.Address  # xlColorIndexNone


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_workbook_api_cell_styles_in_excel(tmp_path, open_workbook, ext):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in _rows(200):
        sheet.append(row)
    for r, row in enumerate(sheet.iter_rows(), start=1):
        for c, cell in enumerate(row):
            if c % 2 == 1:
                cell.font = Font(bold=True)
                if r % 2:
                    cell.number_format = "#,##0.00"
    path = tmp_path / f"api.{ext}"
    workbook.save(path)
    sheet = open_workbook(path).Worksheets("Data")

    assert sheet.Range("B1").Font.Bold and _number_format(sheet.Range("B1"), "#,##0.00")
    assert sheet.Range("B2").Font.Bold and _number_format(sheet.Range("B2"), "General")
    assert not sheet.Range("C1").Font.Bold
    assert sheet.Range("D199").Font.Bold and _number_format(sheet.Range("D199"), "#,##0.00")
    assert sheet.Range("D199").Value == 1983


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_show_grid_lines_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"grid.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("NoGrid")
        writer.write_sheet([["a", 1]], show_grid_lines=False)
        writer.add_sheet("Grid")
        writer.write_sheet([["b", 2]])
        writer.add_sheet("FrozenNoGrid")
        writer.write_sheet([["c", 3], ["d", 4]], freeze_panes="B2", show_grid_lines=False)
    workbook = open_workbook(path)
    window = workbook.Windows(1)
    for name, shown in (("NoGrid", False), ("Grid", True), ("FrozenNoGrid", False)):
        workbook.Worksheets(name).Activate()
        assert window.DisplayGridlines is shown, name
    assert window.FreezePanes  # FrozenNoGrid keeps its frozen pane


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_show_grid_lines_from_workbook_api_in_excel(tmp_path, open_workbook, ext):
    workbook = Workbook()
    workbook.active.title = "NoGrid"
    workbook.active.append(["x"])
    workbook.active.sheet_view.showGridLines = False
    workbook.create_sheet("Grid").append(["y"])
    path = tmp_path / f"api_grid.{ext}"
    workbook.save(path)
    opened = open_workbook(path)
    window = opened.Windows(1)
    for name, shown in (("NoGrid", False), ("Grid", True)):
        opened.Worksheets(name).Activate()
        assert window.DisplayGridlines is shown, name


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_border_edges_in_excel(tmp_path, open_workbook, writer_cls, ext):
    """Every edge set alone reaches Excel; ``start`` is the left edge of a
    left-to-right sheet. An unset ``start`` must not hide ``left``."""
    thin = Side(style="thin")
    edges = {"left": 7, "top": 8, "bottom": 9, "right": 10}
    styles = {f"{letter}2": Style(border=Border(**{edge: thin}))
              for letter, edge in zip("ABCD", edges)}
    if ext == "xlsx":  # XLSB has no start/end edges
        styles["E2"] = Style(border=Border(start=thin))
    styles["F2"] = Style(border=Border(left=thin, right=thin, top=thin, bottom=thin))
    path = tmp_path / f"borders.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["x"] * 6, [1, 2, 3, 4, 5, 6]], styles=styles)
    sheet = open_workbook(path).Worksheets("Data")
    for letter, (edge, index) in zip("ABCD", edges.items()):
        cell = sheet.Range(f"{letter}2")
        assert cell.Borders(index).LineStyle == XL_CONTINUOUS, edge
        for other, other_index in edges.items():
            if other != edge:
                assert cell.Borders(other_index).LineStyle == XL_LINE_NONE, (edge, other)
    if ext == "xlsx":
        assert _has_left_border(sheet.Range("E2"))
    assert all(sheet.Range("F2").Borders(i).LineStyle == XL_CONTINUOUS for i in edges.values())


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_value_style_cells_in_excel(tmp_path, open_workbook, writer_cls, ext):
    """(value, Style) cells go on top of header and range styles; most rows
    take the C encoder once the first row has recorded their xf."""
    money_red = Style(fill="FCE4D6", number_format="#,##0.00")
    rows = [["Name", "Amount", "Flag"]]
    rows += [[f"n{r}", (r * 1.5, money_red), (r % 2 == 0, BOLD)] for r in range(1, 60)]
    path = tmp_path / f"cells.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, header_style=HEADER, styles={"A2:C20": BOXED})
    sheet = open_workbook(path).Worksheets("Data")

    for address, boxed in (("B2", True), ("B20", True), ("B21", False), ("B60", False)):
        cell = sheet.Range(address)
        assert cell.Interior.Color == _bgr("FCE4D6"), address
        assert _number_format(cell, "#,##0.00"), address
        assert _has_left_border(cell) is boxed, address
    assert sheet.Range("C3").Font.Bold and _has_left_border(sheet.Range("C3"))
    assert sheet.Range("C41").Font.Bold and sheet.Range("C41").Value is True  # r = 40
    assert sheet.Range("B60").Value == pytest.approx(88.5)
    assert sheet.Range("A1").Interior.Color == _bgr("1F4E78")


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_open_style_ranges_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"open.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_rows(500), styles={
            "B:B": MONEY, "C3:C": BOLD, "1:1": Style(fill="E2EFDA"),
        })
    sheet = open_workbook(path).Worksheets("Data")
    assert _number_format(sheet.Range("B2"), "#,##0.00")
    assert _number_format(sheet.Range("B500"), "#,##0.00")
    assert sheet.Range("C3").Font.Bold and sheet.Range("C500").Font.Bold
    assert not sheet.Range("C2").Font.Bold
    assert sheet.Range("D1").Interior.Color == _bgr("E2EFDA")
    assert sheet.Range("D2").Interior.ColorIndex == -4142
    # Only written rows are styled: the used range ends with the data.
    assert sheet.UsedRange.Rows.Count == 500


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_styled_rows_with_heights_hidden_and_groups_in_excel(tmp_path, open_workbook, writer_cls, ext):
    from xlspy import Group
    path = tmp_path / f"rows.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [[r, (r * 1.5, "0.00"), (f"t{r}", BOLD)] for r in range(40)],
            styles={"A:A": BOXED},
            row_heights={3: 30, 10: 12.75},
            hidden_rows=[5],
            row_groups=[Group(20, 25)],
        )
    sheet = open_workbook(path).Worksheets("Data")
    # Excel snaps heights to its pixel grid (12.75 can read back as 12.8).
    assert sheet.Rows(4).RowHeight == pytest.approx(30, abs=0.3)
    assert sheet.Rows(11).RowHeight == pytest.approx(12.75, abs=0.3)
    assert sheet.Rows(6).Hidden and not sheet.Rows(7).Hidden
    assert sheet.Rows(22).OutlineLevel == 2 and sheet.Rows(30).OutlineLevel == 1
    for address in ("A4", "A6", "A22", "A40"):
        assert _has_left_border(sheet.Range(address)), address
    for address in ("B4", "B22"):
        assert _number_format(sheet.Range(address), "0.00"), address
    assert sheet.Range("C4").Font.Bold and sheet.Range("C22").Value == "t21"


WIDTH_ROWS = [
    ["Order number", "Customer name and address", "Amount", "Share", "Date", "Long date",
     "Codes", "Mixed", "Flag", "General"],
    ["ORDER-2024-000123", "Wrocław, ul. Świdnicka 12/4", (1234567.891, "#,##0.00"),
     (0.4567, "0.00%"), datetime.date(2024, 9, 30), (datetime.date(2024, 9, 30), "dddd, d mmmm yyyy"),
     "WWWWMMMM@@", "中文字符测试", True, 1 / 3],
    ["ORD-1", "MAXIMUM WIDTH UPPERCASE TEXT", (98765432.1, '#,##0.00 "zł"'), (1.5, "0%"),
     datetime.datetime(2024, 12, 31, 23, 59), (datetime.date(2024, 9, 30), "mmmm yyyy"),
     "QQQQQQQQQQ", "Zażółć gęślą jaźń", False, 1234567890123],
]


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_automatic_widths_are_not_narrower_than_excel_autofit(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"widths.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(WIDTH_ROWS)
    sheet = open_workbook(path).Worksheets("Data")
    ours = [sheet.Columns(i).ColumnWidth for i in range(1, len(WIDTH_ROWS[0]) + 1)]
    # Nothing shows #### (a number or date too wide for its column).
    for row in range(2, len(WIDTH_ROWS) + 1):
        for col in range(1, len(WIDTH_ROWS[0]) + 1):
            assert "#" not in sheet.Cells(row, col).Text, (row, col, sheet.Cells(row, col).Text)
    sheet.Range("A1:J3").Columns.AutoFit()
    fitted = [sheet.Columns(i).ColumnWidth for i in range(1, len(WIDTH_ROWS[0]) + 1)]
    for col, (mine, excel) in enumerate(zip(ours, fitted), start=1):
        assert mine >= excel - 0.05, (col, WIDTH_ROWS[0][col - 1], mine, excel)
        # Not wildly wider either (headers get room for the filter button).
        assert mine <= max(excel, 8.43) + 4.5, (col, WIDTH_ROWS[0][col - 1], mine, excel)
