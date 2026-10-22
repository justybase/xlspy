"""Opt-in checks in Microsoft Excel (COM): document properties, stacked,
area and doughnut charts, and pictures from memory, GIF and BMP.

Set XLSPY_RUN_EXCEL_COM=1 to run.
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

from xlspy import Chart, ChartSeries, Image, XlsbWriter, XlsxWriter  # noqa: E402

# Excel's XlChartType values.
XL_CHART_TYPES = {
    ("bar", None): 51, ("bar", "stacked"): 52, ("bar", "percent_stacked"): 53,
    ("barh", "stacked"): 58, ("barh", "percent_stacked"): 59,
    ("line", "stacked"): 63, ("line", "percent_stacked"): 64,
    ("area", None): 1, ("area", "stacked"): 76, ("area", "percent_stacked"): 77,
    ("doughnut", None): -4120,
}
ROWS = [["Month", "North", "South"], ["Jan", 10, 20], ["Feb", 15, 25], ["Mar", 12, 30]]
SERIES = [
    ChartSeries(values="B2:B4", categories="A2:A4", title_cell="B1"),
    ChartSeries(values="C2:C4", categories="A2:A4", title_cell="C1"),
]


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


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_document_properties_in_excel(tmp_path, open_workbook, writer_cls, ext):
    path = tmp_path / f"props.{ext}"
    with writer_cls(path) as writer:
        writer.set_properties(
            title="Report Q3 & <test>", subject="Sales", author="Zoë Ångström",
            manager="Boss", company="ACME", category="Finance", keywords="q3, sales",
            comments="Report description", status="Draft",
            created=datetime.datetime(2024, 5, 1, 10, 0, tzinfo=datetime.timezone.utc),
            custom={"Project": "Alpha", "Count": 7, "Ratio": 0.25, "Approved": True,
                    "Due": datetime.datetime(2024, 6, 1, 12, 0)},
        )
        writer.add_sheet("Data")
        writer.write_sheet(ROWS)
    workbook = open_workbook(path)
    builtin = workbook.BuiltinDocumentProperties
    expected = {
        "Title": "Report Q3 & <test>", "Subject": "Sales", "Author": "Zoë Ångström",
        "Manager": "Boss", "Company": "ACME", "Category": "Finance",
        "Keywords": "q3, sales", "Comments": "Report description", "Content status": "Draft",
    }
    for name, value in expected.items():
        assert builtin(name).Value == value, name
    created = builtin("Creation date").Value
    assert (created.year, created.month, created.day) == (2024, 5, 1)
    custom = workbook.CustomDocumentProperties
    assert custom.Count == 5
    assert custom("Project").Value == "Alpha"
    assert custom("Count").Value == 7
    assert custom("Ratio").Value == pytest.approx(0.25)
    assert custom("Approved").Value is True
    due = custom("Due").Value
    assert (due.year, due.month, due.day) == (2024, 6, 1)


def test_stacked_area_and_doughnut_charts_in_excel(tmp_path, open_workbook):
    charts = []
    for index, (chart_type, grouping) in enumerate(XL_CHART_TYPES):
        series = SERIES[:1] if chart_type == "doughnut" else SERIES
        charts.append(Chart(
            chart_type, series, title=f"{chart_type} {grouping}", grouping=grouping,
            anchor=f"E{2 + index * 16}",
        ))
    path = tmp_path / "charts.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS, charts=charts)
    sheet = open_workbook(path).Worksheets("Data")
    assert sheet.ChartObjects().Count == len(XL_CHART_TYPES)
    for index, key in enumerate(XL_CHART_TYPES, start=1):
        chart = sheet.ChartObjects(index).Chart
        assert chart.ChartType == XL_CHART_TYPES[key], key
        expected_series = 1 if key[0] == "doughnut" else 2
        assert chart.SeriesCollection().Count == expected_series, key
        assert tuple(chart.SeriesCollection(1).Values) == (10, 15, 12), key
        assert tuple(chart.SeriesCollection(1).XValues) == ("Jan", "Feb", "Mar"), key


def _png(width, height):
    raw = b"".join(b"\x00" + b"\x00\x80\x00" * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _bmp(width, height):
    row = b"\xff\x00\x00" * width
    row += b"\x00" * (-len(row) % 4)
    pixels = row * height
    return (b"BM" + struct.pack("<IHHI", 54 + len(pixels), 0, 0, 54)
            + struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, len(pixels), 2835, 2835, 0, 0)
            + pixels)


GIF_1X1 = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\x00\x00\x00\x00\x00"
    b"!\xf9\x04\x00\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)


def test_pictures_from_memory_gif_and_bmp_in_excel(tmp_path, open_workbook):
    images = [
        Image(_png(40, 20), anchor="B2"),
        Image(io.BytesIO(_png(40, 20)), anchor="B8", width=80),
        Image(GIF_1X1, anchor="F2", width=24),
        Image(_bmp(30, 12), anchor="F8"),
    ]
    path = tmp_path / "pictures.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Pictures")
        writer.write_sheet([["x"]], images=images)
    sheet = open_workbook(path).Worksheets("Pictures")
    shapes = sheet.Shapes
    assert shapes.Count == 4
    # Shape sizes are in points: 0.75 point per pixel.
    sizes = sorted((round(shapes(i).Width / 0.75), round(shapes(i).Height / 0.75))
                   for i in range(1, 5))
    assert sizes == sorted([(40, 20), (80, 40), (24, 24), (30, 12)])
    corners = sorted(shapes(i).TopLeftCell.Address.replace("$", "") for i in range(1, 5))
    assert corners == ["B2", "B8", "F2", "F8"]
