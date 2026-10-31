"""Chart formatting: gap width, doughnut hole, chart and plot area fills, and default text."""

import os
import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import Chart, ChartSeries, XlsbWriter, XlsxWriter

C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
ROWS = [["Region", "Sales"], ["North", 10], ["South", 7], ["East", 5]]


def _chart_root(path, index=1):
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read(f"xl/charts/chart{index}.xml"))


def _write(path, chart, writer_class=XlsxWriter):
    with writer_class(path) as writer:
        if writer_class is XlsbWriter:
            writer.add_sheet("Sheet1")
        writer.write_sheet(ROWS, charts=[chart])
    return path


def _bar(**kwargs):
    return Chart("bar", [ChartSeries("B2:B4", categories="A2:A4")], **kwargs)


def _doughnut(**kwargs):
    return Chart("doughnut", [ChartSeries("B2:B4", categories="A2:A4")], **kwargs)


def test_default_bar_chart_keeps_gap_width_150(tmp_path):
    root = _chart_root(_write(tmp_path / "bar.xlsx", _bar()))
    assert root.find(f".//{C_NS}gapWidth").get("val") == "150"


def test_gap_width_is_written_to_every_bar_group(tmp_path):
    chart = Chart("bar", [
        ChartSeries("B2:B4", categories="A2:A4"),
        ChartSeries("C2:C4", categories="A2:A4", chart_type="line"),
    ], gap_width=40)
    root = _chart_root(_write(tmp_path / "combo.xlsx", chart))
    assert [node.get("val") for node in root.iter(f"{C_NS}gapWidth")] == ["40"]


def test_hole_size_is_written_to_doughnut(tmp_path):
    root = _chart_root(_write(tmp_path / "donut.xlsx", _doughnut(hole_size=65)))
    assert root.find(f".//{C_NS}holeSize").get("val") == "65"


def test_default_doughnut_keeps_hole_50(tmp_path):
    root = _chart_root(_write(tmp_path / "donut.xlsx", _doughnut()))
    assert root.find(f".//{C_NS}holeSize").get("val") == "50"


def test_chart_space_formatting_follows_the_chart_element(tmp_path):
    chart = _bar(fill="#FFFFFF", border="#D9D9D9", font_size=9, font_color="#595959")
    root = _chart_root(_write(tmp_path / "styled.xlsx", chart))
    order = [child.tag.removeprefix(C_NS) for child in root]
    assert order.index("chart") < order.index("spPr") < order.index("txPr")
    space_fill = root.find(f"{C_NS}spPr/{A_NS}solidFill/{A_NS}srgbClr")
    assert space_fill.get("val") == "FFFFFF"
    assert root.find(f"{C_NS}spPr/{A_NS}ln/{A_NS}solidFill/{A_NS}srgbClr").get("val") == "D9D9D9"
    default_run = root.find(f"{C_NS}txPr/{A_NS}p/{A_NS}pPr/{A_NS}defRPr")
    assert default_run.get("sz") == "900"
    assert default_run.find(f"{A_NS}solidFill/{A_NS}srgbClr").get("val") == "595959"


def test_plot_area_fill_is_inside_plot_area(tmp_path):
    root = _chart_root(_write(tmp_path / "plot.xlsx", _bar(plot_fill="#F7F9FC")))
    plot_area = root.find(f"{C_NS}chart/{C_NS}plotArea")
    last = list(plot_area)[-1]
    assert last.tag == f"{C_NS}spPr"
    assert last.find(f"{A_NS}solidFill/{A_NS}srgbClr").get("val") == "F7F9FC"


def test_chart_without_formatting_has_no_extra_parts(tmp_path):
    root = _chart_root(_write(tmp_path / "plain.xlsx", _bar()))
    assert root.find(f"{C_NS}spPr") is None
    assert root.find(f"{C_NS}txPr") is None


def test_formatting_is_written_to_xlsb_too(tmp_path):
    path = tmp_path / "styled.xlsb"
    _write(path, _bar(gap_width=60, fill="#FFFFFF"), writer_class=XlsbWriter)
    with zipfile.ZipFile(path) as archive:
        chart = ET.fromstring(archive.read("xl/charts/chart1.xml"))
    assert chart.find(f".//{C_NS}gapWidth").get("val") == "60"
    assert chart.find(f"{C_NS}spPr") is not None


@pytest.mark.parametrize(
    "build, message",
    [
        (lambda: Chart("pie", [ChartSeries("B2:B4")], gap_width=50), "gap_width"),
        (lambda: _bar(gap_width=600), "gap_width"),
        (lambda: _bar(hole_size=50), "hole_size"),
        (lambda: _doughnut(hole_size=5), "hole_size"),
        (lambda: _bar(fill="white"), "fill"),
        (lambda: _bar(border="#12345"), "border"),
        (lambda: _bar(plot_fill="#GGGGGG"), "plot_fill"),
        (lambda: _bar(font_color="red"), "font_color"),
        (lambda: _bar(font_size=0), "font_size"),
    ],
)
def test_invalid_formatting_is_rejected(build, message):
    with pytest.raises(ValueError, match=message):
        build()


@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
                    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation")
def test_excel_reads_gap_width_hole_size_and_chart_fill(tmp_path):
    win32 = pytest.importorskip("win32com.client")
    bar_path = _write(tmp_path / "bar_com.xlsx", _bar(gap_width=40, fill="#FFFFFF"))
    donut_path = _write(tmp_path / "donut_com.xlsx", _doughnut(hole_size=65))
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(bar_path))
        chart = book.Worksheets(1).ChartObjects(1).Chart
        assert chart.ChartGroups(1).GapWidth == 40
        assert chart.ChartArea.Format.Fill.ForeColor.RGB == 0xFFFFFF
        book.Close(SaveChanges=False)
        book = excel.Workbooks.Open(str(donut_path))
        chart = book.Worksheets(1).ChartObjects(1).Chart
        assert chart.ChartGroups(1).DoughnutHoleSize == 65
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()
