"""Every chart option in XLSB: the same chart XML as XLSX plus the caches."""

import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import (
    Chart, ChartAxis, ChartLegend, ChartMarker, ChartSeries, DataLabels, XlsbWriter,
)

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

ROWS = [
    ["Quarter", "Sales", "Costs", "X"],
    ["Q1", 10, 4, 1.5],
    ["Q2", 30, 12, 2.5],
    ["Q3", 20, 9, 4.0],
]

TWO_SERIES = [
    ChartSeries("B2:B4", title_cell="B1", categories="A2:A4"),
    ChartSeries("C2:C4", title_cell="C1", categories="A2:A4"),
]

CHARTS = {
    "barh": Chart("barh", TWO_SERIES),
    "area": Chart("area", TWO_SERIES),
    "doughnut": Chart("doughnut", [ChartSeries("B2:B4", categories="A2:A4")]),
    "scatter": Chart("scatter", [
        ChartSeries("B2:B4", x_values="D2:D4", title="Sales",
                    marker=ChartMarker("diamond", 9), line=True, color="FF0000"),
    ]),
    "stacked": Chart("bar", TWO_SERIES, grouping="stacked"),
    "percent": Chart("area", TWO_SERIES, grouping="percent_stacked"),
    "axes": Chart("line", TWO_SERIES, x_axis=ChartAxis(title="Quarter"),
                  y_axis=ChartAxis(minimum=0, maximum=40, major_unit=10, number_format="0.0")),
    "legend": Chart("bar", TWO_SERIES, legend=ChartLegend("bottom")),
    "no_legend": Chart("bar", TWO_SERIES, legend=False),
    "labels": Chart("bar", TWO_SERIES, data_labels=DataLabels(number_format="0", position="outside_end")),
    "colours": Chart("bar", [ChartSeries("B2:B4", color="00B050"), ChartSeries("C2:C4", color="7030A0")]),
    "markers": Chart("line", [ChartSeries("B2:B4", marker=ChartMarker("square")), ChartSeries("C2:C4", line=False,
                                                                                         marker=ChartMarker())]),
    "title_size": Chart("bar", TWO_SERIES, title="Small title", title_size=12),
}


def _chart(tmp_path, chart):
    path = tmp_path / "chart.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS, charts=[chart])
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/charts/chart1.xml").decode()


@pytest.mark.parametrize("name", sorted(CHARTS))
def test_every_reference_carries_a_cache(tmp_path, name):
    xml = _chart(tmp_path, CHARTS[name])
    root = ET.fromstring(xml)
    refs = root.findall(f".//{C}numRef") + root.findall(f".//{C}strRef")
    assert refs
    for ref in refs:
        assert ref.find(f"{C}numCache") is not None or ref.find(f"{C}strCache") is not None, ET.tostring(ref)
    # Excel's XLSB loader needs explicit colours and visible axes.
    for series in root.iter(f"{C}ser"):
        assert "srgbClr" in ET.tostring(series).decode()
    for axis in list(root.iter(f"{C}catAx")) + list(root.iter(f"{C}valAx")):
        assert axis.find(f"{C}delete") is not None


def test_scatter_caches_x_and_y_values(tmp_path):
    xml = _chart(tmp_path, CHARTS["scatter"])
    assert "<xVal><numRef><f>'Data'!$D$2:$D$4</f><numCache>" in xml
    assert "<pt idx=\"0\"><v>1.5</v></pt>" in xml
    assert "<yVal><numRef><f>'Data'!$B$2:$B$4</f><numCache>" in xml
    assert '<marker><symbol val="diamond"/><size val="9"/>' in xml
    assert '<a:srgbClr val="FF0000"/>' in xml


def test_options_reach_the_xml(tmp_path):
    assert '<legendPos val="b"/>' in _chart(tmp_path, CHARTS["legend"])
    assert "<legend>" not in _chart(tmp_path, CHARTS["no_legend"])
    assert '<grouping val="stacked"/><varyColors val="0"/>' in _chart(tmp_path, CHARTS["stacked"])
    assert '<overlap val="100"/>' in _chart(tmp_path, CHARTS["stacked"])
    axes = _chart(tmp_path, CHARTS["axes"])
    assert '<max val="40"/><min val="0"/>' in axes and '<majorUnit val="10"/>' in axes
    assert '<numFmt formatCode="0.0" sourceLinked="0"/>' in axes
    assert "<a:t>Quarter</a:t>" in axes
    assert '<dLblPos val="outEnd"/>' in _chart(tmp_path, CHARTS["labels"])
    colours = _chart(tmp_path, CHARTS["colours"])
    assert '<a:srgbClr val="00B050"/>' in colours and '<a:srgbClr val="7030A0"/>' in colours
    assert 'sz="1200" b="1"' in _chart(tmp_path, CHARTS["title_size"])
    markers = _chart(tmp_path, CHARTS["markers"])
    assert '<marker><symbol val="square"/>' in markers
    assert '<a:ln w="19050"><a:noFill/></a:ln>' in markers
    # The XLSB loader smooths lines and drops automatic titles unless told.
    assert markers.count('<smooth val="0"/></ser>') == 2
    assert '<chart><autoTitleDeleted val="0"/><plotArea>' in markers
    assert "autoTitleDeleted" not in _chart(tmp_path, CHARTS["title_size"])
    doughnut = _chart(tmp_path, CHARTS["doughnut"])
    assert doughnut.count("<dPt>") == 3 and '<holeSize val="50"/>' in doughnut
