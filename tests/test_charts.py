import zipfile

import pytest

from xlspy import Chart, ChartSeries, ExcelReader, XlsbWriter, XlsxWriter
from helpers import part_text
from xlspy import Trendline
from xlspy import Workbook
from xml.etree import ElementTree as ET
import re


def _chart_parts(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert "xl/charts/chart1.xml" in names
        return {
            name: archive.read(name).decode()
            for name in names
            if "/drawings/" in name or "/charts/" in name
        }


def test_chart_model_rejects_invalid_input():
    with pytest.raises(ValueError, match="invalid chart type"):
        Chart("funnel", [ChartSeries("B2:B4")])
    with pytest.raises(ValueError, match="at least one series"):
        Chart("bar", [])
    with pytest.raises(TypeError, match="ChartSeries"):
        Chart("bar", ["B2:B4"])
    with pytest.raises(ValueError, match="single cell"):
        Chart("bar", [ChartSeries("B2:B4")], anchor="A1:B2")
    with pytest.raises(ValueError, match="must be positive"):
        Chart("bar", [ChartSeries("B2:B4")], width=0)
    with pytest.raises(TypeError, match="must be a number"):
        Chart("bar", [ChartSeries("B2:B4")], height="tall")
    with pytest.raises(ValueError, match="between 1 and 48"):
        Chart("bar", [ChartSeries("B2:B4")], style=99)
    with pytest.raises(ValueError, match="mutually exclusive"):
        ChartSeries("B2:B4", title="A", title_cell="A1")
    with pytest.raises(ValueError, match="single cell"):
        ChartSeries("B2:B4", title_cell="A1:A2")
    with pytest.raises(ValueError, match="invalid A1 range"):
        ChartSeries("B2:B4", categories="nope")
    with pytest.raises(ValueError, match="chart sheet"):
        ChartSeries("B2:B4", categories="Other!A1:A3")


def test_charts_must_be_chart_instances(tmp_path):
    with XlsxWriter(tmp_path / "bad.xlsx") as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError, match="Chart instances"):
            writer.write_sheet([["x"]], charts=["B2:B4"])


def _sample_charts():
    return [
        Chart(
            "bar",
            [
                ChartSeries("B2:B4", title_cell="B1", categories="A2:A4"),
                ChartSeries("C2:C4", title="Lit", categories="A2:A4"),
            ],
            title="Sales & <Costs>",
            anchor="E2",
            style=10,
        ),
        Chart(
            "line",
            [ChartSeries("B2:B4", title_cell="B1", categories="A2:A4")],
            title="Trend",
            anchor="E18",
        ),
        Chart(
            "pie",
            [ChartSeries("B2:B4", title_cell="B1")],
            title="Share",
            anchor="L2",
        ),
    ]


def _sample_data():
    return [
        ["Q", "A", "B"],
        ["Q1", 10, 20],
        ["Q2", 30, 40],
        ["Q3", 50, 60],
    ]


def test_charts_round_trip_through_openpyxl(tmp_path):
    pytest.importorskip("openpyxl")
    import openpyxl

    path = tmp_path / "charts.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts())
        writer.add_sheet("Other")
        writer.write_sheet([["x", 1]], charts=[Chart("bar", [ChartSeries("B1:B1")], anchor="C3")])

    wb = openpyxl.load_workbook(path)
    by_title = {}
    for ws in wb.worksheets:
        for chart in ws._charts:
            # A single-series chart without a title keeps Excel's empty automatic title.
            title = (chart.title.tx.rich.p[0].r[0].t if chart.title else None) or None
            by_title[(ws.title, title)] = chart
    assert set(by_title) == {
        ("Data", "Sales & <Costs>"), ("Data", "Trend"), ("Data", "Share"), ("Other", None),
    }

    bar = by_title[("Data", "Sales & <Costs>")]
    assert type(bar).__name__ == "BarChart"
    assert [s.val.numRef.f for s in bar.series] == ["'Data'!$B$2:$B$4", "'Data'!$C$2:$C$4"]
    assert bar.series[0].tx.strRef.f == "'Data'!B1"
    assert bar.series[1].tx.v == "Lit"
    assert bar.series[0].cat.numRef.f == "'Data'!$A$2:$A$4"

    line = by_title[("Data", "Trend")]
    assert type(line).__name__ == "LineChart"

    pie = by_title[("Data", "Share")]
    assert type(pie).__name__ == "PieChart"
    assert pie.series[0].cat is None

    anchor = bar.anchor._from
    assert (anchor.col, anchor.row) == (4, 1)
    assert (bar.width, bar.height) == (15, 7.5)

    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data"))[1] == ["Q1", 10, 20]


def test_chart_xml_matches_openpyxl_shape(tmp_path):
    path = tmp_path / "shape.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(_sample_data(), charts=_sample_charts()[:1])

    parts = _chart_parts(path)
    assert set(parts) == {
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
    }
    drawing = parts["xl/drawings/drawing1.xml"]
    assert '<c:chart r:id="rId1"/>' in drawing
    assert "<from><col>4</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>" in drawing
    assert '<ext cx="5400000" cy="2700000"/>' in drawing

    chart = parts["xl/charts/chart1.xml"]
    assert chart.startswith("<chartSpace ")
    assert "<style val=\"10\"/>" in chart
    assert "<a:t>Sales &amp; &lt;Costs&gt;</a:t>" in chart
    assert chart.count("<ser>") == 2
    assert "<f>'Data'!$B$2:$B$4</f>" in chart
    assert "<f>'Data'!B1</f>" in chart
    assert "<tx><v>Lit</v></tx>" in chart
    assert "<barDir val=\"col\"/>" in chart
    assert "<gapWidth val=\"150\"/>" in chart
    # An absent overlay element would let Excel draw the legend over the plot.
    assert "<legend><legendPos val=\"r\"/><overlay val=\"0\"/></legend>" in chart

    with zipfile.ZipFile(path) as archive:
        content_types = archive.read("[Content_Types].xml").decode()
        assert "drawing1.xml" in content_types and "chart1.xml" in content_types
        assert "/xl/theme/theme1.xml" in content_types
        assert archive.read("xl/theme/theme1.xml").startswith(b"<?xml")
        sheet_rels = archive.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        assert 'Target="../drawings/drawing1.xml"' in sheet_rels
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode()
        assert '<drawing r:id="rId1"/>' in sheet_xml
        workbook_rels = archive.read("xl/_rels/workbook.xml.rels").decode()
        assert 'Target="theme/theme1.xml"' in workbook_rels


def test_pie_chart_keeps_its_categories(tmp_path):
    path = tmp_path / "pie.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            _sample_data(),
            charts=[Chart("pie", [ChartSeries("B2:B4", categories="A2:A4")], anchor="E2")],
        )

    chart = _chart_parts(path)["xl/charts/chart1.xml"]
    # The categories name the slices in the legend.
    assert "<cat><numRef><f>'Data'!$A$2:$A$4</f></numRef></cat>" in chart


def test_drawing_element_follows_tables_and_plain_sheets(tmp_path):
    path = tmp_path / "mixed.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Rich")
        writer.write_sheet(
            _sample_data(),
            tables=[__import__("xlspy").ExcelTable("Tbl", "A1:C4")],
            charts=_sample_charts()[:1],
        )
        writer.add_sheet("Plain")
        writer.write_sheet(
            [["x", 1]],
            charts=[Chart("bar", [ChartSeries("B1:B1")], anchor="C3")],
        )

    with zipfile.ZipFile(path) as archive:
        rich_xml = archive.read("xl/worksheets/sheet1.xml").decode()
        assert '<drawing r:id="rId2"/>' in rich_xml
        assert rich_xml.index("<drawing") < rich_xml.index("<tableParts")
        assert '<tablePart r:id="rId1"/>' in rich_xml
        rich_rels = archive.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        assert 'Target="../tables/table1.xml"' in rich_rels
        assert 'Target="../drawings/drawing1.xml"' in rich_rels
        plain_xml = archive.read("xl/worksheets/sheet2.xml").decode()
        assert '<drawing r:id="rId1"/>' in plain_xml


def test_sheet_names_with_spaces_are_quoted(tmp_path):
    pytest.importorskip("openpyxl")
    import openpyxl

    path = tmp_path / "quoted.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sales Data")
        writer.write_sheet(
            [["Q", "A"], ["Q1", 10]],
            charts=[Chart("bar", [ChartSeries("B2:B2", title_cell="B1")], anchor="D2")],
        )
    wb = openpyxl.load_workbook(path)
    chart = wb["Sales Data"]._charts[0]
    assert chart.series[0].val.numRef.f == "'Sales Data'!$B$2:$B$2"


def test_plain_write_stays_free_of_drawing_parts(tmp_path):
    path = tmp_path / "plain.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["h1", "h2"]] + [[1, "a"]] * 10)

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        assert not [name for name in names if "/drawings/" in name or "/charts/" in name]
        assert "xl/worksheets/_rels" not in "|".join(names)
        assert "drawing" not in archive.read("[Content_Types].xml").decode()
        assert "theme" not in archive.read("[Content_Types].xml").decode()
        assert "theme" not in archive.read("xl/_rels/workbook.xml.rels").decode()


@pytest.mark.parametrize("writer_cls", [XlsxWriter, XlsbWriter])
def test_untitled_single_series_chart_keeps_the_automatic_title(tmp_path, writer_cls):
    path = tmp_path / f"auto.{'xlsb' if writer_cls is XlsbWriter else 'xlsx'}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["k", "a", "b"], ["x", 1, 2], ["y", 3, 4]],
            charts=[Chart("bar", [ChartSeries("B2:B3", title_cell="B1")], anchor="E2"),
                    Chart("bar", [ChartSeries("B2:B3"), ChartSeries("C2:C3")], anchor="E20")],
        )
    with zipfile.ZipFile(path) as archive:
        single, double = (archive.read(f"xl/charts/chart{i}.xml").decode() for i in (1, 2))
    # As Excel saves it: an empty title, shown as the series name.
    assert '<chart><title><overlay val="0"/></title><autoTitleDeleted val="0"/>' in single
    assert "<title>" not in double and '<autoTitleDeleted val="0"/>' in double


def test_radar_chart_is_written_as_radar_not_pie(tmp_path):
    path = tmp_path / "radar.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    for row in range(1, 6):
        sheet.cell(row=row, column=2, value=row)
    sheet.add_chart(Chart("radar", [ChartSeries("B1:B5")]), "E2")
    workbook.save(path)
    workbook.close()

    chart = part_text(path, "xl/charts/chart1.xml")
    assert "<radarChart>" in chart
    assert "<radarStyle" in chart
    assert "pieChart" not in chart


@pytest.mark.parametrize("chart_type", [
    "area3d", "bar3d", "bubble", "line3d", "pie3d",
    "projected_pie", "stock", "surface", "surface3d",
])
def test_unsupported_chart_types_raise_instead_of_becoming_pie(chart_type):
    with pytest.raises(ValueError):
        Chart(chart_type, [ChartSeries("B1:B5")])


CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def parts(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def xml_root(data: bytes):
    return ET.fromstring(data)


def _chart_formulas(root):
    return [node.text for node in root.iter(f"{CHART}f")]


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_chart_formula_escapes_sheet_name(tmp_path, ext):
    writer = XlsxWriter if ext == "xlsx" else XlsbWriter
    path = tmp_path / f"amp.{ext}"
    with writer(path) as wb:
        wb.add_sheet("Q&A")
        wb.write_sheet(
            [["Q", "A", "B"], ["Q1", 10, 20], ["Q2", 30, 40]],
            charts=[Chart("bar", [ChartSeries("B2:B3", title_cell="B1", categories="A2:A3")])],
        )
    chart = xml_root(parts(path)["xl/charts/chart1.xml"])  # raises if not well-formed
    assert "'Q&A'!$B$2:$B$3" in _chart_formulas(chart)


def test_radar_trendline_is_rejected():
    with pytest.raises(ValueError, match="radar"):
        Chart("radar", [ChartSeries("B2:B4", trendline=Trendline("linear"))])


def test_chart_title_size_is_written_as_an_integer(tmp_path):
    path = tmp_path / "size.xlsx"
    with XlsxWriter(path) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["a"], [1]], charts=[Chart("bar", [ChartSeries("A2")], title="T", title_size=12.345)])
    chart = parts(path)["xl/charts/chart1.xml"].decode()
    sizes = re.findall(r'sz="([^"]+)"', chart)
    assert sizes and all(size.isdigit() for size in sizes)


def test_control_character_in_chart_title_is_rejected():
    with pytest.raises(ValueError, match="control character"):
        Chart("bar", [ChartSeries("A2")], title="bad\x01title")
