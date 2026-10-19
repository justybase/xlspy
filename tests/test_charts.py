import zipfile

import pytest

from xlspy import Chart, ChartSeries, ExcelReader, XlsxWriter


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
            title = chart.title.tx.rich.p[0].r[0].t if chart.title else None
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
    assert "<legend><legendPos val=\"r\"/></legend>" in chart

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
