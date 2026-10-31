"""The title, axis titles and legend of a loaded chart can be edited."""

import zipfile

import pytest

from xlspy import Chart, ChartAxis, ChartSeries, Workbook, load_workbook
from xlspy.chart_edit import set_axis_title, set_chart_title, set_legend_position


@pytest.fixture(params=["xlsx", "xlsb"])
def source(request, tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    for row in (["name", "value"], ["a", 1], ["b", 2], ["c", 3]):
        sheet.append(row)
    chart = Chart(
        "bar", [ChartSeries("B2:B4", categories="A2:A4")], title="Original",
        x_axis=ChartAxis(title="Names"),
    ) if request.param == "xlsx" else Chart(
        "bar", [ChartSeries("B2:B4", categories="A2:A4")], title="Original",
    )
    sheet.add_chart(chart, "D2")
    path = tmp_path / f"source.{request.param}"
    workbook.save(path)
    workbook.close()
    return path


def _roundtrip(source, tmp_path, edit):
    out = tmp_path / f"out{source.suffix}"
    workbook = load_workbook(source)
    try:
        edit(workbook.active.charts[0])
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        return reloaded.active.charts[0], out
    finally:
        reloaded.close()


def test_title_is_changed_on_save(source, tmp_path):
    chart, out = _roundtrip(source, tmp_path, lambda c: setattr(c, "title", "Sales & <costs>"))
    assert chart.title == "Sales & <costs>"
    assert chart.series_count == 1
    with zipfile.ZipFile(out) as archive:
        assert archive.testzip() is None


def test_title_can_be_removed_and_added(source, tmp_path):
    chart, _ = _roundtrip(source, tmp_path, lambda c: setattr(c, "title", None))
    assert chart.title is None
    chart, _ = _roundtrip(source, tmp_path, lambda c: (
        setattr(c, "title", None), setattr(c, "title", "Back"),
    ))
    assert chart.title == "Back"


def test_axis_titles_and_legend(source, tmp_path):
    def edit(chart):
        chart.x_axis_title = "Categories"
        chart.y_axis_title = "Amount"
        chart.legend_position = "bottom"

    chart, _ = _roundtrip(source, tmp_path, edit)
    assert chart.x_axis_title == "Categories"
    assert chart.y_axis_title == "Amount"
    assert chart.legend_position == "bottom"
    assert chart.title == "Original"


def test_legend_can_be_removed(source, tmp_path):
    chart, _ = _roundtrip(source, tmp_path, lambda c: setattr(c, "legend_position", None))
    assert chart.legend_position is None


def test_unedited_chart_is_not_rewritten(source, tmp_path):
    out = tmp_path / f"same{source.suffix}"
    workbook = load_workbook(source)
    try:
        workbook.active.charts[0].title = "Original"  # same text
        workbook.save(out)
    finally:
        workbook.close()
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(out) as after:
        assert before.read("xl/charts/chart1.xml").count(b"Original") == 1
        assert b"Original" in after.read("xl/charts/chart1.xml")


def test_edits_keep_unrelated_markup():
    xml = (
        b'<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/'
        b'drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        b' xmlns:mc="x" mc:Ignorable="c14"><c:chart><c:title><c:tx><c:rich><a:p>'
        b'<a:r><a:rPr sz="2000"/><a:t>Old</a:t></a:r><a:r><a:t> more</a:t></a:r></a:p>'
        b'</c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/><c:plotArea>'
        b'<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:numFmt formatCode="G"/></c:catAx>'
        b'<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'
        b'</c:plotArea></c:chart></c:chartSpace>'
    )
    edited = set_chart_title(xml, "New").decode()
    assert 'mc:Ignorable="c14"' in edited and 'sz="2000"' in edited
    assert "<a:t>New</a:t>" in edited and "more" not in edited
    edited = set_axis_title(xml, 1, "Y").decode()
    assert edited.index("<c:title>") < edited.index("<c:crossAx") and "<a:t>Y</a:t>" in edited
    edited = set_axis_title(xml, 0, "X").decode()
    assert edited.index("<c:axPos") < edited.rindex("<c:title>") < edited.index("<c:numFmt")
    assert b"<c:legend>" in set_legend_position(xml, "top")
    with pytest.raises(ValueError):
        set_legend_position(xml, "nowhere")


def _series_book(tmp_path, ext):
    workbook = Workbook()
    sheet = workbook.active
    for row in (["n", "a", "b"], ["x", 1, 10], ["y", 2, 30]):
        sheet.append(row)
    sheet.add_chart(Chart("bar", [
        ChartSeries("B2:B3", categories="A2:A3", title="One", color="#FF0000"),
        ChartSeries("C2:C3", categories="A2:A3", title_cell="C1"),
    ]), "E2")
    sheet.add_chart(Chart("line", [ChartSeries("B2:B3", categories="A2:A3", color="#00FF00")]), "E20")
    path = tmp_path / f"series.{ext}"
    workbook.save(path)
    workbook.close()
    return path


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_series_titles_and_colors_are_changed_on_save(tmp_path, ext):
    path = _series_book(tmp_path, ext)
    out = tmp_path / f"out.{ext}"
    workbook = load_workbook(path)
    try:
        bars, line = workbook.active.charts
        assert bars.series_titles[0] == "One"
        bars.set_series_title(0, "Renamed & <new>")
        bars.set_series_title(1, "Second")
        bars.set_series_color(0, "#0000FF")
        bars.set_series_color(1, "#00AA00")
        line.set_series_color(0, "#123456")
        workbook.save(out)
    finally:
        workbook.close()
    with zipfile.ZipFile(out) as archive:
        xmls = [archive.read(n).decode() for n in sorted(archive.namelist()) if n.startswith("xl/charts/chart")]
    joined = "".join(xmls)
    assert "Renamed &amp; &lt;new&gt;" in joined and "Second" in joined
    assert "0000FF" in joined and "00AA00" in joined and "123456" in joined
    assert "FF0000" not in joined and "00FF00" not in joined
    reloaded = load_workbook(out)
    try:
        assert reloaded.active.charts[0].series_titles == ["Renamed & <new>", "Second"]
    finally:
        reloaded.close()


def test_series_color_edits_keep_the_rest_of_the_shape():
    from xlspy.chart_edit import set_series_color, set_series_title

    xml = (
        b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        b'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>'
        b'<c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f>'
        b'</c:strRef></c:tx><c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent1"/>'
        b'</a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:marker><c:symbol val="none"/></c:marker>'
        b'<c:val><c:numRef><c:f>S!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:lineChart>'
        b'<c:barChart><c:ser><c:idx val="1"/><c:order val="1"/><c:invertIfNegative val="0"/>'
        b'<c:val><c:numRef><c:f>S!$C$2:$C$3</c:f></c:numRef></c:val></c:ser></c:barChart>'
        b'</c:plotArea></c:chart></c:chartSpace>'
    )
    edited = set_series_color(xml, 0, "ABCDEF").decode()
    assert 'w="28575" cap="rnd"' in edited and "<a:round/>" in edited and "accent1" not in edited
    assert edited.index("ABCDEF") < edited.index("<a:round/>") and "<a:effectLst/>" in edited
    edited = set_series_color(xml, 1, "112233").decode()
    assert "<c:spPr><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></c:spPr>" in edited
    assert edited.index("<c:order val=\"1\"/>") < edited.index("112233") < edited.index("invertIfNegative")
    assert "<c:v>T</c:v>" in set_series_title(xml, 0, "T").decode()
    assert "<c:tx><c:v>T</c:v></c:tx><c:spPr>" in set_series_title(xml, 0, "T").decode()
    assert "<c:order val=\"1\"/><c:tx><c:v>U</c:v></c:tx>" in set_series_title(xml, 1, "U").decode()
