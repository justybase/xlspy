"""Combined charts, secondary axes, trendlines and error bars."""

import zipfile

import pytest

from xlspy import Chart, ChartAxis, ChartSeries, ErrorBars, Trendline, Workbook, load_workbook


def _book(tmp_path, chart, name="out.xlsx"):
    workbook = Workbook()
    sheet = workbook.active
    for row in (["n", "a", "b"], ["x", 1, 10], ["y", 2, 30], ["z", 3, 20]):
        sheet.append(row)
    sheet.add_chart(chart, "E2")
    path = tmp_path / name
    workbook.save(path)
    workbook.close()
    with zipfile.ZipFile(path) as archive:
        return path, archive.read("xl/charts/chart1.xml").decode()


def _combo(**kwargs):
    return Chart("bar", [
        ChartSeries("B2:B4", categories="A2:A4"),
        ChartSeries("C2:C4", categories="A2:A4", chart_type="line", secondary_axis=True),
    ], **kwargs)


def test_bar_and_line_on_a_secondary_axis(tmp_path):
    path, xml = _book(tmp_path, _combo(y2_axis=ChartAxis(title="Totals")))
    assert xml.index("<barChart>") < xml.index("<lineChart>")
    assert xml.count('<axId val="200"/>') == 2  # the line group and the secondary value axis
    assert '<crosses val="max"/>' in xml and '<axPos val="r"/>' in xml
    assert "<idx val=\"0\"/>" in xml and "<idx val=\"1\"/>" in xml
    workbook = load_workbook(path)
    try:
        chart = workbook.active.charts[0]
        assert chart.chart_type == "bar" and chart.series_count == 2
    finally:
        workbook.close()


def test_single_type_chart_is_unchanged(tmp_path):
    _path, xml = _book(tmp_path, Chart("bar", [ChartSeries("B2:B4", categories="A2:A4")]))
    assert "<lineChart>" not in xml and "<crosses" not in xml


def test_trendline_and_error_bars(tmp_path):
    chart = Chart("line", [ChartSeries(
        "B2:B4", categories="A2:A4",
        trendline=Trendline("polynomial", order=3, forward=1, display_equation=True, color="#FF0000"),
        error_bars=ErrorBars("percentage", 10, direction="plus", end_cap=False),
    )])
    _path, xml = _book(tmp_path, chart)
    assert '<trendlineType val="poly"/><order val="3"/><forward val="1"/>' in xml
    assert '<dispRSqr val="0"/><dispEq val="1"/>' in xml
    assert '<errValType val="percentage"/><noEndCap val="1"/><val val="10"/>' in xml
    assert xml.index("<trendline>") < xml.index("<errBars>") < xml.index("<cat>")


def test_scatter_series_with_x_error_bars(tmp_path):
    chart = Chart("scatter", [ChartSeries(
        "C2:C4", x_values="B2:B4", error_bars=ErrorBars("std_err", axis="x"),
        trendline=Trendline("moving_average", period=2),
    )])
    _path, xml = _book(tmp_path, chart)
    assert '<errDir val="x"/>' in xml and '<period val="2"/>' in xml
    assert xml.index("<errBars>") < xml.index("<xVal>")


@pytest.mark.parametrize("build", [
    lambda: Trendline("nope"),
    lambda: Trendline("linear", order=2),
    lambda: Trendline("polynomial", order=9),
    lambda: Trendline("moving_average", period=1),
    lambda: ErrorBars("fixed"),
    lambda: ErrorBars("std_err", 1),
    lambda: ErrorBars("fixed", 1, direction="up"),
    lambda: ChartSeries("B2:B4", chart_type="pie"),
    lambda: Chart("pie", [ChartSeries("B2:B4", chart_type="line")]),
    lambda: Chart("barh", [ChartSeries("B2:B4", secondary_axis=True)]),
    lambda: Chart("bar", [ChartSeries("B2:B4")], y2_axis=ChartAxis()),
    lambda: Chart("pie", [ChartSeries("B2:B4", trendline=Trendline())]),
    lambda: Chart("line", [ChartSeries("B2:B4", error_bars=ErrorBars("fixed", 1, axis="x"))]),
])
def test_invalid_options_are_refused(build):
    with pytest.raises((ValueError, TypeError)):
        build()


def test_xlsb_writer_writes_the_new_options(tmp_path):
    chart = _combo()
    path, xml = _book(tmp_path, chart, "x.xlsb")
    assert "<lineChart>" in xml and '<crosses val="max"/>' in xml
    workbook = load_workbook(path)
    try:
        assert workbook.active.charts[0].series_count == 2
    finally:
        workbook.close()
