"""Workbook.create_chartsheet() for XLSX files."""

import os

import pytest

from xlspy import Chart, ChartSeries, Workbook, load_workbook


def _data(workbook):
    sheet = workbook.active
    sheet.title = "Data"
    for row in (["name", "value"], ["a", 1], ["b", 4], ["c", 9]):
        sheet.append(row)
    return sheet


def _chart(title="Values"):
    return Chart("bar", [ChartSeries("B2:B4", categories="A2:A4", title="value")], title=title)


def test_new_workbook_with_a_chartsheet(tmp_path):
    workbook = Workbook()
    _data(workbook)
    chartsheet = workbook.create_chartsheet("Overview")
    chartsheet.add_chart(_chart())
    workbook.save(tmp_path / "charts.xlsx")
    workbook.close()

    reloaded = load_workbook(tmp_path / "charts.xlsx")
    try:
        assert reloaded.sheetnames == ["Data"]
        [loaded] = reloaded.chartsheets
        assert loaded.title == "Overview"
        [chart] = loaded.charts
        assert chart.chart_type == "bar" and chart.title == "Values"
        assert chart.series_count == 1
        assert b"Data!$B$2:$B$4" in chart.xml.replace(b"'", b"")
    finally:
        reloaded.close()


def test_chartsheet_added_to_a_loaded_workbook(tmp_path):
    workbook = Workbook()
    _data(workbook)
    workbook.create_sheet("Other")["A1"] = 1
    workbook.save(tmp_path / "plain.xlsx")
    workbook.close()

    loaded = load_workbook(tmp_path / "plain.xlsx")
    try:
        loaded.create_chartsheet().add_chart(_chart("Loaded"), source="Data")
        loaded.save(tmp_path / "more.xlsx")
    finally:
        loaded.close()
    reloaded = load_workbook(tmp_path / "more.xlsx")
    try:
        assert [sheet.title for sheet in reloaded.chartsheets] == ["Chart1"]
        assert reloaded.chartsheets[0].charts[0].title == "Loaded"
        assert reloaded["Other"]["A1"].value == 1
    finally:
        reloaded.close()


def test_chartsheet_checks(tmp_path):
    workbook = Workbook()
    _data(workbook)
    chartsheet = workbook.create_chartsheet("Overview")
    with pytest.raises(ValueError, match="already exists"):
        workbook.create_chartsheet("Data")
    with pytest.raises(NotImplementedError):
        workbook.create_chartsheet("Front", 0)
    with pytest.raises(ValueError, match="exactly one chart"):
        workbook.save(tmp_path / "empty.xlsx")
    chartsheet.add_chart(_chart())
    with pytest.raises(ValueError, match="one chart"):
        chartsheet.add_chart(_chart())
    with pytest.raises(NotImplementedError, match="XLSX"):
        workbook.save(tmp_path / "charts.xlsb")
    workbook.close()


@pytest.mark.com
@pytest.mark.skipif(os.environ.get("XLSPY_RUN_EXCEL_COM") != "1", reason="Set XLSPY_RUN_EXCEL_COM=1")
def test_excel_opens_the_chartsheet(tmp_path):
    import win32com.client

    workbook = Workbook()
    _data(workbook)
    workbook.create_chartsheet("Overview").add_chart(_chart())
    workbook.save(tmp_path / "charts.xlsx")
    workbook.close()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(tmp_path / "charts.xlsx"), CorruptLoad=0)
        try:
            assert book.Charts.Count == 1
            chart = book.Charts("Overview")
            assert chart.HasTitle and chart.ChartTitle.Text == "Values"
            assert list(chart.SeriesCollection(1).Values) == [1, 4, 9]
            assert [sheet.Name for sheet in book.Sheets] == ["Data", "Overview"]
        finally:
            book.Close(False)
    finally:
        excel.Quit()
