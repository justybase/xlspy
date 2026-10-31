"""A loaded XLSB keeps and reads its chartsheets."""

import zipfile
from pathlib import Path

import pytest

from xlspy import load_workbook
from xlspy.xlsb_updater import XlsbUpdater

FIXTURE = Path(__file__).parent / "data" / "chartsheet.xlsb"  # one worksheet, one chartsheet (Excel)


def _save(tmp_path, edit):
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(FIXTURE)
    try:
        edit(workbook)
        workbook.save(out)
    finally:
        workbook.close()
    return out


def _summary(path):
    workbook = load_workbook(path)
    try:
        return (
            workbook.sheetnames,
            [(sheet.title, [chart.title for chart in sheet.charts]) for sheet in workbook.chartsheets],
        )
    finally:
        workbook.close()


def test_chartsheet_is_read():
    names, charts = _summary(FIXTURE)
    assert len(names) == 1 and charts == [("ChartOnly", ["CS title"])]


def test_updater_tells_worksheets_from_chartsheets():
    updater = XlsbUpdater(FIXTURE)
    try:
        assert updater.get_chartsheet_names() == ["ChartOnly"]
        assert "ChartOnly" not in updater.get_worksheet_names()
        assert "ChartOnly" in updater.get_sheet_names()
    finally:
        updater.close()


def test_edits_keep_the_chartsheet(tmp_path):
    out = _save(tmp_path, lambda wb: wb.active.__setitem__("B1", "x"))
    with zipfile.ZipFile(out) as archive:
        assert any(name.startswith("xl/chartsheets/") for name in archive.namelist())
    assert _summary(out)[1] == [("ChartOnly", ["CS title"])]


def test_sheet_structure_edits_keep_the_chartsheet(tmp_path):
    def edit(workbook):
        workbook.active.title = "Renamed"
        extra = workbook.create_sheet("Extra")
        extra["A1"] = 1
        workbook.move_sheet("Extra", offset=-1)

    out = _save(tmp_path, edit)
    names, charts = _summary(out)
    assert names == ["Extra", "Renamed"] and charts == [("ChartOnly", ["CS title"])]


def test_chartsheet_chart_title_can_be_edited(tmp_path):
    out = _save(tmp_path, lambda wb: setattr(wb.chartsheets[0].charts[0], "title", "Changed"))
    assert _summary(out)[1] == [("ChartOnly", ["Changed"])]


def test_chartsheet_title_clash_is_refused():
    workbook = load_workbook(FIXTURE)
    try:
        with pytest.raises(ValueError):
            workbook.active.title = "chartonly"
    finally:
        workbook.close()


def test_new_chartsheet_is_added_after_the_other_sheets(tmp_path):
    from xlspy import Chart, ChartSeries

    def edit(workbook):
        sheet = workbook.active
        for row in (["n", "v"], ["a", 1], ["b", 2]):
            sheet.append(row)
        chartsheet = workbook.create_chartsheet("Summary")
        chartsheet.add_chart(Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], title="New"))

    names, charts = _summary(_save(tmp_path, edit))
    assert sorted(charts) == [("ChartOnly", ["CS title"]), ("Summary", ["New"])]


def test_new_chartsheet_on_a_workbook_without_one(tmp_path):
    from xlspy import Chart, ChartSeries, Workbook

    source = tmp_path / "plain.xlsb"
    workbook = Workbook()
    for row in (["n", "v"], ["a", 1], ["b", 2]):
        workbook.active.append(row)
    workbook.save(source)
    workbook.close()
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        sheet = workbook.create_chartsheet("Only")
        sheet.add_chart(Chart("line", [ChartSeries("B2:B3", categories="A2:A3")], title="T"))
        workbook.save(out)
        workbook.save(tmp_path / "again.xlsb")
    finally:
        workbook.close()
    assert _summary(out)[1] == [("Only", ["T"])]
    assert _summary(tmp_path / "again.xlsb")[1] == [("Only", ["T"])]


def test_active_chartsheet_survives_a_plain_resave(tmp_path):
    import shutil

    from xlspy import load_workbook
    from xlspy.xlsb_updater import XlsbUpdater

    source = tmp_path / "in.xlsb"
    shutil.copy("tests/data/chartsheet.xlsb", source)
    updater = XlsbUpdater(source)
    chart_index = updater.get_sheet_names().index(updater.get_chartsheet_names()[0])
    updater.set_active_sheet(chart_index)
    updater.save(source)
    updater.close()
    loaded = load_workbook(source)
    target = tmp_path / "out.xlsb"
    loaded.save(target)
    loaded.close()
    updater = XlsbUpdater(target)
    try:
        assert updater.get_active_sheet() == chart_index
    finally:
        updater.close()
