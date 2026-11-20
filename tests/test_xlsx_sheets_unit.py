"""XML edge cases in the XLSX sheet management API."""

from __future__ import annotations

import re

import pytest

from xlspy import (
    Chart,
    ChartSeries,
    Comment,
    ExcelTable,
    Slicer,
    UnsupportedFeatureError,
    Workbook,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)


def _workbook(path, names=("One", "Two")):
    with XlsxWriter(path) as writer:
        for name in names:
            writer.add_sheet(name)
            writer.write_sheet([[name]])
    return path


def _rewrite_workbook(updater, change):
    xml = updater._package.get("xl/workbook.xml").decode("utf-8")
    updater._package.set("xl/workbook.xml", change(xml).encode("utf-8"))


def test_add_sheet_creates_a_worksheet_part_and_relationships(tmp_path):
    path = _workbook(tmp_path / "add.xlsx", ("One", "Two"))
    updater = XlsxUpdater(path)
    try:
        added = updater.add_sheet("Three", 1)
        assert added == "xl/worksheets/sheet3.xml"
        package = updater._package
        assert package.has(added)
        assert b"Three" in package.get("xl/workbook.xml")
        rels = package.get("xl/_rels/workbook.xml.rels").decode("utf-8")
        assert 'Target="worksheets/sheet3.xml"' in rels
        assert 'PartName="/xl/worksheets/sheet3.xml"' in package.get(
            "[Content_Types].xml"
        ).decode("utf-8")
        assert updater.get_sheet_names() == ["One", "Three", "Two"]
        updater.save(path)
    finally:
        updater.close()

    reopened = XlsxUpdater(path)
    try:
        assert reopened.get_sheet_names() == ["One", "Three", "Two"]
        assert reopened.get_sheet_states()["Three"] == "visible"
    finally:
        reopened.close()


def test_add_sheet_keeps_the_active_sheet_and_tab_bar_in_place(tmp_path):
    path = _workbook(tmp_path / "add-active.xlsx", ("One", "Two"))
    updater = XlsxUpdater(path)
    try:
        updater.set_active_sheet(1)
        updater.add_sheet("Zero", 0)
        assert updater.get_active_sheet() == 2
        xml = updater._package.get("xl/workbook.xml").decode("utf-8")
        assert re.search(r'activeTab="2"', xml)
    finally:
        updater.close()


def test_active_sheet_is_created_before_the_sheet_list(tmp_path):
    path = _workbook(tmp_path / "active.xlsx")
    updater = XlsxUpdater(path)
    try:
        _rewrite_workbook(updater, lambda xml: re.sub(r"<bookViews\b.*?</bookViews>", "", xml, flags=re.DOTALL))
        updater.set_active_sheet(1)
        xml = updater._package.get("xl/workbook.xml").decode("utf-8")
        assert re.search(r'<bookViews[^>]*><workbookView[^>]*activeTab="1"', xml)
        assert xml.index("<bookViews") < xml.index("<sheets")
        assert updater.get_active_sheet() == 1
        updater.save(path)
    finally:
        updater.close()

    reopened = XlsxUpdater(path)
    try:
        assert reopened.get_active_sheet() == 1
    finally:
        reopened.close()


def test_invalid_active_tab_reads_as_the_first_sheet(tmp_path):
    path = _workbook(tmp_path / "invalid-active.xlsx")
    updater = XlsxUpdater(path)
    try:
        _rewrite_workbook(
            updater,
            lambda xml: re.sub(r'activeTab="\d+"', 'activeTab="not-a-number"', xml),
        )
        assert updater.get_active_sheet() == 0
    finally:
        updater.close()


def test_unknown_sheet_state_reads_as_visible(tmp_path):
    path = _workbook(tmp_path / "invalid-state.xlsx")
    updater = XlsxUpdater(path)
    try:
        _rewrite_workbook(
            updater,
            lambda xml: re.sub(r'(<sheet\b[^>]*?)(/?>)', r'\1 state="bogus"\2', xml, count=1),
        )
        assert updater.get_sheet_states()["One"] == "visible"
    finally:
        updater.close()


def test_active_sheet_pulls_the_tab_bar_to_it(tmp_path):
    path = _workbook(tmp_path / "tab-bar.xlsx")
    updater = XlsxUpdater(path)
    try:
        _rewrite_workbook(
            updater,
            lambda xml: re.sub(r"(<workbookView\b[^>]*?)(/?>)", r'\1 firstSheet="1"\2', xml, count=1),
        )
        updater.set_active_sheet(0)
        xml = updater._package.get("xl/workbook.xml").decode("utf-8")
        assert re.search(r'firstSheet="0"', xml)
        assert re.search(r'activeTab="0"', xml)
    finally:
        updater.close()


def test_reorder_remaps_the_tab_bar_position(tmp_path):
    path = _workbook(tmp_path / "tab-bar-reorder.xlsx", ("One", "Two", "Three"))
    updater = XlsxUpdater(path)
    try:
        _rewrite_workbook(
            updater,
            lambda xml: re.sub(r"(<workbookView\b[^>]*?)(/?>)", r'\1 firstSheet="2"\2', xml, count=1),
        )
        updater.reorder_sheets(["Three", "Two", "One"])
        xml = updater._package.get("xl/workbook.xml").decode("utf-8")
        assert re.search(r'firstSheet="0"', xml)
        assert re.search(r'activeTab="2"', xml)
    finally:
        updater.close()


def test_remove_sheet_drops_its_comments_drawing_and_chart_parts(tmp_path):
    path = tmp_path / "owned.xlsx"
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data["A1"] = 1
    data["A2"] = 2
    data["A1"].comment = Comment("note", "author")
    data.add_chart(Chart("bar", [ChartSeries("A1:A2")]), "D2")
    other = workbook.create_sheet("Other")
    other["A1"] = "keep"
    workbook.save(path)
    workbook.close()

    updater = XlsxUpdater(path)
    try:
        owned = {
            name for name in updater._package.names()
            if "comments" in name or "vmlDrawing" in name
            or "/drawings/" in name or "/charts/" in name
        }
        assert owned
        updater.remove_sheet("Data")
        assert owned.isdisjoint(updater._package.names())
        assert updater.get_sheet_names() == ["Other"]
        updater.save(path)
    finally:
        updater.close()

    loaded = load_workbook(path)
    try:
        assert loaded.get_sheet_names() == ["Other"]
        assert loaded["Other"]["A1"].value == "keep"
    finally:
        loaded.close()


def test_remove_sheet_refuses_a_sheet_with_a_slicer(tmp_path):
    path = tmp_path / "slicer.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Region", "Sales"], ["North", 1]],
            tables=[ExcelTable("Sales", "A1:B2")],
            slicers=[Slicer("Sales", "Region", anchor="D2")],
        )
        writer.add_sheet("Other")
        writer.write_sheet([["keep"]])

    updater = XlsxUpdater(path)
    try:
        with pytest.raises(UnsupportedFeatureError):
            updater.remove_sheet("Data")
        assert updater.get_sheet_names() == ["Data", "Other"]
    finally:
        updater.close()


def test_remove_chartsheet_drops_its_parts(tmp_path):
    path = tmp_path / "chartsheet-remove.xlsx"
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data["A1"] = 1
    data["A2"] = 2
    chart = workbook.create_chartsheet("ChartOnly")
    chart.add_chart(Chart("bar", [ChartSeries("A1:A2")]), source=data)
    workbook.save(path)
    workbook.close()

    updater = XlsxUpdater(path)
    try:
        assert updater.get_chartsheet_names() == ["ChartOnly"]
        updater.remove_sheet("ChartOnly")
        assert updater.get_sheet_names() == ["Data"]
        assert updater.get_chartsheet_names() == []
        updater.save(path)
    finally:
        updater.close()

    loaded = load_workbook(path)
    try:
        assert loaded.get_sheet_names() == ["Data"]
    finally:
        loaded.close()


def test_reorder_moves_away_from_a_hidden_active_sheet(tmp_path):
    path = _workbook(tmp_path / "hidden-active.xlsx", ("One", "Two"))
    updater = XlsxUpdater(path)
    try:
        updater.set_sheet_state("One", "hidden")
        updater.set_active_sheet(0)
        updater.reorder_sheets(["Two", "One"])
        assert updater.get_active_sheet() == 0
        assert updater.get_sheet_states() == {"Two": "visible", "One": "hidden"}
        updater.save(path)
    finally:
        updater.close()

    reopened = XlsxUpdater(path)
    try:
        assert reopened.get_active_sheet() == 0
        assert reopened.get_sheet_names() == ["Two", "One"]
    finally:
        reopened.close()
