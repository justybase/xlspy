"""XML edge cases in the XLSX sheet management API."""

from __future__ import annotations

import re

from xlspy import XlsxUpdater, XlsxWriter


def _workbook(path, names=("One", "Two")):
    with XlsxWriter(path) as writer:
        for name in names:
            writer.add_sheet(name)
            writer.write_sheet([[name]])
    return path


def _rewrite_workbook(updater, change):
    xml = updater._package.get("xl/workbook.xml").decode("utf-8")
    updater._package.set("xl/workbook.xml", change(xml).encode("utf-8"))


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
