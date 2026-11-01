"""The one <sheetData> locator used by the updater, the cell editor and the loader."""

import pytest

from xlspy._sheet_data import replace_sheet_data_body, sheet_data_match, tag_prefix


def test_tag_prefix_plain_and_prefixed():
    assert tag_prefix("<worksheet><sheetData/></worksheet>", "sheetData") == ""
    assert tag_prefix('<x:worksheet xmlns:x="u"><x:sheetData/></x:worksheet>', "sheetData") == "x:"
    assert tag_prefix("<worksheet/>", "sheetData") == ""


def test_match_reports_prefix_for_empty_and_filled_elements():
    empty = sheet_data_match('<x:worksheet><x:sheetData/></x:worksheet>')
    filled = sheet_data_match("<worksheet><sheetData><row r=\"1\"/></sheetData></worksheet>")
    assert empty["prefix"] == "x:" and empty[0] == "<x:sheetData/>"
    assert filled["prefix"] is None and filled[0] == '<sheetData><row r="1"/></sheetData>'


def test_empty_element_is_expanded_and_keeps_attributes():
    xml = '<worksheet><sheetData a="1"/></worksheet>'
    assert replace_sheet_data_body(xml, '<row r="1"/>') == (
        '<worksheet><sheetData a="1"><row r="1"/></sheetData></worksheet>'
    )


def test_space_before_self_closing_slash_is_kept_as_written():
    xml = "<worksheet><sheetData /></worksheet>"
    assert replace_sheet_data_body(xml, "<row/>") == "<worksheet><sheetData ><row/></sheetData></worksheet>"


def test_prefix_and_surrounding_xml_are_untouched():
    xml = '<x:worksheet><x:sheetData><x:row r="1"/></x:sheetData><x:tail/></x:worksheet>'
    assert replace_sheet_data_body(xml, '<x:row r="2"/>') == (
        '<x:worksheet><x:sheetData><x:row r="2"/></x:sheetData><x:tail/></x:worksheet>'
    )


def test_missing_element_is_a_value_error():
    with pytest.raises(ValueError, match="sheetData element was not found"):
        replace_sheet_data_body("<worksheet/>", "")
