"""The shared worksheet-part edits (xlspy._sheet_data): the one <sheetData> locator used by
the updater, the cell editor and the loader, the row and cell walks, and the dimension ref."""
from __future__ import annotations

import pytest

from xlspy._sheet_data import (
    insert_rows,
    patch_cells,
    patch_rows,
    replace_sheet_data_body,
    set_dimension_ref,
    sheet_data_match,
    tag_prefix,
)

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'


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


def test_patch_rows_numbers_rows_without_r_after_the_previous_row() -> None:
    element = '<sheetData><row r="3"><c r="A3"/></row><row/><row r="9"/></sheetData>'
    seen = []

    def patch(number: int, row_xml: str) -> str:
        seen.append(number)
        return row_xml

    assert patch_rows(element, patch) == element
    assert seen == [3, 4, 9]


def test_insert_rows_keeps_row_order() -> None:
    element = '<sheetData><row r="2"/><row r="5"/></sheetData>'
    result = insert_rows(element, {7: '<row r="7"/>', 1: '<row r="1"/>', 3: '<row r="3"/>'})
    assert result == '<sheetData><row r="1"/><row r="2"/><row r="3"/><row r="5"/><row r="7"/></sheetData>'


def test_insert_rows_expands_an_empty_sheet_data() -> None:
    assert insert_rows("<x:sheetData/>", {1: '<x:row r="1"/>'}) == '<x:sheetData><x:row r="1"/></x:sheetData>'
    assert insert_rows("<sheetData/>", {}) == "<sheetData/>"


def test_patch_cells_gives_cells_without_r_their_position() -> None:
    xml = f'<worksheet {NS}><sheetData><row r="2"><c r="B2"/><c/></row><row><c/></row></sheetData></worksheet>'
    refs = []

    def patch(ref: str, cell_xml: str) -> str:
        refs.append(ref)
        return cell_xml.replace("<c", f'<c s="{len(refs)}"', 1)

    assert patch_cells(xml, patch, set()) == xml
    result = patch_cells(xml, patch, {2, 3})
    assert refs == ["B2", "C2", "A3"]
    assert '<c s="2"/>' in result and result.startswith(f"<worksheet {NS}>")


def test_set_dimension_ref_keeps_the_existing_element() -> None:
    xml = f'<x:worksheet {NS}><x:dimension ref="A1" extra="1" /><x:sheetData/></x:worksheet>'
    assert set_dimension_ref(xml, "A1:C3") == (
        f'<x:worksheet {NS}><x:dimension ref="A1:C3" extra="1" /><x:sheetData/></x:worksheet>'
    )


def test_set_dimension_ref_adds_a_missing_element_in_schema_order() -> None:
    xml = f'<x:worksheet {NS}><x:sheetPr/><x:sheetViews/><x:sheetData/></x:worksheet>'
    assert set_dimension_ref(xml, "B2:D4") == (
        f'<x:worksheet {NS}><x:sheetPr/><x:dimension ref="B2:D4"/><x:sheetViews/><x:sheetData/></x:worksheet>'
    )


def test_set_dimension_ref_needs_a_worksheet() -> None:
    with pytest.raises(ValueError, match="Worksheet root"):
        set_dimension_ref("<chartsheet/>", "A1")
