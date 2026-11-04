"""The tag-at-a-time walk over a worksheet part: the same elements as the regular
expression it stands in for, and edit times that follow the size of the sheet."""
from __future__ import annotations

import random
import re
import time
import zipfile

import pytest

from xlspy import Font, Formula, Workbook, XlsxWriter, load_workbook
from xlspy._metadata import element_pattern
from xlspy._sheet_data import _iter_rows, insert_rows, patch_cells, patch_rows, sheet_data_match
from xlspy.xlsx_updater import XlsxUpdater

MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
SHEET = "xl/worksheets/sheet1.xml"

# Written and damaged forms, in any order: the walk and the pattern must not differ.
_PIECES = [
    "<row>", "</row>", "<row/>", "<row />", '<row r="3">', '<row r="3"/>', "<row r='4' a='b'>",
    '<row a="1" r="5">', '<row a="1"/>', '<row\nr="6">', "<x:row>", "</x:row>", "<x:row/>", '<x:row r="7">',
    "</row >", "</row\n>", "</x:row >", "</rows>", "</rowx >", "<rows>", "<rowBreaks/>", "<row-x>", "<y:row>",
    "</y:row>", '<c r="A1"><v>1</v></c>', "<c/>", "text", ">", "/>", "<", "</",
]


def _random_xml(rng: random.Random) -> str:
    return "".join(rng.choice(_PIECES) for _ in range(rng.randint(0, 14)))


@pytest.mark.parametrize("seed", range(4000))
def test_the_row_walk_finds_what_the_pattern_finds(seed):
    xml = _random_xml(random.Random(seed))
    assert [(start, end) for start, end, _ in _iter_rows(xml)] == [
        match.span() for match in element_pattern("row").finditer(xml)
    ]


@pytest.mark.parametrize("seed", range(2000))
def test_the_sheet_data_search_finds_what_the_pattern_finds(seed):
    rng = random.Random(seed)
    pieces = ["<sheetData>", "</sheetData>", "</sheetData >", "<sheetData/>", '<sheetData a="1">', "<x:sheetData>",
              "</x:sheetData>", "<x:sheetData/>", "<sheetDatas>", "</sheetDatas>", "<row/>", "text", "<"]
    xml = "".join(rng.choice(pieces) for _ in range(rng.randint(0, 8)))
    found = sheet_data_match(xml)
    expected = element_pattern("sheetData").search(xml)
    if expected is None:
        assert found is None
    else:
        assert (found.start(), found.end(), found[0], found["prefix"]) == (
            expected.start(), expected.end(), expected[0], expected["prefix"])


def test_a_row_without_a_number_follows_the_previous_row():
    xml = '<row/><row r="5"><c/></row><row></row><row a="1"/><row r=\'9\'/><row/>'
    assert [number for _, _, number in _iter_rows(xml)] == [1, 5, 6, 7, 9, 10]


def test_patch_rows_touches_only_the_rows_it_is_given():
    element = '<sheetData><row r="1"><c/></row>\n<row><c/></row><row r="9"/></sheetData>'
    seen = []

    def patch(number, row_xml):
        seen.append((number, row_xml))
        return f"[{number}]"

    assert patch_rows(element, patch, {2, 9, 40}) == '<sheetData><row r="1"><c/></row>\n[2][9]</sheetData>'
    assert seen == [(2, "<row><c/></row>"), (9, '<row r="9"/>')]
    assert patch_rows(element, lambda number, row_xml: "x") == "<sheetData>x\nxx</sheetData>"
    assert patch_rows(element, patch, set()) == element


def test_insert_rows_goes_before_the_first_higher_row_and_ends_the_rest_at_the_end():
    element = '<sheetData><row r="3"/><row r="7"/></sheetData>'
    new = {1: "<A/>", 2: "<B/>", 5: "<C/>", 9: "<D/>"}
    assert insert_rows(element, new) == '<sheetData><A/><B/><row r="3"/><C/><row r="7"/><D/></sheetData>'
    assert insert_rows(element, {}) is element
    assert insert_rows("<sheetData/>", {1: "<A/>"}) == "<sheetData><A/></sheetData>"


def test_patch_cells_leaves_other_rows_alone():
    xml = f'<worksheet {MAIN}><sheetData><row r="1"><c/></row><row r="2"><c/><c r="D2"/></row></sheetData></worksheet>'
    refs = []
    patch_cells(xml, lambda ref, cell: refs.append(ref) or cell, {2})
    assert refs == ["A2", "D2"]


def _package(tmp_path, body, *, dimension='<dimension ref="A1"/>'):
    source = tmp_path / "source.xlsx"
    with XlsxWriter(str(source)) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x"]])
    path = tmp_path / "hand.xlsx"
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(path, "w") as target:
        for name in original.namelist():
            data = original.read(name)
            if name == SHEET:
                data = f'<?xml version="1.0"?><worksheet {MAIN}>{dimension}<sheetData>{body}</sheetData></worksheet>'
            target.writestr(name, data)
    return path


def _edited(path, cells, tmp_path):
    out = tmp_path / "out.xlsx"
    updater = XlsxUpdater(str(path))
    try:
        updater.update_cells("S", cells)
        updater.save(str(out))
    finally:
        updater.close()
    with zipfile.ZipFile(out) as archive:
        return archive.read(SHEET).decode("utf-8")


def test_the_dimension_counts_cells_in_rows_the_edit_does_not_touch(tmp_path):
    path = _package(tmp_path, '<row r="1"><c r="B1"/></row><row r="30"><c/><c/><c r="H30"/></row>')
    assert '<dimension ref="A1:H30"/>' in _edited(path, {"A1": 1}, tmp_path)


def test_cells_without_a_reference_are_edited_where_they_stand(tmp_path):
    path = _package(tmp_path, "<row><c><v>1</v></c><c><v>2</v></c></row><row><c/><c><v>4</v></c></row>")
    xml = _edited(path, {"B2": 9}, tmp_path)
    assert '<row><c><v>1</v></c><c><v>2</v></c></row>' in xml
    assert '<c r="B2"><v>9</v></c>' in xml and '<v>4</v>' not in xml


@pytest.mark.parametrize("body, message", [
    ('<row r="9"><c r="nonsense"/></row>', "invalid A1"),
    ('<row r="9"><c r="A1:B2"/></row>', "single cell"),
    ('<row r="x"><c r="A1"/></row>', "invalid literal"),
])
def test_a_damaged_row_is_refused_even_when_the_edit_is_elsewhere(tmp_path, body, message):
    with pytest.raises(ValueError, match=message):
        _edited(_package(tmp_path, '<row r="1"><c r="A1"/></row>' + body), {"A1": 1}, tmp_path)


def test_a_shared_formula_member_is_refused_in_a_sheet_of_cells_without_references(tmp_path):
    body = '<row><c><f t="shared" ref="A1:A2" si="0">1+1</f></c></row><row><c><f t="shared" si="0"/></c><c/></row>'
    path = _package(tmp_path, body)
    with pytest.raises(ValueError, match="shared formula"):
        _edited(path, {"A2": 1}, tmp_path)
    assert "<f>" not in _edited(path, {"B2": 1}, tmp_path)
    assert "<f>2+2</f>" in _edited(path, {"C5": Formula("2+2")}, tmp_path)


# Far above what these take now; the quadratic walk took minutes on them.
_ROWS, _LIMIT = 8000, 30.0


def test_the_pattern_does_not_run_to_the_end_of_the_document_per_element():
    xml = "".join(f'<row r="{i}"><c r="A{i}"><v>{i}</v></c></row>' for i in range(1, 40001))
    start = time.perf_counter()
    assert sum(1 for _ in element_pattern("row").finditer(xml)) == 40000
    assert time.perf_counter() - start < _LIMIT
    # A prefix is still looked for, and a closing tag without one still ends the element.
    assert re.fullmatch(element_pattern("row"), "<x:row><c/></row>") is not None
    assert element_pattern("row").search("<row><c/></x:row>") is None


def test_saving_a_loaded_workbook_after_edits_follows_the_size_of_the_sheet(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    for row in range(_ROWS):
        sheet.append([f"text {row}", row, row * 0.5, row + 1, "same"])
    source = tmp_path / "large.xlsx"
    workbook.save(str(source))
    loaded = load_workbook(str(source))
    sheet = loaded.active
    bold = Font(bold=True)
    for row in range(2, _ROWS, 8):
        sheet.cell(row=row, column=2).value = -row
        sheet.cell(row=row, column=3).font = bold
    start = time.perf_counter()
    loaded.save(str(tmp_path / "edited.xlsx"))
    assert time.perf_counter() - start < _LIMIT
    again = load_workbook(str(tmp_path / "edited.xlsx")).active
    assert again.cell(row=10, column=2).value == -10
    assert again.cell(row=10, column=3).font.bold is True
    assert again.cell(row=11, column=2).value == 10
    assert again.max_row == _ROWS
