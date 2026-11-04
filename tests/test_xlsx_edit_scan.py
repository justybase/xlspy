"""The native scan of a sheet for ``XlsxUpdater.update_cells`` reads what the tree does.

``_fast_scan`` (Rust) and ``_tree_scan`` (ElementTree) must answer the same for
every sheet the native scan takes; for the rest it returns None and the tree is
used. The refusals of the old code, all of which came from the tree, stay.
"""

import random
import xml.etree.ElementTree as ET
import zipfile

import pytest

from xlspy import Formula, XlsxUpdater, XlsxWriter, _accel
from xlspy._safe_xml import check_well_formed, fromstring
from xlspy._xlsx_edit import _fast_scan, _fast_sheet, _tree_scan

pytestmark = pytest.mark.skipif(_accel.xlsx_edit_scan is None, reason="native extension not built")

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _sheet(body, *, head='<dimension ref="A1:C3"/>', tail="", root_extra=""):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<worksheet xmlns="{MAIN}"{root_extra}>{head}<sheetData>{body}</sheetData>{tail}</worksheet>'
    )


CASES = {
    "plain": _sheet('<row r="1"><c r="A1"><v>1</v></c><c r="C1" t="s"><v>0</v></c></row><row r="3"><c r="B3"/></row>'),
    "no_refs": _sheet("<row><c><v>1</v></c><c/><c><v>3</v></c></row><row><c/><c><v>2</v></c></row>"),
    "row_gap": _sheet('<row r="2"><c><v>1</v></c></row><row><c><v>2</v></c></row><row r="9"><c><v>3</v></c></row>'),
    "empty_rows": _sheet('<row r="1"/><row r="2" ht="20" customHeight="1"/><row r="4"><c r="A4"/></row>'),
    "formulas": _sheet(
        '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="B1" t="str"><f>"a"</f><v>a</v></c></row>'
    ),
    "shared": _sheet(
        '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1</f><v>1</v></c></row>'
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>'
        '<row r="3"><c r="A3"><f t="shared" si="1"/><v>2</v></c></row>'
    ),
    "array": _sheet(
        '<row r="1"><c r="A1"><f t="array" ref="A1:B2">1</f><v>1</v></c><c r="B1"><v>1</v></c></row>'
        '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>1</v></c></row>'
    ),
    "data_table": _sheet('<row r="1"><c r="A1"><f t="dataTable" ref="A1:C3" dt2D="1"/></c></row>'),
    "merged": _sheet(
        '<row r="1"><c r="A1"><v>1</v></c></row>', tail='<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="D4:E4"/></mergeCells>'
    ),
    "inline": _sheet(
        '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> a </t></is></c>'
        '<c r="B1" t="inlineStr"><is><r><rPr><b/><family val="2"/></rPr><t>x</t></r></is></c></row>'
    ),
    "styled": _sheet('<row r="1" spans="1:3" s="2" customFormat="1"><c r="A1" s="3"/><c r="B1" s="4" t="n"><v>1</v></c></row>'),
    "unicode": _sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>żółć 😀</t></is></c></row>'),
    "no_dimension": _sheet('<row r="1"><c r="A1"><v>1</v></c></row>', head=""),
    "views_first": _sheet(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        head='<sheetPr/><dimension ref="A1:Z9"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>',
    ),
    "empty": _sheet(""),
    "far": _sheet('<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>'),
    "duplicate": _sheet('<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row><row r="1"><c r="B1"/></row>'),
    "out_of_order": _sheet('<row r="3"><c r="C3"/><c r="A3"/></row><row r="1"><c r="B1"/></row>'),
}
SELF_CLOSED = f'<?xml version="1.0"?><worksheet xmlns="{MAIN}"><dimension ref="A1"/><sheetData/></worksheet>'

# Sheets the native scan leaves to the tree.
NOT_PLAIN = {
    "prefixed": f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1"/></x:row></x:sheetData></x:worksheet>',
    "single_quotes": _sheet("<row r='1'><c r='A1'><v>1</v></c></row>"),
    "lower_case": _sheet('<row r="1"><c r="a1"><v>1</v></c></row>'),
    "dollar": _sheet('<row r="1"><c r="$A$1"><v>1</v></c></row>'),
    "leading_zero": _sheet('<row r="1"><c r="A01"><v>1</v></c></row>'),
    "comment": _sheet('<row r="1"><c r="A1"><!-- x --><v>1</v></c></row>'),
    "cdata": _sheet('<row r="1"><c r="A1" t="inlineStr"><is><t><![CDATA[<f t="shared"/>]]></t></is></c></row>'),
    "entity_in_type": _sheet('<row r="1"><c r="A1" t="&#115;"><v>0</v></c></row>'),
    "second_namespace": _sheet('<row r="1"><c r="A1" xmlns="urn:other"/></row>'),
    "row_zero": _sheet('<row r="0"><c r="A1"/></row>'),
    "column_past_xfd": _sheet('<row r="1"><c r="XFE1"/></row>'),
    "formula_outside": _sheet('<row r="1"><c r="A1"/></row>', tail="<f>1</f>"),
    "unknown_row_child": _sheet('<row r="1"><extLst/><c r="A1"/></row>'),
}


def _targets(rng, count):
    return [f"{rng.choice('ABCDEFXYZ')}{rng.choice([1, 2, 3, 4, 9, 1000])}" for _ in range(count)]


def _compare(xml, targets):
    fast = _fast_sheet(xml)
    if fast is None:
        return None
    check_well_formed(xml)
    got = _fast_scan(xml, fast, targets)
    if got is None:
        return None
    return got, _tree_scan(fromstring(xml), dict.fromkeys(targets))


@pytest.mark.parametrize("name", [*CASES, "self_closed"])
def test_fast_scan_answers_as_the_tree_does(name):
    xml = SELF_CLOSED if name == "self_closed" else CASES[name]
    refs = ["A1", "B1", "C1", "A2", "B2", "A3", "B3", "C3", "D4", "XFD1048576", "A9", "Z9", "B1000"]
    for targets in (refs, ["A1"], ["B2", "C1"], ["D4"], ["A3", "B1"]):
        pair = _compare(xml, targets)
        assert pair is not None, name
        got, want = pair
        assert got == want, (name, targets)


@pytest.mark.parametrize("name", NOT_PLAIN)
def test_markup_that_is_not_plain_goes_to_the_tree(name):
    assert _compare(NOT_PLAIN[name], ["A1", "B2"]) is None


def test_random_sheets_are_scanned_as_the_tree_does():
    rng = random.Random(2026)
    scanned = 0
    for _ in range(400):
        rows = []
        row_number = 0
        for _ in range(rng.randint(0, 6)):
            row_number += rng.choice([1, 1, 2, 5])
            numbered = rng.random() < 0.7
            cells = []
            column = 0
            for _ in range(rng.randint(0, 5)):
                column += rng.choice([1, 1, 2, 3])
                label = "ABCDEFGHIJKLMNOP"[column - 1] if column <= 16 else "Q"
                attrs = f' r="{label}{row_number}"' if numbered else ""
                if rng.random() < 0.3:
                    attrs += rng.choice([' t="s"', ' t="b"', ' t="str"', ' s="2"', ' t="inlineStr"'])
                content = rng.choice([
                    "", "<v>1</v>", "<f>1+1</f><v>2</v>", '<f t="shared" ref="A1:A9" si="0">1</f>',
                    '<f t="shared" si="0"/>', '<f t="shared" si="1"/>', '<f t="array" ref="B2:C3">1</f>',
                    '<f t="dataTable" ref="A1:B2"/>', "<is><t>x</t></is>",
                ])
                cells.append(f"<c{attrs}/>" if not content else f"<c{attrs}>{content}</c>")
            if not numbered and rng.random() < 0.5:
                cells = [cell for cell in cells if " r=" not in cell]
            rows.append(f'<row r="{row_number}">{"".join(cells)}</row>' if rng.random() < 0.8 else f'<row>{"".join(cells)}</row>')
        tail = rng.choice(["", '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'])
        xml = _sheet("".join(rows), head=rng.choice(["", '<dimension ref="A1:D9"/>']), tail=tail)
        try:
            ET.fromstring(xml)
        except ET.ParseError:
            continue
        pair = _compare(xml, _targets(rng, rng.randint(1, 4)))
        if pair is not None:
            scanned += 1
            assert pair[0] == pair[1], xml
    assert scanned > 200  # the corpus is mostly plain markup


def _book(tmp_path, rows):
    path = tmp_path / "a.xlsx"
    with XlsxWriter(str(path)) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, auto_filter=False)
    return path


def _edit(path, edits, out):
    updater = XlsxUpdater(str(path))
    try:
        updater.update_cells("S", edits)
        updater.save(str(out))
    finally:
        updater.close()
    with zipfile.ZipFile(out) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


@pytest.mark.parametrize("edits", [
    {"A1": 5},
    {"B2": "text", "C3": 1.5, "E7": Formula("A1*2")},
    {"A1": None, "B2": None},
    {"D4": 1},
])
def test_update_cells_writes_the_same_with_and_without_the_native_scan(tmp_path, monkeypatch, edits):
    path = _book(tmp_path, [[1, "a", 3], [4, "b", 6], ["c", 7, 8]])
    native = _edit(path, edits, tmp_path / "native.xlsx")
    monkeypatch.setattr(_accel, "xlsx_edit_scan", None)
    tree = _edit(path, edits, tmp_path / "tree.xlsx")
    assert native == tree


def _sheet_part(path, body):
    out = path.with_name("patched.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(out, "w") as target:
        for name in source.namelist():
            data = source.read(name)
            target.writestr(name, body.encode() if name == "xl/worksheets/sheet1.xml" else data)
    return out


@pytest.mark.parametrize("native", [True, False])
@pytest.mark.parametrize("name, message", [
    ("shared", "shared formula"),
    ("array", "array/data-table"),
    ("merged", "merged cells"),
])
def test_edits_into_formula_groups_and_merges_are_refused(tmp_path, monkeypatch, native, name, message):
    if not native:
        monkeypatch.setattr(_accel, "xlsx_edit_scan", None)
    path = _sheet_part(_book(tmp_path, [[1]]), CASES[name])
    updater = XlsxUpdater(str(path))
    try:
        with pytest.raises(ValueError, match=message):
            updater.update_cells("S", {"A2" if name != "merged" else "B2": 1})
        # Nothing of the refused edit is left.
        assert updater.to_bytes() == XlsxUpdater(str(path)).to_bytes()
    finally:
        updater.close()


@pytest.mark.parametrize("native", [True, False])
def test_a_sheet_that_is_not_well_formed_or_not_a_worksheet_is_refused_first(tmp_path, monkeypatch, native):
    if not native:
        monkeypatch.setattr(_accel, "xlsx_edit_scan", None)
    broken = _sheet('<row r="1"><c r="A1"><v>1</v></c></row')
    for body, error in [
        (broken, ET.ParseError),
        (CASES["plain"].replace("<worksheet", "<!DOCTYPE worksheet [<!ENTITY x 'y'>]><worksheet"), ET.ParseError),
        (CASES["plain"].replace("worksheet", "chartsheet"), ValueError),
    ]:
        path = _sheet_part(_book(tmp_path, [[1]]), body)
        updater = XlsxUpdater(str(path))
        try:
            # A bad reference comes after the sheet is checked, as before.
            with pytest.raises(error):
                updater.update_cells("S", {"not a ref": 1})
        finally:
            updater.close()
