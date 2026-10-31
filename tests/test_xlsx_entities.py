"""Entities, carriage returns and character references in XLSX cell values must read
the same on every path: the native scanner, the Python fast path and ElementTree.

The guard tests cover text the fast paths already hand to ElementTree. The xfail
tests cover typed values (t="e", t="d", t="b", numbers) with entities, which the
native scanner did not decode before C1.
"""

import pytest

from test_comments_and_editing import replace_parts
from test_xlsx_fast_path import SHEET, _all_paths, _base_file, _same_typed_rows, _sheet

ENTITY_BODIES = {
    "entities_in_inline_text": (
        '<row r="1"><c r="A1" t="inlineStr"><is><t>a &amp; b &lt; c &gt; d &quot;q&quot; &apos;s&apos;</t></is></c></row>'
    ),
    "carriage_returns_in_inline_text": '<row r="1"><c r="A1" t="inlineStr"><is><t>a\r\nb\rc</t></is></c></row>',
    "character_references_in_inline_text": (
        '<row r="1"><c r="A1" t="inlineStr"><is><t>x&#13;y&#x41;&#66;</t></is></c>'
        '<c r="B1" t="inlineStr"><is><t>&#x5F;x0041_</t></is></c></row>'
    ),
    "entities_in_str_cells": (
        '<row r="1"><c r="A1" t="str"><v>x &amp; &#65;</v></c><c r="B1" t="str"><v>a\r\nb</v></c></row>'
    ),
}

TYPED_ENTITY_BODIES = {
    "error_with_entity": '<row r="1"><c r="A1" t="e"><v>#N&amp;A</v></c></row>',
    "date_with_entity": '<row r="1"><c r="A1" t="d"><v>2024-01-05T10&#58;30&#58;00</v></c></row>',
    "boolean_with_entity": '<row r="1"><c r="A1" t="b"><v>&#49;</v></c></row>',
    "number_with_entity": '<row r="1"><c r="A1"><v>&#49;2</v></c></row>',
}


def _read_all(path, monkeypatch):
    from_c, from_python, from_tree = _all_paths(path, monkeypatch)
    _same_typed_rows(from_c, from_tree)
    _same_typed_rows(from_python, from_tree)


@pytest.mark.parametrize("name", sorted(ENTITY_BODIES))
def test_entity_text_reads_the_same_on_every_path(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(ENTITY_BODIES[name])})
    _read_all(path, monkeypatch)


@pytest.mark.xfail(strict=True, reason="native reader does not decode entities in typed values yet (C1)")
@pytest.mark.parametrize("name", sorted(TYPED_ENTITY_BODIES))
def test_typed_values_with_entities_read_the_same_on_every_path(tmp_path, monkeypatch, name):
    path = _base_file(tmp_path)
    replace_parts(path, {SHEET: _sheet(TYPED_ENTITY_BODIES[name])})
    _read_all(path, monkeypatch)
