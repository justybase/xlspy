"""Inline fonts, rich-text runs and the ``<si>`` readers.

The workbook and writer suites exercise rich text end to end; these tests pin
the individual run properties, the XML escaping and every parsing branch that
does not need a whole package.
"""

from __future__ import annotations

import pytest
from xml.etree import ElementTree as ET

from xlspy.rich_text import (
    CellRichText,
    InlineFont,
    TextBlock,
    _color_attrs,
    _font_from_rpr,
    _quote,
    parse_rich_shared_strings,
    rich_text_from_si,
)
from xlspy.styles import Color


def test_inline_font_refuses_unknown_underline_and_vertical_alignment():
    with pytest.raises(ValueError, match="u must be one of"):
        InlineFont(u="wavy")
    with pytest.raises(ValueError, match="vertAlign must be one of"):
        InlineFont(vertAlign="bogus")


@pytest.mark.parametrize("size", [True, 0, 0.5, 410, "12"])
def test_inline_font_refuses_an_invalid_size(size):
    with pytest.raises(ValueError, match="between 1 and 409"):
        InlineFont(sz=size)


def test_inline_font_normalizes_size_and_color():
    font = InlineFont(sz=12, color="c00000")
    assert font.sz == 12.0
    assert font.color == Color(rgb="FFC00000")
    assert InlineFont().color is None


def test_inline_font_writes_every_run_property():
    font = InlineFont(
        rFont="Arial",
        b=True,
        i=True,
        u="double",
        strike=True,
        sz=14,
        color=Color(indexed=3),
        vertAlign="superscript",
        family=2,
        charset=1,
        scheme="minor",
        outline=True,
        shadow=True,
        condense=True,
        extend=True,
    )
    xml = font._xml("x:")
    for fragment in (
        "<x:b/>",
        "<x:i/>",
        "<x:strike/>",
        "<x:condense/>",
        "<x:extend/>",
        "<x:outline/>",
        "<x:shadow/>",
        '<x:u val="double"/>',
        '<x:vertAlign val="superscript"/>',
        '<x:sz val="14"/>',
        '<x:color indexed="3"/>',
        '<x:rFont val="Arial"/>',
        '<x:family val="2"/>',
        '<x:charset val="1"/>',
        '<x:scheme val="minor"/>',
    ):
        assert fragment in xml
    assert xml.startswith("<x:rPr>") and xml.endswith("</x:rPr>")


def test_an_empty_inline_font_writes_no_run_properties():
    assert InlineFont()._xml("") == ""
    assert InlineFont(u="single", vertAlign="subscript")._xml("") == ('<rPr><u/><vertAlign val="subscript"/></rPr>')


def test_text_block_requires_a_font_and_text():
    with pytest.raises(TypeError, match="font must be an InlineFont"):
        TextBlock("not a font", "text")
    with pytest.raises(TypeError, match="text must be a str"):
        TextBlock(InlineFont(), 3)


def test_rich_text_container_helpers():
    rich = CellRichText("a", TextBlock(InlineFont(b=True), "b"))
    rich.extend(["c", TextBlock(InlineFont(i=True), "d")])
    assert str(rich) == "abcd"
    assert rich.as_list() == ["a", "b", "c", "d"]
    assert repr(rich).startswith("CellRichText([")
    with pytest.raises(TypeError, match="str or TextBlock"):
        rich.append(3)


def test_runs_xml_skips_empty_parts_and_escapes_text():
    rich = CellRichText(" lead ", "", TextBlock(InlineFont(b=True), ""), 'a<b&"c"')
    xml = rich.runs_xml()
    assert '<t xml:space="preserve"> lead </t>' in xml
    assert 'a&lt;b&amp;"c"' in xml
    assert xml.count("<r>") == 2
    assert CellRichText("").runs_xml() == "<t/>"


@pytest.mark.parametrize(
    "key, expected",
    [
        (("rgb", "FF112233"), 'rgb="FF112233"'),
        (("indexed", 4), 'indexed="4"'),
        (("auto",), 'auto="1"'),
        (("theme", 2, None), 'theme="2"'),
        (("theme", 2, 0.5), 'theme="2" tint="0.5"'),
    ],
)
def test_color_attributes(key, expected):
    assert _color_attrs(key) == expected


def test_quotes_are_xml_escaped():
    assert _quote('a"b<c') == '"a&quot;b&lt;c"'


def _si(xml: str):
    return ET.fromstring(xml)


def test_a_plain_shared_string_reads_as_text():
    assert rich_text_from_si(_si("<si><t>a_x0041_b</t></si>")) == "aAb"


def test_runs_without_font_properties_stay_plain_strings():
    rich = rich_text_from_si(_si("<si><r><rPr/><t>a</t></r><r><t>b</t></r></si>"))
    assert rich == CellRichText("a", "b")


def test_phonetic_runs_are_skipped():
    xml = "<si><r><t>base</t></r><rPh sb='0' eb='1'><t>phonetic</t></rPh></si>"
    assert rich_text_from_si(_si(xml)) == CellRichText("base")


def test_run_font_properties_are_read_back():
    xml = (
        "<si><r><rPr>"
        "<b/><i val='0'/><strike/><outline/><shadow/><condense/><extend/>"
        "<u val='double'/><vertAlign val='superscript'/>"
        "<sz val='14'/><color rgb='FFC00000'/>"
        "<rFont val='Arial'/><family val='2'/><charset val='1'/><scheme val='minor'/>"
        "</rPr><t>run</t></r></si>"
    )
    (run,) = rich_text_from_si(_si(xml))
    font = run.font
    assert font.b is True and font.i is False
    assert (font.strike, font.outline, font.shadow, font.condense, font.extend) == (
        True,
        True,
        True,
        True,
        True,
    )
    assert font.u == "double"
    assert font.vertAlign == "superscript"
    assert font.sz == 14.0
    assert font.color == Color(rgb="FFC00000")
    assert (font.rFont, font.family, font.charset, font.scheme) == (
        "Arial",
        2,
        1,
        "minor",
    )


@pytest.mark.parametrize(
    "properties, expected",
    [
        ("<u val='none'/><sz val='500'/><color auto='1'/>", InlineFont()),
        ("<color theme='4' tint='0.25'/>", InlineFont(color=Color(theme=4, tint=0.25))),
        ("<color indexed='7'/>", InlineFont(color=Color(indexed=7))),
        ("<color rgb='FF0000FF'/>", InlineFont(color=Color(rgb="FF0000FF"))),
    ],
)
def test_run_font_edge_properties(properties, expected):
    rpr = _si(f"<rPr>{properties}</rPr>")
    assert _font_from_rpr(rpr) == expected


def test_shared_string_table_reads_plain_and_rich_items():
    xml = "<sst><si><t>plain</t></si><si><r><rPr><b/></rPr><t>bold</t></r></si></sst>"
    values = parse_rich_shared_strings(xml.encode("utf-8"))
    assert values[0] == "plain"
    assert isinstance(values[1], CellRichText)
    assert values[1][0].font.b is True
