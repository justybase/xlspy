"""Shared strings of an XLSX workbook must come out the same on every parser: the
native one (where it answers), the regex one and the expat one. Documents the
native parser refuses must still read through the other two."""

import datetime

import pytest

from test_comments_and_editing import replace_parts
from xlspy import ExcelReader, XlsxWriter, _accel
from xlspy import reader as reader_module

SST = "xl/sharedStrings.xml"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def _sst(body, decl=DECL, root=f'<sst xmlns="{MAIN_NS}" count="1" uniqueCount="1">'):
    return (decl + root + body + "</sst>").encode("utf-8")


def _workbook(tmp_path, data):
    path = tmp_path / "sst.xlsx"
    with XlsxWriter(path) as writer:
        writer.write_sheet([["x", 1, datetime.datetime(2024, 1, 5)]])
    replace_parts(path, {SST: data})
    return str(path)


def _strings(path, monkeypatch, *, native=True, regex=True):
    """The shared strings ExcelReader reads, with the chosen parsers switched off."""
    with monkeypatch.context() as patch:
        if not native:
            patch.setattr(reader_module, "_accel_read_xlsx_sst", lambda data: None)
        if not regex:
            patch.setattr(reader_module, "_parse_shared_strings_fast", lambda content: None)
        with ExcelReader(path) as reader:
            return list(reader._shared_string_array)


SUPPORTED = {
    "plain": "<si><t>a</t></si><si><t>b c</t></si>",
    "empty_items": "<si></si><si><t></t></si>",
    "space_preserve": '<si><t xml:space="preserve"> x </t></si>',
    "rich_runs": "<si><r><rPr><b/></rPr><t>bold</t></r><r><t> plain</t></r></si>",
    "phonetic_run": (
        '<si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="0" type="noConversion"/></si>'
    ),
    "entities": "<si><t>a &amp; b &lt; c &gt; &quot;d&quot; &apos;e&apos;</t></si>",
    "numeric_references": "<si><t>&#65;&#x42;&#x1F600;&#xe9;</t></si>",
    "carriage_returns": "<si><t>one\r\ntwo\rthree</t></si>",
    "referenced_carriage_return": "<si><t>x&#13;y&#xD;\r\n</t></si>",
    "xstring_escapes": "<si><t>_x0041_ &amp; _x000D_ _x005F_x0041_</t></si>",
    "escaped_underscore": "<si><t>&#x5F;x0041_</t></si>",
    "split_escape_across_runs": "<si><r><t>_x00</t></r><r><t>41_</t></r></si>",
    "unicode": "<si><t>ąęźż 😀</t></si>",
    "attribute_with_angle_bracket": '<si><rPh sb="0" x="a&gt;b"><t>k</t></rPh><t>v</t></si>',
    "whitespace_text": "<si><t>  </t></si>",
}

UNSUPPORTED = {
    "cdata": "<si><t><![CDATA[x]]></t></si>",
    "comment": "<si><!-- note --><t>x</t></si>",
    "self_closing_item": "<si/><si><t>a</t></si>",
    "self_closing_text": "<si><t/></si>",
    "text_with_attributes": '<si><t xml:lang="en">a</t></si>',
    "nested_markup_in_text": "<si><t>a<b/>c</t></si>",
    "unknown_entity": "<si><t>&nbsp;</t></si>",
    "control_character_reference": "<si><t>&#1;</t></si>",
    "surrogate_reference": "<si><t>&#xD800;</t></si>",
    "unterminated_reference": "<si><t>a &amp b</t></si>",
    "stray_text_close": "<si><t>a</t></t></si>",
    "unbalanced_item": "<si><t>a</t>",
}


@pytest.mark.parametrize("name", sorted(SUPPORTED))
def test_shared_strings_agree_on_every_parser(tmp_path, monkeypatch, name):
    path = _workbook(tmp_path, _sst(SUPPORTED[name]))
    default = _strings(path, monkeypatch)
    regex_only = _strings(path, monkeypatch, native=False)
    expat_only = _strings(path, monkeypatch, native=False, regex=False)
    assert default == regex_only == expat_only


@pytest.mark.parametrize("name", sorted(UNSUPPORTED))
def test_unsupported_shared_strings_still_read_through_the_other_parsers(tmp_path, monkeypatch, name):
    path = _workbook(tmp_path, _sst(UNSUPPORTED[name]))
    try:
        expected = _strings(path, monkeypatch, native=False)
    except Exception as error:  # The XML parser's own error is the reference.
        with pytest.raises(type(error)):
            _strings(path, monkeypatch)
        return
    assert _strings(path, monkeypatch) == expected


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
@pytest.mark.parametrize("declaration", [
    '<?xml version="1.0" encoding="UTF-16"?>',
    '<?xml version="1.0" encoding = "utf-8"?>',
])
def test_native_parser_declines_other_declarations(declaration):
    assert _accel.read_xlsx_shared_strings(_sst(SUPPORTED["plain"], decl=declaration)) is None


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
@pytest.mark.parametrize("name", sorted(SUPPORTED))
def test_native_parser_answers_the_supported_documents(name):
    assert _accel.read_xlsx_shared_strings(_sst(SUPPORTED[name])) is not None


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
@pytest.mark.parametrize("name", sorted(UNSUPPORTED))
def test_native_parser_declines_the_unsupported_documents(name):
    assert _accel.read_xlsx_shared_strings(_sst(UNSUPPORTED[name])) is None


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
def test_native_parser_declines_invalid_utf8():
    data = _sst(SUPPORTED["plain"]).replace(b"b c", b"b \xff")
    assert _accel.read_xlsx_shared_strings(data) is None
