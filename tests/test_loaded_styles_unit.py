"""Style-table parsing and lookup paths not reached by whole-workbook tests.

The workbook suites check what Excel sees; these tests drive the individual
readers (gradient fills, color variants, named-style records) and the table's
copy/lookup helpers with small, explicit ``styles.xml`` documents.
"""

from __future__ import annotations

import pytest

from xlspy import XlsxWriter, load_workbook
from xlspy.styles import (
    Alignment,
    Border,
    Color,
    DifferentialStyle,
    Font,
    GradientFill,
    NamedStyle,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
)
from xlspy.workbook._styles import (
    XlsxStyleTable,
    _attr_flag,
    _border_from_key,
    _color_from_key,
    _fill_from_key,
    _float,
    _font_from_key,
    _int,
    _read_color,
    _read_fill,
)
from xlspy._safe_xml import fromstring as safe_fromstring

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _xml(body: str) -> bytes:
    return f'<styleSheet xmlns="{MAIN}">{body}</styleSheet>'.encode()


def test_a_gradient_fill_survives_an_xlspy_load(tmp_path):
    path = tmp_path / "gradient.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["a"]],
            styles={
                "A1": Style(
                    fill=GradientFill(
                        "linear",
                        degree=90,
                        stop=[Stop("FF0000", 0), Stop("00FF00", 1)],
                    )
                )
            },
        )
    workbook = load_workbook(path)
    try:
        fill = workbook["Data"]["A1"].fill
    finally:
        workbook.close()
    assert isinstance(fill, GradientFill)
    assert fill.fill_type == "linear"
    assert fill.degree == 90
    assert [(stop.position, stop.color.rgb) for stop in fill.stop] == [
        (0, "FFFF0000"),
        (1, "FF00FF00"),
    ]


def test_a_gradient_without_stops_becomes_a_plain_fill():
    fill = _read_fill(
        safe_fromstring(
            f'<fill xmlns="{MAIN}"><gradientFill type="path"/></fill>'
        )
    )
    assert fill.fill_type == "none"


def test_the_style_table_keeps_named_style_records():
    table = XlsxStyleTable.from_xml(
        _xml(
            "<fonts><font><sz val='11'/><name val='Calibri'/></font></fonts>"
            "<fills><fill><patternFill patternType='none'/></fill>"
            "<fill><patternFill patternType='gray125'/></fill></fills>"
            "<borders><border/></borders>"
            "<cellStyleXfs count='2'>"
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'/>"
            "<xf numFmtId='0' fontId='0' fillId='0' borderId='0'/>"
            "</cellStyleXfs>"
            "<cellStyles count='3'>"
            "<cellStyle name='Normal' xfId='0' builtinId='0'/>"
            "<cellStyle name='Good' xfId='1' builtinId='26'/>"
            "<cellStyle/>"
            "</cellStyles>"
            "<cellXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/></cellXfs>"
        )
    )
    assert table.named_styles == ["Normal", "Good"]
    assert table.named_style_defs["Good"].builtinId == 26
    assert table.named_style_defs["Normal"].font.name == "Calibri"


def test_a_minimal_style_table_gets_excel_defaults():
    table = XlsxStyleTable.from_xml(_xml(""))
    assert [font.name for font in table.fonts] == ["Calibri"]
    assert [fill.fill_type for fill in table.fills] == ["none", "gray125"]
    assert len(table.borders) == 1
    assert len(table.styles) == 1
    assert table.named_styles == ["Normal"]


def test_an_empty_document_uses_the_empty_table():
    table = XlsxStyleTable.from_xml(None)
    assert table.fonts[0].name == "Calibri"
    assert table.xfs[0].font_id == 0


def test_the_style_table_clone_is_independent():
    table = XlsxStyleTable.from_xml(_xml(""))
    clone = table.clone()
    clone.fonts.append(None)
    clone.fills.append(None)
    clone.num_formats[999] = "0"
    clone.dxfs.append(DifferentialStyle())
    assert len(table.fonts) == 1
    assert len(table.fills) == 2
    assert 999 not in table.num_formats
    assert table.dxfs == []
    assert clone._root is not table._root


def test_differential_lookups_are_bounds_checked():
    table = XlsxStyleTable.empty()
    assert table.dxf_at(None) is None
    assert table.dxf_at(0) is None
    style = DifferentialStyle(number_format="0.00")
    table.dxfs.append(style)
    assert table.dxf_at(0) is style
    assert table.dxf_at(1) is None
    assert table.dxf_at(-1) is None


@pytest.mark.parametrize(
    "key, expected",
    [
        (None, None),
        (("rgb", "FF112233"), "FF112233"),
        (("indexed", 4), 4),
        (("auto",), True),
        (("theme", 2, 0.5), ("theme", 2, 0.5)),
    ],
)
def test_colors_are_rebuilt_from_their_keys(key, expected):
    color = _color_from_key(key)
    if key is None:
        assert color is None
    elif key[0] == "rgb":
        assert color.rgb == expected
    elif key[0] == "indexed":
        assert color.indexed == expected
    elif key[0] == "auto":
        assert color.auto is True
    else:
        assert (color.theme, color.tint) == expected[1:]


def test_absent_component_keys_return_none():
    assert _fill_from_key(None) is None
    assert _border_from_key(None) is None
    assert _font_from_key(None) is None


def test_numeric_helpers_ignore_values_that_are_not_numbers():
    assert _float("1.5") == 1.5
    assert _float("nope") is None
    assert _int("3") == 3
    assert _int("nope") is None
    assert _attr_flag(None, "auto") is None
    assert _attr_flag(safe_fromstring('<c auto="1"/>'), "auto") is True


def test_the_style_table_appends_cell_formats_and_reuses_equal_ones():
    table = XlsxStyleTable.empty()
    style = Style(
        font=Font(name="Arial", size=12, bold=True, color=Color(rgb="FF112233")),
        fill=PatternFill("solid", start_color="FFFF00"),
        border=Border(left=Side(style="thin", color="FF000000"), diagonal_up=True, outline=False),
        number_format="0.00",
        alignment=Alignment(horizontal="center", wrap_text=True),
        protection=Protection(locked=False, hidden=True),
    )
    index = table.add(style)
    assert index == 1
    assert table.add(style) == index
    assert table.style_at(index).number_format == "0.00"
    assert table.style_at(999) is table.styles[0]
    data = table.to_bytes()
    assert b'<cellXfs count="2">' in data
    assert b'applyNumberFormat="1"' in data
    assert b'applyAlignment="1"' in data
    assert b'applyProtection="1"' in data


def test_named_styles_are_appended_to_the_style_and_name_containers():
    table = XlsxStyleTable.empty()
    named = NamedStyle(
        "Fancy",
        font=Font(name="Georgia", italic=True),
        fill=PatternFill("solid", start_color="FF00FF00"),
        number_format="0.00",
        alignment=Alignment(horizontal="right"),
        protection=Protection(locked=False),
        builtinId=27,
        hidden=True,
    )
    xf_index = table.add_named_style(named)
    assert xf_index == 1
    assert table.named_style_defs["Fancy"] is named
    data = table.to_bytes()
    assert b'name="Fancy"' in data
    assert b'builtinId="27"' in data
    assert b'hidden="1"' in data
    assert b'<cellStyles count="2">' in data


def test_differential_formats_are_appended_and_shared():
    table = XlsxStyleTable.empty()
    fragment = f'<dxf xmlns="{MAIN}"><font><b/></font></dxf>'
    assert table.append_dxfs([fragment, fragment]) == [0, 0]
    assert len(table.dxfs) == 1
    assert table.dxfs[0] is not None
    assert b'<dxfs count="1">' in table.to_bytes()


def test_colors_are_read_from_every_attribute():
    assert _read_color(safe_fromstring('<color rgb="FF112233"/>')).rgb == "FF112233"
    assert _read_color(safe_fromstring('<color indexed="4"/>')).indexed == 4
    assert _read_color(safe_fromstring('<color theme="2" tint="0.5"/>')).tint == 0.5
    assert _read_color(safe_fromstring('<color theme="2" tint="0"/>')).tint is None
    assert _read_color(safe_fromstring('<color auto="1"/>')).auto is True
    assert _read_color(safe_fromstring("<color/>")) is None
    assert _read_color(None) is None
