"""XLSX style-table reader/writer used by the openpyxl-like workbook model.

The fast writers keep their own compact registries.  This module is only used
by :mod:`xlspy.workbook`, where loading an existing file has to preserve the
style components that are already in ``xl/styles.xml``.  New styles are
appended to that table; existing component XML is left untouched.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .._metadata import MAIN
from ..styles import (
    Alignment,
    Border,
    Color,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    _border_key,
    _color_key,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _pattern_fill_key,
    _protection_key_from,
    _alignment_key_from,
    split_style,
)
from ..xlsx_writer import _alignment_xml, _format_number, _protection_xml

# ECMA-376 / openpyxl's built-in number formats.  Only the codes that are
# useful to callers are listed; unknown built-ins fall back to "General".
BUILTIN_NUM_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: "$#,##0_);($#,##0)",
    6: "$#,##0_);[Red]($#,##0)",
    7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}


def _q(tag: str) -> str:
    return f"{{{MAIN}}}{tag}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _flag(node: ET.Element | None, name: str = "val") -> bool | None:
    if node is None:
        return None
    value = node.get(name)
    if value is None:
        return True
    return value.lower() not in {"0", "false", "off"}


def _attr_flag(node: ET.Element | None, name: str) -> bool | None:
    if node is None:
        return None
    value = node.get(name)
    if value is None:
        return None
    return value.lower() not in {"0", "false", "off"}


def _float(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _int(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _read_color(node: ET.Element | None) -> Color | None:
    if node is None:
        return None
    rgb = node.get("rgb")
    if rgb:
        return Color(rgb=rgb)
    indexed = _int(node.get("indexed"))
    if indexed is not None:
        return Color(indexed=indexed)
    theme = _int(node.get("theme"))
    if theme is not None:
        tint = _float(node.get("tint"))
        return Color(theme=theme, tint=tint if tint not in (None, 0.0) else None)
    if _attr_flag(node, "auto"):
        return Color(auto=True)
    return None


def _read_font(node: ET.Element) -> Font:
    color = _read_color(node.find(_q("color")))
    name_node = node.find(_q("name"))
    size_node = node.find(_q("sz"))
    underline_node = node.find(_q("u"))
    underline: str | None = None
    if underline_node is not None:
        underline = underline_node.get("val") or "single"
    vert = node.find(_q("vertAlign"))
    family = node.find(_q("family"))
    charset = node.find(_q("charset"))
    scheme = node.find(_q("scheme"))
    return Font(
        name=name_node.get("val") if name_node is not None else None,
        size=_float(size_node.get("val")) if size_node is not None else None,
        bold=_flag(node.find(_q("b"))),
        italic=_flag(node.find(_q("i"))),
        strike=_flag(node.find(_q("strike"))),
        underline=underline,
        vertAlign=vert.get("val") if vert is not None else None,
        color=color,
        charset=_int(charset.get("val")) if charset is not None else None,
        family=_int(family.get("val")) if family is not None else None,
        scheme=scheme.get("val") if scheme is not None else None,
        outline=_flag(node.find(_q("outline"))),
        shadow=_flag(node.find(_q("shadow"))),
        condense=_flag(node.find(_q("condense"))),
        extend=_flag(node.find(_q("extend"))),
    )


def _read_fill(node: ET.Element) -> PatternFill | GradientFill:
    pattern = node.find(_q("patternFill"))
    if pattern is not None:
        return PatternFill(
            fill_type=pattern.get("patternType"),
            start_color=_read_color(pattern.find(_q("fgColor"))),
            end_color=_read_color(pattern.find(_q("bgColor"))),
        )
    gradient = node.find(_q("gradientFill"))
    if gradient is not None:
        stops = []
        for stop in gradient.findall(_q("stop")):
            color = _read_color(stop.find(_q("color")))
            position = _float(stop.get("position"))
            if color is not None and position is not None:
                stops.append(Stop(color=color, position=position))
        kwargs = {
            "fill_type": gradient.get("type", "linear"),
            "degree": _float(gradient.get("degree")),
            "left": _float(gradient.get("left")),
            "right": _float(gradient.get("right")),
            "top": _float(gradient.get("top")),
            "bottom": _float(gradient.get("bottom")),
            "stop": tuple(stops),
        }
        if not stops:
            # A gradient without stops is invalid; keep it as a plain fill.
            return PatternFill(fill_type="none")
        return GradientFill(**kwargs)
    return PatternFill(fill_type="none")


def _read_side(node: ET.Element | None) -> Side | None:
    if node is None:
        return None
    style = node.get("style")
    color = _read_color(node.find(_q("color")))
    if style is None and color is None:
        return None
    return Side(style=style, color=color)


def _read_border(node: ET.Element) -> Border:
    return Border(
        left=_read_side(node.find(_q("left"))),
        right=_read_side(node.find(_q("right"))),
        top=_read_side(node.find(_q("top"))),
        bottom=_read_side(node.find(_q("bottom"))),
        diagonal=_read_side(node.find(_q("diagonal"))),
        vertical=_read_side(node.find(_q("vertical"))),
        horizontal=_read_side(node.find(_q("horizontal"))),
        start=_read_side(node.find(_q("start"))),
        end=_read_side(node.find(_q("end"))),
        diagonal_up=_attr_flag(node, "diagonalUp"),
        diagonal_down=_attr_flag(node, "diagonalDown"),
        outline=_attr_flag(node, "outline"),
    )


def _read_alignment(node: ET.Element | None) -> Alignment | None:
    if node is None:
        return None
    return Alignment(
        horizontal=node.get("horizontal"),
        vertical=node.get("vertical"),
        wrap_text=_attr_flag(node, "wrapText"),
        shrink_to_fit=_attr_flag(node, "shrinkToFit"),
        indent=_int(node.get("indent")),
        text_rotation=_int(node.get("textRotation")),
        relative_indent=_int(node.get("relativeIndent")),
        justify_last_line=_attr_flag(node, "justifyLastLine"),
        reading_order=_int(node.get("readingOrder")),
    )


def _read_protection(node: ET.Element | None) -> Protection | None:
    if node is None:
        return None
    return Protection(
        locked=_attr_flag(node, "locked"),
        hidden=_attr_flag(node, "hidden"),
    )


def _color_from_key(key) -> Color | None:
    if key is None:
        return None
    kind = key[0]
    if kind == "rgb":
        return Color(rgb=key[1])
    if kind == "indexed":
        return Color(indexed=key[1])
    if kind == "auto":
        return Color(auto=True)
    theme, tint = key[1], key[2]
    return Color(theme=theme, tint=tint)


def _side_from_key(key) -> Side | None:
    style, color = key
    if style is None and color is None:
        return None
    return Side(style=style, color=_color_from_key(color))


def _border_from_key(key) -> Border | None:
    if key is None:
        return None
    return Border(
        left=_side_from_key(key.left),
        right=_side_from_key(key.right),
        top=_side_from_key(key.top),
        bottom=_side_from_key(key.bottom),
        diagonal=_side_from_key(key.diagonal),
        vertical=_side_from_key(key.vertical),
        horizontal=_side_from_key(key.horizontal),
        start=_side_from_key(key.start),
        end=_side_from_key(key.end),
        diagonal_up=key.diagonal_up,
        diagonal_down=key.diagonal_down,
        outline=key.outline,
    )


def _font_from_key(key) -> Font | None:
    if key is None:
        return None
    return Font(
        name=key.name,
        size=key.size,
        bold=key.bold,
        italic=key.italic,
        strike=key.strike,
        underline=key.underline,
        vertAlign=key.vert_align,
        color=_color_from_key(key.color),
        charset=key.charset,
        family=key.family,
        scheme=key.scheme,
        outline=key.outline,
        shadow=key.shadow,
        condense=key.condense,
        extend=key.extend,
    )


def _fill_from_key(key):
    if key is None:
        return None
    if key[0] == "gradient":
        _, fill_type, degree, left, right, top, bottom, stops = key
        return GradientFill(
            fill_type=fill_type,
            degree=degree,
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            stop=tuple(Stop(_color_from_key(color), position) for color, position in stops),
        )
    _, pattern, fg, bg = key
    return PatternFill(
        fill_type=pattern,
        start_color=_color_from_key(fg),
        end_color=_color_from_key(bg),
    )


def _alignment_from_key(key) -> Alignment | None:
    if key is None or not any(value is not None for value in key):
        return None
    return Alignment(
        horizontal=key[0], vertical=key[1], wrap_text=key[2],
        shrink_to_fit=key[3], indent=key[4], text_rotation=key[5],
        relative_indent=key[6], justify_last_line=key[7],
        reading_order=key[8],
    )


def _protection_from_key(key) -> Protection | None:
    if key is None or (key[0] is None and key[1] is None):
        return None
    return Protection(locked=key[0], hidden=key[1])


@dataclass
class _Xf:
    font_id: int = 0
    fill_id: int = 0
    border_id: int = 0
    num_fmt_id: int = 0
    alignment: Alignment | None = None
    protection: Protection | None = None


class XlsxStyleTable:
    """A parsed ``xl/styles.xml`` table with append-only style creation."""

    def __init__(self) -> None:
        self._root: ET.Element | None = None
        self.fonts: list[Font | None] = []
        self.fills: list[PatternFill | GradientFill | None] = []
        self.borders: list[Border | None] = []
        self.num_formats: dict[int, str] = dict(BUILTIN_NUM_FORMATS)
        self.custom_num_formats: set[int] = set()
        self.xfs: list[_Xf] = []
        self.styles: list[Style] = []
        self._font_keys: list[object | None] = []
        self._fill_keys: list[object | None] = []
        self._border_keys: list[object | None] = []
        self._font_index: dict[object, int] = {}
        self._fill_index: dict[object, int] = {}
        self._border_index: dict[object, int] = {}
        self._numfmt_by_code: dict[str, int] = {}
        self._xf_index: dict[tuple, int] = {}
        self._next_numfmt_id = 164
        self._font_xml: list[str] = []
        self._fill_xml: list[str] = []
        self._border_xml: list[str] = []
        self._xf_xml: list[str] = []

    # -- construction -------------------------------------------------
    @classmethod
    def empty(cls) -> "XlsxStyleTable":
        table = cls()
        table._add_default_components()
        return table

    @classmethod
    def from_xml(cls, xml: bytes | str | None) -> "XlsxStyleTable":
        if not xml:
            return cls.empty()
        table = cls()
        root = ET.fromstring(xml)
        table._root = root
        table._source = xml
        table._parse(root)
        if not table.xfs:
            table._add_default_components()
        return table

    def _add_default_components(self) -> None:
        default_font = Font(name="Calibri", size=11, family=2, scheme="minor", color=Color(theme=1))
        self.fonts = [default_font]
        self.fills = [PatternFill(fill_type="none"), PatternFill(fill_type="gray125")]
        self.borders = [Border()]
        self._font_keys = [_font_key_from(default_font)]
        self._fill_keys = [_pattern_fill_key(self.fills[0]), _pattern_fill_key(self.fills[1])]
        self._border_keys = [None]
        self._font_xml = [self._font_xml_for(_font_key_from(default_font))]
        self._fill_xml = [
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
        ]
        self._border_xml = ["<border><left/><right/><top/><bottom/><diagonal/></border>"]
        self.xfs = [_Xf()]
        self.styles = [Style(font=default_font)]
        self._xf_xml = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
        self._rebuild_indexes()

    def _parse(self, root: ET.Element) -> None:
        numfmts = root.find(_q("numFmts"))
        if numfmts is not None:
            for node in numfmts.findall(_q("numFmt")):
                num_id = _int(node.get("numFmtId"))
                code = node.get("formatCode")
                if num_id is not None and code is not None:
                    self.num_formats[num_id] = code
                    self.custom_num_formats.add(num_id)
                    self._next_numfmt_id = max(self._next_numfmt_id, num_id + 1)

        fonts = root.find(_q("fonts"))
        if fonts is not None:
            for node in fonts.findall(_q("font")):
                font = _read_font(node)
                key = _font_key_from(font)
                if _font_is_empty(key):
                    key = None
                self.fonts.append(font)
                self._font_keys.append(key)
                self._font_xml.append(ET.tostring(node, encoding="unicode"))

        fills = root.find(_q("fills"))
        if fills is not None:
            for node in fills.findall(_q("fill")):
                fill = _read_fill(node)
                key = _fill_key_from(fill)
                self.fills.append(fill)
                self._fill_keys.append(key)
                self._fill_xml.append(ET.tostring(node, encoding="unicode"))

        borders = root.find(_q("borders"))
        if borders is not None:
            for node in borders.findall(_q("border")):
                border = _read_border(node)
                key = _border_key(border)
                self.borders.append(border)
                self._border_keys.append(key)
                self._border_xml.append(ET.tostring(node, encoding="unicode"))

        if not self.fonts:
            self.fonts = [Font(name="Calibri", size=11, family=2, scheme="minor", color=Color(theme=1))]
            self._font_keys = [_font_key_from(self.fonts[0])]
            self._font_xml = [self._font_xml_for(self._font_keys[0])]
        if not self.fills:
            self.fills = [PatternFill(fill_type="none"), PatternFill(fill_type="gray125")]
            self._fill_keys = [_pattern_fill_key(self.fills[0]), _pattern_fill_key(self.fills[1])]
            self._fill_xml = [
                '<fill><patternFill patternType="none"/></fill>',
                '<fill><patternFill patternType="gray125"/></fill>',
            ]
        if not self.borders:
            self.borders = [Border()]
            self._border_keys = [None]
            self._border_xml = ["<border><left/><right/><top/><bottom/><diagonal/></border>"]

        xfs = root.find(_q("cellXfs"))
        if xfs is not None:
            for node in xfs.findall(_q("xf")):
                xf = _Xf(
                    font_id=_int(node.get("fontId")) or 0,
                    fill_id=_int(node.get("fillId")) or 0,
                    border_id=_int(node.get("borderId")) or 0,
                    num_fmt_id=_int(node.get("numFmtId")) or 0,
                    alignment=_read_alignment(node.find(_q("alignment"))),
                    protection=_read_protection(node.find(_q("protection"))),
                )
                self.xfs.append(xf)
                self._xf_xml.append(ET.tostring(node, encoding="unicode"))
        self._rebuild_indexes()

    def _rebuild_indexes(self) -> None:
        self._font_index = {}
        for index, key in enumerate(self._font_keys):
            if key is not None and key not in self._font_index:
                self._font_index[key] = index
        self._fill_index = {}
        for index, key in enumerate(self._fill_keys):
            if key is not None and key not in self._fill_index:
                self._fill_index[key] = index
        self._border_index = {None: 0}
        for index, key in enumerate(self._border_keys):
            if key is not None and key not in self._border_index:
                self._border_index[key] = index
        self._numfmt_by_code = {
            code: num_id for num_id, code in self.num_formats.items()
        }
        self._xf_index = {}
        self.styles = []
        for index, xf in enumerate(self.xfs):
            align_key = _alignment_key_from(xf.alignment) if xf.alignment is not None else None
            prot_key = _protection_key_from(xf.protection) if xf.protection is not None else None
            if align_key is not None and all(value is None for value in align_key):
                align_key = None
            if prot_key is not None and all(value is None for value in prot_key):
                prot_key = None
            key = (
                xf.font_id, xf.fill_id, xf.border_id, xf.num_fmt_id,
                align_key, prot_key,
            )
            self._xf_index[key] = index
            font_key = self._font_keys[xf.font_id] if xf.font_id < len(self._font_keys) else None
            fill_key = self._fill_keys[xf.fill_id] if xf.fill_id < len(self._fill_keys) else None
            border_key = self._border_keys[xf.border_id] if xf.border_id < len(self._border_keys) else None
            font = self.fonts[xf.font_id] if font_key is not None and xf.font_id < len(self.fonts) else None
            fill = self.fills[xf.fill_id] if fill_key is not None and xf.fill_id < len(self.fills) else None
            border = self.borders[xf.border_id] if border_key is not None and xf.border_id < len(self.borders) else None
            fmt = None if xf.num_fmt_id == 0 else self.num_formats.get(xf.num_fmt_id)
            self.styles.append(
                Style(
                    font=font,
                    fill=fill,
                    border=border,
                    alignment=xf.alignment,
                    protection=xf.protection,
                    number_format=fmt,
                )
            )

    # -- internals ----------------------------------------------------
    def _intern_font(self, key) -> int:
        if key is None:
            return 0
        if key in self._font_index:
            return self._font_index[key]
        index = len(self.fonts)
        self.fonts.append(_font_from_key(key))
        self._font_keys.append(key)
        self._font_xml.append(self._font_xml_for(key))
        self._font_index[key] = index
        return index

    def _intern_fill(self, key) -> int:
        if key is None:
            return 0
        if key in self._fill_index:
            return self._fill_index[key]
        index = len(self.fills)
        self.fills.append(_fill_from_key(key))
        self._fill_keys.append(key)
        self._fill_xml.append(self._fill_xml_for(key))
        self._fill_index[key] = index
        return index

    def _intern_border(self, key) -> int:
        if key is None:
            return 0
        if key in self._border_index:
            return self._border_index[key]
        index = len(self.borders)
        self.borders.append(_border_from_key(key))
        self._border_keys.append(key)
        self._border_xml.append(self._border_xml_for(key))
        self._border_index[key] = index
        return index

    def _intern_numfmt(self, code: str) -> int:
        if code in self._numfmt_by_code:
            return self._numfmt_by_code[code]
        num_id = self._next_numfmt_id
        self._next_numfmt_id += 1
        self.num_formats[num_id] = code
        self.custom_num_formats.add(num_id)
        self._numfmt_by_code[code] = num_id
        return num_id

    def add(self, style: Style) -> int:
        """Return the XF index for ``style``, appending components as needed."""
        if not isinstance(style, Style):
            raise TypeError("style must be a Style instance")
        if style.pattern is None and isinstance(style.fill, PatternFill) and _pattern_fill_key(style.fill) is None:
            style = replace(style, fill=None)
        if isinstance(style.pattern, PatternFill) and _pattern_fill_key(style.pattern) is None:
            style = replace(style, pattern=None)
        font_key, fill_key, border_key, numfmt, align_key, prot_key = split_style(style)
        font_id = self._intern_font(font_key)
        fill_id = self._intern_fill(fill_key)
        border_id = self._intern_border(border_key)
        numfmt_id = 0 if numfmt is None else self._intern_numfmt(numfmt)
        if align_key is not None and all(value is None for value in align_key):
            align_key = None
        if prot_key is not None and all(value is None for value in prot_key):
            prot_key = None
        key = (font_id, fill_id, border_id, numfmt_id, align_key, prot_key)
        existing = self._xf_index.get(key)
        if existing is not None:
            return existing
        index = len(self.xfs)
        alignment = _alignment_from_key(align_key)
        protection = _protection_from_key(prot_key)
        self.xfs.append(
            _Xf(font_id, fill_id, border_id, numfmt_id, alignment, protection)
        )
        self.styles.append(style)
        self._xf_xml.append(self._xf_xml_for(font_id, fill_id, border_id, numfmt_id, align_key, prot_key))
        self._xf_index[key] = index
        return index

    # -- XML serialisation -------------------------------------------
    @staticmethod
    def _color_attrs(key) -> str:
        kind = key[0]
        if kind == "rgb":
            return f'rgb="{key[1]}"'
        if kind == "indexed":
            return f'indexed="{key[1]}"'
        if kind == "auto":
            return 'auto="1"'
        _, theme, tint = key
        if tint is None:
            return f'theme="{theme}"'
        return f'theme="{theme}" tint="{_format_number(tint)}"'

    def _color_xml(self, key) -> str:
        if key is None:
            return '<color theme="1"/>'
        return f'<color {self._color_attrs(key)}/>'

    def _font_xml_for(self, key) -> str:
        name = key.name
        size = key.size
        parts = []
        if size is not None:
            parts.append(f'<sz val="{_format_number(size)}"/>')
        parts.append(self._color_xml(key.color))
        label = escape(name, {'"': "&quot;"}) if name is not None else "Calibri"
        parts.append(f'<name val="{label}"/>')
        if key.family is not None:
            parts.append(f'<family val="{key.family}"/>')
        else:
            parts.append('<family val="2"/>')
        if key.scheme is not None:
            parts.append(f'<scheme val="{key.scheme}"/>')
        elif name is None:
            parts.append('<scheme val="minor"/>')
        if key.charset is not None:
            parts.append(f'<charset val="{key.charset}"/>')
        if key.bold:
            parts.append("<b/>")
        if key.italic:
            parts.append("<i/>")
        if key.underline is not None:
            if key.underline == "single":
                parts.append("<u/>")
            else:
                parts.append(f'<u val="{key.underline}"/>')
        if key.strike:
            parts.append("<strike/>")
        if key.vert_align is not None:
            parts.append(f'<vertAlign val="{key.vert_align}"/>')
        if key.outline:
            parts.append("<outline/>")
        if key.shadow:
            parts.append("<shadow/>")
        if key.condense:
            parts.append("<condense/>")
        if key.extend:
            parts.append("<extend/>")
        return f"<font>{''.join(parts)}</font>"

    def _fill_xml_for(self, key) -> str:
        if key[0] == "gradient":
            _, fill_type, degree, left, right, top, bottom, stops = key
            attrs = [f'type="{fill_type}"']
            if degree is not None:
                attrs.append(f'degree="{_format_number(degree)}"')
            for name, value in (("left", left), ("right", right), ("top", top), ("bottom", bottom)):
                if value is not None:
                    attrs.append(f'{name}="{_format_number(value)}"')
            stop_xml = "".join(
                f'<stop position="{_format_number(position)}"><color {self._color_attrs(color)}/></stop>'
                for color, position in stops
            )
            return f'<fill><gradientFill {" ".join(attrs)}>{stop_xml}</gradientFill></fill>'
        _, pattern, fg_key, bg_key = key
        pattern_type = pattern if pattern is not None else "none"
        if pattern_type == "solid":
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            bg_xml = (
                f'<bgColor {self._color_attrs(bg_key)}/>'
                if bg_key is not None else '<bgColor indexed="64"/>'
            )
        else:
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            bg_xml = f'<bgColor {self._color_attrs(bg_key)}/>' if bg_key is not None else ""
        return f'<fill><patternFill patternType="{pattern_type}">{fg_xml}{bg_xml}</patternFill></fill>'

    @staticmethod
    def _border_side_xml(tag: str, side_key, color_attrs) -> str:
        style, color = side_key
        if style is None and color is None:
            return f"<{tag}/>"
        inner = f"<color {color_attrs(color)}/>" if color is not None else ""
        if style is None:
            return f"<{tag}>{inner}</{tag}>"
        return f'<{tag} style="{style}">{inner}</{tag}>'

    def _border_xml_for(self, key) -> str:
        attrs = ""
        if key.diagonal_up:
            attrs += ' diagonalUp="1"'
        if key.diagonal_down:
            attrs += ' diagonalDown="1"'
        if key.outline is False:
            attrs += ' outline="0"'
        # Excel reads <start>/<end> as the left/right edges (in a
        # left-to-right sheet), so an empty <start/> would hide <left>:
        # write them only when set, like <vertical>/<horizontal>.
        parts = [
            self._border_side_xml(tag, getattr(key, tag), self._color_attrs)
            for tag in ("start", "end")
            if getattr(key, tag) != (None, None)
        ]
        parts += [
            self._border_side_xml("left", key.left, self._color_attrs),
            self._border_side_xml("right", key.right, self._color_attrs),
            self._border_side_xml("top", key.top, self._color_attrs),
            self._border_side_xml("bottom", key.bottom, self._color_attrs),
            self._border_side_xml("diagonal", key.diagonal, self._color_attrs),
        ]
        if key.vertical != (None, None):
            parts.append(self._border_side_xml("vertical", key.vertical, self._color_attrs))
        if key.horizontal != (None, None):
            parts.append(self._border_side_xml("horizontal", key.horizontal, self._color_attrs))
        return f"<border{attrs}>{''.join(parts)}</border>"

    def _xf_xml_for(self, font_id, fill_id, border_id, numfmt_id, align_key, prot_key) -> str:
        attrs = [
            f'numFmtId="{numfmt_id}"', f'fontId="{font_id}"',
            f'fillId="{fill_id}"', f'borderId="{border_id}"', 'xfId="0"',
        ]
        if numfmt_id:
            attrs.append('applyNumberFormat="1"')
        if font_id:
            attrs.append('applyFont="1"')
        if fill_id:
            attrs.append('applyFill="1"')
        if border_id:
            attrs.append('applyBorder="1"')
        align = _alignment_from_key(align_key)
        align_xml = _alignment_xml(align) if align is not None else ""
        if align_xml:
            attrs.append('applyAlignment="1"')
        prot = _protection_from_key(prot_key)
        prot_xml = _protection_xml(prot) if prot is not None else ""
        if prot_xml:
            attrs.append('applyProtection="1"')
        return f'<xf {" ".join(attrs)}>{align_xml}{prot_xml}</xf>'

    _STYLE_ORDER = (
        "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
        "cellStyles", "dxfs", "tableStyles", "colors", "extLst",
    )

    def _replace_container(self, root: ET.Element, tag: str, fragments: list[str], count: int | None = None) -> None:
        nodes = [child for child in root if _local(child.tag) == tag]
        container = nodes[0] if nodes else None
        if not fragments:
            for node in nodes:
                root.remove(node)
            return
        if container is None:
            container = ET.Element(_q(tag))
            position = self._STYLE_ORDER.index(tag)
            successors = set(self._STYLE_ORDER[position + 1:])
            insertion = len(root)
            for index, child in enumerate(list(root)):
                if _local(child.tag) in successors:
                    insertion = index
                    break
            root.insert(insertion, container)
        else:
            for node in nodes[1:]:
                root.remove(node)
            container.clear()
        if count is not None:
            container.set("count", str(count))
        for fragment in fragments:
            try:
                container.append(ET.fromstring(fragment))
            except ET.ParseError:
                continue

    def to_bytes(self) -> bytes:
        """Serialise the table, preserving unrelated style children."""
        if self._root is None:
            root = ET.Element(_q("styleSheet"))
            ET.SubElement(root, _q("fonts"), {"count": str(len(self._font_xml))})
            ET.SubElement(root, _q("fills"), {"count": str(len(self._fill_xml))})
            ET.SubElement(root, _q("borders"), {"count": str(len(self._border_xml))})
            cell_style_xfs = ET.SubElement(root, _q("cellStyleXfs"), {"count": "1"})
            ET.SubElement(cell_style_xfs, _q("xf"), {
                "numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0",
            })
            ET.SubElement(root, _q("cellXfs"), {"count": str(len(self._xf_xml))})
            cell_styles = ET.SubElement(root, _q("cellStyles"), {"count": "1"})
            ET.SubElement(cell_styles, _q("cellStyle"), {
                "name": "Normal", "xfId": "0", "builtinId": "0",
            })
            ET.SubElement(root, _q("dxfs"), {"count": "0"})
            ET.SubElement(root, _q("tableStyles"), {
                "count": "0", "defaultTableStyle": "TableStyleMedium9",
                "defaultPivotStyle": "PivotStyleLight16",
            })
            self._root = root

        root = self._root
        custom = []
        for num_id, code in sorted(self.num_formats.items()):
            if num_id in self.custom_num_formats:
                quoted = escape(code, {'"': "&quot;"})
                custom.append(f'<numFmt numFmtId="{num_id}" formatCode="{quoted}"/>')
        self._replace_container(root, "numFmts", custom, len(custom) if custom else None)
        self._replace_container(root, "fonts", self._font_xml, len(self._font_xml))
        self._replace_container(root, "fills", self._fill_xml, len(self._fill_xml))
        self._replace_container(root, "borders", self._border_xml, len(self._border_xml))
        self._replace_container(root, "cellXfs", self._xf_xml, len(self._xf_xml))
        source = getattr(self, "_source", None)
        if source is not None:
            # Keep Excel's prefixes: mc:Ignorable names them.
            from .._metadata import serialize_like
            return serialize_like(root, source)
        ET.register_namespace("", MAIN)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def append_dxfs(self, fragments: list[str]) -> int:
        """Append differential formats and return the index of the first.

        ``fragments`` are ``<dxf>`` elements in the main namespace. Their
        number formats are given this table's ids, so that a format code
        keeps one id in the workbook.
        """
        if self._root is None:
            self.to_bytes()
        root = self._root
        container = next((child for child in root if _local(child.tag) == "dxfs"), None)
        if container is None:
            self._replace_container(root, "dxfs", [f'<dxf xmlns="{MAIN}"/>'])
            container = next(child for child in root if _local(child.tag) == "dxfs")
            container.clear()
        first = len(container)
        for fragment in fragments:
            node = ET.fromstring(fragment)
            for numfmt in node.iter(_q("numFmt")):
                numfmt.set("numFmtId", str(self._intern_numfmt(numfmt.get("formatCode", "General"))))
            container.append(node)
        container.set("count", str(len(container)))
        return first

    def style_at(self, index: int) -> Style:
        if not 0 <= index < len(self.styles):
            return self.styles[0]
        return self.styles[index]

    def clone(self) -> "XlsxStyleTable":
        import copy
        clone = XlsxStyleTable()
        clone._root = copy.deepcopy(self._root)
        clone.fonts = list(self.fonts)
        clone.fills = list(self.fills)
        clone.borders = list(self.borders)
        clone.num_formats = dict(self.num_formats)
        clone.custom_num_formats = set(self.custom_num_formats)
        clone.xfs = list(self.xfs)
        clone.styles = list(self.styles)
        clone._font_keys = list(self._font_keys)
        clone._fill_keys = list(self._fill_keys)
        clone._border_keys = list(self._border_keys)
        clone._font_xml = list(self._font_xml)
        clone._fill_xml = list(self._fill_xml)
        clone._border_xml = list(self._border_xml)
        clone._xf_xml = list(self._xf_xml)
        clone._next_numfmt_id = self._next_numfmt_id
        clone._rebuild_indexes()
        return clone
