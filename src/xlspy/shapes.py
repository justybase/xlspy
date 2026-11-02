"""Drawn shapes and text boxes, such as KPI cards, placed over a sheet (XLSX and XLSB).

A ``Shape`` is an ``sp`` object in the sheet's drawing. Its text is static, unless
``link_cell`` names a cell of the same sheet: then Excel fills the text from that
cell and keeps it up to date, which makes a card show a live value. ``text`` is
the value shown until Excel recalculates, so give the formatted value there.
"""

from __future__ import annotations
from typing import Any

from dataclasses import dataclass
import re
import uuid
import xml.sax.saxutils as saxutils

from ._text import check_xml_text
from .refs import column_label, parse_cell_range
from ._text import _attr as _attr

EMU_PER_PIXEL = 9525
_SHAPE_PRESETS = {"rect": "rect", "round": "roundRect", "ellipse": "ellipse", "textbox": "rect"}
_ALIGN = {"left": "l", "center": "ctr", "right": "r"}
_VALIGN = {"top": "t", "center": "ctr", "bottom": "b"}
_COLOR_RE = re.compile(r"#([0-9A-Fa-f]{6})")
_FONT_SIZE_LIMITS = (1, 409)
_NAMESPACE = uuid.UUID("6f1d8c2a-3b7e-4f0a-9d4c-2e8a5b1c7f30")


def _shape_color(value: str, name: str) -> str:
    match = _COLOR_RE.fullmatch(value) if isinstance(value, str) else None
    if match is None:
        raise ValueError(f"{name} must be an RGB hex colour such as '#1F4E79', got {value!r}")
    return match[1].upper()


@dataclass(frozen=True, slots=True)
class Shape:
    """A rectangle, rounded rectangle, ellipse or text box over the sheet.

    ``anchor`` is the top-left cell, ``width`` and ``height`` are in pixels.
    ``fill`` and ``line`` are ``#RRGGBB`` colours, or ``None`` for no fill or
    no outline. ``text`` may hold several lines separated by ``"\\n"``.
    """

    anchor: str
    width: int = 200
    height: int = 80
    kind: str = "rect"
    text: str | None = None
    fill: str | None = None
    line: str | None = None
    font_size: float = 11
    font_color: str = "#000000"
    bold: bool = False
    align: str = "left"
    valign: str = "top"
    link_cell: str | None = None
    name: str | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.anchor, str):
            raise TypeError("shape anchor must be an A1 cell such as 'B2'")
        first_col, first_row, last_col, last_row = parse_cell_range(self.anchor)
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"shape anchor must be a single cell, got {self.anchor!r}")
        for field_name in ("width", "height"):
            value = getattr(self, field_name)
            if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
                raise ValueError(f"shape {field_name} must be a positive integer of pixels")
        if self.kind not in _SHAPE_PRESETS:
            raise ValueError(f"shape kind must be one of {sorted(_SHAPE_PRESETS)}, got {self.kind!r}")
        if self.align not in _ALIGN:
            raise ValueError(f"shape align must be one of {sorted(_ALIGN)}, got {self.align!r}")
        if self.valign not in _VALIGN:
            raise ValueError(f"shape valign must be one of {sorted(_VALIGN)}, got {self.valign!r}")
        if self.text is not None and not isinstance(self.text, str):
            raise TypeError("shape text must be a string")
        if self.text is not None:
            check_xml_text(self.text, "shape text")
        if self.fill is not None:
            _shape_color(self.fill, "fill")
        if self.line is not None:
            _shape_color(self.line, "line")
        _shape_color(self.font_color, "font_color")
        low, high = _FONT_SIZE_LIMITS
        if isinstance(self.font_size, bool) or not isinstance(self.font_size, (int, float)) \
                or not low <= self.font_size <= high:
            raise ValueError(f"shape font_size must be between {low} and {high} points")
        if self.link_cell is not None:
            link_col, link_row, link_end_col, link_end_row = parse_cell_range(self.link_cell)
            if (link_col, link_row) != (link_end_col, link_end_row):
                raise ValueError(f"link_cell must be a single cell, got {self.link_cell!r}")
        if self.name is not None and (not isinstance(self.name, str) or not self.name):
            raise ValueError("shape name must be a non-empty string or None")
        if self.name is not None:
            check_xml_text(self.name, "shape name")


def normalize_shapes(shapes: Any) -> list[Shape]:
    """Validate the shapes of one sheet; raises before any part is written."""
    if not shapes:
        return []
    result = []
    for shape in shapes:
        if not isinstance(shape, Shape):
            raise TypeError("shapes must be Shape instances")
        result.append(shape)
    return result


def _link_reference(cell: str) -> str:
    col, row, _, _ = parse_cell_range(cell)
    return f"${column_label(col)}${row + 1}"


def _run_properties(shape: Shape, tag: str) -> str:
    bold = ' b="1"' if shape.bold else ""
    size = round(shape.font_size * 100)
    return (
        f'<a:{tag} lang="en-US" sz="{size}"{bold} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{_shape_color(shape.font_color, "font_color")}"/></a:solidFill>'
        f"</a:{tag}>"
    )


def _paragraphs(shape: Shape, position: int) -> str:
    """Text paragraphs; with ``link_cell`` the text is a linked field refreshed by Excel."""
    text = shape.text if shape.text is not None else ""
    align = _ALIGN[shape.align]
    lines = text.split("\n") if text else [""]
    paragraphs = []
    for index, line in enumerate(lines):
        escaped = saxutils.escape(line)
        if shape.link_cell is not None and index == 0:
            field_id = "{" + str(uuid.uuid5(_NAMESPACE, f"{position}:{shape.anchor}")).upper() + "}"
            body = (
                f'<a:fld id="{field_id}" type="TxLink">{_run_properties(shape, "rPr")}'
                f'<a:pPr algn="{align}"/><a:t>{escaped}</a:t></a:fld>'
            )
        elif line:
            body = f'<a:r>{_run_properties(shape, "rPr")}<a:t>{escaped}</a:t></a:r>'
        else:
            body = ""
        end = _run_properties(shape, "endParaRPr")
        paragraphs.append(f'<a:p><a:pPr algn="{align}"/>{body}{end}</a:p>')
    return "".join(paragraphs)


def shape_frame_xml(shape: Shape, object_id: int) -> str:
    """The ``oneCellAnchor`` drawing object for one shape, numbered ``object_id``."""
    first_col, first_row, _, _ = parse_cell_range(shape.anchor)
    cx = shape.width * EMU_PER_PIXEL
    cy = shape.height * EMU_PER_PIXEL
    name = shape.name or f"Shape {object_id}"
    textlink = f' textlink="{_link_reference(shape.link_cell)}"' if shape.link_cell else ""
    box = ' txBox="1"' if shape.kind == "textbox" else ""
    fill = (
        f'<a:solidFill><a:srgbClr val="{_shape_color(shape.fill, "fill")}"/></a:solidFill>'
        if shape.fill is not None else "<a:noFill/>"
    )
    outline = (
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{_shape_color(shape.line, "line")}"/></a:solidFill></a:ln>'
        if shape.line is not None else "<a:ln><a:noFill/></a:ln>"
    )
    return (
        "<oneCellAnchor>"
        f"<from><col>{first_col}</col><colOff>0</colOff>"
        f"<row>{first_row}</row><rowOff>0</rowOff></from>"
        f'<ext cx="{cx}" cy="{cy}"/>'
        f'<sp macro=""{textlink}>'
        f'<nvSpPr><cNvPr id="{object_id}" name="{_attr(name)}"/><cNvSpPr{box}/></nvSpPr>'
        '<spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{_SHAPE_PRESETS[shape.kind]}"><a:avLst/></a:prstGeom>'
        f"{fill}{outline}</spPr>"
        f'<txBody><a:bodyPr wrap="square" rtlCol="0" anchor="{_VALIGN[shape.valign]}"/>'
        f"<a:lstStyle/>{_paragraphs(shape, object_id)}</txBody>"
        "</sp><clientData/></oneCellAnchor>"
    )


__all__ = ["Shape", "normalize_shapes", "shape_frame_xml"]
