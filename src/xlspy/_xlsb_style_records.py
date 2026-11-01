"""BIFF12 style records shared by the XLSB writer and the loaded-workbook saver.

Each function encodes one ``split_style`` component key ([MS-XLSB] 2.4.311
BrtBorder, 2.4.674 BrtFill, 2.4.683 BrtFont, 2.4.863 BrtXF, 2.4.334 BrtColor);
``_xlsb_styles`` decodes the same records back into keys.
"""

from __future__ import annotations

import struct
from typing import Any

from .biff_utils import build_record

_BRT_FONT = 0x002B
_BRT_FMT = 0x002C
_BRT_FILL = 0x002D
_BRT_BORDER = 0x002E
_BRT_XF = 0x002F

# Property encodings for BrtFont, BrtFill, BrtBorder and BrtXF; the values come
# from the tables in [MS-XLSB] sections 2.4.311 (BrtBorder), 2.4.674 (BrtFill),
# 2.4.683 (BrtFont) and 2.4.863 (BrtXF).
_UNDERLINE_BYTES = {
    None: 0x00, "none": 0x00, "single": 0x01, "double": 0x02,
    "singleAccounting": 0x21, "doubleAccounting": 0x22,
}
_FONT_BASELINE = {None: 0, "baseline": 0, "superscript": 1, "subscript": 2}
_FONT_SCHEME = {None: 0, "major": 1, "minor": 2}
_PATTERN_BYTES = {
    None: 0, "none": 0, "solid": 1, "mediumGray": 2, "darkGray": 3,
    "lightGray": 4, "darkHorizontal": 5, "darkVertical": 6, "darkDown": 7,
    "darkUp": 8, "darkGrid": 9, "darkTrellis": 10, "lightHorizontal": 11,
    "lightVertical": 12, "lightDown": 13, "lightUp": 14, "lightGrid": 15,
    "lightTrellis": 16, "gray125": 17, "gray0625": 18,
}
_BORDER_BYTES = {
    None: 0x00, "none": 0x00, "thin": 0x01, "medium": 0x02, "dashed": 0x03,
    "dotted": 0x04, "thick": 0x05, "double": 0x06, "hair": 0x07,
    "mediumDashed": 0x08, "dashDot": 0x09, "mediumDashDot": 0x0A,
    "dashDotDot": 0x0B, "mediumDashDotDot": 0x0C, "slantDashDot": 0x0D,
}
_HORIZONTAL_BYTES = {
    None: 0, "general": 0, "left": 1, "center": 2, "right": 3, "fill": 4,
    "justify": 5, "centerContinuous": 6, "distributed": 7,
}
_VERTICAL_BYTES = {
    None: 2, "top": 0, "center": 1, "bottom": 2, "justify": 3, "distributed": 4,
}


def numfmt_record(ifmt: int, fmt_string: str) -> bytes:
    encoded = fmt_string.encode('utf-16-le')
    payload = struct.pack('<HI', ifmt, len(encoded) // 2) + encoded
    return build_record(_BRT_FMT, payload)

def color_payload(color_key) -> bytes:
    """An 8-byte BrtColor ([MS-XLSB] 2.4.334) for a ``split_style`` color."""
    if color_key is None:
        return bytes(8)  # automatic
    kind = color_key[0]
    if kind == "rgb":
        alpha, red, green, blue = bytes.fromhex(color_key[1])
        return bytes((0x05, 0, 0, 0, red, green, blue, alpha))
    if kind == "indexed":
        return bytes((0x03, color_key[1], 0, 0, 0, 0, 0, 0xFF))
    if kind == "theme":
        tint = color_key[2]
        scaled = 0 if tint is None else max(-32768, min(32767, round(tint * 32767)))
        return bytes((0x07, color_key[1])) + struct.pack('<h', scaled) + bytes((0, 0, 0, 0xFF))
    return bytes(8)  # ("auto",)

def font_record(font_key: Any) -> bytes:
    """Build a BrtFont record ([MS-XLSB] 2.4.683) from a font key."""
    flags = 0
    for bit, value in (
        (0x01, font_key.bold), (0x02, font_key.italic), (0x08, font_key.strike),
        (0x10, font_key.outline), (0x20, font_key.shadow),
        (0x40, font_key.condense), (0x80, font_key.extend),
    ):
        if value:
            flags |= bit
    height = 220 if font_key.size is None else max(0x14, min(0x1FFF, round(font_key.size * 20)))
    payload = bytearray()
    payload.extend(struct.pack(
        '<HHHH', height, flags, 700 if font_key.bold else 400,
        _FONT_BASELINE[font_key.vert_align],
    ))
    payload.extend((
        _UNDERLINE_BYTES[font_key.underline],
        font_key.family if font_key.family is not None else 2,
        font_key.charset if font_key.charset is not None else 0,
        0,
    ))
    payload.extend(color_payload(font_key.color))
    scheme = _FONT_SCHEME[font_key.scheme]
    if font_key.scheme is None and font_key.name is None:
        # The default Calibri is the minor (body) theme font; a font that
        # names its typeface must not claim a theme scheme or Excel would
        # substitute the theme font for the requested one.
        scheme = 2
    payload.append(scheme)
    name = (font_key.name or 'Calibri').encode('utf-16-le')
    payload.extend(struct.pack('<I', len(name) // 2))
    payload.extend(name)
    return build_record(_BRT_FONT, bytes(payload))

def fill_record(fill_key: Any) -> bytes:
    """Build a BrtFill record ([MS-XLSB] 2.4.674) from a fill key."""
    if fill_key[0] == "gradient":
        return gradient_fill_record(fill_key)
    _, fill_type, start_color, end_color = fill_key
    payload = bytearray()
    payload.extend(struct.pack('<I', _PATTERN_BYTES[fill_type]))
    payload.extend(color_payload(start_color))
    payload.extend(color_payload(end_color))
    payload.extend(bytes(48))  # unused gradient fields
    return build_record(_BRT_FILL, bytes(payload))

def gradient_fill_record(fill_key: Any) -> bytes:
    _, fill_type, degree, left, right, top, bottom, stops = fill_key
    linear = fill_type == "linear"
    payload = bytearray()
    payload.extend(struct.pack('<I', 0x28))
    payload.extend(color_payload(None))
    payload.extend(color_payload(None))
    payload.extend(struct.pack('<I', 0 if linear else 1))
    # BIFF12 measures the gradient angle counter-clockwise while xlspy (like
    # OOXML) measures it clockwise.
    angle = (360 - degree) % 360 if linear and degree is not None else 0.0
    payload.extend(struct.pack('<d', angle))
    if linear:
        payload.extend(struct.pack('<dddd', 0.0, 0.0, 0.0, 0.0))
    else:
        # OOXML omits unset fillToRect values, i.e. the inner rectangle
        # spans the whole cell.
        payload.extend(struct.pack(
            '<dddd',
            0.0 if left is None else float(left),
            1.0 if right is None else float(right),
            0.0 if top is None else float(top),
            1.0 if bottom is None else float(bottom),
        ))
    payload.extend(struct.pack('<I', len(stops)))
    for color_key, position in stops:
        payload.extend(color_payload(color_key))
        payload.extend(struct.pack('<d', float(position)))
    return build_record(_BRT_FILL, bytes(payload))

def border_record(border_key: Any) -> bytes:
    """Build a BrtBorder record ([MS-XLSB] 2.4.311) from a border key."""
    diagonal = 0
    if border_key.diagonal_down:
        diagonal |= 0x01
    if border_key.diagonal_up:
        diagonal |= 0x02
    payload = bytearray((diagonal,))
    for style, color in (
        border_key.top, border_key.bottom, border_key.left,
        border_key.right, border_key.diagonal,
    ):
        payload.append(_BORDER_BYTES[style])
        payload.append(0)  # reserved
        payload.extend(color_payload(color))
    return build_record(_BRT_BORDER, bytes(payload))

def xf_record(
    ifmt: int = 0, *, flags: int = 0, font_id: int = 0, fill_id: int = 0,
    border_id: int = 0, alignment=None, protection=None, apply_flags: bool = False,
) -> bytes:
    """Build a BrtXF record ([MS-XLSB] 2.4.863).

    ``alignment`` is a ``split_style`` alignment key and ``protection`` a
    protection key.  ``apply_flags`` derives the xfGrbitAtr bits from the
    referenced components; the built-in XFs pass ``flags`` explicitly to
    keep byte parity with Excel's own file.
    """
    rotation = 0
    indent = 0
    alignment_byte = 0x10  # Excel's default: bottom vertical alignment
    flags_byte = 0x10  # Excel's default: locked
    if alignment is not None:
        (
            horizontal, vertical, wrap, shrink, indent_value,
            text_rotation, _relative_indent, justify_last_line,
            reading_order,
        ) = alignment
        rotation = 0 if text_rotation is None else text_rotation
        indent = 0 if indent_value is None else indent_value
        alignment_byte = (
            _HORIZONTAL_BYTES[horizontal]
            | (_VERTICAL_BYTES[vertical] << 3)
            | (0x40 if wrap else 0)
            | (0x80 if justify_last_line else 0)
        )
        flags_byte = (
            0x10
            | (0x01 if shrink else 0)
            | ((reading_order or 0) << 2)
        )
    if protection is not None:
        locked, hidden = protection
        flags_byte = (
            (flags_byte & 0xCF)
            | (0x10 if locked is not False else 0x00)
            | (0x20 if hidden else 0x00)
        )
    if apply_flags:
        flags = (
            (0x01 if ifmt else 0)
            | (0x02 if font_id else 0)
            | (0x04 if alignment is not None else 0)
            | (0x08 if border_id else 0)
            | (0x10 if fill_id else 0)
            | (0x20 if protection is not None else 0)
        )
    payload = bytearray(2)  # ixfeParent: cell style XF 0
    payload.extend(struct.pack('<HHHH', ifmt, font_id, fill_id, border_id))
    payload.append(rotation)
    payload.append(indent)
    payload.append(alignment_byte)
    payload.append(flags_byte)
    payload.extend(struct.pack('<H', flags))
    return build_record(_BRT_XF, bytes(payload))

