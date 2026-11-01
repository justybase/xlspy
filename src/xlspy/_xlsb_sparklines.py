"""Sparkline groups as BIFF12 records: the XLSB form of ``x14:sparklineGroups``.

The layout was mapped from files Excel saved (``tools/capture_xlsb.py`` changes one
option at a time). A worksheet FRT block holds the groups:

    BrtFRTBegin(0x0E02) BrtBeginSparklineGroups
        ( BrtBeginSparklineGroup BrtBeginSparklines BrtSparkline* BrtEndSparklines
          BrtEndSparklineGroup )*
    BrtEndSparklineGroups BrtFRTEnd

BrtBeginSparklineGroup (98 bytes): a zero FRT header; two flag bytes; eight BrtColor
values (series, negative, axis, markers, first, last, high, low); the manual maximum,
manual minimum and line weight as doubles; the type (0 line, 1 column, 2 win/loss)
as a uint16 and two zero bytes. BrtSparkline: an FRT header with sqref and formula
flags, the location as an sqref and the data as a PtgArea3d formula.
"""

from __future__ import annotations

import struct

from ._xlsb_formulas import encode_reference
from .biff_utils import build_record
from .refs import parse_cell_range
from .sparklines import SparklineGroup, _argb, _expand

_FRT_BEGIN, _FRT_END = 0x23, 0x24
_FRT_SPARKLINES = 0x0E02  # the product version Excel 2010 writes for this block
_BEGIN_GROUPS, _END_GROUPS = 0x422, 0x423
_BEGIN_GROUP, _END_GROUP = 0x411, 0x412
_BEGIN_SPARKLINES, _END_SPARKLINES = 0x420, 0x421
_SPARKLINE = 0x413

_XLSB_EMPTY_CELL_CODES = {"zero": 0, "gap": 1, "connect": 2}
_XLSB_SPARKLINE_KIND_CODES = {"line": 0, "column": 1, "win_loss": 2}
# Second flag byte: how each end of the vertical axis is scaled. A manual value
# sets no bit and is stored as a double.
_MIN_AXIS = {None: 0x10, "individual": 0x10, "group": 0x40}
_MAX_AXIS = {None: 0x08, "individual": 0x08, "group": 0x20}


def _sparkline_color_record(value: str, name: str) -> bytes:
    """BrtColor: an RGB colour (xColorType 2, valid), no tint, opaque."""
    rgb = bytes.fromhex(_argb(value, name)[2:])
    return b"\x05\xff\x00\x00" + rgb + b"\xff"


def _sparkline_axis_code(value, bits) -> tuple[int, float]:
    if isinstance(value, str) or value is None:
        return bits[value], 0.0
    return 0, float(value)


def _group_payload(group: SparklineGroup) -> bytes:
    first = _XLSB_EMPTY_CELL_CODES[group.empty_cells] << 1
    for flag, bit in ((group.markers, 0x08), (group.high_point, 0x10), (group.low_point, 0x20),
                      (group.first_point, 0x40), (group.last_point, 0x80)):
        if flag:
            first |= bit
    min_bits, min_value = _sparkline_axis_code(group.axis_min, _MIN_AXIS)
    max_bits, max_value = _sparkline_axis_code(group.axis_max, _MAX_AXIS)
    second = min_bits | max_bits
    for flag, bit in ((group.negative_points, 0x01), (group.show_axis, 0x02),
                      (group.show_hidden, 0x04), (group.right_to_left, 0x80)):
        if flag:
            second |= bit
    colors = (
        ("color", group.color), ("negative_color", group.negative_color),
        ("axis_color", group.axis_color), ("marker_color", group.marker_color),
        ("first_color", group.first_color), ("last_color", group.last_color),
        ("high_color", group.high_color), ("low_color", group.low_color),
    )
    return (b"\x00" * 4 + bytes((first, second))
            + b"".join(_sparkline_color_record(value, name) for name, value in colors)
            + struct.pack("<dddH", max_value, min_value, float(group.line_width), _XLSB_SPARKLINE_KIND_CODES[group.kind])
            + b"\x00\x00")


def _unquote(sheet: str) -> str:
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
        return sheet[1:-1].replace("''", "'")
    return sheet


def _sparkline_payload(cell: str, formula: str, context) -> bytes:
    column, row, _, _ = parse_cell_range(cell)
    sheet, _, ref = formula.rpartition("!")
    rgce = encode_reference(ref, _unquote(sheet), context, "R", (0, 0), False)
    # FRT header (fSqref | fFormula), one sqref (fDoAdjust) of one cell range, then
    # one formula: flags, cce, an empty cb and the tokens.
    return (struct.pack("<IIII", 0x06, 1, 0x02, 1) + struct.pack("<IIII", row, row, column, column)
            + struct.pack("<IIII", 1, 0x02, len(rgce), 0) + rgce)


def sparklines_bin(groups: list[SparklineGroup], sheet_name: str, context) -> bytes:
    """The worksheet FRT block for ``groups``, or ``b""`` when there are none.

    ``context`` is the workbook's ``FormulaContext``: the data ranges become
    ExternSheet references, like sheet-qualified formula references.
    """
    if not groups:
        return b""
    parts = [build_record(_FRT_BEGIN, struct.pack("<I", _FRT_SPARKLINES)), build_record(_BEGIN_GROUPS)]
    for group in groups:
        parts.append(build_record(_BEGIN_GROUP, _group_payload(group)))
        parts.append(build_record(_BEGIN_SPARKLINES))
        parts.extend(build_record(_SPARKLINE, _sparkline_payload(cell, formula, context))
                     for cell, formula in _expand(group, sheet_name))
        parts.append(build_record(_END_SPARKLINES))
        parts.append(build_record(_END_GROUP))
    parts += [build_record(_END_GROUPS), build_record(_FRT_END)]
    return b"".join(parts)
