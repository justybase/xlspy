"""Conditional formats with a differential style as XLSB records.

``BrtDXF`` ([MS-XLSB] 2.4.390) holds the style in styles.bin, and
``BrtBeginCFRule`` (2.4.24) the rule. The field values mirror what Excel
writes when it saves the same rules from an XLSX file as XLSB. Style
properties Excel's saved files were not checked against raise ``TypeError``
instead of being guessed.
"""

from __future__ import annotations

import struct

from ._xlsb_formulas import compile_formula
from .biff_utils import build_record
from .refs import parse_cell_range
from .styles import (
    CF_BLANK_RULES, CF_TEXT_RULES, CF_TIME_PERIOD_FORMULAS,
    AboveAverageRule, BlanksRule, CellIsRule, DifferentialStyle, DuplicateRule, FormulaRule,
    GradientFill, TextRule, TimePeriodRule, TopBottomRule, _color_key,
)

BRT_BEGIN_DXFS = 0x1F9
BRT_END_DXFS = 0x1FA
BRT_DXF = 0x1FB
BRT_BEGIN_CF_RULE = 0x1CF
BRT_END_CF_RULE = 0x1D0

# XFPropType ([MS-XLSB] 2.5.165).
_FILL_PATTERN, _FOREGROUND, _BACKGROUND, _FONT_COLOR = 0x00, 0x01, 0x02, 0x05
_BORDERS = {"top": 0x06, "bottom": 0x07, "left": 0x08, "right": 0x09}
_HORIZONTAL, _WRAP = 0x0F, 0x14
_FONT_NAME, _BOLD, _UNDERLINE, _ITALIC, _STRIKE, _SIZE = 0x18, 0x19, 0x1A, 0x1C, 0x1D, 0x24
_NUMBER_FORMAT, _NUMBER_FORMAT_ID = 0x26, 0x29
_LOCKED, _HIDDEN = 0x2B, 0x2C
_DXF_FLAGS = 0x8000  # as in every BrtDXF Excel writes
_AUTO_COLOR = bytes((0x01, 0x40, 0, 0, 0, 0, 0, 0))  # automatic, system foreground

# CFType and CFTemp ([MS-XLSB] 2.5.17 and 2.5.18).
_CELL_IS, _EXPRESSION, _TOP10 = 1, 2, 5
_FORMULA_TEMPLATE, _FILTER_TEMPLATE = 0x01, 0x05
_UNIQUE, _CONTAINS_TEXT, _DUPLICATE = 0x07, 0x08, 0x1B
_BLANK_TEMPLATES = {"blanks": 0x09, "no_blanks": 0x0A, "errors": 0x0B, "no_errors": 0x0C}
_TEXT_OPERATORS = {"contains": 0, "notContains": 1, "beginsWith": 2, "endsWith": 3}
# timePeriod -> (template, iParam)
_TIME_PERIODS = {
    "today": (0x0F, 0), "yesterday": (0x11, 1), "tomorrow": (0x10, 6), "last7Days": (0x12, 2),
    "thisWeek": (0x15, 3), "lastWeek": (0x17, 4), "nextWeek": (0x16, 7),
    "thisMonth": (0x18, 9), "lastMonth": (0x13, 5), "nextMonth": (0x14, 8),
}
_CELL_IS_OPERATORS = {
    "between": 1, "notBetween": 2, "equal": 3, "notEqual": 4,
    "greaterThan": 5, "lessThan": 6, "greaterThanOrEqual": 7, "lessThanOrEqual": 8,
}
# BrtBeginCFRule flags.
_STOP_IF_TRUE, _ABOVE, _PERCENT, _BOTTOM = 0x02, 0x04, 0x08, 0x10

FORMULA_RULES = (
    CellIsRule, FormulaRule, TopBottomRule, AboveAverageRule, TextRule,
    DuplicateRule, TimePeriodRule, BlanksRule,
)

_UNDERLINE_CODES = {"single": 0x01, "double": 0x02, "singleAccounting": 0x21, "doubleAccounting": 0x22}
_HORIZONTAL_CODES = {
    "general": 0, "left": 1, "center": 2, "right": 3, "fill": 4,
    "justify": 5, "centerContinuous": 6, "distributed": 7,
}
_BORDER_CODES = {
    "none": 0x00, "thin": 0x01, "medium": 0x02, "dashed": 0x03, "dotted": 0x04, "thick": 0x05,
    "double": 0x06, "hair": 0x07, "mediumDashed": 0x08, "dashDot": 0x09, "mediumDashDot": 0x0A,
    "dashDotDot": 0x0B, "mediumDashDotDot": 0x0C, "slantDashDot": 0x0D,
}
_PATTERNS = {
    "none": 0, "solid": 1, "mediumGray": 2, "darkGray": 3, "lightGray": 4, "darkHorizontal": 5,
    "darkVertical": 6, "darkDown": 7, "darkUp": 8, "darkGrid": 9, "darkTrellis": 10,
    "lightHorizontal": 11, "lightVertical": 12, "lightDown": 13, "lightUp": 14, "lightGrid": 15,
    "lightTrellis": 16, "gray125": 17, "gray0625": 18,
}


def _unsupported(name: str):
    return TypeError(f"{name} is not supported in XLSB conditional formats; use XlsxWriter for it")


def _prop(kind: int, data: bytes) -> tuple[int, bytes]:
    return kind, struct.pack("<HH", kind, 4 + len(data)) + data


def _text(value: str) -> bytes:
    encoded = value.encode("utf-16-le")
    return struct.pack("<H", len(encoded) // 2) + encoded


def _color(color) -> bytes:
    """An 8-byte XFPropColor; Excel writes index 0xFF for RGB colors."""
    key = _color_key(color)
    if key[0] == "rgb":
        alpha, red, green, blue = bytes.fromhex(key[1])
        return bytes((0x05, 0xFF, 0, 0, red, green, blue, alpha))
    if key[0] == "theme":
        tint = key[2]
        scaled = 0 if tint is None else max(-32768, min(32767, round(tint * 32767)))
        return bytes((0x07, key[1])) + struct.pack("<h", scaled) + bytes((0, 0, 0, 0xFF))
    if key[0] == "indexed":
        return bytes((0x03, key[1], 0, 0, 0, 0, 0, 0xFF))
    return bytes(8)


def _font_props(font) -> list:
    for name in ("vertAlign", "charset", "family", "scheme", "outline", "shadow", "condense", "extend"):
        if getattr(font, name) is not None:
            raise _unsupported(f"Font.{name}")
    props = []
    if font.color is not None:
        props.append(_prop(_FONT_COLOR, _color(font.color)))
    if font.name is not None:
        props.append(_prop(_FONT_NAME, _text(font.name)))
    if font.bold is not None:
        props.append(_prop(_BOLD, struct.pack("<H", 700 if font.bold else 400)))
    if font.underline is not None:
        underline = "single" if font.underline is True else font.underline
        code = 0 if underline in (False, "none") else _UNDERLINE_CODES[underline]
        props.append(_prop(_UNDERLINE, struct.pack("<H", code)))
    if font.italic is not None:
        props.append(_prop(_ITALIC, bytes((int(font.italic),))))
    if font.strike is not None:
        props.append(_prop(_STRIKE, bytes((int(font.strike),))))
    if font.size is not None:
        props.append(_prop(_SIZE, struct.pack("<I", round(font.size * 20))))
    return props


def _fill_props(fill) -> list:
    if isinstance(fill, GradientFill):
        raise _unsupported("GradientFill")
    props = []
    if fill.fill_type is not None:
        props.append(_prop(_FILL_PATTERN, bytes((_PATTERNS[fill.fill_type],))))
    if fill.start_color is not None:
        props.append(_prop(_FOREGROUND, _color(fill.start_color)))
    background = fill.end_color
    if background is None and fill.fill_type == "solid":
        # A conditional format paints a solid fill with its background color,
        # so a fill given only start_color would not show (as in XLSX).
        background = fill.start_color
    if background is not None:
        props.append(_prop(_BACKGROUND, _color(background)))
    return props


def _border_props(border) -> list:
    for name in ("diagonal", "vertical", "horizontal", "start", "end"):
        if getattr(border, name) is not None:
            raise _unsupported(f"Border.{name}")
    # Excel writes all four sides once any is set; an unset side is "none"
    # with the automatic color, as it reads the XLSX writer's empty elements.
    props = []
    for name, kind in _BORDERS.items():
        side = getattr(border, name)
        style = None if side is None else side.style
        color = _AUTO_COLOR if side is None or side.color is None else _color(side.color)
        props.append(_prop(kind, color + struct.pack("<H", _BORDER_CODES[style or "none"])))
    return props


def _alignment_props(alignment) -> list:
    for name in ("vertical", "shrink_to_fit", "indent", "text_rotation", "relative_indent",
                 "justify_last_line", "reading_order"):
        if getattr(alignment, name) is not None:
            raise _unsupported(f"Alignment.{name}")
    props = []
    if alignment.horizontal is not None:
        props.append(_prop(_HORIZONTAL, bytes((_HORIZONTAL_CODES[alignment.horizontal],))))
    if alignment.wrap_text is not None:
        props.append(_prop(_WRAP, bytes((int(alignment.wrap_text),))))
    return props


def dxf_payload(dxf: DifferentialStyle, register_numfmt) -> bytes:
    """The BrtDXF payload; ``register_numfmt`` returns the id of a number format."""
    props = []
    if dxf.font is not None:
        props += _font_props(dxf.font)
    if dxf.fill is not None:
        props += _fill_props(dxf.fill)
    if dxf.border is not None:
        props += _border_props(dxf.border)
    if dxf.alignment is not None:
        props += _alignment_props(dxf.alignment)
    if dxf.protection is not None:
        if dxf.protection.locked is not None:
            props.append(_prop(_LOCKED, bytes((int(dxf.protection.locked),))))
        if dxf.protection.hidden is not None:
            props.append(_prop(_HIDDEN, bytes((int(dxf.protection.hidden),))))
    if dxf.number_format is not None:
        # Excel writes the format id right before the format string.
        props.append((_NUMBER_FORMAT - 0.5, struct.pack("<HHH", _NUMBER_FORMAT_ID, 6,
                                                       register_numfmt(dxf.number_format))))
        props.append(_prop(_NUMBER_FORMAT, _text(dxf.number_format)))
    props.sort(key=lambda item: item[0])
    return struct.pack("<HHH", _DXF_FLAGS, 0, len(props)) + b"".join(data for _, data in props)


def dxfs_bin(payloads: list[bytes]) -> bytes:
    """BrtBeginDXFs, the BrtDXF records and BrtEndDXFs."""
    return (
        build_record(BRT_BEGIN_DXFS, struct.pack("<I", len(payloads)))
        + b"".join(build_record(BRT_DXF, payload) for payload in payloads)
        + build_record(BRT_END_DXFS)
    )


def _rule_fields(rule, corner: str):
    """``(iType, iTemplate, iParam, flags, formulas, text)`` for a rule."""
    flags = _STOP_IF_TRUE if rule.stop_if_true else 0
    if isinstance(rule, CellIsRule):
        operator = _CELL_IS_OPERATORS.get(rule.operator)
        if operator is None:
            raise TypeError(f"CellIsRule operator {rule.operator!r} is XLSX-only; use TextRule for text tests")
        return _CELL_IS, 0, operator, flags, list(rule.formula), None
    if isinstance(rule, FormulaRule):
        return _EXPRESSION, _FORMULA_TEMPLATE, 0, flags, list(rule.formula), None
    if isinstance(rule, TopBottomRule):
        flags |= (_PERCENT if rule.percent else 0) | (_BOTTOM if rule.bottom else 0)
        return _TOP10, _FILTER_TEMPLATE, rule.rank, flags, [], None
    if isinstance(rule, AboveAverageRule):
        template = {(False, False): 0x19, (True, False): 0x1A,
                    (False, True): 0x1D, (True, True): 0x1E}[(rule.below, rule.equal_average)]
        if not rule.below:
            flags |= _ABOVE
        return _EXPRESSION, template, rule.std_dev or 0, flags, [], None
    if isinstance(rule, DuplicateRule):
        return _EXPRESSION, _UNIQUE if rule.unique else _DUPLICATE, 0, flags, [], None
    if isinstance(rule, TextRule):
        quoted = '"' + rule.text.replace('"', '""') + '"'
        formula = CF_TEXT_RULES[rule.operator][2].format(t=quoted, c=corner)
        return _EXPRESSION, _CONTAINS_TEXT, _TEXT_OPERATORS[rule.operator], flags, [formula], rule.text
    if isinstance(rule, TimePeriodRule):
        template, param = _TIME_PERIODS[rule.period]
        formula = CF_TIME_PERIOD_FORMULAS[rule.period].format(c=corner)
        return _EXPRESSION, template, param, flags, [formula], None
    if isinstance(rule, BlanksRule):
        formula = CF_BLANK_RULES[rule.kind][1].format(c=corner)
        return _EXPRESSION, _BLANK_TEMPLATES[rule.kind], 0, flags, [formula], None
    raise TypeError(f"unsupported conditional formatting rule {type(rule).__name__}")


def rule_bin(rule, priority: int, dxf_id: int, context, letters) -> bytes:
    """BrtBeginCFRule with the rule's formulas, and BrtEndCFRule."""
    first_col, first_row, _, _ = parse_cell_range(rule.sqref.split()[0])
    corner = f"{letters[first_col]}{first_row + 1}"
    kind, template, param, flags, formulas, text = _rule_fields(rule, corner)
    # Formulas are relative to the range's top-left cell, like XLSX.
    compiled, sizes = [], []
    for formula in formulas:
        stats = {}
        compiled.append(compile_formula(
            formula, context, origin=(first_col, first_row), cell_relative=True, stats=stats,
        ))
        # cbFmla: Excel writes the token length plus two bytes per reference.
        sizes.append(struct.unpack_from("<I", compiled[-1])[0] + 2 * stats["references"])
    sizes += [0] * (3 - len(sizes))
    payload = struct.pack("<IIiIIII", kind, template, dxf_id, priority, param, 0, 0)
    payload += struct.pack("<H", flags) + struct.pack("<III", *sizes)
    if text is None:
        payload += b"\xff\xff\xff\xff"
    else:
        encoded = text.encode("utf-16-le")
        payload += struct.pack("<I", len(encoded) // 2) + encoded
    payload += b"".join(compiled)
    return build_record(BRT_BEGIN_CF_RULE, payload) + build_record(BRT_END_CF_RULE)
