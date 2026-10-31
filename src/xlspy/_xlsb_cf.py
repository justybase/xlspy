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
_BORDER_KINDS = {kind: name for name, kind in _BORDERS.items()}
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


BRT_BEGIN_CONDITIONAL_FORMATTING = 0x1CD

_PATTERN_NAMES = {value: name for name, value in _PATTERNS.items()}
_BORDER_NAMES = {value: name for name, value in _BORDER_CODES.items()}
_HORIZONTAL_NAMES = {value: name for name, value in _HORIZONTAL_CODES.items()}
_UNDERLINE_NAMES = {value: name for name, value in _UNDERLINE_CODES.items()}
_CELL_IS_NAMES = {value: name for name, value in _CELL_IS_OPERATORS.items()}
_TEXT_NAMES = {value: name for name, value in _TEXT_OPERATORS.items()}
_TIME_PERIOD_NAMES = {template: name for name, (template, _param) in _TIME_PERIODS.items()}
_BLANK_NAMES = {value: name for name, value in _BLANK_TEMPLATES.items()}
_ABOVE_AVERAGE = {0x19: (False, False), 0x1A: (True, False), 0x1D: (False, True), 0x1E: (True, True)}


def _decoded_text(data: bytes) -> str:
    length = struct.unpack_from("<H", data, 0)[0]
    return data[2:2 + 2 * length].decode("utf-16-le")


def dxf_from_payload(payload: bytes) -> DifferentialStyle:
    """A ``DifferentialStyle`` from a BrtDXF payload (the XFProp list)."""
    from ._xlsb_styles import _color as color_at
    from .styles import Alignment, Border, Font, PatternFill, Protection, Side

    count = struct.unpack_from("<H", payload, 4)[0]
    offset = 6
    font, fill, sides, alignment, protection = {}, {}, {}, {}, {}
    number_format = None
    for _ in range(count):
        kind, size = struct.unpack_from("<HH", payload, offset)
        data = payload[offset + 4:offset + size]
        offset += size
        if kind == _FILL_PATTERN:
            fill["fill_type"] = _PATTERN_NAMES.get(data[0], "solid")
        elif kind == _FOREGROUND:
            fill["start_color"] = color_at(data, 0)
        elif kind == _BACKGROUND:
            fill["end_color"] = color_at(data, 0)
        elif kind == _FONT_COLOR:
            font["color"] = color_at(data, 0)
        elif kind in _BORDER_KINDS:
            style = _BORDER_NAMES.get(struct.unpack_from("<H", data, 8)[0])
            if style and style != "none":
                color = color_at(data, 0)
                sides[_BORDER_KINDS[kind]] = Side(
                    style=style, color=None if color is not None and color.auto else color,
                )
        elif kind == _HORIZONTAL:
            alignment["horizontal"] = _HORIZONTAL_NAMES.get(data[0])
        elif kind == _WRAP:
            alignment["wrap_text"] = bool(data[0])
        elif kind == _FONT_NAME:
            font["name"] = _decoded_text(data)
        elif kind == _BOLD:
            font["bold"] = struct.unpack_from("<H", data, 0)[0] >= 700
        elif kind == _UNDERLINE:
            code = struct.unpack_from("<H", data, 0)[0]
            font["underline"] = _UNDERLINE_NAMES.get(code, "none") if code else "none"
        elif kind == _ITALIC:
            font["italic"] = bool(data[0])
        elif kind == _STRIKE:
            font["strike"] = bool(data[0])
        elif kind == _SIZE:
            font["size"] = struct.unpack_from("<I", data, 0)[0] / 20
        elif kind == _NUMBER_FORMAT:
            number_format = _decoded_text(data)
        elif kind == _LOCKED:
            protection["locked"] = bool(data[0])
        elif kind == _HIDDEN:
            protection["hidden"] = bool(data[0])
    return DifferentialStyle(
        font=Font(**font) if font else None,
        fill=PatternFill(**fill) if fill else None,
        border=Border(**sides) if sides else None,
        number_format=number_format,
        alignment=Alignment(**alignment) if alignment else None,
        protection=Protection(**protection) if protection else None,
    )


def _sqref(payload: bytes, offset: int) -> str:
    from ._xlsb_validation import _sqref as format_sqref

    count = struct.unpack_from("<I", payload, offset)[0]
    ranges = [struct.unpack_from("<IIII", payload, offset + 4 + 16 * index) for index in range(count)]
    return format_sqref(ranges)


def cf_rule(sqref: str, payload: bytes, context, style_table):
    """A rule from a BrtBeginCFRule payload, or None for the kinds the model
    reads from XLSX only (color scales, data bars, icon sets)."""
    from ._xlsb_formulas import decode_formula, read_parsed
    from ._xlsb_metadata import read_wide

    kind, template, dxf_id, priority, param = struct.unpack_from("<IIiII", payload, 0)
    flags = struct.unpack_from("<H", payload, 28)[0]
    sizes = struct.unpack_from("<III", payload, 30)
    text, offset = read_wide(payload, 42, nullable=True)
    first_col, first_row, _, _ = parse_cell_range(sqref.split()[0])
    formulas = []
    for size in sizes:
        if not size:
            continue
        _tokens, _extra, end = read_parsed(payload, offset)
        formulas.append(decode_formula(payload[offset:end], context, origin=(first_col, first_row)))
        offset = end
    dxf = style_table.dxf_at(dxf_id) if dxf_id >= 0 else None
    stop = True if flags & _STOP_IF_TRUE else None
    if kind == _CELL_IS:
        operator = _CELL_IS_NAMES.get(param)
        return None if operator is None else CellIsRule(sqref, operator, formulas, dxf, stop, priority)
    if kind == _TOP10:
        return TopBottomRule(sqref, rank=param, percent=bool(flags & _PERCENT),
                             bottom=bool(flags & _BOTTOM), dxf=dxf, stop_if_true=stop, priority=priority)
    if kind != _EXPRESSION:
        return None
    if template == _FORMULA_TEMPLATE:
        return FormulaRule(sqref, formulas, dxf, stop, priority)
    if template in _ABOVE_AVERAGE:
        below, equal = _ABOVE_AVERAGE[template]
        return AboveAverageRule(sqref, below=below, equal_average=equal, std_dev=param or None,
                                dxf=dxf, stop_if_true=stop, priority=priority)
    if template in (_UNIQUE, _DUPLICATE):
        return DuplicateRule(sqref, unique=template == _UNIQUE, dxf=dxf, stop_if_true=stop,
                             priority=priority)
    if template == _CONTAINS_TEXT and param in _TEXT_NAMES:
        return TextRule(sqref, _TEXT_NAMES[param], text or "", dxf, stop, priority)
    if template in _TIME_PERIOD_NAMES:
        return TimePeriodRule(sqref, _TIME_PERIOD_NAMES[template], dxf, stop, priority)
    if template in _BLANK_NAMES:
        return BlanksRule(sqref, _BLANK_NAMES[template], dxf=dxf, stop_if_true=stop, priority=priority)
    return None


def read_conditional_formats(items, context, style_table):
    """``(sqref, rule)`` pairs from the BrtBeginConditionalFormatting and
    BrtBeginCFRule records of a sheet, in file order."""
    found = []
    sqref = None
    for record_id, payload in items:
        if record_id == BRT_BEGIN_CONDITIONAL_FORMATTING:
            sqref = _sqref(payload, 8)
        elif record_id == BRT_BEGIN_CF_RULE and sqref:
            try:
                rule = cf_rule(sqref, payload, context, style_table)
            except (struct.error, TypeError, ValueError, NotImplementedError):
                continue
            if rule is not None:
                found.append((sqref, rule))
    return found


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
