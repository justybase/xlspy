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
    """A rule from a BrtBeginCFRule payload with formulas, or None for an
    unknown kind (color scales, data bars and icon sets are read by
    ``read_conditional_formats`` with their sub-records)."""
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
    """``(sqref, rule)`` pairs from the BrtBeginConditionalFormatting,
    BrtBeginCFRule and color scale / data bar / icon set records of a sheet,
    in file order."""
    found = []
    sqref = None
    pending = None  # (sqref, rule payload, sub-records) of a scale rule

    def flush():
        if pending is None:
            return
        rule_sqref, payload, parts = pending
        kind, _template, _dxf, priority = struct.unpack_from("<IIiI", payload, 0)
        stop = True if struct.unpack_from("<H", payload, 28)[0] & _STOP_IF_TRUE else None
        try:
            rule = _scale_rule(rule_sqref, kind, priority, stop, parts, context)
        except (struct.error, KeyError, TypeError, ValueError, NotImplementedError):
            return
        if rule is not None:
            found.append((rule_sqref, rule))

    for record_id, payload in items:
        if record_id in SCALE_RECORDS:
            if pending is not None:
                pending[2].append((record_id, payload))
            continue
        flush()
        pending = None
        if record_id == BRT_BEGIN_CONDITIONAL_FORMATTING:
            sqref = _sqref(payload, 8)
        elif record_id == BRT_BEGIN_CF_RULE and sqref:
            if len(payload) >= 4 and struct.unpack_from("<I", payload)[0] in _SCALE_TEMPLATES:
                pending = (sqref, payload, [])
                continue
            try:
                rule = cf_rule(sqref, payload, context, style_table)
            except (struct.error, TypeError, ValueError, NotImplementedError):
                continue
            if rule is not None:
                found.append((sqref, rule))
    flush()
    return found


def rule_bin(rule, priority: int, dxf_id: int, context, letters) -> bytes:
    """BrtBeginCFRule with the rule's formulas (or its color scale, data bar
    or icon set records), and BrtEndCFRule."""
    from .styles import ColorScaleRule, DataBarRule, IconSetRule

    first_col, first_row, _, _ = parse_cell_range(rule.sqref.split()[0])
    if isinstance(rule, (ColorScaleRule, DataBarRule, IconSetRule)):
        kind = _ICON_SET if isinstance(rule, IconSetRule) else (
            _DATA_BAR if isinstance(rule, DataBarRule) else _COLOR_SCALE)
        # No dxf, no formula, no flags: Excel keeps everything in the sub-records.
        payload = struct.pack("<IIiIIII", kind, _SCALE_TEMPLATES[kind], -1, priority, 0, 0, 0)
        payload += struct.pack("<H3Ii", 0, 0, 0, 0, -1)
        return (build_record(BRT_BEGIN_CF_RULE, payload)
                + _scale_records(rule, context, (first_col, first_row))
                + build_record(BRT_END_CF_RULE))
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


# Color scales, data bars and icon sets ([MS-XLSB] 2.1.7.62): the rule
# record carries no formula, the sub-records hold the thresholds and colors.
BRT_BEGIN_ICON_SET, BRT_END_ICON_SET = 0x1D1, 0x1D2
BRT_BEGIN_DATA_BAR, BRT_END_DATA_BAR = 0x1D3, 0x1D4
BRT_BEGIN_COLOR_SCALE, BRT_END_COLOR_SCALE = 0x1D5, 0x1D6
BRT_CFVO, BRT_COLOR = 0x1D7, 0x234
SCALE_RECORDS = frozenset((
    BRT_BEGIN_ICON_SET, BRT_BEGIN_DATA_BAR, BRT_BEGIN_COLOR_SCALE, BRT_CFVO, BRT_COLOR,
))
_COLOR_SCALE, _DATA_BAR, _ICON_SET = 3, 4, 6
_SCALE_TEMPLATES = {_COLOR_SCALE: 0x02, _DATA_BAR: 0x03, _ICON_SET: 0x04}
_CFVO_CODES = {"num": 1, "min": 2, "max": 3, "percent": 4, "percentile": 5, "formula": 7}
_CFVO_NAMES = {code: name for name, code in _CFVO_CODES.items()}
# KPISets ([MS-XLSB] 2.5.94).
ICON_SET_CODES = {
    "3Arrows": 0, "3ArrowsGray": 1, "3Flags": 2, "3TrafficLights1": 3,
    "3TrafficLights2": 4, "3Signs": 5, "3Symbols": 6, "3Symbols2": 7,
    "4Arrows": 8, "4ArrowsGray": 9, "4RedToBlack": 10, "4Rating": 11,
    "4TrafficLights": 12, "5Arrows": 13, "5ArrowsGray": 14,
    "5Rating": 15, "5Quarters": 16,
}
_ICON_SET_NAMES = {code: name for name, code in ICON_SET_CODES.items()}
# BrtBeginIconSet flags as Excel writes them: fIcon hides the value,
# fReverse reverses the icon order.
_ICON_ONLY, _ICON_REVERSE = 0x02, 0x04
_BAR_SHOW_VALUE = 0x01


def _scale_color(color) -> bytes:
    """A BrtColor payload; Excel writes RGB colors opaque."""
    data = _color(color)
    if data[0] == 0x05:
        data = data[:7] + b"\xff"
    return data


def cfvo_bin(kind: str, value, context, origin, gte: bool = False) -> bytes:
    """A BrtCFVO payload; a ``formula`` threshold is compiled like a rule
    formula, relative to the range's top-left cell."""
    code = _CFVO_CODES[kind]
    formula = b""
    number = 0.0
    if kind == "formula":
        if value is None:
            raise ValueError("a formula threshold needs a formula")
        text = str(value)
        stats = {}
        formula = compile_formula(
            text[1:] if text.startswith("=") else text, context,
            origin=origin, cell_relative=True, stats=stats,
        )
        size = struct.unpack_from("<I", formula)[0] + 2 * stats["references"]
    else:
        size = 0
        if kind not in ("min", "max"):
            if value is None:
                raise ValueError(f"CFVO type {kind!r} needs a value")
            number = float(value)
    flag = int(gte)
    return struct.pack("<Id3I", code, number, flag, flag, size) + formula


def _scale_records(rule, context, origin) -> bytes:
    from .styles import ColorScaleRule, DataBarRule

    def threshold(item, default, gte=False):
        if item is None:
            return build_record(BRT_CFVO, cfvo_bin(default[0], default[1], context, origin, gte))
        return build_record(BRT_CFVO, cfvo_bin(item.type, item.value, context, origin, gte))

    if isinstance(rule, ColorScaleRule):
        out = build_record(BRT_BEGIN_COLOR_SCALE)
        out += threshold(rule.minimum, ("min", None))
        if rule.midpoint is not None:
            out += threshold(rule.midpoint, ("percentile", 50))
        out += threshold(rule.maximum, ("max", None))
        out += b"".join(build_record(BRT_COLOR, _scale_color(color)) for color in rule.colors)
        return out + build_record(BRT_END_COLOR_SCALE)
    if isinstance(rule, DataBarRule):
        # Without lengths Excel draws bars from 10% to 90% of the cell.
        flags = 0 if rule.show_value is False else _BAR_SHOW_VALUE
        out = build_record(BRT_BEGIN_DATA_BAR, bytes((
            10 if rule.min_length is None else rule.min_length,
            90 if rule.max_length is None else rule.max_length, flags,
        )))
        out += threshold(rule.minimum, ("min", None)) + threshold(rule.maximum, ("max", None))
        out += build_record(BRT_COLOR, _scale_color(rule.color))
        return out + build_record(BRT_END_DATA_BAR)
    flags = (_ICON_ONLY if rule.show_value is False else 0) | (_ICON_REVERSE if rule.reverse else 0)
    out = build_record(BRT_BEGIN_ICON_SET, struct.pack("<IH", ICON_SET_CODES[rule.icon_style], flags))
    for value in rule.values:
        out += build_record(BRT_CFVO, cfvo_bin(rule.cfvo_type, value, context, origin, gte=True))
    return out + build_record(BRT_END_ICON_SET)


def _read_cfvo(payload: bytes, context, origin):
    from ._xlsb_formulas import decode_formula, read_parsed
    from .styles import FormatObject

    code, number, _save_gte, _gte, size = struct.unpack_from("<Id3I", payload, 0)
    kind = _CFVO_NAMES[code]
    if kind == "formula":
        _tokens, _extra, end = read_parsed(payload, 24)
        return FormatObject(kind, decode_formula(payload[24:end], context, origin=origin))
    if kind in ("min", "max"):
        return FormatObject(kind)
    return FormatObject(kind, int(number) if number.is_integer() else number)


def _scale_rule(sqref, kind, priority, stop, parts, context):
    """A color scale, data bar or icon set from its sub-records."""
    from ._xlsb_styles import _color as color_at
    from .styles import ColorScaleRule, DataBarRule, IconSetRule

    first_col, first_row, _, _ = parse_cell_range(sqref.split()[0])
    origin = (first_col, first_row)
    head = None
    cfvos, colors = [], []
    for record_id, payload in parts:
        if record_id == BRT_CFVO:
            cfvos.append(_read_cfvo(payload, context, origin))
        elif record_id == BRT_COLOR:
            colors.append(color_at(payload, 0))
        else:
            head = (record_id, payload)
    if head is None:
        return None
    if kind == _COLOR_SCALE and head[0] == BRT_BEGIN_COLOR_SCALE and len(cfvos) in (2, 3):
        if len(cfvos) == 2:
            return ColorScaleRule(sqref, cfvos[0], None, cfvos[1], colors, priority)
        return ColorScaleRule(sqref, *cfvos, colors, priority)
    if kind == _DATA_BAR and head[0] == BRT_BEGIN_DATA_BAR and len(cfvos) == 2 and colors:
        low, high, flags = head[1][:3]
        return DataBarRule(sqref, cfvos[0], cfvos[1], colors[0],
                           show_value=False if not flags & _BAR_SHOW_VALUE else None,
                           min_length=low, max_length=high, priority=priority)
    if kind == _ICON_SET and head[0] == BRT_BEGIN_ICON_SET and cfvos:
        code, flags = struct.unpack_from("<IH", head[1], 0)
        values = [0 if item.value is None else item.value for item in cfvos]
        return IconSetRule(sqref, _ICON_SET_NAMES[code], values, cfvo_type=cfvos[0].type,
                           show_value=False if flags & _ICON_ONLY else None,
                           reverse=True if flags & _ICON_REVERSE else None,
                           stop_if_true=stop, priority=priority)
    return None


BRT_END_CONDITIONAL_FORMATTING = 0x1CE
# Excel 2010 conditional formats live in a worksheet FRT extension
# (BrtFRTBegin, BrtBeginConditionalFormattings ... BrtEndConditionalFormattings,
# BrtFRTEnd); BrtCFRuleExt ties a rule of the base block to its x14 rule.
_FRT_BEGIN, _FRT_END = 0x23, 0x24
_BEGIN_CF14_LIST, _CF_RULE14, _CF_RULE_EXT = 0x46F, 0x418, 0x47A


def _letter(index: int) -> str:
    from .refs import column_label

    return column_label(index)


class _Letters:
    __getitem__ = staticmethod(_letter)


def conditional_formats_bin(formats, context, dxf_id) -> bytes:
    """The BrtBeginConditionalFormatting blocks for ``formats``, an iterable
    of ``(sqref, pivot, rules)``; ``dxf_id`` gives a rule's style index.
    Rules without a priority come after the others, in order."""
    formats = [(sqref, pivot, list(rules)) for sqref, pivot, rules in formats]
    taken = set()
    for _sqref, _pivot, rules in formats:
        for rule in rules:
            if rule.priority is None:
                continue
            if rule.priority in taken:
                raise ValueError(f"duplicate conditional formatting priority {rule.priority}")
            taken.add(rule.priority)
    following = max(taken, default=0) + 1
    out = []
    for sqref, pivot, rules in formats:
        if not rules:
            continue
        ranges = [parse_cell_range(part) for part in sqref.split()]
        head = struct.pack("<III", len(rules), 1 if pivot else 0, len(ranges))
        head += b"".join(struct.pack("<IIII", r1, r2, c1, c2) for c1, r1, c2, r2 in ranges)
        out.append(build_record(BRT_BEGIN_CONDITIONAL_FORMATTING, head))
        for rule in rules:
            priority = rule.priority
            if priority is None:
                priority = following
                following += 1
            out.append(rule_bin(rule, priority, dxf_id(getattr(rule, "dxf", None)), context, _Letters()))
        out.append(build_record(BRT_END_CONDITIONAL_FORMATTING))
    return b"".join(out)


def without_conditional_formats(data: bytes) -> bytes:
    """A worksheet part without its conditional formats (with their Excel
    2010 extension). Raises ``NotImplementedError`` when the extension holds
    rules the model cannot read, which a rewrite would lose."""
    from .biff_utils import iter_records

    records = list(iter_records(data))
    kept, linked, extended = [], set(), []
    depth = 0
    index = 0
    while index < len(records):
        record = records[index]
        kind = record.record_id
        payload = data[record.data_start:record.data_end]
        if kind == BRT_BEGIN_CONDITIONAL_FORMATTING:
            depth += 1
        if depth:
            if kind == _CF_RULE_EXT and len(payload) >= 20:
                linked.add(payload[4:20])
            if kind == BRT_END_CONDITIONAL_FORMATTING:
                depth -= 1
            index += 1
            continue
        if (kind == _FRT_BEGIN and index + 1 < len(records)
                and records[index + 1].record_id == _BEGIN_CF14_LIST):
            end = index + 1
            while end < len(records) and records[end].record_id != _FRT_END:
                if records[end].record_id == _CF_RULE14:
                    extended.append(data[records[end].data_start:records[end].data_end])
                end += 1
            index = end + 1
            continue
        kept.append(data[record.header_start:record.data_end])
        index += 1
    for payload in extended:
        if not any(guid in payload for guid in linked):
            raise NotImplementedError(
                "this sheet has Excel 2010 conditional formats the model cannot read; "
                "editing its conditional formats would lose them"
            )
    return b"".join(kept)
