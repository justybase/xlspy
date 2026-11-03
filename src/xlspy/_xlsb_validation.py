"""Data validation as XLSB worksheet records ([MS-XLSB] 2.4: BrtBeginDVals, BrtDVal).

The layout and flag values mirror what Excel writes when it saves the same
rules from an XLSX file as XLSB.
"""

from __future__ import annotations

import struct

from ._xlsb_formulas import compile_formula, decode_formula, read_parsed
from ._xlsb_metadata import read_wide, wide_string
from .biff_utils import build_record
from .data_validation import DataValidation, OPERATORS
from .refs import column_label, parse_cell_range
from ._xlsb_print import _NULL_STRING as _NULL_STRING

BRT_BEGIN_DVALS = 0x23D
BRT_DVAL = 0x40
BRT_END_DVALS = 0x23E

# valType ("none" is 0), errStyle and typOperator in Excel's numbering; the
# operators follow the order of the OOXML enumeration, as OPERATORS does.
_TYPES = {"whole": 1, "decimal": 2, "list": 3, "date": 4, "time": 5, "text_length": 6, "custom": 7}
_ERROR_STYLES = {"stop": 0, "warning": 1, "information": 2}
_XLSB_OPERATOR_CODES = {name: index for index, name in enumerate(OPERATORS)}

_LITERAL_LIST = 0x80        # fStrLookup: formula1 is a quoted list of values
_ALLOW_BLANK = 0x100
_SUPPRESS_DROPDOWN = 0x200  # fSuppressCombo
_SHOW_INPUT = 1 << 18
_SHOW_ERROR = 1 << 19
_OPERATOR_SHIFT = 20
_EMPTY_FORMULA = b"\x00" * 8  # cce = 0, cb = 0


def _nullable_or_empty(text: str | None) -> bytes:
    return _NULL_STRING if not text else wide_string(text)


def _flags(rule: DataValidation) -> int:
    flags = _TYPES[rule.type] | _ERROR_STYLES[rule.error_style] << 4
    if rule.type == "list" and str(rule.formula1).startswith('"'):
        flags |= _LITERAL_LIST
    if rule.allow_blank:
        flags |= _ALLOW_BLANK
    if not rule.show_dropdown:
        flags |= _SUPPRESS_DROPDOWN
    if rule.show_input_message:
        flags |= _SHOW_INPUT
    if rule.show_error_message:
        flags |= _SHOW_ERROR
    if rule.operator is not None:
        flags |= _XLSB_OPERATOR_CODES[rule.operator] << _OPERATOR_SHIFT
    return flags


def _dval(rule: DataValidation, context) -> bytes:
    ranges = [parse_cell_range(part) for part in str(rule.sqref).split()]
    first_col, first_row = ranges[0][0], ranges[0][1]
    payload = bytearray(struct.pack("<II", _flags(rule), len(ranges)))
    for c1, r1, c2, r2 in ranges:
        payload += struct.pack("<IIII", r1, r2, c1, c2)
    payload += _nullable_or_empty(rule.error_title) + _nullable_or_empty(rule.error)
    payload += _nullable_or_empty(rule.prompt_title) + _nullable_or_empty(rule.prompt)
    # A list range is a reference; the other operands are values.
    operand = "R" if rule.type == "list" else "V"
    for text in (rule.formula1, rule.formula2):
        if text is None:
            payload += _EMPTY_FORMULA
        else:
            payload += compile_formula(
                text, context, sheet=None, origin=(first_col, first_row),
                cell_relative=True, operand=operand,
            )
    return build_record(BRT_DVAL, bytes(payload))


def validations_bin(rules, context) -> bytes:
    """BrtBeginDVals, one BrtDVal per rule and BrtEndDVals."""
    if not rules:
        return b""
    # wFlags, xLeft, yTop, unused, idvMac (the rule count).
    begin = struct.pack("<HIIII", 0, 0, 0, 0, len(rules))
    return b"".join((
        build_record(BRT_BEGIN_DVALS, begin),
        *(_dval(rule, context) for rule in rules),
        build_record(BRT_END_DVALS),
    ))


_TYPE_NAMES = {value: name for name, value in _TYPES.items()}
_ERROR_STYLE_NAMES = {value: name for name, value in _ERROR_STYLES.items()}
_OPERATOR_NAMES = list(OPERATORS)
_OPERATOR_TYPES = frozenset({"whole", "decimal", "date", "time", "text_length"})


def _format_sqref(ranges) -> str:
    parts = []
    for r1, r2, c1, c2 in ranges:
        first = f"{column_label(c1)}{r1 + 1}"
        last = f"{column_label(c2)}{r2 + 1}"
        parts.append(first if first == last else f"{first}:{last}")
    return " ".join(parts)


def _formula(payload: bytes, offset: int, context, origin):
    if payload[offset:offset + 8] == _EMPTY_FORMULA:
        return None, offset + 8
    _tokens, _extra, end = read_parsed(payload, offset)
    return decode_formula(payload[offset:end], context, origin=origin), end


def dval_rule(payload: bytes, context) -> DataValidation | None:
    """A ``DataValidation`` from a BrtDVal payload; None for an "any value" rule."""
    flags, count = struct.unpack_from("<II", payload, 0)
    kind = _TYPE_NAMES.get(flags & 0x0F)
    if kind is None:
        return None
    offset = 8
    ranges = []
    for _ in range(count):
        ranges.append(struct.unpack_from("<IIII", payload, offset))
        offset += 16
    if not ranges:
        return None
    error_title, offset = read_wide(payload, offset, nullable=True)
    error, offset = read_wide(payload, offset, nullable=True)
    prompt_title, offset = read_wide(payload, offset, nullable=True)
    prompt, offset = read_wide(payload, offset, nullable=True)
    origin = (ranges[0][2], ranges[0][0])
    formula1, offset = _formula(payload, offset, context, origin)
    formula2, offset = _formula(payload, offset, context, origin)
    operator = (flags >> _OPERATOR_SHIFT) & 0x0F
    return DataValidation(
        type=kind,
        sqref=_format_sqref(ranges),
        formula1=formula1,
        formula2=formula2,
        # "between" is stored as 0, which XLSX leaves out: read it as the default.
        operator=_OPERATOR_NAMES[operator] if kind in _OPERATOR_TYPES and operator else None,
        allow_blank=bool(flags & _ALLOW_BLANK),
        show_dropdown=not flags & _SUPPRESS_DROPDOWN,
        show_input_message=bool(flags & _SHOW_INPUT),
        prompt_title=prompt_title or None,
        prompt=prompt or None,
        show_error_message=bool(flags & _SHOW_ERROR),
        error_title=error_title or None,
        error=error or None,
        error_style=_ERROR_STYLE_NAMES.get((flags >> 4) & 0x07, "stop"),
    )
