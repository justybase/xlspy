"""Data validation as XLSB worksheet records ([MS-XLSB] 2.4: BrtBeginDVals, BrtDVal).

The layout and flag values mirror what Excel writes when it saves the same
rules from an XLSX file as XLSB.
"""

from __future__ import annotations

import struct

from ._xlsb_formulas import compile_formula
from ._xlsb_metadata import wide_string
from .biff_utils import build_record
from .data_validation import DataValidation, OPERATORS
from .refs import parse_cell_range

BRT_BEGIN_DVALS = 0x23D
BRT_DVAL = 0x40
BRT_END_DVALS = 0x23E

# valType ("none" is 0), errStyle and typOperator in Excel's numbering; the
# operators follow the order of the OOXML enumeration, as OPERATORS does.
_TYPES = {"whole": 1, "decimal": 2, "list": 3, "date": 4, "time": 5, "text_length": 6, "custom": 7}
_ERROR_STYLES = {"stop": 0, "warning": 1, "information": 2}
_OPERATORS = {name: index for index, name in enumerate(OPERATORS)}

_LITERAL_LIST = 0x80        # fStrLookup: formula1 is a quoted list of values
_ALLOW_BLANK = 0x100
_SUPPRESS_DROPDOWN = 0x200  # fSuppressCombo
_SHOW_INPUT = 1 << 18
_SHOW_ERROR = 1 << 19
_OPERATOR_SHIFT = 20
_NULL_STRING = b"\xff\xff\xff\xff"
_EMPTY_FORMULA = b"\x00" * 8  # cce = 0, cb = 0


def _nullable(text: str | None) -> bytes:
    return _NULL_STRING if not text else wide_string(text)


def _flags(rule: DataValidation) -> int:
    flags = _TYPES[rule.type] | _ERROR_STYLES[rule.error_style] << 4
    if rule.type == "list" and rule.formula1.startswith('"'):
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
        flags |= _OPERATORS[rule.operator] << _OPERATOR_SHIFT
    return flags


def _dval(rule: DataValidation, context) -> bytes:
    ranges = [parse_cell_range(part) for part in rule.sqref.split()]
    first_col, first_row = ranges[0][0], ranges[0][1]
    payload = bytearray(struct.pack("<II", _flags(rule), len(ranges)))
    for c1, r1, c2, r2 in ranges:
        payload += struct.pack("<IIII", r1, r2, c1, c2)
    payload += _nullable(rule.error_title) + _nullable(rule.error)
    payload += _nullable(rule.prompt_title) + _nullable(rule.prompt)
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
