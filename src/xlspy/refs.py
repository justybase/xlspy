"""A1 cell and range helpers shared by the style, chart and feature modules.

This module has no xlspy imports so that every model module can depend on it
without creating an import cycle.
"""

from __future__ import annotations

import re
from functools import lru_cache

_RANGE_RE = re.compile(
    r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?"
)


@lru_cache(maxsize=4096)
def _column_index(label: str) -> int:
    result = 0
    for char in label.upper():
        result = result * 26 + ord(char) - 64
    return result - 1


def column_label(index: int) -> str:
    """Zero-based column index to its letters (``0`` -> ``"A"``)."""
    if not 0 <= index < 16384:
        raise ValueError(f"column index out of range: {index}")
    label = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label


def parse_cell_range(reference: str) -> tuple[int, int, int, int]:
    """Parse an A1 cell or rectangular range into zero-based bounds."""
    if not isinstance(reference, str):
        raise TypeError("style range keys must be A1 references such as 'A1:D1'")
    match = _RANGE_RE.fullmatch(reference)
    if match is None:
        raise ValueError(f"invalid A1 range {reference!r}")
    c1, r1, c2, r2 = match.groups()
    if c2 is None:
        c2, r2 = c1, r1
    bounds = (_column_index(c1), int(r1) - 1, _column_index(c2), int(r2) - 1)
    if bounds[2] < bounds[0] or bounds[3] < bounds[1]:
        raise ValueError(f"range end precedes its start in {reference!r}")
    if bounds[2] >= 16384 or bounds[3] >= 1048576:
        raise ValueError(f"range exceeds Excel limits: {reference!r}")
    return bounds


_MAX_COLUMN = 16384  # zero-based column indexes run 0 .. 16383
_MAX_ROW = 1048576

# One pass over a formula. String literals, quoted sheet names and structured
# references match first, so nothing inside them is shifted.
_FORMULA_TOKEN = re.compile(
    r'"(?:[^"]|"")*"'
    r"|'(?:[^']|'')*'"
    r"|\[[^\]]*\]"
    r"|(?<![\w.])(?P<cell>(?P<cabs>\$?)(?P<col>[A-Za-z]{1,3})(?P<rabs>\$?)(?P<row>[1-9][0-9]*))(?![\w.!(])"
    r"|(?<![\w.])(?P<cols>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![\w.])"
    r"|(?<![\w.])(?P<rows>\$?[1-9][0-9]*:\$?[1-9][0-9]*)(?![\w.])"
)


def shift_references(text: str, rows: int, cols: int) -> str:
    """Shift the relative references of formula ``text`` by ``rows`` and ``cols``.

    ``$``-anchored parts stay put, as when a formula is copied or filled. Whole
    column (``A:A``) and whole row (``1:2``) references shift as well. A reference
    pushed off the sheet becomes ``#REF!``. Names such as ``XFE1`` are left alone.
    The result matches openpyxl's ``Translator`` for the same input.
    """

    def replace(match: re.Match) -> str:
        if match["cell"] is not None:
            return _shift_cell(match, rows, cols)
        if match["cols"] is not None:
            return ":".join(_shift_column(part, cols) for part in match["cols"].split(":"))
        if match["rows"] is not None:
            return ":".join(_shift_row(part, rows) for part in match["rows"].split(":"))
        return match[0]

    return _FORMULA_TOKEN.sub(replace, text)


def _shift_cell(match: re.Match, rows: int, cols: int) -> str:
    column = _column_index(match["col"])
    row = int(match["row"]) - 1
    if column >= _MAX_COLUMN or row >= _MAX_ROW:
        return match[0]  # a name such as XFE1, not a cell
    if not match["cabs"]:
        column += cols
    if not match["rabs"]:
        row += rows
    if not (0 <= column < _MAX_COLUMN and 0 <= row < _MAX_ROW):
        return "#REF!"
    return f"{match['cabs']}{column_label(column)}{match['rabs']}{row + 1}"


def _shift_column(part: str, cols: int) -> str:
    if part.startswith("$"):
        return part
    column = _column_index(part) + cols
    return column_label(column) if 0 <= column < _MAX_COLUMN else "#REF!"


def _shift_row(part: str, rows: int) -> str:
    if part.startswith("$"):
        return part
    row = int(part) + rows
    return str(row) if 1 <= row <= _MAX_ROW else "#REF!"


_COLUMN_CACHE: dict[str, int] = {}


def cell_position(ref: str) -> tuple[int, int]:
    """Zero-based ``(column, row)`` of a worksheet cell reference (``r="B2"``).

    Every XLSX reader path accepts the same references: column letters that
    ``parse_cell_range`` accepts, then ASCII digits naming row 1-1048576
    (leading zeros allowed). Anything else raises ValueError. Column letters
    are validated once and cached.
    """
    letters = ref.rstrip("0123456789")
    digits = ref[len(letters):]
    column = _COLUMN_CACHE.get(letters)
    if column is not None and digits:
        row = int(digits)
        if 1 <= row <= 1048576:
            return column, row - 1
    column, row, _, _ = parse_cell_range(f"{letters}{int(digits)}" if digits and int(digits) else ref)
    if len(_COLUMN_CACHE) < 16384:
        _COLUMN_CACHE[letters] = column
    return column, row


def quote_sheet_name(sheet_name: str) -> str:
    """Quote a sheet name for use in a formula or defined name."""
    return "'" + sheet_name.replace("'", "''") + "'"


_SHEET_NAME_FORBIDDEN = frozenset(r':\/?*[]')


def check_sheet_name(sheet_name, existing=()) -> None:
    """Raise unless Excel accepts ``sheet_name`` beside the ``existing`` names.

    Excel refuses to open a workbook with a sheet name that is empty, longer
    than 31 characters, holds ``: \\ / ? * [ ]`` or a control character, starts
    or ends with an apostrophe, is the reserved name ``History``, or repeats
    another sheet's name ignoring case.
    """
    if not isinstance(sheet_name, str):
        raise TypeError("sheet name must be a string")
    if not 1 <= len(sheet_name) <= 31:
        raise ValueError(f"sheet name must be 1 to 31 characters, got {sheet_name!r}")
    bad = sorted(set(sheet_name) & _SHEET_NAME_FORBIDDEN)
    if bad:
        raise ValueError(f"sheet name {sheet_name!r} contains {''.join(bad)!r}, which Excel forbids")
    # XML 1.0 cannot carry these at all, so the XLSX workbook part would not parse.
    control = sorted({ch for ch in sheet_name if ord(ch) < 0x20 or ord(ch) in (0xFFFE, 0xFFFF)})
    if control:
        raise ValueError(
            f"sheet name {sheet_name!r} contains control characters {control!r}"
        )
    if sheet_name[0] == "'" or sheet_name[-1] == "'":
        raise ValueError(f"sheet name {sheet_name!r} must not start or end with an apostrophe")
    if sheet_name.casefold() == "history":
        raise ValueError("'History' is a sheet name Excel reserves")
    folded = sheet_name.casefold()
    if any(name.casefold() == folded for name in existing):
        raise ValueError(f"duplicate sheet name {sheet_name!r} (Excel ignores case)")
