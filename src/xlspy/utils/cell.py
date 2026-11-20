"""A1/range helpers compatible with ``openpyxl.utils.cell``."""

from __future__ import annotations

import re
from functools import lru_cache

from .exceptions import CellCoordinatesException

_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_ABS_RE = re.compile(
    r"^\$?(?P<min_col>[A-Za-z]{1,3})?"
    r"\$?(?P<min_row>\d+)?"
    r"(?::\$?(?P<max_col>[A-Za-z]{1,3})?"
    r"\$?(?P<max_row>\d+)?)?$"
)
# Room for every column (A..XFD is 16,384; the helpers accept up to ZZZ). Mixed-case
# names would otherwise grow the letter cache to ~140,000 entries.
_LETTER_CACHE_SIZE = 20_000

_SHEET_RANGE_RE = re.compile(
    r"^(?:(?P<quoted>'(?:[^']|'')*')|(?P<notquoted>[^!]+))!(?P<cells>.+)$"
)


@lru_cache(maxsize=_LETTER_CACHE_SIZE)
def get_column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to letters (1 -> A)."""
    if not 1 <= col_idx <= 18278:
        raise ValueError(f"Invalid column index {col_idx}")
    letters = []
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


@lru_cache(maxsize=_LETTER_CACHE_SIZE)
def column_index_from_string(col: str) -> int:
    """Convert letters to a 1-based column index (A -> 1)."""
    error = f"'{col}' is not a valid column name. Column names are from A to ZZZ"
    if not isinstance(col, str) or not 1 <= len(col) <= 3:
        raise ValueError(error)
    index = 0
    for char in col.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(error)
        index = index * 26 + ord(char) - 64
    if not 0 < index < 18279:
        raise ValueError(error)
    return index


def coordinate_from_string(coord_string: str) -> tuple[str, int]:
    match = _COORD_RE.match(coord_string)
    if not match:
        raise CellCoordinatesException(f"Invalid cell coordinates ({coord_string})")
    column, row = match.groups()
    row = int(row)
    if not row:
        raise CellCoordinatesException(f"There is no row 0 ({coord_string})")
    return column, row


def coordinate_to_tuple(coordinate: str) -> tuple[int, int]:
    match = _COORD_RE.match(coordinate)
    if not match:
        raise CellCoordinatesException(f"Invalid cell coordinates ({coordinate})")
    column, row = match.groups()
    return int(row), column_index_from_string(column)


def range_boundaries(range_string: str) -> tuple[int | None, int | None, int | None, int | None]:
    """Convert an A1 range to ``(min_col, min_row, max_col, max_row)``.

    A single cell fills both ends, and a whole-column (``A:C``) or whole-row
    (``1:3``) range leaves the missing row/column bounds as ``None``, as
    openpyxl does. ``"A"`` and ``"2"`` are one-ended ranges (a whole column and
    a whole row) and ``""`` has every bound ``None``. Raises ValueError for
    anything else, such as ``"A:1"``.
    """
    message = f"{range_string} is not a valid coordinate or range"
    match = _ABS_RE.match(range_string)
    if match is None:
        raise ValueError(message)
    min_col, min_row, max_col, max_row = match.groups()

    if ":" in range_string:
        cols = (min_col, max_col)
        rows = (min_row, max_row)
        if not (
            all(cols + rows)
            or all(cols) and not any(rows)
            or all(rows) and not any(cols)
        ):
            raise ValueError(message)

    if min_col is not None:
        min_col = column_index_from_string(min_col)
    if min_row is not None:
        min_row = int(min_row)
    max_col = column_index_from_string(max_col) if max_col is not None else min_col
    max_row = int(max_row) if max_row is not None else min_row
    return min_col, min_row, max_col, max_row


def range_to_tuple(range_string: str):
    """Return ``(sheetname, (min_col, min_row, max_col, max_row))``."""
    match = _SHEET_RANGE_RE.match(range_string)
    if match is None:
        raise ValueError("Value must be of the form sheetname!A1:E4")
    sheetname = match.group("quoted") or match.group("notquoted")
    if match.group("quoted"):
        sheetname = sheetname[1:-1].replace("''", "'")
    return sheetname, range_boundaries(match.group("cells"))


def rows_from_range(range_string: str):
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError(f"{range_string} is not a rectangular cell range")
    cols = [get_column_letter(col) for col in range(min_col, max_col + 1)]
    for row in range(min_row, max_row + 1):
        yield tuple(f"{col}{row}" for col in cols)


def cols_from_range(range_string: str):
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError(f"{range_string} is not a rectangular cell range")
    for col in range(min_col, max_col + 1):
        letter = get_column_letter(col)
        yield tuple(f"{letter}{row}" for row in range(min_row, max_row + 1))


def get_column_interval(start, end) -> list[str]:
    if isinstance(start, str):
        start = column_index_from_string(start)
    if isinstance(end, str):
        end = column_index_from_string(end)
    return [get_column_letter(index) for index in range(start, end + 1)]


def absolute_coordinate(coord_string: str) -> str:
    match = _ABS_RE.match(coord_string)
    if match is None:
        raise ValueError(f"{coord_string} is not a valid coordinate range")
    c1, r1, c2, r2 = match.groups()
    first = (f"${c1}" if c1 else "") + (f"${r1}" if r1 else "")
    if c2 is None and r2 is None:
        return first
    second = (f"${c2}" if c2 else "") + (f"${r2}" if r2 else "")
    return f"{first}:{second}"


def quote_sheetname(sheetname: str) -> str:
    return "'" + sheetname.replace("'", "''") + "'"
