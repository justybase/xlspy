"""Cell coordinates, sheet ranges and multi-range helpers."""

from __future__ import annotations
from collections.abc import Iterable, Iterator
import re
from ..refs import column_label, parse_cell_range


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")


def _coerce_column(column: int | str) -> int:
    if isinstance(column, str):
        match = _CELL_RE.match(column)
        if match:
            column = match.group(1)
        if not re.fullmatch(r"[A-Za-z]{1,3}", column):
            raise ValueError(f"invalid column {column!r}")
        value = 0
        for char in column.upper():
            value = value * 26 + ord(char) - 64
        return value
    if not isinstance(column, int) or isinstance(column, bool):
        raise TypeError("column must be a 1-based int or a column letter")
    if not 1 <= column <= 16384:
        raise ValueError("column must be 1..16384")
    return column


def _coerce_row(row: int) -> int:
    if not isinstance(row, int) or isinstance(row, bool):
        raise TypeError("row must be a 1-based int")
    if not 1 <= row <= 1048576:
        raise ValueError("row must be 1..1048576")
    return row


def _coordinate(row: int, column: int) -> str:
    return f"{column_label(column - 1)}{row}"


def _normalize_range(reference: str) -> str:
    if not isinstance(reference, str):
        raise TypeError("a range must be an A1 string")
    first, _, last = reference.partition(":")
    if not last:
        last = first
    c1, r1, c2, r2 = parse_cell_range(f"{first}:{last}")
    return f"{_coordinate(r1 + 1, c1 + 1)}:{_coordinate(r2 + 1, c2 + 1)}"


def _normalize_sqref(reference: str) -> str:
    """A range or space-separated ranges, each normalized."""
    if not isinstance(reference, str):
        raise TypeError("a range must be an A1 string")
    parts = reference.split()
    if not parts:
        raise ValueError("a range must not be empty")
    return " ".join(_normalize_range(part) for part in parts)


class CellRange(str):
    """An A1 range string with openpyxl-style boundary properties."""

    def __new__(cls, value: str) -> CellRange:
        normalized = _normalize_range(value)
        return str.__new__(cls, normalized)

    @property
    def min_col(self) -> int:
        return parse_cell_range(self)[0] + 1

    @property
    def min_row(self) -> int:
        return parse_cell_range(self)[1] + 1

    @property
    def max_col(self) -> int:
        return parse_cell_range(self)[2] + 1

    @property
    def max_row(self) -> int:
        return parse_cell_range(self)[3] + 1

    def __contains__(self, item: object) -> bool:
        if not isinstance(item, str):
            return False
        try:
            col, row, _, _ = parse_cell_range(item)
        except (TypeError, ValueError):
            return False
        return (
            self.min_col - 1 <= col <= self.max_col - 1
            and self.min_row - 1 <= row <= self.max_row - 1
        )

    @property
    def coord(self) -> str:
        return str(self)

    @property
    def bounds(self) -> tuple[int, int, int, int]:
        return (self.min_col, self.min_row, self.max_col, self.max_row)

    @property
    def size(self) -> dict[str, int]:
        return {
            "columns": self.max_col - self.min_col + 1,
            "rows": self.max_row - self.min_row + 1,
        }

    def shift(self, col_shift: int = 0, row_shift: int = 0) -> CellRange:
        min_col = self.min_col + col_shift
        min_row = self.min_row + row_shift
        max_col = self.max_col + col_shift
        max_row = self.max_row + row_shift
        if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
            raise ValueError(
                f"Invalid shift value: col_shift={col_shift}, row_shift={row_shift}"
            )
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def expand(self, right: int = 0, down: int = 0, left: int = 0, up: int = 0) -> CellRange:
        min_col = self.min_col - left
        min_row = self.min_row - up
        max_col = self.max_col + right
        max_row = self.max_row + down
        if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
            raise ValueError("expanded range is outside the worksheet")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def shrink(self, right: int = 0, bottom: int = 0, left: int = 0, top: int = 0) -> CellRange:
        return self.expand(right=-right, down=-bottom, left=-left, up=-top)

    def union(self, other: CellRange | str) -> CellRange:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return CellRange(
            f"{column_label(min(self.min_col, other.min_col) - 1)}"
            f"{min(self.min_row, other.min_row)}:"
            f"{column_label(max(self.max_col, other.max_col) - 1)}"
            f"{max(self.max_row, other.max_row)}"
        )

    def intersection(self, other: CellRange | str) -> CellRange:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        min_col = max(self.min_col, other.min_col)
        min_row = max(self.min_row, other.min_row)
        max_col = min(self.max_col, other.max_col)
        max_row = min(self.max_row, other.max_row)
        if min_col > max_col or min_row > max_row:
            raise ValueError(f"Range {self} doesn't intersect {other}")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def isdisjoint(self, other: CellRange | str) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            self.max_col < other.min_col
            or other.max_col < self.min_col
            or self.max_row < other.min_row
            or other.max_row < self.min_row
        )

    def issubset(self, other: CellRange | str) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            other.min_col <= self.min_col <= self.max_col <= other.max_col
            and other.min_row <= self.min_row <= self.max_row <= other.max_row
        )

    def issuperset(self, other: CellRange | str) -> bool:
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return other.issubset(self)


class MultiCellRange:
    def __init__(self, ranges: Iterable[str] | None = None) -> None:
        self.ranges: list[CellRange] = []
        if ranges is not None:
            for value in ranges:
                self.add(value)

    def add(self, value: str) -> None:
        cell_range = CellRange(value)
        if cell_range not in self.ranges:
            self.ranges.append(cell_range)

    def remove(self, value: str) -> None:
        cell_range = CellRange(value)
        for index, existing in enumerate(self.ranges):
            if existing == cell_range:
                del self.ranges[index]
                return
        raise ValueError(f"range {value!r} is not merged")

    def __contains__(self, item: object) -> bool:
        if isinstance(item, str) and ":" not in item:
            return any(item in cell_range for cell_range in self.ranges)
        if not isinstance(item, str):
            return False
        try:
            cell_range = CellRange(item)
        except (TypeError, ValueError):
            return False
        return cell_range in self.ranges

    def __iter__(self) -> Iterator[CellRange]:
        return iter(self.ranges)

    def __len__(self) -> int:
        return len(self.ranges)

    def __bool__(self) -> bool:
        return bool(self.ranges)

    def __str__(self) -> str:
        return " ".join(self.ranges)

    def __repr__(self) -> str:
        return f"MultiCellRange({list(self.ranges)!r})"
