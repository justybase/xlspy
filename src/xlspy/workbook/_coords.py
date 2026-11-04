"""Cell coordinates, sheet ranges and multi-range helpers."""

from __future__ import annotations
from collections.abc import Iterable, Iterator
import re
from typing import Any, Self, SupportsIndex
from ..refs import RangeIndex, column_label, parse_cell_range


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


def _in_sheet_order(refs: Iterable[str]) -> list[str]:
    """A1 cell references sorted by row, then column.

    Pending edits are kept in sets; saving walks them in this order so the same
    edits always produce the same bytes (string hashes vary between runs)."""
    def position(ref: str) -> tuple[int, int]:
        column, row, _, _ = parse_cell_range(ref)
        return row, column
    return sorted(refs, key=position)


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
        """The first column, 1-based."""
        return parse_cell_range(self)[0] + 1

    @property
    def min_row(self) -> int:
        """The first row, 1-based."""
        return parse_cell_range(self)[1] + 1

    @property
    def max_col(self) -> int:
        """The last column, 1-based."""
        return parse_cell_range(self)[2] + 1

    @property
    def max_row(self) -> int:
        """The last row, 1-based."""
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
        """The range as text, such as ``"A1:C3"``."""
        return str(self)

    @property
    def bounds(self) -> tuple[int, int, int, int]:
        """``(min_col, min_row, max_col, max_row)``, 1-based."""
        return (self.min_col, self.min_row, self.max_col, self.max_row)

    @property
    def size(self) -> dict[str, int]:
        """The size as ``{"columns": n, "rows": m}``."""
        return {
            "columns": self.max_col - self.min_col + 1,
            "rows": self.max_row - self.min_row + 1,
        }

    def shift(self, col_shift: int = 0, row_shift: int = 0) -> CellRange:
        """Return the range moved right by ``col_shift`` and down by ``row_shift``.

        Raises ValueError when the result leaves the worksheet.
        """
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
        """Return the range grown by the given number of cells on each side.

        Raises ValueError when the result leaves the worksheet.
        """
        min_col = self.min_col - left
        min_row = self.min_row - up
        max_col = self.max_col + right
        max_row = self.max_row + down
        if min_col < 1 or min_row < 1 or max_col > 16384 or max_row > 1048576:
            raise ValueError("expanded range is outside the worksheet")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def shrink(self, right: int = 0, bottom: int = 0, left: int = 0, top: int = 0) -> CellRange:
        """Return the range made smaller by the given number of cells on each side."""
        return self.expand(right=-right, down=-bottom, left=-left, up=-top)

    def union(self, other: CellRange | str) -> CellRange:
        """Return the smallest range that covers this range and ``other``."""
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return CellRange(
            f"{column_label(min(self.min_col, other.min_col) - 1)}"
            f"{min(self.min_row, other.min_row)}:"
            f"{column_label(max(self.max_col, other.max_col) - 1)}"
            f"{max(self.max_row, other.max_row)}"
        )

    def intersection(self, other: CellRange | str) -> CellRange:
        """Return the cells this range shares with ``other``.

        Raises ValueError when the ranges do not overlap.
        """
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        min_col = max(self.min_col, other.min_col)
        min_row = max(self.min_row, other.min_row)
        max_col = min(self.max_col, other.max_col)
        max_row = min(self.max_row, other.max_row)
        if min_col > max_col or min_row > max_row:
            raise ValueError(f"Range {self} doesn't intersect {other}")
        return CellRange(f"{column_label(min_col - 1)}{min_row}:{column_label(max_col - 1)}{max_row}")

    def isdisjoint(self, other: CellRange | str) -> bool:
        """True when this range and ``other`` share no cell."""
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            self.max_col < other.min_col
            or other.max_col < self.min_col
            or self.max_row < other.min_row
            or other.max_row < self.min_row
        )

    def issubset(self, other: CellRange | str) -> bool:
        """True when every cell of this range is in ``other``."""
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return (
            other.min_col <= self.min_col <= self.max_col <= other.max_col
            and other.min_row <= self.min_row <= self.max_row <= other.max_row
        )

    def issuperset(self, other: CellRange | str) -> bool:
        """True when every cell of ``other`` is in this range."""
        other = other if isinstance(other, CellRange) else CellRange(str(other))
        return other.issubset(self)


class _RangeList(list):
    """A list that counts its changes, so an index built from it knows when it is stale.

    Any change, including a replacement of one item in place, moves ``version`` on.
    """

    def __init__(self, items: Iterable[CellRange] = ()) -> None:
        super().__init__(items)
        self.version = 0

    def _changed(self) -> None:
        self.version += 1

    def __setitem__(self, index: SupportsIndex | slice, value: Any) -> None:
        super().__setitem__(index, value)
        self._changed()

    def __delitem__(self, index: SupportsIndex | slice) -> None:
        super().__delitem__(index)
        self._changed()

    # list.__add__ takes only a list, so mypy objects to a wider __iadd__; ``+=`` must still count.
    def __iadd__(self, other: Iterable[CellRange]) -> Self:  # type: ignore[misc, override]
        result = super().__iadd__(other)
        self._changed()
        return result

    def __imul__(self, count: SupportsIndex) -> Self:
        result = super().__imul__(count)
        self._changed()
        return result

    def append(self, value: CellRange) -> None:
        super().append(value)
        self._changed()

    def extend(self, values: Iterable[CellRange]) -> None:
        super().extend(values)
        self._changed()

    def insert(self, index: SupportsIndex, value: CellRange) -> None:
        super().insert(index, value)
        self._changed()

    def remove(self, value: CellRange) -> None:
        super().remove(value)
        self._changed()

    def pop(self, index: SupportsIndex = -1) -> CellRange:
        value: CellRange = super().pop(index)
        self._changed()
        return value

    def clear(self) -> None:
        super().clear()
        self._changed()

    def sort(self, *args: Any, **kwargs: Any) -> None:
        super().sort(*args, **kwargs)
        self._changed()

    def reverse(self) -> None:
        super().reverse()
        self._changed()


class MultiCellRange:
    """Several cell ranges, such as a sheet's merged cells, in the order added."""
    def __init__(self, ranges: Iterable[str] | None = None) -> None:
        self._ranges = _RangeList()
        self._index = RangeIndex()
        self._present: set[str] = set()
        self._synced = -1
        if ranges is not None:
            for value in ranges:
                self.add(value)

    @property
    def ranges(self) -> list[CellRange]:
        """The ranges in the order added. Code may change the list; the index follows it."""
        return self._ranges

    @ranges.setter
    def ranges(self, value: Iterable[CellRange]) -> None:
        self._ranges = _RangeList(value)

    def _sync(self) -> None:
        """Rebuild the index when ``ranges`` changed since it was last built."""
        ranges = self._ranges
        if self._synced == ranges.version:
            return
        self._index = RangeIndex()
        for cell_range in ranges:
            self._index.add(parse_cell_range(cell_range), cell_range)
        self._present = set(ranges)
        self._synced = ranges.version

    def add(self, value: str) -> None:
        """Add a range unless it is already present."""
        cell_range = CellRange(value)
        self._sync()
        if cell_range not in self._present:
            self._ranges.append(cell_range)
            self._present.add(cell_range)
            self._index.add(parse_cell_range(cell_range), cell_range)
            self._synced = self._ranges.version

    def remove(self, value: str) -> None:
        """Remove a range; raises ValueError when it is not present."""
        cell_range = CellRange(value)
        self._sync()
        if cell_range not in self._present:
            raise ValueError(f"range {value!r} is not merged")
        self._ranges.remove(cell_range)
        self._present.discard(cell_range)
        self._index.remove(parse_cell_range(cell_range), cell_range)
        self._synced = self._ranges.version

    def first_overlapping(self, bounds: tuple[int, int, int, int]) -> CellRange | None:
        """The first range added that shares a cell with zero-based ``bounds``, or None."""
        self._sync()
        found = self._index.overlapping(bounds)
        return found[0][4] if found else None

    def __contains__(self, item: object) -> bool:
        if not isinstance(item, str):
            return False
        if ":" not in item:
            try:
                col, row, _, _ = parse_cell_range(item)
            except (TypeError, ValueError):
                return False
            return self.first_overlapping((col, row, col, row)) is not None
        try:
            cell_range = CellRange(item)
        except (TypeError, ValueError):
            return False
        self._sync()
        return cell_range in self._present

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
