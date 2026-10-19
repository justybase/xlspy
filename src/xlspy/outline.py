"""Row and column groups: the +/- buttons Excel draws beside them.

A group is a run of rows or columns, given as a ``Group`` or a
``(first, last)`` / ``(first, last, collapsed)`` tuple. Groups may nest, and
the nesting depth gives the outline level. Groups that only share some rows
or columns cannot be drawn by Excel, so they are rejected.
"""

from __future__ import annotations

from bisect import bisect_right
import heapq
from dataclasses import dataclass

from .refs import _column_index

MAX_LEVEL = 7  # Excel shows at most seven outline levels.
_LIMITS = {"row": 1048576, "column": 16384}


@dataclass(frozen=True, slots=True)
class Group:
    """Rows or columns from ``first`` to ``last``, both inclusive and zero-based.

    Columns may also be given as letters, such as ``Group("D", "F")``.
    ``collapsed`` hides the group and sets the collapsed flag on the summary
    row or column just after ``last``.
    """

    first: int | str
    last: int | str
    collapsed: bool = False


def _bound(value, kind: str) -> int:
    if kind == "column" and isinstance(value, str) and value.isascii() and value.isalpha():
        index = _column_index(value)
    elif isinstance(value, int) and not isinstance(value, bool):
        index = value
    elif kind == "column":
        raise TypeError("column group bounds must be column letters or 0-based ints")
    else:
        raise TypeError("row group bounds must be 0-based ints")
    if not 0 <= index < _LIMITS[kind]:
        raise ValueError(f"{kind} group bound out of range: {value!r}")
    return index


def _span(item, kind: str) -> tuple[int, int, bool]:
    if isinstance(item, Group):
        first, last, collapsed = item.first, item.last, item.collapsed
    elif isinstance(item, (tuple, list)) and len(item) in (2, 3):
        first, last, *rest = item
        collapsed = rest[0] if rest else False
    else:
        raise TypeError(
            f"{kind}_groups entries must be Group objects or (first, last) tuples"
        )
    if not isinstance(collapsed, bool):
        raise TypeError(f"{kind}_groups collapsed flags must be bool")
    first_index = _bound(first, kind)
    last_index = _bound(last, kind)
    if first_index > last_index:
        raise ValueError(f"{kind} group ends before it starts: {first!r} to {last!r}")
    return first_index, last_index, collapsed


def build_outline(groups, kind: str) -> Outline | None:
    """Validate ``groups`` and return their outline, or ``None`` when there are none.

    ``kind`` is ``"row"`` or ``"column"`` and only changes the bounds and the
    error messages.
    """
    if groups is None:
        return None
    if isinstance(groups, (str, bytes)):
        raise TypeError(f"{kind}_groups must be a collection of groups, not a string")
    try:
        items = list(groups)
    except TypeError:
        raise TypeError(f"{kind}_groups must be a collection or None") from None
    spans = sorted((_span(item, kind) for item in items), key=lambda s: (s[0], -s[1]))
    if not spans:
        return None
    for prev, cur in zip(spans, spans[1:]):
        if prev[:2] == cur[:2]:
            raise ValueError(f"duplicate {kind} group {cur[0]} to {cur[1]}")

    # Sweep the boundaries in order. The stack holds the open groups, outermost
    # first; every group on it contains the current position.
    positions = sorted({first for first, _, _ in spans} | {last + 1 for _, last, _ in spans})
    starts: list[int] = []
    levels: list[int] = []
    hidden: list[bool] = []
    stack: list[tuple[int, int, bool]] = []
    cursor = 0
    for position in positions:
        while stack and stack[-1][1] < position:
            stack.pop()
        while cursor < len(spans) and spans[cursor][0] == position:
            first, last, collapsed = spans[cursor]
            cursor += 1
            if stack and last > stack[-1][1]:
                outer_first, outer_last, _ = stack[-1]
                raise ValueError(
                    f"{kind} groups overlap without nesting: "
                    f"{outer_first} to {outer_last} and {first} to {last}"
                )
            stack.append((first, last, collapsed))
            if len(stack) > MAX_LEVEL:
                raise ValueError(f"{kind} groups nest deeper than {MAX_LEVEL} levels")
        level = len(stack)
        is_hidden = any(collapsed for _, _, collapsed in stack)
        if not levels or levels[-1] != level or hidden[-1] != is_hidden:
            starts.append(position)
            levels.append(level)
            hidden.append(is_hidden)

    # A group ending on the sheet's last row or column has no summary to flag.
    summaries = frozenset(
        last + 1 for _, last, collapsed in spans if collapsed and last + 1 < _LIMITS[kind]
    )
    return Outline(starts, levels, hidden, summaries)


class Outline:
    """Outline levels and flags for one axis, looked up by zero-based index."""

    __slots__ = ("_starts", "_levels", "_hidden", "_summaries", "max_level")

    def __init__(self, starts, levels, hidden, summaries):
        self._starts = starts
        self._levels = levels
        self._hidden = hidden
        self._summaries = summaries
        self.max_level = max(levels)

    def state_at(self, index: int) -> tuple[int, bool]:
        """Outline level and hidden flag at ``index``. Level 0 is outside every group."""
        k = bisect_right(self._starts, index) - 1
        if k < 0:
            return 0, False
        return self._levels[k], self._hidden[k]

    def collapsed_at(self, index: int) -> bool:
        """True when ``index`` is the summary row or column of a collapsed group."""
        return index in self._summaries

    def indices(self) -> set[int]:
        """Every index with an outline level or a collapsed flag, so it needs its own column entry."""
        return self._collect([level > 0 for level in self._levels]) | self._summaries

    def indices_from(self, start: int):
        """Yield, in order, every index from ``start`` on that has an outline level
        or a collapsed flag.

        Rows past the data still need these, or a group reaching beyond the
        last row loses its level and a collapsed group loses its summary flag.
        """
        grouped = (
            index
            for k, begin in enumerate(self._starts)
            if self._levels[k] > 0
            # Every group ends, so the last segment is level 0 and k + 1 exists.
            for index in range(max(begin, start), self._starts[k + 1])
        )
        summaries = sorted(index for index in self._summaries if index >= start)
        previous = None
        for index in heapq.merge(grouped, summaries):
            if index != previous:
                yield index
                previous = index

    def hidden_indices(self) -> set[int]:
        """Every index hidden by a collapsed group."""
        return self._collect(self._hidden)

    def _collect(self, selected: list[bool]) -> set[int]:
        result: set[int] = set()
        for k, start in enumerate(self._starts):
            if selected[k]:
                end = self._starts[k + 1] if k + 1 < len(self._starts) else start
                result.update(range(start, end))
        return result
