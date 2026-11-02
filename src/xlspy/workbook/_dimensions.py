"""Row and column dimension objects and their holder."""

from __future__ import annotations
from collections.abc import Callable
from typing import Any
from ..refs import column_label
from ._coords import _coerce_column, _coerce_row


class Dimension:
    """The size, visibility and outline settings of a row or column, with openpyxl's
    attribute names.
    """
    def __init__(self, index: str | int) -> None:
        self.index = index
        self.min: int | None = None
        self.max: int | None = None
        self.width: float | None = None
        self.height: float | None = None
        self.hidden: bool = False
        self.bestFit: bool | None = None
        self.autoSize: bool | None = None
        self.outlineLevel: int = 0
        self.collapsed: bool = False
        self.style: int | None = None

    @property
    def customWidth(self) -> bool:
        """True when a width is set."""
        return self.width is not None

    @property
    def customHeight(self) -> bool:
        """True when a height is set."""
        return self.height is not None

    @property
    def outline_level(self) -> int:
        """The outline (group) level; the same as ``outlineLevel``."""
        return self.outlineLevel

    @outline_level.setter
    def outline_level(self, value: int) -> None:
        self.outlineLevel = value


class ColumnDimension(Dimension):
    """A column's width, visibility and outline level: ``ws.column_dimensions["B"]``."""


class RowDimension(Dimension):
    """A row's height, visibility and outline level: ``ws.row_dimensions[3]``."""


class DimensionHolder(dict[Any, Any]):
    def __init__(self, factory: Callable[[Any], Dimension], transform: Callable[[Any], Any]) -> None:
        super().__init__()
        self._factory = factory
        self._transform = transform

    def __getitem__(self, key: Any) -> Any:
        key = self._transform(key)
        if key not in self:
            super().__setitem__(key, self._factory(key))
        return super().__getitem__(key)

    def __setitem__(self, key: Any, value: Any) -> None:
        super().__setitem__(self._transform(key), value)

    def get(self, key: Any, default: Any = None) -> Any:
        """Like ``dict.get``, without creating an entry; the key may be a letter or a number."""
        return super().get(self._transform(key), default)

    def __contains__(self, key: object) -> bool:
        return super().__contains__(self._transform(key))

    def group(self, start: int | str, end: int | str | None = None, outline_level: int = 1,
              hidden: bool = False) -> None:
        """Group consecutive rows or columns, as openpyxl does."""
        if end is None:
            end = start
        if self._factory is ColumnDimension:
            first = _coerce_column(start)
            last = _coerce_column(end)
            for index in range(first, last + 1):
                dimension = self[column_label(index - 1)]
                dimension.outlineLevel = outline_level
                dimension.hidden = hidden
                dimension.min = first
                dimension.max = last
        else:
            for index in range(int(start), int(end) + 1):
                dimension = self[index]
                dimension.outlineLevel = outline_level
                dimension.hidden = hidden


def _column_key(key: str | int) -> str:
    if isinstance(key, int) and not isinstance(key, bool):
        if not 1 <= key <= 16384:
            raise ValueError("column index must be 1..16384")
        return column_label(key - 1)
    if not isinstance(key, str):
        raise TypeError("column key must be a letter or 1-based int")
    column = _coerce_column(key)
    return column_label(column - 1) if column else ""


def _row_key(key: int) -> int:
    return _coerce_row(key)
