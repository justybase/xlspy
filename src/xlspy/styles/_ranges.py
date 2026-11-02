"""Row and column style ranges, and sheet-view option checks."""

from __future__ import annotations
from dataclasses import replace
from functools import lru_cache
import re
from typing import Any, overload
from collections.abc import Mapping
from ..refs import _column_index, parse_cell_range
from ._attrs import _coerce_color
from ._keys import _color_key
from ._style import Style


_EMPTY_STYLES: dict[int, Style] = {}


class StyleRanges(tuple):
    """``(first_col, first_row, last_col, last_row, style)`` entries in the
    caller's order, plus a sweep that finds the ranges covering a row.

    Writers ask for rows in increasing order, so the sweep adds ranges as
    their first row is reached and drops them past their last row: the cost
    per row follows the ranges on that row, not all ranges of the sheet.
    Asking for an earlier row restarts the sweep.
    """

    def __init__(self, entries=()):
        self._by_first_row = sorted(range(len(self)), key=lambda i: self[i][1])
        self._restart()

    def _restart(self):
        self._next = 0
        self._row = -1
        self._open = []
        self._expires = None
        self._active: tuple[Any, ...] = ()

    def active(self, row_index: int):
        """``(first_col, last_col, style)`` of the ranges covering ``row_index``,
        in the caller's order."""
        if row_index == self._row:
            return self._active
        if row_index < self._row:
            self._restart()
        changed = False
        if self._expires is not None and row_index > self._expires:
            self._open = [i for i in self._open if self[i][3] >= row_index]
            changed = True
        order = self._by_first_row
        position = self._next
        while position < len(order) and self[order[position]][1] <= row_index:
            index = order[position]
            if self[index][3] >= row_index:
                self._open.append(index)
                changed = True
            position += 1
        self._next = position
        if changed:
            self._open.sort()
            self._expires = min((self[i][3] for i in self._open), default=None)
            self._active = tuple(
                (self[i][0], self[i][2], self[i][4]) for i in self._open
            )
        self._row = row_index
        return self._active


_LAST_ROW = 1048575
_LAST_COL = 16383
# Whole columns ("B:D") or columns from a row down ("B2:D"), and whole rows ("1:3").
_OPEN_COLUMNS_RE = re.compile(r"\$?([A-Za-z]{1,3})(?:\$?([1-9][0-9]*))?:\$?([A-Za-z]{1,3})")
_OPEN_ROWS_RE = re.compile(r"\$?([1-9][0-9]*):\$?([1-9][0-9]*)")


def parse_style_range(reference: str) -> tuple[int, int, int, int]:
    """Bounds of a ``styles`` key: an A1 cell or range, whole columns
    (``"B:D"``), columns from a row down (``"B2:D"``) or whole rows
    (``"1:3"``). Open ranges reach Excel's last row or column; the writers
    style only the cells of the rows they write."""
    if isinstance(reference, str):
        match = _OPEN_COLUMNS_RE.fullmatch(reference)
        if match is not None:
            first, row, last = match.groups()
            bounds = (_column_index(first), int(row) - 1 if row else 0,
                      _column_index(last), _LAST_ROW)
        else:
            match = _OPEN_ROWS_RE.fullmatch(reference)
            if match is not None:
                bounds = (0, int(match[1]) - 1, _LAST_COL, int(match[2]) - 1)
        if match is not None:
            if bounds[2] < bounds[0] or bounds[3] < bounds[1]:
                raise ValueError(f"range end precedes its start in {reference!r}")
            if bounds[2] > _LAST_COL or bounds[1] > _LAST_ROW or bounds[3] > _LAST_ROW:
                raise ValueError(f"range exceeds Excel limits: {reference!r}")
            return bounds
    return parse_cell_range(reference)


def normalize_style_ranges(styles: Mapping[str, Style] | None):
    if styles is None:
        return StyleRanges()
    if not isinstance(styles, Mapping):
        raise TypeError("styles must be a mapping of A1 ranges to Style objects")
    result = []
    for reference, style in styles.items():
        if not isinstance(style, Style):
            raise TypeError(f"style for {reference!r} must be a Style instance")
        result.append((*parse_style_range(reference), style))
    return StyleRanges(result)


def check_header_rows(header_rows) -> int:
    if isinstance(header_rows, bool) or not isinstance(header_rows, int):
        raise TypeError("header_rows must be an int")
    if not 1 <= header_rows <= _LAST_ROW + 1:
        raise ValueError("header_rows must be at least 1")
    return header_rows


def header_style_ranges(style_ranges, header_style, header_rows: int):
    """``(style_ranges, header_style)`` for a header of ``header_rows`` rows.

    One header row keeps the header_style path. A taller header becomes a
    whole-row range ahead of the caller's ranges, so they still overlay it.
    """
    if header_style is None or header_rows == 1:
        return style_ranges, header_style
    header = (0, 0, _LAST_COL, header_rows - 1, merge_styles(Style(bold=True), header_style))
    return StyleRanges((header, *style_ranges)), None


def check_zoom(zoom) -> int | None:
    """A sheet's zoom in percent (10-400), or None for Excel's 100%."""
    if zoom is None:
        return None
    if isinstance(zoom, bool) or not isinstance(zoom, int):
        raise TypeError("zoom must be an int percentage or None")
    if not 10 <= zoom <= 400:
        raise ValueError("zoom must be between 10 and 400 percent")
    return zoom


def check_sheet_view(show_row_col_headers, show_zero_values, right_to_left) -> tuple[bool, bool, bool]:
    """The sheet view switches as ``(headers, zeros, right_to_left)``; each must be a bool."""
    for name, value in (("show_row_col_headers", show_row_col_headers),
                        ("show_zero_values", show_zero_values), ("right_to_left", right_to_left)):
        if not isinstance(value, bool):
            raise TypeError(f"{name} must be a bool")
    return show_row_col_headers, show_zero_values, right_to_left


def tab_color_key(tab_color):
    """Color key of a sheet's ``tab_color`` (a Color or color string), or None."""
    if tab_color is None:
        return None
    return _color_key(_coerce_color(tab_color, "tab_color"))


@overload
def merge_styles(base: Style, overlay: Style | None) -> Style: ...
@overload
def merge_styles(base: Style | None, overlay: Style) -> Style: ...
@overload
def merge_styles(base: None, overlay: None) -> None: ...
def merge_styles(base: Style | None, overlay: Style | None) -> Style | None:
    if base is None:
        return overlay
    if overlay is None:
        return base
    return replace(base, **{
        field: getattr(overlay, field)
        for field in Style.__dataclass_fields__
        if getattr(overlay, field) is not None
    })


@lru_cache(maxsize=128)
def _uniform_row_styles(row_length: int, style: Style):
    # These cached dictionaries are only read by the writers.
    return dict.fromkeys(range(row_length), style)


@lru_cache(maxsize=128)
def _row_styles_for_pattern(row_length: int, ranges, header_style: Style | None):
    """Reuse the same immutable-in-practice map for rows with the same styles."""
    styles: dict[int, Style] = {}
    if header_style is not None:
        for column in range(row_length):
            styles[column] = header_style
    for first_col, last_col, style in ranges:
        for column in range(first_col, min(last_col, row_length - 1) + 1):
            merged = merge_styles(styles.get(column), style)
            if merged is not None:
                styles[column] = merged
    return styles or _EMPTY_STYLES


def styles_for_row(row_index: int, row_length: int, ranges, header_style: Style | None = None):
    """Return sparse column->style mappings; no allocation for unstyled rows."""
    if not ranges:
        if header_style is None or row_index:
            return _EMPTY_STYLES
        return _uniform_row_styles(row_length, merge_styles(Style(bold=True), header_style))
    if len(ranges) == 1 and header_style is None:
        first_col, first_row, last_col, last_row, style = ranges[0]
        if (first_row <= row_index <= last_row and first_col == 0
                and last_col >= row_length - 1):
            return _uniform_row_styles(row_length, style)
    if isinstance(ranges, StyleRanges):
        active = ranges.active(row_index)
    else:
        active = tuple(
            (first_col, last_col, style)
            for first_col, first_row, last_col, last_row, style in ranges
            if first_row <= row_index <= last_row
        )
    if row_index == 0 and header_style is not None:
        effective_header = merge_styles(Style(bold=True), header_style)
    elif not active:
        return _EMPTY_STYLES
    else:
        effective_header = None
    return _row_styles_for_pattern(row_length, active, effective_header)
