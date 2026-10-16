"""Cell formatting shared by the XLSX and XLSB writers."""

from __future__ import annotations

from dataclasses import dataclass, replace
from functools import lru_cache
import re
from typing import Mapping


@dataclass(frozen=True, slots=True)
class Style:
    """A reusable Excel cell style.

    Colors are RGB hex strings (``"008000"`` or ``"#008000"``).  ``None``
    means “leave this property unchanged”, which also makes styles composable
    when a header style and a range style overlap.

    The initial API supports font/fill colors, bold, italic, underline,
    number formats, horizontal/vertical alignment, and wrapping.
    """

    fill: str | None = None
    font_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    number_format: str | None = None
    horizontal: str | None = None
    vertical: str | None = None
    wrap: bool | None = None

    def __post_init__(self):
        for name in ("fill", "font_color"):
            value = getattr(self, name)
            if value is not None:
                object.__setattr__(self, name, normalize_color(value))
        if self.number_format is not None and not isinstance(self.number_format, str):
            raise TypeError("number_format must be a string or None")
        if self.horizontal is not None and self.horizontal not in {
            "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"
        }:
            raise ValueError("invalid horizontal alignment")
        if self.vertical is not None and self.vertical not in {
            "top", "center", "bottom", "justify", "distributed"
        }:
            raise ValueError("invalid vertical alignment")
        for name in ("bold", "italic", "underline", "wrap"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{name} must be bool or None")


def normalize_color(value: str) -> str:
    """Validate and normalize a six-digit RGB color."""
    if not isinstance(value, str):
        raise TypeError("colors must be strings such as '#008000'")
    color = value.removeprefix("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        raise ValueError(f"invalid RGB color {value!r}; expected '#RRGGBB'")
    return color


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_EMPTY_STYLES = {}


def _column_index(label: str) -> int:
    result = 0
    for char in label.upper():
        result = result * 26 + ord(char) - 64
    return result - 1


def parse_cell_range(reference: str) -> tuple[int, int, int, int]:
    """Parse an A1 cell or rectangular range into zero-based bounds."""
    if not isinstance(reference, str):
        raise TypeError("style range keys must be A1 references such as 'A1:D1'")
    parts = reference.split(":")
    if len(parts) not in (1, 2):
        raise ValueError(f"invalid A1 range {reference!r}")
    matches = [_CELL_RE.fullmatch(part) for part in parts]
    if any(match is None for match in matches):
        raise ValueError(f"invalid A1 range {reference!r}")
    (c1, r1), (c2, r2) = (matches[0].groups(), matches[-1].groups())
    bounds = (_column_index(c1), int(r1) - 1, _column_index(c2), int(r2) - 1)
    if bounds[2] < bounds[0] or bounds[3] < bounds[1]:
        raise ValueError(f"range end precedes its start in {reference!r}")
    if bounds[2] >= 16384 or bounds[3] >= 1048576:
        raise ValueError(f"range exceeds Excel limits: {reference!r}")
    return bounds


def normalize_style_ranges(styles: Mapping[str, Style] | None):
    if styles is None:
        return ()
    if not isinstance(styles, Mapping):
        raise TypeError("styles must be a mapping of A1 ranges to Style objects")
    result = []
    for reference, style in styles.items():
        if not isinstance(style, Style):
            raise TypeError(f"style for {reference!r} must be a Style instance")
        result.append((*parse_cell_range(reference), style))
    return tuple(result)


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
    return {column: style for column in range(row_length)}


def styles_for_row(row_index: int, row_length: int, ranges, header_style: Style | None = None):
    """Return sparse column->style mappings; no allocation for unstyled rows."""
    if not ranges and header_style is None:
        return _EMPTY_STYLES
    if row_index == 0 and header_style is not None and not ranges:
        return _uniform_row_styles(row_length, merge_styles(Style(bold=True), header_style))
    if len(ranges) == 1 and header_style is None:
        first_col, first_row, last_col, last_row, style = ranges[0]
        if (first_row <= row_index <= last_row and first_col == 0
                and last_col >= row_length - 1):
            return _uniform_row_styles(row_length, style)
    styles = {}
    if row_index == 0 and header_style is not None:
        effective_header = merge_styles(Style(bold=True), header_style)
        for column in range(row_length):
            styles[column] = effective_header
    for first_col, first_row, last_col, last_row, style in ranges:
        if first_row <= row_index <= last_row:
            for column in range(first_col, min(last_col, row_length - 1) + 1):
                styles[column] = merge_styles(styles.get(column), style)
    return styles or _EMPTY_STYLES
