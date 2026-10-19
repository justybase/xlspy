"""A1 cell and range helpers shared by the style, chart and feature modules.

This module has no xlspy imports so that every model module can depend on it
without creating an import cycle.
"""

from __future__ import annotations

import re

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")


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


def quote_sheet_name(sheet_name: str) -> str:
    """Quote a sheet name for use in a formula or defined name."""
    return "'" + sheet_name.replace("'", "''") + "'"
