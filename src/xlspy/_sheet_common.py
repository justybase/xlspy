"""Sheet names, visibility states and part paths shared by both updaters."""

from __future__ import annotations

import re
from collections.abc import Iterable

SHEET_STATES = frozenset({"visible", "hidden", "veryHidden"})
STATE_CODES = {"visible": 0, "hidden": 1, "veryHidden": 2}
STATE_NAMES = {code: name for name, code in STATE_CODES.items()}


def check_sheet_state(state: str) -> str:
    """Return ``state``, or raise ValueError for a spelling Excel does not use."""
    if state not in SHEET_STATES:
        raise ValueError(f"invalid sheet state {state!r}")
    return state


def next_sheet_part(names: Iterable[str], extension: str) -> str:
    """The first free ``xl/worksheets/sheetN.<extension>`` for a package."""
    pattern = re.compile(rf"xl/worksheets/sheet(\d+)\.{re.escape(extension)}")
    highest = 0
    for name in names:
        match = pattern.fullmatch(name)
        if match:
            highest = max(highest, int(match.group(1)))
    return f"xl/worksheets/sheet{highest + 1}.{extension}"
