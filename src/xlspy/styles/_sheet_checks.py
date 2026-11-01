"""Sheet-level normalisers: hidden rows and columns, conditional-format priorities."""

from __future__ import annotations
from ..refs import _column_index


def explicit_cf_priorities(rules, *taken: set) -> list[int]:
    """The explicit priorities of ``rules``, checked against each other and ``taken``.

    Nothing is registered, so the writer can reserve them only once the whole
    sheet's options are valid.
    """
    priorities: list[int] = []
    for rule in rules:
        priority = getattr(rule, "priority", None)
        if priority is None:
            continue
        if priority in priorities or any(priority in used for used in taken):
            raise ValueError(f"duplicate conditional formatting priority {priority}")
        priorities.append(priority)
    return priorities


def normalize_hidden_columns(items, kind: str = "hidden_columns") -> set[int]:
    """Validate a collection of column letters or 0-based ints; returns the indices."""
    if items is None:
        return set()
    if isinstance(items, (str, bytes)):
        raise TypeError(f"{kind} must be a collection of columns/rows, not a string")
    try:
        values = list(items)
    except TypeError:
        raise TypeError(f"{kind} must be a collection or None") from None
    result = set()
    for key in values:
        if isinstance(key, str):
            col = _column_index(key)
        elif isinstance(key, int) and not isinstance(key, bool):
            col = key
        else:
            raise TypeError(f"{kind} keys must be column letters or 0-based ints")
        if col < 0 or col >= 16384:
            raise ValueError(f"hidden column out of range: {col}")
        result.add(col)
    return result


def normalize_hidden_rows(items) -> set[int]:
    """Validate a collection of 0-based row ints; returns them as a set."""
    if items is None:
        return set()
    try:
        candidates = list(items)
    except TypeError:
        raise TypeError("hidden_rows must be a collection or None") from None
    result = set()
    for key in candidates:
        if not isinstance(key, int) or isinstance(key, bool) or key < 0:
            raise TypeError("hidden_rows entries must be 0-based row ints")
        result.add(key)
    return result
