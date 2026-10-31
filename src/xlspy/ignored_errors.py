"""Ignored errors: the cell ranges whose green-triangle checks Excel skips."""

from __future__ import annotations

from dataclasses import dataclass, fields

from .refs import parse_cell_range

# The ``ignoredError`` flags, as spelled in the file.
FLAGS = (
    "evalError", "twoDigitTextYear", "numberStoredAsText", "formula",
    "formulaRange", "unlockedFormula", "emptyCellReference",
    "listDataValidation", "calculatedColumn",
)


@dataclass(slots=True)
class IgnoredError:
    """Error checks switched off for ``sqref`` (space-separated ranges).

    Each flag names a check: ``numberStoredAsText`` for "number stored as
    text", ``formula`` for "inconsistent formula", ``formulaRange`` for
    "formula omits adjacent cells", ``unlockedFormula``, ``evalError``,
    ``twoDigitTextYear``, ``emptyCellReference``, ``listDataValidation`` and
    ``calculatedColumn``.
    """

    sqref: str
    evalError: bool | None = None
    twoDigitTextYear: bool | None = None
    numberStoredAsText: bool | None = None
    formula: bool | None = None
    formulaRange: bool | None = None
    unlockedFormula: bool | None = None
    emptyCellReference: bool | None = None
    listDataValidation: bool | None = None
    calculatedColumn: bool | None = None

    def __post_init__(self):
        if not isinstance(self.sqref, str) or not self.sqref.strip():
            raise ValueError("ignored error needs a range such as 'A1:B5'")
        for part in self.sqref.split():
            parse_cell_range(part)
        self.sqref = " ".join(self.sqref.split())


def ignored_errors_xml(items, prefix: str = "") -> str:
    """The ``ignoredErrors`` element for ``items``, or an empty string."""
    entries = []
    for item in items:
        attrs = f' sqref="{item.sqref}"' + "".join(
            f' {name}="{int(bool(getattr(item, name)))}"'
            for name in FLAGS if getattr(item, name) is not None
        )
        entries.append(f"<{prefix}ignoredError{attrs}/>")
    if not entries:
        return ""
    return f"<{prefix}ignoredErrors>{''.join(entries)}</{prefix}ignoredErrors>"


def snapshot(items) -> tuple:
    return tuple(tuple(getattr(item, field.name) for field in fields(item)) for item in items)
