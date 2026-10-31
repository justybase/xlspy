"""Sheet references inside A1 formulas."""

from __future__ import annotations

import re

# A string literal, or a sheet prefix followed by "!": quoted, or plain, either
# of them one sheet or a 3D span "First:Last".
_NAME = r"""[^\s'!"()+\-*/^&=<>,;:{}%\[\]]+"""
_PARTS = re.compile(rf"""("(?:[^"]|"")*")|('(?:[^']|'')+'|(?<![\w.\]]){_NAME}(?::{_NAME})?)!""")
_PLAIN = re.compile(r"[A-Za-z_À-￿][\w.]*")
_CELL_LIKE = re.compile(r"(?:[A-Za-z]{1,3}\d+|[RrCc]\d*|[Rr]\d*[Cc]\d*)")


def sheet_prefix(name: str) -> str:
    """``name`` as it appears before ``!`` in a formula, quoted when needed."""
    if _PLAIN.fullmatch(name) and not _CELL_LIKE.fullmatch(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def rename_sheet_references(formula: str, old: str, new: str) -> str:
    """Point the references to sheet ``old`` in ``formula`` at ``new``.

    Text in string literals is left alone; names compare without case, as in
    Excel. Both ends of a 3D reference (``Jan:Mar!B2``) are renamed.
    References to other workbooks (``[1]Sheet!A1``) are not touched.
    """
    folded = old.casefold()

    def replace(match):
        if match.group(1) is not None:
            return match.group(0)
        prefix = match.group(2)
        quoted = prefix.startswith("'")
        # Sheet names cannot hold ":", so a colon always separates a 3D span.
        names = (prefix[1:-1].replace("''", "'") if quoted else prefix).split(":")
        lead = ""
        if not quoted and len(names) == 2 and sheet_prefix(names[0]) != names[0]:
            # "A1:Sheet2!B2": the part before ":" ends a range; it is no sheet.
            lead, names = names[0] + ":", names[1:]
        if not any(name.casefold() == folded for name in names):
            return match.group(0)
        names = [new if name.casefold() == folded else name for name in names]
        if all(sheet_prefix(name) == name for name in names):
            return lead + ":".join(names) + "!"
        return lead + "'" + ":".join(names).replace("'", "''") + "'!"

    return _PARTS.sub(replace, formula)
