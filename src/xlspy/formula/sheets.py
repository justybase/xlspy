"""Sheet references inside A1 formulas."""

from __future__ import annotations

import re

# A string literal, or a sheet prefix (quoted or plain) followed by "!".
_PARTS = re.compile(r"""("(?:[^"]|"")*")|('(?:[^']|'')+'|(?<![\w.\]])[^\s'!"()+\-*/^&=<>,;:{}%\[\]]+)!""")
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
    Excel. References to other workbooks (``[1]Sheet!A1``) are not touched.
    """
    folded = old.casefold()

    def replace(match):
        if match.group(1) is not None:
            return match.group(0)
        prefix = match.group(2)
        name = prefix[1:-1].replace("''", "'") if prefix.startswith("'") else prefix
        if name.casefold() != folded:
            return match.group(0)
        return sheet_prefix(new) + "!"

    return _PARTS.sub(replace, formula)
