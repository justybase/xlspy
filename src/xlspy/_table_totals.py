"""The totals-row cell of a table, as the XLSB and XLSX writers encode it."""

from __future__ import annotations

from typing import Any

from ._table_totals_names import ILTA_CUSTOM, ILTA_NONE, ILTA_SUM
from .formulas import Formula
from .hyperlinks import Hyperlink
from .rich_text import CellRichText


def totals_function(cell: Any) -> int:
    """The XLSB ListTotalRowFunction of a totals-row cell.

    Excel refuses a table whose non-empty totals cell is not a label and has no
    function: ILTA_NONE allows only text or an empty cell. Text is its own label,
    a number is the cached result of a sum, a formula is a custom total.
    """
    if cell is None:
        return ILTA_NONE
    if isinstance(cell, Formula):
        return ILTA_CUSTOM
    if isinstance(cell, str):
        return ILTA_CUSTOM if cell.startswith("=") else ILTA_NONE
    if isinstance(cell, (CellRichText, Hyperlink)):
        return ILTA_NONE
    return ILTA_SUM


def totals_label(cell: Any) -> str | None:
    """The BrtBeginListCol label of a totals cell, or None when it has none."""
    if isinstance(cell, Hyperlink):
        return cell.display_text
    if isinstance(cell, CellRichText):
        return str(cell)
    if isinstance(cell, str) and cell and not cell.startswith("="):
        return cell
    return None


def totals_formula(cell: Any) -> str | None:
    """The formula text of a custom totals cell, without its leading ``=``."""
    if isinstance(cell, Formula):
        return cell.text
    if isinstance(cell, str) and cell.startswith("="):
        return cell[1:].lstrip()
    return None


def totals_choice(cell: Any, chosen: str | None = None) -> tuple[str, str | None, str | None]:
    """``(function, label, formula)`` for a totals cell.

    ``chosen`` is an explicit ``ExcelTable.totals_row_functions`` entry; None
    infers it from the cell. A label is only valid for a text cell and a custom
    formula needs a formula cell, so an explicit choice that the cell cannot
    hold is refused instead of writing a table Excel rejects.
    """
    if chosen is None:
        function = totals_function(cell)
        if function == ILTA_CUSTOM:
            return "custom", None, totals_formula(cell)
        if function == ILTA_SUM:
            return "sum", None, None
        label = totals_label(cell)
        return ("label" if label is not None else "none"), label, None

    if chosen == "label":
        label = totals_label(cell)
        if label is None and cell not in (None, ""):
            raise ValueError("a label totals cell must hold text")
        return "label", label, None
    if chosen == "none":
        label = totals_label(cell)
        if label is None and cell not in (None, ""):
            raise ValueError("a none totals cell must hold text or be empty")
        return "none", label, None
    if chosen == "custom":
        formula = totals_formula(cell)
        if formula is None:
            raise ValueError("a custom totals cell must hold a formula")
        return "custom", None, formula
    if totals_formula(cell) is not None:
        raise ValueError(f"a {chosen} totals cell cannot hold a formula")
    return chosen, None, None
