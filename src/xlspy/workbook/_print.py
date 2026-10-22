"""Print setup of a model worksheet, with openpyxl's attribute names.

The objects are plain mutable holders, as in openpyxl
(``ws.page_setup.orientation = "landscape"``). ``print_settings_for`` turns a
sheet's values into the writers' ``PrintSettings``.
"""

from __future__ import annotations

from dataclasses import dataclass, fields

from ..printing import PageMargins as _WriterMargins
from ..printing import PrintSettings


@dataclass
class PageSetup:
    """``ws.page_setup``: ``orientation`` ("portrait"/"landscape"),
    ``paperSize`` (an Excel code such as 9 for A4, or a name such as "A4"),
    ``fitToWidth``/``fitToHeight`` (used when
    ``ws.sheet_properties.pageSetUpPr.fitToPage`` is true) and ``scale``."""

    orientation: str | None = None
    paperSize: int | str | None = None
    fitToWidth: int | None = None
    fitToHeight: int | None = None
    scale: int | None = None


@dataclass
class PageMargins:
    """``ws.page_margins`` in inches (openpyxl's defaults)."""

    left: float = 0.7
    right: float = 0.7
    top: float = 0.75
    bottom: float = 0.75
    header: float = 0.3
    footer: float = 0.3


@dataclass
class PrintOptions:
    """``ws.print_options``: centring on the page."""

    horizontalCentered: bool | None = None
    verticalCentered: bool | None = None


@dataclass
class PageSetupProperties:
    """``ws.sheet_properties.pageSetUpPr``."""

    fitToPage: bool | None = None


@dataclass
class _HeaderFooterPart:
    text: str | None = None


class HeaderFooterItem:
    """``ws.oddHeader`` / ``ws.oddFooter``: ``left``, ``center`` and ``right``
    parts, each with a ``text`` in Excel's header codes (``&P`` page number,
    ``&N`` page count, ``&D`` date, ``&&`` a literal ampersand)."""

    def __init__(self):
        self.left = _HeaderFooterPart()
        self.center = _HeaderFooterPart()
        self.right = _HeaderFooterPart()

    def code(self) -> str | None:
        parts = [
            f"&{tag}{part.text}"
            for tag, part in (("L", self.left), ("C", self.center), ("R", self.right))
            if part.text
        ]
        return "".join(parts) or None


class HeaderFooter:
    """``ws.HeaderFooter``: the odd-page header and footer (used on every page)."""

    def __init__(self):
        self.oddHeader = HeaderFooterItem()
        self.oddFooter = HeaderFooterItem()


def print_settings_for(sheet) -> PrintSettings | None:
    """The sheet's ``PrintSettings``, or None when nothing was set."""
    setup = sheet.page_setup
    fit = bool(sheet.sheet_properties.pageSetUpPr.fitToPage)
    margins = sheet.page_margins
    options = sheet.print_options
    values = {
        "orientation": setup.orientation,
        "paper_size": setup.paperSize,
        "margins": None if margins == PageMargins() else _WriterMargins(
            **{item.name: getattr(margins, item.name) for item in fields(PageMargins)}
        ),
        "print_area": sheet.print_area,
        "repeat_rows": sheet.print_title_rows,
        "repeat_columns": sheet.print_title_cols,
        # As in Excel, fitToWidth/fitToHeight count only with fitToPage.
        "fit_to_width": (1 if setup.fitToWidth is None else setup.fitToWidth) if fit else None,
        "fit_to_height": (1 if setup.fitToHeight is None else setup.fitToHeight) if fit else None,
        "scale": None if fit else setup.scale,
        "center_horizontally": bool(options.horizontalCentered),
        "center_vertically": bool(options.verticalCentered),
        "header": sheet.oddHeader.code(),
        "footer": sheet.oddFooter.code(),
    }
    if not any(values.values()):
        return None
    return PrintSettings(**values)
