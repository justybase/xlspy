"""Named cell styles compatible with openpyxl."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .core import Alignment, Border, Font, GradientFill, PatternFill, Protection


class NamedStyle:
    """A cell style with a name: font, fill, border, alignment, number format and
    protection.

    Register it with ``Workbook.add_named_style``, then assign its name to
    ``cell.style``.
    """
    def __init__(
        self,
        name: str = "Normal",
        font: Font | None = None,
        fill: str | PatternFill | GradientFill | None = None,
        border: Border | None = None,
        alignment: Alignment | None = None,
        number_format: str | None = None,
        protection: Protection | None = None,
        builtinId: int | None = None,
        hidden: bool = False,
    ) -> None:
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
        self.builtinId = builtinId
        self.hidden = hidden

    def __repr__(self) -> str:
        return f"<NamedStyle {self.name!r}>"


class NamedStyleList(list["str | NamedStyle"]):
    @property
    def names(self) -> list[str]:
        """The style names, in order."""
        return [item if isinstance(item, str) else item.name for item in self]

    def add(self, style: str | NamedStyle) -> None:
        """Add a style, replacing one with the same name."""
        name = style if isinstance(style, str) else style.name
        self[:] = [
            item for item in self
            if (item if isinstance(item, str) else item.name) != name
        ]
        self.append(style)
