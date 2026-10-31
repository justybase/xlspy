"""Named cell styles compatible with openpyxl."""

from __future__ import annotations


class NamedStyle:
    def __init__(
        self,
        name="Normal",
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
        builtinId=None,
        hidden=False,
    ):
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
        self.builtinId = builtinId
        self.hidden = hidden

    def __repr__(self):
        return f"<NamedStyle {self.name!r}>"


class NamedStyleList(list):
    @property
    def names(self):
        return [item if isinstance(item, str) else item.name for item in self]

    def add(self, style) -> None:
        name = style if isinstance(style, str) else style.name
        self[:] = [
            item for item in self
            if (item if isinstance(item, str) else item.name) != name
        ]
        self.append(style)
