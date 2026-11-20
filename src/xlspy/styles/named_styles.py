"""Named cell styles compatible with openpyxl."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .core import Alignment, Border, Font, GradientFill, PatternFill, Protection


class StyleArray(tuple[int, ...]):
    """openpyxl's nine style ids, with the same positional and attribute names.

    The fields are ``fontId, fillId, borderId, numFmtId, protectionId,
    alignmentId, pivotButton, quotePrefix, xfId``. A style that is not bound to
    a workbook has no ids yet, so every field is 0 — which is what openpyxl's
    own array holds in that case.
    """

    __slots__ = ()

    _NAMES = (
        "fontId", "fillId", "borderId", "numFmtId", "protectionId",
        "alignmentId", "pivotButton", "quotePrefix", "xfId",
    )

    def __new__(cls, values: tuple[int, ...] = (0,) * 9) -> StyleArray:
        return super().__new__(cls, values)

    @property
    def fontId(self) -> int:
        """The font id."""
        return self[0]

    @property
    def fillId(self) -> int:
        """The fill id."""
        return self[1]

    @property
    def borderId(self) -> int:
        """The border id."""
        return self[2]

    @property
    def numFmtId(self) -> int:
        """The number-format id."""
        return self[3]

    @property
    def protectionId(self) -> int:
        """The protection id."""
        return self[4]

    @property
    def alignmentId(self) -> int:
        """The alignment id."""
        return self[5]

    @property
    def pivotButton(self) -> int:
        """The pivot-button flag."""
        return self[6]

    @property
    def quotePrefix(self) -> int:
        """The quote-prefix flag."""
        return self[7]

    @property
    def xfId(self) -> int:
        """The style xf id."""
        return self[8]


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

    def as_tuple(self) -> StyleArray:
        """The style ids openpyxl keeps in its style array.

        openpyxl returns a nine-id ``StyleArray`` in the order ``fontId,
        fillId, borderId, numFmtId, protectionId, alignmentId, pivotButton,
        quotePrefix, xfId``. xlspy stores the components directly and assigns
        ids only while saving, so the array is all zeros, as openpyxl's is for
        a style not yet bound to a workbook. Read ``font``, ``fill`` and the
        other components for their values.
        """
        return StyleArray()

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

    def keys(self) -> list[str]:
        """The style names, as a dict-like view over the list."""
        return self.names

    def values(self) -> list[str | NamedStyle]:
        """The styles, as a dict-like view over the list."""
        return list(self)

    def items(self) -> list[tuple[str, str | NamedStyle]]:
        """The names and their styles, as a dict-like view over the list."""
        return [(name, style) for name, style in zip(self.names, self)]

    def get(self, name: str, default: object = None) -> object:
        """The style called ``name``, or ``default`` when there is none."""
        for item in self:
            if (item if isinstance(item, str) else item.name) == name:
                return item
        return default
