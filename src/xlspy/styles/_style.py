"""The Style and DifferentialStyle models."""

from __future__ import annotations
from dataclasses import dataclass
from ._attrs import Alignment, Border, Font, GradientFill, PatternFill, Protection, normalize_color


@dataclass(frozen=True, slots=True)
class Style:
    """A reusable Excel cell style.

    Colors are RGB hex strings (``"008000"`` or ``"#008000"``).  ``None``
    means “leave this property unchanged”, which also makes styles composable
    when a header style and a range style overlap.

    The flat fields (``fill``, ``font_color``, ``bold``, ``italic``,
    ``underline``, ``number_format``, ``horizontal``, ``vertical``, ``wrap``)
    cover the legacy API; the structured fields (``font``, ``pattern``,
    ``border``, ``alignment``, ``protection``) are the openpyxl-like API.
    ``fill`` accepts a color string (solid), ``PatternFill``, or
    ``GradientFill``. On conflict the structured object wins per property; a
    ``None`` sub-property falls back to the corresponding flat shorthand.
    """
    fill: str | PatternFill | GradientFill | None = None
    font_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    number_format: str | None = None
    horizontal: str | None = None
    vertical: str | None = None
    wrap: bool | None = None
    font: Font | None = None
    pattern: PatternFill | None = None
    border: Border | None = None
    alignment: Alignment | None = None
    protection: Protection | None = None

    def __post_init__(self) -> None:
        fill = self.fill
        if fill is not None:
            if type(fill).__name__ == "GradientFill" and not isinstance(
                fill, GradientFill
            ):
                raise TypeError(
                    "fill must be a xlspy GradientFill, PatternFill, or color "
                    "string; openpyxl fills are not interchangeable"
                )
            if isinstance(fill, str):
                object.__setattr__(self, "fill", normalize_color(fill))
            elif not isinstance(fill, (PatternFill, GradientFill)):
                raise TypeError(
                    "fill must be a color string, PatternFill, GradientFill, "
                    "or None"
                )
        if self.font_color is not None:
            object.__setattr__(
                self, "font_color", normalize_color(self.font_color)
            )
        if self.number_format is not None and not isinstance(self.number_format, str):
            raise TypeError("number_format must be a string or None")
        if self.horizontal is not None and self.horizontal not in {
            "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"
        }:
            raise ValueError("invalid horizontal alignment")
        if self.vertical is not None and self.vertical not in {
            "top", "center", "bottom", "justify", "distributed"
        }:
            raise ValueError("invalid vertical alignment")
        for name in ("bold", "italic", "underline", "wrap"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{name} must be bool or None")
        for name, cls in (
            ("font", Font),
            ("pattern", PatternFill),
            ("border", Border),
            ("alignment", Alignment),
            ("protection", Protection),
        ):
            value = getattr(self, name)
            if value is not None and not isinstance(value, cls):
                raise TypeError(f"{name} must be a {cls.__name__} or None")


@dataclass(frozen=True, slots=True)
class DifferentialStyle:
    """The conditional-formatting style (openpyxl ``dxf``).

    Covers font, number format, fill (pattern or gradient), alignment, border,
    and protection — the same six components as openpyxl.
    """

    font: Font | None = None
    fill: PatternFill | GradientFill | None = None
    border: Border | None = None
    number_format: str | None = None
    alignment: Alignment | None = None
    protection: Protection | None = None

    def __post_init__(self) -> None:
        if self.font is not None and not isinstance(self.font, Font):
            raise TypeError("font must be a Font or None")
        if self.fill is not None:
            if type(self.fill).__name__ == "GradientFill" and not isinstance(
                self.fill, GradientFill
            ):
                raise TypeError(
                    "fill must be a xlspy PatternFill or GradientFill"
                )
            if not isinstance(self.fill, (PatternFill, GradientFill)):
                raise TypeError(
                    "fill must be a PatternFill, GradientFill, or None"
                )
        if self.border is not None and not isinstance(self.border, Border):
            raise TypeError("border must be a Border or None")
        if self.number_format is not None and not isinstance(
            self.number_format, str
        ):
            raise TypeError("number_format must be a string or None")
        if self.alignment is not None and not isinstance(
            self.alignment, Alignment
        ):
            raise TypeError("alignment must be an Alignment or None")
        if self.protection is not None and not isinstance(
            self.protection, Protection
        ):
            raise TypeError("protection must be a Protection or None")
