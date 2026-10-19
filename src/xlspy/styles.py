"""Cell formatting shared by the XLSX and XLSB writers.

The model mirrors openpyxl: ``Font``, ``PatternFill``/``GradientFill``,
``Border``/``Side``, ``Alignment``, ``Protection``, and the conditional rules
(``CellIsRule``, ``FormulaRule``, ``ColorScaleRule``, ``DataBarRule``,
``IconSetRule``) with ``DifferentialStyle`` as the Excel ``dxf``. The XLSX
writer serializes all of it; the XLSB writer currently supports the flat
``Style`` subset and rejects the rest with ``NotImplementedError``.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from functools import lru_cache
import re
from typing import Mapping, NamedTuple

# Chart and cell-reference models live in their own modules; they are
# re-exported here so that ``xlspy.styles`` keeps its historical names.
from .charts import Chart, ChartSeries  # noqa: F401
from .hyperlinks import Hyperlink
from .refs import _column_index, column_label, parse_cell_range  # noqa: F401

_EMPTY_STYLES = {}


# ---------------------------------------------------------------------------
# Structured cell-style model (openpyxl-like, XLSB-safe pure dataclasses)
# ---------------------------------------------------------------------------

_ARGB_RE = re.compile(r"[0-9A-F]{8}")
_TABLE_NAME_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
# Table names Excel refuses because they read as cell references.
_A1_NAME_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+)")
_R1C1_NAME_RE = re.compile(r"[Rr][0-9]*[Cc][0-9]*|[Rr][0-9]*|[Cc][0-9]*")


def _is_reference_like(name: str) -> bool:
    match = _A1_NAME_RE.fullmatch(name)
    if match is not None and _column_index(match.group(1).upper()) < 16384 and 1 <= int(match.group(2)) <= 1048576:
        return True
    return _R1C1_NAME_RE.fullmatch(name) is not None

_UNDERLINES = {"none", "single", "double", "singleAccounting", "doubleAccounting"}
_VERT_ALIGNS = {"superscript", "subscript", "baseline"}
_PATTERN_TYPES = {
    "none", "solid", "gray125",
    "darkGray", "mediumGray", "lightGray", "gray0625",
    "darkHorizontal", "darkVertical", "darkDown", "darkUp",
    "darkGrid", "darkTrellis",
    "lightHorizontal", "lightVertical", "lightDown", "lightUp",
    "lightGrid", "lightTrellis",
}
_SIDE_STYLES = {
    "thin", "medium", "thick", "double", "dotted", "dashed",
    "dashDot", "dashDotDot", "slantDashDot", "hair",
    "mediumDashed", "mediumDashDot", "mediumDashDotDot",
}
_DIAGONAL_DIRECTIONS = {"up", "down", "both"}
_HORIZONTAL = {
    "general", "left", "center", "right",
    "fill", "justify", "centerContinuous", "distributed",
}
_VERTICAL = {"top", "center", "bottom", "justify", "distributed"}
_FONT_SCHEMES = {"major", "minor"}
_GRADIENT_TYPES = {"linear", "path"}
_ICON_SETS = {
    "3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2",
    "3Signs", "3Symbols", "3Symbols2",
    "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights",
    "5Arrows", "5ArrowsGray", "5Rating", "5Quarters",
}
_CELL_IS_OPERATORS = {
    "between", "notBetween", "equal", "notEqual",
    "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual",
    "containsText", "notContains", "beginsWith", "endsWith",
}
_FORMAT_OBJECT_TYPES = {"min", "max", "num", "percent", "percentile", "formula"}
_CFVO_TYPES = {"num", "percent", "percentile", "formula"}


def _normalize_argb(value: str) -> str:
    """Validate an ``RRGGBB``/``AARRGGBB`` color into 8-char uppercase ARGB."""
    if not isinstance(value, str):
        raise TypeError("colors must be strings such as '#FF0000'")
    text = value.removeprefix("#").upper()
    if len(text) == 6:
        text = "FF" + text
    if len(text) != 8 or _ARGB_RE.fullmatch(text) is None:
        raise ValueError(
            f"invalid ARGB color {value!r}; expected '#RRGGBB' or '#AARRGGBB'"
        )
    return text


def _coerce_color(value: Color | str | None, field: str) -> Color | None:
    if value is None or isinstance(value, Color):
        return value
    if type(value).__name__ == "GradientFill":
        raise TypeError(
            f"{field} must be a Color or color string; a gradient is only "
            "valid as a fill"
        )
    if isinstance(value, str):
        return Color(rgb=value)
    raise TypeError(f"{field} must be a Color, color string, or None")


@dataclass(frozen=True, slots=True)
class Color:
    """An Excel color: exactly one of ``rgb``/``indexed``/``theme``/``auto``.

    ``rgb`` accepts ``"RRGGBB"`` or ``"AARRGGBB"`` with or without ``"#"`` and
    is normalized to 8-char uppercase ARGB (``FF`` alpha prefix when absent).
    ``auto=True`` requests the automatic (system) color.
    """

    rgb: str | None = None
    indexed: int | None = None
    theme: int | None = None
    tint: float | None = None
    auto: bool | None = None

    def __post_init__(self):
        if self.auto is not None and self.auto is not True:
            raise ValueError("auto must be True or None")
        set_count = sum(
            value is not None
            for value in (self.rgb, self.indexed, self.theme, self.auto)
        )
        if set_count != 1:
            raise ValueError(
                "Color needs exactly one of rgb, indexed, theme, or auto"
            )
        if self.rgb is not None:
            object.__setattr__(self, "rgb", _normalize_argb(self.rgb))
        if self.indexed is not None:
            if not isinstance(self.indexed, int):
                raise TypeError("indexed must be an int or None")
            if not 0 <= self.indexed <= 65:
                raise ValueError("indexed must be 0..65")
        if self.theme is not None:
            if not isinstance(self.theme, int):
                raise TypeError("theme must be an int or None")
            if not 0 <= self.theme <= 9:
                raise ValueError("theme must be 0..9")
        if self.tint is not None:
            if not isinstance(self.tint, (int, float)):
                raise TypeError("tint must be a float or None")
            if self.theme is None:
                raise ValueError("tint needs a theme color")


@dataclass(frozen=True, slots=True)
class Font:
    """An openpyxl-like font. All-``None`` means “no font override”."""

    name: str | None = None
    size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    strike: bool | None = None
    underline: str | bool | None = None
    vertAlign: str | None = None
    color: Color | str | None = None
    charset: int | None = None
    family: int | None = None
    scheme: str | None = None
    outline: bool | None = None
    shadow: bool | None = None
    condense: bool | None = None
    extend: bool | None = None

    def __post_init__(self):
        if self.name is not None and not isinstance(self.name, str):
            raise TypeError("name must be a string or None")
        if self.size is not None:
            if not isinstance(self.size, (int, float)):
                raise TypeError("size must be a number or None")
            if self.size <= 0:
                raise ValueError("size must be > 0")
        for attr in ("bold", "italic", "strike"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")
        underline = self.underline
        if underline is True:
            object.__setattr__(self, "underline", "single")
        elif underline is False:
            object.__setattr__(self, "underline", "none")
        elif underline is not None and underline not in _UNDERLINES:
            raise ValueError(f"invalid underline {underline!r}")
        if self.vertAlign is not None and self.vertAlign not in _VERT_ALIGNS:
            raise ValueError(f"invalid vertAlign {self.vertAlign!r}")
        object.__setattr__(self, "color", _coerce_color(self.color, "color"))
        if self.charset is not None:
            if not isinstance(self.charset, int) or isinstance(self.charset, bool):
                raise TypeError("charset must be an int or None")
            if not 0 <= self.charset <= 255:
                raise ValueError("charset must be 0..255")
        if self.family is not None:
            if not isinstance(self.family, int) or isinstance(self.family, bool):
                raise TypeError("family must be an int or None")
            if not 0 <= self.family <= 14:
                raise ValueError("family must be 0..14")
        if self.scheme is not None and self.scheme not in _FONT_SCHEMES:
            raise ValueError(f"invalid scheme {self.scheme!r}")
        for attr in ("outline", "shadow", "condense", "extend"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")


@dataclass(frozen=True, slots=True)
class PatternFill:
    """An openpyxl-like pattern fill. All-``None`` means “no fill override”."""

    fill_type: str | None = None
    start_color: Color | str | None = None
    end_color: Color | str | None = None

    def __post_init__(self):
        if self.fill_type is not None and self.fill_type not in _PATTERN_TYPES:
            raise ValueError(f"invalid fill_type {self.fill_type!r}")
        object.__setattr__(
            self, "start_color", _coerce_color(self.start_color, "start_color")
        )
        object.__setattr__(
            self, "end_color", _coerce_color(self.end_color, "end_color")
        )


@dataclass(frozen=True, slots=True)
class Stop:
    """One gradient stop: a color at a position between 0 and 1."""

    color: Color | str
    position: float

    def __post_init__(self):
        color = _coerce_color(self.color, "color")
        if color is None:
            raise TypeError("stop color is required")
        object.__setattr__(self, "color", color)
        if isinstance(self.position, bool) or not isinstance(
            self.position, (int, float)
        ):
            raise TypeError("position must be a number")
        if not 0 <= self.position <= 1:
            raise ValueError("position must be between 0 and 1")


@dataclass(frozen=True, slots=True)
class GradientFill:
    """An openpyxl-like gradient fill (``linear`` or ``path``).

    ``stop`` accepts ``Stop`` objects or plain colors; plain colors are spread
    evenly between position 0 and 1, matching openpyxl.
    """

    fill_type: str = "linear"
    degree: float | None = None
    left: float | None = None
    right: float | None = None
    top: float | None = None
    bottom: float | None = None
    stop: tuple[Stop | Color | str, ...] = ()

    def __post_init__(self):
        if self.fill_type not in _GRADIENT_TYPES:
            raise ValueError(f"invalid gradient type {self.fill_type!r}")
        degree = self.degree
        if degree is not None:
            if isinstance(degree, bool) or not isinstance(degree, (int, float)):
                raise TypeError("degree must be a number or None")
            if not 0 <= degree <= 360:
                raise ValueError("degree must be between 0 and 360")
        for name in ("left", "right", "top", "bottom"):
            value = getattr(self, name)
            if value is not None:
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    raise TypeError(f"{name} must be a number or None")
                if not 0 <= value <= 1:
                    raise ValueError(f"{name} must be between 0 and 1")
        if self.fill_type == "path" and self.degree is not None:
            raise ValueError("path gradients have no degree; use linear")
        if not isinstance(self.stop, (list, tuple)):
            raise TypeError("stop must be a list of Stop objects or colors")
        if not self.stop:
            raise ValueError("a gradient needs at least one stop")
        stops = []
        last = len(self.stop) - 1
        for index, item in enumerate(self.stop):
            if isinstance(item, Stop):
                stops.append(item)
            else:
                position = index / last if last else 0.0
                stops.append(Stop(item, position))
        object.__setattr__(self, "stop", tuple(stops))


@dataclass(frozen=True, slots=True)
class Side:
    """One border edge."""

    style: str | None = None
    color: Color | str | None = None

    def __post_init__(self):
        if self.style is not None and self.style not in _SIDE_STYLES:
            raise ValueError(f"invalid border style {self.style!r}")
        object.__setattr__(self, "color", _coerce_color(self.color, "color"))


@dataclass(frozen=True, slots=True)
class Border:
    """An openpyxl-like border. All-``None`` means “no border override”."""

    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    diagonal: Side | None = None
    diagonal_direction: str | None = None
    vertical: Side | None = None
    horizontal: Side | None = None
    start: Side | None = None
    end: Side | None = None
    diagonal_up: bool | None = None
    diagonal_down: bool | None = None
    outline: bool | None = None

    def __post_init__(self):
        for attr in (
            "left", "right", "top", "bottom",
            "diagonal", "vertical", "horizontal", "start", "end",
        ):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, Side):
                raise TypeError(f"{attr} must be a Side or None")
        if (
            self.diagonal_direction is not None
            and self.diagonal_direction not in _DIAGONAL_DIRECTIONS
        ):
            raise ValueError(
                f"invalid diagonal_direction {self.diagonal_direction!r}"
            )
        for attr in ("diagonal_up", "diagonal_down", "outline"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")


@dataclass(frozen=True, slots=True)
class Alignment:
    """An openpyxl-like alignment. All-``None`` means “no override”."""

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool | None = None
    shrink_to_fit: bool | None = None
    indent: int | None = None
    text_rotation: int | None = None
    relative_indent: int | None = None
    justify_last_line: bool | None = None
    reading_order: int | None = None

    def __post_init__(self):
        if self.horizontal is not None and self.horizontal not in _HORIZONTAL:
            raise ValueError(f"invalid horizontal alignment {self.horizontal!r}")
        if self.vertical is not None and self.vertical not in _VERTICAL:
            raise ValueError(f"invalid vertical alignment {self.vertical!r}")
        for attr in ("wrap_text", "shrink_to_fit", "justify_last_line"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")
        for attr in ("indent", "relative_indent"):
            value = getattr(self, attr)
            if value is not None:
                if not isinstance(value, int) or isinstance(value, bool):
                    raise TypeError(f"{attr} must be an int or None")
                if not 0 <= value <= 250:
                    raise ValueError(f"{attr} must be 0..250")
        if self.text_rotation is not None:
            if not isinstance(self.text_rotation, int):
                raise TypeError("text_rotation must be an int or None")
            if not (0 <= self.text_rotation <= 180
                    or self.text_rotation == 255):
                raise ValueError(
                    "text_rotation must be 0..180 or 255 (vertical text)"
                )
        if self.reading_order is not None:
            if not isinstance(self.reading_order, int) or isinstance(
                self.reading_order, bool
            ):
                raise TypeError("reading_order must be an int or None")
            if self.reading_order not in (0, 1, 2):
                raise ValueError("reading_order must be 0, 1, or 2")


@dataclass(frozen=True, slots=True)
class Protection:
    """Cell protection flags. All-``None`` means “no override”."""

    locked: bool | None = None
    hidden: bool | None = None

    def __post_init__(self):
        for attr in ("locked", "hidden"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")


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
    fill: str | PatternFill | None = None
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

    def __post_init__(self):
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


# ---------------------------------------------------------------------------
# Sheet-feature models: conditional formatting, tables
# ---------------------------------------------------------------------------

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

    def __post_init__(self):
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


def _normalize_formula(formula: str | list[str] | tuple[str, ...]) -> tuple[str, ...]:
    if isinstance(formula, str):
        return (formula,)
    if isinstance(formula, (list, tuple)):
        items = tuple(formula)
        for item in items:
            if not isinstance(item, str):
                raise TypeError("formula entries must be strings")
        return items
    raise TypeError("formula must be a string or a list of strings")


def _validate_priority(priority: int | None) -> None:
    if priority is None:
        return
    if not isinstance(priority, int) or isinstance(priority, bool):
        raise TypeError("priority must be a positive int or None")
    if priority <= 0:
        raise ValueError("priority must be a positive int")


def _validate_sqref(sqref: str) -> None:
    # Multi-range "A1:A5 C1:C5" is not supported: the split below yields
    # three parts and parse_cell_range raises ValueError.
    parse_cell_range(sqref)


@dataclass(frozen=True, slots=True)
class CellIsRule:
    """A ``cellIs`` conditional-formatting rule."""

    sqref: str
    operator: str
    formula: list[str] | tuple[str, ...] | str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        if self.operator not in _CELL_IS_OPERATORS:
            raise ValueError(f"invalid operator {self.operator!r}")
        object.__setattr__(self, "formula", _normalize_formula(self.formula))
        if self.dxf is not None and not isinstance(self.dxf, DifferentialStyle):
            raise TypeError("dxf must be a DifferentialStyle or None")
        if self.stop_if_true is not None and not isinstance(self.stop_if_true, bool):
            raise TypeError("stop_if_true must be bool or None")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class FormulaRule:
    """An ``expression`` conditional-formatting rule."""

    sqref: str
    formula: list[str] | tuple[str, ...] | str
    dxf: DifferentialStyle | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        object.__setattr__(self, "formula", _normalize_formula(self.formula))
        if self.dxf is not None and not isinstance(self.dxf, DifferentialStyle):
            raise TypeError("dxf must be a DifferentialStyle or None")
        if self.stop_if_true is not None and not isinstance(self.stop_if_true, bool):
            raise TypeError("stop_if_true must be bool or None")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class FormatObject:
    """A ``cfvo`` data-bar / color-scale endpoint."""

    type: str
    value: str | float | None = None

    def __post_init__(self):
        if self.type not in _FORMAT_OBJECT_TYPES:
            raise ValueError(f"invalid format object type {self.type!r}")
        if (
            self.value is not None
            and not isinstance(self.value, (str, int, float))
        ):
            raise TypeError("value must be a string, number, or None")


@dataclass(frozen=True, slots=True)
class ColorScaleRule:
    """A 2-color (``midpoint`` None) or 3-color scale rule."""

    sqref: str
    minimum: FormatObject | None
    midpoint: FormatObject | None
    maximum: FormatObject | None
    colors: list[Color | str] | tuple[Color | str, ...]
    priority: int | None = None
    def __post_init__(self):
        _validate_sqref(self.sqref)
        for name in ("minimum", "midpoint", "maximum"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, FormatObject):
                raise TypeError(f"{name} must be a FormatObject or None")
        if not isinstance(self.colors, (list, tuple)):
            raise TypeError("colors must be a list of Color or color strings")
        colors = tuple(_coerce_color(item, "colors") for item in self.colors)
        expected = 2 if self.midpoint is None else 3
        if len(colors) != expected:
            raise ValueError(
                f"expected {expected} colors for a "
                f"{'2' if expected == 2 else '3'}-color scale"
            )
        object.__setattr__(self, "colors", colors)
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class DataBarRule:
    """A ``dataBar`` conditional-formatting rule."""

    sqref: str
    minimum: FormatObject | None
    maximum: FormatObject | None
    color: Color | str
    show_value: bool | None = None
    min_length: int | None = None
    max_length: int | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        for name in ("minimum", "maximum"):
            value = getattr(self, name)
            if value is not None and not isinstance(value, FormatObject):
                raise TypeError(f"{name} must be a FormatObject or None")
        object.__setattr__(self, "color", _coerce_color(self.color, "color"))
        if self.color is None:
            raise ValueError("color is required")
        if self.show_value is not None and not isinstance(self.show_value, bool):
            raise TypeError("show_value must be bool or None")
        for name in ("min_length", "max_length"):
            value = getattr(self, name)
            if value is not None:
                if not isinstance(value, int) or isinstance(value, bool):
                    raise TypeError(f"{name} must be an int or None")
                if value < 0:
                    raise ValueError(f"{name} must be >= 0")
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class IconSetRule:
    """An ``iconSet`` conditional-formatting rule (openpyxl-like)."""

    sqref: str
    icon_style: str
    values: list[float | str] | tuple[float | str, ...]
    type: str = "percent"
    show_value: bool | None = None
    percent: bool | None = None
    reverse: bool | None = None
    stop_if_true: bool | None = None
    priority: int | None = None

    def __post_init__(self):
        _validate_sqref(self.sqref)
        if self.icon_style not in _ICON_SETS:
            raise ValueError(f"invalid icon set {self.icon_style!r}")
        if not isinstance(self.values, (list, tuple)):
            raise TypeError("values must be a list of numbers")
        values = tuple(self.values)
        if not values:
            raise ValueError("an icon set rule needs at least one value")
        for value in values:
            if isinstance(value, bool) or not isinstance(value, (int, float, str)):
                raise TypeError("icon set values must be numbers or formulas")
        if self.type not in _CFVO_TYPES:
            raise ValueError(f"invalid icon set value type {self.type!r}")
        if self.type != "formula":
            for value in values:
                if isinstance(value, str):
                    raise ValueError(
                        "string values require type='formula'"
                    )
        elif any(not isinstance(value, str) for value in values):
            raise ValueError("type='formula' requires string values")
        for attr in ("show_value", "percent", "reverse", "stop_if_true"):
            value = getattr(self, attr)
            if value is not None and not isinstance(value, bool):
                raise TypeError(f"{attr} must be bool or None")
        object.__setattr__(self, "values", values)
        _validate_priority(self.priority)


@dataclass(frozen=True, slots=True)
class ExcelTable:
    """An Excel table (ListObject) bound to a rectangular range."""

    display_name: str
    ref: str
    columns: list[str] | tuple[str, ...] | None = None
    header_row: bool = True
    totals_row: bool = False
    style_name: str = "TableStyleMedium9"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False

    def __post_init__(self):
        if not isinstance(self.display_name, str) or _TABLE_NAME_RE.fullmatch(
            self.display_name
        ) is None:
            raise ValueError(
                f"invalid table display name {self.display_name!r}"
            )
        if _is_reference_like(self.display_name):
            raise ValueError(
                f"invalid table display name {self.display_name!r}: "
                "it reads as a cell reference, which Excel refuses"
            )
        first_col, _, last_col, _ = parse_cell_range(self.ref)
        if self.columns is not None:
            if not isinstance(self.columns, (list, tuple)):
                raise TypeError("columns must be a list of strings or None")
            columns = tuple(self.columns)
            for column in columns:
                if not isinstance(column, str):
                    raise TypeError("columns must be a list of strings or None")
            if self.header_row and len(columns) != last_col - first_col + 1:
                raise ValueError(
                    "len(columns) must match the table ref width"
                )
            object.__setattr__(self, "columns", columns)
        for name in (
            "header_row", "totals_row",
            "show_first_column", "show_last_column",
            "show_row_stripes", "show_column_stripes",
        ):
            value = getattr(self, name)
            if not isinstance(value, bool):
                raise TypeError(f"{name} must be bool")
        if not isinstance(self.style_name, str):
            raise TypeError("style_name must be a string")


def table_column_names(table: ExcelTable, header_cells) -> list[str]:
    """Column names for ``table``, checked against the cells of its header row.

    Excel opens a table only when each header cell holds text equal (ignoring
    case) to its column's name and the names are unique; Excel 16.0 refuses
    the whole file otherwise. Names come from ``table.columns`` or, when that
    is not given, from the header cells. ``header_cells`` is the sheet row
    holding the header (``None`` if the data never reached it).
    """
    first_col, first_row, last_col, _ = parse_cell_range(table.ref)
    width = last_col - first_col + 1
    if not table.header_row:
        if table.columns is not None:
            return list(table.columns)
        return [f"Column{i + 1}" for i in range(width)]
    row = header_cells or ()
    names = []
    for offset in range(width):
        col = first_col + offset
        ref = f"{column_label(col)}{first_row + 1}"
        cell = row[col] if col < len(row) else None
        if isinstance(cell, tuple) and len(cell) == 2 and isinstance(cell[1], str):
            cell = cell[0]  # a (value, number format) cell
        if isinstance(cell, Hyperlink):
            cell = cell.display_text  # written to the cell as plain text
        if not isinstance(cell, str) or cell == "":
            raise ValueError(
                f"table {table.display_name!r}: header cell {ref} must hold text "
                f"(found {cell!r}); Excel requires a text header for every table column"
            )
        if table.columns is not None:
            name = table.columns[offset]
            if name.casefold() != cell.casefold():
                raise ValueError(
                    f"table {table.display_name!r}: column name {name!r} does not match "
                    f"header cell {ref} ({cell!r}); Excel requires them to be equal"
                )
        else:
            name = cell
        names.append(name)
    seen = set()
    for name in names:
        if name.casefold() in seen:
            raise ValueError(
                f"table {table.display_name!r}: duplicate column name {name!r}; "
                "Excel requires unique table column names"
            )
        seen.add(name.casefold())
    return names


def capture_table_headers(rows, tables, store: dict):
    """Yield ``rows`` unchanged, checking each table's header row as it passes.

    A copy of every header row is kept in ``store`` (row index -> cells) for
    ``table_column_names``. Checking while the row streams past means a bad
    header fails before the rest of the sheet is encoded.
    """
    by_row: dict[int, list] = {}
    for table in tables:
        if table.header_row:
            by_row.setdefault(parse_cell_range(table.ref)[1], []).append(table)
    if not by_row:
        yield from rows
        return
    for row_idx, row in enumerate(rows):
        header_tables = by_row.get(row_idx)
        if header_tables is not None:
            cells = list(row)
            store[row_idx] = cells
            for table in header_tables:
                table_column_names(table, cells)
        yield row


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


# ---------------------------------------------------------------------------
# Registry-key splitting (used only by the XLSX writer)
# ---------------------------------------------------------------------------

class _FontKey(NamedTuple):
    name: str | None
    size: float | None
    bold: bool | None
    italic: bool | None
    strike: bool | None
    underline: str | None
    vert_align: str | None
    color: tuple | None
    charset: int | None
    family: int | None
    scheme: str | None
    outline: bool | None
    shadow: bool | None
    condense: bool | None
    extend: bool | None


class _AlignmentKey(NamedTuple):
    horizontal: str | None
    vertical: str | None
    wrap_text: bool | None
    shrink_to_fit: bool | None
    indent: int | None
    text_rotation: int | None
    relative_indent: int | None
    justify_last_line: bool | None
    reading_order: int | None


class _BorderKey(NamedTuple):
    left: tuple
    right: tuple
    top: tuple
    bottom: tuple
    diagonal: tuple
    vertical: tuple
    horizontal: tuple
    start: tuple
    end: tuple
    diagonal_up: bool
    diagonal_down: bool
    outline: bool | None


class _ProtectionKey(NamedTuple):
    locked: bool | None
    hidden: bool | None


def _underline_key(value: str | bool | None) -> str | None:
    if value is True:
        return "single"
    if value is False:
        return "none"
    return value


def _color_key(color: Color | str | None):
    """Hashable registry key for a color: ``("rgb", ARGB)``/``indexed``/``theme``/``auto``."""
    color = _coerce_color(color, "color")
    if color is None:
        return None
    if color.rgb is not None:
        return ("rgb", color.rgb)
    if color.indexed is not None:
        return ("indexed", color.indexed)
    if color.auto:
        return ("auto",)
    return ("theme", color.theme, color.tint)


def _side_key(side: Side | None):
    if side is None:
        return (None, None)
    return (side.style, _color_key(side.color))


_EMPTY_SIDE = (None, None)


def _font_key_from(font: Font) -> _FontKey:
    return _FontKey(
        font.name,
        font.size,
        font.bold,
        font.italic,
        font.strike,
        font.underline,
        font.vertAlign,
        _color_key(font.color),
        font.charset,
        font.family,
        font.scheme,
        font.outline,
        font.shadow,
        font.condense,
        font.extend,
    )


def _font_is_empty(font_key: _FontKey | None) -> bool:
    return font_key is None or all(value is None for value in font_key)


def _border_key(border: Border) -> _BorderKey | None:
    """Hashable border key; ``None`` when the border overrides nothing."""
    direction = border.diagonal_direction
    diagonal_up = border.diagonal_up
    diagonal_down = border.diagonal_down
    if diagonal_up is None:
        diagonal_up = direction in ("up", "both")
    if diagonal_down is None:
        diagonal_down = direction in ("down", "both")
    key = _BorderKey(
        _side_key(border.left),
        _side_key(border.right),
        _side_key(border.top),
        _side_key(border.bottom),
        _side_key(border.diagonal),
        _side_key(border.vertical),
        _side_key(border.horizontal),
        _side_key(border.start),
        _side_key(border.end),
        diagonal_up,
        diagonal_down,
        border.outline,
    )
    if (
        all(part == _EMPTY_SIDE for part in key[:9])
        and not key.diagonal_up
        and not key.diagonal_down
        and key.outline in (None, True)
    ):
        return None
    return key


def _pattern_fill_key(fill: PatternFill):
    if (
        fill.fill_type is None
        and fill.start_color is None
        and fill.end_color is None
    ):
        return None
    return (
        "pattern",
        fill.fill_type,
        _color_key(fill.start_color),
        _color_key(fill.end_color),
    )


def _fill_key_from(fill: PatternFill | GradientFill):
    """Key for a pattern or gradient fill; ``None`` means “no fill override”."""
    if isinstance(fill, PatternFill):
        return _pattern_fill_key(fill)
    stops = tuple((_color_key(stop.color), stop.position) for stop in fill.stop)
    return (
        "gradient",
        fill.fill_type,
        fill.degree,
        fill.left,
        fill.right,
        fill.top,
        fill.bottom,
        stops,
    )


def _alignment_key_from(alignment: Alignment) -> _AlignmentKey:
    return _AlignmentKey(
        alignment.horizontal,
        alignment.vertical,
        alignment.wrap_text,
        alignment.shrink_to_fit,
        alignment.indent,
        alignment.text_rotation,
        alignment.relative_indent,
        alignment.justify_last_line,
        alignment.reading_order,
    )


def _protection_key_from(protection: Protection) -> _ProtectionKey:
    return _ProtectionKey(protection.locked, protection.hidden)


def split_style(style: Style) -> tuple:
    """Split a ``Style`` into hashable ``(font, fill, border, numfmt, align, prot)`` keys.

    Flat shorthands (``font_color``/``bold``/``italic``/``underline``,
    ``horizontal``/``vertical``/``wrap``) fill in any sub-property the
    structured object leaves ``None``; the structured object wins on conflict.
    ``fill`` given as ``PatternFill`` behaves like ``pattern``; ``pattern``
    wins when both are set, then a gradient fill, then the flat color. A font
    or border that overrides nothing yields ``None``, which the writer maps to
    id 0 without a registry entry.
    """
    flat_underline = _underline_key(style.underline)
    font = style.font
    if font is None:
        font_key = _FontKey(
            None, None, style.bold, style.italic, None, flat_underline, None,
            _color_key(style.font_color),
            None, None, None, None, None, None, None,
        )
        if _font_is_empty(font_key):
            font_key = None
    else:
        font_key = _FontKey(
            font.name,
            font.size,
            font.bold if font.bold is not None else style.bold,
            font.italic if font.italic is not None else style.italic,
            font.strike,
            font.underline if font.underline is not None else flat_underline,
            font.vertAlign,
            _color_key(
                font.color if font.color is not None else style.font_color
            ),
            font.charset,
            font.family,
            font.scheme,
            font.outline,
            font.shadow,
            font.condense,
            font.extend,
        )

    fill_key = None
    if style.pattern is not None:
        fill_key = _pattern_fill_key(style.pattern)
    if fill_key is None and isinstance(style.fill, PatternFill):
        fill_key = _pattern_fill_key(style.fill)
    if fill_key is None and isinstance(style.fill, GradientFill):
        fill_key = _fill_key_from(style.fill)
    if fill_key is None and style.fill is not None:
        fill_key = ("pattern", "solid", ("rgb", f"FF{style.fill}"), None)

    border_key = None
    if style.border is not None:
        border_key = _border_key(style.border)

    alignment = style.alignment
    if alignment is None:
        align_key = _AlignmentKey(
            style.horizontal, style.vertical, style.wrap,
            None, None, None, None, None, None,
        )
    else:
        align_key = _AlignmentKey(
            alignment.horizontal
            if alignment.horizontal is not None else style.horizontal,
            alignment.vertical
            if alignment.vertical is not None else style.vertical,
            alignment.wrap_text
            if alignment.wrap_text is not None else style.wrap,
            alignment.shrink_to_fit,
            alignment.indent,
            alignment.text_rotation,
            alignment.relative_indent,
            alignment.justify_last_line,
            alignment.reading_order,
        )

    protection = style.protection
    if protection is None:
        prot_key = _ProtectionKey(None, None)
    else:
        prot_key = _protection_key_from(protection)

    return (
        font_key, fill_key, border_key,
        style.number_format, align_key, prot_key,
    )


def normalize_color(value: str) -> str:
    """Validate and normalize a six-digit RGB color."""
    if not isinstance(value, str):
        raise TypeError("colors must be strings such as '#008000'")
    color = value.removeprefix("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        raise ValueError(f"invalid RGB color {value!r}; expected '#RRGGBB'")
    return color


def normalize_style_ranges(styles: Mapping[str, Style] | None):
    if styles is None:
        return ()
    if not isinstance(styles, Mapping):
        raise TypeError("styles must be a mapping of A1 ranges to Style objects")
    result = []
    for reference, style in styles.items():
        if not isinstance(style, Style):
            raise TypeError(f"style for {reference!r} must be a Style instance")
        result.append((*parse_cell_range(reference), style))
    return tuple(result)


def merge_styles(base: Style | None, overlay: Style | None) -> Style | None:
    if base is None:
        return overlay
    if overlay is None:
        return base
    return replace(base, **{
        field: getattr(overlay, field)
        for field in Style.__dataclass_fields__
        if getattr(overlay, field) is not None
    })


@lru_cache(maxsize=128)
def _uniform_row_styles(row_length: int, style: Style):
    # These cached dictionaries are only read by the writers.
    return {column: style for column in range(row_length)}


def styles_for_row(row_index: int, row_length: int, ranges, header_style: Style | None = None):
    """Return sparse column->style mappings; no allocation for unstyled rows."""
    if not ranges and header_style is None:
        return _EMPTY_STYLES
    if row_index == 0 and header_style is not None and not ranges:
        return _uniform_row_styles(row_length, merge_styles(Style(bold=True), header_style))
    if len(ranges) == 1 and header_style is None:
        first_col, first_row, last_col, last_row, style = ranges[0]
        if (first_row <= row_index <= last_row and first_col == 0
                and last_col >= row_length - 1):
            return _uniform_row_styles(row_length, style)
    styles = {}
    if row_index == 0 and header_style is not None:
        effective_header = merge_styles(Style(bold=True), header_style)
        for column in range(row_length):
            styles[column] = effective_header
    for first_col, first_row, last_col, last_row, style in ranges:
        if first_row <= row_index <= last_row:
            for column in range(first_col, min(last_col, row_length - 1) + 1):
                styles[column] = merge_styles(styles.get(column), style)
    return styles or _EMPTY_STYLES
