"""Value objects of the cell style model: colours, fonts, fills, borders, alignment, protection."""

from __future__ import annotations
from dataclasses import dataclass
import re
from ..refs import _column_index


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
_VERT_ALIGNS = frozenset({"superscript", "subscript", "baseline"})
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
            if not 0 <= self.theme <= 11:
                raise ValueError("theme must be 0..11")
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


def normalize_color(value: str) -> str:
    """Validate and normalize a six-digit RGB color."""
    if not isinstance(value, str):
        raise TypeError("colors must be strings such as '#008000'")
    color = value.removeprefix("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        raise ValueError(f"invalid RGB color {value!r}; expected '#RRGGBB'")
    return color
