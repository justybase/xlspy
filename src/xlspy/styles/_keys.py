"""Hashable style keys: the canonical form used to intern fonts, fills, borders and xfs."""

from __future__ import annotations
from typing import NamedTuple, cast
from ._attrs import Alignment, Border, Color, Font, GradientFill, PatternFill, Protection, Side, Stop, _coerce_color
from ._style import Style


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
        _underline_key(font.underline),
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
    # GradientFill.__post_init__ has already turned every stop into a Stop.
    stops = tuple(
        (_color_key(stop.color), stop.position)
        for stop in cast(tuple[Stop, ...], fill.stop)
    )
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


def fill_key_for(fill):
    """The fill component key of a Style fill: a PatternFill, a GradientFill, or a
    colour string, which is a solid fill of that colour (``Style.fill``)."""
    if isinstance(fill, PatternFill):
        return _pattern_fill_key(fill)
    if isinstance(fill, GradientFill):
        return _fill_key_from(fill)
    if fill is not None:
        return ("pattern", "solid", ("rgb", f"FF{fill}"), None)
    return None


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
    font_key: _FontKey | None
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
            _underline_key(font.underline) if font.underline is not None else flat_underline,
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
    if fill_key is None:
        fill_key = fill_key_for(style.fill)

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
