"""Pixel widths of characters in the fonts Excel's AutoFit measures cells with.

Excel measures a cell's text with GDI at 96 DPI (on a display at 100%
scaling): the sum of its characters' pixel advances in the cell's font, no
kerning. ``_font_metrics`` holds those advances for popular Windows fonts in
four styles (``tools/gdi_column_widths.py`` regenerates them). A character
the font lacks is drawn from a fallback font (Segoe UI Emoji, Segoe UI
Symbol, ...), whose widths the module keeps too.

A font without metrics is measured as Calibri widened by ``UNKNOWN_SCALE``:
of the 494 faces of a Windows 11 + Office install, the median is 1.17 times
as wide as Calibri, and 1.4 covers the text fonts without metrics here
(Century, Bookman Old Style, Lucida Sans: up to 1.39, in any style) but not
display faces such as Cooper Black or Elephant (1.5 to 1.6).
"""

from __future__ import annotations

import base64
import bisect
import functools
import json
import math
import zlib

from . import _font_metrics as _metrics

DEFAULT_FAMILY = "calibri"
UNKNOWN = ""  # the family key of a font without metrics
UNKNOWN_SCALE = 1.4
_UNITS = _metrics.UNITS
_SIZES = _metrics.PIXEL_SIZES


# Names of fixed-pitch fonts: like Courier New and Consolas, such a font
# (Lucida Console, Cascadia Mono, ...) gives every character about 0.6 em,
# half again as wide as Calibri, so it is measured as Courier New.
_FIXED_PITCH = ("mono", "console", "courier", "code", "typewriter", "fixed", "terminal")
_FIXED_FAMILY = "courier new"


@functools.lru_cache(maxsize=256)
def family_key(name: str | None) -> str:
    """The metrics a font name is measured with: its lower-case name, a
    fixed-pitch font's stand-in, or ``UNKNOWN``."""
    if not name:
        return DEFAULT_FAMILY
    key = name.strip().lower()
    if key in _metrics.FAMILIES:
        return key
    if any(hint in key for hint in _FIXED_PITCH):
        return _FIXED_FAMILY
    return UNKNOWN


def _unpack(packed):
    return json.loads(zlib.decompress(base64.b85decode(packed)))


@functools.lru_cache(maxsize=None)
def _family_styles(family: str) -> dict:
    return _unpack(_metrics.FAMILIES[family][1])


@functools.lru_cache(maxsize=1)
def _fallback_runs() -> tuple[list[int], list[list[int]]]:
    runs = _unpack(_metrics.FALLBACK)
    return [run[0] for run in runs], runs


def _scaled(units: int, pixels: int) -> int:
    """Design units at a pixel size, rounded half up."""
    return (units * pixels * 2 + _UNITS) // (2 * _UNITS)


def _scaled_up(units: int, pixels: int) -> int:
    """Design units at a pixel size, rounded up: the fallback fonts' hinting is
    not stored, and may add a pixel."""
    return -(-units * pixels // _UNITS)


def fallback_pixels(char: str, pixels: int) -> int | None:
    """Width of a character Excel draws from a fallback font (None: unknown)."""
    starts, runs = _fallback_runs()
    index = bisect.bisect_right(starts, ord(char)) - 1
    if index >= 0:
        first, count, units = runs[index]
        if ord(char) < first + count:
            return _scaled_up(units, pixels)
    return None


@functools.lru_cache(maxsize=128)
def advances(family: str, bold: bool, italic: bool, pixels: int) -> dict[str, int]:
    """Pixel advance of each character GDI draws in a font ``pixels`` high,
    with the fallback font's characters the font lacks."""
    result: dict[str, int] = {}
    starts, runs = _fallback_runs()
    for first, count, units in runs:
        if first > 0xFFFF:
            break
        width = _scaled_up(units, pixels)
        for cp in range(first, first + count):
            result[chr(cp)] = width
    if family == UNKNOWN:
        # Many fonts have no bold face: Windows emboldens them, a pixel or two
        # more for every character (Century bold: "iiilll" outgrows the scale).
        embolden = (pixels + 23) // 24 if bold else 0
        for char, design in zip(_metrics.CHARS, _family_styles(DEFAULT_FAMILY)[f"{int(bold)}{int(italic)}"]["design"]):
            if design >= 0:
                result[char] = _scaled(math.ceil(design * UNKNOWN_SCALE), pixels) + embolden
    else:
        style = _family_styles(family)[f"{int(bold)}{int(italic)}"]
        shift, moved = 0, {}
        hinting = style["hinting"].get(str(pixels))
        if hinting is not None:
            shift, moved = hinting
        for char, design in zip(_metrics.CHARS, style["design"]):
            if design >= 0:
                result[char] = _scaled(design, pixels) + shift
        for delta, chars in moved.items():
            delta = int(delta)
            for char in chars:
                result[char] += delta
    # GDI draws a tab as half the em (measured at 8 to 20 pt), anywhere in the text.
    result["\t"] = (pixels + 1) // 2
    # Excel draws a soft hyphen as a hyphen.
    result["\xad"] = result["-"]
    return result


def ascent(family: str, bold: bool, italic: bool, pixels: int) -> int:
    """The font's ascent in pixels (GDI's tmAscent)."""
    if family == UNKNOWN:
        return math.ceil(ascent(DEFAULT_FAMILY, bold, italic, pixels) * UNKNOWN_SCALE)
    style = _family_styles(family)[f"{int(bold)}{int(italic)}"]
    if pixels in _SIZES:
        return style["ascent"][pixels - _SIZES.start]
    return _scaled(style["ascent_units"], pixels)


def margin(advances: dict[str, int]) -> int:
    """The room AutoFit adds to a text: two pixels for every 4 pixels of the
    font's digit 0, less two (Calibri 11: 2). Measured on 14 fonts, 4 styles,
    8 to 40 pixels."""
    return 2 * math.ceil(advances["0"] / 4) - 2


def italic_pixels(family: str, bold: bool, italic: bool, pixels: int) -> int:
    """The room AutoFit adds for italic text's slant, by the font's ascent
    (Calibri 11: 4 pixels)."""
    if not italic:
        return 0
    return 2 * math.ceil((ascent(family, bold, italic, pixels) - 4) / 5)
