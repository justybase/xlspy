"""Unit conversions used by DrawingML, compatible with openpyxl."""

from __future__ import annotations

import math


def EMU_to_cm(value):
    return round(value / 360000, 4)


def cm_to_EMU(value):
    return int(value * 360000)


def EMU_to_inch(value):
    return round(value / 914400, 4)


def inch_to_EMU(value):
    return int(value * 914400)


def EMU_to_pixels(value):
    return round(value / 9525)


def pixels_to_EMU(value):
    return int(value * 9525)


def pixels_to_points(value, dpi=96):
    return value * 72 / dpi


def points_to_pixels(value, dpi=96):
    return int(math.ceil(value * dpi / 72))


def dxa_to_cm(value):
    return 2.54 * dxa_to_inch(value)


def cm_to_dxa(value):
    emu = cm_to_EMU(value)
    inch = EMU_to_inch(emu)
    return inch_to_dxa(inch)


def dxa_to_inch(value):
    return value / 72 / 20


def inch_to_dxa(value):
    return int(value * 20 * 72)


def short_color(color: str) -> str:
    if len(color) > 6:
        return color[2:]
    return color
