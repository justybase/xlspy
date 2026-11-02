"""Cell style model, re-exported from :mod:`xlspy.styles.core`.

The package layout keeps the historical ``from xlspy.styles import Style``
imports working while adding openpyxl-compatible submodules such as
``xlspy.styles.numbers``.
"""
from __future__ import annotations

from .core import *  # noqa: F401,F403
from .named_styles import NamedStyle as NamedStyle, NamedStyleList as NamedStyleList
from .core import (
    _EMPTY_STYLES as _EMPTY_STYLES,
    _FontKey as _FontKey,
    _alignment_key_from as _alignment_key_from,
    _border_key as _border_key,
    _coerce_color as _coerce_color,
    _color_key as _color_key,
    _column_index as _column_index,
    _date_cell_xf as _date_cell_xf,
    _fill_key_from as _fill_key_from,
    _font_is_empty as _font_is_empty,
    _font_key_from as _font_key_from,
    _pattern_fill_key as _pattern_fill_key,
    _protection_key_from as _protection_key_from,
    _record_cell_xf as _record_cell_xf,
    _style_ids_for_row as _style_ids_for_row,
    _xf_for_style_format as _xf_for_style_format,
    fill_key_for as fill_key_for,
)

__all__ = [name for name in globals() if not name.startswith("__")]
