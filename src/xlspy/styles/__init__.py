"""Cell style model, re-exported from :mod:`xlspy.styles.core`.

The package layout keeps the historical ``from xlspy.styles import Style``
imports working while adding openpyxl-compatible submodules such as
``xlspy.styles.numbers``.
"""

from .core import *  # noqa: F401,F403
from .named_styles import NamedStyle, NamedStyleList  # noqa: F401
from .core import (  # noqa: F401
    _EMPTY_STYLES,
    _FontKey,
    _alignment_key_from,
    _border_key,
    _coerce_color,
    _color_key,
    _column_index,
    _date_cell_xf,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _pattern_fill_key,
    _protection_key_from,
    _record_cell_xf,
    _style_ids_for_row,
    _xf_for_style_format,
)

__all__ = [name for name in globals() if not name.startswith("__")]
