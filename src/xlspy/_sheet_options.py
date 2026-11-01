"""Sheet-level options for ``write_sheet``, bundled in one object.

``write_sheet(rows, options=SheetOptions(zoom=80, freeze_panes="A2"))`` and
``write_sheet(rows, zoom=80, freeze_panes="A2")`` are the same call. The
keyword form stays for short calls; a SheetOptions object is for options that
are built once and reused for several sheets.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, fields, replace
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ._metadata import SheetProtection
    from .printing import PrintSettings
    from .styles import Color, Style


@dataclass(frozen=True)
class SheetOptions:
    """Every option of ``write_sheet`` except the rows. Defaults match the call defaults."""

    header_style: Style | None = None
    styles: dict[str, Style] | None = None
    merges: list[str] | tuple[str, ...] | None = None
    conditional_formats: list | tuple | None = None
    tables: list | tuple | None = None
    column_widths: dict | None = None
    hidden_columns: list | tuple | set | None = None
    row_heights: dict | None = None
    hidden_rows: list | tuple | set | None = None
    row_groups: list | tuple | None = None
    column_groups: list | tuple | None = None
    freeze_panes: str | None = None
    auto_filter: bool | None = None
    auto_filter_ref: str | None = None
    charts: list | tuple | None = None
    data_validations: list | tuple | None = None
    images: list | tuple | None = None
    print_settings: PrintSettings | None = None
    comments: Mapping[str, Any] | None = None
    sheet_protection: SheetProtection | None = None
    show_grid_lines: bool = True
    header_rows: int = 1
    tab_color: Color | str | None = None
    zoom: int | None = None
    ignored_errors: list | tuple | None = None
    sparklines: list | tuple | None = None
    slicers: list | tuple | None = None
    pivots: list | tuple | None = None
    shapes: list | tuple | None = None
    show_row_col_headers: bool = True
    show_zero_values: bool = True
    right_to_left: bool = False

    def as_kwargs(self, names: frozenset[str] | None = None) -> dict[str, Any]:
        """The fields in ``names`` (all of them by default) as keyword arguments.
        Values are not copied."""
        return {name: getattr(self, name) for name in _ORDER if names is None or name in names}


_ORDER = tuple(field.name for field in fields(SheetOptions))
_FIELD_NAMES = frozenset(_ORDER)
_DEFAULTS = {field.name: field.default for field in fields(SheetOptions)}

# The options each writer accepts. XlsbWriter has no ignored_errors.
XLSX_OPTIONS = _FIELD_NAMES
XLSB_OPTIONS = _FIELD_NAMES - {"ignored_errors"}


def resolve_sheet_options(
    options: SheetOptions | None,
    overrides: dict[str, Any],
    supported: frozenset[str],
    writer: str,
) -> SheetOptions:
    """Merge a SheetOptions object with keyword arguments; keyword arguments win.

    Raises TypeError for an unknown name, for an option the writer does not
    support, or for an ``options`` argument that is not a SheetOptions.
    """
    if options is None:
        base = SheetOptions()
    elif isinstance(options, SheetOptions):
        base = options
    else:
        raise TypeError("options must be a SheetOptions instance")
    for name in overrides:
        if name not in _FIELD_NAMES:
            raise TypeError(f"write_sheet() got an unexpected keyword argument {name!r}")
    for name in _FIELD_NAMES - supported:
        if name in overrides or getattr(base, name) != _DEFAULTS[name]:
            raise TypeError(f"write_sheet() got an unexpected keyword argument {name!r} for {writer}")
    return replace(base, **overrides) if overrides else base
