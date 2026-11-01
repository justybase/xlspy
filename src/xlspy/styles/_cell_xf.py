"""Cell xf (style index) registration shared by the XLSX and XLSB writers."""

from __future__ import annotations
from dataclasses import replace
from ._ranges import merge_styles
from ._style import Style


def _xf_for_style_format(style: Style, overlay: str | Style, register_style, cache: dict) -> int:
    """XF of ``style`` with a cell's number format or ``Style`` on top,
    merged once per pair."""
    key: tuple[object, ...]
    if isinstance(overlay, str):
        key = (id(style), overlay)
        cached = cache.get(key)
        if cached is not None and cached[0] is style:
            return cached[2]
        merged = merge_styles(style, Style(number_format=overlay))
    else:
        key = (id(style), id(overlay))
        cached = cache.get(key)
        if cached is not None and cached[0] is style and cached[1] is overlay:
            return cached[2]
        merged = merge_styles(style, overlay)
    xf_index = register_style(merged)
    # Holding both styles keeps their ids from being reused by other objects.
    cache[key] = (style, overlay, xf_index)
    return xf_index


def _record_cell_xf(fmt_xf: dict, keep: list, base_xf, overlay, xf_index: int) -> None:
    """Record the xf of a ``(value, format)`` or ``(value, Style)`` cell for
    the native row encoders: keyed by the format or the Style's ``id()``, paired
    with the cell's range-style xf when it has one. ``keep`` holds every
    keyed Style so its id cannot be reused while the writer is open."""
    key: object = overlay if isinstance(overlay, str) else id(overlay)
    if base_xf is not None:
        key = (base_xf, key)
    if key not in fmt_xf:
        fmt_xf[key] = xf_index
        if not isinstance(overlay, str):
            keep.append(overlay)


def _date_cell_xf(fmt_xf: dict, xf_styles: dict, register_style, xf_index: int,
                  date_format: str, is_datetime: bool) -> int:
    """XF for a date cell whose style is ``xf_index``.

    A style with no number format gets ``date_format`` on top, so Excel does
    not show the serial number. The result is recorded in ``fmt_xf`` under
    ``("date" | "datetime", xf_index)`` for the native row encoders. ``xf_styles``
    maps registered style xfs to their ``Style``; other xfs (format-only
    ones) already carry a number format and are kept.
    """
    key = ("datetime" if is_datetime else "date", xf_index)
    cached = fmt_xf.get(key)
    if cached is not None:
        return cached
    style = xf_styles.get(xf_index)
    result = xf_index
    if style is not None and style.number_format is None:
        result = register_style(replace(style, number_format=date_format))
    fmt_xf[key] = result
    return result


def _style_ids_for_row(row_styles, row_length: int, register_style, cache: dict):
    """Convert a shared row-style map once per distinct row-style pattern."""
    if not row_styles:
        return None
    key = id(row_styles)
    cached = cache.get(key)
    if cached is not None and cached[0] is row_styles:
        return cached[1]
    first_style = next(iter(row_styles.values()))
    if (len(row_styles) == row_length
            and all(style is first_style for style in row_styles.values())):
        style_ids = register_style(first_style)
    else:
        style_ids = tuple(
            (column, register_style(row_styles[column]))
            for column in sorted(row_styles)
        )
    # Keep the style map alive along with its cached ids so the identity key
    # cannot be reused for a different map during this sheet's write.
    cache[key] = (row_styles, style_ids)
    return style_ids


# Built-in look of a hyperlink cell that has no style of its own.
HYPERLINK_STYLE = Style(font_color="#0563C1", underline=True)
