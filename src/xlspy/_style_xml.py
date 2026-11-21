"""Style XML fragments shared by the XLSX writer and the loaded style table.

Both sides store the same hashable component keys (``("rgb", "FF112233")``,
``("theme", 2, 0.5)``, border sides as ``(style, color)``). Rendering them in
one place keeps a new workbook and a saved loaded workbook byte-compatible.
"""

from __future__ import annotations

from typing import Any

from ._xlsx_writer_xml import _format_number


def color_attrs(key: Any) -> str:
    """The attributes of a ``<color>`` element for a color key."""
    kind = key[0]
    if kind == "rgb":
        return f'rgb="{key[1]}"'
    if kind == "indexed":
        return f'indexed="{key[1]}"'
    if kind == "auto":
        return 'auto="1"'
    _, theme, tint = key
    if tint is None:
        return f'theme="{theme}"'
    return f'theme="{theme}" tint="{_format_number(tint)}"'


def border_side_xml(tag: str, side_key: Any) -> str:
    """One border edge; an empty ``(None, None)`` is a self-closing element."""
    style, color_key = side_key
    if style is None and color_key is None:
        return f"<{tag}/>"
    inner = f"<color {color_attrs(color_key)}/>" if color_key is not None else ""
    if style is None:
        return f"<{tag}>{inner}</{tag}>"
    return f'<{tag} style="{style}">{inner}</{tag}>'


def border_xml(key: Any) -> str:
    """The ``<border>`` element for a border key."""
    attrs = ""
    if key.diagonal_up:
        attrs += ' diagonalUp="1"'
    if key.diagonal_down:
        attrs += ' diagonalDown="1"'
    if key.outline is False:
        attrs += ' outline="0"'
    # Excel reads <start>/<end> as the left/right edges (in a left-to-right
    # sheet), so an empty <start/> would hide <left>: write them only when
    # set, like <vertical>/<horizontal>.
    parts = [border_side_xml(tag, getattr(key, tag)) for tag in ("start", "end") if getattr(key, tag) != (None, None)]
    parts += [
        border_side_xml("left", key.left),
        border_side_xml("right", key.right),
        border_side_xml("top", key.top),
        border_side_xml("bottom", key.bottom),
        border_side_xml("diagonal", key.diagonal),
    ]
    if key.vertical != (None, None):
        parts.append(border_side_xml("vertical", key.vertical))
    if key.horizontal != (None, None):
        parts.append(border_side_xml("horizontal", key.horizontal))
    return f"<border{attrs}>{''.join(parts)}</border>"
