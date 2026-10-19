"""Hyperlinks: a cell value that links to a web address or a cell in this workbook."""

from __future__ import annotations

import re
from dataclasses import dataclass

from .refs import parse_cell_range

MAX_TARGET_LENGTH = 2079
MAX_TOOLTIP_LENGTH = 255
ALLOWED_SCHEMES = frozenset({"http", "https", "mailto", "ftp"})
_SCHEME_RE = re.compile(r"^([A-Za-z][A-Za-z0-9+.\-]*):")
_CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")


@dataclass(frozen=True, slots=True)
class Hyperlink:
    """A link written as the value of a cell.

    Put it in a row where the text belongs::

        ["Report", Hyperlink("https://example.com", text="Open report")]

    ``target`` is an ``http``, ``https``, ``mailto`` or ``ftp`` URL, or an
    internal location written as ``"#Sheet Name!B2"`` (quote the sheet name
    with single quotes when it contains spaces: ``"#'Sales Data'!B2"``). Other
    schemes and relative file paths are refused. ``text`` is the displayed
    label and defaults to the target; ``tooltip`` is shown on hover. Linked
    cells use the built-in hyperlink look unless the cell has its own style.
    """

    target: str
    text: str | None = None
    tooltip: str | None = None

    def __post_init__(self):
        if not isinstance(self.target, str) or not self.target:
            raise TypeError("hyperlink target must be a non-empty string")
        if len(self.target) > MAX_TARGET_LENGTH:
            raise ValueError(f"hyperlink target may be at most {MAX_TARGET_LENGTH} characters")
        if _CONTROL_RE.search(self.target) or self.target != self.target.strip():
            raise ValueError("hyperlink target must not contain control characters or surrounding spaces")
        if self.target.startswith("#"):
            _check_location(self.target[1:])
        else:
            match = _SCHEME_RE.match(self.target)
            scheme = match.group(1).lower() if match else None
            if scheme not in ALLOWED_SCHEMES:
                raise ValueError(
                    f"unsupported hyperlink target {self.target!r}; use an http, https, "
                    "mailto or ftp URL, or '#Sheet!A1' for a cell in this workbook"
                )
        if self.text is not None and not isinstance(self.text, str):
            raise TypeError("hyperlink text must be a string or None")
        if self.tooltip is not None:
            if not isinstance(self.tooltip, str):
                raise TypeError("hyperlink tooltip must be a string or None")
            if len(self.tooltip) > MAX_TOOLTIP_LENGTH:
                raise ValueError(f"hyperlink tooltip may be at most {MAX_TOOLTIP_LENGTH} characters")

    @property
    def is_internal(self) -> bool:
        return self.target.startswith("#")

    @property
    def location(self) -> str:
        """The ``Sheet!A1`` form of an internal link."""
        return self.target[1:]

    @property
    def display_text(self) -> str:
        return self.target if self.text is None else self.text


def _check_location(location: str) -> None:
    sheet, separator, cell = location.rpartition("!")
    if not separator or not sheet:
        raise ValueError("internal hyperlink must look like '#Sheet!A1'")
    if _CONTROL_RE.search(location):
        raise ValueError("internal hyperlink must not contain control characters")
    parse_cell_range(cell)
