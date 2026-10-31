"""OOXML ``_xHHHH_`` escaping, compatible with openpyxl."""

from __future__ import annotations

import re

_ESCAPE_RE = re.compile(r"[\001-\031]")
_UNESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def escape(value: str) -> str:
    """Escape control characters 0x01..0x1F as ``_xHHHH_``."""
    return _ESCAPE_RE.sub(lambda match: f"_x{ord(match.group(0)):0>4x}_", value)


def unescape(value: str) -> str:
    """Decode ``_xHHHH_`` escapes."""
    if "_x" in value:
        value = _UNESCAPE_RE.sub(lambda match: chr(int(match.group(1), 16)), value)
    return value
