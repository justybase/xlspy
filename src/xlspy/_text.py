"""Cell text rules shared by the writers and the XLSX reader.

Length: Excel stores at most 32,767 characters in a cell, counted in UTF-16
code units, so a character outside the Basic Multilingual Plane counts as two.
The same limit applies to XLSX and XLSB, to shared and inline strings.

XML text: OOXML strings (ST_Xstring) carry characters that XML 1.0 cannot
hold as ``_xHHHH_``. Control characters, carriage returns (which an XML parser
would turn into line feeds) and U+FFFE/U+FFFF are written that way, and the
underscore of literal text that already looks like ``_xHHHH_`` is written as
``_x005F_``, so Excel reads back exactly the original text.
"""
from __future__ import annotations

import re
from xml.sax import saxutils


def _attr(value: object) -> str:
    """A value for a double-quoted XML attribute. Newlines become ``&#10;``, so
    a parser does not turn them into spaces."""
    return saxutils.escape(str(value), {'"': "&quot;", "\n": "&#10;"})

MAX_CELL_TEXT_UNITS = 32767
# Each code point is one or two UTF-16 units, so text this short always fits.
_ALWAYS_FITS = MAX_CELL_TEXT_UNITS // 2

_XSTRING_SPECIAL = re.compile(r"[\x00-\x08\x0b-\x1f￾￿]|_(?=x[0-9A-Fa-f]{4}_)")
_XSTRING_ESCAPE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def utf16_units(text: str) -> int:
    """Length of ``text`` in UTF-16 code units; a lone surrogate counts as one."""
    return len(text.encode("utf-16-le", "surrogatepass")) // 2


def check_cell_text(text: str) -> None:
    """Raise ``ValueError`` if ``text`` exceeds Excel's per-cell limit."""
    if len(text) > _ALWAYS_FITS:
        units = utf16_units(text)
        if units > MAX_CELL_TEXT_UNITS:
            raise ValueError(
                f"cell text exceeds Excel's {MAX_CELL_TEXT_UNITS}-character limit "
                f"({units} UTF-16 code units)"
            )


# Characters XML 1.0 cannot carry at all (tab, line feed and carriage return are fine).
_XML_FORBIDDEN = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff" + chr(0xFFFE) + chr(0xFFFF) + "]"
)


def check_xml_text(value: str, what: str) -> None:
    """Raise ``ValueError`` for text that cannot be written into an XML part.

    Parts that the writers build by hand (chart, drawing and shape XML) have no
    escape for a control character, so a title holding one made the part unparseable.
    """
    if _XML_FORBIDDEN.search(value) is not None:
        raise ValueError(f"{what} contains a control character, which XML cannot hold")


def _escape_match(match: re.Match[str]) -> str:
    return f"_x{ord(match.group()):04X}_"


def xstring_escape(text: str) -> str:
    """Apply the ``_xHHHH_`` escapes; text without special characters is returned as is.

    The input may already be XML-escaped: ``&amp;``, ``&lt;`` and ``&gt;`` contain
    neither control characters nor ``_x``, so the order does not matter.
    """
    # Fast pre-check: isprintable() is False for every control character and
    # for U+FFFE/U+FFFF, and the substring test is a plain scan.
    if "_x" not in text and text.isprintable():
        return text
    if _XSTRING_SPECIAL.search(text) is None:
        return text
    return _XSTRING_SPECIAL.sub(_escape_match, text)


def _unescape_match(match: re.Match[str]) -> str:
    return chr(int(match.group(1), 16))


def xstring_unescape(text: str) -> str:
    """Decode ``_xHHHH_`` escapes, as Excel does when it reads a cell."""
    if "_x" not in text:
        return text
    return _XSTRING_ESCAPE.sub(_unescape_match, text)
