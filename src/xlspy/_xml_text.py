"""Small XML text and attribute helpers shared by the updaters and the model.

These helpers never parse a whole worksheet: they operate on the bytes or on
an opening tag, so callers can keep markup they do not understand.
"""

from __future__ import annotations

import html
import posixpath
import re
from typing import Any
from xml.parsers import expat
from xml.sax import saxutils

from ._safe_xml import harden_expat
from ._text import check_cell_text, xstring_escape


def xml_attribute(attrs: dict[str, str], wanted: str) -> str | None:
    for name, value in attrs.items():
        if local_xml_name(name) == wanted:
            return value
    return None


def escape_xml_text(value: Any) -> str:
    """Escape a value for use as XML text or an XML attribute value.

    Control characters are written as ``_xHHHH_`` escapes, as the writers write
    them, so they survive; text over Excel's cell limit is refused.
    """

    text = str(value)
    check_cell_text(text)
    return saxutils.escape(xstring_escape(text), {"\"": "&quot;", "'": "&apos;"})


def unescape_xml_text(value: str) -> str:
    """Decode XML/HTML character references in a value read from XML text."""

    return html.unescape(value)


def column_index_to_letter(index: int) -> str:
    """Convert a zero-based column index to an Excel column name."""

    if index < 0:
        raise ValueError("Column index cannot be negative")
    result: list[str] = []
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def column_letter_to_index(reference: str) -> int:
    """Extract a zero-based column index from an A1-style cell reference."""

    letters = "".join(character for character in reference if character.isalpha())
    if not letters:
        raise ValueError(f"Invalid cell reference: {reference!r}")
    result = 0
    for character in letters.upper():
        if not "A" <= character <= "Z":
            raise ValueError(f"Invalid cell reference: {reference!r}")
        result = result * 26 + ord(character) - ord("A") + 1
    return result - 1


def local_xml_name(name: str) -> str:
    """Return an XML local name for either prefixed or namespace-expanded input."""

    if "}" in name:
        name = name.rsplit("}", 1)[1]
    if ":" in name:
        name = name.rsplit(":", 1)[1]
    return name


_XML_ATTRIBUTE_RE_TEMPLATE = r"(?P<prefix>(?:[A-Za-z_][\w.-]*:)?{name})\s*=\s*(?P<quote>[\"\'])(?P<value>.*?)(?P=quote)"


def replace_or_add_xml_attribute(opening_tag: str, name: str, value: Any) -> str:
    """Replace an attribute in an opening tag, or add it before its closing marker."""

    escaped = escape_xml_text(value)
    pattern = re.compile(_XML_ATTRIBUTE_RE_TEMPLATE.format(name=re.escape(name)), re.DOTALL)
    match = pattern.search(opening_tag)
    if match:
        start, end = match.span("value")
        return opening_tag[:start] + escaped + opening_tag[end:]

    insertion = f' {name}="{escaped}"'
    if opening_tag.endswith("/>"):
        return opening_tag[:-2] + insertion + "/>"
    return opening_tag[:-1] + insertion + ">"


def parse_relationships_xml(xml_data: bytes) -> dict[str, str]:
    """Read relationship IDs and targets from an OOXML relationships part."""

    relationships: dict[str, str] = {}
    parser = harden_expat(expat.ParserCreate())

    def start_element(name: str, attrs: dict[str, str]) -> None:
        if local_xml_name(name) != "Relationship":
            return
        relationship_id = attrs.get("Id")
        target = attrs.get("Target")
        if relationship_id and target:
            relationships[relationship_id] = target

    parser.StartElementHandler = start_element
    parser.Parse(xml_data, True)
    return relationships


def resolve_zip_target(target: str, base_directory: str = "xl") -> str:
    """Resolve an OOXML relationship target to a ZIP member name."""

    if target.startswith("/"):
        resolved = target[1:]
    else:
        resolved = posixpath.normpath(posixpath.join(base_directory, target))
    return resolved.replace("\\", "/")
