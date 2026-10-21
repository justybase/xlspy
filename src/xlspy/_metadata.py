"""Small OOXML metadata helpers; never require a spreadsheet engine."""

from __future__ import annotations

from dataclasses import dataclass, fields
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from .refs import parse_cell_range, column_label

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
# ISO/IEC 29500 Strict ("Strict Open XML Spreadsheet") namespaces.
STRICT_MAIN = "http://purl.oclc.org/ooxml/spreadsheetml/main"
STRICT_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships"


def namespace_of(tag: str) -> str:
    """Namespace URI of an ElementTree tag such as ``{uri}local``."""
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def relationship_namespace(main: str) -> str:
    """The r:id namespace matching a Transitional or Strict main namespace."""
    return STRICT_REL if main == STRICT_MAIN else REL


def relationship_kind(rel_type: str) -> str | None:
    """Last segment of a Transitional or Strict officeDocument relationship type."""
    for base in (REL, STRICT_REL):
        if rel_type.startswith(base + "/"):
            return rel_type[len(base) + 1:]
    return None


def cell_ref(ref: str) -> str:
    c, r, c2, r2 = parse_cell_range(ref)
    if (c, r) != (c2, r2) or ":" in ref:
        raise ValueError("a single cell reference is required")
    return f"{column_label(c)}{r + 1}"


def xml_text(text: str) -> str:
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    if any(not (ch in "\t\n\r" or 0x20 <= ord(ch) <= 0xD7FF
                or 0xE000 <= ord(ch) <= 0xFFFD or 0x10000 <= ord(ch) <= 0x10FFFF)
           for ch in text):
        raise ValueError("text contains a character forbidden in XML")
    return text


_DECLARATION = re.compile(r"""\sxmlns(?::([\w.-]+))?\s*=\s*(?:"([^"]*)"|'([^']*)')""")


def serialize_like(root, original) -> bytes:
    """Serialize an ElementTree root with the source part's namespace prefixes.

    ElementTree renames prefixes (``ns0``...) and drops declarations that no
    element uses. Excel parts list prefixes in ``mc:Ignorable`` and
    ``mc:Choice Requires``, so both changes make Excel reject the part. The
    original prefixes are reused and every root declaration is kept.
    """
    text = original.decode("utf-8") if isinstance(original, bytes) else original
    opening = re.search(r"<(?![?!])[^>]+>", text)[0]
    declarations = [(m[1] or "", m[2] if m[2] is not None else m[3]) for m in _DECLARATION.finditer(opening)]
    default = next((uri for prefix, uri in declarations if not prefix), None)
    # Nested elements (extLst, AlternateContent) declare prefixes of their own.
    nested = [(m[1], m[2] if m[2] is not None else m[3]) for m in _DECLARATION.finditer(text) if m[1]]
    for prefix, uri in dict.fromkeys(nested[::-1] + [d for d in declarations if d[0]][::-1]):
        if not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)
    if default is not None:
        ET.register_namespace("", default)
    try:
        output = ET.tostring(root, encoding="unicode")
    finally:
        if default is not None:
            # The empty prefix is global; do not leak it into later calls.
            ET._namespace_map.pop(default, None)
    start = re.match(r"<[^>]+?(?=/?>)", output)[0]
    present = {m[1] or "" for m in _DECLARATION.finditer(start)}
    missing = "".join(
        f' xmlns{":" + prefix if prefix else ""}="{uri}"'
        for prefix, uri in declarations if prefix not in present
    )
    output = start + missing + output[len(start):]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + output).encode("utf-8")


def element_pattern(tag: str):
    return re.compile(
        rf"<(?P<prefix>[A-Za-z_][\w.-]*:)?{tag}\b[^>]*?"
        rf"(?:/>|>.*?</(?P=prefix){tag}\s*>|>.*?</{tag}\s*>)", re.DOTALL,
    )


def replace_element(xml: str, tag: str, fragment: str, *, before=()) -> str:
    """Replace a top-level metadata element, preserving all other XML bytes."""
    pattern = element_pattern(tag)
    match = pattern.search(xml)
    if match:
        return xml[:match.start()] + fragment + xml[match.end():]
    if not fragment:
        return xml
    for successor in before:
        match = re.search(rf"<(?:[A-Za-z_][\w.-]*:)?{successor}\b", xml)
        if match:
            return xml[:match.start()] + fragment + xml[match.start():]
    closing = list(re.finditer(r"</(?:[A-Za-z_][\w.-]*:)?[\w.-]+\s*>", xml))
    if not closing:
        # Empty documents such as <Relationships/>.
        match = re.search(r"<([\w:.-]+)\b([^>]*)/>", xml)
        if not match:
            raise ValueError("XML root element was not found")
        return xml[:match.start()] + f"<{match[1]}{match[2]}>" + fragment + f"</{match[1]}>" + xml[match.end():]
    pos = closing[-1].start()
    return xml[:pos] + fragment + xml[pos:]


def append_child(xml: str, container: str, fragment: str) -> str:
    match = element_pattern(container).search(xml)
    if not match:
        raise ValueError(f"{container} element was not found")
    body = match[0]
    if body.endswith("/>"):
        name = re.match(r"<([\w:.-]+)", body)[1]
        body = body[:-2] + ">" + fragment + f"</{name}>"
    else:
        pos = body.rfind("</")
        body = body[:pos] + fragment + body[pos:]
    return xml[:match.start()] + body + xml[match.end():]


def password_hash(password: str) -> str:
    """Excel legacy worksheet/workbook password hash (not encryption)."""
    value = 0
    for index, ch in enumerate(password, 1):
        code = ord(ch) << index
        value ^= (code & 0x7FFF) | (code >> 15)
    return f"{(value ^ len(password) ^ 0xCE4B) & 0xFFFF:04X}"


@dataclass(frozen=True, slots=True)
class DefinedName:
    """A global name, or a local name scoped to a worksheet's title."""

    name: str
    formula: str
    scope: str | None = None

    def __post_init__(self):
        if not isinstance(self.name, str) or not re.fullmatch(r"[^\W\d][\w.]*|\\[\w.]*", self.name):
            raise ValueError("invalid defined name")
        if len(self.name) > 255 or self.name.upper() in {"R", "C"} or self.name.lower().startswith("_xlnm."):
            raise ValueError("reserved or oversized defined name")
        try:
            parse_cell_range(self.name)
        except ValueError:
            pass
        else:
            raise ValueError("defined name must not be a cell reference")
        if re.fullmatch(r"R[1-9]\d*C[1-9]\d*", self.name, re.I):
            raise ValueError("defined name must not be an R1C1 reference")
        if not xml_text(self.formula).removeprefix("="):
            raise ValueError("defined name formula must not be empty")
        if self.scope is not None and not isinstance(self.scope, str):
            raise TypeError("scope must be a worksheet title or None")
        object.__setattr__(self, "formula", self.formula.removeprefix("="))

    def to_xml(self, sheets) -> str:
        scope = ""
        if self.scope is not None:
            if self.scope not in sheets:
                raise KeyError(f"Worksheet not found: {self.scope}")
            scope = f' localSheetId="{list(sheets).index(self.scope)}"'
        return f'<definedName xmlns="{MAIN}" name={quoteattr(self.name)}{scope}>{escape(self.formula)}</definedName>'


def read_defined_names(xml: str, sheets) -> list[DefinedName]:
    result = []
    root = ET.fromstring(xml)
    ns = namespace_of(root.tag) or MAIN
    for node in root.findall(f"{{{ns}}}definedNames/{{{ns}}}definedName"):
        if node.get("name", "").lower().startswith("_xlnm."):
            continue
        scope = node.get("localSheetId")
        result.append(DefinedName(node.attrib["name"], node.text or "", None if scope is None else list(sheets)[int(scope)]))
    return result


@dataclass(frozen=True, slots=True)
class SheetProtection:
    """Enable sheet protection. Selection flags mean 'allow selection'."""

    enabled: bool = True
    select_locked_cells: bool = True
    select_unlocked_cells: bool = True
    password: str | None = None

    def __post_init__(self):
        _validate_protection(self)

    def to_xml(self):
        if not self.enabled:
            return ""
        password = "" if self.password is None else f' password="{password_hash(self.password)}"'
        return (f'<sheetProtection xmlns="{MAIN}" sheet="1" objects="1" scenarios="1"'
                f' selectLockedCells="{int(not self.select_locked_cells)}"'
                f' selectUnlockedCells="{int(not self.select_unlocked_cells)}"{password}/>')


@dataclass(frozen=True, slots=True)
class WorkbookProtection:
    """Protect workbook structure; this does not encrypt workbook data."""

    lock_structure: bool = True
    password: str | None = None

    def __post_init__(self):
        _validate_protection(self)

    def to_xml(self):
        if not self.lock_structure:
            return ""
        password = "" if self.password is None else f' workbookPassword="{password_hash(self.password)}"'
        return f'<workbookProtection xmlns="{MAIN}" lockStructure="1"{password}/>'


def _validate_protection(obj):
    for field in fields(obj):
        value = getattr(obj, field.name)
        if field.name == "password":
            if value is not None and (not isinstance(value, str) or len(value) > 15 or not value.isascii()):
                raise ValueError("legacy protection passwords must be ASCII strings of at most 15 characters")
        elif not isinstance(value, bool):
            raise TypeError(f"{field.name} must be bool")

