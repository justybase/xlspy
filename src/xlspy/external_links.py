"""Links to other workbooks (``[1]Sheet!A1`` in formulas), read from the package."""

from __future__ import annotations

import posixpath
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET
from ._safe_xml import fromstring as safe_fromstring

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


@dataclass(slots=True)
class ExternalLink:
    """A workbook that formulas refer to, as ``[index]`` in their text.

    ``index`` counts from 1 as in the formulas, ``target`` is the path or URL
    of the other workbook (``None`` when the file does not say) and
    ``sheet_names`` its sheets. Links are read, not edited; they are saved as
    they were.
    """

    index: int
    target: str | None
    sheet_names: list[str] = field(default_factory=list)
    part: str | None = None


def _relationships(package, part: str) -> dict[str, tuple[str, str]]:
    path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if not package.has(path):
        return {}
    try:
        root = safe_fromstring(package.get(path))
    except ET.ParseError:
        return {}
    return {
        node.get("Id") or "": (node.get("Target") or "", node.get("Type") or "")
        for node in root.iter(_PKG_REL) if node.get("Id")
    }


def _resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _path_target(package, part: str, rid: str | None) -> str | None:
    if rid is None:
        return None
    target = _relationships(package, part).get(rid)
    return target[0] if target else None


def read_external_links(package, *, binary: bool) -> list[ExternalLink]:
    """The workbook's external links in formula order."""
    book = "xl/workbook.bin" if binary else "xl/workbook.xml"
    if not package.has(book):
        return []
    relationships = _relationships(package, book)
    order: list[str] = []
    if binary:
        from ._xlsb_metadata import read_wide, records

        for record, payload in records(package.get(book)):
            if record.record_id == 0x163:
                order.append(read_wide(payload, 0)[0])
    else:
        try:
            root = safe_fromstring(package.get(book))
        except ET.ParseError:
            return []
        order = [
            node.get(f"{{{_REL}}}id") or ""
            for node in root.iter(f"{{{_MAIN}}}externalReference")
        ]
    links: list[ExternalLink] = []
    for position, rid in enumerate(order, start=1):
        target = relationships.get(rid)
        if target is None or not package.has(_resolve(book, target[0])):
            # Keep a placeholder: formulas number links by their position, so
            # dropping this one moved every later [n] onto the wrong workbook.
            links.append(ExternalLink(position, None))
            continue
        part = _resolve(book, target[0])
        link = ExternalLink(position, None, _sheet_names(package, part, binary), part)
        link.target = _path_target(package, part, _target_rid(package, part, binary))
        links.append(link)
    return links


def _target_rid(package, part: str, binary: bool) -> str | None:
    data = package.get(part)
    if binary:
        from ._xlsb_metadata import read_wide, records

        for record, payload in records(data):
            if record.record_id == 0x168:  # BrtSupBookSrc of the part
                return read_wide(payload, 2)[0] or None
        return None
    try:
        root = safe_fromstring(data)
    except ET.ParseError:
        return None
    node = next(root.iter(f"{{{_MAIN}}}externalBook"), None)
    return None if node is None else node.get(f"{{{_REL}}}id")


def _sheet_names(package, part: str, binary: bool) -> list[str]:
    data = package.get(part)
    if binary:
        from ._xlsb_formulas import external_sheet_names

        return external_sheet_names(data)
    try:
        root = safe_fromstring(data)
    except ET.ParseError:
        return []
    return [node.get("val") or "" for node in root.iter(f"{{{_MAIN}}}sheetName")]
