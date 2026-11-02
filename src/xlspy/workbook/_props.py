"""Document properties of a loaded workbook, saved back into its package.

XLSX and XLSB keep them in the same XML parts: ``docProps/core.xml`` for the
core properties and ``docProps/custom.xml`` for user-defined ones.
"""

from __future__ import annotations

import datetime as _datetime
from xml.etree import ElementTree as ET
from .._safe_xml import fromstring as safe_fromstring
from xml.sax.saxutils import escape

from .._metadata import CT, PKG_REL
from ..doc_properties import (
    _CORE_REL, _CORE_TYPE, _CUSTOM_FMTID, _CUSTOM_REL, _CUSTOM_TYPE, _check_doc_text, _custom_value, _utc,
)

_CORE_PATH, _CUSTOM_PATH = "docProps/core.xml", "docProps/custom.xml"
# Model field -> element, in the order Excel writes them.
_CORE = (
    ("title", "dc:title"), ("subject", "dc:subject"), ("creator", "dc:creator"),
    ("keywords", "cp:keywords"), ("description", "dc:description"),
    ("lastModifiedBy", "cp:lastModifiedBy"), ("revision", "cp:revision"),
    ("category", "cp:category"), ("contentStatus", "cp:contentStatus"),
    ("identifier", "dc:identifier"), ("language", "dc:language"), ("version", "cp:version"),
)
_DATES = (("lastPrinted", "cp:lastPrinted"), ("created", "dcterms:created"), ("modified", "dcterms:modified"))


def signature(workbook) -> tuple:
    return repr(workbook.properties), repr(list(workbook.custom_doc_props))


def remember(workbook) -> None:
    workbook._saved_properties = signature(workbook)


def core_xml(props) -> bytes:
    body = []
    for name, tag in _CORE:
        value = getattr(props, name)
        if value is not None:
            body.append(f"<{tag}>{escape(_check_doc_text(name, str(value)))}</{tag}>")
    for name, tag in _DATES:
        value = getattr(props, name)
        if isinstance(value, _datetime.datetime):
            kind = ' xsi:type="dcterms:W3CDTF"' if tag.startswith("dcterms") else ""
            body.append(f"<{tag}{kind}>{_utc(value, name)}</{tag}>")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        + "".join(body) + "</cp:coreProperties>"
    ).encode("utf-8")


def custom_xml(items) -> bytes:
    body = "".join(
        f'<property fmtid="{_CUSTOM_FMTID}" pid="{pid}" name="{escape(item.name, {chr(34): "&quot;"})}">'
        f"{_custom_value(item.name, item.value)}</property>"
        for pid, item in enumerate(items, start=2)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"{body}</Properties>"
    ).encode()


def _serialize(root) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def save_properties(updater, workbook) -> None:
    """Write the properties when they changed since the workbook was loaded or saved."""
    if getattr(workbook, "_saved_properties", None) == signature(workbook):
        return
    package = updater._package
    types = safe_fromstring(package.get("[Content_Types].xml"))
    rels = safe_fromstring(package.get("_rels/.rels"))
    parts = {_CORE_PATH: core_xml(workbook.properties)}
    removed = []

    def ensure(path, content_type, rel_type):
        if not any(node.get("PartName") == "/" + path for node in types):
            ET.SubElement(types, f"{{{CT}}}Override", {"PartName": "/" + path, "ContentType": content_type})
        if not any(node.get("Type") == rel_type for node in rels):
            taken = {node.get("Id") for node in rels}
            number = 1
            while f"rId{number}" in taken:
                number += 1
            ET.SubElement(rels, f"{{{PKG_REL}}}Relationship", {
                "Id": f"rId{number}", "Type": rel_type, "Target": path,
            })

    ensure(_CORE_PATH, _CORE_TYPE, _CORE_REL)
    if workbook.custom_doc_props:
        parts[_CUSTOM_PATH] = custom_xml(workbook.custom_doc_props)
        ensure(_CUSTOM_PATH, _CUSTOM_TYPE, _CUSTOM_REL)
    elif package.has(_CUSTOM_PATH):
        removed.append(_CUSTOM_PATH)
        for node in list(types):
            if node.get("PartName") == "/" + _CUSTOM_PATH:
                types.remove(node)
        for node in list(rels):
            if node.get("Type") == _CUSTOM_REL:
                rels.remove(node)
    ET.register_namespace("", CT)
    parts["[Content_Types].xml"] = _serialize(types)
    ET.register_namespace("", PKG_REL)
    parts["_rels/.rels"] = _serialize(rels)
    updater._apply_parts(parts, removed)
    remember(workbook)
