"""Macro-enabled workbooks and templates: the workbook part's content type
and, for .xlsm/.xltm, the VBA project taken from ``Workbook.vba_archive``."""

from __future__ import annotations
import zipfile
from typing import TYPE_CHECKING
from xml.etree import ElementTree as ET

from .._metadata import CT, MAIN, element_bytes
from .._safe_xml import fromstring as safe_fromstring
from .._vba import VBA_CONTENT_TYPE as _VBA_TYPE, VBA_PART as _VBA_PART, VBA_RELATIONSHIP as _VBA_REL
from ._objects import _Package
from ._chartsheets import _set_rels

if TYPE_CHECKING:
    from ..xlsx_updater import XlsxUpdater

_MAIN_TYPES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}


def set_package_kind(
    updater: XlsxUpdater,
    extension: str,
    vba_archive: zipfile.ZipFile | None = None,
    *,
    vba_codename: str | None = None,
) -> None:
    """Give the workbook part the content type of ``extension``; a
    macro-enabled kind also gets the VBA project of ``vba_archive`` (a
    ``zipfile.ZipFile`` such as ``load_workbook(..., keep_vba=True)`` sets)
    and, when given, the project's ``vba_codename``."""
    main = _MAIN_TYPES.get(extension)
    if main is None or extension == "xlsx":
        return
    package = _Package(updater)
    for node in package.types:
        if node.tag == f"{{{CT}}}Override" and node.get("PartName") == "/xl/workbook.xml":
            node.set("ContentType", main)
    if extension in ("xlsm", "xltm") and vba_archive is not None:
        if _VBA_PART not in vba_archive.namelist():
            raise ValueError("vba_archive holds no xl/vbaProject.bin")
        package.set(_VBA_PART, vba_archive.read(_VBA_PART))
        package.override(_VBA_PART, _VBA_TYPE)
        rels = package.rels("xl/workbook.xml")
        rid = package.add_rel(rels, "xl/workbook.xml", "vbaProject", _VBA_PART)
        # add_rel builds officeDocument relationship types; the VBA one is Microsoft's.
        next(node for node in rels if node.get("Id") == rid).set("Type", _VBA_REL)

        _set_rels(package, "xl/workbook.xml", rels)
        if vba_codename:
            root = safe_fromstring(package.require("xl/workbook.xml"))
            workbook_pr = root.find(f"{{{MAIN}}}workbookPr")
            if workbook_pr is None:
                workbook_pr = ET.Element(f"{{{MAIN}}}workbookPr")
                file_version = root.find(f"{{{MAIN}}}fileVersion")
                index = list(root).index(file_version) + 1 if file_version is not None else 0
                root.insert(index, workbook_pr)
            workbook_pr.set("codeName", vba_codename)
            package.set("xl/workbook.xml", element_bytes(root))
    package.commit()
