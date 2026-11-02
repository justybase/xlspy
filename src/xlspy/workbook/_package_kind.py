"""Macro-enabled workbooks and templates: the workbook part's content type
and, for .xlsm/.xltm, the VBA project taken from ``Workbook.vba_archive``."""

from __future__ import annotations
import zipfile
from typing import TYPE_CHECKING

from .._metadata import CT
from ._objects import _Package

if TYPE_CHECKING:
    from ..xlsx_updater import XlsxUpdater

_MAIN_TYPES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
_VBA_PART = "xl/vbaProject.bin"
_VBA_TYPE = "application/vnd.ms-office.vbaProject"
_VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def set_package_kind(updater: XlsxUpdater, extension: str,
                     vba_archive: zipfile.ZipFile | None = None) -> None:
    """Give the workbook part the content type of ``extension``; a
    macro-enabled kind also gets the VBA project of ``vba_archive`` (a
    ``zipfile.ZipFile`` such as ``load_workbook(..., keep_vba=True)`` sets)."""
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
        from ._chartsheets import _set_rels

        _set_rels(package, "xl/workbook.xml", rels)
    package.commit()
