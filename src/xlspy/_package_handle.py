"""The object model's view of an updater's package.

``load_workbook`` keeps an ``XlsxUpdater`` or ``XlsbUpdater`` behind a loaded
workbook, and ``Workbook.save`` writes its changes through it. Beyond the
updater's public methods, the model needs the staged package, the sheet-to-part
map and a few part helpers. This class is the only way ``xlspy.workbook`` reaches
them, so the updaters' private state has one typed contract with the model
(``tests/test_package_handle.py`` checks that nothing else is used).

A handle holds no state of its own: every property reads the updater, so it
stays valid while the updater changes, and creating one is cheap.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, MutableMapping
from typing import TYPE_CHECKING, Any
from .xlsx_updater import XlsxUpdater
from .xlsb_updater import XlsbUpdater

if TYPE_CHECKING:
    from ._xlsb_sheets import _Entry
    from .updater_utils import ZipPackage


class PackageHandle:
    """Package-level access to an updater for the object model."""

    __slots__ = ("updater",)

    def __init__(self, updater: XlsxUpdater | XlsbUpdater) -> None:
        self.updater = updater

    @property
    def package(self) -> ZipPackage:
        """The staged package: ``get``/``set``/``has``/``names`` over its parts."""
        return self.updater._package

    @property
    def sheets(self) -> MutableMapping[str, str]:
        """Sheet name -> part path, in workbook order. The model updates it when it
        adds, renames or removes a sheet; the updater reads the same mapping."""
        return self.updater._sheets

    @property
    def date_1904(self) -> bool:
        return self.updater._date_1904

    @property
    def chartsheets(self) -> set[str]:
        """Names of the chartsheets of an XLSB package (they are also in ``sheets``)."""
        return self._xlsb()._chartsheets

    def apply_parts(self, parts: Mapping[str, str | bytes], removed: Iterable[str] = ()) -> None:
        """Write ``parts`` and remove ``removed`` as one transaction."""
        self.package.apply_parts(parts, removed)

    def recalculate(self, updates: MutableMapping[str, Any], removed: list[str]) -> None:
        """Add to ``updates``/``removed`` what makes Excel recalculate an XLSX on open."""
        self._xlsx()._recalculate(updates, removed)

    def sheet_xml(self, name: str) -> tuple[str, str]:
        """The part path and XML of an XLSX worksheet."""
        return self._xlsx()._sheet_xml(name)

    def sheet_bin(self, name: str) -> tuple[str, bytes]:
        """The part path and records of an XLSB worksheet."""
        return self._xlsb()._sheet_bin(name)

    def part_or_none(self, path: str) -> bytes | None:
        return self.package.get(path) if self.package.has(path) else None

    def tag_prefix(self, xml: str, tag: str) -> str:
        """The namespace prefix (``""`` or ``"x:"``) the XML uses for ``tag``."""
        return self._xlsx()._tag_prefix(xml, tag)

    def rewrite_sheet_list(self, workbook: bytes, entries: list[_Entry], parts: dict[str, bytes]) -> bytes:
        """``xl/workbook.bin`` with its sheet list replaced by ``entries``."""
        return self._xlsb()._rewrite_sheet_list(workbook, entries, parts)

    def _xlsx(self) -> XlsxUpdater:
        if not isinstance(self.updater, XlsxUpdater):
            raise TypeError("an XLSX package operation was used on an XLSB workbook")
        return self.updater

    def _xlsb(self) -> XlsbUpdater:
        if not isinstance(self.updater, XlsbUpdater):
            raise TypeError("an XLSB package operation was used on an XLSX workbook")
        return self.updater
