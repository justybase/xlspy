"""The package transaction and defined-name helpers shared by both edit mixins.

The XLSX and XLSB edit mixins keep their wire formats separate; only the parts
that are identical in both formats live here.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ._metadata import DefinedName
    from .updater_utils import ZipPackage


class SheetEditMixin:
    """Shared edit plumbing for ``XlsxEditMixin`` and ``XlsbEditMixin``."""

    if TYPE_CHECKING:
        _package: ZipPackage

    def _apply_parts(self, parts: Mapping[str, Any], removed: Iterable[str] = ()) -> None:
        # A streaming replacement already holds the package transaction; nesting
        # another one would fail, so its parts are written directly into that
        # transaction (a failure rolls the whole replacement back).
        if self._package.in_transaction:
            for path, data in parts.items():
                self._package.set(path, data.encode("utf-8") if isinstance(data, str) else data)
            for path in removed:
                self._package.remove(path)
            return
        # The package owns the write-back transaction (ZipPackage.apply_parts).
        self._package.apply_parts(parts, removed)

    @staticmethod
    def _resolve_defined_name(definition: DefinedName, sheets: list[str]) -> DefinedName:
        """Resolve a scope title to the index of that sheet in ``sheets``."""
        return definition.resolved(sheets)
