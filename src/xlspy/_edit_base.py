"""The package transaction and defined-name helpers shared by both edit mixins.

The XLSX and XLSB edit mixins keep their wire formats separate; only the parts
that are identical in both formats live here.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Any

from ._metadata import cell_ref
from .updater_utils import unwrap_cell

if TYPE_CHECKING:
    from ._metadata import DefinedName
    from .updater_utils import ZipPackage


class SheetEditMixin:
    """Shared edit plumbing for ``XlsxEditMixin`` and ``XlsbEditMixin``."""

    if TYPE_CHECKING:
        _package: ZipPackage
        _shared_strings_loaded: bool
        _shared_strings_total: int
        _stream_shared_strings: Any

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

    def _normalize_cell_edits(
        self,
        cells: Mapping[str, object],
        styles: Mapping[str, int] | None,
        *,
        cells_error: str,
        styles_error: str,
        duplicate_error: str,
        keep: object,
    ) -> tuple[dict[str, Any], dict[str, int], dict[str, Any]]:
        """Check and normalize an ``update_cells`` call for either wire format.

        Returns the values, the style ids and the targets (values plus the
        addresses that only set a style); ``keep`` is the format's sentinel for
        an existing value the style update must leave alone.
        """
        if not isinstance(cells, Mapping):
            raise TypeError(cells_error)
        if styles is not None and not isinstance(styles, Mapping):
            raise TypeError(styles_error)
        values: dict[str, Any] = {}
        for ref, value in cells.items():
            ref = cell_ref(ref)
            if ref in values:
                raise ValueError(duplicate_error.format(ref=ref))
            values[ref] = unwrap_cell(value)
        style_ids: dict[str, int] = {}
        for ref, xf in (styles or {}).items():
            ref = cell_ref(ref)
            if not isinstance(xf, int) or isinstance(xf, bool) or not 0 <= xf < 0xFFFFFF:
                raise ValueError(f"invalid cell style index {xf!r} for {ref}")
            style_ids[ref] = xf
        targets = dict(values)
        for ref in style_ids:
            targets.setdefault(ref, keep)
        return values, style_ids, targets

    def _account_removed_shared_strings(self, removed_shared_refs: int) -> None:
        """Keep the cached shared-string totals in step with a replacement."""
        if self._shared_strings_loaded:
            self._shared_strings_total = max(0, self._shared_strings_total - removed_shared_refs)
        if self._stream_shared_strings is not None:
            self._stream_shared_strings.total_count = max(
                0, self._stream_shared_strings.total_count - removed_shared_refs
            )
