"""Lifecycle and shared-string bookkeeping shared by the XLSX and XLSB updaters.

Only code that is identical in both formats lives here.  The format-specific
parts stay in the subclasses: the ``_load_*`` readers, ``replace_sheet_data``
and the two hooks declared below.  Subclasses list this base first in their
bases, so these methods take precedence over the edit mixins, as they did when
they were defined on the classes themselves.
"""

from __future__ import annotations

import datetime as _datetime
import os
from pathlib import Path
from typing import Any, Literal
from collections.abc import Callable, Iterable, Sequence

from .excel_dates import datetime_to_serial
from .updater_utils import (
    DiskStringIndex,
    TableHeader,
    TableHeaderGuard,
    ZipPackage,
    copy_stream,
    date_like_format,
    header_row_texts,
    materialize_rows,
    sheet_table_headers,
    trim_trailing_empty_rows,
)

StyleFallback = Literal["inherit", "general"]

# Built-in number formats that show a date or time (ECMA-376 18.8.30).
_BUILT_IN_DATE_FORMATS = frozenset(range(14, 23)) | {45, 46, 47}


class _UpdaterBase:
    """Attributes and methods common to XlsxUpdater and XlsbUpdater."""

    # State set by the subclass __init__; shared-string fields are used by the
    # shared-string helpers below.
    _package: ZipPackage
    _date_1904: bool
    _sheets: dict[str, str]
    _shared_strings_loaded: bool
    _shared_strings_values: list[str]
    _shared_string_indices: dict[str, int]
    _shared_strings_total: int
    _shared_strings_pending: list[str]
    _shared_strings_dirty: bool
    _stream_shared_strings: DiskStringIndex | None

    def _ensure_stream_shared_strings(self) -> None:
        """Load the shared-string store used by streaming replacements (subclass hook)."""
        raise NotImplementedError

    def _replace_sheet_data_stream_impl(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Format-specific body of ``replace_sheet_data_stream`` (subclass hook)."""
        raise NotImplementedError

    def close(self) -> None:
        """Release temporary files held by the updater.

        While in use, the updater keeps shared-string data and staged package
        parts on disk.  Calling ``close`` (or simply dropping the updater)
        removes them.
        """

        if self._stream_shared_strings is not None:
            self._stream_shared_strings.close()
        self._package.close()

    def __del__(self) -> None:
        # An exception raised in __del__ cannot reach a caller; it would only be
        # reported as unraisable, so the cleanup swallows any failure.
        try:
            self.close()
        except Exception:
            pass

    def get_sheet_names(self) -> list[str]:
        """Return worksheet names in workbook order."""

        return list(self._sheets)

    def _read_table_header(self, table_part: bytes) -> TableHeader | None:
        """Parse one table part of this format into its header row (subclass hook)."""
        raise NotImplementedError

    def _table_header_guard(self, sheet_name: str, sheet_path: str) -> TableHeaderGuard:
        """Guard the header rows of the tables on one sheet during a data replacement."""

        return TableHeaderGuard(
            sheet_name,
            sheet_table_headers(self._package, sheet_path, self._read_table_header),
        )

    # -- steps shared by replace_sheet_data in both formats ----------------

    def _replacement_sheet_path(self, sheet_name: str, style_fallback: str) -> str:
        """Check the arguments of a data replacement; return the sheet's part path."""

        if style_fallback not in ("inherit", "general"):
            raise ValueError("style_fallback must be 'inherit' or 'general'")
        try:
            return self._sheets[sheet_name]
        except KeyError as exc:
            raise KeyError(f"Worksheet not found: {sheet_name}") from exc

    def _checked_rows(
        self,
        sheet_name: str,
        sheet_path: str,
        rows: Iterable[Sequence[Any]],
        headers: Sequence[str] | None,
    ) -> tuple[list[list[Any]], list[str] | None]:
        """The rows (materialized, trailing empty rows dropped) and the header
        values, checked against the header rows of the sheet's tables."""

        materialized = trim_trailing_empty_rows(materialize_rows(rows))
        header_values = list(headers) if headers is not None else None
        guard = self._table_header_guard(sheet_name, sheet_path)
        first_data_row = 0
        if header_values is not None:
            guard.check_row(0, header_row_texts(header_values))
            first_data_row = 1
        for index, row in enumerate(materialized, start=first_data_row):
            guard.check_row(index, row)
        guard.finish()
        return materialized, header_values

    @staticmethod
    def _region_size(materialized: list[list[Any]], header_values: list[str] | None) -> tuple[int, int]:
        """``(rows, columns)`` of the replacement region, the header row included."""

        width = max(
            [len(row) for row in materialized]
            + ([len(header_values)] if header_values is not None else [])
            + [0]
        )
        return len(materialized) + (1 if header_values is not None else 0), width

    def _cell_xf_number_formats(self) -> tuple[list[int], dict[int, str]] | None:
        """The number format id of each cell XF, in order, and the custom format
        codes by id; None without a styles part (subclass hook)."""
        raise NotImplementedError

    def _find_date_style(self) -> int | None:
        """Index of the first cell XF with a date number format, or None.

        Dates written without a style of their own get this one
        (``style_fallback="inherit"``).
        """

        found = self._cell_xf_number_formats()
        if found is None:
            return None
        xf_formats, custom_formats = found
        for index, number_format_id in enumerate(xf_formats):
            if number_format_id in _BUILT_IN_DATE_FORMATS:
                return index
            if date_like_format(custom_formats.get(number_format_id, "")):
                return index
        return None

    def _stage_part_copy(self, part: str, suffix: str) -> Path:
        """A temporary copy of a package part, for a streaming step to read."""

        copy = self._package.temporary_path(suffix=suffix)
        with self._package.open_entry(part) as source, copy.open("wb") as target:
            copy_stream(source, target)
        return copy

    def _rewrite_part_on_disk(self, part: str, suffix: str, rewrite: Callable[[Path, Path], None]) -> None:
        """Stage ``rewrite(source, output)`` of a part, file to file.

        ``source`` is a temporary copy of the part. The output replaces the
        part only when ``rewrite`` succeeds; both temporary files are removed
        otherwise (the source always).
        """

        source_path = self._stage_part_copy(part, suffix)
        output_path = self._package.temporary_path(suffix=suffix)
        output_staged = False
        try:
            rewrite(source_path, output_path)
            self._package.set_file(part, output_path)
            output_staged = True
        finally:
            source_path.unlink(missing_ok=True)
            if not output_staged:
                output_path.unlink(missing_ok=True)

    def replace_sheet_data_stream(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace worksheet data atomically while streaming rows.

        The package and disk-backed shared-string index are committed only
        after the one-pass row source and all metadata updates succeed.  A
        failure therefore leaves this updater reusable for a later attempt.
        """

        self._replacement_sheet_path(sheet_name, style_fallback)
        self._ensure_stream_shared_strings()
        store = self._stream_shared_strings
        self._package.begin_transaction()
        try:
            if store is not None:
                store.begin_transaction()
            self._replace_sheet_data_stream_impl(
                sheet_name,
                rows,
                headers=headers,
                style_fallback=style_fallback,
            )
            self._package.commit_transaction()
            if store is not None:
                store.commit_transaction()
        except BaseException:
            if store is not None:
                store.rollback_transaction()
            self._package.rollback_transaction()
            raise

    def save(self, output_path: str | os.PathLike[str] | None = None) -> None:
        """Save to a new path or atomically replace the input file."""

        self._package.save(output_path)

    def to_bytes(self) -> bytes:
        """Return the current package as XLSX bytes."""

        return self._package.to_bytes()

    def _shared_string_index(self, value: str) -> int:
        if self._stream_shared_strings is not None:
            return self._stream_shared_strings.get_or_add(value)
        index = self._shared_string_indices.get(value)
        if index is None:
            index = len(self._shared_strings_values)
            self._shared_strings_values.append(value)
            self._shared_string_indices[value] = index
            self._shared_strings_pending.append(value)
        self._shared_strings_total += 1
        self._shared_strings_dirty = True
        return index

    def _date_to_serial(self, value: _datetime.date | _datetime.datetime) -> float:
        if isinstance(value, _datetime.datetime):
            value = value.replace(tzinfo=None)
        return datetime_to_serial(value, self._date_1904)

    def _package_names(self) -> list[str]:
        return self._package.names()
