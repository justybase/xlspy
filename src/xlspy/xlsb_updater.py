"""Update worksheet data in an existing XLSB package."""

from __future__ import annotations

import datetime as _datetime
import math
import numbers
import os
import re
import struct
from dataclasses import replace
from ._safe_xml import fromstring as safe_fromstring
from decimal import Decimal
from pathlib import Path
from typing import Any, BinaryIO
from collections.abc import Iterable, Iterator, Sequence
from ._accel import encode_update_xlsb_row, patch_xlsb_row_headers
from ._updater_base import StyleFallback, _UpdaterBase
from ._xlsb_edit import XlsbEditMixin
from ._xlsb_formulas import ArrayValue, FormulaContext, formula_cell
from ._xlsb_table_part import parse_table_part
from ._xlsb_sheets import XlsbSheetsMixin
from .exceptions import InvalidFileError, UnsupportedFeatureError, reading
from .formulas import ArrayFormula, DataTableFormula, Formula
from .refs import parse_cell_range

from .biff_utils import (
    Biff12Record,
    build_record,
    iter_records,
    iter_records_stream,
    iter_xlsb_sheets,
    parse_shared_strings_bin,
    xlsb_date_1904,
    read_utf16,
)
from .updater_utils import (
    TrailingRowTrimmer,
    dominant_styles,
    row_cells,
    DiskStringIndex,
    TableHeader,
    ZipPackage,
    copy_file_range,
    copy_stream,
    header_row_texts,
    table_header,
    parse_relationships_xml,
    resolve_zip_target,
    unwrap_cell,
)
from ._text import check_cell_text


_CELL_RECORDS = set(range(0x01, 0x0C))
# BrtWsDim (sheet dimension). Not BrtSel (0x98), which is the selection record.
_WS_DIM_RECORD = 0x94
_RK_INT_LOWER = -536870912
_RK_INT_UPPER = 536870911


class XlsbUpdater(_UpdaterBase, XlsbSheetsMixin, XlsbEditMixin):
    """Replace the data region of sheets in an existing XLSB file."""

    def __init__(self, path: str | os.PathLike[str]) -> None:
        self._package = ZipPackage(path)
        try:
            if self._package.has("xl/workbook.xml"):
                raise ValueError("The input is XLSX; use XlsxUpdater")
            if not self._package.has("xl/workbook.bin"):
                raise InvalidFileError("XLSB workbook.bin was not found")
            with reading(path):
                self._date_1904 = self._load_date_system()
                self._sheets = self._load_sheet_structure()
        except BaseException:
            self._package.close()
            raise
        self._shared_strings_loaded = False
        self._shared_strings_data: bytes | None = None
        self._shared_strings_values: list[str] = []
        self._shared_string_indices: dict[str, int] = {}
        self._shared_strings_total = 0
        self._shared_strings_pending: list[str] = []
        self._shared_strings_dirty = False
        self._stream_shared_strings: DiskStringIndex | None = None
        self._replace_formula_context: FormulaContext | None = None

    def _part_or_none(self, path: str) -> bytes | None:
        return self._package.get(path) if self._package.has(path) else None

    def get_worksheet_names(self) -> list[str]:
        """The names of the worksheets alone, in workbook order."""

        return [name for name in self._sheets if name not in self._chartsheets]

    def get_chartsheet_names(self) -> list[str]:
        """The chartsheet names, in workbook order."""
        return [name for name in self._sheets if name in self._chartsheets]

    def _read_table_header(self, table_part: bytes) -> TableHeader | None:
        table = parse_table_part(table_part)
        if table is None:
            return None
        return table_header(
            table.name or "",
            table.ref,
            table.header_rows != 0,
            [column.name or "" for column in table.columns],
        )

    def replace_sheet_data(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace one worksheet's used data region and preserve sheet wrappers."""

        sheet_path = self._replacement_sheet_path(sheet_name, style_fallback)
        if self._stream_shared_strings is not None:
            self.replace_sheet_data_stream(
                sheet_name,
                rows,
                headers=headers,
                style_fallback=style_fallback,
            )
            return

        materialized, header_values = self._checked_rows(sheet_name, sheet_path, rows, headers)
        sheet_data = self._package.get(sheet_path)
        self._ensure_shared_strings_loaded()
        if self._shared_strings_data is not None:
            self._shared_strings_total = max(
                0,
                self._shared_strings_total
                - self._shared_string_ref_count(sheet_data),
            )
            self._shared_strings_dirty = True

        header_styles, data_styles = self._collect_styles(
            sheet_data, headers_provided=headers is not None
        )
        date_xf = self._find_date_style() if style_fallback == "inherit" else None
        total_rows, width = self._region_size(materialized, header_values)
        last_row = max(0, total_rows - 1)
        last_col = max(0, width - 1)
        new_rows, wrote_formula, wrote_spill = self._build_rows(
            materialized,
            header_values,
            data_styles,
            header_styles,
            date_xf,
            style_fallback,
            sheet_name,
        )
        # Extra spill/array rows past the provided data extend the dimension.
        if wrote_spill or wrote_formula:
            for row in materialized:
                for raw in row:
                    value = unwrap_cell(raw)
                    ref: str | None
                    if isinstance(value, ArrayFormula):
                        ref = value.ref
                    elif isinstance(value, Formula):
                        ref = value.array_ref or value.spill_ref
                    else:
                        ref = None
                    if ref is not None:
                        _c1, _r1, c2, r2 = parse_cell_range(ref)
                        last_row = max(last_row, r2)
                        last_col = max(last_col, c2)

        region = self._find_rows_region(sheet_data)
        old_rows_length = region.rows_end - region.rows_start
        updated = bytearray(
            sheet_data[: region.rows_start]
            + new_rows
            + sheet_data[region.rows_end :]
        )
        delta = len(new_rows) - old_rows_length

        for filter_offset in region.auto_filter_ranges:
            new_offset = filter_offset
            if filter_offset >= region.rows_end:
                new_offset += delta
            if new_offset + 16 <= len(updated):
                struct.pack_into("<i", updated, new_offset + 4, last_row)
                struct.pack_into("<i", updated, new_offset + 12, last_col)

        self._patch_dimension(updated, last_row, last_col)
        self._package.set(sheet_path, bytes(updated))
        if wrote_formula:
            # New names (future functions, _xlpm params) live in the context;
            # write them back like update_cells does.
            original_workbook = self._package.get("xl/workbook.bin")
            context = self._replace_context()
            self._package.set("xl/workbook.bin", context.write_workbook(original_workbook))
            self._invalidate_replaced_formulas()
        if wrote_spill:
            self._ensure_dynamic_metadata()
        self._commit_shared_strings()
        self._patch_pivot_caches(sheet_name, last_row, last_col)

    def _replace_sheet_data_stream_impl(
        self,
        sheet_name: str,
        rows: Iterable[Sequence[Any]],
        *,
        headers: Sequence[str] | None = None,
        style_fallback: StyleFallback = "inherit",
    ) -> None:
        """Replace worksheet data while streaming rows through temporary files."""

        sheet_path = self._replacement_sheet_path(sheet_name, style_fallback)
        header_values = list(headers) if headers is not None else None
        guard = self._table_header_guard(sheet_name, sheet_path)
        original_sheet = self._stage_part_copy(sheet_path, ".bin")

        self._ensure_stream_shared_strings()
        header_styles, data_styles = self._collect_styles_stream(
            original_sheet, headers_provided=headers is not None
        )
        if self._stream_shared_strings is not None:
            self._shared_strings_total = max(
                0,
                self._stream_shared_strings.total_count
                - self._shared_string_ref_count_stream(original_sheet),
            )
            self._stream_shared_strings.total_count = self._shared_strings_total
            self._stream_shared_strings.dirty = True
        date_xf = self._find_date_style() if style_fallback == "inherit" else None

        rows_file = self._package.temporary_path(suffix=".rows")

        try:
            with rows_file.open("wb") as output:
                trimmer = TrailingRowTrimmer(output, len(header_values) if header_values is not None else 0)
                next_row = 0

                # The final column is not known until the generator ends: the row
                # headers say 0 until _patch_row_headers_file.
                writer = _XlsbRowWriter(
                    self, header_styles, data_styles, date_xf, style_fallback, 0, sheet_name
                )

                if header_values is not None:
                    header_texts = header_row_texts(header_values)
                    guard.check_row(0, header_texts)
                    output.write(writer.header(header_texts, next_row))
                    trimmer.keep()
                    next_row += 1

                for source_row in rows:
                    row = row_cells(source_row)
                    guard.check_row(next_row, row)
                    prepared = writer.prepare_row(row, next_row)
                    output.write(writer.row(prepared, next_row))
                    trimmer.wrote(prepared)
                    next_row += 1

                # Array and spill members past the source rows.
                for extra in writer.extra_row_numbers(next_row):
                    prepared = writer.prepare_row([], extra)
                    output.write(writer.row(prepared, extra))
                    trimmer.wrote(prepared)

                guard.finish()
                trimmer.finish()

            last_col = max(0, trimmer.width - 1)
            self._patch_row_headers_file(rows_file, last_col)
            region = self._find_rows_region_stream(original_sheet)
            updated_sheet = self._package.temporary_path(suffix=".bin")
            with updated_sheet.open("wb") as output:
                copy_file_range(original_sheet, output, 0, region.rows_start)
                with rows_file.open("rb") as generated_rows:
                    copy_stream(generated_rows, output)
                copy_file_range(
                    original_sheet,
                    output,
                    region.rows_end,
                    original_sheet.stat().st_size,
                )

            total_rows = trimmer.kept_rows + (1 if header_values is not None else 0)
            last_row = max(0, total_rows - 1)
            rows_length = rows_file.stat().st_size
            self._patch_sheet_file(
                updated_sheet, last_row, last_col, (region.rows_start, region.rows_start + rows_length)
            )
            self._package.set_file(sheet_path, updated_sheet)
            if writer.wrote_formula:
                original_workbook = self._package.get("xl/workbook.bin")
                context = self._replace_context()
                self._package.set("xl/workbook.bin", context.write_workbook(original_workbook))
                self._invalidate_replaced_formulas()
            if writer.wrote_spill:
                self._ensure_dynamic_metadata()
            self._commit_stream_shared_strings()
            self._patch_pivot_caches(sheet_name, last_row, last_col)
        finally:
            for temporary_path in (original_sheet, rows_file):
                try:
                    temporary_path.unlink()
                except FileNotFoundError:
                    pass

    def _load_sheet_structure(self) -> dict[str, str]:
        rels_path = "xl/_rels/workbook.bin.rels"
        if not self._package.has(rels_path):
            raise ValueError("XLSB workbook relationships were not found")
        relationships = parse_relationships_xml(self._package.get(rels_path))
        chart_ids = {
            node.get("Id") for node in safe_fromstring(self._package.get(rels_path))
            if (node.get("Type") or "").endswith("/chartsheet")
        }
        self._chartsheets: set[str] = set()
        workbook_data = self._package.get("xl/workbook.bin")
        sheets: dict[str, str] = {}

        for relationship_id, sheet_name in iter_xlsb_sheets(workbook_data):
            target = relationships.get(relationship_id or "")
            if target is None:
                raise InvalidFileError(f"Relationship {relationship_id!r} for sheet is missing")
            sheet_path = resolve_zip_target(target, "xl")
            if not self._package.has(sheet_path):
                raise ValueError(f"Worksheet part was not found: {sheet_path}")
            sheets[sheet_name] = sheet_path
            if relationship_id in chart_ids:
                self._chartsheets.add(sheet_name)

        if not sheets:
            raise InvalidFileError("No worksheets were found in workbook.bin")
        return sheets

    def _load_date_system(self) -> bool:
        return xlsb_date_1904(self._package.get("xl/workbook.bin"))

    def _ensure_shared_strings_loaded(self) -> None:
        if self._shared_strings_loaded:
            return
        self._shared_strings_loaded = True
        if not self._package.has("xl/sharedStrings.bin"):
            return
        data = self._package.get("xl/sharedStrings.bin")
        parsed = parse_shared_strings_bin(data)
        self._shared_strings_data = data
        self._shared_strings_values = parsed.values
        self._shared_string_indices = {
            value: index for index, value in enumerate(parsed.values)
        }
        self._shared_strings_total = parsed.total_count

    def _ensure_stream_shared_strings(self) -> None:
        if self._stream_shared_strings is not None:
            return
        if not self._package.has("xl/sharedStrings.bin"):
            return

        store = DiskStringIndex(self._package.temporary_directory)
        total_count: int | None = None
        value_count = 0
        with self._package.open_entry("xl/sharedStrings.bin") as source:
            for record, payload in iter_records_stream(source):
                if record.record_id == 0x9F:
                    if len(payload) < 8:
                        raise ValueError("Malformed XLSB shared-string header")
                    total_count = struct.unpack_from("<I", payload, 0)[0]
                elif record.record_id == 0x13:
                    if len(payload) < 5:
                        raise ValueError("Malformed XLSB shared-string item")
                    character_count = struct.unpack_from("<I", payload, 1)[0]
                    value, string_end = read_utf16(payload, 5, character_count)
                    if string_end > len(payload):
                        raise ValueError("Malformed XLSB shared-string item length")
                    store.add_existing(value, value_count)
                    value_count += 1

        store.total_count = total_count if total_count is not None else value_count
        self._stream_shared_strings = store

    def _commit_stream_shared_strings(self) -> None:
        store = self._stream_shared_strings
        if store is None or not store.dirty:
            return

        def write_pending(target: BinaryIO) -> None:
            for value in store.pending_values():
                encoded = value.encode("utf-16le")
                target.write(
                    build_record(
                        0x13,
                        b"\x00" + struct.pack("<I", len(encoded) // 2) + encoded,
                    )
                )

        def rewrite(source_path: Path, output_path: Path) -> None:
            # New items go before BrtEndSst (0xA0); the BrtBeginSst counts are updated.
            inserted = False
            with source_path.open("rb") as source, output_path.open("wb") as target:
                for record, payload in iter_records_stream(source):
                    if record.record_id == 0xA0 and not inserted:
                        write_pending(target)
                        inserted = True

                    if record.record_id == 0x9F and len(payload) >= 8:
                        header_length = record.data_start - record.header_start
                        source.seek(record.header_start)
                        target.write(source.read(header_length))
                        updated_payload = bytearray(payload)
                        struct.pack_into(
                            "<II",
                            updated_payload,
                            0,
                            store.total_count,
                            store.unique_count,
                        )
                        target.write(updated_payload)
                        source.seek(record.data_end)
                    else:
                        copy_file_range(
                            source_path,
                            target,
                            record.header_start,
                            record.data_end,
                        )

                if not inserted:
                    write_pending(target)

        self._rewrite_part_on_disk("xl/sharedStrings.bin", ".bin", rewrite)
        store.clear_pending()

    def _collect_styles_stream(
        self, sheet_path: str | os.PathLike[str], *, headers_provided: bool
    ) -> tuple[dict[int, int], dict[int, int]]:
        with open(sheet_path, "rb") as stream:
            return _record_cell_styles(iter_records_stream(stream), headers_provided=headers_provided)

    @staticmethod
    def _shared_string_ref_count_stream(sheet_path: str | os.PathLike[str]) -> int:
        with open(sheet_path, "rb") as stream:
            return _record_shared_string_cells(iter_records_stream(stream))

    def _commit_shared_strings(self) -> None:
        if self._shared_strings_data is None or (
            not self._shared_strings_pending and not self._shared_strings_dirty
        ):
            return

        new_items = bytearray()
        for value in self._shared_strings_pending:
            encoded = value.encode("utf-16le")
            payload = b"\x00" + struct.pack("<I", len(encoded) // 2) + encoded
            new_items.extend(build_record(0x13, payload))

        parsed = parse_shared_strings_bin(self._shared_strings_data)
        insertion = parsed.end_offset if parsed.end_offset is not None else len(self._shared_strings_data)
        updated = bytearray(
            self._shared_strings_data[:insertion]
            + bytes(new_items)
            + self._shared_strings_data[insertion:]
        )
        for record in iter_records(bytes(updated)):
            if record.record_id == 0x9F and record.length >= 8:
                struct.pack_into("<II", updated, record.data_start, self._shared_strings_total, len(self._shared_strings_values))
                break

        self._shared_strings_data = bytes(updated)
        self._shared_strings_pending.clear()
        self._shared_strings_dirty = False
        self._package.set("xl/sharedStrings.bin", self._shared_strings_data)

    @staticmethod
    def _shared_string_ref_count(sheet_data: bytes) -> int:
        return _record_shared_string_cells(_with_payloads(sheet_data))

    def _collect_styles(
        self, sheet_data: bytes, *, headers_provided: bool
    ) -> tuple[dict[int, int], dict[int, int]]:
        return _record_cell_styles(_with_payloads(sheet_data), headers_provided=headers_provided)

    def _cell_xf_number_formats(self) -> tuple[list[int], dict[int, str]] | None:
        if not self._package.has("xl/styles.bin"):
            return None
        styles_data = self._package.get("xl/styles.bin")
        in_cell_xfs = False
        xf_formats: list[int] = []
        custom_formats: dict[int, str] = {}
        in_number_formats = False
        for record in iter_records(styles_data):
            payload = styles_data[record.data_start : record.data_end]
            if record.record_id == 0x267:
                in_number_formats = True
                continue
            if record.record_id == 0x268:
                in_number_formats = False
                continue
            if in_number_formats and record.record_id == 0x2C and record.length >= 6:
                number_format_id = struct.unpack_from("<H", payload, 0)[0]
                character_count = struct.unpack_from("<I", payload, 2)[0]
                format_end = 6 + character_count * 2
                if format_end <= record.length:
                    custom_formats[number_format_id] = payload[6:format_end].decode(
                        "utf-16le"
                    )
                continue
            if record.record_id == 0x269:
                in_cell_xfs = True
                continue
            if record.record_id == 0x26A:
                in_cell_xfs = False
                continue
            if in_cell_xfs and record.record_id == 0x2F and record.length >= 4:
                xf_formats.append(struct.unpack_from("<H", payload, 2)[0])
        return xf_formats, custom_formats

    def _build_rows(
        self,
        rows: list[list[Any]],
        headers: list[Any] | None,
        data_styles: dict[int, int],
        header_styles: dict[int, int],
        date_xf: int | None,
        style_fallback: StyleFallback,
        sheet_name: str,
    ) -> tuple[bytes, bool, bool]:
        chunks: list[bytes] = []
        max_col = max(
            [len(row) for row in rows] + ([len(headers)] if headers is not None else []) + [0]
        )
        first_row = 1 if headers is not None else 0
        # A spill or array formula widens its row even when the row source is
        # shorter; the row headers must name that last column.
        for row in rows:
            for raw in row:
                value = unwrap_cell(raw)
                ref: str | None
                if isinstance(value, ArrayFormula):
                    ref = value.ref
                elif isinstance(value, Formula):
                    ref = value.array_ref or value.spill_ref
                else:
                    ref = None
                if ref is not None:
                    _c1, _r1, c2, _r2 = parse_cell_range(ref)
                    max_col = max(max_col, c2 + 1)
        writer = _XlsbRowWriter(
            self, header_styles, data_styles, date_xf, style_fallback,
            max(0, max_col - 1), sheet_name,
        )

        if headers is not None:
            header_texts = [
                "" if unwrap_cell(value) is None else str(unwrap_cell(value))
                for value in headers
            ]
            chunks.append(writer.header(header_texts, 0))

        row_number = first_row
        for row in rows:
            chunks.append(writer.row(row, row_number))
            row_number += 1
        # Array and spill ranges that reach past the provided rows still need
        # their member cells, as the workbook writer emits them.
        for extra in writer.extra_row_numbers(row_number):
            chunks.append(writer.row([], extra))
        return b"".join(chunks), writer.wrote_formula, writer.wrote_spill

    def _replace_context(self) -> FormulaContext:
        """The formula context of a ``replace_sheet_data`` call, built once."""
        if self._replace_formula_context is None:
            self._replace_formula_context = self._context(
                self._package.get("xl/workbook.bin"), self._part_or_none
            )
        return self._replace_formula_context

    def _invalidate_replaced_formulas(self) -> None:
        """Drop the calculation chain and ask Excel to recalculate on load.

        The parts are set one by one: the streaming path already holds a package
        transaction, and the in-memory path writes them straight through.
        """
        updates: dict[str, Any] = {}
        removed: list[str] = []
        self._invalidate_calculation(updates, removed)
        for path, data in updates.items():
            self._package.set(path, data.encode("utf-8") if isinstance(data, str) else data)
        for path in removed:
            self._package.remove(path)
        # workbook.bin changed (the recalculation flag), so a cached formula
        # context no longer describes the package.
        self._replace_formula_context = None

    def _value_cell_bytes(
        self,
        raw: Any,
        column: int,
        data_styles: dict[int, int],
        date_xf: int | None,
        style_fallback: StyleFallback,
        *,
        sheet_name: str = "",
        row_number: int = 0,
    ) -> bytes:
        value = unwrap_cell(raw)
        if isinstance(value, ArrayFormula):
            value = Formula(value.text, array_ref=value.ref)
        elif isinstance(value, DataTableFormula):
            raise UnsupportedFeatureError(
                "an XLSB workbook cannot store a data-table formula; write it to XLSX"
            )
        style = 0 if style_fallback == "general" else data_styles.get(column, 0)
        if isinstance(value, ArrayValue):
            return formula_cell(
                value, column, style, self._replace_context(),
                sheet=sheet_name, row=row_number,
            )
        if isinstance(value, Formula):
            spill = value.spill_ref is not None and value.array_ref is None
            if spill:
                assert value.spill_ref is not None
                # Dynamic array: BrtCellMeta + array formula group, like the writer.
                # Tracked by the row writer for metadata ensuring.
                converted = replace(value, array_ref=value.spill_ref, spill_ref=None)
                return build_record(0x31, struct.pack("<I", 1)) + formula_cell(
                    converted, column, style, self._replace_context(),
                    sheet=sheet_name, row=row_number,
                )
            return formula_cell(
                value, column, style, self._replace_context(),
                sheet=sheet_name, row=row_number,
            )
        if value is None:
            return self._string_cell_bytes(column, style, "")
        # numpy and other numeric types register with numbers; write them as Python numbers.
        if isinstance(value, numbers.Integral) and not isinstance(value, bool):
            value = int(value)
        elif isinstance(value, numbers.Real) and not isinstance(value, (float, bool)):
            value = float(value)
        if isinstance(value, bool):
            payload = struct.pack("<IIB", column, style, 1 if value else 0)
            return build_record(0x04, payload)
        if isinstance(value, int) and not isinstance(value, bool):
            if _RK_INT_LOWER <= value <= _RK_INT_UPPER:
                payload = struct.pack("<IIi", column, style, (value << 2) | 2)
                return build_record(0x02, payload)
            payload = struct.pack("<IId", column, style, float(value))
            return build_record(0x05, payload)
        if isinstance(value, (float, Decimal)):
            numeric_value = float(value)
            if not math.isfinite(numeric_value):
                raise ValueError("numeric cell values must be finite")
            payload = struct.pack("<IId", column, style, numeric_value)
            return build_record(0x05, payload)
        if isinstance(value, (_datetime.datetime, _datetime.date)):
            serial = self._date_to_serial(value)
            if math.isfinite(serial):
                date_style = style if style > 0 else (date_xf or 0)
                payload = struct.pack("<IId", column, date_style, serial)
                return build_record(0x05, payload)
        return self._string_cell_bytes(column, style, str(value))

    def _string_cell_bytes(self, column: int, style: int, text: str) -> bytes:
        # XLSB stores text as UTF-16 records, not XML: control characters are kept,
        # as the XLSB writer keeps them.
        check_cell_text(text)
        if self._shared_strings_data is None and self._stream_shared_strings is None:
            encoded = text.encode("utf-16le")
            payload = struct.pack("<III", column, style, len(encoded) // 2) + encoded
            return build_record(0x06, payload)
        index = self._shared_string_index(text)
        return build_record(0x07, struct.pack("<III", column, style, index))

    @staticmethod
    def _patch_row_headers_file(path: str | os.PathLike[str], last_col: int) -> None:
        if patch_xlsb_row_headers is not None:
            patch_xlsb_row_headers(os.fspath(path), last_col)
            return
        with open(path, "r+b") as stream:
            for record, _ in iter_records_stream(stream):
                if record.record_id == 0x00 and record.length >= 25:
                    stream.seek(record.data_start + 21)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)

    @staticmethod
    def _rows_region(records: Iterable[Any], size: int) -> _RowsRegion:
        """Where a sheet's rows are, from its records: what new rows replace. A sheet
        without rows gets them at BrtEndSheetData; records past it are not cells, and
        the wrapper records there belong to other parts of the sheet."""
        rows_start = -1
        last_cell_end = -1
        last_row_header_end = -1
        last_wrapper = -1
        end_of_data = -1
        auto_filter_ranges: list[int] = []
        for record in records:
            kind = record.record_id
            if kind == 0xA1 and record.length >= 16:
                auto_filter_ranges.append(record.data_start)
            elif end_of_data >= 0:
                continue
            elif kind == 0x00:
                if rows_start == -1:
                    rows_start = record.header_start
                last_row_header_end = record.data_end
            elif 0x01 <= kind <= 0x12:
                last_cell_end = record.data_end
            elif kind == 0x25 and record.length == 6:
                last_wrapper = record.header_start
            elif kind == 0x92:
                end_of_data = record.header_start

        if rows_start == -1:
            insert_at = end_of_data if end_of_data >= 0 else (last_wrapper if last_wrapper >= 0 else size)
            return _RowsRegion(insert_at, insert_at, auto_filter_ranges)
        return _RowsRegion(rows_start, max(last_cell_end, last_row_header_end), auto_filter_ranges)

    @classmethod
    def _find_rows_region_stream(cls, sheet_path: str | os.PathLike[str]) -> _RowsRegion:
        with open(sheet_path, "rb") as stream:
            return cls._rows_region((record for record, _ in iter_records_stream(stream)), os.path.getsize(sheet_path))

    @staticmethod
    def _patch_sheet_file(
        sheet_path: str | os.PathLike[str], last_row: int, last_col: int,
        rows: tuple[int, int] | None = None,
    ) -> None:
        """Patch the dimension and the autofilter ranges; ``rows`` (start, end)
        is the region of the new rows, which holds neither and is skipped."""
        dimension_patched = False
        with open(sheet_path, "r+b") as stream:
            for record, _ in _records_around(stream, rows):
                if (
                    record.record_id == _WS_DIM_RECORD
                    and record.length >= 16
                    and not dimension_patched
                ):
                    # BrtWsDim: rwFirst, rwLast, colFirst, colLast (int32 each).
                    stream.seek(record.data_start + 4)
                    stream.write(struct.pack("<i", last_row))
                    stream.seek(record.data_start + 12)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)
                    dimension_patched = True
                elif record.record_id == 0xA1 and record.length >= 16:
                    stream.seek(record.data_start + 4)
                    stream.write(struct.pack("<i", last_row))
                    stream.seek(record.data_start + 12)
                    stream.write(struct.pack("<i", last_col))
                    stream.seek(record.data_end)

    def _find_rows_region(self, sheet_data: bytes) -> _RowsRegion:
        return self._rows_region(iter_records(sheet_data), len(sheet_data))

    @staticmethod
    def _patch_dimension(sheet_data: bytearray, last_row: int, last_col: int) -> None:
        for record in iter_records(bytes(sheet_data)):
            if record.record_id == _WS_DIM_RECORD and record.length >= 16:
                # BrtWsDim: rwFirst, rwLast, colFirst, colLast (int32 each).
                struct.pack_into("<i", sheet_data, record.data_start + 4, last_row)
                struct.pack_into("<i", sheet_data, record.data_start + 12, last_col)
                return

    def _patch_pivot_caches(self, sheet_name: str, last_row: int, last_col: int) -> None:
        for path in sorted(
            name
            for name in self._package_names()
            if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d*\.bin", name)
        ):
            data = bytearray(self._package.get(path))
            patched = False
            source_matched = False
            for record in iter_records(bytes(data)):
                payload = data[record.data_start : record.data_end]
                if record.record_id == 0xBB and record.length >= 23:
                    if len(payload) < 7:
                        continue
                    character_count = struct.unpack_from("<I", payload, 3)[0]
                    name_start = record.data_start + 7
                    reference_start = name_start + character_count * 2
                    if reference_start + 16 > record.data_end:
                        continue
                    source_sheet, _ = read_utf16(bytes(data), name_start, character_count)
                    if source_sheet == sheet_name:
                        struct.pack_into("<i", data, reference_start + 4, last_row)
                        struct.pack_into("<i", data, reference_start + 12, last_col)
                        patched = True
                        source_matched = True
            if source_matched:
                for record in iter_records(bytes(data)):
                    if record.record_id == 0xB3 and record.length >= 4:
                        data[record.data_start + 3] |= 0x04
                        break
            if patched:
                self._package.set(path, bytes(data))


class _XlsbRowWriter:
    """The records of a replacement's rows, natively when the extension is
    built (``encode_update_xlsb_row``): the common cells there, any other cell
    through ``XlsbUpdater._value_cell_bytes``, one at a time. Row headers name
    ``last_col`` as the last column."""

    def __init__(self, updater: XlsbUpdater, header_styles: dict[int, int], data_styles: dict[int, int],
                 date_xf: int | None, style_fallback: StyleFallback, last_col: int,
                 sheet_name: str = "") -> None:
        self._updater = updater
        self._header_styles = header_styles
        self._data_styles = data_styles
        self._date_xf = date_xf
        self._style_fallback = style_fallback
        self._last_col = last_col
        self._sheet_name = sheet_name
        # Set when a Formula cell is encoded: the save then asks for a recalculation.
        self.wrote_formula = False
        self.wrote_spill = False
        # Array and spill ranges started in the rows written so far, and the
        # anchor cells that started them (prepare_row may run twice per row).
        self.array_ranges: list[tuple[int, int, int, int]] = []
        self._array_anchors: set[tuple[int, int]] = set()
        general = style_fallback == "general"
        # The native encoder takes the column xfs as a list (none for "general").
        self._style_lists = tuple(
            [] if general else [styles.get(column, 0) for column in range(max(styles, default=-1) + 1)]
            for styles in (header_styles, data_styles)
        )
        shared = updater._shared_strings_data is not None or updater._stream_shared_strings is not None
        self._string_index = updater._shared_string_index if shared else None

    def _row_header(self, row_number: int) -> bytes:
        payload = bytearray(25)
        struct.pack_into("<i", payload, 0, row_number)
        struct.pack_into("<i", payload, 8, 300)
        payload[13] = 1
        struct.pack_into("<i", payload, 17, 0)
        struct.pack_into("<i", payload, 21, self._last_col)
        return build_record(0x00, bytes(payload))

    def _cell(self, raw: Any, column: int, row_number: int) -> bytes:
        unwrapped = unwrap_cell(raw)
        if isinstance(unwrapped, ArrayFormula):
            self.wrote_formula = True
            self._register_array_range(
                Formula(unwrapped.text, array_ref=unwrapped.ref), column, row_number
            )
        elif isinstance(unwrapped, Formula):
            self.wrote_formula = True
            if unwrapped.spill_ref is not None:
                self.wrote_spill = True
        return self._updater._value_cell_bytes(
            raw, column, self._data_styles, self._date_xf, self._style_fallback,
            sheet_name=self._sheet_name, row_number=row_number,
        )

    def _register_array_range(
        self, formula: Formula, column: int, row_number: int
    ) -> None:
        ref = formula.array_ref or formula.spill_ref
        if ref is None:
            return
        bounds = parse_cell_range(ref)
        c1, r1, c2, r2 = bounds
        if formula.array_ref is not None and (column, row_number) != (c1, r1):
            raise ValueError("array formula must be written in its top-left cell")
        if (column, row_number) in self._array_anchors:
            return
        self._array_anchors.add((column, row_number))
        for a, b, c, d in self.array_ranges:
            if not (c2 < a or c1 > c or r2 < b or r1 > d):
                raise ValueError("overlapping array formulas")
        self.array_ranges.append(bounds)

    def prepare_row(self, row: Sequence[Any], row_number: int) -> list[Any]:
        """Add the member cells an array or spill range needs to ``row``.

        The workbook writer does the same: cells inside a range become
        ``ArrayValue`` followers (the anchor stays a formula), so Excel opens a
        file whose spills are not lost when the source row is shorter or ends
        before the range does.
        """
        prepared = list(row)
        for column, raw in enumerate(prepared):
            value = unwrap_cell(raw)
            if isinstance(value, ArrayFormula):
                self._register_array_range(
                    Formula(value.text, array_ref=value.ref), column, row_number
                )
            elif isinstance(value, Formula):
                self._register_array_range(value, column, row_number)
        for c1, r1, c2, r2 in self.array_ranges:
            if not r1 <= row_number <= r2:
                continue
            if len(prepared) <= c2:
                prepared.extend([None] * (c2 + 1 - len(prepared)))
            for column in range(c1, c2 + 1):
                if row_number == r1 and column == c1:
                    continue  # the anchor formula stays in place
                value = unwrap_cell(prepared[column]) if prepared[column] is not None else None
                if isinstance(value, (Formula, ArrayFormula, DataTableFormula)):
                    raise ValueError("overlapping array formulas")
                if isinstance(value, ArrayValue):
                    continue
                prepared[column] = ArrayValue(r1, c1, value)
        return prepared

    def extra_row_numbers(self, next_row: int) -> Iterator[int]:
        """Rows from ``next_row`` through the last member of an array range."""
        if not self.array_ranges:
            return
        final_row = max(r2 for _c1, _r1, _c2, r2 in self.array_ranges)
        yield from range(next_row, final_row + 1)

    def _header_cell(self, text: Any, column: int) -> bytes:
        style = 0 if self._style_fallback == "general" else self._header_styles.get(column, 0)
        return self._updater._string_cell_bytes(column, style, text)

    def header(self, texts: list[str], row_number: int) -> bytes:
        if encode_update_xlsb_row is not None:
            return bytes(encode_update_xlsb_row(
                texts, row_number, self._last_col, self._style_lists[0], 0, self._string_index,
                self._header_cell, self._updater._date_1904,
            ))
        return self._row_header(row_number) + b"".join(
            self._header_cell(text, column) for column, text in enumerate(texts)
        )

    def row(self, row: Sequence[Any], row_number: int) -> bytes:
        prepared = self.prepare_row(row, row_number)

        def cell(raw: Any, column: int, _row_number: int = row_number) -> bytes:
            return self._cell(raw, column, _row_number)

        if encode_update_xlsb_row is not None:
            return bytes(encode_update_xlsb_row(
                prepared, row_number, self._last_col, self._style_lists[1], self._date_xf or 0,
                self._string_index, cell, self._updater._date_1904,
            ))
        return self._row_header(row_number) + b"".join(
            self._cell(raw, column, row_number) for column, raw in enumerate(prepared) if raw is not None
        )


def _records_around(stream: BinaryIO, rows: tuple[int, int] | None) -> Iterator[tuple[Biff12Record, bytes]]:
    """The records of a sheet stream, but for those in ``rows`` (start, end)."""
    if rows is None:
        yield from iter_records_stream(stream)
        return
    start, end = rows
    for item in iter_records_stream(stream):
        if item[0].header_start >= start:
            break
        yield item
    yield from iter_records_stream(stream, end)


class _RowsRegion:
    def __init__(self, rows_start: int, rows_end: int, auto_filter_ranges: list[int]) -> None:
        self.rows_start = rows_start
        self.rows_end = rows_end
        self.auto_filter_ranges = auto_filter_ranges


def _with_payloads(data: bytes) -> Iterator[tuple[Biff12Record, bytes]]:
    """The records of ``data`` with their payloads, as ``iter_records_stream`` yields them."""
    for record in iter_records(data):
        yield record, data[record.data_start : record.data_end]


def _record_shared_string_cells(records: Iterable[tuple[Biff12Record, bytes]]) -> int:
    """The number of BrtCellIsst (shared-string) cells."""
    return sum(1 for record, _ in records if record.record_id == 0x07)


def _record_cell_styles(
    records: Iterable[tuple[Biff12Record, bytes]], *, headers_provided: bool
) -> tuple[dict[int, int], dict[int, int]]:
    """Per column: the first style of the header row and the dominant data style."""
    header_styles: dict[int, int] = {}
    data_counts: dict[int, dict[int, int]] = {}
    row = -1
    for record, payload in records:
        if record.record_id == 0x00 and record.length >= 4:
            row = struct.unpack_from("<i", payload, 0)[0]
            continue
        if record.record_id not in _CELL_RECORDS or record.length < 8:
            continue
        column, style = struct.unpack_from("<II", payload, 0)
        style &= 0xFFFFFF
        if style <= 0:
            continue
        if headers_provided and row == 0:
            header_styles.setdefault(column, style)
        elif row >= 0:
            counts = data_counts.setdefault(column, {})
            counts[style] = counts.get(style, 0) + 1
    return header_styles, dominant_styles(data_counts)
