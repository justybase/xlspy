"""Formula-aware, incremental XLSB rows (separate from the bulk native/Python path)."""
from __future__ import annotations

import datetime
import struct
from typing import Any

from ._errors import CellError
from ._untrusted import check_column_span
from ._xlsb_formulas import FormulaContext, ERRORS, decode_formula, read_parsed
from ._xlsb_metadata import read_wide
from ._xlsx_read import rows_in_place
from .biff_reader import BiffReader
from .biff_utils import iter_record_payloads, iter_records_stream
from .excel_dates import serial_to_datetime
from .exceptions import InvalidFileError


def cached_cell(kind, payload, reader):
    if len(payload) < 8:
        raise ValueError("truncated XLSB cell header")
    style = struct.unpack_from('<I', payload, 4)[0] & 0xFFFFFF
    formula_offset = None
    if kind == 1:
        value = None
    elif kind == 2:
        value = BiffReader._get_rk_number(payload, 8)
    elif kind in (3, 11):
        value = CellError('error:' + ERRORS.get(payload[8], '#VALUE!'))
        formula_offset = 11 if kind == 11 else None
    elif kind in (4, 10):
        value = payload[8] == 1  # only 1 is True, as in the row readers
        formula_offset = 11 if kind == 10 else None
    elif kind in (5, 9):
        value = struct.unpack_from('<d', payload, 8)[0]
        formula_offset = 18 if kind == 9 else None
    elif kind in (6, 8):
        value, offset = read_wide(payload, 8)
        formula_offset = offset + 2 if kind == 8 else None
    elif kind == 7:
        index = struct.unpack_from('<I', payload, 8)[0]
        value = reader._shared_string_array[index] if index < len(reader._shared_string_array) else None
    else:
        raise ValueError(f"unknown XLSB cell kind {kind}")
    if isinstance(value, float):
        if style < len(reader._styles_cell_xfs_array) and reader._number_formats_type_dict.get(reader._styles_cell_xfs_array[style]['numFmtId']) is datetime.datetime:
            value = _date_or_number(value, reader._date_1904)
        elif value.is_integer():
            value = int(value)
    return value, payload[formula_offset:] if formula_offset is not None else None, style


# Shortest body of a value cell (BrtCellBlank..BrtCellIsst) that is read; a
# shorter one, or a string whose characters overrun its record, is ignored, as
# by the native loader and the row readers.
_VALUE_CELL_MIN = {1: 8, 2: 12, 3: 9, 4: 9, 5: 16, 6: 12, 7: 12}
# Last zero-based row: a row record past it is refused (0xFFFFFFFF: no row).
_MAX_ROW_INDEX = 1048575


def _readable_value_cell(kind: int, payload: bytes) -> bool:
    if len(payload) < _VALUE_CELL_MIN[kind]:
        return False
    if kind == 6:
        chars: int = struct.unpack_from('<I', payload, 8)[0]
        return chars * 2 + 12 <= len(payload)
    return True


def _date_or_number(serial: float, date_1904: bool) -> Any:
    """The datetime of a date-formatted serial; a serial outside the datetime
    range (NaN, infinite, past 9999) stays a number, as Excel shows it as ####."""
    try:
        return serial_to_datetime(serial, date_1904)
    except (ValueError, OverflowError):
        return int(serial) if serial.is_integer() else serial


def formula_groups(records):
    """Shared (0x1AB) and array (0x1AA) formula definitions by their master cell."""
    groups = {}
    for kind, payload in records:
        r1, r2, c1, c2 = struct.unpack_from('<IIII', payload)
        groups[(r1, c1)] = (kind, payload[17 if kind == 0x1AA else 16:])
    return groups


def formula_text(formula, groups, context, sheet_name, row, column):
    """The text (without ``=``) of a cell formula, or None for a member of an
    array formula other than its master cell, which keeps its cached value."""
    tokens, extra, _ = read_parsed(formula)
    if tokens[:1] == bytes((0x01,)):  # PtgExp: the cell belongs to a formula group
        if len(tokens) != 5 or len(extra) < 4:
            raise InvalidFileError("invalid XLSB shared/array formula reference")
        source = (struct.unpack('<I', tokens[1:])[0], struct.unpack_from('<I', extra)[0])
        if source not in groups:
            raise InvalidFileError("XLSB formula group definition was not found")
        group_kind, group_formula = groups[source]
        if group_kind == 0x1AA and (row, column) != source:
            return None
        formula = group_formula
    return decode_formula(formula, context, sheet=sheet_name, origin=(column, row))


def iter_xlsb_indexed_rows(reader, sheet_name, *, data_only=True, data=None, with_styles=False):
    """Yield ``(row_number, [(column, value), ...])`` for every BrtRowHdr, with
    1-based rows and 0-based columns; formulas are decoded when not data_only.
    With ``with_styles`` the cells are ``(column, value, style_id)``."""
    path = None
    if data is None:
        path = reader._sheet_part(sheet_name)
        if path is None:
            return  # a chartsheet: no cells
    context = FormulaContext(reader.get_sheet_names(), reader._archive.read('xl/workbook.bin'), reader.read_part) if not data_only else None
    groups: dict[Any, Any] = {}
    current_row: int | None = None
    cells: list[Any] = []
    previous_column = -1

    def cell_value(column, value, formula):
        if formula is None or data_only:
            return value
        text = formula_text(formula, groups, context, sheet_name, current_row, column)
        return value if text is None else '=' + text

    def finish():
        if current_row is None:
            raise ValueError("XLSB row end without a row")
        # Formula groups follow their master cell, so decode at row end.
        if with_styles:
            return current_row + 1, [
                (column, cell_value(column, value, formula), style)
                for column, value, formula, style in cells
            ]
        return current_row + 1, [
            (column, cell_value(column, value, formula)) for column, value, formula, _style in cells
        ]

    def scan(source):
        nonlocal current_row, cells, previous_column
        in_data = False
        for kind, payload in source:
            if kind == 0x91:
                in_data = True
            elif kind == 0x92:
                break
            elif in_data and kind == 0:
                # A row record too short for its index is ignored: its cells
                # stay in the row before (as in the native loader).
                if len(payload) < 4:
                    continue
                index = struct.unpack_from('<I', payload)[0]
                if index != 0xFFFFFFFF and index > _MAX_ROW_INDEX:
                    raise InvalidFileError(
                        f"XLSB row record has row index {index}, past the last row ({_MAX_ROW_INDEX})")
                if current_row is not None and current_row >= 0:
                    yield finish()
                current_row = -1 if index == 0xFFFFFFFF else index  # -1: no row, its cells are skipped
                cells = []
                previous_column = -1
            elif in_data and 1 <= kind <= 0x12:
                if current_row is None or current_row < 0:
                    continue
                if kind >= 0xC:
                    if len(payload) < 4:
                        continue
                    column = previous_column + 1
                    payload = struct.pack('<I', column) + payload
                    kind -= 0xB
                else:
                    if len(payload) < 8:
                        continue
                    column = struct.unpack_from('<I', payload)[0]
                if not 0 <= column < 16384:
                    raise InvalidFileError("XLSB column exceeds Excel limits")
                previous_column = column
                if kind < 8 and not _readable_value_cell(kind, payload):
                    continue
                value, formula, style = cached_cell(kind, payload, reader)
                cells.append((column, value, formula, style))
            elif in_data and kind in (0x1AA, 0x1AB) and not data_only:
                groups.update(formula_groups([(kind, payload)]))
        if current_row is not None and current_row >= 0:
            yield finish()

    if data is not None:
        yield from scan(iter_record_payloads(data))
        return
    with reader._archive.open(path) as stream:
        yield from scan(
            (record.record_id, payload) for record, payload in iter_records_stream(stream)
        )


def sheet_dimension(reader, sheet_name):
    """Return ``(min_col, min_row, max_col, max_row)`` (1-based) from BrtWsDim,
    or None when the sheet has no usable dimension record."""
    path = reader._sheet_part(sheet_name)
    if path is None:
        return None
    with reader._archive.open(path) as stream:
        for record, payload in iter_records_stream(stream):
            if record.record_id == 0x94 and len(payload) >= 16:
                r1, r2, c1, c2 = struct.unpack_from('<IIII', payload)
                # XlsbWriter before 0.11 stored an A1:A1 placeholder whatever the
                # data; a sheet that really is A1:A1 reads the same when unsized.
                if (r1, r2, c1, c2) != (0, 0, 0, 0) and r1 <= r2 < 1048576 and c1 <= c2 < 16384:
                    return c1 + 1, r1 + 1, c2 + 1, r2 + 1
                return None
            if record.record_id == 0x91:
                return None
    return None


def read_xlsb_rows(reader, sheet_name, *, data_only=True):
    """Yield rows with each value at its own position, as the bulk reader and XLSX."""
    def checked():
        for row_number, cells in iter_xlsb_indexed_rows(reader, sheet_name, data_only=data_only):
            if cells:
                check_column_span(0, max(column for column, _ in cells))
            yield row_number, cells

    return rows_in_place(checked())
