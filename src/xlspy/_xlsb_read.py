"""Formula-aware, incremental XLSB rows (separate from the bulk C/Python path)."""

import datetime
import struct

from ._xlsb_formulas import FormulaContext, ERRORS, decode_formula, read_parsed
from ._xlsb_metadata import read_wide
from .biff_reader import BiffReader
from .biff_utils import iter_records_stream
from .excel_dates import serial_to_datetime


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
        value = 'error:' + ERRORS.get(payload[8], '#VALUE!')
        formula_offset = 11 if kind == 11 else None
    elif kind in (4, 10):
        value = bool(payload[8])
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
            value = serial_to_datetime(value, reader._date_1904)
        elif value.is_integer():
            value = int(value)
    return value, payload[formula_offset:] if formula_offset is not None else None


def iter_indexed_rows(reader, sheet_name, *, data_only=True):
    """Yield ``(row_number, [(column, value), ...])`` for every BrtRowHdr, with
    1-based rows and 0-based columns; formulas are decoded when not data_only."""
    rid = reader._worksheet_name_to_id[sheet_name]
    path = reader._resolve_path(reader._worksheet_id_to_location[rid])
    context = FormulaContext(reader.get_sheet_names(), reader._archive.read('xl/workbook.bin')) if not data_only else None
    groups = {}
    current_row = None
    cells = []
    previous_column = -1

    def cell_value(column, value, formula):
        if formula is None or data_only:
            return value
        tokens, extra, _ = read_parsed(formula)
        if tokens[:1] == b'':
            if len(tokens) != 5 or len(extra) < 4:
                raise ValueError("invalid XLSB shared/array formula reference")
            source = (struct.unpack('<I', tokens[1:])[0], struct.unpack_from('<I', extra)[0])
            if source not in groups:
                raise NotImplementedError("XLSB formula group definition was not found")
            group_kind, group_formula = groups[source]
            if group_kind == 0x1AA and (current_row, column) != source:
                return value
            formula = group_formula
        return '=' + decode_formula(formula, context, sheet=sheet_name, origin=(column, current_row))

    def finish():
        # Formula groups follow their master cell, so decode at row end.
        return current_row + 1, [(column, cell_value(column, value, formula)) for column, value, formula in cells]

    with reader._archive.open(path) as stream:
        in_data = False
        for record, payload in iter_records_stream(stream):
            kind = record.record_id
            if kind == 0x91:
                in_data = True
            elif kind == 0x92:
                break
            elif in_data and kind == 0:
                if current_row is not None:
                    yield finish()
                current_row = struct.unpack_from('<I', payload)[0]
                cells = []
                previous_column = -1
            elif in_data and 1 <= kind <= 0x12:
                if kind >= 0xC:
                    column = previous_column + 1
                    payload = struct.pack('<I', column) + payload
                    kind -= 0xB
                else:
                    column = struct.unpack_from('<I', payload)[0]
                if not 0 <= column < 16384:
                    raise ValueError("XLSB column exceeds Excel limits")
                previous_column = column
                value, formula = cached_cell(kind, payload, reader)
                cells.append((column, value, formula))
            elif in_data and kind in (0x1AA, 0x1AB) and not data_only:
                r1, r2, c1, c2 = struct.unpack_from('<IIII', payload)
                groups[(r1, c1)] = (kind, payload[17 if kind == 0x1AA else 16:])
        if current_row is not None:
            yield finish()


def sheet_dimension(reader, sheet_name):
    """Return ``(min_col, min_row, max_col, max_row)`` (1-based) from BrtWsDim,
    or None when the sheet has no usable dimension record."""
    rid = reader._worksheet_name_to_id[sheet_name]
    path = reader._resolve_path(reader._worksheet_id_to_location[rid])
    with reader._archive.open(path) as stream:
        for record, payload in iter_records_stream(stream):
            if record.record_id == 0x94 and len(payload) >= 16:
                r1, r2, c1, c2 = struct.unpack_from('<IIII', payload)
                if r1 <= r2 < 1048576 and c1 <= c2 < 16384:
                    return c1 + 1, r1 + 1, c2 + 1, r2 + 1
                return None
            if record.record_id == 0x91:
                return None
    return None


def iter_rows(reader, sheet_name, *, data_only=True):
    """Yield rows shaped like the bulk reader: the first data row fixes the first
    column, missing and empty rows are synthesized, and later wider rows widen."""
    first_row_record = None
    last_row_record = None
    first_column = None
    width = 0
    previous_row = None
    for row_number, cells in iter_indexed_rows(reader, sheet_name, data_only=data_only):
        if first_row_record is None:
            first_row_record = row_number
        last_row_record = row_number
        if first_column is None:
            if not cells:
                continue
            first_column = cells[0][0]
            width = cells[-1][0] - first_column + 1
            # Leading row records without cells become empty rows.
            for _ in range(row_number - first_row_record):
                yield [None] * width
        else:
            for _ in range(row_number - previous_row - 1):
                yield [None] * width
        previous_row = row_number
        if cells and cells[-1][0] - first_column + 1 > width:
            width = cells[-1][0] - first_column + 1
        row = [None] * width
        for column, value in cells:
            if column >= first_column:
                row[column - first_column] = value
        yield row
    if first_column is None and first_row_record is not None:
        # Only all-empty row records, as in the bulk reader.
        for _ in range(last_row_record - first_row_record + 1):
            yield [None]
