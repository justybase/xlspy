from __future__ import annotations
import struct
from enum import IntEnum

from .exceptions import InvalidFileError

# Last zero-based row Excel can hold. Row 0xFFFFFFFF reads as -1, the same as
# a row record too short to hold an index.
MAX_ROW_INDEX = 1048575
# Last zero-based column (XFD).
MAX_COLUMN_INDEX = 16383


class CellType(IntEnum):
    doubleVal = 0
    boolVal = 1
    stringVal = 2
    sharedString = 3
    nullValue = 4
    errorVal = 5


class BiffReader:
    _sheet = 0x9C

    _xf = 0x2F
    _cellXfStart = 0x269
    _cellXfEnd = 0x26A
    _cellStyleXfStart = 0x272
    _cellStyleXfEnd = 0x273
    _numberFormatStart = 0x267
    _numberFormat = 0x2C
    _numberFormatEnd = 0x268
    _sharedStringStart = 159
    _stringItem = 0x13

    _row = 0x00
    _blank = 0x01
    _number = 0x02
    _boolError = 0x03
    _bool = 0x04
    _float = 0x05
    _string = 0x06
    _sharedString = 0x07
    _formulaString = 0x08
    _formulaNumber = 0x09
    _formulaBool = 0x0A
    _formulaError = 0x0B
    _endSheetData = 0x92

    def __init__(self, stream):
        self._stream = stream
        self._buffer = bytearray(128)

        self.in_cell_xf = False
        self.in_cell_style_xf = False
        self.in_number_format = False
        self.parent_cell_style_xf = 0
        self.number_format_index = 0
        self.format = 0
        self.format_string = ""

        self.shared_string_value: str | None = None
        self.shared_string_unique_count = 0

        self.cell_type = CellType.nullValue
        self.error_code = 0
        self.int_value = 0
        self.double_val = 0.0
        self.bool_value = False
        self.string_value = ""
        self.column_num = -1
        self.xf_index = 0
        self.read_cell = False
        self.row_index = -1
        # Id of the record read_worksheet read last.
        self.record_id = -1

    def _try_read_variable_value(self):
        """``(True, value)``, or ``(False, 0)`` when the data ends before its first byte."""
        value = 0
        b = self._stream.read(1)
        if len(b) == 0:
            return False, 0

        b1 = b[0]
        value = b1 & 0x7F

        if (b1 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            raise InvalidFileError("XLSB record header is cut short")
        b2 = b[0]
        value = ((b2 & 0x7F) << 7) | value

        if (b2 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            raise InvalidFileError("XLSB record header is cut short")
        b3 = b[0]
        value = ((b3 & 0x7F) << 14) | value

        if (b3 & 0x80) == 0:
            return True, value

        b = self._stream.read(1)
        if len(b) == 0:
            raise InvalidFileError("XLSB record header is cut short")
        b4 = b[0]
        value = ((b4 & 0x7F) << 21) | value

        return True, value

    def _read_record(self):
        """The next record as ``(True, id, body)``, or ``(False, 0, None)`` at the end.

        Only the end of the data between records is a clean end; a header or body
        cut short raises InvalidFileError instead of dropping the rest silently.
        """
        ok1, record_id = self._try_read_variable_value()
        if not ok1:
            return False, 0, None
        ok2, record_length = self._try_read_variable_value()
        if not ok2:
            raise InvalidFileError("XLSB record header is cut short")
        if record_length < len(self._buffer):
            buf = self._buffer[:record_length]
        else:
            buf = bytearray(record_length)
        if record_length > 0:
            read = self._stream.readinto(buf)
            if read != record_length:
                raise InvalidFileError(
                    f"XLSB record 0x{record_id:X} is cut short: {read} of {record_length} bytes"
                )
        return True, record_id, bytes(buf)

    @staticmethod
    def _get_dword(buffer, offset):
        return (buffer[offset + 3] << 24) + (buffer[offset + 2] << 16) + (buffer[offset + 1] << 8) + buffer[offset]

    @staticmethod
    def _get_int32(buffer, offset):
        result = (buffer[offset + 3] << 24) + (buffer[offset + 2] << 16) + (buffer[offset + 1] << 8) + buffer[offset]
        return result - 0x100000000 if result >= 0x80000000 else result

    @staticmethod
    def _get_word(buffer, offset):
        return (buffer[offset + 1] << 8) + buffer[offset]

    @staticmethod
    def _get_string(buffer, offset, length):
        end = offset + 2 * length
        return buffer[offset:end].decode('utf-16-le', errors='surrogatepass')

    @staticmethod
    def _get_nullable_string(buffer, offset):
        length = BiffReader._get_dword(buffer, offset)
        offset += 4
        if length == 0xFFFFFFFF:
            return None, offset
        end = offset + length * 2
        return buffer[offset:end].decode('utf-16-le', errors='surrogatepass'), end

    @staticmethod
    def _get_rk_number(buffer, offset):
        flags = buffer[offset]
        val = struct.unpack_from('<I', buffer, offset)[0]
        if (flags & 0x02) != 0:
            result = float(BiffReader._get_int32(buffer, offset) >> 2)
        else:
            masked = val & 0xFFFFFFFC
            packed = struct.pack('<Q', masked << 32)
            result = struct.unpack('<d', packed)[0]
        if (flags & 0x01) != 0:
            result /= 100
        return result

    @staticmethod
    def _get_double(buffer, offset):
        return struct.unpack_from('<d', buffer, offset)[0]

    def read_styles(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        if record_id == self._cellXfStart:
            self.in_cell_xf = True
        elif record_id == self._cellXfEnd:
            self.in_cell_xf = False
        elif record_id == self._cellStyleXfStart:
            self.in_cell_style_xf = True
        elif record_id == self._cellStyleXfEnd:
            self.in_cell_style_xf = False
        elif record_id == self._numberFormatStart:
            self.in_number_format = True
        elif record_id == self._numberFormatEnd:
            self.in_number_format = False
        elif record_id == self._xf and self.in_cell_style_xf:
            pass
        elif record_id == self._xf and self.in_cell_xf:
            self.parent_cell_style_xf = self._get_word(buffer, 0)
            self.number_format_index = self._get_word(buffer, 2)
        elif record_id == self._numberFormat and self.in_number_format:
            self.format = self._get_word(buffer, 0)
            length = self._get_dword(buffer, 2)
            self.format_string = self._get_string(buffer, 2 + 4, length)

        return True

    def read_shared_strings(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        if record_id == self._stringItem:
            length = self._get_dword(buffer, 1)
            self.shared_string_value = self._get_string(buffer, 1 + 4, length)
        elif record_id == self._sharedStringStart:
            # Only a hint: the readers never size anything by it.
            if len(buffer) < 5:
                raise IndexError("truncated BrtBeginSst record")
            self.shared_string_unique_count = self._get_dword(buffer, 4) if len(buffer) >= 8 else 0
            self.shared_string_value = None
        else:
            self.shared_string_value = None

        return True

    def read_worksheet(self):
        ok, record_id, buffer = self._read_record()
        if not ok:
            return False

        self.read_cell = False
        self.column_num = -1
        self.record_id = record_id

        # BrtEndSheetData ends the cells: records after it are not cells of
        # the sheet (the native reader and load_workbook stop there too).
        if record_id == self._endSheetData:
            return False
        if record_id == self._row:
            if len(buffer) >= 4:
                self.row_index = self._get_int32(buffer, 0)
                # A row past Excel's last would be padded with billions of
                # empty rows by the readers, so the record is corrupt.
                if not -1 <= self.row_index <= MAX_ROW_INDEX:
                    raise InvalidFileError(
                        f"XLSB row record has row index {self.row_index & 0xFFFFFFFF}, "
                        f"past the last row ({MAX_ROW_INDEX})"
                    )
            else:
                self.row_index = -1
        elif len(buffer) < 8:
            pass
        elif record_id == self._blank:
            self.read_cell = True
            self.cell_type = CellType.nullValue
        elif record_id in (self._boolError, self._formulaError):
            # BrtCellError / BrtFmlaError: one byte after the cell header holds the error code.
            if len(buffer) >= 9:
                self.error_code = buffer[8]
                self.read_cell = True
                self.cell_type = CellType.errorVal
            else:
                self.read_cell = True
                self.cell_type = CellType.nullValue
        elif record_id == self._number:
            if len(buffer) >= 12:
                self.double_val = self._get_rk_number(buffer, 8)
                self.read_cell = True
                self.cell_type = CellType.doubleVal
        elif record_id in (self._bool, self._formulaBool):
            if len(buffer) >= 9:
                self.bool_value = (buffer[8] == 1)
                self.read_cell = True
                self.cell_type = CellType.boolVal
        elif record_id in (self._formulaNumber, self._float):
            if len(buffer) >= 16:
                self.double_val = self._get_double(buffer, 8)
                self.read_cell = True
                self.cell_type = CellType.doubleVal
        elif record_id in (self._string, self._formulaString):
            if len(buffer) >= 12:
                length = self._get_dword(buffer, 8)
                if len(buffer) >= 12 + 2 * length:
                    self.string_value = self._get_string(buffer, 12, length)
                    self.read_cell = True
                    self.cell_type = CellType.stringVal
        elif record_id == self._sharedString:
            if len(buffer) >= 12:
                self.int_value = self._get_dword(buffer, 8)
                self.read_cell = True
                self.cell_type = CellType.sharedString

        if self.read_cell:
            self.column_num = self._get_dword(buffer, 0)
            # A cell past column XFD is corrupt (load_workbook refuses it too);
            # the row readers would pad up to it with billions of None.
            if self.column_num > MAX_COLUMN_INDEX:
                raise InvalidFileError("XLSB column exceeds Excel limits")
            self.xf_index = self._get_dword(buffer, 4) & 0xFFFFFF

        return True

    def close(self):
        self._stream.close()
