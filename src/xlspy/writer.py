from dataclasses import dataclass
import datetime
from decimal import Decimal
import zipfile
import struct
from typing import  Tuple, Iterable
import io
import itertools
from ._accel import encode_xlsb_row, calc_column_widths as accel_calc_column_widths
from .formats import _is_formatted_cell, _unwrap_cell, _get_format
from .excel_dates import datetime_to_serial
from .styles import Style, merge_styles, normalize_style_ranges, styles_for_row


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0  
    start_row: int = 0
    end_row: int = 0


class XlsbWriter:
    def __init__(self, file: str | io.IOBase, compressionLevel:int = 4):
        self._file = file
        self._worksheet_data: list[Tuple[str, Iterable[list[any]],bool]] = []
        self._shared_strings: list[str] = []
        self._shared_strings_dict: dict[str, int] = {}
        self._sheet_count = 0
        self._sst_unique_count = 0
        self._sst_all_count = 0
        self._filtered_data_list: list[FilterData] = []
        self._sheetCnt = 1
        self._compressionLevel = compressionLevel
        self._zf: zipfile.ZipFile = None

        self._format_registry: dict[str, int] = {}
        self._format_xf_map: dict[str, int] = {}
        self._style_xf_map: dict[Style, int] = {}
        self._style_components: dict[int, tuple[int, int, int, Style]] = {}
        self._font_registry: dict[tuple, int] = {(None, True, None, None): 1}
        self._fill_registry: dict[str, int] = {}
        self._next_numfmt_id = 167
        self._next_xf_index = 4

    def __enter__(self):
        if self._zf is None:
            self._zf = zipfile.ZipFile(self._file, 'w', zipfile.ZIP_DEFLATED, compresslevel=self._compressionLevel)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self._zf is not None:
                self.save()
        finally:
            if self._zf is not None:
                self._zf.close()
                self._zf = None
        
        return False

    def add_sheet(self, sheet_name: str, hidden : bool = False):
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]),hidden))

    def _register_format(self, fmt_string: str) -> int:
        if fmt_string in self._format_xf_map:
            return self._format_xf_map[fmt_string]
        # A style may already have registered this number format; reuse its id
        # so the style XF keeps referencing a format that is still emitted.
        numfmt_id = self._format_registry.get(fmt_string)
        if numfmt_id is None:
            numfmt_id = self._next_numfmt_id
            self._next_numfmt_id += 1
            self._format_registry[fmt_string] = numfmt_id
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._format_xf_map[fmt_string] = xf_index
        return xf_index

    def _register_style(self, style: Style) -> int:
        if style in self._style_xf_map:
            return self._style_xf_map[style]
        if any(value is not None for value in (style.horizontal, style.vertical, style.wrap)):
            raise NotImplementedError(
                "horizontal/vertical alignment and wrapping are not yet supported for XLSB output"
            )
        font_key = (style.font_color, style.bold, style.italic, style.underline)
        if font_key == (None, None, None, None):
            font_id = 0
        elif font_key not in self._font_registry:
            self._font_registry[font_key] = max(self._font_registry.values(), default=1) + 1
            font_id = self._font_registry[font_key]
        else:
            font_id = self._font_registry[font_key]
        fill_id = 0
        if style.fill is not None:
            if style.fill not in self._fill_registry:
                self._fill_registry[style.fill] = 2 + len(self._fill_registry)
            fill_id = self._fill_registry[style.fill]
        numfmt_id = 0
        if style.number_format is not None:
            if style.number_format not in self._format_registry:
                self._format_registry[style.number_format] = self._next_numfmt_id
                self._next_numfmt_id += 1
            numfmt_id = self._format_registry[style.number_format]
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._style_xf_map[style] = xf_index
        self._style_components[xf_index] = (font_id, fill_id, numfmt_id, style)
        return xf_index

    def write_sheet(
        self,
        data: Iterable[list[any]],
        *,
        header_style: Style | None = None,
        styles: dict[str, Style] | None = None,
    ):
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        style_ranges = normalize_style_ranges(styles)
        for candidate in ([header_style] if header_style is not None else []) + [item[-1] for item in style_ranges]:
            if any(value is not None for value in (candidate.horizontal, candidate.vertical, candidate.wrap)):
                raise NotImplementedError(
                    "horizontal/vertical alignment and wrapping are not yet supported for XLSB output"
                )
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        
        sheet_name,_, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)

        sheet_id = self._sheet_count
        with self._zf.open(f"xl/worksheets/sheet{sheet_id}.bin", 'w') as sheet_file:
            self._write_worksheet_bin(
                sheet_file, data, self._sheet_count - 1,
                style_ranges=style_ranges, header_style=header_style,
            )
        
        self._zf.writestr(f"xl/worksheets/_rels/sheet{sheet_id}.bin.rels", self._create_worksheet_rels(sheet_id))
        self._zf.writestr(f"xl/worksheets/binaryIndex{sheet_id}.bin", self._binaryIndexBin)

    def write_dataframe(
        self, df, include_index: bool = False, na_rep=None, formats: dict = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
    ):
        """Write a pandas or polars DataFrame to the current sheet.

        pandas and polars are optional: they are imported lazily only when a
        DataFrame is actually passed in.

        Parameters
        ----------
        df:
            A ``pandas.DataFrame`` or ``polars.DataFrame``.
        include_index:
            pandas only -- prepend the index as leading column(s).
        na_rep:
            Value for missing cells, defaults to ``None`` (empty cell).
        formats:
            Optional ``{column_name: format_string}`` mapping (see ``xlspy.F``).
        """
        from .dataframe import dataframe_to_rows

        self.write_sheet(
            dataframe_to_rows(df, include_index=include_index, na_rep=na_rep, formats=formats),
            header_style=header_style,
            styles=styles,
        )

    def save(self):
        if self._zf is None:
            raise RuntimeError("Zip file not initialized. Use context manager or initialize manually.")
            
        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.bin", self._create_workbook_bin())
        self._zf.writestr("xl/styles.bin", self._build_styles_bin())
        self._zf.writestr("xl/_rels/workbook.bin.rels", self._create_workbook_rels())

        with self._zf.open("xl/sharedStrings.bin", 'w') as sst_file:
            self._write_shared_strings_bin(sst_file)
        
    def close(self):
        if self._zf is not None:
            self._zf.close()
            self._zf = None

    def _create_content_types(self) -> str:
        parts = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.bin" ContentType="application/vnd.ms-excel.worksheet"/>'
            f'<Override PartName="/xl/worksheets/binaryIndex{i + 1}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>'
            for i in range(self._sheet_count)
        )

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
{parts}
<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>'''

    def _create_root_rels(self) -> str:
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>'''

    def _create_workbook_rels(self) -> str:
        relationships = [
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.bin"/>'
            for i in range(self._sheet_count)
        ]
        
        style_rid = self._sheet_count + 1
        shared_strings_rid = self._sheet_count + 2
        relationships.append(f'<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>')
        relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>')

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _create_worksheet_rels(self, sheet_id) -> str:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex{sheet_id}.bin"/>
</Relationships>'''

    def _create_workbook_bin(self) -> bytes:
        data = bytearray(self._workbookBinStart)
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            r_id = f"rId{sheet_id}"
            
            sheet_name_utf16 = sheet_name.encode('utf-16-le')
            r_id_utf16 = r_id.encode('utf-16-le')
            record_length = 4 + 3 * 4 + len(sheet_name_utf16) + len(r_id_utf16)
            
            data.extend(b'\x9C\x01')
            data.extend(self._write_custom_variable_int(record_length))
            
            if hidden:
                data.extend(struct.pack('<I', 1))
            else:
                data.extend(struct.pack('<I', 0)) 

            data.extend(struct.pack('<I', sheet_id))
            data.extend(struct.pack('<I', len(r_id_utf16) // 2))
            data.extend(r_id_utf16)
            data.extend(struct.pack('<I', len(sheet_name_utf16) // 2))
            data.extend(sheet_name_utf16)

        data.extend(self._workbookBinMiddle)
        self._write_filter_defined_names(data)
        data.extend(self._workbookBinEnd)
        return bytes(data)

    def _write_custom_variable_int(self, value: int) -> bytes:
        if value < 0:
            raise ValueError("Value must be non-negative")
        
        buffer = bytearray()
        temp_val = value
        while True:
            byte = temp_val & 0x7F
            temp_val >>= 7
            if temp_val > 0:
                byte |= 0x80
                buffer.append(byte)
            else:
                buffer.append(byte)
                break
        return bytes(buffer)

    def _write_shared_strings_bin(self, sst_file: io.BufferedWriter):
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536

        buffer.write(b'\x9F\x01\x08')
        buffer.write(struct.pack('<II', self._sst_all_count, self._sst_unique_count))
        
        for s in self._shared_strings:
            encoded = s.encode('utf-16-le')
            str_len = len(encoded) // 2
            if str_len > 32767:
                encoded = encoded[: 32767 * 2]
                # Do not leave a dangling high surrogate at the truncation edge.
                if len(encoded) >= 2 and 0xD800 <= struct.unpack_from('<H', encoded, len(encoded) - 2)[0] <= 0xDBFF:
                    encoded = encoded[:-2]
                str_len = len(encoded) // 2
            
            record_length = 1 + 4 + 2 * str_len
            
            buffer.write(b'\x13')
            buffer.write(self._write_custom_variable_int(record_length))
            buffer.write(b'\x00')
            buffer.write(struct.pack('<I', str_len))
            buffer.write(encoded)

            if buffer.tell() > BUFFER_SIZE:
                sst_file.write(buffer.getvalue())
                buffer.seek(0)
                buffer.truncate()
            
        buffer.write(b'\xA0\x01\x00')
        
        if buffer.tell() > 0:
            sst_file.write(buffer.getvalue())

    def _build_font_numfmt_record(self, ifmt: int, fmt_string: str) -> bytes:
        encoded = fmt_string.encode('utf-16-le')
        cch = len(encoded) // 2
        remaining = 2 + 4 + cch * 2
        buf = bytearray()
        buf.append(0x2C)
        buf.append(remaining)
        buf.extend(struct.pack('<H', ifmt))
        buf.extend(struct.pack('<I', cch))
        buf.extend(encoded)
        return bytes(buf)

    def _build_style_record(
        self, ifmt: int, flags: int, font_id: int = 0, fill_id: int = 0,
    ) -> bytes:
        buf = bytearray()
        buf.extend(b'\x2F\x10')
        buf.extend(b'\x00\x00')
        buf.extend(struct.pack('<H', ifmt))
        buf.extend(struct.pack('<H', font_id))
        buf.extend(struct.pack('<H', fill_id))
        buf.extend(b'\x00\x00')
        buf.extend(b'\x00\x00')
        buf.extend(b'\x10\x10')
        buf.extend(struct.pack('<H', flags))
        return bytes(buf)

    @staticmethod
    def _xlsb_record(record_id: bytes, payload: bytes) -> bytes:
        """Encode a record using the compact identifiers used in styles.bin."""
        return record_id + bytes((len(payload),)) + payload

    @staticmethod
    def _xlsb_color(color: str | None) -> bytes:
        if color is None:
            # Automatic color, no explicit RGB value.
            return bytes(8)
        red, green, blue = bytes.fromhex(color)
        # fValidRGB=1, xColorType=2 (RGB), followed by index/tint/RGBA.
        return bytes((0x05, 0, 0, 0, red, green, blue, 0xFF))

    def _xlsb_font_record(self, key: tuple) -> bytes:
        color, bold, italic, underline = key
        payload = bytearray()
        payload.extend(struct.pack('<H', 220))  # 11pt in twips
        flags = (1 if bold else 0) | (2 if italic else 0)
        payload.extend(struct.pack('<H', flags))
        payload.extend(struct.pack('<H', 700 if bold else 400))
        payload.extend(struct.pack('<H', 0))  # baseline, not sub/superscript
        payload.extend((1 if underline else 0, 2, 0, 0))  # underline/family/charset/reserved
        payload.extend(self._xlsb_color(color))
        payload.append(2)  # minor theme font
        name = 'Calibri'.encode('utf-16-le')
        payload.extend(struct.pack('<I', len(name) // 2))
        payload.extend(name)
        return self._xlsb_record(b'\x2B', bytes(payload))

    def _xlsb_fill_record(self, color: str) -> bytes:
        payload = bytearray(struct.pack('<I', 1))  # solid fill
        payload.extend(self._xlsb_color(color))
        payload.extend(self._xlsb_color(None))
        payload.extend(struct.pack('<I', 0))  # gradient fields (unused)
        payload.extend(bytes(40))
        payload.extend(struct.pack('<I', 0))
        return self._xlsb_record(b'\x2D', bytes(payload))

    def _build_styles_bin(self) -> bytes:
        if not self._format_registry and not self._style_components:
            return self._stylesBin

        buf = bytearray()

        # BrtBeginStyleSheet
        buf.extend(self._stylesBin[0:3])

        # ── Fonts section (contains font/numfmt records) ──
        total_records = 2 + len(self._format_registry)
        buf.extend(b'\xE7\x04\x04')
        buf.extend(struct.pack('<H', total_records))
        buf.extend(b'\x00\x00')

        # Font/numfmt record 1: ifmt=164, "yyyy\-mm\-dd\ hh:mm"
        buf.extend(self._build_font_numfmt_record(164, 'yyyy\\-mm\\-dd\\ hh:mm'))
        # Font/numfmt record 2: ifmt=166, "yyyy\-mm\-dd"
        buf.extend(self._build_font_numfmt_record(166, 'yyyy\\-mm\\-dd'))

        # Custom format records
        for fmt_string, numfmt_id in self._format_registry.items():
            buf.extend(self._build_font_numfmt_record(numfmt_id, fmt_string))

        # BrtEndFonts
        buf.extend(self._stylesBin[88:91])

        # ── Static sections from default styles, with custom font/fill records ──
        static_styles = bytearray(self._stylesBin[91:409])
        custom_font_keys = [
            key for key, font_id in sorted(self._font_registry.items(), key=lambda item: item[1])
            if font_id >= 2
        ]
        font_records = b''.join(self._xlsb_font_record(key) for key in custom_font_keys)
        if font_records:
            # Add font records immediately before BrtEndFonts.
            struct.pack_into('<I', static_styles, 3, 2 + len(custom_font_keys))
            static_styles[193 - 91:193 - 91] = font_records

        fill_records = b''.join(self._xlsb_fill_record(color) for color in self._fill_registry)
        if fill_records:
            fill_begin = 196 - 91 + len(font_records)
            fill_count = struct.unpack_from('<I', static_styles, fill_begin + 3)[0]
            struct.pack_into('<I', static_styles, fill_begin + 3, fill_count + len(self._fill_registry))
            # The inserted font bytes shift this boundary; locate BrtEndFills
            # by its complete encoded marker (record id 0x25C, empty payload).
            end_fills = b'\xDC\x04\x00'
            end_offset = static_styles.find(end_fills, fill_begin + 7)
            if end_offset < 0:
                raise RuntimeError('Could not locate BrtEndFills in the XLSB style template')
            static_styles[end_offset:end_offset] = fill_records
        buf.extend(static_styles)

        # ── CellXFs section (BrtBeginCellXFs + XF + BrtEndCellXFs) ──
        buf.extend(self._stylesBin[409:437])

        # ── NumFmts section (style records) ──
        total_styles = 4 + len(self._format_xf_map) + len(self._style_components)
        buf.extend(b'\xE9\x04\x04')
        buf.extend(struct.pack('<H', total_styles))
        buf.extend(struct.pack('<H', 0))

        # Style 0: default (fontId=0, ifmt=0)
        buf.extend(self._build_style_record(0, 0x0000, 0x00))
        # Style 1: date (fontId=0, ifmt=164)
        buf.extend(self._build_style_record(164, 0x0001, 0x00))
        # Style 2: datetime (fontId=0, ifmt=166)
        buf.extend(self._build_style_record(166, 0x0001, 0x00))
        # Style 3: bold header (fontId=1, ifmt=256, byte6=0x01)
        buf.extend(self._build_style_record(256, 0x0000, 0x01))

        # Custom format/style records follow the XF indices assigned while
        # streaming cells, so preserve that exact order.
        custom_xfs = {index: (self._format_registry[fmt], 0, 0, 0x0001)
                      for fmt, index in self._format_xf_map.items()}
        custom_xfs.update({
            index: (numfmt_id, font_id, fill_id, 0x0001 if numfmt_id else 0x0000)
            for index, (font_id, fill_id, numfmt_id, _style) in self._style_components.items()
        })
        for _index, (numfmt_id, font_id, fill_id, flags) in sorted(custom_xfs.items()):
            buf.extend(self._build_style_record(numfmt_id, flags, font_id, fill_id))

        # BrtEndNumFmts
        buf.extend(b'\xEA\x04\x00')

        # ── Remaining sections (BrtBeginDXFs, tableStyles, BrtEndStyleSheet) ──
        buf.extend(self._stylesBin[519:])

        return bytes(buf)

    def _write_cols_width(self, buffer: io.BytesIO, start_col: int, end_col: int, col_widths_array: list[float]):
        buffer.write(b'\x86\x03')
        
        for i in range(start_col, end_col):
            buffer.write(b'\x00\x3C\x12')
            buffer.write(struct.pack('<I', i))
            buffer.write(struct.pack('<I', i))
            buffer.write(b'\x00')
            width_byte = struct.pack('<B', int(col_widths_array[i - start_col]))
            buffer.write(width_byte)
            buffer.write(b'\x00\x00')
            buffer.write(b'\x00\x00\x00\x00')
            buffer.write(b'\x02')

        buffer.write(b'\x00\x87\x03\x00')

    def _write_worksheet_bin(
        self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[any]],
        worksheet_index: int, *, style_ranges=(), header_style: Style | None = None,
    ):
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536

        buffer.write(self._sheet1Bytes[:54])
        if self._sheetCnt != 1:
            buffer.write(b'\x9C')
        else:
            buffer.write(b'\xDC')
        self._sheetCnt += 1

        buffer.write(self._sheet1Bytes[55:84])

        buffer.write(self._stickHeaderA1bytes)

        data_iterator = iter(worksheet_data)
        rows_to_analyze = []
        max_cols = 1

        for _ in range(100):
            try:
                row = next(data_iterator)
                rows_to_analyze.append(row)
                row_len = 0
                for c in row:
                    cv = _unwrap_cell(c) if c is not None else None
                    if cv is not None:
                        row_len += 1
                max_cols = max(max_cols, row_len if row else 0)
            except StopIteration:
                break

        if max_cols == 0:
            max_cols = 1

        start_col = 0
        end_col = max_cols

        col_widths_array = accel_calc_column_widths(rows_to_analyze, max_cols)
        if col_widths_array is None:
            col_widths_array = []
            for i in range(max_cols):
                max_width = 0
                for row in rows_to_analyze:
                    if i < len(row):
                        cell = row[i] if i < len(row) else None
                        cell = _unwrap_cell(cell) if cell is not None else None
                        if cell is not None:
                            if isinstance(cell, Decimal):
                                cell = float(cell)
                            if isinstance(cell, datetime.datetime):
                                temp_width = 18
                            else:
                                str_cell = str(cell)
                                if '\n' in str_cell:
                                    str_cell = str_cell.split('\n')[0]
                                lenn = len(str_cell)
                                temp_width = 1.25 * lenn + 4
                            max_width = max(max_width, temp_width)
                
                if max_width == 0:
                    max_width = 4.0
                if max_width > 255:
                    max_width = 255
                col_widths_array.append(max_width)

        buffer.write(self._sheet1Bytes[84:159])
        self._write_cols_width(buffer, start_col, end_col, col_widths_array)
        buffer.write(self._sheet1Bytes[159:175])
        buffer.write(b'\x26\x00')
        buffer.write(b'\x91\x01\x00')

        shared_strings_dict = self._shared_strings_dict
        shared_strings_append = self._shared_strings.append

        if rows_to_analyze:
            data_iterator = itertools.chain(rows_to_analyze, data_iterator)
        else:
            data_iterator = iter([])

        row_idx = -1  
        for row_idx, row in enumerate(data_iterator):
            has_formatted = any(_is_formatted_cell(c) for c in row if c is not None)
            row_styles = styles_for_row(row_idx, len(row), style_ranges, header_style)

            if not has_formatted:
                style_ids = None
                if row_styles:
                    first_style = next(iter(row_styles.values()))
                    if (len(row_styles) == len(row)
                            and all(style is first_style for style in row_styles.values())):
                        style_ids = self._register_style(first_style)
                    else:
                        style_ids = tuple(
                            (col, self._register_style(row_styles[col]))
                            for col in sorted(row_styles)
                        )
                c_result = encode_xlsb_row(
                    row, self._shared_strings_dict, self._shared_strings,
                    self._sst_unique_count, self._sst_all_count, row_idx, style_ids
                )
                if c_result is not None:
                    cell_data, self._sst_unique_count, self._sst_all_count = c_result
                    buffer.write(cell_data)
                    if buffer.tell() > BUFFER_SIZE:
                        sheet_file.write(buffer.getvalue())
                        buffer.seek(0)
                        buffer.truncate()
                    continue

            last_col = len(row) - 1 if row else 0

            buffer.write(b'\x00\x19')
            buffer.write(struct.pack('<I', row_idx))
            buffer.write(b'\x00\x00\x00\x00')
            buffer.write(struct.pack('<I', 0x12c))
            buffer.write(b'\x00\x01\x00\x00\x00')
            buffer.write(struct.pack('<II', 0, last_col))

            for col_idx, cell in enumerate(row):
                if cell is None:
                    cell_style = row_styles.get(col_idx)
                    if cell_style is not None:
                        xf_index = self._register_style(cell_style)
                        buffer.write(b'\x01\x08')
                        buffer.write(struct.pack('<II', col_idx, xf_index))
                    continue

                fmt_string = _get_format(cell)
                if fmt_string is not None:
                    cell = _unwrap_cell(cell)
                    cell_style = row_styles.get(col_idx)
                    if cell_style is not None:
                        cell_style = merge_styles(cell_style, Style(number_format=fmt_string))
                else:
                    cell_style = row_styles.get(col_idx)
                if cell_style is not None:
                    xf_index = self._register_style(cell_style)
                    has_custom_fmt = True
                elif fmt_string is not None:
                    xf_index = self._register_format(fmt_string)
                    has_custom_fmt = True
                else:
                    xf_index = 0
                    has_custom_fmt = False

                if isinstance(cell, str):
                    if has_custom_fmt:
                        self._write_str(buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx, xf_index)
                    else:
                        self._write_str(buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx)
                elif isinstance(cell, bool):
                    style_ref = xf_index if has_custom_fmt else 0
                    buffer.write(b'\x04\x09')
                    buffer.write(struct.pack('<IIB', col_idx, style_ref, 1 if cell else 0))
                elif isinstance(cell, int):
                    style_ref = xf_index if has_custom_fmt else 0
                    if -536870912 <= cell <= 536870911:
                        rk_val = (cell << 2) | 0b10
                        buffer.write(b'\x02\x0C')
                        buffer.write(struct.pack('<IIi', col_idx, style_ref, rk_val))
                    else:
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, float(cell)))
                elif isinstance(cell, (float, Decimal)):
                    style_ref = xf_index if has_custom_fmt else 0
                    buffer.write(b'\x05\x10')
                    buffer.write(struct.pack('<IId', col_idx, style_ref, float(cell)))
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = xf_index if has_custom_fmt else 1
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = xf_index if has_custom_fmt else 2
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                else:
                    self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                    xf_index if has_custom_fmt else None)

            if buffer.tell() > BUFFER_SIZE:
                sheet_file.write(buffer.getvalue())
                buffer.seek(0)
                buffer.truncate()

        buffer.write(self._sheet1Bytes[218:290])

        last_row_idx = row_idx if row_idx > -1 else 0 

        if row_idx >= 0:
            buffer.write(self._autoFilterStartBytes)
            buffer.write(struct.pack('<I', 0))
            buffer.write(struct.pack('<I', last_row_idx))
            buffer.write(struct.pack('<I', 0))
            buffer.write(struct.pack('<I', 0 + end_col - 1))
            buffer.write(self._autoFilterEndBytes)
            self._filtered_data_list.append(
                FilterData(
                sheet_index=worksheet_index,
                start_column=0,
                end_column=end_col - 1,
                start_row=0,
                end_row=last_row_idx
            ))

        buffer.write(self._sheet1Bytes[290:])
        if buffer.tell() > 0:
            sheet_file.write(buffer.getvalue())

    def _write_str(self, buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx, style_override=None):
        self._sst_all_count += 1
        if cell not in shared_strings_dict:
            string_index = self._sst_unique_count
            shared_strings_dict[cell] = string_index
            shared_strings_append(cell)
            self._sst_unique_count += 1
        else:
            string_index = shared_strings_dict[cell]
        if style_override is not None:
            style_ref = style_override
        else:
            style_ref = 3 if row_idx == 0 else 0
        buffer.write(b'\x07\x0C')
        buffer.write(struct.pack('<III', col_idx, style_ref, string_index))

    def _write_filter_defined_names(self, sw: bytearray) -> None:
        filtered_dict_items_cnt = len(self._filtered_data_list) if self._filtered_data_list else 0
        
        if (filtered_dict_items_cnt > 0 and 
            (0x80 + (filtered_dict_items_cnt - 21) * 0x0c) <= 255):
            
            sw.extend(self._magicFilterExcel2016Fix0)
            
            if filtered_dict_items_cnt <= 10:
                sw.extend(bytes([0x10 + (filtered_dict_items_cnt - 1) * 0x0c, 
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
                
            elif filtered_dict_items_cnt <= 20:
                sw.extend(bytes([0x10 + (filtered_dict_items_cnt - 1) * 0x0c,
                               (filtered_dict_items_cnt - 1) // 10,
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
                
            else:
                sw.extend(bytes([0x80 + (filtered_dict_items_cnt - 21) * 0x0c,
                               (filtered_dict_items_cnt - 1) // 10,
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
            
            for nm in range(filtered_dict_items_cnt):
                sw.extend(bytes([0x00, 0x00, 0x00, 0x00]))
                sw.extend(bytes([self._filtered_data_list[nm].sheet_index, 0x00, 0x00, 0x00]))
                sw.extend(bytes([self._filtered_data_list[nm].sheet_index, 0x00, 0x00, 0x00]))
            
            sw.extend(bytes([0xE2, 0x02, 0x00]))
            
            for sheet_num in range(len(self._filtered_data_list)):
                start_column = self._filtered_data_list[sheet_num].start_column
                end_column = self._filtered_data_list[sheet_num].end_column
                start_row = self._filtered_data_list[sheet_num].start_row
                end_row = self._filtered_data_list[sheet_num].end_row
                sheet_index = self._filtered_data_list[sheet_num].sheet_index
                
                magic_filter_fix = self._magicFilterExcel2016Fix1.copy()
                magic_filter_fix[7] = sheet_index
                magic_filter_fix[-2] = sheet_num
                sw.extend(magic_filter_fix)

                sw.extend(struct.pack('<I', start_row))
                sw.extend(struct.pack('<I', end_row))
                sw.extend(struct.pack('<H', start_column))
                sw.extend(struct.pack('<H', end_column))
                
                sw.extend(self._magicFilterExcel2016Fix2)


    _autoFilterStartBytes = bytes([0xA1, 0x01, 0x10])
    _sheet1Bytes = bytes([
        0x81,0x01,0x00,0x93,0x01,0x17,0xCB,0x04,0x02,0x00,0x40,0x00,0x00,0x00,0x00,0x00,0x00,0xFF,0xFF,0xFF,
        0xFF,0xFF,0xFF,0xFF,0xFF,0x00,0x00,0x00,0x00,0x94,0x01,0x10,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x85,0x01,0x00,0x89,0x01,0x1E,0xDC,0x03,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x00,0x00,0x64,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x98,0x01,0x24,0x03,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x8A,0x01,0x00,0x86,0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x95,0x08,0x02,
        0x05,0x00,0x26,0x00,0xE5,0x03,0x0C,0xFF,0xFF,0xFF,0xFF,0x08,0x00,0x2C,0x01,0x00,0x00,0x00,0x00,0x91,
        0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x80,0x08,0x02,0x05,0x00,0x26,0x00,0x00,0x19,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2C,0x01,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x07,0x0C,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x92,0x01,
        0x00,0x97,0x04,0x42,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0xDD,0x03,0x02,0x10,0x00,0xDC,0x03,0x30,0x66,0x66,
        0x66,0x66,0x66,0x66,0xE6,0x3F,0x66,0x66,0x66,0x66,0x66,0x66,0xE6,0x3F,0x00,0x00,0x00,0x00,0x00,0x00,
        0xE8,0x3F,0x00,0x00,0x00,0x00,0x00,0x00,0xE8,0x3F,0x33,0x33,0x33,0x33,0x33,0x33,0xD3,0x3F,0x33,0x33,
        0x33,0x33,0x33,0x33,0xD3,0x3F,0x25,0x06,0x01,0x00,0x00,0x10,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x82,0x01,0x00
    ])
    _stylesBin = bytes([
        0x96,0x02,0x00,0xE7,0x04,0x04,0x02,
        0x00,0x00,0x00,0x2C,0x2C,0xA4,0x00,0x13,
        0x00,0x00,0x00,0x79,0x00,0x79,0x00,0x79,
        0x00,0x79,0x00,0x5C,0x00,0x2D,0x00,0x6D,
        0x00,0x6D,0x00,0x5C,0x00,0x2D,0x00,0x64,
        0x00,0x64,0x00,0x5C,0x00,0x20,0x00,0x68,
        0x00,0x68,0x00,0x3A,0x00,0x6D,0x00,0x6D,
        0x00,0x2C,0x1E,0xA6,0x00,0x0C,0x00,0x00,
        0x00,0x79,0x00,0x79,0x00,0x79,0x00,0x79,
        0x00,0x5C,0x00,0x2D,0x00,0x6D,0x00,0x6D,
        0x00,0x5C,0x00,0x2D,0x00,0x64,0x00,0x64,
        0x00,0xE8,0x04,0x00,0xE3,0x04,0x04,0x01,
        0x00,0x00,0x00,
        0x2B,0x27,0xDC,0x00,0x00,0x00,0x90,0x01,
        0x00,0x00,0x00,0x02,0x00,0x00,0x07,0x01,
        0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
        0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
        0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
        0x00,
        0x2B,0x27,0xDC,0x00,0x01,0x00,0xBC,0x02,
        0x00,0x00,0x00,0x02,0xEE,0x00,0x07,0x01,
        0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
        0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
        0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
        0x00,
        0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x81,0x08,0x00,0x26,0x00,0xE4,0x04,0x00,0xDB,0x04,0x04,0x02,0x00,0x00,0x00,
        0x2D,0x44,0x00,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x2D,0x44,0x11,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0xDC,0x04,0x00,0xE5,0x04,0x04,0x01,0x00,0x00,0x00,0x2E,0x33,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xE6,0x04,0x00,0xF2,
        0x04,0x04,0x01,0x00,0x00,0x00,0x2F,0x10,0xFF,0xFF,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,
        0x00,0xF3,0x04,0x00,0xE9,0x04,0x04,
        0x04,
        0x00,0x00,0x00,
        0x2F,0x10,0x00,0x00,0x00,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,
        0x2F,0x10,0x00,0x00,0xA4,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,
        0x2F,0x10,0x00,0x00,0xA6,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,
        0x2F,0x10,0x00,0x00,0x00,0x01,    0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,
        0xEA,0x04,0x00,0xEB,0x04,0x04,0x01,0x00,
        0x00,0x00,0x25,0x06,0x01,0x00,0x02,0x11,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x30,0x1C,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x08,0x00,0x00,0x00,0x4E,
        0x00,0x6F,0x00,0x72,0x00,0x6D,0x00,0x61,0x00,0x6C,0x00,0x6E,0x00,0x79,0x00,0xEC,0x04,0x00,0xF9,0x03,0x04,0x00,0x00,
        0x00,0x00,0xFA,0x03,0x00,0xFC,0x03,0x50,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x54,0x00,0x61,0x00,0x62,0x00,0x6C,
        0x00,0x65,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4D,0x00,0x65,0x00,0x64,0x00,0x69,0x00,0x75,0x00,
        0x6D,0x00,0x32,0x00,0x11,0x00,0x00,0x00,0x50,0x00,0x69,0x00,0x76,0x00,0x6F,0x00,0x74,0x00,0x53,0x00,0x74,0x00,0x79,
        0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0x36,0x00,0xFD,0x03,0x00,0x23,
        0x04,0x02,0x0E,0x00,0x00,0xEB,0x08,0x00,0xF6,0x08,0x2A,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x53,0x00,0x6C,0x00,
        0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,
        0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xF7,0x08,0x00,0xEC,0x08,0x00,0x24,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xB0,0x10,
        0x00,0xB2,0x10,0x32,0x00,0x00,0x00,0x00,0x15,0x00,0x00,0x00,0x54,0x00,0x69,0x00,0x6D,0x00,0x65,0x00,0x53,0x00,0x6C,
        0x00,0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,
        0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xB3,0x10,0x00,0xB1,0x10,0x00,0x24,0x00,0x97,0x02,0x00
    ])
    _workbookBinStart = bytes([
        0x83,0x01,0x00,0x80,0x01,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x02,0x00,0x00,0x00,0x78,0x00,0x6C,0x00,0x01,0x00,0x00,0x00,0x37,0x00,0x01,0x00,0x00,0x00,
        0x36,0x00,0x05,0x00,0x00,0x00,0x32,0x00,0x34,0x00,0x33,0x00,0x32,0x00,0x36,0x00,0x99,0x01,0x0C,0x20,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x25,0x06,0x01,0x00,0x03,0x0F,0x00,0x80,0x97,
        0x10,0x34,0x18,0x00,0x00,0x00,        0x43,0x00,0x3A,0x00,0x5C,0x00,0x78,0x00,0x6C,0x00,0x73,0x00,0x70,0x00,0x79,0x00,
        0x5C,0x00,0x74,0x00,0x65,0x00,0x73,0x00,0x74,0x00,0x5C,0x00,0x70,0x00,0x61,0x00,0x74,0x00,0x68,0x00,
        0x5C,0x00,0x64,0x00,0x6F,0x00,0x63,0x00,0x5C,0x00,0x74,0x00,0x6D,0x00,0x70,0x00,0x5C,0x00,0x26,0x00,0x25,0x06,0x01,0x00,
        0x00,0x10,0x00,0x80,0x81,0x18,0x82,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2F,0x00,0x00,0x00,
        0x31,0x00,0x33,0x00,0x5F,0x00,0x6E,0x00,0x63,0x00,0x72,0x00,0x3A,0x00,0x31,0x00,0x5F,0x00,0x7B,0x00,
        0x31,0x00,0x36,0x00,0x35,0x00,0x30,0x00,0x38,0x00,0x44,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x43,0x00,
        0x46,0x00,0x38,0x00,0x37,0x00,0x2D,0x00,0x34,0x00,0x37,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x38,0x00,
        0x34,0x00,0x35,0x00,0x36,0x00,0x2D,0x00,0x44,0x00,0x34,0x00,0x41,0x00,0x34,0x00,0x30,0x00,0x31,0x00,
        0x31,0x00,0x33,0x00,0x31,0x00,0x35,0x00,0x36,0x00,0x37,0x00,0x7D,0x00,0x2F,0x00,0x00,0x00,0x2F,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,
        0x87,0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x10,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x0D,0x00,
        0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x00,0x00,0x00,0x00,0x26,0x00,0x9E,0x01,0x1D,0x00,0x00,0x00,0x00,0x9E,
        0x16,0x00,0x00,0xB4,0x69,0x00,0x00,0xE8,0x26,0x00,0x00,0x58,0x02,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x78,0x88,0x01,0x00,0x8F,0x01,0x00
    ])
    _workbookBinMiddle = bytes([0x90,0x01,0x00])
    _workbookBinEnd = bytes([
        0x9D,0x01,0x1A,0x35,0xEA,0x02,0x00,0x01,0x00,0x00,0x00,0x64,0x00,0x00,0x00,0xFC,0xA9,0xF1,0xD2,0x4D,
        0x62,0x50,0x3F,0x01,0x00,0x00,0x00,0x6A,0x00,0x9B,0x01,0x01,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xAB,
        0x10,0x01,0x01,0x24,0x00,0x84,0x01,0x00
    ])
    _binaryIndexBin = bytes([
        0x2A,0x18,0x00,0x00,0x00,0x00,0x20,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x95,0x02,0x00
    ])
    _stickHeaderA1bytes =bytes([
        0x97,0x01,0x1D,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0xF0,0x3F,0x01,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x03,
    ])
    _autoFilterEndBytes = bytes([0xA2, 0x01, 0x00])
    _magicFilterExcel2016Fix0 = bytes([0xE1, 0x02, 0x00, 0xE5, 0x02, 0x00, 0xEA, 0x02])
    _magicFilterExcel2016Fix1 = bytearray([
            0x27,
            0x46,
            0x21,
            0x00,
            0x00,
            0x00,
            0x00,
            255,
            0x00,
            0x00,
            0x00,
            0x0F,
            0x00,
            0x00,
            0x00,
            0x5F,
            0x00,
            0x46,
            0x00,
            0x69,
            0x00,
            0x6C,
            0x00,
            0x74,
            0x00,
            0x65,
            0x00,
            0x72,
            0x00,
            0x44,
            0x00,
            0x61,
            0x00,
            0x74,
            0x00,
            0x61,
            0x00,
            0x62,
            0x00,
            0x61,
            0x00,
            0x73,
            0x00,
            0x65,
            0x00,
            0x0F,
            0x00,
            0x00,
            0x00,
            0x3B,
            255,
            0x00
        ])
    _magicFilterExcel2016Fix2 = bytes([0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF])
