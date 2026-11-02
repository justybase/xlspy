"""XlsbWriter cell string records: inline and rich strings, filter defined names."""

import struct
from .styles import (
    _FontKey,
    _color_key,
)
from .biff_utils import build_record
from .rich_text import CellRichText, InlineFont, TextBlock
from ._text import check_cell_text
from ._xlsb_writer_consts import _MAX_STR_RUNS
from ._xlsb_writer_host import _XlsbWriterHost


class _XlsbStringsMixin(_XlsbWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

    @staticmethod
    def _write_inline_str(buffer, text: str, col_idx: int, style_ref: int):
        """BrtCellSt: the text is stored in the cell, so no shared-string table grows.

        The count is in UTF-16 code units, so a non-BMP character counts as two.
        Lone surrogates raise UnicodeEncodeError here, as they do in the shared table.
        """
        check_cell_text(text)
        encoded = text.encode('utf-16-le')
        units = len(encoded) // 2
        buffer.write(build_record(
            0x06, struct.pack('<IIi', col_idx, style_ref, units) + encoded,
        ))

    def _rich_font_id(self, font: InlineFont) -> int:
        """The font table index of a run's font, shared with cell styles."""
        if font.rFont is not None and len(font.rFont) > 31:
            raise ValueError("XLSB font names are limited to 31 characters")
        key = _FontKey(
            font.rFont, font.sz, font.b, font.i, font.strike, font.u, font.vertAlign,
            _color_key(font.color), font.charset, font.family, font.scheme,
            font.outline, font.shadow, font.condense, font.extend,
        )
        return self._register_component(self._font_registry, key, 2)

    def _rich_runs(self, rich: CellRichText) -> tuple[str, tuple]:
        """The text and the StrRun list ``(ich, ifnt)`` of a rich cell.

        As in the files Excel saves, text before the first formatted run has
        no run (it takes the cell's font) and later plain text uses font 0.
        """
        texts, runs, position, current = [], [], 0, 0
        for part in rich:
            text = str(part)
            if not text:
                continue
            font_id = self._rich_font_id(part.font) if isinstance(part, TextBlock) else 0
            if font_id != current:
                runs.append((position, font_id))
                current = font_id
            texts.append(text)
            position += len(text.encode('utf-16-le')) // 2
        if len(runs) > _MAX_STR_RUNS:
            raise ValueError(f"XLSB rich text supports at most {_MAX_STR_RUNS} formatting runs")
        text = "".join(texts)
        check_cell_text(text)
        return text, tuple(runs)

    @staticmethod
    def _rich_str_payload(text: str, runs: tuple) -> bytes:
        """A RichStr ([MS-XLSB] 2.5.125): fRichStr, the text and its StrRuns."""
        encoded = text.encode('utf-16-le')
        payload = struct.pack('<BI', 0x01 if runs else 0x00, len(encoded) // 2) + encoded
        if runs:
            payload += struct.pack('<I', len(runs))
            payload += b"".join(struct.pack('<HH', ich, ifnt) for ich, ifnt in runs)
        return payload

    def _write_rich_str(self, buffer, rich: CellRichText, shared_strings_dict, shared_strings_append,
                        row_idx, col_idx, style_override=None):
        text, runs = self._rich_runs(rich)
        if not runs:
            self._write_str(buffer, text, shared_strings_dict, shared_strings_append, row_idx, col_idx,
                            style_override)
            return
        style_ref = style_override if style_override is not None else (3 if row_idx == 0 else 0)
        # Rich text always goes to the shared-string table, also with
        # use_shared_strings=False: Excel saves it there too, and the in-cell
        # BrtCellRString is a record readers rarely know. A tuple key never
        # equals a str, so the same text plain and formatted stays two entries.
        key = ("rich", text, runs)
        self._sst_all_count += 1
        string_index = shared_strings_dict.get(key)
        if string_index is None:
            string_index = self._sst_unique_count
            shared_strings_dict[key] = string_index
            shared_strings_append(key)
            self._sst_unique_count += 1
            self._sst_has_rich = True
        buffer.write(b'\x07\x0C')
        buffer.write(struct.pack('<III', col_idx, style_ref, string_index))

    def _write_str(self, buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx, style_override=None):
        if not self._use_shared_strings:
            style_ref = style_override if style_override is not None else (3 if row_idx == 0 else 0)
            self._write_inline_str(buffer, cell, col_idx, style_ref)
            return
        check_cell_text(cell)
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
