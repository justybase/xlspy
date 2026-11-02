"""XlsbWriter worksheet record stream for one sheet."""
from __future__ import annotations

from typing import Any
import datetime
import math
from decimal import Decimal
import struct
from collections.abc import Iterable
import io
import itertools
from ._accel import encode_xlsb_row, encode_xlsb_row_inline, xlsb_row_has_formatted as accel_row_has_formatted
from ._widths import estimate_column_widths
from .formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT, _is_formatted_cell, _unwrap_cell
from .excel_dates import datetime_to_serial
from .styles import (
    HYPERLINK_STYLE,
    Style,
    merge_styles,
    parse_cell_range,
    styles_for_row,
    _date_cell_xf,
    _record_cell_xf,
    _style_ids_for_row,
    _xf_for_style_format,
)
from .biff_utils import build_record
from ._xlsb_print import FIT_TO_PAGE_BIT, FIT_TO_PAGE_BYTE, print_records
from ._xlsb_validation import validations_bin
from .rich_text import CellRichText
from .formulas import Formula
from .hyperlinks import Hyperlink
from ._scalars import numpy_value
from ._xlsb_slicers import (
    sheet_slicers_bin,
)
from ._xlsb_sparklines import sparklines_bin
from .printing import PrintSettings
from ._writer_common import FilterData
from ._xlsb_writer_consts import _BRT_BEGIN_MERGE_CELLS, _BRT_DRAWING, _BRT_END_MERGE_CELLS, _BRT_HLINK, _BRT_MERGE_CELL, _hlink_payload
from ._xlsb_writer_host import _XlsbWriterHost
from ._xlsb_formulas import ArrayValue, formula_cell
from ._xlsb_metadata import replace_records, sheet_protection_bin, wide_string


class _XlsbBinMixin(_XlsbWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

    def _write_worksheet_bin(
        self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[Any]],
        worksheet_index: int, *, style_ranges=(), header_style: Style | None = None,
        merge_refs=(), widths=None, hidden_cols=frozenset(), heights=None,
        hidden_row_set=frozenset(), freeze_panes: str | None = None,
        auto_filter: bool | None = None, auto_filter_ref: str | None = None,
        drawing_rel: str | None = None,
        table_rel_ids: list[str] | None = None,
        cf_rules: list | None = None,
        row_outline=None, col_outline=None,
        sheet_protection=None, comment_vml_rid=None, show_grid_lines=True, view=(True, True, False),
        filter_cells=(), header_rows=1, tab_key=None, zoom=None,
        link_rids: dict | None = None, link_rid_start: int = 2,
        print_spec: PrintSettings | None = None, validations=(), sparklines=(),
        slicer_rel: str | None = None,
    ):
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536
        # (row, column, Hyperlink, relationship id or None) for the BrtHLink records.
        hyperlinks = []
        if link_rids is None:
            link_rids = {}

        if tab_key is None:
            head = self._sheet1Bytes[:54]
        else:
            # Bytes 9-16 are brtcolorTab, the BrtColor inside BrtWsProp.
            head = self._sheet1Bytes[:9] + self._color_payload(tab_key) + self._sheet1Bytes[17:54]
        if print_spec is not None and print_spec.fit_to_page:
            # The BrtWsProp payload starts at byte 6; fFitToPage is in its second byte.
            patched = bytearray(head)
            patched[6 + FIT_TO_PAGE_BYTE] |= FIT_TO_PAGE_BIT
            head = bytes(patched)
        buffer.write(head)
        # First flag byte of BrtBeginWsView: fDspGrid (0x04), fDspRwCol,
        # fDspZeros, fSelected (0x40, first sheet only) and fDspRuler.
        view_flags = 0xDC if self._sheet_number == 1 else 0x9C
        headers, zeros, right_to_left = view
        if not show_grid_lines:
            view_flags &= ~0x04
        if not headers:
            view_flags &= ~0x08  # fDspRwCol
        if not zeros:
            view_flags &= ~0x10  # fDspZeros
        if right_to_left:
            view_flags |= 0x20  # fRightToLeft
        buffer.write(bytes((view_flags,)))
        self._sheet_number += 1

        view_tail = self._sheet1Bytes[55:84]
        if zoom is not None:
            # wScale and wScaleNormal of BrtBeginWsView ([MS-XLSB] 2.4.304),
            # payload offsets 18 and 20.
            scale = struct.pack('<HH', zoom, zoom)
            view_tail = view_tail[:17] + scale + view_tail[21:]
        buffer.write(view_tail)

        buffer.write(self._pane_record(freeze_panes))

        data_iterator = iter(worksheet_data)
        rows_to_analyze = []
        max_cols = 1

        for _ in range(100):
            try:
                row = next(data_iterator)
            except StopIteration:
                break
            rows_to_analyze.append(row)
            # Columns up to the last cell with a value.
            for col_idx in range(len(row) - 1, -1, -1):
                cell = row[col_idx]
                if cell is not None and _unwrap_cell(cell) is not None:
                    max_cols = max(max_cols, col_idx + 1)
                    break

        start_col = 0
        end_col = max_cols

        if auto_filter is not False and len(rows_to_analyze) >= header_rows:
            filter_cells = [*filter_cells, (header_rows - 1, 0, max_cols - 1)]
        col_widths_array = estimate_column_widths(
            rows_to_analyze, max_cols,
            lambda row_idx, length: styles_for_row(row_idx, length, style_ranges, header_style),
            filter_cells,
        )

        # The last two bytes of BrtWsFmtInfo are the maximum row and column
        # outline levels; the template holds zeros.
        outline_levels = bytes((
            row_outline.max_level if row_outline is not None else 0,
            col_outline.max_level if col_outline is not None else 0,
        ))
        if freeze_panes is None:
            buffer.write(self._sheet1Bytes[84:157])
        else:
            buffer.write(self._selection_record(freeze_panes))
            # Template offsets 84..123 hold the default BrtSel record; the
            # remainder (EndWsView .. BrtWsFmtInfo) is layout-independent.
            buffer.write(self._sheet1Bytes[123:157])
        buffer.write(outline_levels)
        self._write_cols_width(
            buffer, start_col, end_col, col_widths_array, widths, hidden_cols,
            col_outline,
        )
        # Template offsets 159..177 hold BrtBeginSheetData plus the complete
        # ACCELLTABLE block (ACBegin/RwDescent/ACEnd); rows follow directly.
        buffer.write(self._sheet1Bytes[159:177])

        shared_strings_dict = self._shared_strings_dict
        shared_strings_append = self._shared_strings.append

        if rows_to_analyze:
            data_iterator = itertools.chain(rows_to_analyze, data_iterator)
        else:
            data_iterator = iter([])

        # Rows with a custom height or hidden flag cannot use the fast C
        # encoder, which only emits default row headers.
        heights = heights or {}
        custom_rows = set(heights) | set(hidden_row_set)
        array_ranges: list[Any] = []
        row_style_ids_cache: dict[Any, Any] = {}

        row_idx = -1
        last_written = -1  # index of the last row record written
        # Width of the whole sheet, measured as the rows stream past. The column
        # records come before the rows and use the sample, but the autofilter is
        # written after the rows and must span every column that holds a value.
        stream_width = end_col
        for row_idx, row in enumerate(data_iterator):
            for col_idx in range(len(row) - 1, -1, -1):
                cell = row[col_idx]
                if cell is not None and _unwrap_cell(cell) is not None:
                    if col_idx + 1 > stream_width:
                        stream_width = col_idx + 1
                    break
            c_has_formatted = accel_row_has_formatted(row)
            if c_has_formatted is None:
                has_formatted = any(_is_formatted_cell(c) for c in row if c is not None)
            else:
                has_formatted = c_has_formatted
            row_styles = styles_for_row(row_idx, len(row), style_ranges, header_style)

            if row_outline is None:
                outline_level, outline_hidden, outline_collapsed = 0, False, False
            else:
                outline_level, outline_hidden = row_outline.state_at(row_idx)
                outline_collapsed = row_outline.collapsed_at(row_idx)
            if (
                len(row) == 0 and not array_ranges and not row_styles and not outline_level
                and not outline_collapsed and row_idx not in custom_rows
            ):
                # An empty row with no height, flags or style needs no row record;
                # XLSB rows carry their own index, so the gap is kept.
                continue
            last_written = row_idx
            use_native = not array_ranges and (not has_formatted or self._fmt_xf)
            style_ids = None
            if use_native and row_styles:
                if has_formatted:
                    # Only look styles up: the Python encoder registers the
                    # styles of formatted cells in its own order.
                    try:
                        style_ids = _style_ids_for_row(
                            row_styles, len(row), self._style_xf_map.__getitem__,
                            row_style_ids_cache,
                        )
                    except KeyError:
                        use_native = False
                else:
                    style_ids = _style_ids_for_row(
                        row_styles, len(row), self._register_style,
                        row_style_ids_cache,
                    )
            if use_native:
                fmt_xf = self._fmt_xf if has_formatted else None
                row_header = None
                if row_idx in custom_rows or outline_level or outline_collapsed:
                    row_header = self._row_header_record(
                        row_idx, len(row) - 1 if row else 0,
                        self._row_height_twips(heights[row_idx]) if row_idx in heights else None,
                        row_idx in hidden_row_set or outline_hidden,
                        outline_level, outline_collapsed,
                    )
                cell_data = None
                if self._use_shared_strings:
                    c_result = encode_xlsb_row(
                        row, self._shared_strings_dict, self._shared_strings,
                        self._sst_unique_count, self._sst_all_count, row_idx,
                        style_ids, fmt_xf, row_header,
                    )
                    if c_result is not None:
                        cell_data, self._sst_unique_count, self._sst_all_count = c_result
                else:
                    cell_data = encode_xlsb_row_inline(row, row_idx, style_ids, fmt_xf, row_header)
                if cell_data is not None:
                    buffer.write(cell_data)
                    if buffer.tell() > BUFFER_SIZE:
                        sheet_file.write(buffer.getvalue())
                        buffer.seek(0)
                        buffer.truncate()
                    continue

            last_col = len(row) - 1 if row else 0
            if array_ranges:
                row = list(row)
                for c1, r1, c2, r2 in array_ranges:
                    if r1 < row_idx <= r2:
                        row.extend([None] * max(0, c2 + 1 - len(row)))
                        for col in range(c1, c2 + 1):
                            value = _unwrap_cell(row[col]) if row[col] is not None else None
                            if isinstance(value, Formula):
                                raise ValueError("overlapping array formulas")
                            row[col] = ArrayValue(r1, c1, value)
                last_col = max(0, len(row) - 1)
                row_styles = styles_for_row(row_idx, len(row), style_ranges, header_style)

            height_twips = None
            if row_idx in heights:
                height_twips = self._row_height_twips(heights[row_idx])
            row_header_position = buffer.tell()
            buffer.write(self._row_header_record(
                row_idx, last_col, height_twips,
                row_idx in hidden_row_set or outline_hidden,
                outline_level, outline_collapsed,
            ))

            row_cells = enumerate(row)
            for col_idx, cell in row_cells:
                if cell is None:
                    cell_style = row_styles.get(col_idx)
                    if cell_style is not None:
                        xf_index = self._register_style(cell_style)
                        buffer.write(b'\x01\x08')
                        buffer.write(struct.pack('<II', col_idx, xf_index))
                    continue

                # (value, format) and (value, Style) cells: the overlay goes
                # on top of the range style.
                overlay = None
                if _is_formatted_cell(cell):
                    cell, overlay = cell
                cell = numpy_value(cell)
                cell_style = row_styles.get(col_idx)
                if isinstance(cell, Hyperlink):
                    if isinstance(overlay, str):
                        raise TypeError("a hyperlink cell cannot carry a number format; style the cell instead")
                    # The link look sits under the range style and the cell's
                    # own style, as in the XLSX writer.
                    cell_style = merge_styles(HYPERLINK_STYLE, merge_styles(cell_style, overlay))
                    overlay = None
                    rid = None
                    if not cell.is_internal:
                        rid = link_rids.get(cell.target)
                        if rid is None:
                            rid = link_rids[cell.target] = f"rId{link_rid_start + len(link_rids)}"
                    hyperlinks.append((row_idx, col_idx, cell, rid))
                    cell = cell.display_text
                if overlay is not None:
                    if cell_style is not None:
                        xf_index = _xf_for_style_format(
                            cell_style, overlay, self._register_style,
                            self._style_format_xf,
                        )
                        # Only a registered style has the id the native encoder sees.
                        base_xf = self._style_xf_map.get(cell_style)
                        if base_xf is not None:
                            _record_cell_xf(self._fmt_xf, self._fmt_xf_styles,
                                            base_xf, overlay, xf_index)
                    else:
                        if isinstance(overlay, str):
                            xf_index = self._register_format(overlay)
                        else:
                            xf_index = self._register_style(overlay)
                        _record_cell_xf(self._fmt_xf, self._fmt_xf_styles,
                                        None, overlay, xf_index)
                    has_custom_fmt = True
                elif cell_style is not None:
                    xf_index = self._register_style(cell_style)
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
                    if not math.isfinite(float(cell)):
                        raise ValueError("numeric cell values must be finite")
                    style_ref = xf_index if has_custom_fmt else 0
                    buffer.write(b'\x05\x10')
                    buffer.write(struct.pack('<IId', col_idx, style_ref, float(cell)))
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = 1 if not has_custom_fmt else _date_cell_xf(
                            self._fmt_xf, self._xf_styles, self._register_style,
                            xf_index, DEFAULT_DATETIME_FORMAT, True,
                        )
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = 2 if not has_custom_fmt else _date_cell_xf(
                            self._fmt_xf, self._xf_styles, self._register_style,
                            xf_index, DEFAULT_DATE_FORMAT, False,
                        )
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                else:
                    # The native encoder hands rows with other objects to this path.
                    if isinstance(cell, Formula):
                        buffer.write(formula_cell(cell, col_idx, xf_index if has_custom_fmt else 0,
                                                  self._formula_context(), sheet=self._worksheet_data[worksheet_index][0], row=row_idx))
                        if cell.array_ref is not None:
                            bounds = parse_cell_range(cell.array_ref)
                            c1, r1, c2, r2 = bounds
                            if any(not (c2 < a or c1 > c or r2 < b or r1 > d) for a, b, c, d in array_ranges):
                                raise ValueError("overlapping array formulas")
                            array_ranges.append(bounds)
                            if c2 > last_col:
                                position = buffer.tell()
                                buffer.seek(row_header_position + 2 + 21)
                                buffer.write(struct.pack('<I', c2))
                                buffer.seek(position)
                                last_col = c2
                            for follower_col in range(c1 + 1, c2 + 1):
                                _, cache = next(row_cells, (follower_col, None))
                                cache = _unwrap_cell(cache) if cache is not None else None
                                if isinstance(cache, Formula):
                                    raise ValueError("overlapping array formulas")
                                follower_style = row_styles.get(follower_col)
                                style = self._register_style(follower_style) if follower_style is not None else 0
                                buffer.write(formula_cell(ArrayValue(r1, c1, cache), follower_col, style,
                                                          self._formula_context(), sheet=self._worksheet_data[worksheet_index][0], row=row_idx))
                        continue
                    if array_ranges:
                        if isinstance(cell, ArrayValue):
                            buffer.write(formula_cell(cell, col_idx, xf_index if has_custom_fmt else 0,
                                                      self._formula_context(), sheet=self._worksheet_data[worksheet_index][0], row=row_idx))
                            continue
                    if isinstance(cell, CellRichText):
                        self._write_rich_str(buffer, cell, shared_strings_dict, shared_strings_append,
                                             row_idx, col_idx, xf_index if has_custom_fmt else None)
                        continue
                    self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                    xf_index if has_custom_fmt else None)

            if buffer.tell() > BUFFER_SIZE:
                sheet_file.write(buffer.getvalue())
                buffer.seek(0)
                buffer.truncate()

        # Empty rows at the end of the data were skipped in the loop; their records are
        # written now, so the sheet keeps its shape (an empty frame is still one row).
        for trailing_row in range(last_written + 1, row_idx + 1):
            buffer.write(self._row_header_record(trailing_row, 0, None, False, 0, False))
        if array_ranges:
            final_row = max(bounds[3] for bounds in array_ranges)
            for extra_row in range(row_idx + 1, final_row + 1):
                active = [bounds for bounds in array_ranges if bounds[1] <= extra_row <= bounds[3]]
                final_col = max((bounds[2] for bounds in active), default=0)
                # Array rows past the data keep the row options of real rows.
                if row_outline is None:
                    extra_level, extra_hidden, extra_collapsed = 0, False, False
                else:
                    extra_level, extra_hidden = row_outline.state_at(extra_row)
                    extra_collapsed = row_outline.collapsed_at(extra_row)
                buffer.write(self._row_header_record(
                    extra_row, final_col,
                    self._row_height_twips(heights[extra_row]) if extra_row in heights else None,
                    extra_row in hidden_row_set or extra_hidden,
                    extra_level, extra_collapsed,
                ))
                extra_styles = styles_for_row(extra_row, final_col + 1, style_ranges, header_style)
                for c1, r1, c2, r2 in active:
                    for col in range(c1, c2 + 1):
                        style = self._register_style(extra_styles[col]) if col in extra_styles else 0
                        buffer.write(formula_cell(ArrayValue(r1, c1), col, style, self._formula_context(),
                                                  sheet=self._worksheet_data[worksheet_index][0], row=extra_row))
                if buffer.tell() > BUFFER_SIZE:
                    sheet_file.write(buffer.getvalue())
                    buffer.seek(0)
                    buffer.truncate()
            row_idx = max(row_idx, final_row)

        if row_outline is not None:
            # Grouped rows and summary rows past the data get an empty BrtRowHdr.
            for extra_idx in row_outline.indices_from(row_idx + 1):
                height_twips = None
                if extra_idx in heights:
                    height_twips = self._row_height_twips(heights[extra_idx])
                outline_level, outline_hidden = row_outline.state_at(extra_idx)
                buffer.write(self._row_header_record(
                    extra_idx, 0, height_twips,
                    extra_idx in hidden_row_set or outline_hidden,
                    outline_level, row_outline.collapsed_at(extra_idx),
                ))
                if buffer.tell() > BUFFER_SIZE:
                    sheet_file.write(buffer.getvalue())
                    buffer.seek(0)
                    buffer.truncate()

        if sheet_protection is None:
            buffer.write(self._sheet1Bytes[218:290])
        else:
            buffer.write(replace_records(self._sheet1Bytes[218:290], {0x217}, sheet_protection_bin(sheet_protection), before=(0x217,)))

        last_row_idx = row_idx if row_idx > -1 else 0 

        # The filter buttons sit on the last header row.
        filter_row = header_rows - 1
        if row_idx >= filter_row and auto_filter is not False:
            filter_first_col, filter_last_col = 0, stream_width - 1
            filter_first_row, filter_last_row = filter_row, last_row_idx
            if auto_filter_ref is not None:
                filter_first_col, filter_first_row, filter_last_col, filter_last_row = (
                    parse_cell_range(auto_filter_ref)
                )
            buffer.write(self._autoFilterStartBytes)
            buffer.write(struct.pack('<I', filter_first_row))
            buffer.write(struct.pack('<I', filter_last_row))
            buffer.write(struct.pack('<I', filter_first_col))
            buffer.write(struct.pack('<I', filter_last_col))
            buffer.write(self._autoFilterEndBytes)
            self._filtered_data_list.append(
                FilterData(
                sheet_index=worksheet_index,
                start_column=filter_first_col,
                end_column=filter_last_col,
                start_row=filter_first_row,
                end_row=filter_last_row
            ))

        if merge_refs:
            buffer.write(build_record(
                _BRT_BEGIN_MERGE_CELLS, struct.pack('<I', len(merge_refs)),
            ))
            for ref in merge_refs:
                first_col, first_row, last_col, last_row = parse_cell_range(ref)
                buffer.write(build_record(
                    _BRT_MERGE_CELL,
                    struct.pack('<IIII', first_row, last_row, first_col, last_col),
                ))
            buffer.write(build_record(_BRT_END_MERGE_CELLS, b''))

        if cf_rules:
            buffer.write(self._cf_bin_grouped(cf_rules))
        if validations:
            # DVALS follow the conditional formats and precede the hyperlinks.
            buffer.write(validations_bin(validations, self._formula_context()))
        for row_idx, col_idx, link, rid in hyperlinks:
            buffer.write(build_record(_BRT_HLINK, _hlink_payload(row_idx, col_idx, link, rid)))

        # Template offsets 290..346 hold BrtPrintOptions + BrtMargins; 346..
        # holds the FRT alternative-content block and BrtEndSheet. BrtDrawing
        # ([MS-XLSB] 2.4.351, record 550) sits between them per the Worksheet
        # ABNF: ... [BrtMargins] [BrtPageSetup] [HEADERFOOTER] [RWBRK] [COLBRK]
        # ... [BrtDrawing] ... [LISTPARTS] ... BrtEndSheet.
        if print_spec is None:
            buffer.write(self._sheet1Bytes[290:346])
        else:
            buffer.write(print_records(print_spec))
        if drawing_rel is not None:
            rel_encoded = drawing_rel.encode('utf-16-le')
            buffer.write(build_record(
                _BRT_DRAWING,
                struct.pack('<I', len(drawing_rel)) + rel_encoded,
            ))
        if table_rel_ids:
            if comment_vml_rid is not None:
                buffer.write(build_record(0x227, wide_string(comment_vml_rid)))
            buffer.write(self._list_parts_bin(table_rel_ids))
        elif comment_vml_rid is not None:
            buffer.write(build_record(0x227, wide_string(comment_vml_rid)))
        if sparklines:
            # Future-record blocks follow LISTPARTS, before the trailing uid block and BrtEndSheet.
            buffer.write(sparklines_bin(sparklines, self._worksheet_data[worksheet_index][0],
                                        self._formula_context()))
        if slicer_rel is not None:
            buffer.write(sheet_slicers_bin(slicer_rel))
        buffer.write(self._sheet1Bytes[346:])
        if buffer.tell() > 0:
            sheet_file.write(buffer.getvalue())
