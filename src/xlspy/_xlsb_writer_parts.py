"""XlsbWriter package methods: content types, relationships, workbook, shared strings, tables, CF."""
from __future__ import annotations

import struct
import io
import re
import uuid
import xml.sax.saxutils as saxutils
from ._accel import encode_xlsb_sst as accel_encode_xlsb_sst
from . import chart_parts
from .styles import (
    parse_cell_range,
)
from ._extlst import REL_SLICER, REL_SLICER_CACHE
from .biff_utils import build_record, iter_records
from ._xlsb_cf import FORMULA_RULES, dxf_payload, rule_bin
from ._text import check_cell_text
from ._xlsb_slicers import (
    CT_SLICER_BIN, CT_SLICER_CACHE_BIN, slicer_cache_bin, slicers_part_bin,
    workbook_slicer_caches_bin,
)
from ._xlsb_pivots import (
    CT_PIVOT_CACHE_BIN, CT_PIVOT_TABLE_BIN, pivot_cache_bin, pivot_table_bin,
    workbook_pivot_caches_bin,
)
from .pivots import (
    pivot_table_rels_xml,
)
from ._xlsb_writer_consts import _ACUID_NAMESPACE, _BRT_AC_BEGIN, _BRT_AC_END, _BRT_BEGIN_AUTOFILTER, _BRT_BEGIN_CF, _BRT_BEGIN_LIST, _BRT_BEGIN_LIST_COL, _BRT_BEGIN_LIST_COLS, _BRT_BEGIN_LIST_PARTS, _BRT_END_AUTOFILTER, _BRT_END_CF, _BRT_END_LIST, _BRT_END_LIST_COL, _BRT_END_LIST_COLS, _BRT_END_LIST_PARTS, _BRT_FRT_BEGIN, _BRT_LIST_PART, _BRT_LIST_TR_FMLA, _BRT_TABLE_STYLE_CLIENT, _REL_PIVOT_CACHE, _REL_PIVOT_TABLE
from ._metadata_part import METADATA_BIN
from ._xlsb_writer_host import _XlsbWriterHost
from .styles import ColorScaleRule, DataBarRule, DifferentialStyle, IconSetRule
from ._xlsb_metadata import comments_bin
from .comments import vml_xml
from ._xlsb_formulas import compile_formula
from ._table_totals import totals_choice
from ._table_totals_names import ILTA_BY_NAME, ILTA_CUSTOM


class _XlsbPartsMixin(_XlsbWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

    def _normalize_cf_rules(self, conditional_formats):
        if not conditional_formats:
            return []
        for rule in conditional_formats:
            self._check_unknown_cf_rule(rule)
            if not isinstance(
                rule, FORMULA_RULES + (ColorScaleRule, DataBarRule, IconSetRule),
            ):
                raise TypeError(
                    f"unsupported conditional formatting rule {type(rule).__name__}"
                )
            rule_dxf = getattr(rule, "dxf", None)
            if rule_dxf is not None and not isinstance(rule_dxf, DifferentialStyle):
                raise TypeError("dxf must be a DifferentialStyle or None")
            if isinstance(rule, FORMULA_RULES) and rule.dxf is not None:
                # Reject styles XLSB cannot hold before anything is recorded.
                dxf_payload(rule.dxf, lambda _fmt: 0)
        return list(conditional_formats)

    @staticmethod
    def _encode_nullable_wstring(value: str | None) -> bytes:
        if value is None:
            return struct.pack("<I", 0xFFFFFFFF)
        encoded = value.encode("utf-16-le")
        return struct.pack("<I", len(value)) + encoded

    @staticmethod
    def _encode_rfx(first_col: int, first_row: int,
                    last_col: int, last_row: int) -> bytes:
        return struct.pack("<IIII", first_row, last_row, first_col, last_col)

    @staticmethod
    def _acuid(key: str) -> bytes:
        # ACBegin / FRTBegin (revision GUID) / ACEnd. Excel writes a random
        # GUID; one derived from the table and column keeps the output the
        # same for the same input. It only has to be unique in the workbook.
        guid = uuid.uuid5(_ACUID_NAMESPACE, key)
        payload = guid.bytes_le
        out = build_record(_BRT_AC_BEGIN, bytes((1, 0, 2, 0x10, 0, 0x80)))
        out += build_record(_BRT_FRT_BEGIN, payload)
        out += build_record(_BRT_AC_END, b"")
        return out

    def _table_bin(self, table, table_id: int, names: list[str], totals_cells=None,
                   formula_context=None, sheet_name=None) -> bytes:
        first_col, first_row, last_col, last_row = parse_cell_range(table.ref)
        totals_row = bool(table.totals_row)
        out = bytearray()
        out.extend(self._acuid(f"{table_id}:{table.display_name}"))
        # BrtBeginList
        payload = bytearray()
        payload.extend(self._encode_rfx(first_col, first_row, last_col, last_row))
        payload.extend(struct.pack("<I", 0))  # lt = LTRANGE
        payload.extend(struct.pack("<I", table_id))
        payload.extend(struct.pack("<I", 1 if table.header_row else 0))
        payload.extend(struct.pack("<I", 1 if totals_row else 0))
        payload.extend(struct.pack("<I", 1 if totals_row else 0))  # fShowTotals
        for _ in range(6):
            payload.extend(struct.pack("<i", -1))  # dxfIds
        payload.extend(struct.pack("<I", 0))  # dwConnID
        # stName and stDisplayName hold the same name in files Excel saves. The
        # display name is what a pivot cache's table source and structured
        # references resolve, so it must equal the name the table is used by.
        display = table.display_name
        payload.extend(self._encode_nullable_wstring(display))
        payload.extend(self._encode_nullable_wstring(display))
        payload.extend(struct.pack("<I", 0))  # stComment empty
        for _ in range(3):
            payload.extend(self._encode_nullable_wstring(None))
        out.extend(build_record(_BRT_BEGIN_LIST, bytes(payload)))
        out.extend(self._acuid(f"{table_id}:{table.display_name}:list"))
        # AutoFilter range inside the table part; the totals row is not filtered.
        # A table without a header row has no filter.
        if table.header_row:
            filter_last_row = last_row - 1 if totals_row else last_row
            out.extend(build_record(
                _BRT_BEGIN_AUTOFILTER,
                self._encode_rfx(first_col, first_row, last_col, filter_last_row),
            ))
            out.extend(build_record(_BRT_END_AUTOFILTER, b""))
        # ListCols
        out.extend(build_record(
            _BRT_BEGIN_LIST_COLS, struct.pack("<I", len(names)),
        ))
        explicit = getattr(table, "totals_row_functions", None)
        for idx, col_name in enumerate(names, start=1):
            cell = totals_cells[idx - 1] if totals_row and totals_cells is not None and idx <= len(totals_cells) else None
            chosen = explicit[idx - 1] if explicit is not None and idx <= len(explicit) else None
            function, label, formula = totals_choice(cell, chosen)
            ilta = ILTA_BY_NAME.get(function, 0)
            out.extend(self._acuid(f"{table_id}:{table.display_name}:{idx}"))
            col_payload = bytearray()
            col_payload.extend(struct.pack("<I", idx))
            col_payload.extend(struct.pack("<I", ilta))
            for _ in range(3):
                col_payload.extend(struct.pack("<i", -1))
            col_payload.extend(struct.pack("<I", 0))  # idqsif
            col_payload.extend(self._encode_nullable_wstring(None))
            col_payload.extend(self._encode_nullable_wstring(col_name))
            col_payload.extend(self._encode_nullable_wstring(label))
            for _ in range(3):
                col_payload.extend(self._encode_nullable_wstring(None))
            out.extend(build_record(_BRT_BEGIN_LIST_COL, bytes(col_payload)))
            if ilta == ILTA_CUSTOM:
                if formula is None or formula_context is None:
                    raise ValueError(
                        f"table {table.display_name!r}: the totals formula needs a formula context")
                parsed = compile_formula(
                    formula, formula_context, sheet=sheet_name,
                    origin=(first_col + idx - 1, last_row),
                )
                out.extend(build_record(_BRT_LIST_TR_FMLA, b"\x00" + parsed))
            out.extend(build_record(_BRT_END_LIST_COL, b""))
        out.extend(build_record(_BRT_END_LIST_COLS, b""))
        # TableStyleClient
        flags = 0
        if table.show_first_column:
            flags |= 0x01
        if table.show_last_column:
            flags |= 0x02
        if table.show_row_stripes:
            flags |= 0x04
        if table.show_column_stripes:
            flags |= 0x08
        style_payload = struct.pack("<H", flags)
        style_payload += self._encode_nullable_wstring(table.style_name)
        out.extend(build_record(_BRT_TABLE_STYLE_CLIENT, style_payload))
        out.extend(build_record(_BRT_END_LIST, b""))
        return bytes(out)

    def _list_parts_bin(self, rel_ids: list[str]) -> bytes:
        out = bytearray()
        out.extend(build_record(
            _BRT_BEGIN_LIST_PARTS, struct.pack("<I", len(rel_ids)),
        ))
        for rel_id in rel_ids:
            encoded = rel_id.encode("utf-16-le")
            out.extend(build_record(
                _BRT_LIST_PART,
                struct.pack("<I", len(rel_id)) + encoded,
            ))
        out.extend(build_record(_BRT_END_LIST_PARTS, b""))
        return bytes(out)

    def _register_dxf(self, dxf) -> int:
        """Index of ``dxf`` in the styles part; equal styles share one BrtDXF."""
        if dxf is None:
            return -1
        payload = dxf_payload(dxf, self._register_numfmt)
        return self._dxf_index.setdefault(payload, len(self._dxf_index))

    def _cf_bin_for_rule(self, rule, priority: int) -> bytes:
        return rule_bin(rule, priority, self._register_dxf(getattr(rule, "dxf", None)),
                        self._formula_context(), self._letters)

    def _cf_bin_grouped(self, rules) -> bytes:
        # Group by sqref like XLSX, claim priorities in order.
        by_range: dict[str, list] = {}
        for rule in rules:
            by_range.setdefault(rule.sqref, []).append(rule)
        out = bytearray()
        for sqref, group in by_range.items():
            ranges = sqref.split()
            # BeginCF: ccf, fPivot=0, then the sqref's ranges.
            payload = bytearray(struct.pack("<III", len(group), 0, len(ranges)))
            for item in ranges:
                payload.extend(self._encode_rfx(*parse_cell_range(item)))
            out.extend(build_record(_BRT_BEGIN_CF, bytes(payload)))
            for rule in group:
                priority = self._claim_cf_priority(rule)
                out.extend(self._cf_bin_for_rule(rule, priority))
            out.extend(build_record(_BRT_END_CF, b""))
        return bytes(out)

    def _write_pivot_parts(self) -> None:
        """Write each pivot's cache and table parts; they need the registered number formats."""
        if self._pending_pivots:
            names = ", ".join(sorted(self._pending_pivots))
            raise ValueError(
                f"pivot target sheet(s) {names} were never written; write each one "
                "with write_sheet (an empty list is enough) after the source sheet"
            )
        for render in self._pivot_renders:
            numfmts = [
                self._register_numfmt(spec.number_format) if spec.number_format else None
                for spec in render.pivot.value_specs
            ]
            self._zf.writestr(
                f"xl/pivotCache/pivotCacheDefinition{render.cache_id}.bin", pivot_cache_bin(render),
            )
            self._zf.writestr(
                f"xl/pivotTables/pivotTable{render.pivot_id}.bin", pivot_table_bin(render, numfmts),
            )
            self._zf.writestr(
                f"xl/pivotTables/_rels/pivotTable{render.pivot_id}.bin.rels",
                pivot_table_rels_xml(render.cache_id, ".bin"),
            )

    def _pivot_cache_rel_ids(self) -> list[str]:
        """Workbook relationship ids of the pivot caches, after the slicer caches."""
        first = self._sheet_count + 4 + len(self._slicer_cache_names)
        return [f"rId{first + index}" for index in range(len(self._pivot_renders))]

    def _write_workbook_parts(self):
        # A sheet that was added but never written still has a part: the content
        # types and workbook relationships name it, and XlsxWriter writes one too.
        for sheet_index in range(self._sheet_count):
            if sheet_index not in self._written_sheets:
                self._write_empty_sheet(sheet_index)
        if self._sheet_comments:
            for sheet_id, comments in self._sheet_comments.items():
                part = f'xl/comments/comment{sheet_id}.bin'
                self._zf.writestr(part, comments_bin(comments, part=part))
                self._zf.writestr(f'xl/drawings/commentsDrawing{sheet_id}.vml', vml_xml(comments))
        for sheet_id in sorted(self._sheet_drawings):
            if sheet_id in self._sheet_chart_ids:
                continue
            charts = self._sheet_charts.get(sheet_id, [])
            sheet_name = self._worksheet_data[sheet_id - 1][0]
            cells = self._sheet_chart_cells.get(sheet_id)
            chart_ids = []
            for chart in charts:
                chart_id = self._next_chart_id
                self._next_chart_id += 1
                chart_ids.append(chart_id)
                self._zf.writestr(
                    f"xl/charts/chart{chart_id}.xml",
                    chart_parts.chart_xml(
                        self._letters, chart, sheet_name,
                        cells=cells, xlsb=True,
                    ),
                )
            self._sheet_chart_ids[sheet_id] = chart_ids
            self._zf.writestr(
                f"xl/drawings/drawing{sheet_id}.xml",
                chart_parts.drawing_xml(charts, chart_ids, images=self._sheet_images.get(sheet_id, []),
                                        shapes=self._sheet_shapes.get(sheet_id, []),
                                        slicers=self._sheet_slicers.get(sheet_id, [])),
            )
            self._zf.writestr(
                f"xl/drawings/_rels/drawing{sheet_id}.xml.rels",
                chart_parts.drawing_rels_xml(chart_ids, image_names=self._sheet_image_media.get(sheet_id, [])),
            )
        for name, data in self._media.parts:
            self._zf.writestr(f"xl/media/{name}", data)

        # Always written: without a theme part Excel opens an XLSB with its
        # current default theme (Aptos Narrow in Microsoft 365) instead of
        # Calibri, which the fonts and the XLSX writer's files use.
        self._zf.writestr("xl/theme/theme1.xml", chart_parts.THEME_XML)

        for table_id, table_bin in self._tables_bin.items():
            self._zf.writestr(f"xl/tables/table{table_id}.bin", table_bin)
        if self._has_dynamic:
            self._zf.writestr("xl/metadata.bin", METADATA_BIN)

        self._write_pivot_parts()
        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.bin", self._create_workbook_bin())
        self._zf.writestr("xl/styles.bin", self._build_styles_bin())
        self._zf.writestr("xl/_rels/workbook.bin.rels", self._create_workbook_rels())
        self._write_doc_props()

        with self._zf.open_threaded("xl/sharedStrings.bin") as sst_file:
            self._write_shared_strings_bin(sst_file)

    def _create_content_types(self) -> str:
        parts = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.bin" ContentType="application/vnd.ms-excel.worksheet"/>'
            f'<Override PartName="/xl/worksheets/binaryIndex{i + 1}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>'
            for i in range(self._sheet_count)
        )

        drawing_parts = "".join(
            f'<Override PartName="/xl/drawings/drawing{sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            for sheet_id in sorted(self._sheet_drawings)
        )
        table_parts = "".join(
            f'<Override PartName="/xl/tables/table{table_id}.bin" ContentType="application/vnd.ms-excel.table"/>'
            for table_id in sorted(self._tables_bin)
        )
        chart_parts_overrides = "".join(
            f'<Override PartName="/xl/charts/chart{chart_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            for chart_id in range(1, self._next_chart_id)
        )
        comment_parts = ''.join(
            f'<Override PartName="/xl/comments/comment{sheet_id}.bin" ContentType="application/vnd.ms-excel.comments"/>'
            for sheet_id in self._sheet_comments
        )
        if self._sheet_comments:
            comment_parts += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
        comment_parts += ''.join(
            f'<Override PartName="/xl/slicers/slicer{part_no}.bin" ContentType="{CT_SLICER_BIN}"/>'
            for part_no, _rel_id in self._sheet_slicer_parts.values()
        ) + ''.join(
            f'<Override PartName="/xl/slicerCaches/slicerCache{index + 1}.bin" ContentType="{CT_SLICER_CACHE_BIN}"/>'
            for index in range(len(self._slicer_cache_names))
        )
        comment_parts += ''.join(
            f'<Override PartName="/xl/pivotTables/pivotTable{render.pivot_id}.bin" ContentType="{CT_PIVOT_TABLE_BIN}"/>'
            f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{render.cache_id}.bin" ContentType="{CT_PIVOT_CACHE_BIN}"/>'
            for render in self._pivot_renders
        )
        theme_part = '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        metadata_part = (
            '<Override PartName="/xl/metadata.bin" ContentType="application/vnd.ms-excel.sheetMetadata"/>'
            if self._has_dynamic else ""
        )

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>{self._media.content_type_defaults()}
<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
{parts}
{drawing_parts}
{table_parts}
{chart_parts_overrides}
{theme_part}
{metadata_part}
{comment_parts}
<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
{self._doc_props_content_types()}
</Types>'''

    def _create_root_rels(self) -> str:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
{self._doc_props_root_rels()}
</Relationships>'''

    def _slicer_cache_rel_ids(self) -> list[str]:
        """Workbook relationship ids of the slicer caches, after sheets, styles, strings and theme."""
        return [f"rId{self._sheet_count + 4 + index}" for index in range(len(self._slicer_cache_names))]

    def _create_workbook_rels(self) -> str:
        relationships = [
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.bin"/>'
            for i in range(self._sheet_count)
        ]
        
        style_rid = self._sheet_count + 1
        shared_strings_rid = self._sheet_count + 2
        relationships.append(f'<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>')
        relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>')
        theme_rid = self._sheet_count + 3
        relationships.append(f'<Relationship Id="rId{theme_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
        relationships.extend(
            f'<Relationship Id="{rel_id}" Type="{REL_SLICER_CACHE}" Target="slicerCaches/slicerCache{index + 1}.bin"/>'
            for index, rel_id in enumerate(self._slicer_cache_rel_ids())
        )
        relationships.extend(
            f'<Relationship Id="{rel_id}" Type="{_REL_PIVOT_CACHE}" '
            f'Target="pivotCache/pivotCacheDefinition{render.cache_id}.bin"/>'
            for render, rel_id in zip(self._pivot_renders, self._pivot_cache_rel_ids())
        )
        if self._has_dynamic:
            relationships.append(
                f'<Relationship Id="rId{len(relationships) + 1}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" '
                'Target="metadata.bin"/>'
            )

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _write_slicer_parts(self, sheet_id, rel_id, slicers, table_columns) -> None:
        """Write the sheet's slicer part and one cache part per slicer.

        Runs after the rows streamed, when the tables' column names are known.
        """
        for slicer in slicers:
            table_id, names = table_columns[slicer.table.casefold()]
            column_id = next(
                (index for index, name in enumerate(names, start=1)
                 if name.casefold() == slicer.column.casefold()),
                None,
            )
            if column_id is None:
                raise ValueError(f"slicer column {slicer.column!r} is not a column of table {slicer.table!r}")
            self._slicer_cache_names.append(slicer.cache_name)
            self._zf.writestr(f"xl/slicerCaches/slicerCache{len(self._slicer_cache_names)}.bin",
                              slicer_cache_bin(slicer, table_id, column_id))
        part_no = len(self._sheet_slicer_parts) + 1
        self._sheet_slicer_parts[sheet_id] = (part_no, rel_id)
        self._zf.writestr(f"xl/slicers/slicer{part_no}.bin", slicers_part_bin(slicers))

    def _create_worksheet_rels(self, sheet_id) -> str:
        tables = self._sheet_tables.get(sheet_id, [])
        rels = (
            '<Relationship Id="rId1" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" '
            f'Target="binaryIndex{sheet_id}.bin"/>'
        )
        for table_id, rel_id in tables:
            rels += (
                f'<Relationship Id="{rel_id}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
                f'Target="../tables/table{table_id}.bin"/>'
            )
        if sheet_id in self._sheet_drawings:
            drawing_rid = f"rId{1 + len(tables) + 1}"
            rels += (
                f'<Relationship Id="{drawing_rid}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/drawing{sheet_id}.xml"/>'
            )
        if self._sheet_comments.get(sheet_id):
            index = 2 + len(tables) + int(sheet_id in self._sheet_drawings)
            rels += (
                f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment{sheet_id}.bin"/>'
                f'<Relationship Id="rId{index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/commentsDrawing{sheet_id}.vml"/>'
            )
        if sheet_id in self._sheet_slicer_parts:
            part_no, rel_id = self._sheet_slicer_parts[sheet_id]
            rels += f'<Relationship Id="{rel_id}" Type="{REL_SLICER}" Target="../slicers/slicer{part_no}.bin"/>'
        for target, rel_id in self._sheet_link_rels.get(sheet_id, {}).items():
            rels += (
                f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                f'Target="{saxutils.escape(target, {chr(34): "&quot;"})}" TargetMode="External"/>'
            )
        # Pivot tables take the ids after every other relationship of the sheet.
        next_id = 1 + max((int(rid) for rid in re.findall(r'Id="rId(\d+)"', rels)), default=0)
        for offset, render in enumerate(self._sheet_pivots.get(sheet_id, [])):
            rels += (
                f'<Relationship Id="rId{next_id + offset}" Type="{_REL_PIVOT_TABLE}" '
                f'Target="../pivotTables/pivotTable{render.pivot_id}.bin"/>'
            )
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}
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
        if self._sheet_print or self._slicer_cache_names:
            # Print areas, titles and slicer links are names with formulas: use the name table.
            self._formula_context()
        if self._xlsb_formula_context is None:
            self._write_filter_defined_names(data)
        else:
            self._register_filter_names()
            self._register_print_names()
            self._register_slicer_names()
        end = self._workbookBinEnd
        pivot_caches = workbook_pivot_caches_bin([
            (render.cache_id, rel_id)
            for render, rel_id in zip(self._pivot_renders, self._pivot_cache_rel_ids())
        ])
        if pivot_caches:
            # The pivot cache list sits between BrtCalcProp and BrtFileRecover.
            split = next(record.header_start for record in iter_records(end) if record.record_id == 0x9B)
            end = end[:split] + pivot_caches + end[split:]
        caches = workbook_slicer_caches_bin(self._slicer_cache_rel_ids())
        if caches:
            # The slicer cache list follows BrtFileRecover, before the other FRT blocks.
            split = next(record.data_end for record in iter_records(end) if record.record_id == 0x9B)
            end = end[:split] + caches + end[split:]
        data.extend(end)
        if self._xlsb_formula_context is not None or self._workbook_protection is not None:
            return self._metadata_workbook(bytes(data))
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
        c_sst = None if self._sst_has_rich else accel_encode_xlsb_sst(
            self._shared_strings, self._sst_all_count, self._sst_unique_count
        )
        if c_sst is not None:
            view = memoryview(c_sst)
            for offset in range(0, len(view), 65536):
                sst_file.write(view[offset:offset + 65536])
            return
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536

        buffer.write(b'\x9F\x01\x08')
        buffer.write(struct.pack('<II', self._sst_all_count, self._sst_unique_count))
        
        for s in self._shared_strings:
            if isinstance(s, tuple):  # ("rich", text, runs) from _write_rich_str
                buffer.write(build_record(0x13, self._rich_str_payload(s[1], s[2])))
                if buffer.tell() > BUFFER_SIZE:
                    sst_file.write(buffer.getvalue())
                    buffer.seek(0)
                    buffer.truncate()
                continue
            # An over-limit string is an error, never a cut: the cell writers
            # check every value before it reaches this table.
            check_cell_text(s)
            encoded = s.encode('utf-16-le')
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
