"""XlsxWriter package methods: workbook, relationships, content types, charts, shared strings."""

import io
import xml.sax.saxutils as saxutils
from typing import TYPE_CHECKING, Any
from . import chart_parts
from ._text import xstring_escape
from .comments import comments_xml, vml_xml, COMMENTS_TYPE
from ._extlst import (
    EXT_WORKBOOK_SLICER_CACHES,
    REL_SLICER_CACHE,
    X14,
    X15,
    ext,
    ext_list,
)
from .pivots import (
    pivot_cache_xml,
    pivot_table_rels_xml,
    pivot_table_xml,
)
from .printing import (
    print_area_name,
    print_titles_name,
)
from .refs import quote_sheet_name
from ._xlsx_writer_xml import _ATTR_ENTITIES, _CT_PIVOT_CACHE, _CT_PIVOT_TABLE, _REL_PIVOT_CACHE, _SST_CHUNK_STRINGS


class _XlsxPartsMixin:
    """Methods of the writer class, mixed into it; see the facade module."""
    if TYPE_CHECKING:
        # Attributes of the writer class that these methods use; it sets them.
        _create_styles_xml: Any
        _defined_names: Any
        _doc_props_content_types: Any
        _doc_props_root_rels: Any
        _filtered_data_list: Any
        _internal_links: Any
        _letters: Any
        _media: Any
        _next_chart_id: Any
        _number_format_id: Any
        _pending_pivots: Any
        _pivot_renders: Any
        _shared_strings: Any
        _sheet_chart_ids: Any
        _sheet_charts: Any
        _sheet_comments: Any
        _sheet_count: Any
        _sheet_drawings: Any
        _sheet_image_media: Any
        _sheet_images: Any
        _sheet_print: Any
        _sheet_rels: Any
        _sheet_shapes: Any
        _sheet_slicers: Any
        _sheet_written: Any
        _slicer_cache_names: Any
        _slicer_part_files: Any
        _sst_all_count: Any
        _sst_unique_count: Any
        _tables_xml: Any
        _use_shared_strings: Any
        _workbook_protection: Any
        _workbook_rel_base: Any
        _worksheet_data: Any
        _write_doc_props: Any
        _write_empty_sheet: Any
        _zf: Any

    def _write_workbook_parts(self):
        # Excel compares sheet names without regard to case.
        sheet_names = {name.casefold() for name, _data, _hidden in self._worksheet_data}
        for sheet_index, location in self._internal_links:
            target = location.rpartition("!")[0]
            if len(target) >= 2 and target[0] == target[-1] == "'":
                target = target[1:-1].replace("''", "'")
            if target.casefold() not in sheet_names:
                raise ValueError(f"hyperlink points to unknown sheet {target!r}")

        for sheet_index in range(self._sheet_count):
            if sheet_index not in self._sheet_written:
                self._write_empty_sheet(sheet_index)
        if self._pending_pivots:
            names = ", ".join(sorted(self._pending_pivots))
            raise ValueError(
                f"pivot target sheet(s) {names} were never added to the workbook"
            )

        for render in self._pivot_renders:
            render.numfmt_ids = [
                self._number_format_id(spec.number_format)
                for spec in render.pivot.value_specs
            ]
            cache_path = f"xl/pivotCache/pivotCacheDefinition{render.cache_id}.xml"
            table_path = f"xl/pivotTables/pivotTable{render.pivot_id}.xml"
            self._zf.writestr(cache_path, pivot_cache_xml(render))
            self._zf.writestr(table_path, pivot_table_xml(render, render.numfmt_ids))
            self._zf.writestr(
                f"xl/pivotTables/_rels/pivotTable{render.pivot_id}.xml.rels",
                pivot_table_rels_xml(render.cache_id),
            )

        for sheet_id in sorted(self._sheet_drawings):
            sheet_name = self._worksheet_data[sheet_id - 1][0]
            charts = self._sheet_charts.get(sheet_id, [])
            images = self._sheet_images.get(sheet_id, [])
            media = self._sheet_image_media.get(sheet_id, [])
            chart_ids = []
            for chart in charts:
                chart_id = self._next_chart_id
                self._next_chart_id += 1
                chart_ids.append(chart_id)
                self._zf.writestr(
                    f"xl/charts/chart{chart_id}.xml",
                    self._chart_xml(chart, sheet_name),
                )
            self._sheet_chart_ids[sheet_id] = chart_ids
            self._zf.writestr(
                f"xl/drawings/drawing{sheet_id}.xml",
                chart_parts.drawing_xml(charts, chart_ids, images=images,
                                        shapes=self._sheet_shapes.get(sheet_id, []),
                                        slicers=self._sheet_slicers.get(sheet_id, [])),
            )
            self._zf.writestr(
                f"xl/drawings/_rels/drawing{sheet_id}.xml.rels",
                chart_parts.drawing_rels_xml(chart_ids, image_names=media),
            )

        for sheet_id, comments in self._sheet_comments.items():
            self._zf.writestr(f"xl/comments/comment{sheet_id}.xml", comments_xml(comments))
            self._zf.writestr(f"xl/drawings/commentsDrawing{sheet_id}.vml", vml_xml(comments))

        for name, data in self._media.parts:
            self._zf.writestr(f"xl/media/{name}", data)

        if self._sheet_drawings:
            self._zf.writestr("xl/theme/theme1.xml", chart_parts.THEME_XML)

        for sheet_id, rels in self._sheet_rels.items():
            self._zf.writestr(
                f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels",
                self._create_worksheet_rels(sheet_id),
            )
        for table_id, table_xml in self._tables_xml.items():
            self._zf.writestr(f"xl/tables/table{table_id}.xml", table_xml)
        for path, slicer_xml, _content_type in self._slicer_part_files:
            self._zf.writestr(path, slicer_xml)

        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.xml", self._create_workbook_xml())
        self._zf.writestr("xl/styles.xml", self._create_styles_xml())
        self._zf.writestr("xl/_rels/workbook.xml.rels", self._create_workbook_rels())
        self._write_doc_props()

        if self._use_shared_strings and self._shared_strings:
            with self._zf.open_threaded("xl/sharedStrings.xml") as sst_file:
                self._write_shared_strings_xml(sst_file)

    def _chart_xml(self, chart, sheet_name: str) -> str:
        return chart_parts.chart_xml(self._letters, chart, sheet_name)

    def _create_worksheet_rels(self, sheet_id: int) -> str:
        parts = []
        for rid, rel_type, target, external in self._sheet_rels.get(sheet_id, []):
            mode = ' TargetMode="External"' if external else ""
            parts.append(
                f'<Relationship Id="{rid}" Type="{rel_type}" '
                f'Target="{saxutils.escape(target, {chr(34): "&quot;"})}"{mode}/>'
            )
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(parts)}
</Relationships>'''

    def _create_content_types(self) -> str:
        parts = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(self._sheet_count)
        )

        table_parts = "".join(
            f'<Override PartName="/xl/tables/table{table_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
            for table_id in sorted(self._tables_xml)
        )

        drawing_parts = "".join(
            f'<Override PartName="/xl/drawings/drawing{sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            for sheet_id in sorted(self._sheet_drawings)
        )
        chart_overrides = "".join(
            f'<Override PartName="/xl/charts/chart{chart_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            for chart_id in range(1, self._next_chart_id)
        )
        image_defaults = self._media.content_type_defaults()
        theme_part = ""
        if self._sheet_drawings:
            theme_part = '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'

        shared_strings_part = ""
        comment_parts = ''.join(
            f'<Override PartName="/xl/comments/comment{sheet_id}.xml" ContentType="{COMMENTS_TYPE}"/>'
            for sheet_id in self._sheet_comments
        )
        if self._sheet_comments:
            comment_parts += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
        if self._use_shared_strings and self._shared_strings:
            shared_strings_part = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        slicer_parts = "".join(
            f'<Override PartName="/{path}" ContentType="{content_type}"/>'
            for path, _xml, content_type in self._slicer_part_files
        )
        pivot_parts = "".join(
            f'<Override PartName="/xl/pivotTables/pivotTable{render.pivot_id}.xml" ContentType="{_CT_PIVOT_TABLE}"/>'
            f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{render.cache_id}.xml" ContentType="{_CT_PIVOT_CACHE}"/>'
            for render in self._pivot_renders
        )

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
{image_defaults}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{parts}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{table_parts}
{drawing_parts}
{chart_overrides}
{theme_part}
{shared_strings_part}
{comment_parts}
{slicer_parts}
{pivot_parts}
{self._doc_props_content_types()}
</Types>'''

    def _create_root_rels(self) -> str:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
{self._doc_props_root_rels()}
</Relationships>'''

    def _create_workbook_rels(self) -> str:
        relationships = [
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(self._sheet_count)
        ]
        
        style_rid = self._sheet_count + 1
        relationships.append(f'<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        
        if self._use_shared_strings and self._shared_strings:
            shared_strings_rid = self._sheet_count + 2
            relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')

        if self._sheet_drawings:
            theme_rid = len(relationships) + 1
            relationships.append(f'<Relationship Id="rId{theme_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')

        # Cache i is the (i+1)th relationship after the base count; see _workbook_rel_base.
        for index in range(len(self._slicer_cache_names)):
            relationships.append(
                f'<Relationship Id="rId{self._workbook_rel_base() + index + 1}" '
                f'Type="{REL_SLICER_CACHE}" Target="slicerCaches/slicerCache{index + 1}.xml"/>'
            )

        pivot_base = self._workbook_rel_base() + len(self._slicer_cache_names)
        for index, render in enumerate(self._pivot_renders):
            relationships.append(
                f'<Relationship Id="rId{pivot_base + index + 1}" '
                f'Type="{_REL_PIVOT_CACHE}" '
                f'Target="pivotCache/pivotCacheDefinition{render.cache_id}.xml"/>'
            )

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _create_workbook_xml(self) -> str:
        sheets = ""
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            hidden_attr = ' state="hidden"' if hidden else ''
            sheets += f'<sheet name="{saxutils.escape(sheet_name, _ATTR_ENTITIES)}" sheetId="{sheet_id}"{hidden_attr} r:id="rId{sheet_id}"/>'

        names = []
        for filter_data in self._filtered_data_list:
            sheet_name = self._worksheet_data[filter_data.sheet_index][0]
            start_col = self._letters[filter_data.start_column]
            end_col = self._letters[filter_data.end_column]
            range_ref = (
                f"{quote_sheet_name(sheet_name)}!${start_col}${filter_data.start_row + 1}"
                f":${end_col}${filter_data.end_row + 1}"
            )
            names.append(
                f'<definedName name="_xlnm._FilterDatabase" localSheetId="{filter_data.sheet_index}" hidden="1">'
                f'{saxutils.escape(range_ref)}</definedName>'
            )
        for sheet_index in sorted(self._sheet_print):
            settings = self._sheet_print[sheet_index]
            sheet_name = self._worksheet_data[sheet_index][0]
            if settings.print_area is not None:
                names.append(
                    f'<definedName name="_xlnm.Print_Area" localSheetId="{sheet_index}">'
                    f'{saxutils.escape(print_area_name(sheet_name, settings.print_area))}</definedName>'
                )
            titles = print_titles_name(sheet_name, settings)
            if titles is not None:
                names.append(
                    f'<definedName name="_xlnm.Print_Titles" localSheetId="{sheet_index}">'
                    f'{saxutils.escape(titles)}</definedName>'
                )
        sheet_titles = [name for name, _, _ in self._worksheet_data]
        names.extend(definition.to_xml(sheet_titles) for definition in self._defined_names.values())
        # Excel links each slicer cache to a defined name that holds #N/A.
        names.extend(
            f'<definedName name="{saxutils.escape(cache_name, _ATTR_ENTITIES)}">#N/A</definedName>'
            for cache_name in self._slicer_cache_names
        )
        defined_names = f"<definedNames>{''.join(names)}</definedNames>" if names else ""
        protection = self._workbook_protection.to_xml() if self._workbook_protection else ""
        workbook_ext = self._workbook_extensions()
        pivot_caches = ""
        if self._pivot_renders:
            pivot_base = self._workbook_rel_base() + len(self._slicer_cache_names)
            entries = "".join(
                f'<pivotCache cacheId="{render.cache_id}" '
                f'r:id="rId{pivot_base + index + 1}"/>'
                for index, render in enumerate(self._pivot_renders)
            )
            pivot_caches = f"<pivotCaches>{entries}</pivotCaches>"

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
{protection}
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>{sheets}</sheets>
{defined_names}
<calcPr calcId="124519" fullCalcOnLoad="1"/>
{pivot_caches}
{workbook_ext}
</workbook>'''

    def _workbook_extensions(self) -> str:
        """The workbook ``extLst``: the list of slicer caches, one relationship each."""
        if not self._slicer_cache_names:
            return ""
        base = self._workbook_rel_base()
        caches = "".join(
            f'<x14:slicerCache r:id="rId{base + index + 1}"/>'
            for index in range(len(self._slicer_cache_names))
        )
        return ext_list([
            ext(EXT_WORKBOOK_SLICER_CACHES, f"<x15:slicerCaches>{caches}</x15:slicerCaches>",
                ("x15", X15), ("x14", X14)),
        ])

    def _write_shared_strings_xml(self, sst_file: io.BufferedWriter):
        # Written in chunks: the table can hold millions of strings, and one
        # growing string would keep the whole encoded part in memory.
        header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                  f'count="{self._sst_all_count}" uniqueCount="{self._sst_unique_count}">')
        sst_file.write(header.encode('utf-8'))
        strings = self._shared_strings
        for start in range(0, len(strings), _SST_CHUNK_STRINGS):
            chunk: Any = strings[start:start + _SST_CHUNK_STRINGS]
            # Plain-text chunks are escaped in one pass: joined on NUL (which a
            # string could only hold as a control character), then split again.
            # Rich runs (tuples) make join() fail; those chunks go item by item.
            try:
                joined = '\0'.join(chunk)
            except TypeError:
                joined = None
            if joined is None or joined.count('\0') != len(chunk) - 1:
                self._write_shared_strings_items(sst_file, chunk)
                continue
            body = '</t></si><si><t>'.join(
                joined.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').split('\0')
            )
            # Escaped text holds no '<', so "<si><t>" only ever starts an item.
            body = ('<si><t>' + body + '</t></si>').replace(
                '<si><t> ', '<si><t xml:space="preserve"> ').replace(
                '<si><t>\t', '<si><t xml:space="preserve">\t')
            sst_file.write(xstring_escape(body).encode('utf-8'))
        sst_file.write(b'</sst>')

    def _write_shared_strings_items(self, sst_file, strings) -> None:
        parts: list[str] = []
        for s in strings:
            if type(s) is tuple:  # ("rich", runs) from _write_rich_cell, already escaped
                # The runs already carry their _xHHHH_ escapes: escaping them
                # with the chunk again would turn "_x000D_" into "_x005F_x000D_".
                if parts:
                    sst_file.write(xstring_escape(''.join(parts)).encode('utf-8'))
                    parts.clear()
                sst_file.write(f'<si>{s[1]}</si>'.encode('utf-8'))
                continue
            # The same three replacements as xml.sax.saxutils.escape, minus its call overhead.
            text = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            if s and (s[0] == ' ' or s[0] == '\t'):
                parts.append(f'<si><t xml:space="preserve">{text}</t></si>')
            else:
                parts.append(f'<si><t>{text}</t></si>')
        if parts:
            # The markup holds no escapable characters, so a chunk is escaped at once.
            sst_file.write(xstring_escape(''.join(parts)).encode('utf-8'))
