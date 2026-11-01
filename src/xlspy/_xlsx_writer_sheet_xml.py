"""XlsxWriter sheet-level XML: head, views, columns, tail, tables, hyperlinks, page setup."""

from typing import TYPE_CHECKING, Any
import xml.sax.saxutils as saxutils
from ._writer_common import FilterData
from ._text import xstring_escape
from .data_validation import data_validation_xml
from ._widths import estimate_column_widths, file_width
from ._extlst import (
    CT_SLICER,
    CT_SLICER_CACHE,
)
from .slicers import slicer_cache_xml, slicer_part_xml
from .printing import (
    _paper_code,
)
from .styles import (
    ExcelTable,
    parse_cell_range,
    table_column_names,
    styles_for_row,
)
from ._xlsx_writer_xml import _DEFAULT_MARGINS, _REL_HYPERLINK, _REL_TABLE, _format_number, _outline_col_attr


class _XlsxSheetXmlMixin:
    """Methods of the writer class, mixed into it; see the facade module."""
    if TYPE_CHECKING:
        # Attributes of the writer class that these methods use; it sets them.
        _cf_rule_xml: Any
        _color_attrs: Any
        _filtered_data_list: Any
        _letters: Any
        _next_slicer_part: Any
        _next_table_id: Any
        _shared_strings: Any
        _sheet_count: Any
        _sheet_drawings: Any
        _slicer_cache_names: Any
        _slicer_part_files: Any
        _table_info: Any
        _tables_xml: Any
        _useSharedStrings: Any


    def _table_xml(self, table: ExcelTable, table_id: int, names: list[str]) -> str:
        first_col, first_row, last_col, last_row = parse_cell_range(table.ref)
        ref = (
            f"{self._letters[first_col]}{first_row + 1}:"
            f"{self._letters[last_col]}{last_row + 1}"
        )
        name = saxutils.escape(table.display_name, {'"': '&quot;'})
        # Names must equal the header cells as Excel reads them, so they get the
        # same _xHHHH_ escapes; line feeds and tabs need character references
        # in an attribute, where XML would read them as spaces.
        attr_entities = {'"': '&quot;', '\n': '&#10;', '\t': '&#9;'}
        columns_xml = "".join(
            f'<tableColumn id="{i + 1}" name="{xstring_escape(saxutils.escape(nm, attr_entities))}"/>'
            for i, nm in enumerate(names)
        )
        totals = ' totalsRowShown="1"' if table.totals_row else ""
        header_count = "1" if table.header_row else "0"
        if table.header_row:
            auto = f'<autoFilter ref="{ref}"/>'
        else:
            auto = ""
        style_name = saxutils.escape(table.style_name, {'"': "&quot;"})
        style_xml = (
            f'<tableStyleInfo name="{style_name}" '
            f'showFirstColumn="{1 if table.show_first_column else 0}" '
            f'showLastColumn="{1 if table.show_last_column else 0}" '
            f'showRowStripes="{1 if table.show_row_stripes else 0}" '
            f'showColumnStripes="{1 if table.show_column_stripes else 0}"/>'
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'id="{table_id}" name="{name}" displayName="{name}" ref="{ref}"{totals} '
            f'headerRowCount="{header_count}">{auto}'
            f'<tableColumns count="{len(names)}">{columns_xml}</tableColumns>'
            f"{style_xml}</table>"
        )

    def _register_tables(self, table_specs, sheet_id, body):
        """Assign table ids and relationships; returns ``(rels, bounds)``."""
        rels = []
        bounds = [parse_cell_range(table.ref) for table in table_specs]
        sheet_tables = []
        for position, (table, (c1, r1, c2, r2)) in enumerate(zip(table_specs, bounds), start=1):
            table_id = self._next_table_id
            self._next_table_id += 1
            body.table_refs.add(f"{self._letters[c1]}{r1 + 1}:{self._letters[c2]}{r2 + 1}")
            names = table_column_names(table, body.header_rows.get(r1))
            self._table_info[table.display_name.casefold()] = (table_id, list(names))
            self._tables_xml[table_id] = self._table_xml(table, table_id, names)
            rid = f"rId{position}"
            rels.append((rid, _REL_TABLE, f"../tables/table{table_id}.xml", False))
            sheet_tables.append((rid, table_id))
        body.tables = sheet_tables
        return rels, bounds

    def _resolve_slicers(self, slicer_specs):
        """Write the slicer and cache parts; returns ``[(slicer, part number)]``.

        Runs once the sheet's tables are registered, so their ids and column
        names are known. A header-derived column name is only known now, so a
        missing column is reported here rather than before the rows.
        """
        resolved = []
        for slicer in slicer_specs:
            table_id, names = self._table_info[slicer.table.casefold()]
            column_id = next(
                (index + 1 for index, name in enumerate(names)
                 if name.casefold() == slicer.column.casefold()),
                None,
            )
            if column_id is None:
                raise ValueError(
                    f"slicer column {slicer.column!r} is not a column of table {slicer.table!r}"
                )
            part_no = self._next_slicer_part
            self._next_slicer_part += 1
            self._slicer_part_files.append((
                f"xl/slicers/slicer{part_no}.xml", slicer_part_xml(slicer), CT_SLICER,
            ))
            self._slicer_part_files.append((
                f"xl/slicerCaches/slicerCache{part_no}.xml",
                slicer_cache_xml(slicer, table_id, column_id), CT_SLICER_CACHE,
            ))
            self._slicer_cache_names.append(slicer.cache_name)
            resolved.append((slicer, part_no))
        return resolved

    def _workbook_rel_base(self) -> int:
        """Relationship count before the slicer caches: sheets, styles, strings, theme."""
        count = self._sheet_count + 1
        if self._useSharedStrings and self._shared_strings:
            count += 1
        if self._sheet_drawings:
            count += 1
        return count

    def _link_relationships(self, hyperlinks, rels):
        """External hyperlink relationships; ``rels`` supplies the next free ids."""
        link_rels = []
        seen: dict[str, str] = {}
        next_index = len(rels) + 1
        for _row, _col, link in hyperlinks:
            if link.is_internal or link.target in seen:
                continue
            rid = f"rId{next_index}"
            next_index += 1
            seen[link.target] = rid
            link_rels.append((rid, _REL_HYPERLINK, link.target, True))
        return link_rels

    def _sheet_head(self, body, sheet_index, split_freeze, freeze_panes, widths,
                    hidden_cols, print_spec, row_outline=None, col_outline=None,
                    show_grid_lines=True, width_styles=((), None), filter_cells=(),
                    tab_key=None, zoom=None, view=(True, True, False)):
        """Everything before ``<sheetData>`` content: properties, views, widths."""
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        sheet_pr = ""
        if tab_key is not None:
            sheet_pr += f'<tabColor {self._color_attrs(tab_key)}/>'
        if print_spec is not None and print_spec.fit_to_page:
            sheet_pr += '<pageSetUpPr fitToPage="1"/>'
        if sheet_pr:
            parts.append(f'<sheetPr>{sheet_pr}</sheetPr>')
        if body.num_rows > 0 and body.max_cols > 0:
            parts.append(f'<dimension ref="A1:{self._letters[body.max_cols - 1]}{body.num_rows}"/>')
        parts.append(self._sheet_views_xml(
            sheet_index, split_freeze, freeze_panes, show_grid_lines, zoom, view,
        ))
        max_row_level = row_outline.max_level if row_outline is not None else 0
        max_col_level = col_outline.max_level if col_outline is not None else 0
        if max_row_level or max_col_level:
            level_attrs = ""
            if max_row_level:
                level_attrs += f' outlineLevelRow="{max_row_level}"'
            if max_col_level:
                level_attrs += f' outlineLevelCol="{max_col_level}"'
            parts.append(f'<sheetFormatPr defaultRowHeight="15"{level_attrs}/>')
        style_ranges, header_style = width_styles
        col_widths = estimate_column_widths(
            body.head, body.max_cols,
            lambda row_idx, length: styles_for_row(row_idx, length, style_ranges, header_style),
            filter_cells,
        )
        for col, width in widths.items():
            if 0 <= col < body.max_cols:
                col_widths[col] = width
        parts.append(self._cols_xml(col_widths, widths, hidden_cols, body.max_cols, col_outline))
        parts.append('<sheetData>')
        return "".join(parts)

    def _sheet_views_xml(self, sheet_index, split_freeze, freeze_panes,
                         show_grid_lines=True, zoom=None, view=(True, True, False)) -> str:
        headers, zeros, right_to_left = view
        # CT_SheetView attribute order: showGridLines, showRowColHeaders, showZeros,
        # rightToLeft, tabSelected, ..., zoomScale.
        flags = "".join((
            "" if show_grid_lines else ' showGridLines="0"',
            "" if headers else ' showRowColHeaders="0"',
            "" if zeros else ' showZeros="0"',
            ' rightToLeft="1"' if right_to_left else "",
        ))
        tab = flags + (' tabSelected="1"' if sheet_index == 0 else "")
        if zoom is not None:
            tab += f' zoomScale="{zoom}" zoomScaleNormal="{zoom}"'
        if split_freeze is None and freeze_panes is None:
            return (
                f'<sheetViews><sheetView{tab} workbookViewId="0">'
                '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
            )
        if split_freeze is None:
            return f'<sheetViews><sheetView{tab} workbookViewId="0"/></sheetViews>'
        split_col, split_row = split_freeze
        top_left = f"{self._letters[split_col]}{split_row + 1}"
        parts = []
        if split_col > 0:
            parts.append(f'xSplit="{split_col}"')
        if split_row > 0:
            parts.append(f'ySplit="{split_row}"')
        parts.append(f'topLeftCell="{top_left}"')
        if split_col > 0 and split_row > 0:
            active = "bottomRight"
        elif split_col > 0:
            active = "topRight"
        else:
            active = "bottomLeft"
        parts.append(f'activePane="{active}"')
        parts.append('state="frozen"')
        return (
            f'<sheetViews><sheetView{tab} workbookViewId="0">'
            f'<pane {" ".join(parts)}/>'
            f'<selection pane="{active}" activeCell="{top_left}" sqref="{top_left}"/>'
            "</sheetView></sheetViews>"
        )

    def _cols_xml(self, col_widths, widths, hidden_cols, max_cols, col_outline=None) -> str:
        covered = set(widths) | set(hidden_cols)
        if col_outline is not None:
            covered |= col_outline.indices()
        if not (any(width != 8.43 for width in col_widths) or covered):
            return ""
        parts = ['<cols>']
        emitted = set()
        for i, width in enumerate(col_widths):
            explicit = widths.get(i)
            hidden_attr = ' hidden="1"' if i in hidden_cols else ""
            outline_attr = _outline_col_attr(col_outline, i)
            if explicit is not None:
                parts.append(f'<col min="{i + 1}" max="{i + 1}" width="{explicit:.2f}" bestFit="1" customWidth="1"{hidden_attr}{outline_attr}/>')
                emitted.add(i)
            elif width != 8.43:
                # Automatic widths count characters; the file adds the cell padding.
                parts.append(f'<col min="{i + 1}" max="{i + 1}" width="{file_width(width):.2f}" bestFit="1" customWidth="1"{hidden_attr}{outline_attr}/>')
                emitted.add(i)
            elif hidden_attr or outline_attr:
                parts.append(f'<col min="{i + 1}" max="{i + 1}" width="8.43" customWidth="0"{hidden_attr}{outline_attr}/>')
                emitted.add(i)
        # Columns past the data (explicit widths, hidden flags or groups) come
        # last, which keeps the <col> entries in ascending order.
        for i in sorted(covered):
            if i in emitted:
                continue
            explicit = widths.get(i)
            hidden_attr = ' hidden="1"' if i in hidden_cols else ""
            outline_attr = _outline_col_attr(col_outline, i)
            width_attr = f"{explicit:.2f}" if explicit is not None else "8.43"
            custom = "1" if explicit is not None else "0"
            parts.append(f'<col min="{i + 1}" max="{i + 1}" width="{width_attr}" customWidth="{custom}"{hidden_attr}{outline_attr}/>')
        parts.append('</cols>')
        return "".join(parts)

    def _sheet_tail(self, body, sheet_index, merge_refs, cf_rules, table_specs,
                    table_bounds, validations, link_rels, auto_filter, drawing_rid,
                    print_spec, *, auto_filter_ref=None, header_rows=1):
        """Everything after the row data, in schema order."""
        parts = ['</sheetData>']
        data_filter = None
        if body.num_rows > header_rows and body.max_cols > 0:
            # The filter buttons sit on the last header row.
            data_filter = (
                f"{self._letters[0]}{header_rows}:"
                f"{self._letters[body.max_cols - 1]}{body.num_rows}"
            )
        if auto_filter_ref is not None:
            data_filter = auto_filter_ref
        if auto_filter is False:
            data_filter = None
        if data_filter is not None and data_filter not in body.table_refs:
            parts.append(f'<autoFilter ref="{data_filter}"/>')
            first_col, first_row, last_col, last_row = parse_cell_range(data_filter)
            self._filtered_data_list.append(FilterData(
                sheet_index=sheet_index,
                start_column=first_col,
                end_column=last_col,
                start_row=first_row,
                end_row=last_row,
            ))
        if merge_refs:
            parts.append(f'<mergeCells count="{len(merge_refs)}">')
            for ref in merge_refs:
                parts.append(f'<mergeCell ref="{ref}"/>')
            parts.append('</mergeCells>')
        if cf_rules:
            by_range: dict[str, list] = {}
            for rule in cf_rules:
                by_range.setdefault(rule.sqref, []).append(rule)
            for sqref, rules in by_range.items():
                parts.append(f'<conditionalFormatting sqref="{sqref}">')
                for rule in rules:
                    parts.append(self._cf_rule_xml(rule, set(table_bounds)))
                parts.append('</conditionalFormatting>')
        if validations:
            parts.append(f'<dataValidations count="{len(validations)}">')
            parts.extend(data_validation_xml(rule) for rule in validations)
            parts.append('</dataValidations>')
        if body.hyperlinks:
            parts.append(self._hyperlinks_xml(body.hyperlinks, link_rels))
        if print_spec is not None:
            if print_spec.center_horizontally or print_spec.center_vertically:
                attrs = ""
                if print_spec.center_horizontally:
                    attrs += ' horizontalCentered="1"'
                if print_spec.center_vertically:
                    attrs += ' verticalCentered="1"'
                parts.append(f"<printOptions{attrs}/>")
        parts.append(self._page_margins_xml(print_spec))
        if print_spec is not None:
            setup = self._page_setup_xml(print_spec)
            if setup:
                parts.append(setup)
            header_footer = self._header_footer_xml(print_spec)
            if header_footer:
                parts.append(header_footer)
            if print_spec.row_breaks:
                parts.append(self._breaks_xml("rowBreaks", print_spec.row_breaks, 16383))
            if print_spec.col_breaks:
                parts.append(self._breaks_xml("colBreaks", print_spec.col_breaks, 1048575))
        if drawing_rid is not None:
            parts.append(f'<drawing r:id="{drawing_rid}"/>')
        if table_specs:
            table_rels = body.tables
            parts.append(f'<tableParts count="{len(table_rels)}">')
            for rid, _table_id in table_rels:
                parts.append(f'<tablePart r:id="{rid}"/>')
            parts.append('</tableParts>')
        parts.append('</worksheet>')
        return "".join(parts)

    def _hyperlinks_xml(self, hyperlinks, link_rels) -> str:
        rid_by_target = {target: rid for rid, _type, target, _ext in link_rels}
        parts = [f'<hyperlinks>']
        for row_idx, col_idx, link in hyperlinks:
            ref = f"{self._letters[col_idx]}{row_idx + 1}"
            attrs = f'ref="{ref}"'
            if link.is_internal:
                attrs += f' location="{saxutils.escape(link.location, {chr(34): "&quot;"})}"'
            else:
                attrs += f' r:id="{rid_by_target[link.target]}"'
            if link.tooltip:
                attrs += f' tooltip="{saxutils.escape(link.tooltip, {chr(34): "&quot;"})}"'
            parts.append(f"<hyperlink {attrs}/>")
        parts.append('</hyperlinks>')
        return "".join(parts)

    @staticmethod
    def _page_margins_xml(print_spec) -> str:
        margins = _DEFAULT_MARGINS
        if print_spec is not None and print_spec.margins is not None:
            margins = print_spec.margins
        return (
            f'<pageMargins left="{_format_number(margins.left)}" right="{_format_number(margins.right)}" '
            f'top="{_format_number(margins.top)}" bottom="{_format_number(margins.bottom)}" '
            f'header="{_format_number(margins.header)}" footer="{_format_number(margins.footer)}"/>'
        )

    @staticmethod
    def _page_setup_xml(print_spec) -> str:
        attrs = []
        if print_spec.paper_size is not None:
            attrs.append(f'paperSize="{_paper_code(print_spec.paper_size)}"')
        if print_spec.fit_to_page:
            width = print_spec.fit_to_width if print_spec.fit_to_width is not None else 1
            height = print_spec.fit_to_height if print_spec.fit_to_height is not None else 1
            attrs.append(f'fitToWidth="{width}" fitToHeight="{height}"')
        elif print_spec.scale is not None:
            attrs.append(f'scale="{print_spec.scale}"')
        if print_spec.orientation is not None:
            attrs.append(f'orientation="{print_spec.orientation}"')
        if not attrs:
            return ""
        return f'<pageSetup {" ".join(attrs)}/>'

    @staticmethod
    def _header_footer_xml(print_spec) -> str:
        parts = []
        if print_spec.header is not None:
            parts.append(f"<oddHeader>{saxutils.escape(print_spec.header)}</oddHeader>")
        if print_spec.footer is not None:
            parts.append(f"<oddFooter>{saxutils.escape(print_spec.footer)}</oddFooter>")
        if not parts:
            return ""
        return f"<headerFooter>{''.join(parts)}</headerFooter>"

    @staticmethod
    def _breaks_xml(tag: str, breaks, max_value: int) -> str:
        items = "".join(f'<brk id="{value}" max="{max_value}" man="1"/>' for value in breaks)
        return f'<{tag} count="{len(breaks)}" manualBreakCount="{len(breaks)}">{items}</{tag}>'
