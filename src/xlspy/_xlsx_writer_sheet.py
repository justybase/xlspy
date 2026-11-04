"""XlsxWriter sheet methods: sheet loop, rows, cells, pivot and empty sheets."""
from __future__ import annotations

from typing import Any
import datetime
from decimal import Decimal
import io
import math
import shutil
import tempfile
import xml.sax.saxutils as saxutils
from collections.abc import Iterable, Mapping
from . import _zip_threads
from ._accel import encode_xlsx_row as accel_encode_xlsx_row
from ._text import check_cell_text, xml_space_attribute, xstring_escape
from .charts import Chart
from .data_validation import DataValidation
from .excel_dates import datetime_to_serial
from .formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT, _is_formatted_cell, _unwrap_cell
from .formulas import Formula
from .comments import Comment
from ._metadata import SheetProtection, replace_element
from .ignored_errors import ignored_errors_xml
from ._extlst import (
    EXT_SHEET_SLICER_LIST,
    REL_SLICER,
    X14,
    X15,
    ext,
    ext_list,
)
from ._writer_common import SheetSpec
from ._sheet_options import XLSX_OPTIONS, SheetOptions, resolve_sheet_options
from .sparklines import sparklines_ext
from .pivots import (
    PivotFeeder,
    merge_cells,
    normalize_pivots,
    overlay_rows,
    tee_rows,
)
from .hyperlinks import Hyperlink
from ._scalars import is_pandas_missing, native_cell_value, plain_value
from .images import Image
from .printing import (
    PrintSettings,
)
from .rich_text import CellRichText
from .styles import (
    HYPERLINK_STYLE,
    Color,
    ExcelTable,
    Style,
    _date_cell_xf,
    _record_cell_xf,
    _style_ids_for_row,
    _xf_for_style_format,
    capture_table_headers,
    capture_table_totals,
    merge_styles,
    parse_cell_range,
    styles_for_row,
)
from ._widths import ColumnWidthMeter
from ._xlsx_writer_xml import _EMPTY_ROW_STYLES, _FLUSH_BYTES, _REL_DRAWING, _REL_PIVOT_TABLE, _REL_WORKSHEET, _SheetBody, _Tee, _format_number
from ._xlsx_writer_host import _XlsxWriterHost
from ._dataframe_rows import dataframe_to_rows


def _filter_cells(
    tables: list[Any], table_bounds: list[Any], auto_filter: bool | None,
    auto_filter_ref: str | None, body: Any, header_rows: int,
) -> list[tuple[int, int, int]]:
    """The header cells under a filter drop-down button, as ``(row, first column,
    last column)``: they need room for it."""
    filter_cells = [
        (r1, c1, c2) for table, (c1, r1, c2, _r2) in zip(tables, table_bounds)
        if table.header_row
    ]
    if auto_filter is not False:
        if auto_filter_ref is not None:
            c1, r1, c2, _r2 = parse_cell_range(auto_filter_ref)
            filter_cells.append((r1, c1, c2))
        elif body.num_rows > header_rows and body.value_cols > 0:
            filter_cells.append((header_rows - 1, 0, body.value_cols - 1))
    return filter_cells


# Cells the native encoder's hook leaves to the Python writer.
_XLSX_PYTHON_CELLS = (Formula, Hyperlink, CellRichText, datetime.date)


class _XlsxSheetMixin(_XlsxWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""


    def write_sheet(self, data: Iterable[list[Any]], *, options: SheetOptions | None = None,
                    **kwargs: Any) -> None:
        """Write rows to the current sheet.

        ``data`` is consumed once, row by row. Sheet settings come from ``options``
        (a ``SheetOptions``) or from keyword arguments with the same names; a
        keyword argument overrides the matching field of ``options``. The rows are
        encoded into a temporary spool (kept in memory up to ``spool_max_bytes``,
        then on disk), so memory does not grow with the number of rows. Hyperlinks,
        formulas and the other options are described in the documentation.
        """
        opts = resolve_sheet_options(options, kwargs, XLSX_OPTIONS, "XlsxWriter")
        self._write_sheet(data, **opts.as_kwargs(XLSX_OPTIONS))

    def _write_sheet(
        self,
        data: Iterable[list[Any]],
        *,
        header_style: Style | None = None,
        styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        conditional_formats: list | tuple | None = None,
        tables: list[ExcelTable] | tuple[ExcelTable, ...] | None = None,
        column_widths: dict | None = None,
        autofit_rows: int | None = 1000,
        autofit_padding: float = 0.0,
        hidden_columns: list | tuple | set | None = None,
        row_heights: dict | None = None,
        hidden_rows: list | tuple | set | None = None,
        row_groups: list | tuple | None = None,
        column_groups: list | tuple | None = None,
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        auto_filter_ref: str | None = None,
        charts: list[Chart] | tuple[Chart, ...] | None = None,
        data_validations: list[DataValidation] | tuple[DataValidation, ...] | None = None,
        images: list[Image] | tuple[Image, ...] | None = None,
        print_settings: PrintSettings | None = None,
        comments: Mapping[str, Comment] | None = None,
        sheet_protection: SheetProtection | None = None,
        show_grid_lines: bool = True,
        header_rows: int = 1,
        tab_color: Color | str | None = None,
        zoom: int | None = None,
        ignored_errors: list | tuple | None = None,
        sparklines: list | tuple | None = None,
        slicers: list | tuple | None = None,
        pivots: list | tuple | None = None,
        shapes: list | tuple | None = None,
        show_row_col_headers: bool = True,
        show_zero_values: bool = True,
        right_to_left: bool = False,
    ):
        """Write rows to the current sheet.

        ``data`` is consumed once, row by row. The rows are encoded into a
        temporary spool (kept in memory up to ``spool_max_bytes``, then on
        disk), so memory does not grow with the number of rows. Hyperlinks,
        formulas and the other options are described in the documentation.
        """
        self._require_writable()
        spec = self._check_sheet_options(
            header_style=header_style, show_grid_lines=show_grid_lines, header_rows=header_rows,
            tab_color=tab_color, zoom=zoom, show_row_col_headers=show_row_col_headers,
            show_zero_values=show_zero_values, right_to_left=right_to_left, styles=styles,
            merges=merges, comments=comments, sheet_protection=sheet_protection,
            auto_filter_ref=auto_filter_ref, conditional_formats=conditional_formats, tables=tables,
            sparklines=sparklines, slicers=slicers, charts=charts, data_validations=data_validations,
            images=images, shapes=shapes, print_settings=print_settings, column_widths=column_widths,
            hidden_columns=hidden_columns, row_heights=row_heights, hidden_rows=hidden_rows,
            row_groups=row_groups, column_groups=column_groups, freeze_panes=freeze_panes,
            autofit_rows=autofit_rows, autofit_padding=autofit_padding,
        )
        # Every option is valid; only now record what this sheet claims. A second
        # write to the same sheet is refused after the options are checked, so a
        # bad argument still reports its own error.
        self._reject_second_write()
        self._claim_sheet_options(spec)
        if not self._worksheet_data:
            self.add_sheet("Sheet1")

        sheet_index = self._sheet_count - 1
        sheet_name, _, hidden = self._worksheet_data[sheet_index]
        self._worksheet_data[sheet_index] = (sheet_name, None, hidden)
        sheet_id = self._sheet_count
        has_styles = bool(spec.style_ranges) or spec.header_style is not None

        pivot_plans = normalize_pivots(pivots, spec.tables, self._pivot_names, sheet_name)
        for plan in pivot_plans:
            target_index = self._sheet_index_by_name(plan.target_sheet)
            if (
                target_index is not None
                and target_index != sheet_index
                and target_index in self._sheet_written
            ):
                raise ValueError(
                    f"pivot target sheet {plan.target_sheet!r} was written before its "
                    "source sheet; write the pivot's sheet after the source sheet"
                )
        feeders = [PivotFeeder(plan) for plan in pivot_plans]

        spool = tempfile.SpooledTemporaryFile(max_size=self._spool_max_bytes, mode="w+b")
        # The rows are compressed on a worker thread while they are encoded;
        # the head, which depends on them, is compressed separately at the end.
        compressed = segment = None
        if self._zf.threaded_deflate():
            compressed = tempfile.SpooledTemporaryFile(
                max_size=self._spool_max_bytes, mode="w+b")
            segment = _zip_threads.DeflateSegment(
                self._zf.deflate_pool(), self._zf.compresslevel, compressed)
        try:
            row_sink = spool if segment is None else _Tee(spool, segment)
            pending = self._pending_pivots.pop(sheet_name, None)
            # Feeders read the sheet's own rows; pivot output from other sheets is merged after.
            if feeders:
                data = tee_rows(data, feeders)
            if pending:
                data = overlay_rows(data, pending)
            button_rows = {spec.header_rows - 1, *(parse_cell_range(t.ref)[1] for t in spec.tables)}
            if auto_filter_ref is not None:
                button_rows.add(parse_cell_range(auto_filter_ref)[1])
            width_styles = (spec.style_ranges, spec.header_style if has_styles else None)
            meter = ColumnWidthMeter(
                lambda r, length: styles_for_row(r, length, *width_styles), spec.autofit[0], button_rows,
            )
            table_totals: dict = {}
            body = self._write_body(
                data, row_sink,
                spec.style_ranges, spec.header_style if has_styles else None,
                spec.heights, spec.hidden_rows, has_styles,
                spec.tables,
                row_outline=spec.row_outline, meter=meter, table_totals=table_totals,
            )
            same_sheet = self._place_pivots(feeders, sheet_name, row_sink, body)
            if segment is not None:
                segment.finish()
                segment = None
            table_rels, table_bounds = self._register_tables(spec.tables, sheet_id, body, table_totals)
            rels, link_rels, drawing_rid, slicer_rids = self._register_sheet_parts(
                sheet_id, sheet_index, spec, body, table_rels, (pending or []) + same_sheet,
            )
            if auto_filter is None and spec.tables:
                auto_filter = False
            filter_cells = _filter_cells(
                spec.tables, table_bounds, auto_filter, auto_filter_ref, body, spec.header_rows,
            )
            head = self._sheet_head(
                body, sheet_index, spec.split_freeze, freeze_panes, spec.widths, spec.hidden_cols,
                spec.print_settings, row_outline=spec.row_outline, col_outline=spec.col_outline,
                show_grid_lines=show_grid_lines,
                meter=meter, autofit_padding=spec.autofit[1],
                filter_cells=filter_cells, tab_key=spec.tab_key, zoom=spec.zoom, view=spec.view,
            )
            tail = self._sheet_tail(
                body, sheet_index, spec.merge_refs, spec.cf_rules, spec.tables, table_bounds,
                spec.validations, link_rels, auto_filter, drawing_rid, spec.print_settings,
                auto_filter_ref=auto_filter_ref, header_rows=spec.header_rows,
            )
            tail = self._finish_tail(
                tail, rels, sheet_id, sheet_name, spec, sheet_protection, ignored_errors, slicer_rids,
            )
            if rels:
                self._sheet_rels[sheet_id] = rels
            spool.seek(0)
            if compressed is not None:
                compressed.seek(0)
                _zip_threads.replay_member(
                    self._zf, f"xl/worksheets/sheet{sheet_id}.xml",
                    head.encode('utf-8'), spool, compressed, tail.encode('utf-8'),
                )
            else:
                with self._zf.open(f"xl/worksheets/sheet{sheet_id}.xml", 'w') as sheet_file:
                    sheet_file.write(head.encode('utf-8'))
                    shutil.copyfileobj(spool, sheet_file, 1 << 20)
                    sheet_file.write(tail.encode('utf-8'))
            self._sheet_written.add(sheet_index)
        except BaseException:
            # The sheet part is incomplete, so the workbook cannot be saved.
            self._aborted = True
            if segment is not None:
                segment.cancel()
            raise
        finally:
            spool.close()
            if compressed is not None:
                compressed.close()

    def _register_sheet_parts(
        self, sheet_id: int, sheet_index: int, spec: SheetSpec, body: Any,
        table_rels: list[tuple[str, str, str, bool]], pivots: list[Any],
    ) -> tuple[list[tuple[str, str, str, bool]], list[Any], str | None, list[str]]:
        """Record the parts a written sheet refers to (hyperlinks, pivots, drawing,
        slicers, shapes, charts, images, print settings) and number their
        relationships after the tables'. Returns the relationships, the hyperlink
        ones, the drawing's id and the slicers' ids."""
        rels = list(table_rels)
        link_rels = self._link_relationships(body.hyperlinks, rels)
        rels.extend(link_rels)
        for render in pivots:
            rid = f"rId{len(rels) + 1}"
            rels.append((
                rid, _REL_PIVOT_TABLE,
                f"../pivotTables/pivotTable{render.pivot_id}.xml", False,
            ))
        slicer_parts = self._resolve_slicers(spec.slicers)
        drawing_rid = None
        if spec.charts or spec.images or spec.shapes or slicer_parts:
            drawing_rid = f"rId{len(rels) + 1}"
            rels.append((drawing_rid, _REL_DRAWING,
                         f"../drawings/drawing{sheet_id}.xml", False))
            self._sheet_drawings.add(sheet_id)
        slicer_rids = []
        for _slicer, part_no in slicer_parts:
            rid = f"rId{len(rels) + 1}"
            rels.append((rid, REL_SLICER, f"../slicers/slicer{part_no}.xml", False))
            slicer_rids.append(rid)
        if slicer_parts:
            self._sheet_slicers[sheet_id] = [slicer for slicer, _ in slicer_parts]
        if spec.shapes:
            self._sheet_shapes[sheet_id] = spec.shapes
        if spec.charts:
            self._sheet_charts[sheet_id] = spec.charts
        if spec.images:
            self._sheet_images[sheet_id] = spec.images
            self._sheet_image_media[sheet_id] = [
                self._media.name(image) for image in spec.images
            ]
        if spec.print_settings is not None:
            self._sheet_print[sheet_index] = spec.print_settings
        for row_idx, col_idx, link in body.hyperlinks:
            if link.is_internal:
                self._internal_links.append((sheet_index, link.location))
        return rels, link_rels, drawing_rid, slicer_rids

    def _place_pivots(self, feeders: list[PivotFeeder], sheet_name: str, row_sink: Any, body: Any) -> list[Any]:
        """Number the pivots this sheet's rows fed, append those that land on this
        sheet below its rows and queue the others for their target sheet;
        returns the ones placed here."""
        own_renders = []
        for feeder in feeders:
            render = feeder.result()
            render.assign_ids(self._next_pivot_id)
            self._next_pivot_id += 1
            self._pivot_renders.append(render)
            own_renders.append(render)
        same_sheet = [
            render for render in own_renders
            if render.plan.target_sheet.casefold() == sheet_name.casefold()
        ]
        if same_sheet:
            self._append_pivot_rows(row_sink, same_sheet, body)
        for render in own_renders:
            if any(render is item for item in same_sheet):
                continue
            target_index = self._sheet_index_by_name(render.plan.target_sheet)
            if target_index is not None and target_index in self._sheet_written:
                raise ValueError(
                    f"pivot target sheet {render.plan.target_sheet!r} was written "
                    "before its source sheet"
                )
            self._pending_pivots.setdefault(render.plan.target_sheet, []).append(render)
        return same_sheet

    def _finish_tail(
        self, tail: str, rels: list[tuple[str, str, str, bool]], sheet_id: int, sheet_name: str,
        spec: SheetSpec, sheet_protection: SheetProtection | None, ignored_errors: Any, slicer_rids: list[str],
    ) -> str:
        """Add the protection, ignored errors, comment parts (to ``rels``) and the
        sparkline and slicer extensions to the sheet's closing XML."""
        if sheet_protection is not None:
            tail = '</sheetData>' + sheet_protection.to_xml() + tail[len('</sheetData>'):]
        if ignored_errors:
            tail = replace_element(
                tail, 'ignoredErrors', ignored_errors_xml(ignored_errors),
                before=('drawing', 'legacyDrawing', 'tableParts'),
            )
        if spec.comments:
            self._sheet_comments[sheet_id] = spec.comments
            rid = f"rId{len(rels) + 1}"
            rels.append((rid, _REL_WORKSHEET.rsplit('/', 1)[0] + '/comments',
                         f'../comments/comment{sheet_id}.xml', False))
            vml_rid = f"rId{len(rels) + 1}"
            rels.append((vml_rid, _REL_WORKSHEET.rsplit('/', 1)[0] + '/vmlDrawing',
                         f'../drawings/commentsDrawing{sheet_id}.vml', False))
            tail = replace_element(tail, 'legacyDrawing', f'<legacyDrawing r:id="{vml_rid}"/>', before=('tableParts',))
        # Sparklines and slicers live in the sheet's extLst, which must be the last child.
        sheet_exts = [sparklines_ext(spec.sparklines, sheet_name)]
        if slicer_rids:
            items = "".join(f'<x14:slicer r:id="{rid}"/>' for rid in slicer_rids)
            sheet_exts.append(ext(EXT_SHEET_SLICER_LIST,
                                  f"<x14:slicerList>{items}</x14:slicerList>",
                                  ("x14", X14), ("x15", X15)))
        extensions = ext_list(sheet_exts)
        if extensions:
            tail = replace_element(tail, 'extLst', extensions)
        return tail

    def write_dataframe(
        self, df: Any, include_index: bool = False, na_rep: Any = None, formats: dict | None = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        conditional_formats: list | tuple | None = None,
        tables: list[ExcelTable] | tuple[ExcelTable, ...] | None = None,
        column_widths: dict | None = None,
        autofit_rows: int | None = 1000,
        autofit_padding: float = 0.0,
        hidden_columns: list | tuple | set | None = None,
        row_heights: dict | None = None,
        hidden_rows: list | tuple | set | None = None,
        row_groups: list | tuple | None = None,
        column_groups: list | tuple | None = None,
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        charts: list | tuple | None = None,
        data_validations: list | tuple | None = None,
        images: list | tuple | None = None,
        print_settings: PrintSettings | None = None,
        comments: Mapping[str, Comment] | None = None,
        sheet_protection: SheetProtection | None = None,
        show_grid_lines: bool = True,
        tab_color: Color | str | None = None,
        zoom: int | None = None,
        show_row_col_headers: bool = True,
        show_zero_values: bool = True,
        right_to_left: bool = False,
        pivots: list | tuple | None = None,
    ) -> None:
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

        self.write_sheet(
            dataframe_to_rows(df, include_index=include_index, na_rep=na_rep, formats=formats),
            header_style=header_style,
            styles=styles,
            merges=merges,
            conditional_formats=conditional_formats,
            tables=tables,
            column_widths=column_widths,
            autofit_rows=autofit_rows,
            autofit_padding=autofit_padding,
            hidden_columns=hidden_columns,
            row_heights=row_heights,
            hidden_rows=hidden_rows,
            row_groups=row_groups,
            column_groups=column_groups,
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            charts=charts,
            data_validations=data_validations,
            images=images,
            print_settings=print_settings,
            comments=comments,
            sheet_protection=sheet_protection,
            show_grid_lines=show_grid_lines,
            tab_color=tab_color,
            zoom=zoom,
            show_row_col_headers=show_row_col_headers,
            show_zero_values=show_zero_values,
            right_to_left=right_to_left,
            pivots=pivots,
        )

    def _write_body(self, data, spool, style_ranges, header_style, heights,
                    hidden_row_set, has_styles, table_headers=(),
                    row_outline=None, meter=None, table_totals=None):
        """Encode every row into ``spool`` and return a ``_SheetBody``.

        ``meter`` measures the automatic column widths of the first rows (its
        ``limit``); the header rows of ``table_headers`` are kept, and checked
        as soon as they stream past.
        Hyperlink cells are recorded so the trailer can link them.
        """
        head: list = []  # kept empty: the meter measures the rows
        measure_until = meter.limit if meter is not None else 0
        fast_add = meter.fast_add if meter is not None else None
        slow_rows = meter.slow_rows if meter is not None else frozenset()
        hyperlinks: list = []
        header_rows: dict = {}
        num_rows = 0
        max_cols = 0
        value_cols = 0
        arrays: list = []  # (c1, r1, c2, r2) of the array formulas written
        buf = io.BytesIO()  # encoded rows waiting for the spool
        text = io.StringIO()  # one Python-written row at a time, encoded once
        row_style_ids_cache: dict[Any, Any] = {}
        ss_dict, ss_list = self._shared_strings_dict, self._shared_strings
        use_sst, fmt_xf = self._use_shared_strings, self._fmt_xf
        # Style id of the built-in link style, 0 until a link registers it. Only the
        # Python row path registers styles, so the lookup is repeated after those rows.
        link_xf = self._style_xf_map.get(HYPERLINK_STYLE, 0)
        # Cells of types the native encoder does not know reach this hook one at
        # a time (see call_hook in cells.rs): it gives the value the Python
        # writer would write. A link in a styled cell is kept in pending_links
        # until its row has been written natively.
        pending_links: list = []

        def cell_hook(cell, col, style, row):
            if isinstance(cell, Hyperlink):
                # The link's xf, once the Python writer has written a link in
                # this range style (see _write_row_cells).
                xf = fmt_xf.get(("link", style))
                text = cell.display_text
                if xf is None or type(text) is not str:
                    return NotImplemented
                pending_links.append((row, col, cell))
                return text, xf
            return native_cell_value(cell, _XLSX_PYTHON_CELLS)

        if table_headers:
            data = capture_table_headers(data, table_headers, header_rows)
            if table_totals is not None:
                data = capture_table_totals(data, table_headers, table_totals)
        last_written = -1  # index of the last <row> written
        for row_idx, row in enumerate(data):
            row_styles = (
                styles_for_row(row_idx, len(row), style_ranges, header_style)
                if has_styles else _EMPTY_ROW_STYLES
            )
            if row_idx < measure_until:
                if fast_add is not None and row_idx not in slow_rows:
                    fast_add(row, row_styles or None)
                else:
                    meter.add(row_idx, row, row_styles)
            num_rows = row_idx + 1
            if len(row) > max_cols:
                max_cols = len(row)
            # Only a value past the columns seen so far can widen value_cols.
            last_cell = len(row) - 1
            while last_cell >= value_cols and row[last_cell] is None:
                last_cell -= 1
            for col_idx in range(last_cell, value_cols - 1, -1):
                cell = row[col_idx]
                if cell is not None and _unwrap_cell(cell) is not None:
                    value_cols = col_idx + 1
                    break
            row_attrs = self._row_attrs(row_idx, heights, hidden_row_set, row_outline)
            if len(row) == 0 and not row_attrs and not row_styles:
                # Nothing to show: no <row> at all. The next row written takes its
                # number from a gap (see write_r below), so the sheet keeps its place.
                continue
            # Excel numbers a <row> without r as the row after the previous one. Dense
            # data therefore has no r attribute, and only a row after a gap carries it.
            write_r = row_idx != last_written + 1
            last_written = row_idx
            # The native encoder handles common scalar rows with optional style ids
            # and <row> attributes; it hands back rows it cannot write as None
            # for the Python path.
            style_ids = (
                _style_ids_for_row(
                    row_styles, len(row), self._register_style,
                    row_style_ids_cache,
                )
                if row_styles else None
            )
            # Link cells need the built-in link style registered first: the Python
            # path registers it on the first link, so the xf order is unchanged.
            if pending_links:
                pending_links.clear()  # from a row that went to Python
            c_row = accel_encode_xlsx_row(
                row, ss_dict, ss_list,
                self._sst_unique_count, self._sst_all_count,
                row_idx, use_sst, style_ids, fmt_xf or None, row_attrs or None,
                link_xf, write_r, cell_hook,
            )
            if c_row is not None:
                row_xml, self._sst_unique_count, self._sst_all_count, links = c_row
                buf.write(row_xml)
                if pending_links:
                    # In column order, as the Python writer lists them.
                    links = sorted([*links, *pending_links], key=lambda link: link[1])
                hyperlinks.extend(links)
            else:
                row_number = f' r="{row_idx + 1}"' if write_r else ''
                text.write(f'<row{row_number}{row_attrs}>')
                self._write_row_cells(text, row, row_idx, row_styles, hyperlinks, arrays)
                text.write('</row>')
                buf.write(text.getvalue().encode('utf-8'))
                text.seek(0)
                text.truncate()
                if not link_xf:
                    link_xf = self._style_xf_map.get(HYPERLINK_STYLE, 0)
            if buf.tell() >= _FLUSH_BYTES:
                spool.write(buf.getvalue())
                buf = io.BytesIO()
        # Empty rows at the end of the data were skipped above; they are written now so
        # the sheet keeps its shape (an empty frame is still one row). They follow the
        # last written row directly, so none of them needs r.
        trailing = num_rows - last_written - 1
        while trailing > 0:
            chunk = min(trailing, 4096)
            buf.write(b"<row></row>" * chunk)
            trailing -= chunk
            if buf.tell() >= _FLUSH_BYTES:
                spool.write(buf.getvalue())
                buf = io.BytesIO()
        if row_outline is not None:
            # Grouped rows and summary rows past the data are written empty.
            for row_idx in row_outline.indices_from(num_rows):
                row_attrs = self._row_attrs(row_idx, heights, hidden_row_set, row_outline)
                buf.write(f'<row r="{row_idx + 1}"{row_attrs}/>'.encode())
                if buf.tell() >= _FLUSH_BYTES:
                    spool.write(buf.getvalue())
                    buf = io.BytesIO()
        spool.write(buf.getvalue())
        if num_rows == 0:
            max_cols = 1
        return _SheetBody(
            num_rows, max_cols, head, hyperlinks, header_rows=header_rows, value_cols=value_cols,
            array_rows=max((r2 + 1 for _c1, _r1, _c2, r2 in arrays), default=0),
        )

    @staticmethod
    def _row_attrs(row_idx, heights, hidden_row_set, row_outline) -> str:
        """``<row>`` attributes for height, outline level, collapsed and hidden flags."""
        row_attrs = ""
        if row_idx in heights:
            row_attrs += f' ht="{heights[row_idx]}" customHeight="1"'
        outline_hidden = False
        if row_outline is not None:
            outline_level, outline_hidden = row_outline.state_at(row_idx)
            if outline_level:
                row_attrs += f' outlineLevel="{outline_level}"'
            if row_outline.collapsed_at(row_idx):
                row_attrs += ' collapsed="1"'
        if row_idx in hidden_row_set or outline_hidden:
            row_attrs += ' hidden="1"'
        return row_attrs

    def _write_row_cells(self, buf, row, row_idx, row_styles, hyperlinks, arrays=None):
        """Write one row's cells; array formulas are added to ``arrays`` when given."""
        uniform_style = None
        uniform_xf = None
        if row_styles and len(row_styles) == len(row):
            first_style = next(iter(row_styles.values()))
            if all(style is first_style for style in row_styles.values()):
                uniform_style = first_style
                uniform_xf = self._register_style(first_style)

        for col_idx, cell in enumerate(row):
            cell_style = uniform_style if uniform_style is not None else row_styles.get(col_idx)
            if cell is not None and type(cell).__module__.startswith("pandas") and is_pandas_missing(cell):
                cell = None  # NaT, NA: an empty cell
            if cell is None:
                # Every cell of the row is written, so none needs its r (as the
                # native encoder writes them).
                if cell_style is None:
                    buf.write('<c/>')
                else:
                    xf_index = uniform_xf if cell_style is uniform_style else self._register_style(cell_style)
                    buf.write(f'<c s="{xf_index}"/>')
                continue

            # (value, format) and (value, Style) cells: the overlay goes on
            # top of the range style.
            overlay = None
            link_base = None
            if _is_formatted_cell(cell):
                cell, overlay = cell
                if isinstance(cell, Hyperlink):
                    if isinstance(overlay, str):
                        raise TypeError("a hyperlink cell cannot carry a number format; style the cell instead")
                    cell_style = merge_styles(cell_style, overlay)
                    overlay = None
            elif isinstance(cell, Hyperlink) and cell_style is not None:
                # The native encoder's hook writes the next links in this range
                # style with the xf found under ("link", its xf).
                link_base = self._style_xf_map.get(cell_style)
            cell = plain_value(cell)
            if overlay is None and isinstance(cell, Hyperlink):
                hyperlinks.append((row_idx, col_idx, cell))
                cell = cell.display_text
                cell_style = HYPERLINK_STYLE if cell_style is None else merge_styles(HYPERLINK_STYLE, cell_style)

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
            elif cell_style is not None:
                xf_index = uniform_xf if cell_style is uniform_style else self._register_style(cell_style)
            else:
                xf_index = None
            if link_base is not None and xf_index is not None:
                self._fmt_xf.setdefault(("link", link_base), xf_index)

            if cell is None:
                # (None, format) or (None, Style): a formatted cell without a
                # value, written like a blank with a range style.
                buf.write('<c/>' if xf_index is None else f'<c s="{xf_index}"/>')
            elif isinstance(cell, Formula):
                self._write_formula_cell(buf, cell, xf_index)
                if arrays is not None and cell.array_ref is not None:
                    arrays.append(parse_cell_range(cell.array_ref))
            elif isinstance(cell, str):
                self._write_string_cell(buf, cell, row_idx == 0, xf_index)
            elif isinstance(cell, bool):
                s_attr = f' s="{xf_index}"' if xf_index is not None else ''
                buf.write(f'<c t="b"{s_attr}><v>{"1" if cell else "0"}</v></c>')
            elif isinstance(cell, (int, float, Decimal)):
                if isinstance(cell, (float, Decimal)) and not math.isfinite(float(cell)):
                    raise ValueError("numeric cell values must be finite")
                try:
                    number = float(cell)
                except OverflowError:
                    if not isinstance(cell, int):
                        raise
                    # An int past the float range: the digits, as text (as XlsbWriter writes it).
                    self._write_string_cell(buf, str(cell), row_idx == 0, xf_index)
                else:
                    s_attr = f' s="{xf_index}"' if xf_index is not None else ''
                    buf.write(f'<c{s_attr}><v>{number}</v></c>')
            elif isinstance(cell, datetime.datetime):
                if cell.year < 1900 or cell.year > 9999:
                    self._write_string_cell(buf, str(cell), row_idx == 0, xf_index)
                else:
                    if cell.tzinfo is not None:
                        cell = cell.replace(tzinfo=None)
                    excel_date = datetime_to_serial(cell)
                    style_ref = 2 if xf_index is None else _date_cell_xf(
                        self._fmt_xf, self._xf_styles, self._register_style,
                        xf_index, DEFAULT_DATETIME_FORMAT, True,
                    )
                    buf.write(f'<c s="{style_ref}"><v>{excel_date}</v></c>')
            elif isinstance(cell, datetime.date):
                if cell.year < 1900 or cell.year > 9999:
                    self._write_string_cell(buf, str(cell), row_idx == 0, xf_index)
                else:
                    excel_date = datetime_to_serial(cell)
                    style_ref = 1 if xf_index is None else _date_cell_xf(
                        self._fmt_xf, self._xf_styles, self._register_style,
                        xf_index, DEFAULT_DATE_FORMAT, False,
                    )
                    buf.write(f'<c s="{style_ref}"><v>{excel_date}</v></c>')
            elif isinstance(cell, CellRichText):
                self._write_rich_cell(buf, cell, row_idx == 0, xf_index)
            else:
                self._write_string_cell(buf, str(cell), row_idx == 0, xf_index)

    def _write_formula_cell(self, buf, formula: Formula, xf_index):
        s_attr = f' s="{xf_index}"' if xf_index is not None else ''
        value = formula.value
        t_attr = ''
        v_xml = ''
        if isinstance(value, str):
            check_cell_text(value)
            t_attr = ' t="str"'
            v_xml = f'<v>{xstring_escape(saxutils.escape(value))}</v>'
        elif value is not None:
            v_xml = f'<v>{_format_number(float(value) if isinstance(value, Decimal) else value)}</v>'
        f_attr = f' t="array" ref="{formula.array_ref}"' if formula.array_ref is not None else ''
        buf.write(
            f'<c{s_attr}{t_attr}><f{f_attr}>{saxutils.escape(formula.text)}</f>{v_xml}</c>'
        )

    def _append_pivot_rows(self, sink, renders, body) -> None:
        """Write pivot cells below the streamed rows of the source sheet."""
        merged = merge_cells(renders)
        if not merged:
            return
        by_row: dict[int, dict[int, object]] = {}
        for (row, column), value in merged.items():
            by_row.setdefault(row, {})[column] = value
        first_row = min(by_row)
        if body.num_rows > 0 and first_row < body.num_rows:
            raise ValueError(
                "a pivot table on the source sheet must start below the sheet's "
                f"data (first free row {body.num_rows + 1}); got row {first_row + 1}"
            )
        for row_idx in sorted(by_row):
            cells = by_row[row_idx]
            width = max(cells) + 1
            cell_row: list[Any] = [None] * width
            for column, value in cells.items():
                cell_row[column] = value
            buf = io.StringIO()
            buf.write(f'<row r="{row_idx + 1}">')
            self._write_row_cells(buf, cell_row, row_idx, {}, [])
            buf.write('</row>')
            sink.write(buf.getvalue().encode('utf-8'))
            body.num_rows = max(body.num_rows, row_idx + 1)
            body.max_cols = max(body.max_cols, width)
            body.value_cols = max(body.value_cols, width)

    def _write_pivot_sheet(self, sheet_index: int, renders) -> None:
        """A sheet that only holds pivot output (and its worksheet rels)."""
        merged = merge_cells(renders)
        by_row: dict[int, dict[int, object]] = {}
        for (row, column), value in merged.items():
            by_row.setdefault(row, {})[column] = value
        rows_xml = []
        max_row = max_col = 0
        for row_idx in sorted(by_row):
            cells = by_row[row_idx]
            width = max(cells) + 1
            cell_row: list[Any] = [None] * width
            for column, value in cells.items():
                cell_row[column] = value
            buf = io.StringIO()
            buf.write(f'<row r="{row_idx + 1}">')
            self._write_row_cells(buf, cell_row, row_idx, {}, [])
            buf.write('</row>')
            rows_xml.append(buf.getvalue())
            max_row = max(max_row, row_idx + 1)
            max_col = max(max_col, width)
        dimension = f"A1:{self._letters[max_col - 1]}{max_row}" if max_row else "A1"
        self._zf.writestr(
            f"xl/worksheets/sheet{sheet_index + 1}.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="{dimension}"/>'
            f"<sheetData>{''.join(rows_xml)}</sheetData>"
            "</worksheet>",
        )
        sheet_id = sheet_index + 1
        rels: list[Any] = []
        for render in renders:
            rid = f"rId{len(rels) + 1}"
            rels.append((
                rid, _REL_PIVOT_TABLE,
                f"../pivotTables/pivotTable{render.pivot_id}.xml", False,
            ))
        if rels:
            self._sheet_rels[sheet_id] = rels

    def _write_empty_sheet(self, sheet_index: int) -> None:
        """A sheet that was added but never written holds no rows."""
        sheet_name = self._worksheet_data[sheet_index][0]
        renders = self._pending_pivots.pop(sheet_name, None)
        if renders:
            self._write_pivot_sheet(sheet_index, renders)
            return
        self._zf.writestr(
            f"xl/worksheets/sheet{sheet_index + 1}.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData/></worksheet>',
        )

    def _write_rich_cell(self, buffer, rich: CellRichText, is_header: bool, xf_index):
        runs = rich.runs_xml()
        if xf_index is not None:
            style_ref = f' s="{xf_index}"'
        else:
            style_ref = ' s="3"' if is_header else ''
        if not self._use_shared_strings:
            buffer.write(f'<c t="inlineStr"{style_ref}><is>{runs}</is></c>')
            return
        # Rich entries share the table with plain strings; a tuple key never
        # equals a str, so the same text plain and formatted stays two entries.
        key = ("rich", runs)
        self._sst_all_count += 1
        string_index = self._shared_strings_dict.get(key)
        if string_index is None:
            string_index = self._sst_unique_count
            self._shared_strings_dict[key] = string_index
            self._shared_strings.append(key)
            self._sst_unique_count += 1
        buffer.write(f'<c t="s"{style_ref}><v>{string_index}</v></c>')

    def _write_string_cell(self, buffer: io.StringIO, cell_value: str, is_header: bool = False, xf_index: int | None = None):
        check_cell_text(cell_value)
        if self._use_shared_strings:
            self._sst_all_count += 1
            if cell_value not in self._shared_strings_dict:
                string_index = self._sst_unique_count
                self._shared_strings_dict[cell_value] = string_index
                self._shared_strings.append(cell_value)
                self._sst_unique_count += 1
            else:
                string_index = self._shared_strings_dict[cell_value]

            if xf_index is not None:
                style_ref = f' s="{xf_index}"'
            else:
                style_ref = ' s="3"' if is_header else ''
            buffer.write(f'<c t="s"{style_ref}><v>{string_index}</v></c>')
        else:
            escaped_value = xstring_escape(saxutils.escape(cell_value))
            if xf_index is not None:
                style_ref = f' s="{xf_index}"'
            else:
                style_ref = ' s="3"' if is_header else ''
            space = xml_space_attribute(cell_value)
            buffer.write(f'<c t="inlineStr"{style_ref}><is><t{space}>{escaped_value}</t></is></c>')
