"""XlsbWriter sheet methods: write_sheet, the sheet loop and the dataframe entry point."""
from __future__ import annotations

from typing import Any, TYPE_CHECKING
from collections.abc import Iterable, Mapping
from ._sheet_options import XLSB_OPTIONS, SheetOptions, resolve_sheet_options
from . import chart_parts
from .styles import (
    Chart,
    Style,
    capture_table_headers,
    capture_table_totals,
    parse_cell_range,
    table_column_names,
)
from .images import Image
from .pivots import (
    PivotFeeder, append_rows, normalize_pivots, overlay_rows, tee_rows,
)
from .printing import PrintSettings
from .data_validation import DataValidation
from ._xlsb_writer_consts import _capture_cells
from ._xlsb_writer_host import _XlsbWriterHost
from ._writer_common import SheetSpec
from ._metadata import SheetProtection
from ._dataframe_rows import dataframe_to_rows

if TYPE_CHECKING:
    from .comments import Comment
    from .styles import Color


class _XlsbSheetMixin(_XlsbWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

    def write_sheet(self, data: Iterable[list[Any]], *, options: SheetOptions | None = None,
                    **kwargs: Any) -> None:
        """Write rows to the current sheet.

        Sheet settings come from ``options`` (a ``SheetOptions``) or from keyword
        arguments with the same names; a keyword argument overrides the matching
        field of ``options``.
        """
        opts = resolve_sheet_options(options, kwargs, XLSB_OPTIONS, "XlsbWriter")
        self._write_sheet(data, **opts.as_kwargs(XLSB_OPTIONS))

    def _write_sheet(
        self,
        data: Iterable[list[Any]],
        *,
        header_style: Style | None = None,
        styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
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
        images: list[Image] | tuple[Image, ...] | None = None,
        print_settings: PrintSettings | None = None,
        data_validations: list[DataValidation] | tuple[DataValidation, ...] | None = None,
        tables: list | tuple | None = None,
        conditional_formats: list | tuple | None = None,
        comments: Mapping[str, Comment] | None = None,
        sheet_protection: SheetProtection | None = None,
        show_grid_lines: bool = True,
        header_rows: int = 1,
        tab_color: Color | str | None = None,
        zoom: int | None = None,
        ignored_errors: list | tuple | None = None,
        sparklines: Any = None,
        slicers: Any = None,
        shapes: Any = None,
        show_row_col_headers: bool = True,
        show_zero_values: bool = True,
        right_to_left: bool = False,
        pivots: list | tuple | None = None,
        sort_state: Any | None = None,
    ) -> None:
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
            check_style=self._reject_rich_style, sort_state=sort_state,
        )
        current_sheet = self._worksheet_data[-1][0] if self._worksheet_data else "Sheet1"
        pivot_names = set(self._pivot_names)
        pivot_plans = normalize_pivots(pivots, spec.tables, pivot_names, current_sheet)
        for plan in pivot_plans:
            target_index = self._sheet_index_by_name(plan.target_sheet)
            if (
                target_index is not None
                and target_index != max(self._sheet_count - 1, 0)
                and target_index in self._written_sheets
            ):
                raise ValueError(
                    f"pivot target sheet {plan.target_sheet!r} was written before its "
                    "source sheet; write the pivot's sheet after the source sheet"
                )
        # Every option is valid; only now record what this sheet claims, so a
        # rejected call leaves the writer as it was.
        self._reject_second_write()
        self._claim_sheet_options(spec)
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        sheet_name,_, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)
        sheet_id = self._sheet_count
        if spec.comments:
            self._sheet_comments[sheet_id] = spec.comments
        if spec.charts:
            self._sheet_charts[sheet_id] = spec.charts
            # Chart caches need only the cells the charts read. Capture those
            # while the rows stream past instead of holding the whole sheet.
            store = chart_parts.new_cell_store(chart_parts.chart_cell_spans(spec.charts))
            data = _capture_cells(data, store)
            self._sheet_chart_cells[sheet_id] = store
        if spec.images:
            self._sheet_images[sheet_id] = spec.images
            self._sheet_image_media[sheet_id] = [self._media.name(image) for image in spec.images]
        if spec.print_settings is not None:
            self._sheet_print[sheet_id - 1] = spec.print_settings
        if spec.shapes:
            self._sheet_shapes[sheet_id] = spec.shapes
        if spec.slicers:
            self._sheet_slicers[sheet_id] = spec.slicers
        has_drawing = bool(spec.charts or spec.images or spec.shapes or spec.slicers)
        if has_drawing:
            self._sheet_drawings.add(sheet_id)
        table_ids: list[int] = []
        table_rel_ids: list[str] = []
        table_header_cells: dict = {}
        table_totals_cells: dict = {}
        if spec.tables:
            # rId1 is binaryIndex; tables take rId2..; drawing (if any) last.
            for table in spec.tables:
                table_id = self._next_table_id
                self._next_table_id += 1
                table_ids.append(table_id)
                rel_id = f"rId{1 + len(table_ids)}"
                table_rel_ids.append(rel_id)
            self._sheet_tables[sheet_id] = list(zip(table_ids, table_rel_ids))
            # Column names must match the header cells, which stream past with
            # the rows; a mismatch fails as soon as the header row is read.
            data = capture_table_headers(data, spec.tables, table_header_cells)
            data = capture_table_totals(data, spec.tables, table_totals_cells)
        # Pivots aggregate the sheet's own rows as they stream; the cells of pivots
        # from earlier sheets are merged in after, and this sheet's own pivots are
        # appended below the data once the stream has ended.
        pending_pivots = self._pending_pivots.pop(sheet_name, None) or []
        pivot_renders: list = []
        if pivot_plans:
            feeders = [PivotFeeder(plan, track_all_fields=True) for plan in pivot_plans]
            data = tee_rows(data, feeders)
        if pending_pivots:
            data = overlay_rows(data, pending_pivots)
        if pivot_plans:
            data = append_rows(data, feeders, sheet_name, pivot_renders)
        if spec.tables and auto_filter is None:
            # Tables carry their own AutoFilter; skip sheet-level filter.
            auto_filter = False
        drawing_rel = f"rId{2 + len(table_ids)}" if has_drawing else None
        # External hyperlinks take the relationship ids after the others
        # (binary index, tables, drawing, comments).
        link_rids: dict[str, str] = {}
        link_rid_start = 2 + len(table_ids) + int(has_drawing) + (2 if spec.comments else 0)
        # The sheet's slicer part comes next; its id is known before the rows stream.
        slicer_rid = f"rId{link_rid_start}" if spec.slicers else None
        link_rid_start += int(slicer_rid is not None)
        # Formulas written while the rows stream may refer to this sheet's tables.
        self._sheet_table_scope = (sheet_name, spec.tables, table_ids, table_header_cells) if spec.tables else None
        try:
            self._write_worksheet_bin(
                f"xl/worksheets/sheet{sheet_id}.bin", data, self._sheet_count - 1,
                style_ranges=spec.style_ranges, header_style=spec.header_style,
                merge_refs=spec.merge_refs, widths=spec.widths, autofit=spec.autofit, hidden_cols=spec.hidden_cols,
                heights=spec.heights, hidden_row_set=spec.hidden_rows,
                freeze_panes=freeze_panes, auto_filter=auto_filter,
                auto_filter_ref=auto_filter_ref, sort_state=spec.sort_state,
                drawing_rel=drawing_rel, table_rel_ids=table_rel_ids,
                cf_rules=spec.cf_rules, row_outline=spec.row_outline, col_outline=spec.col_outline,
                sheet_protection=sheet_protection,
                show_grid_lines=show_grid_lines, view=spec.view,
                header_rows=spec.header_rows, tab_key=spec.tab_key, zoom=spec.zoom,
                filter_cells=[
                    (r1, c1, c2) for table in spec.tables if table.header_row
                    for c1, r1, c2, _r2 in (parse_cell_range(table.ref),)
                ],
                table_bounds=frozenset(parse_cell_range(table.ref) for table in spec.tables),
                comment_vml_rid=(f'rId{3 + len(table_ids) + int(has_drawing)}' if spec.comments else None),
                link_rids=link_rids, link_rid_start=link_rid_start,
                print_spec=spec.print_settings, validations=spec.validations,
                sparklines=spec.sparklines, slicer_rel=slicer_rid,
                ignored_errors=ignored_errors,
            )
            self._sheet_link_rels[sheet_id] = link_rids
            self._register_sheet_parts(
                sheet_id, sheet_name, spec, pivot_renders, pending_pivots, pivot_names,
                table_ids, table_header_cells, table_totals_cells, slicer_rid,
            )
            self._zf.writestr(f"xl/worksheets/_rels/sheet{sheet_id}.bin.rels", self._create_worksheet_rels(sheet_id))
            self._zf.writestr(f"xl/worksheets/binaryIndex{sheet_id}.bin", self._binaryIndexBin)
        except BaseException:
            # The sheet part is incomplete, so the workbook cannot be saved.
            self._aborted = True
            raise
        finally:
            self._sheet_table_scope = None

    def _register_sheet_parts(
        self, sheet_id: int, sheet_name: str, spec: SheetSpec, pivot_renders: list[Any],
        pending_pivots: list[Any], pivot_names: set[str], table_ids: list[int],
        table_header_cells: dict, table_totals_cells: dict, slicer_rid: str | None,
    ) -> None:
        """After the sheet part is written: number its pivots and queue those for
        later sheets, encode its table parts from the header and totals cells seen,
        and write its slicer parts."""
        own_pivots = []
        for render in pivot_renders:
            render.assign_ids(len(self._pivot_renders) + 1)
            self._pivot_renders.append(render)
            if render.plan.target_sheet.casefold() == sheet_name.casefold():
                own_pivots.append(render)
            else:
                self._pending_pivots.setdefault(render.plan.target_sheet, []).append(render)
        if pending_pivots or own_pivots:
            self._sheet_pivots[sheet_id] = pending_pivots + own_pivots
        self._pivot_names = pivot_names
        self._written_sheets.add(self._sheet_count - 1)
        table_columns = {}
        for table, table_id in zip(spec.tables, table_ids):
            first_col, first_row, last_col, last_row = parse_cell_range(table.ref)
            names = table_column_names(table, table_header_cells.get(first_row))
            totals = table_totals_cells.get(last_row)
            if totals is not None:
                totals = totals[first_col:last_col + 1]
            self._table_registry[table_id] = (table.display_name, names, sheet_name)
            self._tables_bin[table_id] = self._table_bin(
                table, table_id, names, totals, self._formula_context(), sheet_name,
            )
            table_columns[table.display_name.casefold()] = (table_id, names)
        if slicer_rid is not None:
            self._write_slicer_parts(sheet_id, slicer_rid, spec.slicers, table_columns)

    def write_dataframe(
        self, df: Any, include_index: bool = False, na_rep: Any = None, formats: dict | None = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
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
        charts: list | tuple | None = None,
        images: list | tuple | None = None,
        print_settings: PrintSettings | None = None,
        data_validations: list | tuple | None = None,
        tables: list | tuple | None = None,
        conditional_formats: list | tuple | None = None,
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
            auto_filter_ref=auto_filter_ref,
            charts=charts,
            images=images,
            print_settings=print_settings,
            data_validations=data_validations,
            tables=tables,
            conditional_formats=conditional_formats,
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

    def _write_empty_sheet(self, sheet_index: int) -> None:
        """A sheet that was added but never written holds no rows."""
        self._write_worksheet_bin(f"xl/worksheets/sheet{sheet_index + 1}.bin", [], sheet_index)
