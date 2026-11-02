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
    check_header_rows,
    check_sheet_view,
    check_zoom,
    header_style_ranges,
    tab_color_key,
    explicit_cf_priorities,
    normalize_hidden_columns,
    normalize_hidden_rows,
    normalize_style_ranges,
    parse_cell_range,
    table_column_names,
)
from .images import Image, normalize_images
from .shapes import normalize_shapes
from .pivots import (
    PivotFeeder, append_rows, normalize_pivots, overlay_rows, tee_rows,
)
from .slicers import normalize_slicers
from .sparklines import normalize_sparklines
from .printing import PrintSettings, normalize_print_settings
from .data_validation import DataValidation, normalize_validations
from .outline import build_outline
from ._xlsb_writer_consts import _capture_cells
from ._xlsb_writer_host import _XlsbWriterHost
from .comments import normalize_comments
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
        field of ``options``. ``ignored_errors`` is an XLSX-only option.
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
        sparklines=None,
        slicers=None,
        shapes=None,
        show_row_col_headers: bool = True,
        show_zero_values: bool = True,
        right_to_left: bool = False,
        pivots: list | tuple | None = None,
    ):
        self._require_writable()
        view = check_sheet_view(show_row_col_headers, show_zero_values, right_to_left)
        if not isinstance(show_grid_lines, bool):
            raise TypeError("show_grid_lines must be a bool")
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        header_rows = check_header_rows(header_rows)
        tab_key = tab_color_key(tab_color)
        zoom = check_zoom(zoom)
        style_ranges = normalize_style_ranges(styles)
        for candidate in ([header_style] if header_style is not None else []) + [item[-1] for item in style_ranges]:
            self._reject_rich_style(candidate)
        style_ranges, header_style = header_style_ranges(style_ranges, header_style, header_rows)
        merge_refs = self._normalize_merges(merges)
        comment_specs = {}
        if comments is not None:
            comment_specs = normalize_comments(comments, merges=merge_refs)
        if sheet_protection is not None:
            if not isinstance(sheet_protection, SheetProtection):
                raise TypeError("sheet_protection must be SheetProtection")
        widths = self._normalize_widths(column_widths, "column_widths")
        hidden_cols = normalize_hidden_columns(hidden_columns)
        heights = self._normalize_row_heights(row_heights)
        hidden_row_set = normalize_hidden_rows(hidden_rows)
        row_outline = build_outline(row_groups, "row")
        col_outline = build_outline(column_groups, "column")
        if col_outline is not None:
            hidden_cols = hidden_cols | col_outline.hidden_indices()
        if freeze_panes is not None and freeze_panes != "":
            self._split_freeze_cell(freeze_panes)
        chart_specs = self._normalize_charts(charts)
        image_specs = normalize_images(images)
        shape_specs = normalize_shapes(shapes)
        sparkline_groups = normalize_sparklines(sparklines)
        print_spec = normalize_print_settings(print_settings)
        validations = normalize_validations(data_validations)
        table_specs = self._normalize_tables(tables)
        slicer_specs = normalize_slicers(
            slicers, table_specs, self._slicer_names, self._slicer_caches, self._slicer_targets,
        )
        current_sheet = self._worksheet_data[-1][0] if self._worksheet_data else "Sheet1"
        pivot_names = set(self._pivot_names)
        pivot_plans = normalize_pivots(pivots, table_specs, pivot_names, current_sheet)
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
        cf_specs = self._normalize_cf_rules(conditional_formats)
        # Pre-reserve explicit priorities (like XLSX) so auto-assigned ones skip them.
        cf_priorities = explicit_cf_priorities(
            cf_specs, self._seen_cf_priorities, self._reserved_cf_priorities,
        )

        # Every option is valid; only now record what this sheet claims, so a
        # rejected call leaves the writer as it was.
        self._reject_second_write()
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        sheet_name,_, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)
        sheet_id = self._sheet_count
        if comment_specs:
            self._sheet_comments[sheet_id] = comment_specs
        if chart_specs:
            self._sheet_charts[sheet_id] = chart_specs
            # Chart caches need only the cells the charts read. Capture those
            # while the rows stream past instead of holding the whole sheet.
            store = chart_parts.new_cell_store(chart_parts.chart_cell_spans(chart_specs))
            data = _capture_cells(data, store)
            self._sheet_chart_cells[sheet_id] = store
        if image_specs:
            self._sheet_images[sheet_id] = image_specs
            self._sheet_image_media[sheet_id] = [self._media.name(image) for image in image_specs]
        if print_spec is not None:
            self._sheet_print[sheet_id - 1] = print_spec
        if shape_specs:
            self._sheet_shapes[sheet_id] = shape_specs
        if slicer_specs:
            self._sheet_slicers[sheet_id] = slicer_specs
            self._slicer_names.update(slicer.display_name.casefold() for slicer in slicer_specs)
            self._slicer_caches.update(slicer.cache_name.casefold() for slicer in slicer_specs)
            self._slicer_targets.update(
                (slicer.table.casefold(), slicer.column.casefold()) for slicer in slicer_specs
            )
        has_drawing = bool(chart_specs or image_specs or shape_specs or slicer_specs)
        if has_drawing:
            self._sheet_drawings.add(sheet_id)
        table_ids: list[int] = []
        table_rel_ids: list[str] = []
        table_header_cells: dict = {}
        if table_specs:
            self._table_names.update(table.display_name for table in table_specs)
            # rId1 is binaryIndex; tables take rId2..; drawing (if any) last.
            for table in table_specs:
                table_id = self._next_table_id
                self._next_table_id += 1
                table_ids.append(table_id)
                rel_id = f"rId{1 + len(table_ids)}"
                table_rel_ids.append(rel_id)
            self._sheet_tables[sheet_id] = list(zip(table_ids, table_rel_ids))
            # Column names must match the header cells, which stream past with
            # the rows; a mismatch fails as soon as the header row is read.
            data = capture_table_headers(data, table_specs, table_header_cells)
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
        if table_specs and auto_filter is None:
            # Tables carry their own AutoFilter; skip sheet-level filter.
            auto_filter = False
        for priority in cf_priorities:
            self._reserved_cf_priorities.add(priority)
            self._next_cf_priority = max(self._next_cf_priority, priority + 1)
        drawing_rel = f"rId{2 + len(table_ids)}" if has_drawing else None
        # External hyperlinks take the relationship ids after the others
        # (binary index, tables, drawing, comments).
        link_rids: dict[str, str] = {}
        link_rid_start = 2 + len(table_ids) + int(has_drawing) + (2 if comment_specs else 0)
        # The sheet's slicer part comes next; its id is known before the rows stream.
        slicer_rid = f"rId{link_rid_start}" if slicer_specs else None
        link_rid_start += int(slicer_rid is not None)
        try:
            with self._zf.open_threaded(f"xl/worksheets/sheet{sheet_id}.bin") as sheet_file:
                self._write_worksheet_bin(
                    sheet_file, data, self._sheet_count - 1,
                    style_ranges=style_ranges, header_style=header_style,
                    merge_refs=merge_refs, widths=widths, hidden_cols=hidden_cols,
                    heights=heights, hidden_row_set=hidden_row_set,
                    freeze_panes=freeze_panes, auto_filter=auto_filter,
                    auto_filter_ref=auto_filter_ref,
                    drawing_rel=drawing_rel, table_rel_ids=table_rel_ids,
                    cf_rules=cf_specs, row_outline=row_outline, col_outline=col_outline,
                    sheet_protection=sheet_protection,
                    show_grid_lines=show_grid_lines, view=view,
                    header_rows=header_rows, tab_key=tab_key, zoom=zoom,
                    filter_cells=[
                        (r1, c1, c2) for table in table_specs if table.header_row
                        for c1, r1, c2, _r2 in (parse_cell_range(table.ref),)
                    ],
                    comment_vml_rid=(f'rId{3 + len(table_ids) + int(has_drawing)}' if comment_specs else None),
                    link_rids=link_rids, link_rid_start=link_rid_start,
                    print_spec=print_spec, validations=validations,
                    sparklines=sparkline_groups, slicer_rel=slicer_rid,
                )
            self._sheet_link_rels[sheet_id] = link_rids
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
            for table, table_id in zip(table_specs, table_ids):
                names = table_column_names(table, table_header_cells.get(parse_cell_range(table.ref)[1]))
                self._tables_bin[table_id] = self._table_bin(table, table_id, names)
                table_columns[table.display_name.casefold()] = (table_id, names)
            if slicer_rid is not None:
                self._write_slicer_parts(sheet_id, slicer_rid, slicer_specs, table_columns)
            self._zf.writestr(f"xl/worksheets/_rels/sheet{sheet_id}.bin.rels", self._create_worksheet_rels(sheet_id))
            self._zf.writestr(f"xl/worksheets/binaryIndex{sheet_id}.bin", self._binaryIndexBin)
        except BaseException:
            # The sheet part is incomplete, so the workbook cannot be saved.
            self._aborted = True
            raise

    def write_dataframe(
        self, df: Any, include_index: bool = False, na_rep: Any = None, formats: dict | None = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        column_widths: dict | None = None,
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
        with self._zf.open_threaded(f"xl/worksheets/sheet{sheet_index + 1}.bin") as sheet_file:
            self._write_worksheet_bin(sheet_file, [], sheet_index)
