"""The XlsxWriter state and cross-mixin methods, typed once for all its mixins.

Each XlsxWriter mixin derives from ``_XlsxWriterHost``, so mypy checks
the attributes a mixin reads against the types the writer sets, and every
method a mixin calls on another one against that method's signature (a
mismatch between a stub here and the implementation is an override error).
Everything is declared under ``TYPE_CHECKING``: at runtime the class is empty
and the writer's ``__init__`` sets the attributes.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ._metadata import DefinedName, WorkbookProtection
    from ._output import PackageZip
    from ._writer_common import FilterData, SheetSpec
    from .comments import Comment
    from .images import Image, MediaParts
    from .printing import PrintSettings
    from .shapes import Shape
    from .slicers import Slicer
    from .styles import Style


class _XlsxWriterHost:
    if TYPE_CHECKING:
        # Workbook and sheet bookkeeping
        _zf: PackageZip
        _aborted: bool
        _spool_max_bytes: int
        _worksheet_data: list[tuple[str, Iterable[list[Any]] | None, bool]]
        _sheet_count: int
        _sheet_written: set[int]
        _letters: list[str]
        _filtered_data_list: list[FilterData]
        _defined_names: dict[tuple[str, str | None], DefinedName]
        _workbook_protection: WorkbookProtection | None
        _internal_links: list[tuple[int, str]]

        # Shared strings
        _use_shared_strings: bool
        _shared_strings: list[str | tuple[str, str]]
        _shared_strings_dict: dict[str | tuple[str, str], int]
        _sst_unique_count: int
        _sst_all_count: int

        # Styles, number formats and differential (conditional format) styles
        _format_registry: dict[str, int]
        _format_xf_map: dict[str, int]
        _style_xf_map: dict[Style, int]
        _xf_styles: dict[int, Style]
        _style_components: dict[int, tuple[Any, ...]]
        _style_format_xf: dict[tuple[Any, ...], tuple[Any, ...]]
        # xf of each (value, format) cell the native row encoder may reuse
        _fmt_xf: dict[Any, Any]
        _fmt_xf_styles: list[Any]
        _font_registry: dict[tuple[Any, ...], int]
        _fill_registry: dict[tuple[Any, ...], int]
        _border_registry: dict[tuple[Any, ...], int]
        _dxf_font_registry: dict[tuple[Any, ...], int]
        _dxf_fill_registry: dict[tuple[Any, ...], int]
        _dxf_border_registry: dict[tuple[Any, ...], int]
        _dxf_index: dict[tuple[Any, ...], int]
        _dxf_list: list[tuple[Any, ...]]
        _next_numfmt_id: int
        _next_xf_index: int
        _next_cf_priority: int
        _seen_cf_priorities: set[int]
        _reserved_cf_priorities: set[int]

        # Objects on sheets: tables, charts, images, shapes, slicers, pivots, notes
        _next_table_id: int
        _table_names: set[str]
        _table_info: dict[str, tuple[int, list[str]]]
        _tables_xml: dict[int, str]
        _next_chart_id: int
        _sheet_charts: dict[int, list[Any]]
        _sheet_chart_ids: dict[int, list[int]]
        _sheet_drawings: set[int]
        _media: MediaParts
        _sheet_images: dict[int, list[Image]]
        _sheet_image_media: dict[int, list[str]]
        _sheet_shapes: dict[int, list[Shape]]
        _sheet_slicers: dict[int, list[Slicer]]
        _slicer_cache_names: list[str]
        _slicer_part_files: list[tuple[str, str, str]]
        _next_slicer_part: int
        _slicer_caches: set[str]
        _slicer_names: set[str]
        _slicer_targets: set[tuple[str, str]]
        _next_pivot_id: int
        _pivot_names: set[str]
        _pending_pivots: dict[str, list[Any]]
        _pivot_renders: list[Any]
        _sheet_comments: dict[int, dict[str, Comment]]
        _sheet_print: dict[int, PrintSettings]
        _sheet_rels: dict[int, list[tuple[str, str, str, bool]]]

        # Methods that one mixin (or a shared base) defines and another calls.
        # Writer and front end (xlsx_writer.py, _writer_common.WriterFrontEnd, _output.StagedWriter)
        def add_sheet(self, sheet_name: str, hidden: bool = False) -> None: ...
        def _require_writable(self) -> Any: ...
        def _reject_second_write(self) -> None: ...
        def _sheet_index_by_name(self, name: str) -> int | None: ...
        def _split_freeze_cell(self, ref: str) -> tuple[int, int]: ...
        def _claim_cf_priority(self, rule: Any) -> int: ...
        def _check_sheet_options(
            self, *, header_style: Any, show_grid_lines: Any, header_rows: Any, tab_color: Any, zoom: Any,
            show_row_col_headers: Any, show_zero_values: Any, right_to_left: Any, styles: Any, merges: Any,
            comments: Any, sheet_protection: Any, auto_filter_ref: Any, conditional_formats: Any, tables: Any,
            sparklines: Any, slicers: Any, charts: Any, data_validations: Any, images: Any, shapes: Any,
            print_settings: Any, column_widths: Any, hidden_columns: Any, row_heights: Any, hidden_rows: Any,
            row_groups: Any, column_groups: Any, freeze_panes: Any,
            check_style: Any | None = None,
        ) -> SheetSpec: ...
        def _claim_sheet_options(self, spec: SheetSpec) -> None: ...
        def _normalize_cf_rules(self, conditional_formats: Any) -> Any: ...
        def _normalize_tables(self, tables: Any) -> Any: ...
        @staticmethod
        def _normalize_charts(charts: Any) -> list[Any]: ...
        @staticmethod
        def _normalize_merges(merges: Any) -> Any: ...
        @staticmethod
        def _normalize_row_heights(mapping: Any) -> Any: ...
        @staticmethod
        def _normalize_widths(mapping: Any, kind: Any) -> Any: ...
        def _number_format_id(self, fmt_string: str | None) -> int | None: ...
        def _register_format(self, fmt_string: str) -> int: ...

        # Styles (_xlsx_writer_styles)
        def _register_style(self, style: Style) -> int: ...
        def _cf_rule_xml(self, rule: Any, table_ranges: set[tuple[Any, ...]]) -> str: ...
        def _create_styles_xml(self) -> str: ...
        @staticmethod
        def _color_attrs(color_key: Any) -> str: ...

        # Sheet XML (_xlsx_writer_sheet_xml) and sheets (_xlsx_writer_sheet)
        def _workbook_rel_base(self) -> int: ...
        def _link_relationships(self, hyperlinks: Any, rels: Any) -> Any: ...
        def _register_tables(self, table_specs: Any, sheet_id: Any, body: Any) -> Any: ...
        def _resolve_slicers(self, slicer_specs: Any) -> Any: ...
        def _sheet_head(self, body: Any, sheet_index: Any, split_freeze: Any, freeze_panes: Any, widths: Any,
                        hidden_cols: Any, print_spec: Any, row_outline: Any = None, col_outline: Any = None,
                        show_grid_lines: Any = True, width_styles: Any = ((), None), filter_cells: Any = (),
                        tab_key: Any = None, zoom: Any = None, view: Any = (True, True, False)) -> Any: ...
        def _sheet_tail(self, body: Any, sheet_index: Any, merge_refs: Any, cf_rules: Any, table_specs: Any,
                        table_bounds: Any, validations: Any, link_rels: Any, auto_filter: Any, drawing_rid: Any,
                        print_spec: Any, *, auto_filter_ref: Any = None, header_rows: Any = 1) -> Any: ...
        def _write_empty_sheet(self, sheet_index: int) -> None: ...

        # Document properties (doc_properties.DocumentPropertiesMixin)
        def _doc_props_content_types(self) -> str: ...
        def _doc_props_root_rels(self) -> str: ...
        def _write_doc_props(self) -> None: ...
