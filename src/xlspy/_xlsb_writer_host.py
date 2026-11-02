"""The XlsbWriter state and cross-mixin methods, typed once for all its mixins.

Each XlsbWriter mixin derives from :class:`_XlsbWriterHost`; see
``_xlsx_writer_host`` for how mypy uses it. Everything is declared under
``TYPE_CHECKING``: at runtime the class is empty and the writer sets the
attributes.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ._metadata import WorkbookProtection
    from ._writer_common import FilterData
    from ._xlsb_formulas import FormulaContext
    from .images import Image, MediaParts
    from .printing import PrintSettings
    from .rich_text import CellRichText
    from .shapes import Shape
    from .slicers import Slicer
    from .styles import Style
    from io import BufferedWriter, BytesIO


class _XlsbWriterHost:
    if TYPE_CHECKING:
        _aborted: bool
        _autoFilterEndBytes: bytes
        _autoFilterStartBytes: bytes
        _binaryIndexBin: bytes
        _border_registry: dict[tuple[Any, ...], int]
        _dxf_index: dict[bytes, int]
        _fill_registry: dict[tuple[Any, ...], int]
        _filtered_data_list: list[FilterData]
        _fmt_xf: dict[Any, Any]
        _fmt_xf_styles: list[Any]
        _font_registry: dict[tuple[Any, ...], int]
        _format_registry: dict[str, int]
        _format_xf_map: dict[str, int]
        _letters: list[str]
        _magicFilterExcel2016Fix0: bytes
        _magicFilterExcel2016Fix1: bytearray
        _magicFilterExcel2016Fix2: bytes
        _media: MediaParts
        _next_cf_priority: int
        _next_chart_id: int
        _next_numfmt_id: int
        _next_table_id: int
        _next_xf_index: int
        _pending_pivots: dict[str, list[Any]]
        _pivot_names: set[str]
        _pivot_renders: list[Any]
        _reserved_cf_priorities: set[int]
        _seen_cf_priorities: set[int]
        # A rich-text entry is ("rich", text, runs)
        _shared_strings: list[str | tuple[str, str, tuple[Any, ...]]]
        _shared_strings_dict: dict[str | tuple[str, str, tuple[Any, ...]], int]
        _sheet1Bytes: bytes
        _sheet_chart_cells: dict[int, dict[Any, Any]]
        _sheet_chart_ids: dict[int, list[int]]
        _sheet_charts: dict[int, list[Any]]
        _sheet_comments: dict[int, Any]
        _sheet_count: int
        _sheet_drawings: set[int]
        _sheet_image_media: dict[int, list[str]]
        _sheet_images: dict[int, list[Image]]
        _sheet_link_rels: dict[int, Any]
        _sheet_number: int
        _sheet_pivots: dict[int, list[Any]]
        _sheet_print: dict[int, PrintSettings]
        _sheet_shapes: dict[int, list[Shape]]
        _sheet_slicer_parts: dict[int, tuple[int, str]]
        _sheet_slicers: dict[int, list[Slicer]]
        _sheet_tables: dict[int, list[tuple[Any, ...]]]
        _slicer_cache_names: list[str]
        _slicer_caches: set[str]
        _slicer_names: set[str]
        _slicer_targets: set[tuple[str, str]]
        _sst_all_count: int
        _sst_has_rich: bool
        _sst_unique_count: int
        _stickHeaderA1bytes: bytes
        _style_components: dict[int, tuple[Any, ...]]
        _style_format_xf: dict[tuple[Any, ...], tuple[Any, ...]]
        _style_xf_map: dict[Style, int]
        _stylesBin: bytes
        _table_names: set[str]
        _tables_bin: dict[int, bytes]
        _use_shared_strings: bool
        _workbookBinEnd: bytes
        _workbookBinMiddle: bytes
        _workbookBinStart: bytes
        _workbook_protection: WorkbookProtection | None
        _worksheet_data: list[tuple[str, Iterable[list[Any]], bool]]
        _written_sheets: set[int]
        _xf_styles: dict[int, Style]
        _xlsb_formula_context: FormulaContext | None
        # A threaded ZipFile (_zip_threads) adds methods zipfile.ZipFile lacks.
        _zf: Any

        # Record builders the writer class binds as static methods
        @staticmethod
        def _build_numfmt_record(ifmt: int, fmt_string: str) -> bytes: ...
        @staticmethod
        def _color_payload(color_key: Any) -> bytes: ...

        def _build_styles_bin(self) -> bytes: ...
        def _cf_bin_grouped(self, rules: Any) -> bytes: ...
        def _check_unknown_cf_rule(self, rule: Any) -> None: ...
        def _claim_cf_priority(self, rule: Any) -> int: ...
        def _create_worksheet_rels(self, sheet_id: Any) -> str: ...
        def _doc_props_content_types(self) -> str: ...
        def _doc_props_root_rels(self) -> str: ...
        def _formula_context(self) -> Any: ...
        def _list_parts_bin(self, rel_ids: list[str]) -> bytes: ...
        def _metadata_workbook(self, data: bytes) -> bytes: ...
        def _normalize_cf_rules(self, conditional_formats: Any) -> Any: ...
        @staticmethod
        def _normalize_charts(charts: Any) -> list: ...
        @staticmethod
        def _normalize_merges(merges: Any) -> Any: ...
        @staticmethod
        def _normalize_row_heights(mapping: Any) -> Any: ...
        def _normalize_tables(self, tables: Any) -> Any: ...
        @staticmethod
        def _normalize_widths(mapping: Any, kind: Any) -> Any: ...
        def _pane_record(self, freeze_panes: str | None) -> bytes: ...
        @staticmethod
        def _register_component(registry: dict, key: Any, first_id: int) -> int: ...
        def _register_filter_names(self) -> Any: ...
        def _register_format(self, fmt_string: str) -> int: ...
        def _register_numfmt(self, fmt_string: str) -> int: ...
        def _register_print_names(self) -> Any: ...
        def _register_slicer_names(self) -> Any: ...
        def _register_style(self, style: Style) -> int: ...
        @staticmethod
        def _reject_rich_style(style: Style) -> None: ...
        def _reject_second_write(self) -> None: ...
        def _require_writable(self) -> Any: ...
        @staticmethod
        def _rich_str_payload(text: str, runs: tuple) -> bytes: ...
        @staticmethod
        def _row_header_record(row_idx: int, last_col: int, height_twips: int | None=None, hidden: bool=False, outline_level: int=0, collapsed: bool=False) -> bytes: ...
        @staticmethod
        def _row_height_twips(points: float) -> int: ...
        def _selection_record(self, freeze_panes: str) -> bytes: ...
        def _sheet_index_by_name(self, name: str) -> int | None: ...
        def _split_freeze_cell(self, ref: str) -> tuple[int, int]: ...
        def _table_bin(self, table: Any, table_id: int, names: list[str]) -> bytes: ...
        def _write_cols_width(self, buffer: BytesIO, start_col: int, end_col: int, col_widths_array: list[float], widths: dict | None=None, hidden_cols: set | frozenset=frozenset(), col_outline: Any=None) -> Any: ...
        def _write_doc_props(self) -> None: ...
        def _write_empty_sheet(self, sheet_index: int) -> None: ...
        def _write_filter_defined_names(self, sw: bytearray) -> None: ...
        def _write_rich_str(self, buffer: Any, rich: CellRichText, shared_strings_dict: Any, shared_strings_append: Any, row_idx: Any, col_idx: Any, style_override: Any=None) -> Any: ...
        def _write_slicer_parts(self, sheet_id: Any, rel_id: Any, slicers: Any, table_columns: Any) -> None: ...
        def _write_str(self, buffer: Any, cell: Any, shared_strings_dict: Any, shared_strings_append: Any, row_idx: Any, col_idx: Any, style_override: Any=None) -> Any: ...
        def _write_worksheet_bin(self, sheet_file: BufferedWriter, worksheet_data: Iterable[list[Any]], worksheet_index: int, *, style_ranges: Any=(), header_style: Style | None=None, merge_refs: Any=(), widths: Any=None, hidden_cols: Any=frozenset(), heights: Any=None, hidden_row_set: Any=frozenset(), freeze_panes: str | None=None, auto_filter: bool | None=None, auto_filter_ref: str | None=None, drawing_rel: str | None=None, table_rel_ids: list[str] | None=None, cf_rules: list | None=None, row_outline: Any=None, col_outline: Any=None, sheet_protection: Any=None, comment_vml_rid: Any=None, show_grid_lines: Any=True, view: Any=(True, True, False), filter_cells: Any=(), header_rows: Any=1, tab_key: Any=None, zoom: Any=None, link_rids: dict | None=None, link_rid_start: int=2, print_spec: PrintSettings | None=None, validations: Any=(), sparklines: Any=(), slicer_rel: str | None=None) -> Any: ...
        def add_sheet(self, sheet_name: str, hidden: bool = False) -> None: ...
