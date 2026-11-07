from __future__ import annotations
from typing import Any
from collections.abc import Iterable
from . import _zip_threads as _zip_threads, chart_parts as chart_parts
from ._output import PackageZip as PackageZip, StagedWriter as StagedWriter, WriteTarget as WriteTarget
from ._writer_common import FilterData as FilterData, WriterFrontEnd as WriterFrontEnd, reject_unknown_arguments
from .doc_properties import DocumentPropertiesMixin as DocumentPropertiesMixin
from ._text import check_cell_text as check_cell_text, xstring_escape as xstring_escape
from .charts import Chart as Chart
from .data_validation import DataValidation as DataValidation, data_validation_xml as data_validation_xml, normalize_validations as normalize_validations
from .excel_dates import datetime_to_serial as datetime_to_serial
from ._widths import estimate_column_widths as estimate_column_widths, file_width as file_width
from .formats import DEFAULT_DATE_FORMAT as DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT as DEFAULT_DATETIME_FORMAT, _is_formatted_cell as _is_formatted_cell
from .formulas import Formula as Formula
from .comments import Comment as Comment, normalize_comments as normalize_comments, comments_xml as comments_xml, vml_xml as vml_xml, COMMENTS_TYPE as COMMENTS_TYPE
from ._metadata import DefinedName as DefinedName, SheetProtection as SheetProtection, WorkbookProtection as WorkbookProtection, replace_element as replace_element
from .ignored_errors import ignored_errors_xml as ignored_errors_xml
from ._extlst import (
    CT_SLICER as CT_SLICER,
    CT_SLICER_CACHE as CT_SLICER_CACHE,
    EXT_SHEET_SLICER_LIST as EXT_SHEET_SLICER_LIST,
    EXT_WORKBOOK_SLICER_CACHES as EXT_WORKBOOK_SLICER_CACHES,
    REL_SLICER as REL_SLICER,
    REL_SLICER_CACHE as REL_SLICER_CACHE,
    X14 as X14,
    X15 as X15,
    ext as ext,
    ext_list as ext_list,
)
from ._sheet_options import XLSX_OPTIONS as XLSX_OPTIONS, SheetOptions as SheetOptions, resolve_sheet_options as resolve_sheet_options
from .shapes import Shape as Shape, normalize_shapes as normalize_shapes
from .slicers import Slicer as Slicer, normalize_slicers as normalize_slicers, slicer_cache_xml as slicer_cache_xml, slicer_part_xml as slicer_part_xml
from .sparklines import normalize_sparklines as normalize_sparklines, sparklines_ext as sparklines_ext
from .pivots import (
    PivotFeeder as PivotFeeder,
    merge_cells as merge_cells,
    normalize_pivots as normalize_pivots,
    overlay_rows as overlay_rows,
    pivot_cache_xml as pivot_cache_xml,
    pivot_table_rels_xml as pivot_table_rels_xml,
    pivot_table_xml as pivot_table_xml,
    tee_rows as tee_rows,
)
from .hyperlinks import Hyperlink as Hyperlink
from ._scalars import numpy_value as numpy_value
from .images import Image as Image, MediaParts as MediaParts, normalize_images as normalize_images
from .outline import build_outline as build_outline
from .printing import (
    PageMargins as PageMargins,
    PrintSettings as PrintSettings,
    _paper_code as _paper_code,
    normalize_print_settings as normalize_print_settings,
    print_area_name as print_area_name,
    print_titles_name as print_titles_name,
)
from .refs import quote_sheet_name as quote_sheet_name
from .rich_text import CellRichText as CellRichText
from .styles import (
    CF_BLANK_RULES as CF_BLANK_RULES,
    CF_TEXT_RULES as CF_TEXT_RULES,
    CF_TIME_PERIOD_FORMULAS as CF_TIME_PERIOD_FORMULAS,
    EXTRA_CF_RULES as EXTRA_CF_RULES,
    HYPERLINK_STYLE as HYPERLINK_STYLE,
    AboveAverageRule as AboveAverageRule,
    CellIsRule as CellIsRule,
    Color as Color,
    DuplicateRule as DuplicateRule,
    TextRule as TextRule,
    TimePeriodRule as TimePeriodRule,
    TopBottomRule as TopBottomRule,
    ColorScaleRule as ColorScaleRule,
    DataBarRule as DataBarRule,
    DifferentialStyle as DifferentialStyle,
    ExcelTable as ExcelTable,
    FormatObject as FormatObject,
    FormulaRule as FormulaRule,
    IconSetRule as IconSetRule,
    Style as Style,
    check_header_rows as check_header_rows,
    check_sheet_view as check_sheet_view,
    check_zoom as check_zoom,
    header_style_ranges as header_style_ranges,
    tab_color_key as tab_color_key,
    _alignment_key_from as _alignment_key_from,
    _border_key as _border_key,
    _fill_key_from as _fill_key_from,
    _font_is_empty as _font_is_empty,
    _font_key_from as _font_key_from,
    _protection_key_from as _protection_key_from,
    _date_cell_xf as _date_cell_xf,
    _record_cell_xf as _record_cell_xf,
    _style_ids_for_row as _style_ids_for_row,
    _xf_for_style_format as _xf_for_style_format,
    capture_table_headers as capture_table_headers,
    explicit_cf_priorities as explicit_cf_priorities,
    merge_styles as merge_styles,
    normalize_hidden_columns as normalize_hidden_columns,
    normalize_hidden_rows as normalize_hidden_rows,
    normalize_style_ranges as normalize_style_ranges,
    parse_cell_range as parse_cell_range,
    split_style as split_style,
    table_column_names as table_column_names,
    styles_for_row as styles_for_row,
)
from ._xlsx_writer_xml import (
    _ATTR_ENTITIES as _ATTR_ENTITIES,
    _WIDTH_SAMPLE_ROWS as _WIDTH_SAMPLE_ROWS,
    _FLUSH_BYTES as _FLUSH_BYTES,
    _SST_CHUNK_STRINGS as _SST_CHUNK_STRINGS,
    DEFAULT_SPOOL_BYTES as DEFAULT_SPOOL_BYTES,
    _EMPTY_ROW_STYLES as _EMPTY_ROW_STYLES,
    _Tee as _Tee,
    _REL_TABLE as _REL_TABLE,
    _REL_DRAWING as _REL_DRAWING,
    _REL_HYPERLINK as _REL_HYPERLINK,
    _REL_WORKSHEET as _REL_WORKSHEET,
    _REL_PIVOT_TABLE as _REL_PIVOT_TABLE,
    _REL_PIVOT_CACHE as _REL_PIVOT_CACHE,
    _CT_PIVOT_TABLE as _CT_PIVOT_TABLE,
    _CT_PIVOT_CACHE as _CT_PIVOT_CACHE,
    _DEFAULT_MARGINS as _DEFAULT_MARGINS,
    _SheetBody as _SheetBody,
    _format_number as _format_number,
    _outline_col_attr as _outline_col_attr,
    _alignment_xml as _alignment_xml,
    _protection_xml as _protection_xml,
)
from ._xlsx_writer_sheet import (
    _XlsxSheetMixin as _XlsxSheetMixin,
)
from ._xlsx_writer_styles import (
    _XlsxStylesMixin as _XlsxStylesMixin,
)
from ._xlsx_writer_parts import (
    _XlsxPartsMixin as _XlsxPartsMixin,
)
from ._xlsx_writer_sheet_xml import (
    _XlsxSheetXmlMixin as _XlsxSheetXmlMixin,
)


class XlsxWriter(_XlsxSheetMixin, _XlsxStylesMixin, _XlsxPartsMixin, _XlsxSheetXmlMixin, WriterFrontEnd, StagedWriter, DocumentPropertiesMixin):
    """Write XLSX workbooks with a row-oriented API.

    A writer collects sheets, styles, tables, charts and other objects, then
    publishes the package when the ``with`` block exits. Rows are consumed once,
    so a generator or a database cursor can be passed straight to
    ``write_sheet``.

    Args:
        file: A path or a binary file object to write.
        compression_level: ZIP compression level, 0-9 (default 3).
        use_shared_strings: Store every distinct string once (default). Set it
            to ``False`` for high-cardinality exports.
        spool_max_bytes: Bytes kept in memory before sheet rows spill to a
            temporary file.

    Examples:
        Write a styled report and read it back:

        ```python
        from xlspy import ExcelReader, F, Style, XlsxWriter

        with XlsxWriter("report.xlsx") as workbook:
            workbook.add_sheet("Sales")
            workbook.write_sheet(
                [["Product", "Revenue"], ["Coffee", (12500, F.THOUSANDS_SEP)]],
                header_style=Style(bold=True, fill="183C52", font_color="FFFFFF"),
                freeze_panes="A2",
            )

        with ExcelReader("report.xlsx") as reader:
            print(reader.read_all("Sales"))
        ```

    See Also:
        `XlsbWriter` for the binary XLSB format with the same API, and
        `to_xlsx` for a DataFrame one-liner.
    """
    def __init__(self, file: WriteTarget, compression_level: int = 3, use_shared_strings: bool = True,
                 spool_max_bytes: int = DEFAULT_SPOOL_BYTES, **unknown: object) -> None:
        reject_unknown_arguments(type(self).__name__, unknown)
        self._file = file
        self._worksheet_data: list[tuple[str, Iterable[list[Any]] | None, bool]] = []
        self._shared_strings: list[str | tuple[str, str]] = []
        self._shared_strings_dict: dict[str | tuple[str, str], int] = {}
        self._sheet_count = 0
        self._sst_unique_count = 0
        self._sst_all_count = 0
        self._filtered_data_list: list[FilterData] = []
        self._sheet_number = 1
        self._compression_level = compression_level
        self._use_shared_strings = use_shared_strings
        self._spool_max_bytes = spool_max_bytes
        self._zf: PackageZip = None  # type: ignore[assignment]  # set when the workbook opens
        self._output = None  # StagedOutput while open
        self._letters = self._generate_column_letters()
        self._aborted = False
        self._saved = False
        # Set when a dynamic-array formula is written: xl/metadata.xml follows.
        self._has_dynamic = False

        # The default date and date-time formats, the same as XlsbWriter's.
        # Excel's built-in formats 14 and 22 would follow the reader's locale.
        self._format_registry: dict[str, int] = {DEFAULT_DATE_FORMAT: 164, DEFAULT_DATETIME_FORMAT: 165}
        self._format_xf_map: dict[str, int] = {}
        self._style_xf_map: dict[Style, int] = {}
        self._xf_styles: dict[int, Style] = {}
        self._style_format_xf: dict[tuple, tuple] = {}
        # xf of each (value, format) cell the native row encoder may reuse: keyed
        # by the format, or by (cell style xf, format) for a styled cell.
        self._fmt_xf: dict = {}
        self._fmt_xf_styles: list = []
        # Font id 0 is “no font”, id 1 is the builtin bold default emitted by
        # _create_styles_xml; custom fonts start at 2.
        self._font_registry: dict[tuple, int] = {}
        self._fill_registry: dict[tuple, int] = {}
        self._border_registry: dict[tuple, int] = {}
        self._style_components: dict[int, tuple] = {}
        self._dxf_list: list[tuple] = []
        self._dxf_index: dict[tuple, int] = {}
        self._dxf_font_registry: dict[tuple, int] = {}
        self._dxf_fill_registry: dict[tuple, int] = {}
        self._dxf_border_registry: dict[tuple, int] = {}
        self._seen_cf_priorities: set[int] = set()
        self._reserved_cf_priorities: set[int] = set()
        self._next_cf_priority = 1
        self._next_table_id = 1
        self._table_names: set[str] = set()
        self._tables_xml: dict[int, str] = {}
        # Casefolded table name -> (table id, column names), for slicer caches.
        self._table_info: dict[str, tuple[int, list[str]]] = {}
        # Slicer names, cache names and (table, column) targets used workbook-wide.
        self._slicer_names: set[str] = set()
        self._slicer_caches: set[str] = set()
        self._slicer_targets: set[tuple[str, str]] = set()
        # Each slicer owns one slicer part and one cache, both numbered alike.
        self._next_slicer_part = 1
        self._sheet_slicers: dict[int, list[Slicer]] = {}
        self._sheet_shapes: dict[int, list[Shape]] = {}
        # Pivot-table names, rendered pivots and cells waiting for their target sheet.
        self._pivot_names: set[str] = set()
        self._next_pivot_id = 1
        self._pivot_renders: list = []
        self._pending_pivots: dict[str, list] = {}
        # (path, xml, content type) of each slicer and cache part, in write order.
        self._slicer_part_files: list[tuple[str, str, str]] = []
        # Cache defined names; cache i is the workbook relationship of index i.
        self._slicer_cache_names: list[str] = []
        self._sheet_charts: dict[int, list] = {}
        self._sheet_chart_ids: dict[int, list[int]] = {}
        self._sheet_drawings: set[int] = set()
        self._sheet_images: dict[int, list[Image]] = {}
        self._sheet_image_media: dict[int, list[str]] = {}
        self._media = MediaParts()
        self._sheet_written: set[int] = set()
        self._sheet_rels: dict[int, list[tuple[str, str, str, bool]]] = {}
        self._sheet_print: dict[int, PrintSettings] = {}
        self._internal_links: list[tuple[int, str]] = []
        self._charts_xml: dict[int, tuple[int, int, str]] = {}
        self._next_chart_id = 1
        self._next_numfmt_id = 166
        self._next_xf_index = 4
        self._sheet_comments: dict[int, dict[str, Comment]] = {}
        self._defined_names: dict[tuple[str, str | None], DefinedName] = {}
        self._workbook_protection: WorkbookProtection | None = None

    def set_defined_name(self, definition: DefinedName) -> None:
        """Set a global or worksheet-local name; local titles are checked on save."""
        self._require_writable()
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be a DefinedName")
        self._defined_names[(definition.name.casefold(), definition.scope)] = definition

    def get_defined_names(self) -> list[DefinedName]:
        """The defined names set so far."""
        return list(self._defined_names.values())

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        """Remove a defined name; with ``scope``, the one local to that sheet."""
        self._require_writable()
        DefinedName(name, "0", scope)
        self._defined_names.pop((name.casefold(), scope), None)

    def _staging_bytes(self) -> int:
        # File-object targets are staged under the same in-memory limit as the rows.
        return self._spool_max_bytes

    def _generate_column_letters(self) -> list[str]:
        letters = []
        for i in range(26):
            letters.append(chr(65 + i))
        for i in range(26):
            for j in range(26):
                letters.append(chr(65 + i) + chr(65 + j))
        for i in range(24):
            for j in range(26):
                for k in range(26):
                    if i == 23 and j > 24:
                        break
                    letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))
        return letters

    def _register_format(self, fmt_string: str) -> int:
        if fmt_string in self._format_xf_map:
            return self._format_xf_map[fmt_string]

        # A style may already have registered this number format; reuse its id
        # so the style XF keeps referencing a format that is still emitted.
        numfmt_id = self._format_registry.get(fmt_string)
        if numfmt_id is None:
            numfmt_id = self._next_numfmt_id
            self._next_numfmt_id += 1
            self._format_registry[fmt_string] = numfmt_id

        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._format_xf_map[fmt_string] = xf_index

        return xf_index

    def _number_format_id(self, fmt_string: str | None) -> int | None:
        """The numFmt id for ``fmt_string``, registering it when it is new."""
        if fmt_string is None:
            return None
        numfmt_id = self._format_registry.get(fmt_string)
        if numfmt_id is None:
            numfmt_id = self._next_numfmt_id
            self._next_numfmt_id += 1
            self._format_registry[fmt_string] = numfmt_id
        return numfmt_id

    def _normalize_cf_rules(self, conditional_formats: Any) -> list[Any]:
        if not conditional_formats:
            return []
        for rule in conditional_formats:
            self._check_unknown_cf_rule(rule)
            if not isinstance(
                rule,
                (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule,
                 IconSetRule, *EXTRA_CF_RULES),
            ):
                raise TypeError(
                    f"unsupported conditional formatting rule {type(rule).__name__}"
                )
            rule_dxf = getattr(rule, "dxf", None)
            if rule_dxf is not None and not isinstance(rule_dxf, DifferentialStyle):
                raise TypeError("dxf must be a DifferentialStyle or None")
        return list(conditional_formats)

    def _written_sheet_indexes(self) -> set[int]:
        return self._sheet_written


    _CHART_AXES_XML = chart_parts._CHART_AXES_XML


