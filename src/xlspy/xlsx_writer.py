import datetime
from dataclasses import dataclass, field
from decimal import Decimal
import io
import shutil
import tempfile
import zipfile
import xml.sax.saxutils as saxutils
from typing import Iterable, Mapping, Tuple
from . import chart_parts
from ._accel import encode_xlsx_row as accel_encode_xlsx_row
from ._output import StagedWriter
from .doc_properties import DocumentPropertiesMixin
from ._text import check_cell_text, xstring_escape
from .charts import Chart
from .data_validation import DataValidation, data_validation_xml, normalize_validations
from .excel_dates import datetime_to_serial
from ._widths import estimate_column_widths, file_width
from .formats import DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT, _is_formatted_cell
from .formulas import Formula
from .comments import Comment, normalize_comments, comments_xml, vml_xml, COMMENTS_TYPE
from ._metadata import MAIN, DefinedName, SheetProtection, WorkbookProtection, replace_element
from .hyperlinks import Hyperlink
from .images import Image, MediaParts, normalize_images
from .outline import build_outline
from .printing import (
    PageMargins,
    PrintSettings,
    _paper_code,
    normalize_print_settings,
    print_area_name,
    print_titles_name,
)
from .refs import check_sheet_name, quote_sheet_name
from .rich_text import CellRichText
from .styles import (
    CF_BLANK_RULES,
    CF_TEXT_RULES,
    CF_TIME_PERIOD_FORMULAS,
    EXTRA_CF_RULES,
    HYPERLINK_STYLE,
    AboveAverageRule,
    CellIsRule,
    Color,
    DuplicateRule,
    TextRule,
    TimePeriodRule,
    TopBottomRule,
    ColorScaleRule,
    DataBarRule,
    DifferentialStyle,
    ExcelTable,
    FormatObject,
    FormulaRule,
    IconSetRule,
    Style,
    check_header_rows,
    check_zoom,
    header_style_ranges,
    tab_color_key,
    _alignment_key_from,
    _border_key,
    _column_index,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _protection_key_from,
    _date_cell_xf,
    _record_cell_xf,
    _style_ids_for_row,
    _xf_for_style_format,
    capture_table_headers,
    explicit_cf_priorities,
    merge_styles,
    normalize_hidden_columns,
    normalize_hidden_rows,
    normalize_style_ranges,
    parse_cell_range,
    split_style,
    table_column_names,
    styles_for_row,
)

# Rows sampled for automatic column widths (the historical limit).
_WIDTH_SAMPLE_ROWS = 100
# Encoded row text is handed to the spool once it grows past this size.
_FLUSH_CHARS = 1 << 20
# Shared-string entries written per chunk of sharedStrings.xml.
_SST_CHUNK_STRINGS = 4096
# Larger spools cost resident memory without making writes faster: 1 MiB
# and 8 MiB write 50,000 x 50 rows in the same time, 1 MiB with 14 MB less.
DEFAULT_SPOOL_BYTES = 1 << 20
_EMPTY_ROW_STYLES: dict = {}
_REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
_REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
_REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
_REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

_DEFAULT_MARGINS = PageMargins()


@dataclass
class _SheetBody:
    """What the row pass learned: sizes and the cells that need trailer entries."""

    num_rows: int
    max_cols: int
    head: list
    hyperlinks: list
    table_refs: set = field(default_factory=set)
    tables: list = field(default_factory=list)
    header_rows: dict = field(default_factory=dict)  # row index -> cells, for table headers


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0
    start_row: int = 0
    end_row: int = 0

def _format_number(value) -> str:
    """Shortest numeric literal: ``1.0`` -> ``"1"``, ``0.5`` -> ``"0.5"``."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _outline_col_attr(col_outline, index: int) -> str:
    """``<col>`` attributes for a column's outline level and collapsed flag."""
    if col_outline is None:
        return ""
    level, _ = col_outline.state_at(index)
    attrs = f' outlineLevel="{level}"' if level else ""
    if col_outline.collapsed_at(index):
        attrs += ' collapsed="1"'
    return attrs


def _alignment_xml(align) -> str:
    attrs = []
    if align.horizontal is not None:
        attrs.append(f'horizontal="{align.horizontal}"')
    if align.vertical is not None:
        attrs.append(f'vertical="{align.vertical}"')
    if align.wrap_text is not None:
        attrs.append(f'wrapText="{1 if align.wrap_text else 0}"')
    if align.shrink_to_fit is not None:
        attrs.append(f'shrinkToFit="{1 if align.shrink_to_fit else 0}"')
    if align.indent is not None:
        attrs.append(f'indent="{align.indent}"')
    if align.text_rotation is not None:
        attrs.append(f'textRotation="{align.text_rotation}"')
    if align.relative_indent is not None:
        attrs.append(f'relativeIndent="{align.relative_indent}"')
    if align.justify_last_line is not None:
        attrs.append(
            f'justifyLastLine="{1 if align.justify_last_line else 0}"'
        )
    if align.reading_order is not None:
        attrs.append(f'readingOrder="{align.reading_order}"')
    return f'<alignment {" ".join(attrs)}/>' if attrs else ""


def _protection_xml(prot) -> str:
    attrs = []
    if prot.locked is not None:
        attrs.append(f'locked="{1 if prot.locked else 0}"')
    if prot.hidden is not None:
        attrs.append(f'hidden="{1 if prot.hidden else 0}"')
    return f'<protection {" ".join(attrs)}/>' if attrs else ""


class XlsxWriter(StagedWriter, DocumentPropertiesMixin):
    def __init__(self, file: str | io.IOBase, compressionLevel: int = 4, useSharedStrings: bool = True,
                 spool_max_bytes: int = DEFAULT_SPOOL_BYTES):
        self._file = file
        self._worksheet_data: list[Tuple[str, Iterable[list[any]], bool]] = []
        self._shared_strings: list[str] = []
        self._shared_strings_dict: dict[str, int] = {}
        self._sheet_count = 0
        self._sst_unique_count = 0
        self._sst_all_count = 0
        self._filtered_data_list: list[FilterData] = []
        self._sheetCnt = 1
        self._compressionLevel = compressionLevel
        self._useSharedStrings = useSharedStrings
        self._spool_max_bytes = spool_max_bytes
        self._zf: zipfile.ZipFile = None
        self._output = None  # StagedOutput while open
        self._letters = self._generate_column_letters()
        self._aborted = False
        self._saved = False

        # The default date and date-time formats, the same as XlsbWriter's.
        # Excel's built-in formats 14 and 22 would follow the reader's locale.
        self._format_registry: dict[str, int] = {DEFAULT_DATE_FORMAT: 164, DEFAULT_DATETIME_FORMAT: 165}
        self._format_xf_map: dict[str, int] = {}
        self._style_xf_map: dict[Style, int] = {}
        self._xf_styles: dict[int, Style] = {}
        self._style_format_xf: dict[tuple, tuple] = {}
        # xf of each (value, format) cell the C row encoder may reuse: keyed
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
        return list(self._defined_names.values())

    def remove_defined_name(self, name: str, scope: str | None = None) -> None:
        self._require_writable()
        DefinedName(name, "0", scope)
        self._defined_names.pop((name.casefold(), scope), None)

    def set_workbook_protection(self, protection: WorkbookProtection) -> None:
        self._require_writable()
        if not isinstance(protection, WorkbookProtection):
            raise TypeError("protection must be WorkbookProtection")
        self._workbook_protection = protection

    def _staging_bytes(self):
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

    def add_sheet(self, sheet_name: str, hidden: bool = False):
        check_sheet_name(sheet_name, (name for name, _data, _hidden in self._worksheet_data))
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]), hidden))

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

    def _check_unknown_cf_rule(self, rule) -> None:
        if hasattr(rule, "sqref") and not isinstance(
            rule,
            (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule,
             *EXTRA_CF_RULES),
        ):
            raise TypeError(
                f"unsupported conditional formatting rule {type(rule).__name__}; "
                "use xlspy rule classes (CellIsRule, FormulaRule, "
                "ColorScaleRule, DataBarRule, IconSetRule, TopBottomRule, "
                "AboveAverageRule, TextRule, DuplicateRule, TimePeriodRule, BlanksRule)"
            )

    def _claim_cf_priority(self, rule) -> int:
        priority = rule.priority
        if priority is None:
            while (
                self._next_cf_priority in self._seen_cf_priorities
                or self._next_cf_priority in self._reserved_cf_priorities
            ):
                self._next_cf_priority += 1
            priority = self._next_cf_priority
            self._next_cf_priority += 1
        elif priority in self._seen_cf_priorities:
            # Explicit priorities are pre-registered per sheet, so reaching
            # this branch means the same rule was claimed twice.
            raise ValueError(f"duplicate conditional formatting priority {priority}")
        if priority <= 0:
            raise ValueError("conditional formatting priority must be positive")
        self._seen_cf_priorities.add(priority)
        self._next_cf_priority = max(self._next_cf_priority, priority + 1)
        return priority

    def _normalize_merges(self, merges):
        if not merges:
            return []
        refs = []
        for ref in merges:
            first_col, first_row, last_col, last_row = parse_cell_range(ref)
            refs.append((ref, first_col, first_row, last_col, last_row))
        for i, (ref_a, c1a, r1a, c2a, r2a) in enumerate(refs):
            for ref_b, c1b, r1b, c2b, r2b in refs[i + 1:]:
                if c1a <= c2b and c1b <= c2a and r1a <= r2b and r1b <= r2a:
                    raise ValueError(
                        f"overlapping merged ranges {ref_a!r} and {ref_b!r}"
                    )
        return [ref for ref, *_ in refs]

    def _normalize_cf_rules(self, conditional_formats):
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
            if getattr(rule, "dxf", None) is not None and not isinstance(
                rule.dxf, DifferentialStyle
            ):
                raise TypeError("dxf must be a DifferentialStyle or None")
        return list(conditional_formats)

    def _normalize_tables(self, tables):
        if not tables:
            return []
        # Names are only checked here; write_sheet registers them once every
        # option is valid, so a rejected call leaves no name behind.
        names = set(self._table_names)
        for table in tables:
            if not isinstance(table, ExcelTable):
                raise TypeError("tables must be ExcelTable instances")
            if table.display_name in names:
                raise ValueError(
                    f"duplicate table display name {table.display_name!r}"
                )
            names.add(table.display_name)
        return list(tables)

    @staticmethod
    def _normalize_widths(mapping, kind):
        if mapping is None:
            return {}
        if not isinstance(mapping, Mapping):
            raise TypeError(f"{kind} must be a mapping or None")
        result = {}
        for key, value in mapping.items():
            if isinstance(key, str):
                col = _column_index(key)
            elif isinstance(key, int) and not isinstance(key, bool):
                col = key
            else:
                raise TypeError(f"{kind} keys must be column letters or 0-based ints")
            if col < 0 or col >= 16384:
                raise ValueError(f"{kind} column out of range: {key!r}")
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise TypeError(f"{kind} widths must be numbers")
            if not 0.5 <= value <= 255:
                raise ValueError(f"{kind} width must be 0.5..255")
            result[col] = float(value)
        return result

    @staticmethod
    def _normalize_row_heights(mapping):
        if mapping is None:
            return {}
        if not isinstance(mapping, Mapping):
            raise TypeError("row_heights must be a mapping or None")
        result = {}
        for key, value in mapping.items():
            if not isinstance(key, int) or isinstance(key, bool) or key < 0:
                raise TypeError("row_heights keys must be 0-based row ints")
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise TypeError("row_heights values must be numbers")
            if value <= 0:
                raise ValueError("row height must be > 0")
            result[key] = float(value)
        return result

    def _column_letter(self, index: int) -> str:
        return self._letters[index]

    def _split_freeze_cell(self, ref: str) -> tuple[int, int]:
        first_col, first_row, last_col, last_row = parse_cell_range(ref)
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"freeze_panes must be a single cell, got {ref!r}")
        return first_col, first_row

    @staticmethod
    def _normalize_charts(charts) -> list:
        if not charts:
            return []
        if not isinstance(charts, (list, tuple)):
            raise TypeError("charts must be a list of Chart objects or None")
        for chart in charts:
            if not isinstance(chart, Chart):
                raise TypeError("charts must be Chart instances")
        return list(charts)

    def write_sheet(
        self,
        data: Iterable[list[any]],
        *,
        header_style: Style | None = None,
        styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        conditional_formats: list | tuple | None = None,
        tables: list[ExcelTable] | tuple[ExcelTable, ...] | None = None,
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
        data_validations: list[DataValidation] | tuple[DataValidation, ...] | None = None,
        images: list[Image] | tuple[Image, ...] | None = None,
        print_settings: PrintSettings | None = None,
        comments: Mapping[str, Comment] | None = None,
        sheet_protection: SheetProtection | None = None,
        show_grid_lines: bool = True,
        header_rows: int = 1,
        tab_color: Color | str | None = None,
        zoom: int | None = None,
    ):
        """Write rows to the current sheet.

        ``data`` is consumed once, row by row. The rows are encoded into a
        temporary spool (kept in memory up to ``spool_max_bytes``, then on
        disk), so memory does not grow with the number of rows. Hyperlinks,
        formulas and the other options are described in the documentation.
        """
        self._require_writable()
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        if not isinstance(show_grid_lines, bool):
            raise TypeError("show_grid_lines must be a bool")
        header_rows = check_header_rows(header_rows)
        tab_key = tab_color_key(tab_color)
        zoom = check_zoom(zoom)
        style_ranges, header_style = header_style_ranges(
            normalize_style_ranges(styles), header_style, header_rows,
        )
        merge_refs = self._normalize_merges(merges)
        comment_specs = normalize_comments(comments, merges=merge_refs)
        if sheet_protection is not None and not isinstance(sheet_protection, SheetProtection):
            raise TypeError("sheet_protection must be SheetProtection")
        if auto_filter_ref is not None:
            # Validate before rows are streamed: a late error aborts the writer.
            if not isinstance(auto_filter_ref, str):
                raise TypeError("auto_filter_ref must be an A1 range string")
            parse_cell_range(auto_filter_ref)
        cf_rules = self._normalize_cf_rules(conditional_formats)
        table_specs = self._normalize_tables(tables)
        chart_specs = self._normalize_charts(charts)
        validations = normalize_validations(data_validations)
        image_specs = normalize_images(images)
        print_spec = normalize_print_settings(print_settings)
        widths = self._normalize_widths(column_widths, "column_widths")
        hidden_cols = normalize_hidden_columns(hidden_columns)
        heights = self._normalize_row_heights(row_heights)
        hidden_row_set = normalize_hidden_rows(hidden_rows)
        row_outline = build_outline(row_groups, "row")
        col_outline = build_outline(column_groups, "column")
        if col_outline is not None:
            hidden_cols = hidden_cols | col_outline.hidden_indices()
        split_freeze = None
        if freeze_panes is not None and freeze_panes != "":
            split_freeze = self._split_freeze_cell(freeze_panes)
        cf_priorities = explicit_cf_priorities(
            cf_rules, self._seen_cf_priorities, self._reserved_cf_priorities,
        )
        # Every option is valid; only now record what this sheet claims.
        self._table_names.update(table.display_name for table in table_specs)
        for priority in cf_priorities:
            self._reserved_cf_priorities.add(priority)
            self._next_cf_priority = max(self._next_cf_priority, priority + 1)
        if not self._worksheet_data:
            self.add_sheet("Sheet1")

        sheet_index = self._sheet_count - 1
        sheet_name, _, hidden = self._worksheet_data[sheet_index]
        self._worksheet_data[sheet_index] = (sheet_name, None, hidden)
        sheet_id = self._sheet_count
        has_styles = bool(style_ranges) or header_style is not None

        spool = tempfile.SpooledTemporaryFile(max_size=self._spool_max_bytes, mode="w+b")
        try:
            body = self._write_body(
                data, spool, style_ranges, header_style if has_styles else None,
                heights, hidden_row_set, has_styles,
                table_specs,
                row_outline=row_outline,
            )
            table_rels, table_bounds = self._register_tables(table_specs, sheet_id, body)
            rels = list(table_rels)
            link_rels = self._link_relationships(body.hyperlinks, rels)
            rels.extend(link_rels)
            drawing_rid = None
            if chart_specs or image_specs:
                drawing_rid = f"rId{len(rels) + 1}"
                rels.append((drawing_rid, _REL_DRAWING,
                             f"../drawings/drawing{sheet_id}.xml", False))
                self._sheet_drawings.add(sheet_id)
            if chart_specs:
                self._sheet_charts[sheet_id] = chart_specs
            if image_specs:
                self._sheet_images[sheet_id] = image_specs
                self._sheet_image_media[sheet_id] = [
                    self._media.name(image) for image in image_specs
                ]
            if print_spec is not None:
                self._sheet_print[sheet_index] = print_spec
            for row_idx, col_idx, link in body.hyperlinks:
                if link.is_internal:
                    self._internal_links.append((sheet_index, link.location))
            if auto_filter is None and table_specs:
                auto_filter = False
            # Header cells under a filter drop-down button need room for it.
            filter_cells = [
                (r1, c1, c2) for table, (c1, r1, c2, _r2) in zip(table_specs, table_bounds)
                if table.header_row
            ]
            if auto_filter is not False:
                if auto_filter_ref is not None:
                    c1, r1, c2, _r2 = parse_cell_range(auto_filter_ref)
                    filter_cells.append((r1, c1, c2))
                elif body.num_rows > header_rows:
                    filter_cells.append((header_rows - 1, 0, body.max_cols - 1))
            head = self._sheet_head(
                body, sheet_index, split_freeze, freeze_panes, widths, hidden_cols,
                print_spec, row_outline=row_outline, col_outline=col_outline,
                show_grid_lines=show_grid_lines,
                width_styles=(style_ranges, header_style if has_styles else None),
                filter_cells=filter_cells, tab_key=tab_key, zoom=zoom,
            )
            tail = self._sheet_tail(
                body, sheet_index, merge_refs, cf_rules, table_specs, table_bounds,
                validations, link_rels, auto_filter, drawing_rid, print_spec,
                auto_filter_ref=auto_filter_ref, header_rows=header_rows,
            )
            if sheet_protection is not None:
                tail = '</sheetData>' + sheet_protection.to_xml() + tail[len('</sheetData>'):]
            if comment_specs:
                self._sheet_comments[sheet_id] = comment_specs
                rid = f"rId{len(rels) + 1}"
                rels.append((rid, _REL_WORKSHEET.rsplit('/', 1)[0] + '/comments',
                             f'../comments/comment{sheet_id}.xml', False))
                vml_rid = f"rId{len(rels) + 1}"
                rels.append((vml_rid, _REL_WORKSHEET.rsplit('/', 1)[0] + '/vmlDrawing',
                             f'../drawings/commentsDrawing{sheet_id}.vml', False))
                tail = replace_element(tail, 'legacyDrawing', f'<legacyDrawing r:id="{vml_rid}"/>', before=('tableParts',))
            if rels:
                self._sheet_rels[sheet_id] = rels
            with self._zf.open(f"xl/worksheets/sheet{sheet_id}.xml", 'w') as sheet_file:
                sheet_file.write(head.encode('utf-8'))
                spool.seek(0)
                shutil.copyfileobj(spool, sheet_file, 1 << 20)
                sheet_file.write(tail.encode('utf-8'))
            self._sheet_written.add(sheet_index)
        except BaseException:
            # The sheet part is incomplete, so the workbook cannot be saved.
            self._aborted = True
            raise
        finally:
            spool.close()

    def _intern_font(self, font_key) -> int:
        if _font_is_empty(font_key):
            return 0
        existing = self._font_registry.get(font_key)
        if existing is not None:
            return existing
        font_id = 2 + len(self._font_registry)
        self._font_registry[font_key] = font_id
        return font_id

    def _intern_fill(self, fill_key) -> int:
        if fill_key is None:
            return 0
        existing = self._fill_registry.get(fill_key)
        if existing is not None:
            return existing
        fill_id = 2 + len(self._fill_registry)
        self._fill_registry[fill_key] = fill_id
        return fill_id

    def _intern_border(self, border_key) -> int:
        if border_key is None:
            return 0
        existing = self._border_registry.get(border_key)
        if existing is not None:
            return existing
        border_id = 1 + len(self._border_registry)
        self._border_registry[border_key] = border_id
        return border_id

    def _register_dxf_style(self, dxf: DifferentialStyle) -> int:
        font_key = None if dxf.font is None else _font_key_from(dxf.font)
        if _font_is_empty(font_key):
            font_id = None
        else:
            if font_key in self._dxf_font_registry:
                font_id = self._dxf_font_registry[font_key]
            else:
                font_id = len(self._dxf_font_registry)
                self._dxf_font_registry[font_key] = font_id
        if dxf.fill is None:
            fill_id = None
        else:
            fill_key = _fill_key_from(dxf.fill)
            if fill_key is None:
                fill_id = None
            elif fill_key in self._dxf_fill_registry:
                fill_id = self._dxf_fill_registry[fill_key]
            else:
                fill_id = len(self._dxf_fill_registry)
                self._dxf_fill_registry[fill_key] = fill_id
        if dxf.border is None:
            border_id = None
        else:
            border_key = _border_key(dxf.border)
            if border_key is None:
                border_id = None
            elif border_key in self._dxf_border_registry:
                border_id = self._dxf_border_registry[border_key]
            else:
                border_id = len(self._dxf_border_registry)
                self._dxf_border_registry[border_key] = border_id
        align_key = (
            _alignment_key_from(dxf.alignment)
            if dxf.alignment is not None else None
        )
        if align_key is not None and not any(
            value is not None for value in align_key
        ):
            align_key = None
        prot_key = (
            _protection_key_from(dxf.protection)
            if dxf.protection is not None else None
        )
        if prot_key is not None and not any(
            value is not None for value in prot_key
        ):
            prot_key = None
        if (
            dxf.number_format is not None
            and dxf.number_format not in self._format_registry
        ):
            self._format_registry[dxf.number_format] = self._next_numfmt_id
            self._next_numfmt_id += 1
        dxf_entry = (
            font_id, fill_id, border_id, dxf.number_format,
            align_key, prot_key,
        )
        existing = self._dxf_index.get(dxf_entry)
        if existing is not None:
            return existing
        index = len(self._dxf_list)
        self._dxf_list.append(dxf_entry)
        self._dxf_index[dxf_entry] = index
        return index

    def _register_style(self, style: Style) -> int:
        if style in self._style_xf_map:
            return self._style_xf_map[style]
        font_key, fill_key, border_key, numfmt, _align, _prot = split_style(style)
        font_id = self._intern_font(font_key)
        fill_id = self._intern_fill(fill_key)
        border_id = self._intern_border(border_key)
        numfmt_id = 0
        if numfmt is not None:
            if numfmt not in self._format_registry:
                self._format_registry[numfmt] = self._next_numfmt_id
                self._next_numfmt_id += 1
            numfmt_id = self._format_registry[numfmt]
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._style_xf_map[style] = xf_index
        self._xf_styles[xf_index] = style
        self._style_components[xf_index] = (
            font_id, fill_id, border_id, numfmt_id, _align, _prot,
        )
        return xf_index

    def write_dataframe(
        self, df, include_index: bool = False, na_rep=None, formats: dict = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        conditional_formats: list | tuple | None = None,
        tables: list[ExcelTable] | tuple[ExcelTable, ...] | None = None,
        column_widths: dict | None = None,
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
    ):
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
        from .dataframe import dataframe_to_rows

        self.write_sheet(
            dataframe_to_rows(df, include_index=include_index, na_rep=na_rep, formats=formats),
            header_style=header_style,
            styles=styles,
            merges=merges,
            conditional_formats=conditional_formats,
            tables=tables,
            column_widths=column_widths,
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
        )

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
                chart_parts.drawing_xml(charts, chart_ids, images=images),
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

        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.xml", self._create_workbook_xml())
        self._zf.writestr("xl/styles.xml", self._create_styles_xml())
        self._zf.writestr("xl/_rels/workbook.xml.rels", self._create_workbook_rels())
        self._write_doc_props()

        if self._useSharedStrings and self._shared_strings:
            with self._zf.open("xl/sharedStrings.xml", 'w') as sst_file:
                self._write_shared_strings_xml(sst_file)

    @staticmethod
    def _quote_sheet_name(sheet_name: str) -> str:
        return chart_parts.quote_sheet_name(sheet_name)

    def _chart_ref(self, sheet_name: str, ref: str, *, absolute: bool = True) -> str:
        return chart_parts.chart_ref(self._letters, sheet_name, ref, absolute=absolute)

    def _chart_series_xml(self, series, sheet_name: str, index: int, chart_type: str) -> str:
        return chart_parts.chart_series_xml(
            self._letters, series, sheet_name, index, chart_type,
        )

    _CHART_AXES_XML = chart_parts._CHART_AXES_XML

    def _chart_xml(self, chart, sheet_name: str) -> str:
        return chart_parts.chart_xml(self._letters, chart, sheet_name)

    def _drawing_xml(self, charts: list, chart_ids: list[int]) -> str:
        return chart_parts.drawing_xml(charts, chart_ids)

    @staticmethod
    def _drawing_rels_xml(chart_ids: list[int]) -> str:
        return chart_parts.drawing_rels_xml(chart_ids)

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
        if self._useSharedStrings and self._shared_strings:
            shared_strings_part = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'

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
        
        if self._useSharedStrings and self._shared_strings:
            shared_strings_rid = self._sheet_count + 2
            relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')

        if self._sheet_drawings:
            theme_rid = len(relationships) + 1
            relationships.append(f'<Relationship Id="rId{theme_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _create_workbook_xml(self) -> str:
        sheets = ""
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            hidden_attr = ' state="hidden"' if hidden else ''
            sheets += f'<sheet name="{saxutils.escape(sheet_name)}" sheetId="{sheet_id}"{hidden_attr} r:id="rId{sheet_id}"/>'

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
        defined_names = f"<definedNames>{''.join(names)}</definedNames>" if names else ""
        protection = self._workbook_protection.to_xml() if self._workbook_protection else ""

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
{protection}
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>{sheets}</sheets>
{defined_names}
<calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>'''

    @staticmethod
    def _color_attrs(color_key) -> str:
        kind = color_key[0]
        if kind == "rgb":
            return f'rgb="{color_key[1]}"'
        if kind == "indexed":
            return f'indexed="{color_key[1]}"'
        if kind == "auto":
            return 'auto="1"'
        _, theme, tint = color_key
        if tint is None:
            return f'theme="{theme}"'
        return f'theme="{theme}" tint="{tint}"'

    def _color_xml(self, color_key) -> str:
        if color_key is None:
            return '<color theme="1"/>'
        return f'<color {self._color_attrs(color_key)}/>'

    def _font_xml(self, font_key, *, for_dxf: bool = False) -> str:
        name = font_key.name
        size = font_key.size
        parts = []
        if size is not None:
            parts.append(f'<sz val="{size}"/>')
        if for_dxf:
            # A dxf font applies only the properties it names; emitting the
            # Calibri/theme-1 defaults would override the cell's own font.
            if font_key.color is not None:
                parts.append(self._color_xml(font_key.color))
            if name is not None:
                label = saxutils.escape(name, {'"': '&quot;'})
                parts.append(f'<name val="{label}"/>')
            if font_key.family is not None:
                parts.append(f'<family val="{font_key.family}"/>')
            if font_key.scheme is not None:
                parts.append(f'<scheme val="{font_key.scheme}"/>')
        else:
            parts.append(self._color_xml(font_key.color))
            label = saxutils.escape(name, {'"': '&quot;'}) if name is not None else "Calibri"
            parts.append(f'<name val="{label}"/>')
            if font_key.family is not None:
                parts.append(f'<family val="{font_key.family}"/>')
            else:
                parts.append('<family val="2"/>')
            if font_key.scheme is not None:
                parts.append(f'<scheme val="{font_key.scheme}"/>')
            elif name is None:
                # Only the built-in Calibri default carries scheme="minor".
                # A scheme next to an explicit font name would make Excel
                # substitute the theme font for that name.
                parts.append('<scheme val="minor"/>')
        if font_key.charset is not None:
            parts.append(f'<charset val="{font_key.charset}"/>')
        if font_key.bold:
            parts.append('<b/>')
        if font_key.italic:
            parts.append('<i/>')
        if font_key.underline is not None:
            if font_key.underline == "single":
                parts.append('<u/>')
            else:
                parts.append(f'<u val="{font_key.underline}"/>')
        if font_key.strike:
            parts.append('<strike/>')
        if font_key.vert_align is not None:
            parts.append(f'<vertAlign val="{font_key.vert_align}"/>')
        if font_key.outline:
            parts.append('<outline/>')
        if font_key.shadow:
            parts.append('<shadow/>')
        if font_key.condense:
            parts.append('<condense/>')
        if font_key.extend:
            parts.append('<extend/>')
        return f"<font>{''.join(parts)}</font>"

    def _fill_xml(self, fill_key, *, for_dxf: bool = False) -> str:
        if fill_key[0] == "gradient":
            (
                _, fill_type, degree, left, right, top, bottom, stops,
            ) = fill_key
            attrs = [f'type="{fill_type}"']
            if degree is not None:
                attrs.append(f'degree="{_format_number(degree)}"')
            for name, value in (
                ("left", left), ("right", right), ("top", top),
                ("bottom", bottom),
            ):
                if value is not None:
                    attrs.append(f'{name}="{_format_number(value)}"')
            stop_xml = "".join(
                f'<stop position="{_format_number(position)}">'
                f'<color {self._color_attrs(color_key)}/></stop>'
                for color_key, position in stops
            )
            return (
                f'<fill><gradientFill {" ".join(attrs)}>'
                f"{stop_xml}</gradientFill></fill>"
            )
        _, pattern, fg_key, bg_key = fill_key
        pattern_type = pattern if pattern is not None else "none"
        if pattern_type == "solid":
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            if bg_key is not None:
                bg_xml = f'<bgColor {self._color_attrs(bg_key)}/>'
            elif for_dxf:
                # A conditional format paints a solid fill with its background
                # color, so a fill given only start_color would not show.
                bg_xml = f'<bgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            else:
                bg_xml = '<bgColor indexed="64"/>'
        else:
            fg_xml = f'<fgColor {self._color_attrs(fg_key)}/>' if fg_key is not None else ""
            bg_xml = f'<bgColor {self._color_attrs(bg_key)}/>' if bg_key is not None else ""
        return f'<fill><patternFill patternType="{pattern_type}">{fg_xml}{bg_xml}</patternFill></fill>'

    @staticmethod
    def _border_side_xml(tag, side_key, color_attrs) -> str:
        style, color_key = side_key
        if style is None and color_key is None:
            return f"<{tag}/>"
        inner = f"<color {color_attrs(color_key)}/>" if color_key is not None else ""
        if style is None:
            return f"<{tag}>{inner}</{tag}>"
        return f'<{tag} style="{style}">{inner}</{tag}>'

    def _border_xml(self, border_key) -> str:
        attrs = ""
        if border_key.diagonal_up:
            attrs += ' diagonalUp="1"'
        if border_key.diagonal_down:
            attrs += ' diagonalDown="1"'
        if border_key.outline is False:
            attrs += ' outline="0"'
        # Excel reads <start>/<end> as the left/right edges (in a
        # left-to-right sheet), so an empty <start/> would hide <left>:
        # write them only when set, like <vertical>/<horizontal>.
        parts = [
            self._border_side_xml(tag, getattr(border_key, tag), self._color_attrs)
            for tag in ("start", "end")
            if getattr(border_key, tag) != (None, None)
        ]
        parts += [
            self._border_side_xml("left", border_key.left, self._color_attrs),
            self._border_side_xml("right", border_key.right, self._color_attrs),
            self._border_side_xml("top", border_key.top, self._color_attrs),
            self._border_side_xml("bottom", border_key.bottom, self._color_attrs),
            self._border_side_xml("diagonal", border_key.diagonal, self._color_attrs),
        ]
        if border_key.vertical != (None, None):
            parts.append(self._border_side_xml("vertical", border_key.vertical, self._color_attrs))
        if border_key.horizontal != (None, None):
            parts.append(self._border_side_xml("horizontal", border_key.horizontal, self._color_attrs))
        return f"<border{attrs}>{''.join(parts)}</border>"

    def _create_styles_xml(self) -> str:
        numfmts = ""
        xf_entries = []

        numfmt_count = len(self._format_registry)
        if numfmt_count > 0:
            numfmts_parts = []
            for fmt_string, numfmt_id in self._format_registry.items():
                escaped_fmt = saxutils.escape(fmt_string, {'"': '&quot;'})
                numfmts_parts.append(f'<numFmt numFmtId="{numfmt_id}" formatCode="{escaped_fmt}"/>')
            numfmts = f'<numFmts count="{numfmt_count}">{"".join(numfmts_parts)}</numFmts>'

        xf_entries.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
        xf_entries.append('<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>')

        # Format XFs and style XFs share one index space, so they must be
        # emitted in index order: the position in cellXfs is the value a cell
        # references through its s= attribute.
        custom_xfs = {
            xf_index: f'<xf numFmtId="{self._format_registry[fmt_string]}" fontId="0" '
                      f'fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            for fmt_string, xf_index in self._format_xf_map.items()
        }
        for xf_index, components in self._style_components.items():
            font_id, fill_id, border_id, numfmt_id, align, prot = components
            attrs = [f'numFmtId="{numfmt_id}"', f'fontId="{font_id}"',
                     f'fillId="{fill_id}"', f'borderId="{border_id}"', 'xfId="0"']
            if numfmt_id != 0:
                attrs.append('applyNumberFormat="1"')
            if font_id != 0:
                attrs.append('applyFont="1"')
            if fill_id != 0:
                attrs.append('applyFill="1"')
            if border_id != 0:
                attrs.append('applyBorder="1"')
            align_xml = _alignment_xml(align)
            if align_xml:
                attrs.append('applyAlignment="1"')
            prot_xml = _protection_xml(prot)
            if prot_xml:
                attrs.append('applyProtection="1"')
            custom_xfs[xf_index] = f'<xf {" ".join(attrs)}>{align_xml}{prot_xml}</xf>'

        for xf_index in sorted(custom_xfs):
            xf_entries.append(custom_xfs[xf_index])

        total_xf = len(xf_entries)

        fonts = [
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
            '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/><b/></font>',
        ]
        for key, font_id in sorted(self._font_registry.items(), key=lambda item: item[1]):
            if font_id < 2:
                continue
            fonts.append(self._font_xml(key))
        fills = [
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
        ]
        fills.extend(
            self._fill_xml(key)
            for key, _fill_id in sorted(self._fill_registry.items(), key=lambda item: item[1])
        )
        borders = [
            '<border><left/><right/><top/><bottom/><diagonal/></border>',
        ]
        borders.extend(
            self._border_xml(key)
            for key, _border_id in sorted(self._border_registry.items(), key=lambda item: item[1])
        )

        dxf_fonts: list = [None] * len(self._dxf_font_registry)
        for key, dxf_font_id in self._dxf_font_registry.items():
            dxf_fonts[dxf_font_id] = key
        dxf_fills: list = [None] * len(self._dxf_fill_registry)
        for key, dxf_fill_id in self._dxf_fill_registry.items():
            dxf_fills[dxf_fill_id] = key
        dxf_borders: list = [None] * len(self._dxf_border_registry)
        for key, dxf_border_id in self._dxf_border_registry.items():
            dxf_borders[dxf_border_id] = key
        dxf_parts = []
        for font_id, fill_id, border_id, numfmt, align, prot in self._dxf_list:
            inner = ""
            if font_id is not None:
                inner += self._font_xml(dxf_fonts[font_id], for_dxf=True)
            if numfmt is not None:
                escaped_fmt = saxutils.escape(numfmt, {'"': '&quot;'})
                numfmt_id = self._format_registry[numfmt]
                inner += (
                    f'<numFmt numFmtId="{numfmt_id}" '
                    f'formatCode="{escaped_fmt}"/>'
                )
            if fill_id is not None:
                inner += self._fill_xml(dxf_fills[fill_id], for_dxf=True)
            if align is not None:
                inner += _alignment_xml(align)
            if border_id is not None:
                inner += self._border_xml(dxf_borders[border_id])
            if prot is not None:
                inner += _protection_xml(prot)
            dxf_parts.append(f"<dxf>{inner}</dxf>")
        dxfs = f'<dxfs count="{len(dxf_parts)}">{"".join(dxf_parts)}</dxfs>' if dxf_parts else '<dxfs count="0"/>'

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{numfmts}
<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>
<fills count="{len(fills)}">{"".join(fills)}</fills>
<borders count="{len(borders)}">
{"".join(borders)}
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="{total_xf}">
{"".join(xf_entries)}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
{dxfs}
<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    def _write_shared_strings_xml(self, sst_file: io.BufferedWriter):
        # Written in chunks: the table can hold millions of strings, and one
        # growing string would keep the whole encoded part in memory.
        header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                  f'count="{self._sst_all_count}" uniqueCount="{self._sst_unique_count}">')
        sst_file.write(header.encode('utf-8'))
        parts: list[str] = []
        for s in self._shared_strings:
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
            if len(parts) >= _SST_CHUNK_STRINGS:
                # The markup holds no escapable characters, so a chunk is escaped at once.
                sst_file.write(xstring_escape(''.join(parts)).encode('utf-8'))
                parts.clear()
        parts.append('</sst>')
        sst_file.write(xstring_escape(''.join(parts)).encode('utf-8'))

    @staticmethod
    def _format_object_xml(cfvo) -> str:
        attrs = f'type="{cfvo.type}"'
        if cfvo.value is not None:
            attrs += f' val="{saxutils.escape(str(cfvo.value), {"\"": "&quot;"})}"'
        return f"<cfvo {attrs}/>"

    def _extra_cf_rule_parts(self, rule) -> tuple[str, str | None]:
        """``(cfRule attributes, formula or None)`` for the rules beyond the original five."""
        first_col, first_row, _, _ = parse_cell_range(rule.sqref)
        corner = f"{self._letters[first_col]}{first_row + 1}"
        if isinstance(rule, TopBottomRule):
            attrs = f'type="top10" rank="{rule.rank}"'
            if rule.percent:
                attrs += ' percent="1"'
            if rule.bottom:
                attrs += ' bottom="1"'
            return attrs, None
        if isinstance(rule, AboveAverageRule):
            attrs = 'type="aboveAverage"'
            if rule.below:
                attrs += ' aboveAverage="0"'
            if rule.equal_average:
                attrs += ' equalAverage="1"'
            if rule.std_dev is not None:
                attrs += f' stdDev="{rule.std_dev}"'
            return attrs, None
        if isinstance(rule, DuplicateRule):
            return f'type="{"uniqueValues" if rule.unique else "duplicateValues"}"', None
        if isinstance(rule, TextRule):
            kind, operator, template = CF_TEXT_RULES[rule.operator]
            quoted = '"' + rule.text.replace('"', '""') + '"'
            attrs = (
                f'type="{kind}" operator="{operator}" '
                f'text="{saxutils.escape(rule.text, {chr(34): "&quot;"})}"'
            )
            return attrs, template.format(t=quoted, c=corner)
        if isinstance(rule, TimePeriodRule):
            return (
                f'type="timePeriod" timePeriod="{rule.period}"',
                CF_TIME_PERIOD_FORMULAS[rule.period].format(c=corner),
            )
        kind, template = CF_BLANK_RULES[rule.kind]
        return f'type="{kind}"', template.format(c=corner)

    def _cf_rule_xml(self, rule, table_ranges: set[tuple]) -> str:
        priority = self._claim_cf_priority(rule)
        stop = ""
        if getattr(rule, "stop_if_true", None):
            stop = ' stopIfTrue="1"'
        dxf_part = ""
        dxf_index = None
        dxf = getattr(rule, "dxf", None)
        if dxf is not None:
            dxf_index = self._register_dxf_style(dxf)
            dxf_part = f' dxfId="{dxf_index}"'
        if isinstance(rule, CellIsRule):
            formulas = "".join(
                f"<formula>{saxutils.escape(text)}</formula>"
                for text in rule.formula
            )
            return (
                f'<cfRule type="cellIs" operator="{rule.operator}" '
                f'priority="{priority}"{dxf_part}{stop}>{formulas}</cfRule>'
            )
        if isinstance(rule, FormulaRule):
            formulas = "".join(
                f"<formula>{saxutils.escape(text)}</formula>"
                for text in rule.formula
            )
            return (
                f'<cfRule type="expression" priority="{priority}"'
                f'{dxf_part}{stop}>{formulas}</cfRule>'
            )
        if isinstance(rule, EXTRA_CF_RULES):
            attrs, formula = self._extra_cf_rule_parts(rule)
            formula_xml = f"<formula>{saxutils.escape(formula)}</formula>" if formula else ""
            return (
                f'<cfRule {attrs} priority="{priority}"{dxf_part}{stop}'
                + (f">{formula_xml}</cfRule>" if formula_xml else "/>")
            )
        if isinstance(rule, ColorScaleRule):
            if rule.midpoint is None:
                cfvos = (rule.minimum, rule.maximum)
            else:
                cfvos = (rule.minimum, rule.midpoint, rule.maximum)
            inner = "".join(
                self._format_object_xml(cfvo) if cfvo is not None
                else '<cfvo type="min"/>' if i == 0 else '<cfvo type="max"/>'
                for i, cfvo in enumerate(cfvos)
            )
            colors = "".join(
                f'<color rgb="{self._dxf_color_argb(color)}"/>'
                for color in rule.colors
            )
            return (
                f'<cfRule type="colorScale" priority="{priority}"{stop}>'
                f"<colorScale>{inner}{colors}</colorScale></cfRule>"
            )
        if isinstance(rule, IconSetRule):
            attrs = [f'iconSet="{rule.icon_style}"']
            if rule.show_value is not None:
                attrs.append(f'showValue="{1 if rule.show_value else 0}"')
            if rule.percent is not None:
                attrs.append(f'percent="{1 if rule.percent else 0}"')
            if rule.reverse is not None:
                attrs.append(f'reverse="{1 if rule.reverse else 0}"')
            cfvos = "".join(
                self._format_object_xml(FormatObject(rule.type, value))
                for value in rule.values
            )
            return (
                f'<cfRule type="iconSet" priority="{priority}"{stop}>'
                f'<iconSet {" ".join(attrs)}>{cfvos}</iconSet></cfRule>'
            )
        # DataBarRule
        min_xml = (
            self._format_object_xml(rule.minimum)
            if rule.minimum is not None else '<cfvo type="min"/>'
        )
        max_xml = (
            self._format_object_xml(rule.maximum)
            if rule.maximum is not None else '<cfvo type="max"/>'
        )
        show = "" if rule.show_value is None else (
            ' showValue="1"' if rule.show_value else ' showValue="0"'
        )
        min_len = "" if rule.min_length is None else f' minLength="{rule.min_length}"'
        max_len = "" if rule.max_length is None else f' maxLength="{rule.max_length}"'
        color = (
            f'<color rgb="{self._dxf_color_argb(rule.color)}"/>'
        )
        return (
            f'<cfRule type="dataBar" priority="{priority}"{stop}>'
            f"<dataBar{min_len}{max_len}{show}>{min_xml}{max_xml}{color}</dataBar>"
            "</cfRule>"
        )

    @staticmethod
    def _dxf_color_argb(color) -> str:
        from .styles import _coerce_color
        resolved = _coerce_color(color, "color")
        if resolved is None or resolved.rgb is None:
            raise ValueError("CF colorScale/dataBar colors must be RGB colors")
        return resolved.rgb

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
        style_xml = (
            f'<tableStyleInfo name="{saxutils.escape(table.style_name, {"\"": "&quot;"})}" '
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

    def _write_body(self, data, spool, style_ranges, header_style, heights,
                    hidden_row_set, has_styles, table_headers=(),
                    row_outline=None):
        """Encode every row into ``spool`` and return a ``_SheetBody``.

        Only the first ``_WIDTH_SAMPLE_ROWS`` rows are kept, for automatic
        column widths, plus the header rows of ``table_headers``, which are
        checked as soon as they stream past.
        Hyperlink cells are recorded so the trailer can link them.
        """
        head: list = []
        hyperlinks: list = []
        header_rows: dict = {}
        num_rows = 0
        max_cols = 0
        buf = io.StringIO()
        row_style_ids_cache = {}
        ss_dict, ss_list = self._shared_strings_dict, self._shared_strings
        use_sst, fmt_xf = self._useSharedStrings, self._fmt_xf
        if table_headers:
            data = capture_table_headers(data, table_headers, header_rows)
        for row_idx, row in enumerate(data):
            if row_idx < _WIDTH_SAMPLE_ROWS:
                head.append(row)
            num_rows = row_idx + 1
            if len(row) > max_cols:
                max_cols = len(row)
            row_attrs = self._row_attrs(row_idx, heights, hidden_row_set, row_outline)
            row_styles = (
                styles_for_row(row_idx, len(row), style_ranges, header_style)
                if has_styles else _EMPTY_ROW_STYLES
            )
            # The C encoder handles common scalar rows with optional style ids
            # and <row> attributes; it hands back rows it cannot write as None
            # for the Python path.
            style_ids = (
                _style_ids_for_row(
                    row_styles, len(row), self._register_style,
                    row_style_ids_cache,
                )
                if row_styles else None
            )
            c_row = accel_encode_xlsx_row(
                row, ss_dict, ss_list,
                self._sst_unique_count, self._sst_all_count,
                row_idx, use_sst, style_ids, fmt_xf or None, row_attrs or None,
            )
            if c_row is not None:
                row_xml, self._sst_unique_count, self._sst_all_count = c_row
                buf.write(row_xml)
                if buf.tell() >= _FLUSH_CHARS:
                    spool.write(buf.getvalue().encode('utf-8'))
                    buf = io.StringIO()
                continue
            buf.write(f'<row r="{row_idx + 1}"{row_attrs}>')
            self._write_row_cells(buf, row, row_idx, row_styles, hyperlinks)
            buf.write('</row>')
            if buf.tell() >= _FLUSH_CHARS:
                spool.write(buf.getvalue().encode('utf-8'))
                buf = io.StringIO()
        if row_outline is not None:
            # Grouped rows and summary rows past the data are written empty.
            for row_idx in row_outline.indices_from(num_rows):
                row_attrs = self._row_attrs(row_idx, heights, hidden_row_set, row_outline)
                buf.write(f'<row r="{row_idx + 1}"{row_attrs}/>')
                if buf.tell() >= _FLUSH_CHARS:
                    spool.write(buf.getvalue().encode('utf-8'))
                    buf = io.StringIO()
        spool.write(buf.getvalue().encode('utf-8'))
        if num_rows == 0:
            max_cols = 1
        return _SheetBody(num_rows, max_cols, head, hyperlinks, header_rows=header_rows)

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

    def _write_row_cells(self, buf, row, row_idx, row_styles, hyperlinks):
        uniform_style = None
        uniform_xf = None
        if row_styles and len(row_styles) == len(row):
            first_style = next(iter(row_styles.values()))
            if all(style is first_style for style in row_styles.values()):
                uniform_style = first_style
                uniform_xf = self._register_style(first_style)

        for col_idx, cell in enumerate(row):
            cell_style = uniform_style if uniform_style is not None else row_styles.get(col_idx)
            if cell is None:
                if cell_style is None:
                    buf.write('<c/>')
                else:
                    xf_index = uniform_xf if cell_style is uniform_style else self._register_style(cell_style)
                    cell_ref = f'{self._letters[col_idx]}{row_idx + 1}'
                    buf.write(f'<c r="{cell_ref}" s="{xf_index}"/>')
                continue

            # (value, format) and (value, Style) cells: the overlay goes on
            # top of the range style.
            overlay = None
            if _is_formatted_cell(cell):
                cell, overlay = cell
                if isinstance(cell, Hyperlink):
                    if isinstance(overlay, str):
                        raise TypeError("a hyperlink cell cannot carry a number format; style the cell instead")
                    cell_style = merge_styles(cell_style, overlay)
                    overlay = None
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
                    # Only a registered style has the id the C encoder sees.
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

            if isinstance(cell, Formula):
                self._write_formula_cell(buf, cell, xf_index)
            elif isinstance(cell, str):
                self._write_string_cell(buf, cell, row_idx == 0, xf_index)
            elif isinstance(cell, bool):
                s_attr = f' s="{xf_index}"' if xf_index is not None else ''
                buf.write(f'<c t="b"{s_attr}><v>{"1" if cell else "0"}</v></c>')
            elif isinstance(cell, (int, float, Decimal)):
                if isinstance(cell, float):
                    if cell != cell:
                        self._write_string_cell(buf, "NaN", row_idx == 0, xf_index)
                        continue
                    elif cell == float('inf'):
                        self._write_string_cell(buf, "∞", row_idx == 0, xf_index)
                        continue
                    elif cell == float('-inf'):
                        self._write_string_cell(buf, "-∞", row_idx == 0, xf_index)
                        continue
                s_attr = f' s="{xf_index}"' if xf_index is not None else ''
                buf.write(f'<c{s_attr}><v>{float(cell)}</v></c>')
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
            self._tables_xml[table_id] = self._table_xml(table, table_id, names)
            rid = f"rId{position}"
            rels.append((rid, _REL_TABLE, f"../tables/table{table_id}.xml", False))
            sheet_tables.append((rid, table_id))
        body.tables = sheet_tables
        return rels, bounds

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
                    tab_key=None, zoom=None):
        """Everything before ``<sheetData>`` content: properties, views, widths."""
        sheet_id = sheet_index + 1
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
            sheet_index, split_freeze, freeze_panes, show_grid_lines, zoom,
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
                         show_grid_lines=True, zoom=None) -> str:
        tab = ' tabSelected="1"' if sheet_index == 0 else ""
        if not show_grid_lines:
            tab = ' showGridLines="0"' + tab
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

    def _write_empty_sheet(self, sheet_index: int) -> None:
        """A sheet that was added but never written holds no rows."""
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
        if not self._useSharedStrings:
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

    def _write_string_cell(self, buffer: io.StringIO, cell_value: str, is_header: bool = False, xf_index: int = None):
        check_cell_text(cell_value)
        if self._useSharedStrings:
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
            if cell_value and (cell_value[0] == ' ' or cell_value[0] == '\t'):
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t xml:space="preserve">{escaped_value}</t></is></c>')
            else:
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t>{escaped_value}</t></is></c>')
