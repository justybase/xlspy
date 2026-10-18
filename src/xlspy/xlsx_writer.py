from dataclasses import dataclass
import datetime
from decimal import Decimal
import zipfile
import xml.sax.saxutils as saxutils
import io
from typing import Iterable, Mapping, Tuple
import itertools
from . import chart_parts
from .formats import _is_formatted_cell, _unwrap_cell, _get_format
from .excel_dates import datetime_to_serial
from .styles import (
    CellIsRule,
    Chart,
    ColorScaleRule,
    DataBarRule,
    DifferentialStyle,
    ExcelTable,
    FormatObject,
    FormulaRule,
    IconSetRule,
    Style,
    _alignment_key_from,
    _border_key,
    _column_index,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _protection_key_from,
    merge_styles,
    normalize_style_ranges,
    parse_cell_range,
    split_style,
    styles_for_row,
)


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


class XlsxWriter:
    def __init__(self, file: str | io.IOBase, compressionLevel: int = 4, useSharedStrings: bool = True):
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
        self._zf: zipfile.ZipFile = None
        self._letters = self._generate_column_letters()

        self._format_registry: dict[str, int] = {}
        self._format_xf_map: dict[str, int] = {}
        self._style_xf_map: dict[Style, int] = {}
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
        self._sheet_tables: dict[int, list[tuple]] = {}
        self._tables_xml: dict[int, str] = {}
        self._sheet_charts: dict[int, list] = {}
        self._sheet_chart_ids: dict[int, list[int]] = {}
        self._charts_xml: dict[int, tuple[int, int, str]] = {}
        self._next_chart_id = 1
        self._next_numfmt_id = 164
        self._next_xf_index = 4

    def __enter__(self):
        if self._zf is None:
            self._zf = zipfile.ZipFile(self._file, 'w', zipfile.ZIP_DEFLATED, compresslevel=self._compressionLevel)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self._zf is not None:
                self.save()
        finally:
            if self._zf is not None:
                self._zf.close()
                self._zf = None
        
        return False

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
            (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule),
        ):
            raise TypeError(
                f"unsupported conditional formatting rule {type(rule).__name__}; "
                "use xlspy rule classes (CellIsRule, FormulaRule, "
                "ColorScaleRule, DataBarRule, IconSetRule)"
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
                 IconSetRule),
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
        for table in tables:
            if not isinstance(table, ExcelTable):
                raise TypeError("tables must be ExcelTable instances")
            if table.display_name in self._table_names:
                raise ValueError(
                    f"duplicate table display name {table.display_name!r}"
                )
            self._table_names.add(table.display_name)
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
    def _normalize_hidden(items, kind):
        if items is None:
            return set()
        if isinstance(items, (str, bytes)):
            raise TypeError(f"{kind} must be a collection of columns/rows, not a string")
        try:
            values = list(items)
        except TypeError:
            raise TypeError(f"{kind} must be a collection or None")
        result = set()
        for key in values:
            if isinstance(key, str):
                col = _column_index(key)
            elif isinstance(key, int) and not isinstance(key, bool):
                col = key
            else:
                raise TypeError(f"{kind} keys must be column letters or 0-based ints")
            result.add(col)
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
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        charts: list[Chart] | tuple[Chart, ...] | None = None,
    ):
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        style_ranges = normalize_style_ranges(styles)
        merge_refs = self._normalize_merges(merges)
        cf_rules = self._normalize_cf_rules(conditional_formats)
        table_specs = self._normalize_tables(tables)
        chart_specs = self._normalize_charts(charts)
        if not self._worksheet_data:
            self.add_sheet("Sheet1")

        sheet_name, _, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)
        if chart_specs:
            self._sheet_charts[self._sheet_count] = chart_specs

        sheet_id = self._sheet_count
        drawing_rid = len(table_specs) + 1 if chart_specs else None
        with self._zf.open(f"xl/worksheets/sheet{sheet_id}.xml", 'w') as sheet_file:
            self._write_worksheet_xml(
                sheet_file, data, self._sheet_count - 1,
                style_ranges=style_ranges, header_style=header_style,
                merges=merge_refs, conditional_formats=cf_rules,
                tables=table_specs, column_widths=column_widths,
                hidden_columns=hidden_columns, row_heights=row_heights,
                hidden_rows=hidden_rows, freeze_panes=freeze_panes,
                auto_filter=auto_filter, drawing_rid=drawing_rid,
            )
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
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        charts: list | tuple | None = None,
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
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            charts=charts,
        )

    def save(self):
        if self._zf is None:
            raise RuntimeError("Zip file not initialized. Use context manager or initialize manually.")

        for sheet_id, charts in self._sheet_charts.items():
            if not charts or sheet_id in self._sheet_chart_ids:
                continue
            sheet_name = self._worksheet_data[sheet_id - 1][0]
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
                self._drawing_xml(charts, chart_ids),
            )
            self._zf.writestr(
                f"xl/drawings/_rels/drawing{sheet_id}.xml.rels",
                self._drawing_rels_xml(chart_ids),
            )
            if sheet_id not in self._sheet_tables:
                self._zf.writestr(
                    f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels",
                    self._create_worksheet_rels(sheet_id),
                )

        if any(self._sheet_charts.values()):
            self._zf.writestr("xl/theme/theme1.xml", chart_parts.THEME_XML)

        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.xml", self._create_workbook_xml())
        self._zf.writestr("xl/styles.xml", self._create_styles_xml())
        self._zf.writestr("xl/_rels/workbook.xml.rels", self._create_workbook_rels())

        for sheet_id, tables in self._sheet_tables.items():
            if tables:
                self._zf.writestr(
                    f"xl/worksheets/_rels/sheet{sheet_id}.xml.rels",
                    self._create_worksheet_rels(sheet_id),
                )
        for table_id, table_xml in self._tables_xml.items():
            self._zf.writestr(f"xl/tables/table{table_id}.xml", table_xml)

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
        tables = self._sheet_tables.get(sheet_id, [])
        parts = "".join(
            f'<Relationship Id="rId{i + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
            f'Target="../tables/table{table_id}.xml"/>'
            for i, (table_id, _ref) in enumerate(tables)
        )
        if sheet_id in self._sheet_charts:
            parts += (
                f'<Relationship Id="rId{len(tables) + 1}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/drawing{sheet_id}.xml"/>'
            )
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{parts}
</Relationships>'''
    def close(self):
        if self._zf is not None:
            self._zf.close()
            self._zf = None

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
            for sheet_id in sorted(self._sheet_charts)
            if self._sheet_charts[sheet_id]
        )
        chart_overrides = "".join(
            f'<Override PartName="/xl/charts/chart{chart_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            for chart_id in range(1, self._next_chart_id)
        )
        theme_part = ""
        if any(self._sheet_charts.values()):
            theme_part = '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'

        shared_strings_part = ""
        if self._useSharedStrings and self._shared_strings:
            shared_strings_part = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{parts}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{table_parts}
{drawing_parts}
{chart_overrides}
{theme_part}
{shared_strings_part}
</Types>'''

    def _create_root_rels(self) -> str:
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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

        if any(self._sheet_charts.values()):
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

        defined_names = ""
        if self._filtered_data_list:
            defined_names = "<definedNames>"
            for filter_data in self._filtered_data_list:
                sheet_name = self._worksheet_data[filter_data.sheet_index][0]
                start_col = self._letters[filter_data.start_column]
                end_col = self._letters[filter_data.end_column]
                range_ref = f"{sheet_name}!${start_col}${filter_data.start_row + 1}:${end_col}${filter_data.end_row + 1}"
                defined_names += f'<definedName name="_xlnm._FilterDatabase" localSheetId="{filter_data.sheet_index}" hidden="1">{range_ref}</definedName>'
            defined_names += "</definedNames>"

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
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
                bg_xml = ""
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
        parts = [
            self._border_side_xml("start", border_key.start, self._color_attrs),
            self._border_side_xml("end", border_key.end, self._color_attrs),
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
        xf_entries.append('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        xf_entries.append('<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
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
        content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{self._sst_all_count}" uniqueCount="{self._sst_unique_count}">'''
        
        for s in self._shared_strings:
            escaped_string = saxutils.escape(s)
            if s and (s[0] == ' ' or s[0] == '\t'):
                content += f'<si><t xml:space="preserve">{escaped_string}</t></si>'
            else:
                content += f'<si><t>{escaped_string}</t></si>'
        
        content += '</sst>'
        sst_file.write(content.encode('utf-8'))

    def _calculate_column_widths(self, data_iterator, max_cols: int) -> list[float]:
        col_widths = [8.43] * max_cols
        
        rows_analyzed = 0
        for row in data_iterator:
            if rows_analyzed >= 100:
                break
            
            for col_idx, cell in enumerate(row):
                if col_idx >= max_cols:
                    break
                    
                cell_val = _unwrap_cell(cell) if cell is not None else None
                if cell_val is not None:
                    if isinstance(cell_val, datetime.datetime):
                        width = 18.0
                    elif isinstance(cell_val, datetime.date):
                        width = 10.14
                    else:
                        width = max(8.43, len(str(cell_val)) * 1.25 + 2)
                    
                    col_widths[col_idx] = max(col_widths[col_idx], min(width, 255))
            
            rows_analyzed += 1
        
        return col_widths

    @staticmethod
    def _format_object_xml(cfvo) -> str:
        attrs = f'type="{cfvo.type}"'
        if cfvo.value is not None:
            attrs += f' val="{saxutils.escape(str(cfvo.value), {"\"": "&quot;"})}"'
        return f"<cfvo {attrs}/>"

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

    def _table_xml(self, table: ExcelTable, table_id: int) -> str:
        from .styles import _column_index as col_index
        first_col, first_row, last_col, last_row = parse_cell_range(table.ref)
        ref = (
            f"{self._letters[first_col]}{first_row + 1}:"
            f"{self._letters[last_col]}{last_row + 1}"
        )
        name = saxutils.escape(table.display_name, {'"': '&quot;'})
        if table.columns is not None:
            names = list(table.columns)
        else:
            width = last_col - first_col + 1
            names = [f"Column{i + 1}" for i in range(width)]
        columns_xml = "".join(
            f'<tableColumn id="{i + 1}" name="{saxutils.escape(nm, {"\"": "&quot;"})}"/>'
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

    def _write_worksheet_xml(
        self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[any]],
        worksheet_index: int, *, style_ranges=(), header_style: Style | None = None,
        merges=(), conditional_formats=(), tables=(),
        column_widths=None, hidden_columns=None,
        row_heights=None, hidden_rows=None,
        freeze_panes: str | None = None, auto_filter=None,
        drawing_rid: int | None = None,
    ):
        rich_sheet = (
            header_style is not None or bool(style_ranges) or bool(merges)
            or bool(conditional_formats) or bool(tables)
            or column_widths is not None or hidden_columns is not None
            or row_heights is not None or hidden_rows is not None
            or freeze_panes is not None or auto_filter is not None
        )
        widths = self._normalize_widths(column_widths, "column_widths")
        hidden_cols = self._normalize_hidden(hidden_columns or (), "hidden_columns")
        if hidden_columns is not None:
            hidden_cols = self._normalize_hidden(hidden_columns, "hidden_columns")
        heights = self._normalize_row_heights(row_heights)
        hidden_row_set: set[int] = set()
        if hidden_rows is not None:
            try:
                candidates = list(hidden_rows)
            except TypeError:
                raise TypeError("hidden_rows must be a collection or None")
            for key in candidates:
                if not isinstance(key, int) or isinstance(key, bool) or key < 0:
                    raise TypeError("hidden_rows entries must be 0-based row ints")
                hidden_row_set.add(key)
        if freeze_panes is not None and freeze_panes != "":
            self._split_freeze_cell(freeze_panes)
        for rule in conditional_formats:
            if getattr(rule, "priority", None) is not None:
                if (
                    rule.priority in self._seen_cf_priorities
                    or rule.priority in self._reserved_cf_priorities
                ):
                    raise ValueError(
                        f"duplicate conditional formatting priority {rule.priority}"
                    )
                self._reserved_cf_priorities.add(rule.priority)
                self._next_cf_priority = max(
                    self._next_cf_priority, rule.priority + 1
                )
        table_bounds = [parse_cell_range(table.ref) for table in tables]
        buffer = io.StringIO()

        buffer.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        buffer.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

        data_list = list(worksheet_data)
        max_cols = max(len(row) for row in data_list) if data_list else 1
        num_rows = len(data_list)

        if rich_sheet:
            self._write_rich_sheet(
                buffer, sheet_file, data_list, worksheet_index, max_cols, num_rows,
                style_ranges, header_style, merges, conditional_formats, tables,
                table_bounds, widths, hidden_cols, heights, hidden_row_set,
                freeze_panes, auto_filter, drawing_rid,
            )
            return

        if num_rows > 0 and max_cols > 0:
            end_cell = f"{self._letters[max_cols - 1]}{num_rows}"
            buffer.write(f'<dimension ref="A1:{end_cell}"/>')

        if worksheet_index == 0:
            buffer.write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        else:
            buffer.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')

        col_widths = self._calculate_column_widths(iter(data_list), max_cols)

        if any(width != 8.43 for width in col_widths):
            buffer.write('<cols>')
            for i, width in enumerate(col_widths):
                if width != 8.43:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{width:.2f}" bestFit="1" customWidth="1"/>')
            buffer.write('</cols>')

        buffer.write('<sheetData>')

        last_row_idx = 0
        for row_idx, row in enumerate(data_list):
            last_row_idx = row_idx
            buffer.write(f'<row r="{row_idx + 1}">')

            for col_idx, cell in enumerate(row):
                if cell is None:
                    buffer.write('<c/>')
                    continue

                fmt_string = _get_format(cell)
                if fmt_string is not None:
                    cell = _unwrap_cell(cell)
                    xf_index = self._register_format(fmt_string)
                    has_custom_fmt = True
                else:
                    has_custom_fmt = False
                    xf_index = None

                if isinstance(cell, str):
                    if has_custom_fmt:
                        self._write_string_cell(buffer, cell, row_idx == 0, xf_index)
                    else:
                        self._write_string_cell(buffer, cell, row_idx == 0)
                elif isinstance(cell, bool):
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c t="b"{s_attr}><v>{"1" if cell else "0"}</v></c>')
                elif isinstance(cell, (int, float, Decimal)):
                    if isinstance(cell, float):
                        if cell != cell:
                            self._write_string_cell(buffer, "NaN", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                        elif cell == float('inf'):
                            self._write_string_cell(buffer, "\u221e", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                        elif cell == float('-inf'):
                            self._write_string_cell(buffer, "-\u221e", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c{s_attr}><v>{float(cell)}</v></c>')
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0, xf_index if has_custom_fmt else None)
                    else:
                        if cell.tzinfo is not None:
                            cell = cell.replace(tzinfo=None)
                        excel_date = datetime_to_serial(cell)
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="2"><v>{excel_date}</v></c>')
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0, xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="1"><v>{excel_date}</v></c>')
                else:
                    self._write_string_cell(buffer, str(cell), row_idx == 0,
                                            xf_index if has_custom_fmt else None)

            buffer.write('</row>')

        buffer.write('</sheetData>')

        if last_row_idx > 0:
            start_col = self._letters[0]
            end_col = self._letters[max_cols - 1]
            buffer.write(f'<autoFilter ref="{start_col}1:{end_col}{last_row_idx + 1}"/>')

            self._filtered_data_list.append(FilterData(
                sheet_index=worksheet_index,
                start_column=0,
                end_column=max_cols - 1,
                start_row=0,
                end_row=last_row_idx
            ))

        buffer.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        if drawing_rid is not None:
            buffer.write(f'<drawing r:id="rId{drawing_rid}"/>')
        buffer.write('</worksheet>')

        sheet_file.write(buffer.getvalue().encode('utf-8'))

    def _write_rich_sheet(
        self, buffer, sheet_file, data_list, worksheet_index, max_cols, num_rows,
        style_ranges, header_style, merges, conditional_formats, tables,
        table_bounds, widths, hidden_cols, heights, hidden_row_set,
        freeze_panes, auto_filter, drawing_rid=None,
    ):
        if num_rows > 0 and max_cols > 0:
            end_cell = f"{self._letters[max_cols - 1]}{num_rows}"
            buffer.write(f'<dimension ref="A1:{end_cell}"/>')

        if freeze_panes is None:
            if worksheet_index == 0:
                buffer.write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
            else:
                buffer.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        elif freeze_panes == "":
            if worksheet_index == 0:
                buffer.write('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>')
            else:
                buffer.write('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
        else:
            split_col, split_row = self._split_freeze_cell(freeze_panes)
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
            tab = " tabSelected=\"1\"" if worksheet_index == 0 else ""
            buffer.write(
                f'<sheetViews><sheetView{tab} workbookViewId="0">'
                f'<pane {" ".join(parts)}/>'
                f'<selection pane="{active}" activeCell="{top_left}" sqref="{top_left}"/>'
                "</sheetView></sheetViews>"
            )

        col_widths = self._calculate_column_widths(iter(data_list), max_cols)
        for col, width in widths.items():
            if 0 <= col < max_cols:
                col_widths[col] = width
        covered = set(widths) | set(hidden_cols)
        if any(width != 8.43 for width in col_widths) or covered:
            buffer.write('<cols>')
            emitted = set()
            for i, width in enumerate(col_widths):
                explicit = widths.get(i)
                hidden_attr = ' hidden="1"' if i in hidden_cols else ""
                if explicit is not None:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{explicit:.2f}" bestFit="1" customWidth="1"{hidden_attr}/>')
                    emitted.add(i)
                elif width != 8.43:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{width:.2f}" bestFit="1" customWidth="1"{hidden_attr}/>')
                    emitted.add(i)
                elif hidden_attr:
                    buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="8.43" customWidth="0"{hidden_attr}/>')
                    emitted.add(i)
            for i in sorted(covered):
                if i in emitted or i >= max_cols:
                    continue
                explicit = widths.get(i)
                hidden_attr = ' hidden="1"' if i in hidden_cols else ""
                width_attr = f"{explicit:.2f}" if explicit is not None else "8.43"
                custom = "1" if explicit is not None else "0"
                buffer.write(f'<col min="{i + 1}" max="{i + 1}" width="{width_attr}" customWidth="{custom}"{hidden_attr}/>')
            for i in sorted(hidden_cols):
                if i >= 16384:
                    raise ValueError(f"hidden column out of range: {i}")
            buffer.write('</cols>')

        buffer.write('<sheetData>')

        last_row_idx = 0
        for row_idx, row in enumerate(data_list):
            last_row_idx = row_idx
            row_attrs = f' r="{row_idx + 1}"'
            if row_idx in heights:
                row_attrs += f' ht="{heights[row_idx]}" customHeight="1"'
            if row_idx in hidden_row_set:
                row_attrs += ' hidden="1"'
            buffer.write(f"<row{row_attrs}>")
            row_styles = styles_for_row(row_idx, len(row), style_ranges, header_style)
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
                        buffer.write('<c/>')
                    else:
                        xf_index = uniform_xf if cell_style is uniform_style else self._register_style(cell_style)
                        cell_ref = f'{self._letters[col_idx]}{row_idx + 1}'
                        buffer.write(f'<c r="{cell_ref}" s="{xf_index}"/>')
                    continue

                fmt_string = _get_format(cell)
                if fmt_string is not None:
                    cell = _unwrap_cell(cell)
                    if cell_style is not None:
                        cell_style = merge_styles(cell_style, Style(number_format=fmt_string))
                has_custom_fmt = cell_style is not None or fmt_string is not None
                if cell_style is not None:
                    if cell_style is uniform_style and fmt_string is None:
                        xf_index = uniform_xf
                    else:
                        xf_index = self._register_style(cell_style)
                elif fmt_string is not None:
                    xf_index = self._register_format(fmt_string)

                if isinstance(cell, str):
                    if has_custom_fmt:
                        self._write_string_cell(buffer, cell, row_idx == 0, xf_index)
                    else:
                        self._write_string_cell(buffer, cell, row_idx == 0)
                elif isinstance(cell, bool):
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c t="b"{s_attr}><v>{"1" if cell else "0"}</v></c>')
                elif isinstance(cell, (int, float, Decimal)):
                    if isinstance(cell, float):
                        if cell != cell:
                            self._write_string_cell(buffer, "NaN", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                        elif cell == float('inf'):
                            self._write_string_cell(buffer, "\u221e", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                        elif cell == float('-inf'):
                            self._write_string_cell(buffer, "-\u221e", row_idx == 0, xf_index if has_custom_fmt else None)
                            continue
                    s_attr = f' s="{xf_index}"' if has_custom_fmt else ''
                    buffer.write(f'<c{s_attr}><v>{float(cell)}</v></c>')
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0, xf_index if has_custom_fmt else None)
                    else:
                        if cell.tzinfo is not None:
                            cell = cell.replace(tzinfo=None)
                        excel_date = datetime_to_serial(cell)
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="2"><v>{excel_date}</v></c>')
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_string_cell(buffer, str(cell), row_idx == 0, xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        if has_custom_fmt:
                            buffer.write(f'<c s="{xf_index}"><v>{excel_date}</v></c>')
                        else:
                            buffer.write(f'<c s="1"><v>{excel_date}</v></c>')
                else:
                    self._write_string_cell(buffer, str(cell), row_idx == 0,
                                            xf_index if has_custom_fmt else None)

            buffer.write('</row>')

        buffer.write('</sheetData>')

        data_filter = None
        if num_rows > 1 and max_cols > 0:
            data_filter = (
                f"{self._letters[0]}1:{self._letters[max_cols - 1]}{num_rows}"
            )
        table_refs = {
            (
                f"{self._letters[c1]}{r1 + 1}:"
                f"{self._letters[c2]}{r2 + 1}"
            )
            for c1, r1, c2, r2 in table_bounds
        }
        if data_filter is not None and auto_filter is not False:
            if data_filter not in table_refs:
                buffer.write(f'<autoFilter ref="{data_filter}"/>')
                self._filtered_data_list.append(FilterData(
                    sheet_index=worksheet_index,
                    start_column=0,
                    end_column=max_cols - 1,
                    start_row=0,
                    end_row=num_rows - 1
                ))
        if merges:
            buffer.write(f'<mergeCells count="{len(merges)}">')
            for ref in merges:
                buffer.write(f'<mergeCell ref="{ref}"/>')
            buffer.write('</mergeCells>')
        if conditional_formats:
            by_range: dict[str, list] = {}
            for rule in conditional_formats:
                by_range.setdefault(rule.sqref, []).append(rule)
            for sqref, rules in by_range.items():
                buffer.write(f'<conditionalFormatting sqref="{sqref}">')
                for rule in rules:
                    buffer.write(self._cf_rule_xml(rule, set(table_bounds)))
                buffer.write('</conditionalFormatting>')
        if tables:
            sheet_tables = []
            for table, bounds in zip(tables, table_bounds):
                table_id = self._next_table_id
                self._next_table_id += 1
                c1, r1, c2, r2 = bounds
                ref = (
                    f"{self._letters[c1]}{r1 + 1}:"
                    f"{self._letters[c2]}{r2 + 1}"
                )
                self._tables_xml[table_id] = self._table_xml(table, table_id)
                sheet_tables.append((table_id, ref))
            self._sheet_tables[worksheet_index + 1] = sheet_tables

        buffer.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        if drawing_rid is not None:
            buffer.write(f'<drawing r:id="rId{drawing_rid}"/>')
        if tables:
            buffer.write(f'<tableParts count="{len(sheet_tables)}">')
            for position, (table_id, _ref) in enumerate(sheet_tables, start=1):
                buffer.write(f'<tablePart r:id="rId{position}"/>')
            buffer.write('</tableParts>')
        buffer.write('</worksheet>')

        sheet_file.write(buffer.getvalue().encode('utf-8'))

    def _write_string_cell(self, buffer: io.StringIO, cell_value: str, is_header: bool = False, xf_index: int = None):
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
            escaped_value = saxutils.escape(cell_value)
            if xf_index is not None:
                style_ref = f' s="{xf_index}"'
            else:
                style_ref = ' s="3"' if is_header else ''
            if cell_value and (cell_value[0] == ' ' or cell_value[0] == '\t'):
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t xml:space="preserve">{escaped_value}</t></is></c>')
            else:
                buffer.write(f'<c t="inlineStr"{style_ref}><is><t>{escaped_value}</t></is></c>')
