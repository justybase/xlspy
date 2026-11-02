"""What XlsxWriter and XlsbWriter share before the format encoders:
argument checks and normalization, sheet bookkeeping and conditional
format priorities. The writers keep only what differs between the formats.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from ._metadata import SheetProtection
from .charts import Chart
from .comments import normalize_comments
from .data_validation import normalize_validations
from .images import normalize_images
from .outline import build_outline
from .printing import PrintSettings, normalize_print_settings
from .shapes import normalize_shapes
from .slicers import normalize_slicers
from .sparklines import normalize_sparklines
from .refs import check_sheet_name, parse_cell_range
from .styles import (
    EXTRA_CF_RULES, CellIsRule, ColorScaleRule, DataBarRule, ExcelTable, FormulaRule,
    IconSetRule, Style, _column_index, check_header_rows, check_sheet_view, check_zoom,
    explicit_cf_priorities, header_style_ranges, normalize_hidden_columns, normalize_hidden_rows,
    normalize_style_ranges, tab_color_key,
)

CF_RULES = (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule, *EXTRA_CF_RULES)


# Writer keywords renamed to snake_case; the old spelling gets a pointer, not just "unexpected".
RENAMED_WRITER_ARGUMENTS = {"compressionLevel": "compression_level", "useSharedStrings": "use_shared_strings"}


def check_autofit(rows: Any, padding: Any) -> tuple[int | None, float]:
    """``autofit_rows`` (rows measured for automatic widths; None: all, 0: none)
    and ``autofit_padding`` (characters added to each measured column)."""
    if rows is not None and (isinstance(rows, bool) or not isinstance(rows, int)):
        raise TypeError("autofit_rows must be an int or None")
    if rows is not None and rows < 0:
        raise ValueError("autofit_rows must be >= 0 (or None for every row)")
    if isinstance(padding, bool) or not isinstance(padding, (int, float)):
        raise TypeError("autofit_padding must be a number of characters")
    if not 0 <= padding <= 255:
        raise ValueError("autofit_padding must be between 0 and 255 characters")
    return rows, float(padding)


def reject_unknown_arguments(owner: str, extra: Mapping[str, object]) -> None:
    """Raise TypeError for keywords ``owner`` does not take, naming the new
    spelling of a renamed one."""
    for name in extra:
        new = RENAMED_WRITER_ARGUMENTS.get(name)
        if new is not None:
            raise TypeError(f"{owner}() argument {name!r} was renamed to {new!r}")
        raise TypeError(f"{owner}() got an unexpected keyword argument {name!r}")


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0
    start_row: int = 0
    end_row: int = 0


@dataclass
class SheetSpec:
    """The options of one ``write_sheet`` call, checked and normalized; nothing of
    the writer's state has changed yet when this exists."""

    view: tuple[bool, bool, bool]
    header_rows: int
    tab_key: Any
    zoom: int | None
    style_ranges: list[Any]
    header_style: Style | None
    merge_refs: list[str]
    comments: dict[str, Any]
    cf_rules: list[Any]
    cf_priorities: list[int]
    tables: list[Any]
    sparklines: list[Any]
    slicers: list[Any]
    charts: list[Any]
    validations: list[Any]
    images: list[Any]
    shapes: list[Any]
    print_settings: PrintSettings | None
    widths: Any
    autofit: tuple[int | None, float]
    hidden_cols: set[int]
    heights: Any
    hidden_rows: set[int]
    row_outline: Any
    col_outline: Any
    split_freeze: tuple[int, int] | None


class WriterFrontEnd:
    """Methods both writers use unchanged; they read the writer's own state
    (``_worksheet_data``, ``_table_names``, the CF priority sets)."""

    _worksheet_data: list
    _sheet_count: int
    _table_names: set[str]
    _next_cf_priority: int
    _seen_cf_priorities: set[int]
    _reserved_cf_priorities: set[int]
    _slicer_names: set[str]
    _slicer_caches: set[str]
    _slicer_targets: set[tuple[str, str]]

    if TYPE_CHECKING:
        # Each writer encodes its own conditional format records.
        def _normalize_cf_rules(self, conditional_formats: Any) -> list[Any]: ...

    def add_sheet(self, sheet_name: str, hidden: bool = False) -> None:
        """Start a new sheet; the next ``write_sheet`` call fills it.

        ``hidden=True`` hides the sheet.
        """
        check_sheet_name(sheet_name, (name for name, _data, _hidden in self._worksheet_data))
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]), hidden))

    def _written_sheet_indexes(self) -> set[int]:
        """Indexes of the sheets whose data has been written (each writer tracks its own)."""
        raise NotImplementedError

    def _reject_second_write(self) -> None:
        """``write_sheet`` fills the most recently added sheet. A second call for that
        sheet would write its part twice, and a ZIP with duplicate names is corrupt."""
        if not self._worksheet_data:
            return
        index = self._sheet_count - 1
        if index in self._written_sheet_indexes():
            name = self._worksheet_data[index][0]
            raise ValueError(
                f"sheet {name!r} already has its data; call add_sheet() before write_sheet() again"
            )

    def _sheet_index_by_name(self, name: str) -> int | None:
        folded = name.casefold()
        for index, (sheet_name, _data, _hidden) in enumerate(self._worksheet_data):
            if sheet_name.casefold() == folded:
                return index
        return None

    def _check_sheet_options(
        self, *, header_style: Any, show_grid_lines: Any, header_rows: Any, tab_color: Any, zoom: Any,
        show_row_col_headers: Any, show_zero_values: Any, right_to_left: Any, styles: Any, merges: Any,
        comments: Any, sheet_protection: Any, auto_filter_ref: Any, conditional_formats: Any, tables: Any,
        sparklines: Any, slicers: Any, charts: Any, data_validations: Any, images: Any, shapes: Any,
        print_settings: Any, column_widths: Any, hidden_columns: Any, row_heights: Any, hidden_rows: Any,
        row_groups: Any, column_groups: Any, freeze_panes: Any,
        autofit_rows: Any = 1000, autofit_padding: Any = 0.0,
        check_style: Callable[[Style], None] | None = None,
    ) -> SheetSpec:
        """Check and normalize the options of ``write_sheet``; raise before anything
        is recorded, so a rejected call leaves the writer as it was.

        ``check_style`` sees the header style and every range style before they are
        combined (the XLSB writer refuses rich-text styles there)."""
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        if not isinstance(show_grid_lines, bool):
            raise TypeError("show_grid_lines must be a bool")
        header_rows = check_header_rows(header_rows)
        tab_key = tab_color_key(tab_color)
        zoom = check_zoom(zoom)
        view = check_sheet_view(show_row_col_headers, show_zero_values, right_to_left)
        style_ranges = normalize_style_ranges(styles)
        if check_style is not None:
            for candidate in ([header_style] if header_style is not None else []) + [item[-1] for item in style_ranges]:
                check_style(candidate)
        style_ranges, header_style = header_style_ranges(style_ranges, header_style, header_rows)
        merge_refs = self._normalize_merges(merges)
        comment_specs = normalize_comments(comments, merges=merge_refs)
        if sheet_protection is not None and not isinstance(sheet_protection, SheetProtection):
            raise TypeError("sheet_protection must be SheetProtection")
        if auto_filter_ref is not None:
            # Checked before the rows stream: a late error aborts the writer.
            if not isinstance(auto_filter_ref, str):
                raise TypeError("auto_filter_ref must be an A1 range string")
            parse_cell_range(auto_filter_ref)
        cf_rules = self._normalize_cf_rules(conditional_formats)
        table_specs = self._normalize_tables(tables)
        sparkline_groups = normalize_sparklines(sparklines)
        slicer_specs = normalize_slicers(
            slicers, table_specs, self._slicer_names, self._slicer_caches, self._slicer_targets,
        )
        chart_specs = self._normalize_charts(charts)
        validations = normalize_validations(data_validations)
        image_specs = normalize_images(images)
        shape_specs = normalize_shapes(shapes)
        print_spec = normalize_print_settings(print_settings)
        widths = self._normalize_widths(column_widths, "column_widths")
        autofit = check_autofit(autofit_rows, autofit_padding)
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
        # Explicit priorities are reserved up front, so automatic ones skip them.
        cf_priorities = explicit_cf_priorities(
            cf_rules, self._seen_cf_priorities, self._reserved_cf_priorities,
        )
        return SheetSpec(
            view=view, header_rows=header_rows, tab_key=tab_key, zoom=zoom,
            style_ranges=style_ranges, header_style=header_style, merge_refs=merge_refs,
            comments=comment_specs, cf_rules=cf_rules, cf_priorities=cf_priorities,
            tables=table_specs, sparklines=sparkline_groups, slicers=slicer_specs,
            charts=chart_specs, validations=validations, images=image_specs, shapes=shape_specs,
            print_settings=print_spec, widths=widths, autofit=autofit, hidden_cols=hidden_cols, heights=heights,
            hidden_rows=hidden_row_set, row_outline=row_outline, col_outline=col_outline,
            split_freeze=split_freeze,
        )

    def _claim_sheet_options(self, spec: SheetSpec) -> None:
        """Record the names and priorities a checked sheet takes."""
        self._table_names.update(table.display_name for table in spec.tables)
        self._slicer_names.update(slicer.display_name.casefold() for slicer in spec.slicers)
        self._slicer_caches.update(slicer.cache_name.casefold() for slicer in spec.slicers)
        self._slicer_targets.update(
            (slicer.table.casefold(), slicer.column.casefold()) for slicer in spec.slicers
        )
        for priority in spec.cf_priorities:
            self._reserved_cf_priorities.add(priority)
            self._next_cf_priority = max(self._next_cf_priority, priority + 1)

    def _check_unknown_cf_rule(self, rule) -> None:
        if hasattr(rule, "sqref") and not isinstance(rule, CF_RULES):
            raise TypeError(
                f"unsupported conditional formatting rule {type(rule).__name__}; "
                "use xlspy rule classes (CellIsRule, FormulaRule, "
                "ColorScaleRule, DataBarRule, IconSetRule, TopBottomRule, "
                "AboveAverageRule, TextRule, DuplicateRule, TimePeriodRule, BlanksRule)"
            )

    def _claim_cf_priority(self, rule) -> int:
        priority: int = rule.priority
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

    @staticmethod
    def _normalize_merges(merges):
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

    def _normalize_tables(self, tables):
        if not tables:
            return []
        # Names are only checked here; write_sheet registers them once every
        # option is valid, so a rejected call leaves no name behind.
        # Excel compares table names without regard to case: "Sales" and "sales"
        # in one workbook make Excel repair the file.
        names = {name.casefold() for name in self._table_names}
        for table in tables:
            if not isinstance(table, ExcelTable):
                raise TypeError("tables must be ExcelTable instances")
            if table.display_name.casefold() in names:
                raise ValueError(
                    f"duplicate table display name {table.display_name!r} (Excel ignores case)"
                )
            names.add(table.display_name.casefold())
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
