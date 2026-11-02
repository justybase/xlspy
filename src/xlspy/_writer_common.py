"""What XlsxWriter and XlsbWriter share before the format encoders:
argument checks and normalization, sheet bookkeeping and conditional
format priorities. The writers keep only what differs between the formats.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

from .charts import Chart
from .refs import check_sheet_name, parse_cell_range
from .styles import (
    EXTRA_CF_RULES, CellIsRule, ColorScaleRule, DataBarRule, ExcelTable, FormulaRule,
    IconSetRule, _column_index,
)

CF_RULES = (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule, *EXTRA_CF_RULES)


# Writer keywords renamed to snake_case; the old spelling gets a pointer, not just "unexpected".
RENAMED_WRITER_ARGUMENTS = {"compressionLevel": "compression_level", "useSharedStrings": "use_shared_strings"}


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


class WriterFrontEnd:
    """Methods both writers use unchanged; they read the writer's own state
    (``_worksheet_data``, ``_table_names``, the CF priority sets)."""

    _worksheet_data: list
    _sheet_count: int
    _table_names: set[str]
    _next_cf_priority: int
    _seen_cf_priorities: set[int]
    _reserved_cf_priorities: set[int]

    def add_sheet(self, sheet_name: str, hidden: bool = False) -> None:
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
