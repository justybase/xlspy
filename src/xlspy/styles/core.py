"""Cell formatting shared by the XLSX and XLSB writers.

The model mirrors openpyxl: ``Font``, ``PatternFill``/``GradientFill``,
``Border``/``Side``, ``Alignment``, ``Protection``, and the conditional rules
(``CellIsRule``, ``FormulaRule``, ``ColorScaleRule``, ``DataBarRule``,
``IconSetRule``) with ``DifferentialStyle`` as the Excel ``dxf``. The XLSX
writer serializes all of it; the XLSB writer currently supports the flat
``Style`` subset and rejects the rest with ``NotImplementedError``.
"""

from __future__ import annotations

# Chart and cell-reference models live in their own modules; they are
# re-exported here so that ``xlspy.styles`` keeps its historical names.
from ..charts import Chart as Chart, ChartSeries as ChartSeries  # re-exported: historical import path
from ._attrs import (
    _ARGB_RE as _ARGB_RE,
    _TABLE_NAME_RE as _TABLE_NAME_RE,
    _A1_NAME_RE as _A1_NAME_RE,
    _R1C1_NAME_RE as _R1C1_NAME_RE,
    _is_reference_like as _is_reference_like,
    _UNDERLINES as _UNDERLINES,
    _VERT_ALIGNS as _VERT_ALIGNS,
    _PATTERN_TYPES as _PATTERN_TYPES,
    _SIDE_STYLES as _SIDE_STYLES,
    _DIAGONAL_DIRECTIONS as _DIAGONAL_DIRECTIONS,
    _HORIZONTAL as _HORIZONTAL,
    _VERTICAL as _VERTICAL,
    _FONT_SCHEMES as _FONT_SCHEMES,
    _GRADIENT_TYPES as _GRADIENT_TYPES,
    _ICON_SETS as _ICON_SETS,
    _CELL_IS_OPERATORS as _CELL_IS_OPERATORS,
    _FORMAT_OBJECT_TYPES as _FORMAT_OBJECT_TYPES,
    _CFVO_TYPES as _CFVO_TYPES,
    _normalize_argb as _normalize_argb,
    _coerce_color as _coerce_color,
    Color as Color,
    Font as Font,
    PatternFill as PatternFill,
    Stop as Stop,
    GradientFill as GradientFill,
    Side as Side,
    Border as Border,
    Alignment as Alignment,
    Protection as Protection,
    normalize_color as normalize_color,
)
from ._style import (
    Style as Style,
    DifferentialStyle as DifferentialStyle,
)
from ._cf_rules import (
    _normalize_formula as _normalize_formula,
    _validate_priority as _validate_priority,
    _validate_sqref as _validate_sqref,
    Rule as Rule,
    CellIsRule as CellIsRule,
    FormulaRule as FormulaRule,
    FormatObject as FormatObject,
    ColorScaleRule as ColorScaleRule,
    DataBarRule as DataBarRule,
    IconSetRule as IconSetRule,
    _check_dxf_rule as _check_dxf_rule,
    _check_flag as _check_flag,
    TopBottomRule as TopBottomRule,
    AboveAverageRule as AboveAverageRule,
    TEXT_OPERATORS as TEXT_OPERATORS,
    TextRule as TextRule,
    DuplicateRule as DuplicateRule,
    TIME_PERIODS as TIME_PERIODS,
    TimePeriodRule as TimePeriodRule,
    BLANK_KINDS as BLANK_KINDS,
    BlanksRule as BlanksRule,
    CF_TEXT_RULES as CF_TEXT_RULES,
    CF_TIME_PERIOD_FORMULAS as CF_TIME_PERIOD_FORMULAS,
    CF_BLANK_RULES as CF_BLANK_RULES,
    EXTRA_CF_RULES as EXTRA_CF_RULES,
    _text_rule_type as _text_rule_type,
    _duplicate_rule_type as _duplicate_rule_type,
    _blanks_rule_type as _blanks_rule_type,
    _RULE_TYPES as _RULE_TYPES,
)
from ._tables import (
    ExcelTable as ExcelTable,
    table_column_names as table_column_names,
    capture_table_headers as capture_table_headers,
)
from ._sheet_checks import (
    explicit_cf_priorities as explicit_cf_priorities,
    normalize_hidden_columns as normalize_hidden_columns,
    normalize_hidden_rows as normalize_hidden_rows,
)
from ._keys import (
    _FontKey as _FontKey,
    _AlignmentKey as _AlignmentKey,
    _BorderKey as _BorderKey,
    _ProtectionKey as _ProtectionKey,
    _underline_key as _underline_key,
    _color_key as _color_key,
    _side_key as _side_key,
    _EMPTY_SIDE as _EMPTY_SIDE,
    _font_key_from as _font_key_from,
    _font_is_empty as _font_is_empty,
    _border_key as _border_key,
    _pattern_fill_key as _pattern_fill_key,
    _fill_key_from as _fill_key_from,
    _alignment_key_from as _alignment_key_from,
    _protection_key_from as _protection_key_from,
    fill_key_for as fill_key_for,
    split_style as split_style,
)
from ._ranges import (
    _EMPTY_STYLES as _EMPTY_STYLES,
    StyleRanges as StyleRanges,
    _LAST_ROW as _LAST_ROW,
    _LAST_COL as _LAST_COL,
    _OPEN_COLUMNS_RE as _OPEN_COLUMNS_RE,
    _OPEN_ROWS_RE as _OPEN_ROWS_RE,
    parse_style_range as parse_style_range,
    normalize_style_ranges as normalize_style_ranges,
    check_header_rows as check_header_rows,
    header_style_ranges as header_style_ranges,
    check_zoom as check_zoom,
    check_sheet_view as check_sheet_view,
    tab_color_key as tab_color_key,
    merge_styles as merge_styles,
    _uniform_row_styles as _uniform_row_styles,
    _row_styles_for_pattern as _row_styles_for_pattern,
    styles_for_row as styles_for_row,
)
from ._cell_xf import (
    _xf_for_style_format as _xf_for_style_format,
    _record_cell_xf as _record_cell_xf,
    _date_cell_xf as _date_cell_xf,
    _style_ids_for_row as _style_ids_for_row,
    HYPERLINK_STYLE as HYPERLINK_STYLE,
)
from ..hyperlinks import Hyperlink as Hyperlink
from ..refs import _column_index as _column_index
from ..refs import column_label as column_label
from ..refs import parse_cell_range as parse_cell_range


# ---------------------------------------------------------------------------
# Structured cell-style model (openpyxl-like, XLSB-safe pure dataclasses)
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Sheet-feature models: conditional formatting, tables
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Registry-key splitting (used only by the XLSX writer)
# ---------------------------------------------------------------------------


