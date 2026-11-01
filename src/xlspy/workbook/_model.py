"""Openpyxl-like workbook model.

This module deliberately contains no XLSX parsing or serialisation code.  The
fast path (``XlsxWriter``/``XlsbWriter``/``ExcelReader``) does not import it.
"""

from __future__ import annotations

from ..exceptions import UnsupportedFeatureError as UnsupportedFeatureError


from ..charts import Chart as Chart
from ..comments import Comment as Comment
from ..data_validation import DataValidation as DataValidation
from ..formats import DEFAULT_DATE_FORMAT as DEFAULT_DATE_FORMAT, DEFAULT_DATETIME_FORMAT as DEFAULT_DATETIME_FORMAT
from ..formulas import Formula as Formula
from ..hyperlinks import Hyperlink as Hyperlink
from ..images import Image as Image
from .._metadata import DefinedName as DefinedName
from ..refs import check_sheet_name as check_sheet_name, column_label as column_label, parse_cell_range as parse_cell_range, quote_sheet_name as quote_sheet_name
from ..rich_text import CellRichText as CellRichText
from ..styles import (
    Alignment as Alignment,
    Border as Border,
    CellIsRule as CellIsRule,
    Color as Color,
    ColorScaleRule as ColorScaleRule,
    DataBarRule as DataBarRule,
    EXTRA_CF_RULES as EXTRA_CF_RULES,
    ExcelTable as ExcelTable,
    Font as Font,
    FormulaRule as FormulaRule,
    GradientFill as GradientFill,
    IconSetRule as IconSetRule,
    NamedStyle as NamedStyle,
    NamedStyleList as NamedStyleList,
    PatternFill as PatternFill,
    Protection as Protection,
    Style as Style,
)
from . import _shift as _shift
from ._print import (
    HeaderFooter as HeaderFooter,
    PageMargins as PageMargins,
    PageSetup as PageSetup,
    PageSetupProperties as PageSetupProperties,
    PrintOptions as PrintOptions,
)
from ._styles import XlsxStyleTable as XlsxStyleTable
from ._coords import (
    _CELL_RE as _CELL_RE,
    _coerce_column as _coerce_column,
    _coerce_row as _coerce_row,
    _coordinate as _coordinate,
    _normalize_range as _normalize_range,
    _normalize_sqref as _normalize_sqref,
    CellRange as CellRange,
    MultiCellRange as MultiCellRange,
)
from ._properties import (
    WorkbookSecurity as WorkbookSecurity,
    CalcProperties as CalcProperties,
    BookView as BookView,
    CustomPropertyList as CustomPropertyList,
    DefinedNameList as DefinedNameList,
    DocumentProperties as DocumentProperties,
    Outline as Outline,
    SheetProperties as SheetProperties,
    SheetFormatProperties as SheetFormatProperties,
    Pane as Pane,
    Selection as Selection,
    SheetView as SheetView,
    FilterColumn as FilterColumn,
    SortCondition as SortCondition,
    SortState as SortState,
    AutoFilter as AutoFilter,
    Break as Break,
    RowBreak as RowBreak,
    ColBreak as ColBreak,
    SheetProtection as SheetProtection,
    _as_model_protection as _as_model_protection,
    _as_model_page_margins as _as_model_page_margins,
    CustomProperty as CustomProperty,
    DocumentSecurity as DocumentSecurity,
)
from ._tables import (
    _CF_RULES as _CF_RULES,
    _expect as _expect,
    ConditionalFormatting as ConditionalFormatting,
    ConditionalFormattingList as ConditionalFormattingList,
    TableColumn as TableColumn,
    TableStyleInfo as TableStyleInfo,
    Table as Table,
    TableList as TableList,
    DataValidationList as DataValidationList,
)
from ._dimensions import (
    Dimension as Dimension,
    ColumnDimension as ColumnDimension,
    RowDimension as RowDimension,
    DimensionHolder as DimensionHolder,
    _column_key as _column_key,
    _row_key as _row_key,
)
from ._cell import (
    _new_cell as _new_cell,
    _PLAIN_TYPES as _PLAIN_TYPES,
    Cell as Cell,
    CELL_SLOTS as CELL_SLOTS,
    MergedCell as MergedCell,
    WriteOnlyCell as WriteOnlyCell,
)
from ._worksheet import (
    Worksheet as Worksheet,
    WriteOnlyWorksheet as WriteOnlyWorksheet,
    Chartsheet as Chartsheet,
)
from ._workbook import (
    Workbook as Workbook,
)


