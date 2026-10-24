from .writer import XlsbWriter
from .xlsx_writer import XlsxWriter
from .xlsb_updater import XlsbUpdater
from .xlsx_updater import XlsmUpdater, XlsxUpdater
from .reader import ExcelReader
from .formats import F
from .styles import (
    Alignment,
    Border,
    CellIsRule,
    Color,
    ColorScaleRule,
    DataBarRule,
    DifferentialStyle,
    ExcelTable,
    Font,
    FormatObject,
    FormulaRule,
    GradientFill,
    IconSetRule,
    TopBottomRule,
    AboveAverageRule,
    TextRule,
    DuplicateRule,
    TimePeriodRule,
    BlanksRule,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
)
from .rich_text import CellRichText, InlineFont, TextBlock
from .charts import (
    Chart,
    ChartAxis,
    ChartLegend,
    ChartMarker,
    ChartSeries,
    DataLabels,
)
from .data_validation import DataValidation
from .formulas import Formula
from .comments import Comment
from ._metadata import DefinedName, SheetProtection, WorkbookProtection
from .hyperlinks import Hyperlink
from .images import Image
from .outline import Group
from .printing import PageMargins, PrintSettings
from .dataframe import dataframe_to_rows, to_xlsb, to_xlsx

__version__ = "0.10.0"

_WORKBOOK_EXPORTS = {
    "Workbook",
    "Worksheet",
    "WriteOnlyWorksheet",
    "ReadOnlyWorksheet",
    "Cell",
    "MergedCell",
    "CellRange",
    "MultiCellRange",
    "DocumentProperties",
    "ColumnDimension",
    "RowDimension",
    "Dimension",
    "DimensionHolder",
    "SheetProperties",
    "SheetView",
    "AutoFilter",
    "load_workbook",
}


def __getattr__(name):
    if name in _WORKBOOK_EXPORTS:
        from . import workbook as _workbook

        value = getattr(_workbook, name)
        globals()[name] = value
        return value
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    'XlsbWriter',
    'XlsxWriter',
    'XlsbUpdater',
    'XlsxUpdater',
    'XlsmUpdater',
    'ExcelReader',
    'F',
    'Style',
    'Font',
    'PatternFill',
    'GradientFill',
    'Stop',
    'Side',
    'Border',
    'Alignment',
    'Protection',
    'Color',
    'FormatObject',
    'DifferentialStyle',
    'CellIsRule',
    'FormulaRule',
    'IconSetRule',
    'ColorScaleRule',
    'DataBarRule',
    'TopBottomRule',
    'AboveAverageRule',
    'TextRule',
    'DuplicateRule',
    'TimePeriodRule',
    'BlanksRule',
    'ExcelTable',
    'Chart',
    'ChartSeries',
    'ChartAxis',
    'ChartLegend',
    'ChartMarker',
    'DataLabels',
    'Formula',
    'Comment',
    'DefinedName',
    'SheetProtection',
    'WorkbookProtection',
    'DataValidation',
    'Hyperlink',
    'Image',
    'CellRichText',
    'TextBlock',
    'InlineFont',
    'Group',
    'PageMargins',
    'PrintSettings',
    'dataframe_to_rows',
    'to_xlsx',
    'to_xlsb',
    'Workbook',
    'Worksheet',
    'WriteOnlyWorksheet',
    'ReadOnlyWorksheet',
    'Cell',
    'MergedCell',
    'CellRange',
    'MultiCellRange',
    'DocumentProperties',
    'ColumnDimension',
    'RowDimension',
    'Dimension',
    'DimensionHolder',
    'SheetProperties',
    'SheetView',
    'AutoFilter',
    'load_workbook',
]
