"""An openpyxl-like workbook model layered beside xlspy's fast paths.

``XlsxWriter``, ``XlsbWriter`` and ``ExcelReader`` remain unchanged and do not
import this package.  Use this API when you need random access, styles,
merges, metadata, or round-trip editing of an existing XLSX/XLSM file.
"""

from ._model import (
    AutoFilter,
    Cell,
    CellRange,
    ColumnDimension,
    Dimension,
    DimensionHolder,
    DocumentProperties,
    MergedCell,
    MultiCellRange,
    RowDimension,
    SheetProperties,
    SheetView,
    Workbook,
    Worksheet,
    WriteOnlyWorksheet,
)
from ._reader import ReadOnlyWorksheet, load_workbook
from ._styles import XlsxStyleTable

__all__ = [
    "AutoFilter",
    "Cell",
    "CellRange",
    "ColumnDimension",
    "Dimension",
    "DimensionHolder",
    "DocumentProperties",
    "MergedCell",
    "MultiCellRange",
    "ReadOnlyWorksheet",
    "RowDimension",
    "SheetProperties",
    "SheetView",
    "Workbook",
    "Worksheet",
    "WriteOnlyWorksheet",
    "XlsxStyleTable",
    "load_workbook",
]
