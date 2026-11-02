"""Exceptions compatible with ``openpyxl.utils.exceptions``.

The names match openpyxl; every class also derives from :class:`xlspy.XlspyError`.
"""

from __future__ import annotations

from ..exceptions import InvalidFileError, XlspyError


class CellCoordinatesException(XlspyError):
    """Raised when a cell coordinate cannot be parsed."""


class IllegalCharacterError(XlspyError):
    """Raised when text contains a character a cell cannot store."""


class InvalidFileException(InvalidFileError):
    """Raised when a workbook file is not a supported format."""


class NamedRangeException(XlspyError):
    """Raised for an invalid defined name or named range."""


class ReadOnlyWorkbookException(XlspyError):
    """Raised when editing a read-only workbook."""


class SheetTitleException(XlspyError):
    """Raised for an invalid worksheet title."""


class WorkbookAlreadySaved(XlspyError):
    """Raised when a write-only workbook is saved twice."""
