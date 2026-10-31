"""Exceptions compatible with ``openpyxl.utils.exceptions``."""

from __future__ import annotations

from ..exceptions import InvalidFileError


class CellCoordinatesException(Exception):
    """Raised when a cell coordinate cannot be parsed."""


class IllegalCharacterError(Exception):
    """Raised when text contains a character a cell cannot store."""


class InvalidFileException(InvalidFileError):
    """Raised when a workbook file is not a supported format."""


class NamedRangeException(Exception):
    """Raised for an invalid defined name or named range."""


class ReadOnlyWorkbookException(Exception):
    """Raised when editing a read-only workbook."""


class SheetTitleException(Exception):
    """Raised for an invalid worksheet title."""


class WorkbookAlreadySaved(Exception):
    """Raised when a write-only workbook is saved twice."""
