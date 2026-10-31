"""Exceptions raised by xlspy.

Every exception xlspy raises on purpose derives from :class:`XlspyError`, so
``except XlspyError`` catches all of them. The classes also keep the built-in
type that earlier versions raised, so existing ``except ValueError`` or
``except NotImplementedError`` clauses still match:

- :class:`InvalidFileError`: the file is not a readable Excel package. It is a
  ``ValueError`` and a ``zipfile.BadZipFile``.
- :class:`UnsupportedFeatureError`: the workbook uses, or the call asks for, a
  feature xlspy does not implement. It is a ``NotImplementedError``.

Other errors (bad arguments, out-of-range values) still use the built-in
types; they are not wrapped.
"""

from __future__ import annotations

import zipfile

__all__ = ["XlspyError", "InvalidFileError", "UnsupportedFeatureError"]


class XlspyError(Exception):
    """Base class for the errors xlspy raises deliberately."""


class InvalidFileError(XlspyError, ValueError, zipfile.BadZipFile):
    """The file is not a readable XLSX, XLSM or XLSB package.

    Raised when the file is not a ZIP archive, or when the native reader finds
    data it cannot decode. The message says what failed.
    """


class UnsupportedFeatureError(XlspyError, NotImplementedError):
    """The workbook or the call uses a feature xlspy does not support yet."""
