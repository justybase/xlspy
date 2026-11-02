"""Exceptions raised by xlspy.

Every exception xlspy raises on purpose derives from ``XlspyError``, so
``except XlspyError`` catches all of them. The classes also keep the built-in
type that earlier versions raised, so existing ``except ValueError`` or
``except NotImplementedError`` clauses still match:

- ``InvalidFileError``: the file is not a readable Excel package. It is a
  ``ValueError`` and a ``zipfile.BadZipFile``.
- ``UnsupportedFeatureError``: the workbook uses, or the call asks for, a
  feature xlspy does not implement. It is a ``NotImplementedError``.

Other errors (bad arguments, out-of-range values) still use the built-in
types; they are not wrapped.
"""

from __future__ import annotations

import os
import struct
import zipfile
import zlib
from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from typing import TypeVar
from xml.etree.ElementTree import ParseError
from xml.parsers.expat import ExpatError

__all__ = ["XlspyError", "InvalidFileError", "UnsupportedFeatureError"]


class XlspyError(Exception):
    """Base class for the errors xlspy raises deliberately."""


class InvalidFileError(XlspyError, ValueError, zipfile.BadZipFile):
    """The file is not a readable XLSX, XLSM or XLSB package.

    Raised when the file is not a ZIP archive, or when a part holds data the
    readers cannot decode (malformed XML, bad text encoding, truncated or
    inconsistent records). The message names the file and what failed; the
    decoder's own error is chained as ``__cause__``.
    """


class UnsupportedFeatureError(XlspyError, NotImplementedError):
    """The workbook or the call uses a feature xlspy does not support yet."""


# What the XML parser, the text decoders and the record readers raise on bytes
# they cannot read. ValueError covers int()/float() of damaged attribute text and
# the model refusing a damaged value: the callers check their own arguments
# before the block starts, so inside it a ValueError comes from the file.
# KeyError is left out: a missing sheet name is the caller's.
_DECODE_ERRORS = (ParseError, ExpatError, ValueError, struct.error, IndexError, zlib.error, EOFError)

_T = TypeVar("_T")

# The native module's PanicException, registered by xlspy._accel when it loads
# the module. A panic is a BaseException; in a reader it means data the native
# code did not expect, so it is reported like any other undecodable input.
_panic_types: tuple[type[BaseException], ...] = ()


def _register_panic_type(panic_type: type[BaseException]) -> None:
    global _panic_types
    _panic_types = (panic_type,)


def _invalid(path: str | os.PathLike[str], exc: BaseException) -> InvalidFileError:
    return InvalidFileError(f"cannot read {os.fspath(path)!r}: {type(exc).__name__}: {exc}")


@contextmanager
def reading(path: str | os.PathLike[str]) -> Iterator[None]:
    """Report a decoding error raised inside the block as InvalidFileError."""
    try:
        yield
    except InvalidFileError:
        raise
    except _DECODE_ERRORS as exc:
        raise _invalid(path, exc) from exc
    except LookupError as exc:
        if isinstance(exc, KeyError):
            raise
        raise _invalid(path, exc) from exc  # an unknown text encoding
    except BaseException as exc:
        if isinstance(exc, _panic_types):
            raise _invalid(path, exc) from exc
        raise


def reading_rows(path: str | os.PathLike[str], rows: Iterable[_T]) -> Iterator[_T]:
    """``rows``, with decoding errors raised while iterating reported as InvalidFileError."""
    with reading(path):
        yield from rows
