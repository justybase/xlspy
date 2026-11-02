"""The one place where xlspy parses XML from a package.

Excel never writes a document type declaration (DTD), so xlspy refuses one.
Refusing the DTD also refuses every entity declaration, which closes the usual
routes of entity-expansion and external-entity attacks. A DTD can only appear in
the prolog (before the root element), so the check reads the prolog and no more;
the rest of the document goes to the standard library unchanged.

Three entry points cover the call sites:

- ``fromstring`` replaces ``ElementTree.fromstring``.
- ``iterparse`` replaces ``ElementTree.iterparse`` on a binary stream.
- ``harden_expat`` sets the same refusal on a raw ``pyexpat`` parser.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import IO, Any

__all__ = ["fromstring", "iterparse", "harden_expat"]

# Items allowed in the prolog: a byte-order mark, white space, processing
# instructions (including the XML declaration) and comments.
_PROLOG_BYTES = re.compile(rb"(?:\xef\xbb\xbf)?\s*(?:<\?.*?\?>|<!--.*?-->)", re.DOTALL)
_PROLOG_STR = re.compile(r"﻿?\s*(?:<\?.*?\?>|<!--.*?-->)", re.DOTALL)
_DOCTYPE_BYTES = re.compile(rb"\s*<!DOCTYPE")
_DOCTYPE_STR = re.compile(r"\s*<!DOCTYPE")
_ROOT_BYTES = re.compile(rb"\s*<[A-Za-z_:]")
_ROOT_STR = re.compile(r"\s*<[A-Za-z_:]")

# A prolog longer than this without a root element is refused: it is not a
# spreadsheet part, and it would otherwise keep the guard reading forever.
_MAX_PROLOG_CHARS = 1 << 20

_REFUSAL = "document type declarations are not supported: Excel packages do not contain them"


def _scan(text) -> tuple[bool, bool]:
    """Return ``(has_doctype, decided)`` for a prolog that may be partly read.

    ``decided`` is true once the root element (or a DTD) has been seen, so no
    later text can be part of the prolog.
    """
    item: Any
    doctype: Any
    root: Any
    if isinstance(text, str):
        item, doctype, root = _PROLOG_STR, _DOCTYPE_STR, _ROOT_STR
    else:
        item, doctype, root = _PROLOG_BYTES, _DOCTYPE_BYTES, _ROOT_BYTES
    position = 0
    while True:
        match = item.match(text, position)
        if match is None or match.end() == position:
            break
        position = match.end()
    if doctype.match(text, position):
        return True, True
    return False, root.match(text, position) is not None


def _refuse_if_doctype(text) -> None:
    has_doctype, _ = _scan(text)
    if has_doctype:
        raise ET.ParseError(_REFUSAL)


def fromstring(text):
    """``ElementTree.fromstring`` that refuses a document type declaration."""
    _refuse_if_doctype(text)
    return ET.fromstring(text)


class _PrologGuard:
    """A binary stream that checks the prolog of what it returns.

    Reads pass through unchanged. Until the root element has been seen, the
    returned bytes are also kept so the prolog can be checked; the check stops
    once the root element is found.
    """

    def __init__(self, stream: IO[bytes]):
        self._stream = stream
        self._head = b""
        self._decided = False

    def read(self, size: int = -1) -> bytes:
        data = self._stream.read(size)
        if not self._decided:
            self._head += data
            has_doctype, decided = _scan(self._head)
            if has_doctype:
                raise ET.ParseError(_REFUSAL)
            if decided:
                self._decided = True
                self._head = b""
            elif len(self._head) > _MAX_PROLOG_CHARS:
                raise ET.ParseError("XML prolog is too long: no root element found")
        return data


def iterparse(stream: IO[bytes], events=None):
    """``ElementTree.iterparse`` over a binary stream, refusing a DTD."""
    return ET.iterparse(_PrologGuard(stream), events=events)


def _refuse_declaration(*_args: Any) -> None:
    raise ET.ParseError(_REFUSAL)


def harden_expat(parser):
    """Refuse DTDs and entity declarations on a raw ``pyexpat`` parser.

    Returns the parser so the call can wrap ``ParserCreate(...)``.
    """
    parser.StartDoctypeDeclHandler = _refuse_declaration
    parser.EntityDeclHandler = _refuse_declaration
    return parser
