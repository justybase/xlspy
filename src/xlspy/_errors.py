"""Error values (``#DIV/0!``, ``#N/A``, ...) as the readers return them."""
from __future__ import annotations


class CellError(str):
    """An error cell, ``"error:#DIV/0!"``: a ``str`` with the legacy prefix.

    Code that compares or slices the string keeps working; the type tells a
    real error apart from text that happens to start with ``error:``.
    """

    __slots__ = ()
