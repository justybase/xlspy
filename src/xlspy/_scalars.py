"""Numpy scalars and pandas' missing values in raw rows, as the Python values they hold.

A row passed to ``write_sheet`` may hold numpy scalars. The DataFrame path has
always converted them with ``.item()``; the raw row path does the same for the
kinds that Excel stores as numbers or booleans: integers, ``bool_``, and
float16/float32 (``float64`` already is a ``float``). numpy is never imported
here: a scalar is recognised by its class's module and its dtype.

pandas' ``NaT`` and ``NA`` are empty cells, as in the DataFrame path (``NaT``
is a ``datetime`` whose fields are NaN, ``NA`` would be the text ``<NA>``).
"""
from __future__ import annotations
from typing import Any

# The classes of pandas.NaT and pandas.NA, recognised without importing pandas
# (their module is pandas._libs..., or pandas.api.typing from pandas 3).
_PANDAS_MISSING = frozenset({"NaTType", "NAType"})

# What the native row encoders hand back to the Python encoder.
_NOT_NATIVE = NotImplemented


def numpy_value(value: Any) -> Any:
    """``value.item()`` for a numpy int, unsigned, bool or float16/32 scalar; else ``value``."""
    if type(value).__module__ != "numpy" or getattr(value, "shape", None) != ():
        return value
    dtype = value.dtype
    kind = dtype.kind
    if kind in ("i", "u", "b") or (kind == "f" and dtype.itemsize < 8):
        return value.item()
    return value


def is_pandas_missing(value: Any) -> bool:
    """Whether ``value`` is pandas' ``NaT`` or ``NA``."""
    cls = type(value)
    return cls.__name__ in _PANDAS_MISSING and cls.__module__.startswith("pandas")


def plain_value(value: Any) -> Any:
    """The value a writer stores for ``value``: a numpy scalar's Python value
    (see ``numpy_value``), None for pandas' ``NaT`` and ``NA``, else ``value``."""
    if type(value).__module__.startswith("pandas") and is_pandas_missing(value):
        return None
    return numpy_value(value)


def native_cell_value(value: Any, keep: tuple[type, ...]) -> Any:
    """For the native row encoders' cell hook: what the Python encoders write
    for a value of a type the native encoders do not know.

    A numpy scalar becomes its Python value, ``NaT`` and ``NA`` None, and any
    other object the text ``str(value)``, as the Python encoders write them.
    Values of the ``keep`` types (formulas, links, rich text, dates the native
    encoders could not convert) are left to the Python encoder: NotImplemented.
    """
    plain = plain_value(value)
    if plain is not value:
        return plain
    if isinstance(value, keep):
        return _NOT_NATIVE
    return str(value)
