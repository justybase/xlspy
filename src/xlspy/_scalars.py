"""Numpy scalars in raw rows, as the Python values they hold.

A row passed to ``write_sheet`` may hold numpy scalars. The DataFrame path has
always converted them with ``.item()``; the raw row path does the same for the
kinds that Excel stores as numbers or booleans: integers, ``bool_``, and
float16/float32 (``float64`` already is a ``float``). numpy is never imported
here: a scalar is recognised by its class's module and its dtype.
"""


def numpy_value(value):
    """``value.item()`` for a numpy int, unsigned, bool or float16/32 scalar; else ``value``."""
    if type(value).__module__ != "numpy" or getattr(value, "shape", None) != ():
        return value
    dtype = value.dtype
    kind = dtype.kind
    if kind in ("i", "u", "b") or (kind == "f" and dtype.itemsize < 8):
        return value.item()
    return value
