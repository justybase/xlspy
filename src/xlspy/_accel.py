"""Native accelerator for the hot paths: ``xlspy._rs_core`` (Rust/PyO3).

``XLSPY_ACCEL`` chooses the backend:

- ``auto`` (default) or ``rust``: the Rust module when it is built.
- ``python``: no native code; same as ``XLSPY_DISABLE_C_EXT=1`` (the name
  predates the Rust port and is kept for compatibility).

Every wrapper below returns None when no backend provides the function, and
the caller then uses the pure-Python implementation.
"""
import importlib as _importlib
import os as _os

from .exceptions import InvalidFileError

_NATIVE_NAMES = (
    "encode_xlsb_row", "encode_xlsb_row_inline", "calc_column_widths",
    "read_xlsb_worksheet", "read_xlsx_worksheet", "xlsb_row_has_formatted",
    "encode_xlsb_sst", "read_xlsb_shared_strings", "encode_xlsx_row",
    "read_xlsx_shared_strings",
    "set_style_type", "set_error_type", "xlsb_load_cells",
)
_MODULES = {"rust": "xlspy._rs_core"}
_ORDER = {"auto": ("rust",), "rust": ("rust",), "python": ()}

ACCEL_MODE = _os.environ.get("XLSPY_ACCEL", "auto").strip().lower() or "auto"
if _os.environ.get("XLSPY_DISABLE_C_EXT", "0") == "1":
    ACCEL_MODE = "python"
if ACCEL_MODE not in _ORDER:
    raise ImportError(f"XLSPY_ACCEL must be one of {sorted(_ORDER)}, not {ACCEL_MODE!r}")


def _import(backend):
    try:
        return _importlib.import_module(_MODULES[backend])
    except ImportError:
        return None


_loaded = {b: m for b in _ORDER[ACCEL_MODE] if (m := _import(b)) is not None}

# Function name -> backend ("rust") that provides it.
ACCEL_BACKEND = {}
_impl = {}
for _name in _NATIVE_NAMES:
    for _backend, _module in _loaded.items():
        _func = getattr(_module, _name, None)
        if _func is not None:
            _impl[_name] = _func
            ACCEL_BACKEND[_name] = _backend
            break

# The _c_* names predate the Rust port; other modules import some of them.
_c_encode_xlsb_row = _impl.get("encode_xlsb_row")
_c_calc_column_widths = _impl.get("calc_column_widths")
_c_read_xlsb_worksheet = _impl.get("read_xlsb_worksheet")
_c_read_xlsx_worksheet = _impl.get("read_xlsx_worksheet")
_c_row_has_formatted = _impl.get("xlsb_row_has_formatted")
_c_encode_xlsb_sst = _impl.get("encode_xlsb_sst")
_c_read_sst = _impl.get("read_xlsb_shared_strings")
_c_read_xlsx_sst = _impl.get("read_xlsx_shared_strings")
_c_encode_xlsx_row = _impl.get("encode_xlsx_row")
_c_encode_xlsb_row_inline = _impl.get("encode_xlsb_row_inline")
# Workbook model loader for XLSB sheet data (see workbook/_xlsb_reader.py).
xlsb_load_cells = _impl.get("xlsb_load_cells")

# True when the native backend provides the core readers and writers.
HAVE_NATIVE = all(n in _impl for n in ("encode_xlsb_row", "calc_column_widths", "read_xlsb_worksheet"))
# The pre-Rust name, still used across the code base and by callers.
HAVE_C_EXT = HAVE_NATIVE

# The native module keeps a registry of the xlspy types it checks.
for _module in _loaded.values():
    if hasattr(_module, "set_error_type"):
        from xlspy._errors import CellError as _CellError
        _module.set_error_type(_CellError)
    if hasattr(_module, "set_invalid_file_type"):
        # Data errors the native readers raise are InvalidFileError (a ValueError).
        _module.set_invalid_file_type(InvalidFileError)
    if hasattr(_module, "set_style_type"):
        # (value, Style) cells: the native encoders and width scan recognise them.
        from xlspy.styles import Style as _Style
        _module.set_style_type(_Style)


def encode_xlsb_row_inline(row, row_idx, style_ids=None, fmt_xf=None, row_header=None):
    """Row bytes with inline text cells natively, or None to use the Python
    writer (over-long or unpaired-surrogate text, tuple-formatted cells whose
    xf is not in ``fmt_xf`` yet). ``row_header`` replaces the default
    BrtRowHdr record."""
    if _c_encode_xlsb_row_inline is not None:
        if fmt_xf is not None or row_header is not None:
            return _c_encode_xlsb_row_inline(row, row_idx, style_ids, fmt_xf, row_header)
        if style_ids is not None:
            return _c_encode_xlsb_row_inline(row, row_idx, style_ids)
        return _c_encode_xlsb_row_inline(row, row_idx)
    return None


def encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx,
                    style_ids=None, fmt_xf=None, row_header=None):
    if HAVE_NATIVE:
        if fmt_xf is not None or row_header is not None:
            return _c_encode_xlsb_row(
                row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx,
                style_ids, fmt_xf, row_header,
            )
        if style_ids is not None:
            return _c_encode_xlsb_row(
                row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx, style_ids
            )
        return _c_encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx)
    return None


def calc_column_widths(rows, max_cols):
    if HAVE_NATIVE:
        return _c_calc_column_widths(rows, max_cols)
    return None


def read_xlsb_worksheet(data, shared_strings, styles_list, date_num_fmts, date_1904=False):
    if HAVE_NATIVE:
        return _c_read_xlsb_worksheet(
            data, shared_strings, styles_list, date_num_fmts, date_1904
        )
    return None


def read_xlsx_worksheet(stream, head, shared_strings, is_date_style, date_1904=False):
    """Row iterator over an XLSX worksheet stream, or None to use the Python
    fast path.  Items are ``(row_number, cells)``; a ``bytes`` item is a row
    the caller must parse with ElementTree (then call ``rows.resume(n)``)."""
    if _c_read_xlsx_worksheet is not None:
        return _c_read_xlsx_worksheet(
            stream, head, shared_strings, is_date_style, date_1904
        )
    return None


def xlsb_row_has_formatted(row):
    """True/False natively pre-scan, or None to use the Python check."""
    if _c_row_has_formatted is not None:
        result = _c_row_has_formatted(row)
        return None if result is None else bool(result)
    return None


def encode_xlsb_sst(ss_list, sst_all_count, sst_unique_count):
    """Complete sharedStrings.bin payload, or None to use the Python writer."""
    if _c_encode_xlsb_sst is not None:
        return _c_encode_xlsb_sst(ss_list, sst_all_count, sst_unique_count)
    return None


def read_xlsb_shared_strings(data):
    """Shared-strings list from a sharedStrings.bin part, or None to use
    the Python reader."""
    if _c_read_sst is not None:
        return _c_read_sst(data)
    return None


def read_xlsx_shared_strings(data):
    """Shared strings from an XLSX sharedStrings part (bytes), or None when the
    native parser does not take the part; the regex and expat parsers then read it."""
    if _c_read_xlsx_sst is not None:
        return _c_read_xlsx_sst(data)
    return None


def encode_xlsx_row(row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
                    use_shared_strings, style_ids=None, fmt_xf=None, row_attrs=None):
    """(row_xml as UTF-8 bytes, uniq, all) natively, or None to use the Python row writer
    (unencodable rows, and tuple-formatted cells whose xf is not in
    ``fmt_xf`` yet).  ``row_attrs`` are extra ``<row>`` attributes.  Native errors
    propagate."""
    if _c_encode_xlsx_row is not None:
        if style_ids is not None or fmt_xf is not None or row_attrs is not None:
            return _c_encode_xlsx_row(
                row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
                use_shared_strings, style_ids, fmt_xf, row_attrs,
            )
        return _c_encode_xlsx_row(row, ss_dict, ss_list, sst_uniq,
                                  sst_all, row_idx, use_shared_strings)
    return None
