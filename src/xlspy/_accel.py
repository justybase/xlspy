"""Native accelerator for the hot paths: ``xlspy._rs_core`` (Rust/PyO3).

``XLSPY_ACCEL`` chooses the backend:

- ``auto`` (default) or ``rust``: the native module when it is built.
- ``python``: no native code.

``XLSPY_DISABLE_C_EXT``, the switch from before the Rust port, was removed;
setting it raises ImportError so the choice is not silently ignored.

Every wrapper below returns None when the native module does not provide the
function (or is disabled), and the caller then uses the pure-Python
implementation.
"""
from __future__ import annotations

import importlib as _importlib
import os as _os
from typing import Any

from .exceptions import InvalidFileError, _register_panic_type
from xlspy._errors import CellError as _CellError
from xlspy.styles import Style as _Style
from xlspy.formulas import Formula as _Formula
from xlspy.hyperlinks import Hyperlink as _Hyperlink

_MODULE = "xlspy._rs_core"
_MODES = ("auto", "rust", "python")

ACCEL_MODE = _os.environ.get("XLSPY_ACCEL", "auto").strip().lower() or "auto"
if "XLSPY_DISABLE_C_EXT" in _os.environ:
    raise ImportError("XLSPY_DISABLE_C_EXT was removed; set XLSPY_ACCEL=python to disable the native module")
if ACCEL_MODE not in _MODES:
    raise ImportError(f"XLSPY_ACCEL must be one of {sorted(_MODES)}, not {ACCEL_MODE!r}")

_NATIVE_NAMES = (
    "encode_xlsb_row", "encode_xlsb_row_inline",
    "read_xlsb_worksheet", "read_xlsx_worksheet", "xlsb_row_has_formatted",
    "encode_xlsb_sst", "read_xlsb_shared_strings", "encode_xlsx_row",
    "read_xlsx_shared_strings",
    "set_style_type", "set_error_type", "xlsb_load_cells", "xlsb_sheet_scan", "ModelCells", "xlsx_edit_scan",
    "WidthMeter", "general_number_text",
    "encode_update_xlsx_row", "encode_update_xlsb_row", "patch_xlsb_row_headers",
)


def _load_native():
    if ACCEL_MODE == "python":
        return None
    try:
        return _importlib.import_module(_MODULE)
    except ImportError:
        return None


_native = _load_native()

# Function name -> backend that provides it. CI and the benchmarks read it.
ACCEL_BACKEND: dict[str, str] = {}
_impl = {}
if _native is not None:
    for _name in _NATIVE_NAMES:
        _func = getattr(_native, _name, None)
        if _func is not None:
            _impl[_name] = _func
            ACCEL_BACKEND[_name] = "rust"


def _register_types(module) -> None:
    """Hand the native module the xlspy classes it checks and builds."""
    if hasattr(module, "set_error_type"):
        module.set_error_type(_CellError)
    panic_type = getattr(module, "PanicException", None)
    if panic_type is not None:
        _register_panic_type(panic_type)
    if hasattr(module, "set_invalid_file_type"):
        # Data errors the native readers raise are InvalidFileError (a ValueError).
        module.set_invalid_file_type(InvalidFileError)
    if hasattr(module, "set_style_type"):
        # (value, Style) cells: the native encoders and width scan recognise them.
        module.set_style_type(_Style)
    if hasattr(module, "set_formula_type"):
        # Formula cells: the native row encoder writes them (see xlsx_write.rs).
        module.set_formula_type(_Formula)
    if hasattr(module, "set_hyperlink_type"):
        # Link cells: the native row encoder writes the text and returns the link.
        module.set_hyperlink_type(_Hyperlink)


if _native is not None:
    _register_types(_native)

# True when the native backend provides the core readers and writers.
HAVE_NATIVE = all(n in _impl for n in ("encode_xlsb_row", "read_xlsb_worksheet"))

# Raw functions of the native module (None when it does not provide them).
_native_encode_xlsb_row: Any = _impl.get("encode_xlsb_row")
_native_encode_xlsb_row_inline: Any = _impl.get("encode_xlsb_row_inline")
_native_read_xlsb_worksheet: Any = _impl.get("read_xlsb_worksheet")
_native_read_xlsx_worksheet: Any = _impl.get("read_xlsx_worksheet")
_native_row_has_formatted: Any = _impl.get("xlsb_row_has_formatted")
_native_encode_xlsb_sst: Any = _impl.get("encode_xlsb_sst")
_native_read_xlsb_sst: Any = _impl.get("read_xlsb_shared_strings")
_native_read_xlsx_sst: Any = _impl.get("read_xlsx_shared_strings")
_native_encode_xlsx_row: Any = _impl.get("encode_xlsx_row")
# Workbook model loader for XLSB sheet data (see workbook/_xlsb_reader.py).
xlsb_load_cells: Any = _impl.get("xlsb_load_cells")
# One pass over an XLSB sheet for update_cells (see _xlsb_edit._scan_sheet).
xlsb_sheet_scan: Any = _impl.get("xlsb_sheet_scan")
# One pass over an XLSX sheetData for update_cells (see _xlsx_edit._fast_scan).
xlsx_edit_scan: Any = _impl.get("xlsx_edit_scan")
# The model's hot loops (see workbook/_cell.py).
ModelCells: Any = _impl.get("ModelCells")
# The rows of XlsxUpdater / XlsbUpdater.replace_sheet_data (None: the Python code).
encode_update_xlsx_row: Any = _impl.get("encode_update_xlsx_row")
encode_update_xlsb_row: Any = _impl.get("encode_update_xlsb_row")
patch_xlsb_row_headers: Any = _impl.get("patch_xlsb_row_headers")
# Column-width meter for the rows after the first (see _widths.ColumnWidthMeter).
WidthMeter: Any = _impl.get("WidthMeter")
general_number_text: Any = _impl.get("general_number_text")


def encode_xlsb_row_inline(row, row_idx, style_ids=None, fmt_xf=None, row_header=None, cell_hook=None):
    """Row bytes with inline text cells natively, or None to use the Python
    writer (over-long or unpaired-surrogate text, tuple-formatted cells whose
    xf is not in ``fmt_xf`` yet). ``row_header`` replaces the default
    BrtRowHdr record; ``cell_hook(cell, column, xf, row)`` encodes the cells
    of other types (see ``cells.rs``, ``call_hook``)."""
    native = _native_encode_xlsb_row_inline
    if native is None:
        return None
    if cell_hook is not None:
        return native(row, row_idx, style_ids, fmt_xf, row_header, cell_hook)
    if fmt_xf is not None or row_header is not None:
        return native(row, row_idx, style_ids, fmt_xf, row_header)
    if style_ids is not None:
        return native(row, row_idx, style_ids)
    return native(row, row_idx)


def encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx,
                    style_ids=None, fmt_xf=None, row_header=None, cell_hook=None):
    if not HAVE_NATIVE:
        return None
    if cell_hook is not None:
        return _native_encode_xlsb_row(
            row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx,
            style_ids, fmt_xf, row_header, cell_hook,
        )
    if fmt_xf is not None or row_header is not None:
        return _native_encode_xlsb_row(
            row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx,
            style_ids, fmt_xf, row_header,
        )
    if style_ids is not None:
        return _native_encode_xlsb_row(
            row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx, style_ids
        )
    return _native_encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx)


def read_xlsb_worksheet(data, shared_strings, styles_list, date_num_fmts, date_1904=False):
    if HAVE_NATIVE:
        return _native_read_xlsb_worksheet(
            data, shared_strings, styles_list, date_num_fmts, date_1904
        )
    return None


def read_xlsx_worksheet(stream, head, shared_strings, is_date_style, date_1904=False):
    """Row iterator over an XLSX worksheet stream, or None to use the Python
    fast path.  Items are ``(row_number, cells)``; a ``bytes`` item is a row
    the caller must parse with ElementTree (then call ``rows.resume(n)``)."""
    if _native_read_xlsx_worksheet is None:
        return None
    return _native_read_xlsx_worksheet(stream, head, shared_strings, is_date_style, date_1904)


def xlsb_row_has_formatted(row):
    """True/False natively pre-scan, or None to use the Python check."""
    if _native_row_has_formatted is None:
        return None
    result = _native_row_has_formatted(row)
    return None if result is None else bool(result)


def encode_xlsb_sst(ss_list, sst_all_count, sst_unique_count):
    """Complete sharedStrings.bin payload, or None to use the Python writer."""
    if _native_encode_xlsb_sst is None:
        return None
    return _native_encode_xlsb_sst(ss_list, sst_all_count, sst_unique_count)


def read_xlsb_shared_strings(data):
    """Shared-strings list from a sharedStrings.bin part, or None to use
    the Python reader."""
    if _native_read_xlsb_sst is None:
        return None
    return _native_read_xlsb_sst(data)


def read_xlsx_shared_strings(data):
    """Shared strings from an XLSX sharedStrings part (bytes), or None when the
    native parser does not take the part; the regex and expat parsers then read it."""
    if _native_read_xlsx_sst is None:
        return None
    return _native_read_xlsx_sst(data)


def encode_xlsx_row(row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
                    use_shared_strings, style_ids=None, fmt_xf=None, row_attrs=None,
                    link_xf=None, write_r=True, cell_hook=None):
    """(row_xml as UTF-8 bytes, uniq, all, hyperlinks) natively, or None to use the
    Python row writer (unencodable rows, and tuple-formatted cells whose xf is not
    in ``fmt_xf`` yet).  ``row_attrs`` are extra ``<row>`` attributes.  ``link_xf``
    is the style id of the built-in link style, once the writer has registered it;
    without it, link cells go to the Python writer.  ``hyperlinks`` lists
    (row, column, Hyperlink) for the link cells written.  ``write_r`` False leaves
    the ``r`` attribute off (see the native function).  ``cell_hook(cell, column,
    xf, row)`` encodes the cells of other types (see ``cells.rs``, ``call_hook``).
    Native errors propagate."""
    if _native_encode_xlsx_row is None:
        return None
    if cell_hook is not None:
        return _native_encode_xlsx_row(
            row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
            use_shared_strings, style_ids, fmt_xf, row_attrs, link_xf, write_r, cell_hook,
        )
    if not write_r:
        return _native_encode_xlsx_row(
            row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
            use_shared_strings, style_ids, fmt_xf, row_attrs, link_xf, False,
        )
    if style_ids is not None or fmt_xf is not None or row_attrs is not None or link_xf:
        return _native_encode_xlsx_row(
            row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
            use_shared_strings, style_ids, fmt_xf, row_attrs, link_xf,
        )
    return _native_encode_xlsx_row(row, ss_dict, ss_list, sst_uniq,
                                   sst_all, row_idx, use_shared_strings)
