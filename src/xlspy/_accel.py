import os as _os

HAVE_C_EXT = False
_c_encode_xlsb_row = None
_c_calc_column_widths = None
_c_read_xlsb_worksheet = None
_c_row_has_formatted = None
_c_encode_xlsb_sst = None
_c_read_sst = None
_c_encode_xlsx_row = None

if _os.environ.get('XLSPY_DISABLE_C_EXT', '0') != '1':
    try:
        from xlspy._c_core import (
            encode_xlsb_row as _c_encode_xlsb_row,
            calc_column_widths as _c_calc_column_widths,
            read_xlsb_worksheet as _c_read_xlsb_worksheet,
        )
        HAVE_C_EXT = True
    except ImportError:
        pass
    try:
        from xlspy._c_core import (
            xlsb_row_has_formatted as _c_row_has_formatted,
            encode_xlsb_sst as _c_encode_xlsb_sst,
            read_xlsb_shared_strings as _c_read_sst,
            encode_xlsx_row as _c_encode_xlsx_row,
        )
    except ImportError:
        pass


def encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx, style_ids=None):
    if HAVE_C_EXT:
        if style_ids is not None:
            return _c_encode_xlsb_row(
                row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx, style_ids
            )
        return _c_encode_xlsb_row(row, ss_dict, ss_list, sst_unique_count, sst_all_count, row_idx)
    return None


def calc_column_widths(rows, max_cols):
    if HAVE_C_EXT:
        return _c_calc_column_widths(rows, max_cols)
    return None


def read_xlsb_worksheet(data, shared_strings, styles_list, date_num_fmts, date_1904=False):
    if HAVE_C_EXT:
        return _c_read_xlsb_worksheet(
            data, shared_strings, styles_list, date_num_fmts, date_1904
        )
    return None


def xlsb_row_has_formatted(row):
    """True/False via C pre-scan, or None to use the Python check."""
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


def encode_xlsx_row(row, ss_dict, ss_list, sst_uniq, sst_all, row_idx,
                    use_shared_strings):
    """(row_xml, uniq, all) via C, or None to use the Python row writer
    (tuple-formatted rows, unencodable rows).  C errors propagate."""
    if _c_encode_xlsx_row is not None:
        return _c_encode_xlsx_row(row, ss_dict, ss_list, sst_uniq,
                                  sst_all, row_idx, use_shared_strings)
    return None
