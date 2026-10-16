import os as _os

HAVE_C_EXT = False
_c_encode_xlsb_row = None
_c_calc_column_widths = None
_c_read_xlsb_worksheet = None

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
