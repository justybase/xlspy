"""Object-model iter_rows pads every row to the sheet width, as openpyxl does.

The rule "the first row present sets the width" belongs to the compact
ExcelReader.get_rows() / read_all() and is tested in test_xlsx_reader_streaming.py.
It does not apply here: a short first row is still padded to the widest row.
"""

from xlspy import Workbook, load_workbook

RAGGED = [(1, None, None), (2, 3, 4), (5, None, None)]


def _ragged_sheet():
    wb = Workbook()
    ws = wb.active
    ws.append([1])
    ws.append([2, 3, 4])
    ws.append([5])
    return wb, ws


def test_new_sheet_pads_short_rows_to_sheet_width():
    _, ws = _ragged_sheet()
    assert [tuple(row) for row in ws.iter_rows(values_only=True)] == RAGGED


def test_reloaded_sheet_pads_short_rows_to_sheet_width(tmp_path):
    wb, _ = _ragged_sheet()
    path = tmp_path / "ragged.xlsx"
    wb.save(path)
    loaded = load_workbook(path)
    assert [tuple(row) for row in loaded.active.iter_rows(values_only=True)] == RAGGED


def test_read_only_iter_rows_pads_to_dimension_width(tmp_path):
    wb, _ = _ragged_sheet()
    path = tmp_path / "ragged.xlsx"
    wb.save(path)
    read_only = load_workbook(path, read_only=True)
    assert [tuple(row) for row in read_only.active.iter_rows(values_only=True)] == RAGGED
