"""The regex sheetData loader must build the same model as the element reader."""

import datetime
from pathlib import Path

import pytest

from xlspy import Workbook, load_workbook
from xlspy.workbook import _reader

DATA = Path(__file__).parent / "data"
SAMPLES = sorted(DATA.rglob("*.xlsx"))


def snapshot(workbook):
    result = {}
    for sheet in workbook.worksheets:
        cells = {}
        for row, columns in sheet._cells.items():
            for column, cell in columns.items():
                cells[(row, column)] = (
                    cell._value, cell._data_type, cell._formula, cell._style_id,
                )
        dimensions = {
            number: (d.height, d.hidden, d.outlineLevel, d.collapsed, d.bestFit, d.style)
            for number, d in sheet.row_dimensions.items()
        }
        result[sheet.title] = (
            cells, dimensions, sheet._max_row, sheet._max_column,
            sorted(map(str, sheet.merged_cells.ranges)),
        )
    return result


def load_both(path, monkeypatch, **options):
    monkeypatch.setattr(_reader, "_FAST_SHEET_DATA", True)
    fast = load_workbook(path, **options)
    monkeypatch.setattr(_reader, "_FAST_SHEET_DATA", False)
    slow = load_workbook(path, **options)
    try:
        return snapshot(fast), snapshot(slow)
    finally:
        fast.close()
        slow.close()


@pytest.mark.parametrize("path", SAMPLES, ids=lambda p: p.name)
@pytest.mark.parametrize("data_only", [False, True])
def test_sample_files_load_identically(path, data_only, monkeypatch):
    fast, slow = load_both(path, monkeypatch, data_only=data_only)
    assert fast == slow


@pytest.mark.parametrize("data_only", [False, True])
def test_generated_workbook_loads_identically(tmp_path, monkeypatch, data_only):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["text", 1, 2.5, True, datetime.date(2024, 2, 29), "a & b < c"])
    sheet.append(["=A1", "=SUM(B1:C1)", None, datetime.datetime(2024, 1, 1, 13, 30)])
    sheet.append([datetime.time(8, 15), datetime.timedelta(hours=30), "_x0041_", "line\nbreak"])
    sheet["H10"] = "far"
    sheet["B3"].number_format = "0.00%"
    sheet.row_dimensions[2].height = 30
    sheet.row_dimensions[4].hidden = True
    sheet.merge_cells("A12:C13")
    path = tmp_path / "generated.xlsx"
    workbook.save(path)
    workbook.close()
    fast, slow = load_both(path, monkeypatch, data_only=data_only)
    assert fast == slow
    assert fast["Sheet"][0][(1, 1)][0] == "text"


def test_rich_text_option_uses_the_element_reader(tmp_path, monkeypatch):
    workbook = Workbook()
    workbook.active["A1"] = "plain"
    path = tmp_path / "rich.xlsx"
    workbook.save(path)
    workbook.close()
    fast, slow = load_both(path, monkeypatch, rich_text=True)
    assert fast == slow


def test_record_payload_iterator_matches_record_iterator(tmp_path):
    import zipfile

    from xlspy.biff_utils import iter_record_payloads, iter_records

    workbook = Workbook()
    sheet = workbook.active
    for number in range(1, 200):
        sheet.append([f"s{number}", number, number / 7, number % 2 == 0])
    sheet.merge_cells("F1:G2")
    path = tmp_path / "records.xlsb"
    workbook.save(path)
    workbook.close()
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    expected = [(r.record_id, data[r.data_start:r.data_end]) for r in iter_records(data)]
    assert list(iter_record_payloads(data)) == expected
    with pytest.raises(ValueError, match="Truncated"):
        list(iter_record_payloads(b"\x01\x0aabc"))  # claims 10 bytes, has 3


def _cell_state(sheet):
    return (
        sheet._max_row, sheet._max_column, sorted(sheet._dirty_cells),
        sorted(sheet._dirty_styles), sorted(sheet._dirty_meta), sheet._workbook._styles_dirty,
        {
            (row, column): (cell._value, cell._data_type, str(cell._formula),
                            cell._style_id, cell._hyperlink, cell._dirty, cell.number_format)
            for row, cells in sheet._cells.items() for column, cell in cells.items()
        },
    )


APPEND_ROWS = [
    ["text", 1, 2.5, True, None, "=A1*2", datetime.date(2024, 1, 2)],
    [None, None],
    [],
    ["", 0, -0.0, False, datetime.datetime(2024, 1, 2, 3, 4), "x" * 10, None],
    (value for value in ("gen", 7)),
    # Later date rows take the cached default date styles.
    [datetime.date(2024, 5, 6), datetime.datetime(2024, 5, 6, 7, 8, 9), 3],
    [datetime.datetime(2025, 1, 1), None, datetime.date(1999, 12, 31)],
    [datetime.time(1, 2), datetime.timedelta(hours=5)],
]


def test_append_matches_cell_by_cell():
    """append() into a new row stores what cell(row, column, value) would."""
    fast, slow = Workbook(), Workbook()
    for values in APPEND_ROWS:
        values = list(values)
        fast.active.append(values)
        row = slow.active._max_row + 1
        for column, value in enumerate(values, 1):
            if value is not None:
                slow.active.cell(row, column, value)
            else:
                slow.active._max_column = max(slow.active._max_column, column)
        if values:
            slow.active._max_row = max(slow.active._max_row, row)
    assert _cell_state(fast.active) == _cell_state(slow.active)
    assert fast.active._dirty is slow.active._dirty is True


def test_append_dates_round_trip_with_their_formats(tmp_path):
    workbook = Workbook()
    for day in range(1, 4):
        workbook.active.append([datetime.date(2024, 1, day), datetime.datetime(2024, 1, day, 12)])
    workbook.save(tmp_path / "dates.xlsx")
    loaded = load_workbook(tmp_path / "dates.xlsx")
    try:
        sheet = loaded.active
        assert sheet["A3"].value == datetime.datetime(2024, 1, 3)
        assert sheet["B3"].value == datetime.datetime(2024, 1, 3, 12)
        assert sheet["A3"].number_format == workbook.active["A1"].number_format
        assert sheet["B3"].number_format == workbook.active["B1"].number_format
    finally:
        loaded.close()


def test_append_into_merged_range_still_refused():
    sheet = Workbook().active
    sheet.append([1, 2])
    sheet.merge_cells("A2:B2")
    with pytest.raises(AttributeError, match="merged"):
        sheet.append([None, "inside merge"])


def test_append_keeps_cells_already_in_the_row():
    # An empty cell object does not count towards max_row, but append() must
    # fill it rather than replace it.
    sheet = Workbook().active
    sheet.append(["a"])
    existing = sheet._set_cell_raw(2, 1, None)
    sheet.append(["b", 2])
    assert sheet._existing_cell(2, 1) is existing and existing.value == "b"
    assert sheet["B2"].value == 2


def _model_snapshot(workbook):
    return {
        sheet.title: sorted(
            (cell.row, cell.column, repr(cell.value), cell.data_type, cell._style_id)
            for row in sheet._cells.values() for cell in row.values()
        ) + [("max", sheet.max_row, sheet.max_column)]
        + sorted((row, dim.height, dim.hidden) for row, dim in sheet.row_dimensions.items())
        for sheet in workbook.worksheets
    }


@pytest.mark.parametrize("data_only", [False, True])
def test_native_model_rows_match_python_rows(tmp_path, monkeypatch, data_only):
    import datetime

    import openpyxl

    from xlspy import _accel
    from xlspy.workbook import load_workbook as xlspy_load

    if _accel._c_read_xlsx_worksheet is None:
        pytest.skip("native extension not built")
    source = openpyxl.Workbook()
    sheet = source.active
    sheet.append(["text", 1, 2.5, True, datetime.datetime(2024, 5, 6, 7, 8), None, "=A1&\"x\""])
    sheet.append(["a & b", -3, 1e300, False, datetime.date(1900, 1, 1), " lead", "#N/A"])
    sheet["A3"] = datetime.time(13, 45, 30)
    sheet["A3"].number_format = "hh:mm:ss"
    sheet["B3"] = datetime.timedelta(hours=30)
    sheet["B3"].number_format = "[h]:mm"
    sheet["C3"] = 45000.5
    sheet["C3"].number_format = "yyyy-mm-dd hh:mm"
    sheet["D4"].font = openpyxl.styles.Font(bold=True)  # a styled blank cell
    sheet["Z10"] = "far"
    sheet.row_dimensions[5].height = 30
    sheet["A5"] = "tall"
    sheet["A6"] = "=SUM(B1:B2)"
    other = source.create_sheet("Other")
    other["B2"] = "=Sheet!B1*2"
    other["C3"] = "x\ny"
    path = tmp_path / "kinds.xlsx"
    source.save(path)

    with_c = xlspy_load(path, data_only=data_only)
    try:
        expected = _model_snapshot(with_c)
    finally:
        with_c.close()
    monkeypatch.setattr(_accel, "_c_read_xlsx_worksheet", None)
    python = xlspy_load(path, data_only=data_only)
    try:
        assert _model_snapshot(python) == expected
    finally:
        python.close()
