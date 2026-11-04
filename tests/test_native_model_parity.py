"""The model's native loops (ModelCells) and the XLSB sheet scan against their Python twins."""
from __future__ import annotations

import datetime
import decimal
import hashlib
import random
import struct
import zipfile

import pytest

from xlspy import (
    Comment,
    Font,
    Formula,
    Hyperlink,
    Workbook,
    XlsbWriter,
    _accel,
    _xlsb_edit,
    load_workbook,
)
from xlspy.biff_utils import build_record
from xlspy.exceptions import InvalidFileError
from xlspy.rich_text import CellRichText
from xlspy.workbook import _worksheet, _writer
from xlspy.xlsb_updater import XlsbUpdater

pytestmark = pytest.mark.skipif(_accel.ModelCells is None, reason="needs the native module")


class _Int(int):
    pass


class _Text(str):
    pass


class _Day(datetime.date):
    pass


_VALUES = [
    None, None, 0, 1, -7, 2**70, 1.5, float("inf"), True, False, "", "text", "=A1*2", " =x", "żółć",
    "\ud800", _Int(4), _Text("t"), _Text("=1"), _Day(2024, 2, 3), datetime.date(2024, 1, 2),
    datetime.datetime(2024, 1, 2, 3, 4), datetime.time(1, 2), datetime.timedelta(hours=5),
    decimal.Decimal("1.5"), Formula("1+1"), Hyperlink("https://x.example/", "x"),
    CellRichText(["a", "b"]), b"bytes",
]


def _python_only(patch):
    patch.setattr(_worksheet, "NATIVE_CELLS", None)
    patch.setattr(_writer, "NATIVE_CELLS", None)


def _state(sheet):
    return (
        sheet._max_row, sheet._max_column, sorted(sheet._dirty_cells), sorted(sheet._dirty_styles),
        sorted(sheet._dirty_hyperlinks), sorted(sheet._dirty_meta), sheet._dirty, sheet._workbook._styles_dirty,
        {
            row: [(column, type(cell), type(cell._value), repr(cell._value), repr(cell._formula), cell._data_type,
                   cell._style_id, repr(cell._hyperlink), cell._comment, cell._named_style, cell._sheet is sheet,
                   cell._row, cell._column)
                  for column, cell in cells.items()]
            for row, cells in sheet._cells.items()
        },
    )


def _rows(seed):
    rng = random.Random(seed)
    return [[rng.choice(_VALUES) for _ in range(rng.randint(0, 8))] for _ in range(rng.randint(1, 8))]


def _appended(rows):
    workbook = Workbook()
    for row in rows:
        workbook.active.append(list(row))
    return workbook


def _decorate(sheet, seed):
    rng = random.Random(seed)
    for _ in range(rng.randint(0, 6)):
        cell = sheet.cell(rng.randint(1, 10), rng.randint(1, 10))
        choice = rng.randint(0, 3)
        if choice == 0:
            cell.font = Font(bold=True)
        elif choice == 1:
            cell.comment = Comment("note", "Jan")
        elif choice == 2:
            cell.number_format = "0.00"
        else:
            cell.font = Font(italic=True)
            cell.comment = Comment("both", "Ola")


@pytest.mark.parametrize("seed", range(300))
def test_append_stores_what_python_stores(seed, monkeypatch):
    rows = _rows(seed)
    native = _appended(rows)
    with monkeypatch.context() as patch:
        _python_only(patch)
        python = _appended(rows)
    assert _state(native.active) == _state(python.active)


@pytest.mark.parametrize("seed", range(300))
def test_the_writer_gets_the_rows_and_notes_python_gives_it(seed, monkeypatch):
    workbook = _appended(_rows(seed))
    sheet = workbook.active
    _decorate(sheet, seed)
    native = (repr(list(_writer._cell_values(sheet))), repr(_writer._styles_and_comments(sheet)))
    with monkeypatch.context() as patch:
        _python_only(patch)
        python = (repr(list(_writer._cell_values(sheet))), repr(_writer._styles_and_comments(sheet)))
    assert native == python


def test_a_new_workbook_keeps_no_list_of_changed_cells():
    workbook = _appended([["text", 1, datetime.date(2024, 1, 2)]])
    sheet = workbook.active
    sheet["D5"] = 3
    sheet.cell(6, 1, "x")
    assert not sheet._dirty_cells and sheet._dirty is True


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: hashlib.sha256(archive.read(name)).hexdigest()
                for name in archive.namelist() if name != "docProps/core.xml"}


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_rows_appended_to_a_loaded_workbook_are_saved(tmp_path, ext, monkeypatch):
    # A loaded workbook is patched from the list of changed cells: append has to fill it.
    source = tmp_path / f"source.{ext}"
    workbook = Workbook()
    workbook.active.append(["head", 1])
    workbook.save(str(source))
    outputs = []
    for label in ("native", "python"):
        with monkeypatch.context() as patch:
            if label == "python":
                _python_only(patch)
            loaded = load_workbook(str(source))
            sheet = loaded.active
            sheet.append(["new", 2, None, datetime.date(2024, 1, 2), "=1+1"])
            assert sorted(sheet._dirty_cells) == [(2, 1), (2, 2), (2, 4), (2, 5)]
            outputs.append(tmp_path / f"{label}.{ext}")
            loaded.save(str(outputs[-1]))
            loaded.close()
    assert _parts(outputs[0]) == _parts(outputs[1])
    again = load_workbook(str(outputs[0])).active
    assert [again.cell(2, column).value for column in (1, 2)] == ["new", 2]
    assert again.cell(2, 5).value == "=1+1"


@pytest.fixture
def xlsb_sheet(tmp_path):
    path = tmp_path / "scan.xlsb"
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["id", "text"], [1, "x"], [2, Formula("A2+1")], [None, None, None, 5]], merges=["F2:G3"])
    return path


def test_the_sheet_scan_answers_like_python_on_a_written_sheet(xlsb_sheet, monkeypatch):
    with zipfile.ZipFile(xlsb_sheet) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    native = _accel.xlsb_sheet_scan(data)
    assert native is not None
    with monkeypatch.context() as patch:
        patch.setattr(_accel, "xlsb_sheet_scan", None)
        assert _xlsb_edit._scan_sheet(data) == native
    assert native[4] == [(1, 2, 5, 6)]


def test_update_cells_writes_the_same_bytes_with_either_scan(xlsb_sheet, tmp_path, monkeypatch):
    outputs = []
    for label in ("native", "python"):
        with monkeypatch.context() as patch:
            if label == "python":
                patch.setattr(_accel, "xlsb_sheet_scan", None)
            updater = XlsbUpdater(str(xlsb_sheet))
            try:
                updater.update_cells("Data", {"B2": "new", "A3": None, "H9": 4, "Z1": True})
                outputs.append(tmp_path / f"{label}.xlsb")
                updater.save(str(outputs[-1]))
            finally:
                updater.close()
    assert _parts(outputs[0]) == _parts(outputs[1])


_BEGIN, _END = build_record(0x91, b""), build_record(0x92, b"")
_ROW = build_record(0, struct.pack("<I", 0) + bytes(13))


@pytest.mark.parametrize("data, message", [
    (b"", "no complete sheet data"),
    (_BEGIN, "no complete sheet data"),
    (_END, "no complete sheet data"),
    (_BEGIN + build_record(0, b"\x00") + _END, "row record"),
    (_BEGIN + _ROW + build_record(2, b"\x00") + _END, "cell record"),
    (build_record(0x94, bytes(15)) + _BEGIN + _END, "dimension"),
    (_BEGIN + _END + build_record(0xB0, bytes(15)), "cut short"),
    (_BEGIN + _END + build_record(0x1AA, bytes(4)), "cut short"),
])
def test_a_sheet_the_native_scan_leaves_to_python_is_refused_there(data, message):
    assert _accel.xlsb_sheet_scan(data) is None
    with pytest.raises(InvalidFileError, match=message):
        _xlsb_edit._scan_sheet(data)


def test_compact_cells_follow_the_column_before_them():
    # BrtCellBlank at column 3, then two compact cells (0xC and up carry no column).
    data = (_BEGIN + _ROW + build_record(1, struct.pack("<II", 3, 0)) + build_record(0xD, bytes(4))
            + build_record(0xE, bytes(4)) + _END)
    assert _accel.xlsb_sheet_scan(data)[7] == (3, 0, 5, 0)
