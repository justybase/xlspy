"""Invariants for every archive xlspy produces.

Writers, updaters and the DataFrame shortcuts must all emit valid ZIP archives
whose member names are unique.  Duplicate members make different readers
(Excel, ``zipfile``, openpyxl) disagree about the file contents, so xlspy must
never create them.
"""

import datetime
import zipfile

import pytest

from xlspy import (
    ExcelReader,
    F,
    XlsbUpdater,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
)

HEADERS = ["txt", "int", "bool", "float", "date", "fmt"]
ROWS = [
    ["text & <xml>", 1, True, 2.5, datetime.date(2024, 5, 6), F.DATE_SHORT],
    ["zażółć gęślą jaźń", -2, False, None, None, None],
    [None, None, None, None, None, None],
]
NEW_ROWS = [["updated", 9, True, 1.5, datetime.date(2025, 1, 2), None]]
# A serial-date cell without an explicit date format reads back as a datetime.
EXPECTED_ROW = ["updated", 9, True, 1.5, datetime.datetime(2025, 1, 2, 0, 0), None]


def _write_xlsx(path, *, use_shared_strings=True, sheets=2):
    with XlsxWriter(path, useSharedStrings=use_shared_strings) as writer:
        for index in range(sheets):
            writer.add_sheet(f"Sheet{index + 1}")
            writer.write_sheet([HEADERS, *ROWS])


def _write_xlsb(path, *, sheets=2):
    with XlsbWriter(path) as writer:
        for index in range(sheets):
            writer.add_sheet(f"Sheet{index + 1}")
            writer.write_sheet([HEADERS, *ROWS])


def _assert_unique_members(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        duplicates = sorted({name for name in names if names.count(name) > 1})
        corrupt = archive.testzip()
    assert not duplicates, f"duplicate ZIP members in {path.name}: {duplicates}"
    assert corrupt is None, f"corrupt member {corrupt!r} in {path.name}"


def _read_sheet(path, name):
    with ExcelReader() as reader:
        reader.open(str(path))
        return [list(row) for row in reader.get_rows(name)]


@pytest.mark.parametrize("use_shared_strings", [True, False])
def test_xlsx_writer_emits_unique_members(tmp_path, use_shared_strings):
    path = tmp_path / "writer.xlsx"
    _write_xlsx(path, use_shared_strings=use_shared_strings)
    _assert_unique_members(path)


def test_xlsb_writer_emits_unique_members(tmp_path):
    path = tmp_path / "writer.xlsb"
    _write_xlsb(path)
    _assert_unique_members(path)


@pytest.mark.parametrize("streaming", [False, True])
def test_xlsx_updater_emits_unique_members(tmp_path, streaming):
    source = tmp_path / "source.xlsx"
    _write_xlsx(source)
    updater = XlsxUpdater(source)
    if streaming:
        updater.replace_sheet_data_stream("Sheet1", iter(NEW_ROWS), headers=HEADERS)
    else:
        updater.replace_sheet_data("Sheet1", NEW_ROWS, headers=HEADERS)

    as_bytes = tmp_path / "as_bytes.xlsx"
    as_bytes.write_bytes(updater.to_bytes())
    updater.save(tmp_path / "as_file.xlsx")
    updater.close()

    for path in (as_bytes, tmp_path / "as_file.xlsx"):
        _assert_unique_members(path)
        assert _read_sheet(path, "Sheet1")[1] == EXPECTED_ROW


@pytest.mark.parametrize("streaming", [False, True])
def test_xlsb_updater_emits_unique_members(tmp_path, streaming):
    source = tmp_path / "source.xlsb"
    _write_xlsb(source)
    updater = XlsbUpdater(source)
    if streaming:
        updater.replace_sheet_data_stream("Sheet1", iter(NEW_ROWS), headers=HEADERS)
    else:
        updater.replace_sheet_data("Sheet1", NEW_ROWS, headers=HEADERS)

    as_bytes = tmp_path / "as_bytes.xlsb"
    as_bytes.write_bytes(updater.to_bytes())
    updater.save(tmp_path / "as_file.xlsb")
    updater.close()

    for path in (as_bytes, tmp_path / "as_file.xlsb"):
        _assert_unique_members(path)
        assert _read_sheet(path, "Sheet1")[1] == EXPECTED_ROW


def test_dataframe_shortcuts_emit_unique_members(tmp_path):
    pandas = pytest.importorskip("pandas")
    from xlspy import to_xlsb, to_xlsx

    frame = pandas.DataFrame(
        {
            "number": [1, 2.5, None],
            "label": ["a", "b", None],
            "when": pandas.to_datetime(["2024-01-02", None, "2024-03-04"]),
        }
    )
    for filename, shortcut in (("shortcut.xlsx", to_xlsx), ("shortcut.xlsb", to_xlsb)):
        path = tmp_path / filename
        assert shortcut(frame, path) == path
        _assert_unique_members(path)
