"""Sort state of an auto filter in a loaded XLSB."""

import shutil
import struct
import zipfile
from pathlib import Path

import pytest

from xlspy import load_workbook
from xlspy.biff_utils import iter_records

FIXTURE = Path(__file__).parent / "data" / "xlsb" / "sorted_autofilter.xlsb"


def _records(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    return [
        (record.record_id, data[record.data_start:record.data_end])
        for record in iter_records(data)
    ]


def _sort_flags(path):
    records = dict(_records(path))
    state = struct.unpack_from("<H", records[0x212], 0)[0]
    condition = struct.unpack_from("<H", records[0x214], 0)[0]
    return state, condition


def test_sort_state_is_read_and_written(tmp_path):
    target = tmp_path / "sorted.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        sort = book["Data"].auto_filter.sortState
        assert sort is not None
        assert sort.ref == "A2:C4"
        assert sort.sortCondition[0].ref == "B2:B4"
        assert sort.sortCondition[0].descending
    finally:
        book.close()

    # A no-op save keeps the sheet part byte for byte.
    with zipfile.ZipFile(target) as archive:
        before = archive.read("xl/worksheets/sheet1.bin")
    book = load_workbook(target)
    try:
        book.save(target)
    finally:
        book.close()
    with zipfile.ZipFile(target) as archive:
        assert archive.read("xl/worksheets/sheet1.bin") == before

    # Changing the direction rewrites the sort condition, not the filter range.
    book = load_workbook(target)
    try:
        book["Data"].auto_filter.sortState.sortCondition[0].descending = False
        book.save(target)
    finally:
        book.close()
    _state_flags, condition_flags = _sort_flags(target)
    assert condition_flags & 0x01 == 0  # ascending

    reopened = load_workbook(target)
    try:
        sort = reopened["Data"].auto_filter.sortState
        assert sort.ref == "A2:C4"
        assert not sort.sortCondition[0].descending
    finally:
        reopened.close()


@pytest.mark.com
def test_excel_reads_the_saved_sort_state(tmp_path, open_workbook):
    target = tmp_path / "sorted.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        book["Data"].auto_filter.sortState.sortCondition[0].descending = False
        book.save(target)
    finally:
        book.close()

    worksheet = open_workbook(target).Worksheets("Data")
    sort = worksheet.AutoFilter.Sort
    assert sort.SortFields.Count == 1
    assert sort.SortFields(1).Order == 1  # xlAscending
