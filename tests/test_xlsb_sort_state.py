"""Sort state of an auto filter in a loaded XLSB."""

import shutil
import struct
import zipfile
from pathlib import Path

import pytest

from xlspy import (
    UnsupportedFeatureError, Workbook, XlsxWriter, load_workbook,
)
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


def _rewrite_sheet_part(source, target, change):
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(
        target, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.bin":
                data = change(data)
            dst.writestr(item, data)


def test_sort_cond14_is_read_as_an_icon_sort(tmp_path):
    target = tmp_path / "icon.xlsb"

    def change(data):
        data = bytearray(data)
        for record in iter_records(bytes(data)):
            if record.record_id == 0x214:
                data[record.header_start:record.header_start + 2] = b"\x80\x09"
                data[record.data_start:record.data_start + 2] = struct.pack("<H", 0x07)
                break
        return bytes(data)

    _rewrite_sheet_part(FIXTURE, target, change)
    book = load_workbook(target)
    try:
        condition = book["Data"].auto_filter.sortState.sortCondition[0]
        assert condition.sortBy == "icon"
        # An unrelated edit keeps the record; changing the filter refuses it.
        book["Data"]["D10"] = 1
        book.save(target)
    finally:
        book.close()

    book = load_workbook(target)
    try:
        book["Data"].auto_filter.add_filter_column(0, ["a"])
        with pytest.raises(UnsupportedFeatureError, match="icon"):
            book.save(target)
    finally:
        book.close()


def test_sort_state_fields_are_saved(tmp_path):
    target = tmp_path / "fields.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        sort = book["Data"].auto_filter.sortState
        sort.columnSort = True
        sort.caseSensitive = True
        sort.sortMethod = "stroke"
        book.save(target)
    finally:
        book.close()

    state_flags, _condition_flags = _sort_flags(target)
    assert state_flags & 0x07 == 0x07
    reopened = load_workbook(target)
    try:
        sort = reopened["Data"].auto_filter.sortState
        assert sort.columnSort is True
        assert sort.caseSensitive is True
        assert sort.sortMethod == "stroke"
    finally:
        reopened.close()


def test_sort_condition_details_are_saved(tmp_path):
    target = tmp_path / "details.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        condition = book["Data"].auto_filter.sortState.sortCondition[0]
        condition.sortBy = "cellColor"
        condition.dxfId = 7
        condition.customList = "alpha,beta"
        book.save(target)
    finally:
        book.close()

    payload = dict(_records(target))[0x214]
    _flags, _r1, _r2, _c1, _c2, dxf_id, _second = struct.unpack_from(
        "<HIIIIII", payload
    )
    length = struct.unpack_from("<I", payload, 26)[0]
    assert dxf_id == 7
    assert payload[30:30 + 2 * length].decode("utf-16-le") == "alpha,beta"
    reopened = load_workbook(target)
    try:
        condition = reopened["Data"].auto_filter.sortState.sortCondition[0]
        assert condition.sortBy == "cellColor"
        assert condition.dxfId == 7
        assert condition.customList == "alpha,beta"
    finally:
        reopened.close()


def test_add_sort_condition_uses_the_data_range(tmp_path):
    target = tmp_path / "public.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        auto_filter = book["Data"].auto_filter
        auto_filter.sortState = None
        auto_filter.add_sort_condition("B2:B4", descending=True)
        book.save(target)
    finally:
        book.close()

    payload = dict(_records(target))[0x212]
    _flags, r1, r2, c1, c2 = struct.unpack_from("<HIIII", payload)
    assert (r1, r2, c1, c2) == (1, 3, 0, 2)  # A2:C4, not the header row
    reopened = load_workbook(target)
    try:
        assert reopened["Data"].auto_filter.sortState.ref == "A2:C4"
    finally:
        reopened.close()


def test_new_workbook_sort_state_is_refused(tmp_path):
    workbook = Workbook()
    try:
        sheet = workbook.active
        sheet.append(["h1", "h2", "h3"])
        sheet.append([1, 3, 9])
        sheet.append([2, 1, 8])
        sheet.auto_filter.ref = "A1:C3"
        sheet.auto_filter.add_sort_condition("B2:B3")
        with pytest.raises(UnsupportedFeatureError, match="sort state"):
            workbook.save(tmp_path / "new.xlsx")
    finally:
        workbook.close()


def test_add_sort_condition_in_xlsx_uses_the_data_range(tmp_path):
    source = tmp_path / "filter.xlsx"
    with XlsxWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["h1", "h2", "h3"], [1, 3, 9], [2, 1, 8], [3, 2, 7]],
            auto_filter=True,
        )
    target = tmp_path / "filtered.xlsx"
    book = load_workbook(source)
    try:
        book["Data"].auto_filter.add_sort_condition("B2:B4", descending=True)
        book.save(target)
    finally:
        book.close()
    with zipfile.ZipFile(target) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    assert 'sortState ref="A2:C4"' in xml


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
