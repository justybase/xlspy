"""A hashed (ISO) sheet protection of a loaded XLSB survives an edit."""

import shutil
import struct
import zipfile
from pathlib import Path

import pytest

from xlspy import UnsupportedFeatureError, load_workbook
from xlspy.biff_utils import iter_records

FIXTURE = Path(__file__).parent / "data" / "xlsb" / "protected_iso.xlsb"


def _records(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    return [
        (record.record_id, data[record.data_start:record.data_end])
        for record in iter_records(data)
    ]


def test_hashed_protection_is_read_and_flags_are_kept(tmp_path):
    target = tmp_path / "protection.xlsb"
    shutil.copyfile(FIXTURE, target)
    source_iso = next(payload for record_id, payload in _records(target) if record_id == 0x2A6)

    book = load_workbook(target)
    try:
        sheet = book["Data"]
        assert sheet.protection.sheet
        assert sheet.protection.algorithmName == "SHA-512"
        sheet.protection.formatCells = False
        book.save(target)
    finally:
        book.close()

    records = _records(target)
    iso = next(payload for record_id, payload in records if record_id == 0x2A6)
    flags = struct.unpack_from("<16I", iso, 4)
    # objects, scenarios, formatCells: formatCells is no longer allowed.
    assert flags[2] == 1
    assert len(iso) == len(source_iso)
    assert 0x217 in {record_id for record_id, _ in records}

    reopened = load_workbook(target)
    try:
        assert reopened["Data"].protection.sheet
        assert reopened["Data"].protection.algorithmName == "SHA-512"
    finally:
        reopened.close()


def test_spin_count_change_is_saved(tmp_path):
    target = tmp_path / "spin.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        book["Data"].protection.spinCount = 12345
        book.save(target)
    finally:
        book.close()
    iso = next(payload for record_id, payload in _records(target) if record_id == 0x2A6)
    assert struct.unpack_from("<I", iso, 0)[0] == 12345
    reopened = load_workbook(target)
    try:
        assert reopened["Data"].protection.spinCount == 12345
    finally:
        reopened.close()


def test_hash_algorithm_change_is_refused(tmp_path):
    target = tmp_path / "algorithm.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        book["Data"].protection.algorithmName = "SHA-1"
        with pytest.raises(UnsupportedFeatureError, match="hash algorithm"):
            book.save(target)
    finally:
        book.close()


def test_unrelated_edit_leaves_the_iso_record_byte_for_byte(tmp_path):
    target = tmp_path / "width.xlsb"
    shutil.copyfile(FIXTURE, target)
    source_iso = next(payload for record_id, payload in _records(target) if record_id == 0x2A6)

    book = load_workbook(target)
    try:
        book["Data"].column_dimensions["A"].width = 20
        book.save(target)
    finally:
        book.close()

    iso = next(payload for record_id, payload in _records(target) if record_id == 0x2A6)
    assert iso == source_iso


@pytest.mark.com
def test_excel_still_accepts_the_password(tmp_path, excel):
    target = tmp_path / "protection.xlsb"
    shutil.copyfile(FIXTURE, target)
    book = load_workbook(target)
    try:
        book["Data"].protection.formatCells = False
        book.save(target)
    finally:
        book.close()

    workbook = excel.Workbooks.Open(str(target))
    try:
        sheet = workbook.Worksheets("Data")
        assert sheet.ProtectContents
        sheet.Unprotect("secret")
        assert not sheet.ProtectContents
    finally:
        workbook.Close(False)
