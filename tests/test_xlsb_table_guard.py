"""The removal guard for a loaded XLSB table: a formula's PtgList refers to the table whatever its ixti is."""

import struct

import pytest

from xlspy.workbook._xlsb_tables import _is_referred_to


class _FakePackage:
    def __init__(self, parts):
        self._parts = parts
        self.names = list(parts)

    def get(self, name):
        return self._parts.get(name)


def _ptg_list(ixti, flags, ident):
    return b"\x18\x19" + struct.pack("<HHIHH", ixti, flags, ident, 0, 0)


@pytest.mark.parametrize("ixti", [0, 1, 3, 0xFFFF])
def test_a_formula_on_any_sheet_refers_to_the_table_whatever_its_extern_sheet_index(ixti):
    package = _FakePackage({"xl/worksheets/sheet2.bin": b"\x00" * 4 + _ptg_list(ixti, 0x0001, 7) + b"\x00"})
    assert _is_referred_to(package, 7)


def test_the_workbook_part_is_searched_too():
    package = _FakePackage({"xl/workbook.bin": _ptg_list(2, 0x0021, 9)})
    assert _is_referred_to(package, 9)


def test_another_table_is_not_referred_to():
    package = _FakePackage({"xl/worksheets/sheet1.bin": _ptg_list(1, 0x0001, 7)})
    assert not _is_referred_to(package, 8)


def test_the_bytes_of_a_ptg_list_must_be_in_their_places():
    # The table id alone, without the PtgExtend, PtgList pair before it, is not a reference.
    package = _FakePackage({"xl/worksheets/sheet1.bin": b"\x18\x18" + struct.pack("<HHI", 1, 1, 7) + b"\x00\x00"})
    assert not _is_referred_to(package, 7)


def test_parts_other_than_the_sheets_and_the_workbook_are_not_searched():
    package = _FakePackage({
        "xl/tables/table1.bin": _ptg_list(1, 0x0001, 7),
        "xl/worksheets/sheet1.xml": _ptg_list(1, 0x0001, 7),
    })
    assert not _is_referred_to(package, 7)
