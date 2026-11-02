"""CellRange, MultiCellRange and the row/column dimension holders, checked against
openpyxl where the API is meant to match it."""
from __future__ import annotations

import pytest
from openpyxl.worksheet.cell_range import CellRange as OpenpyxlCellRange
from openpyxl.worksheet.cell_range import MultiCellRange as OpenpyxlMultiCellRange

from xlspy import CellRange, MultiCellRange, Workbook

RANGES = ["B2:D5", "A1", "C3:C3", "$B$2:$D$5", "XFA1048570:XFD1048576"]


@pytest.mark.parametrize("ref", RANGES)
def test_bounds_and_size_match_openpyxl(ref: str) -> None:
    ours, theirs = CellRange(ref), OpenpyxlCellRange(ref.replace("$", ""))
    assert ours.bounds == theirs.bounds
    assert ours.size == theirs.size
    assert (ours.min_col, ours.min_row, ours.max_col, ours.max_row) == theirs.bounds


def test_a_single_cell_is_normalized_to_a_range() -> None:
    assert str(CellRange("$a$1")) == "A1:A1"
    assert CellRange("B2:D5").coord == "B2:D5"


@pytest.mark.parametrize(("cell", "inside"), [("B2", True), ("D5", True), ("C4", True), ("A1", False),
                                              ("E5", False), ("D6", False), ("not a cell", False)])
def test_contains_a_cell(cell: str, inside: bool) -> None:
    assert (cell in CellRange("B2:D5")) is inside
    assert 5 not in CellRange("B2:D5")


def test_shift_expand_and_shrink_match_openpyxl() -> None:
    theirs = OpenpyxlCellRange("B2:D5")
    theirs.shift(col_shift=2, row_shift=1)
    assert CellRange("B2:D5").shift(col_shift=2, row_shift=1).bounds == theirs.bounds
    theirs = OpenpyxlCellRange("B2:D5")
    theirs.expand(right=1, down=2, left=1, up=1)
    assert CellRange("B2:D5").expand(right=1, down=2, left=1, up=1).bounds == theirs.bounds
    theirs = OpenpyxlCellRange("B2:D5")
    theirs.shrink(right=1, bottom=1, left=1, top=1)
    assert CellRange("B2:D5").shrink(right=1, bottom=1, left=1, top=1).bounds == theirs.bounds


def test_moves_off_the_sheet_are_refused() -> None:
    with pytest.raises(ValueError, match="Invalid shift"):
        CellRange("A1:B2").shift(col_shift=-1)
    with pytest.raises(ValueError, match="outside the worksheet"):
        CellRange("A1:B2").expand(up=1)
    with pytest.raises(ValueError, match="outside the worksheet"):
        CellRange("XFC1:XFD1").expand(right=1)


@pytest.mark.parametrize(("left", "right"), [("B2:D5", "C4:F8"), ("A1:B2", "D4:E5"), ("B2:D5", "C3:C4"),
                                             ("C3:C4", "B2:D5"), ("A1:A10", "A5:C5")])
def test_set_operations_match_openpyxl(left: str, right: str) -> None:
    ours, theirs = CellRange(left), OpenpyxlCellRange(left)
    other = OpenpyxlCellRange(right)
    assert ours.isdisjoint(right) is theirs.isdisjoint(other)
    assert ours.issubset(right) is theirs.issubset(other)
    assert ours.issuperset(CellRange(right)) is theirs.issuperset(other)
    assert ours.union(right).bounds == theirs.union(other).bounds
    if theirs.isdisjoint(other):
        with pytest.raises(ValueError, match="doesn't intersect"):
            ours.intersection(right)
    else:
        assert ours.intersection(right).bounds == theirs.intersection(other).bounds


def test_ranges_must_be_strings() -> None:
    with pytest.raises(TypeError):
        CellRange(5)  # type: ignore[arg-type]


def test_multi_cell_range_adds_once_and_removes() -> None:
    ranges = MultiCellRange(["A1:B2", "D4", "a1:b2"])
    assert len(ranges) == 2 and bool(ranges)
    assert str(ranges) == "A1:B2 D4:D4"
    assert repr(ranges) == "MultiCellRange(['A1:B2', 'D4:D4'])"
    assert [str(r) for r in ranges] == ["A1:B2", "D4:D4"]
    ranges.remove("D4")
    assert str(ranges) == "A1:B2"
    with pytest.raises(ValueError, match="not merged"):
        ranges.remove("C3")
    assert not MultiCellRange()


@pytest.mark.parametrize("item", ["B2", "A1:B2", "C3", "C3:D4", 7, "bad range:"])
def test_multi_cell_range_membership_matches_openpyxl(item: object) -> None:
    ours = MultiCellRange(["A1:B2", "D4:E5"])
    theirs = OpenpyxlMultiCellRange("A1:B2 D4:E5")
    expected = item in theirs if isinstance(item, str) and item != "bad range:" else False
    assert (item in ours) is expected


def test_dimensions_report_custom_sizes_and_outline_alias() -> None:
    sheet = Workbook().active
    column = sheet.column_dimensions["B"]
    assert column.customWidth is False
    column.width = 20
    assert column.customWidth is True
    row = sheet.row_dimensions[3]
    assert row.customHeight is False
    row.height = 30
    assert row.customHeight is True
    row.outline_level = 2
    assert row.outlineLevel == 2 and row.outline_level == 2


def test_column_holder_accepts_letters_and_numbers() -> None:
    columns = Workbook().active.column_dimensions
    columns[3].width = 12
    assert columns["C"].width == 12
    assert "C" in columns and 3 in columns
    assert columns.get("c") is columns["C"]
    assert columns.get("Z") is None
    with pytest.raises(ValueError, match="1..16384"):
        columns[0]
    with pytest.raises(TypeError, match="letter or 1-based int"):
        columns[1.5]


def test_grouping_columns_and_rows_as_openpyxl_does() -> None:
    sheet = Workbook().active
    sheet.column_dimensions.group("B", "D", outline_level=2, hidden=True)
    for letter in "BCD":
        dimension = sheet.column_dimensions[letter]
        assert (dimension.outlineLevel, dimension.hidden, dimension.min, dimension.max) == (2, True, 2, 4)
    sheet.column_dimensions.group("F")
    assert sheet.column_dimensions["F"].outlineLevel == 1
    sheet.row_dimensions.group(2, 4, hidden=True)
    assert all(sheet.row_dimensions[row].hidden for row in (2, 3, 4))
    assert sheet.row_dimensions[5].hidden is False
