"""The small A1 helpers keep openpyxl's edge-case answers."""

from __future__ import annotations

import pytest

from xlspy.utils.cell import (
    absolute_coordinate,
    cols_from_range,
    column_index_from_string,
    coordinate_from_string,
    coordinate_to_tuple,
    get_column_interval,
    get_column_letter,
    range_boundaries,
    range_to_tuple,
    rows_from_range,
)
from xlspy.utils.exceptions import CellCoordinatesException


def test_column_letter_bounds():
    assert get_column_letter(1) == "A"
    assert get_column_letter(18278) == "ZZZ"
    with pytest.raises(ValueError):
        get_column_letter(0)
    with pytest.raises(ValueError):
        get_column_letter(18279)


@pytest.mark.parametrize("value", ["", "A1", "ABCD", "A-", 1, None])
def test_column_index_rejects_non_column_names(value):
    with pytest.raises(ValueError):
        column_index_from_string(value)


def test_coordinates_and_their_failures():
    assert coordinate_from_string("$AB$12") == ("AB", 12)
    assert coordinate_to_tuple("C3") == (3, 3)
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string("A0")
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string("A")
    with pytest.raises(CellCoordinatesException):
        coordinate_to_tuple("1A")


def test_range_boundaries_for_every_endpoint_form():
    assert range_boundaries("A1:B2") == (1, 1, 2, 2)
    assert range_boundaries("$B$2") == (2, 2, 2, 2)
    assert range_boundaries("A:C") == (1, None, 3, None)
    assert range_boundaries("1:3") == (None, 1, None, 3)
    assert range_boundaries("A") == (1, None, 1, None)
    assert range_boundaries("2") == (None, 2, None, 2)
    assert range_boundaries("") == (None, None, None, None)
    with pytest.raises(ValueError):
        range_boundaries("A:1")


def test_sheet_ranges_and_cells():
    assert range_to_tuple("'A B'!A1:B2") == ("A B", (1, 1, 2, 2))
    assert range_to_tuple("Data!$A$1") == ("Data", (1, 1, 1, 1))
    with pytest.raises(ValueError):
        range_to_tuple("A1:B2")
    assert list(rows_from_range("A1:B2")) == [("A1", "B1"), ("A2", "B2")]
    assert list(cols_from_range("A1:B2")) == [("A1", "A2"), ("B1", "B2")]
    with pytest.raises(ValueError):
        list(rows_from_range("A:C"))
    with pytest.raises(ValueError):
        list(cols_from_range("1:3"))


def test_intervals_and_absolute_coordinates():
    assert get_column_interval(1, 3) == ["A", "B", "C"]
    assert get_column_interval("A", "C") == ["A", "B", "C"]
    assert absolute_coordinate("A1") == "$A$1"
    assert absolute_coordinate("A1:B2") == "$A$1:$B$2"
    assert absolute_coordinate("A") == "$A"
    with pytest.raises(ValueError):
        absolute_coordinate("A1:B2:C3")
