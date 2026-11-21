"""The sort-state rules both formats share, one validation at a time.

``check_sort_state`` is exercised through the writers elsewhere; these tests
call it directly so every refusal (and the valid edge cases) is pinned to its
own message without building a workbook.
"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from xlspy._sort_state import ICON_SET_SIZES, check_sort_state
from xlspy.exceptions import UnsupportedFeatureError


def _condition(**values):
    data = {"ref": "B2:B3"}
    data.update(values)
    return SimpleNamespace(**data)


def _state(**values):
    data = {"ref": "A1:C9", "sortCondition": [_condition()]}
    data.update(values)
    return SimpleNamespace(**data)


def test_a_valid_sort_state_returns_its_conditions():
    condition = _condition(descending=True, sortBy="value", customList="a,b")
    state = _state(
        sortCondition=[condition],
        columnSort=True,
        caseSensitive=False,
        sortMethod="stroke",
    )
    assert check_sort_state(state, "A1:C9") == [condition]


def test_a_sort_state_without_a_ref_is_accepted():
    condition = _condition()
    state = SimpleNamespace(ref=None, sortCondition=[condition])
    assert check_sort_state(state) == [condition]


@pytest.mark.parametrize(
    "state, message",
    [
        (SimpleNamespace(sortCondition=None, ref="A1"), "sort_state must be"),
        (SimpleNamespace(sortCondition=[]), "sort_state must be"),
        (_state(sortCondition="x"), "list of sort conditions"),
        (_state(ref=3), "range must be an A1 range string"),
        (_state(columnSort="yes"), "columnSort must be a bool"),
        (_state(caseSensitive=1), "caseSensitive must be a bool"),
        (_state(sortMethod=3), "sortMethod must be a string"),
        (_state(sortMethod="bogus"), "invalid sortMethod"),
    ],
)
def test_sort_state_fields_are_checked(state, message):
    with pytest.raises((TypeError, ValueError), match=message):
        check_sort_state(state)


@pytest.mark.parametrize(
    "condition, message",
    [
        (SimpleNamespace(ref=None), "needs a range"),
        (SimpleNamespace(ref=3), "range must be an A1 range string"),
        (SimpleNamespace(ref="not-a-range"), "invalid A1 range"),
        (_condition(descending=1), "descending must be a bool"),
        (_condition(sortBy=3), "sortBy must be a string"),
        (_condition(sortBy="bogus"), "invalid sortBy"),
        (_condition(customList=3), "customList must be a string"),
        (_condition(dxfId=-1), "dxfId must be at least"),
        (_condition(dxfId=True), "dxfId must be an int"),
    ],
)
def test_sort_conditions_are_checked(condition, message):
    with pytest.raises((TypeError, ValueError), match=message):
        check_sort_state(_state(sortCondition=[condition]))


def test_sorting_by_icon_needs_a_known_set_and_a_valid_index():
    with pytest.raises(UnsupportedFeatureError, match="needs an iconSet"):
        check_sort_state(_state(sortCondition=[_condition(sortBy="icon")]))
    with pytest.raises(UnsupportedFeatureError, match="unsupported icon set"):
        check_sort_state(_state(sortCondition=[_condition(sortBy="icon", iconSet="bogus", iconId=0)]))
    with pytest.raises(ValueError, match="invalid iconId"):
        check_sort_state(_state(sortCondition=[_condition(sortBy="icon", iconSet="3Arrows", iconId=3)]))
    with pytest.raises(ValueError, match="invalid iconId"):
        check_sort_state(_state(sortCondition=[_condition(sortBy="icon", iconSet="3Arrows", iconId="1")]))


@pytest.mark.parametrize("icon_set, size", sorted(ICON_SET_SIZES.items())[:3])
def test_every_icon_of_a_set_is_accepted(icon_set, size):
    condition = _condition(sortBy="icon", iconSet=icon_set, iconId=size - 1)
    assert check_sort_state(_state(sortCondition=[condition])) == [condition]


def test_a_filter_range_is_validated_even_without_a_state_range():
    condition = _condition()
    state = SimpleNamespace(ref=None, sortCondition=[condition])
    assert check_sort_state(state, "A2:C10") == [condition]
    with pytest.raises(ValueError, match="invalid A1 range"):
        check_sort_state(state, "not-a-range")
