"""Colour scales, data bars and icon sets survive a loaded workbook's save."""

import pytest

from xlspy import Workbook, load_workbook
from xlspy.styles import (
    CellIsRule, ColorScaleRule, DataBarRule, DifferentialStyle, Font, FormatObject,
    IconSetRule,
)


def _rules():
    return [
        CellIsRule("A1", "greaterThan", ["5"],
                   dxf=DifferentialStyle(font=Font(bold=True))),
        ColorScaleRule("B1", FormatObject("min"), FormatObject("formula", "$D$1"),
                       FormatObject("max"), ["FFF8696B", "FFFFEB84", "FF63BE7B"]),
        DataBarRule("C1", FormatObject("min"), FormatObject("max"),
                    "FF638EC6", show_value=False),
        IconSetRule("D1", "3Arrows", [0, 33, 67], reverse=True, show_value=False),
    ]


def _found(sheet):
    return {
        type(rule).__name__: rule
        for item in sheet.conditional_formatting
        for rule in item.rules
    }


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_scale_rules_survive_an_unrelated_edit(tmp_path, suffix):
    source = tmp_path / f"source.{suffix}"
    workbook = Workbook()
    try:
        for row in range(1, 15):
            workbook.active.append([row, row * 2, row * 3, row * 4])
        for rule in _rules():
            workbook.active.conditional_formatting.add(rule.sqref, rule)
        workbook.save(source)
    finally:
        workbook.close()

    loaded = load_workbook(source)
    try:
        found = _found(loaded.active)
        assert {"CellIsRule", "ColorScaleRule", "DataBarRule", "IconSetRule"} <= set(found)
        scale = found["ColorScaleRule"]
        assert scale.midpoint == FormatObject("formula", "$D$1")
        bar = found["DataBarRule"]
        assert bar.minimum.type == "min"
        assert bar.show_value is False
        icons = found["IconSetRule"]
        assert icons.icon_style == "3Arrows" and icons.reverse is True
        assert list(icons.values) == [0, 33, 67]

        # A save for another reason (an unrelated cell) keeps every rule.
        loaded.active["A20"] = 1
        out = tmp_path / f"out.{suffix}"
        loaded.save(out)
    finally:
        loaded.close()

    reloaded = load_workbook(out)
    try:
        assert {"ColorScaleRule", "DataBarRule", "IconSetRule"} <= set(_found(reloaded.active))
    finally:
        reloaded.close()


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_scale_rule_can_be_changed(tmp_path, suffix):
    source = tmp_path / f"source.{suffix}"
    workbook = Workbook()
    try:
        workbook.active.append([1, 2, 3, 4])
        workbook.active.conditional_formatting.add(
            "B1", ColorScaleRule("B1", FormatObject("min"), None,
                                 FormatObject("max"), ["FFF8696B", "FF63BE7B"])
        )
        workbook.save(source)
    finally:
        workbook.close()

    loaded = load_workbook(source)
    try:
        rules = [rule for item in loaded.active.conditional_formatting for rule in item.rules]
        assert [type(rule).__name__ for rule in rules] == ["ColorScaleRule"]
        loaded.active.conditional_formatting.remove("B1")
        loaded.active.conditional_formatting.add(
            "B1", ColorScaleRule("B1", FormatObject("min"), None,
                                 FormatObject("max"), ["FF00FF00", "FF63BE7B"])
        )
        out = tmp_path / f"out.{suffix}"
        loaded.save(out)
    finally:
        loaded.close()

    reloaded = load_workbook(out)
    try:
        rules = [rule for item in reloaded.active.conditional_formatting for rule in item.rules]
        assert rules[0].colors[0].rgb == "FF00FF00"
    finally:
        reloaded.close()
