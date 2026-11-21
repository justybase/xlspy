"""Hyperlink objects validate targets and the optional display strings."""

from __future__ import annotations

import pytest

from xlspy import Hyperlink
from xlspy.hyperlinks import MAX_TARGET_LENGTH, MAX_TOOLTIP_LENGTH


def test_target_forms():
    external = Hyperlink("https://example.com/report", text="Report", tooltip="Open")
    assert not external.is_internal
    assert external.display_text == "Report"
    internal = Hyperlink("#'Data Set'!A1")
    assert internal.is_internal
    assert internal.location == "'Data Set'!A1"
    assert internal.display_text == "#'Data Set'!A1"


@pytest.mark.parametrize("target", ["", None, 1])
def test_target_must_be_a_non_empty_string(target):
    with pytest.raises(TypeError):
        Hyperlink(target)


def test_target_length_and_characters():
    with pytest.raises(ValueError):
        Hyperlink("https://example.com/" + "x" * MAX_TARGET_LENGTH)
    with pytest.raises(ValueError):
        Hyperlink(" https://example.com")
    with pytest.raises(ValueError):
        Hyperlink("https://example.com\n")


@pytest.mark.parametrize(
    "target",
    [
        "file:///C:/report.xlsx",
        "example.com/report",
        "ftp.example.com/report",
    ],
)
def test_unsupported_targets_are_refused(target):
    with pytest.raises(ValueError):
        Hyperlink(target)


@pytest.mark.parametrize("target", ["#Data", "#!A1", "#Data!A0", "#Data!XFE1"])
def test_internal_locations_are_validated(target):
    with pytest.raises(ValueError):
        Hyperlink(target)


def test_text_and_tooltip_types_and_lengths():
    with pytest.raises(TypeError):
        Hyperlink("https://example.com", text=1)
    with pytest.raises(TypeError):
        Hyperlink("https://example.com", tooltip=1)
    with pytest.raises(ValueError):
        Hyperlink("https://example.com", tooltip="x" * (MAX_TOOLTIP_LENGTH + 1))
    with pytest.raises(ValueError):
        Hyperlink("https://example.com", tooltip="bad\x01tooltip")
