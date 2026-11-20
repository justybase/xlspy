"""The totals-row cell decisions, independent of either writer.

The XLSX and XLSB table tests check the saved files; these tests pin how a
totals cell maps to ``(function, label, formula)``, including the choices that
Excel would reject.
"""

from __future__ import annotations

import pytest

from xlspy import CellRichText, Formula, Hyperlink, InlineFont, TextBlock
from xlspy._table_totals import (
    totals_choice,
    totals_formula,
    totals_function,
    totals_label,
)
from xlspy._table_totals_names import ILTA_CUSTOM, ILTA_NONE, ILTA_SUM


@pytest.mark.parametrize(
    "cell, function",
    [
        (None, ILTA_NONE),
        ("text", ILTA_NONE),
        ("=SUM(A1:A2)", ILTA_CUSTOM),
        (Formula("SUM(A1:A2)"), ILTA_CUSTOM),
        (CellRichText("To", TextBlock(InlineFont(b=True), "tal")), ILTA_NONE),
        (Hyperlink("https://example.com/", text="link"), ILTA_NONE),
        (3, ILTA_SUM),
        (3.5, ILTA_SUM),
        (True, ILTA_SUM),
    ],
)
def test_totals_function(cell, function):
    assert totals_function(cell) == function


@pytest.mark.parametrize(
    "cell, label",
    [
        (None, None),
        ("text", "text"),
        ("", None),
        ("=A1", None),
        (Formula("A1"), None),
        (CellRichText("a", TextBlock(InlineFont(i=True), "b")), "ab"),
        (Hyperlink("https://example.com/", text="link"), "link"),
        (Hyperlink("https://example.com/"), "https://example.com/"),
    ],
)
def test_totals_label(cell, label):
    assert totals_label(cell) == label


@pytest.mark.parametrize(
    "cell, formula",
    [
        (None, None),
        ("text", None),
        ("=A1", "A1"),
        ("=  SUM(A1:A2)", "SUM(A1:A2)"),
        (Formula("A1*2"), "A1*2"),
        (3, None),
    ],
)
def test_totals_formula(cell, formula):
    assert totals_formula(cell) == formula


@pytest.mark.parametrize(
    "cell, expected",
    [
        (None, ("none", None, None)),
        ("label", ("label", "label", None)),
        (7, ("sum", None, None)),
        ("=A1", ("custom", None, "A1")),
        (Formula("A1"), ("custom", None, "A1")),
        (CellRichText("Total"), ("label", "Total", None)),
        (Hyperlink("https://example.com/", text="link"), ("label", "link", None)),
    ],
)
def test_totals_choice_infers_the_cell(cell, expected):
    assert totals_choice(cell) == expected


def test_explicit_choices_accept_matching_cells():
    assert totals_choice("Total", "label") == ("label", "Total", None)
    assert totals_choice(None, "label") == ("label", None, None)
    assert totals_choice("", "none") == ("none", None, None)
    assert totals_choice("=A1", "custom") == ("custom", None, "A1")
    assert totals_choice(3, "average") == ("average", None, None)


@pytest.mark.parametrize(
    "cell, chosen, message",
    [
        (3, "label", "label totals cell must hold text"),
        (Formula("A1"), "label", "label totals cell must hold text"),
        (3, "none", "none totals cell must hold text or be empty"),
        ("label", "custom", "custom totals cell must hold a formula"),
        (None, "custom", "custom totals cell must hold a formula"),
        ("=A1", "sum", "sum totals cell cannot hold a formula"),
        (Formula("A1"), "average", "average totals cell cannot hold a formula"),
    ],
)
def test_explicit_choices_refuse_cells_they_cannot_describe(cell, chosen, message):
    with pytest.raises(ValueError, match=message):
        totals_choice(cell, chosen)
