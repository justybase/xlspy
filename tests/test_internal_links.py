"""Hyperlinks to places in the workbook: sheet names, quoting, order."""
from __future__ import annotations

import pytest
from helpers import part_text

from xlspy import Hyperlink, XlsxWriter

openpyxl = pytest.importorskip("openpyxl")


@pytest.mark.parametrize(
    "sheet_name,location",
    [
        ("it's", "#'it''s'!B2"),
        ("O'Brien Q1", "#'o''brien q1'!A1"),
        ("Sheet1", "#sheet1!A1"),
        ("Sales Data", "#'SALES DATA'!B2"),
        ("Data", "#Data!A1:B3"),
    ],
)
def test_internal_links_resolve_quoted_and_differently_cased_names(tmp_path, sheet_name, location):
    path = tmp_path / "links.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet(sheet_name)
        writer.write_sheet([[Hyperlink(location, text="go")]])

    assert "<hyperlink " in part_text(path, "xl/worksheets/sheet1.xml")


def test_internal_link_to_a_later_sheet_is_accepted(tmp_path):
    path = tmp_path / "later.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Index")
        writer.write_sheet([[Hyperlink("#'Details 2'!A1", text="details")]])
        writer.add_sheet("Details 2")
        writer.write_sheet([["here"]])
    openpyxl.load_workbook(path)


def test_internal_link_with_apostrophe_but_wrong_name_still_fails(tmp_path):
    # Only the enclosing quotes are removed: "'its'" names "its", not "it's".
    with pytest.raises(ValueError, match="unknown sheet 'its'"):
        with XlsxWriter(tmp_path / "bad.xlsx") as writer:
            writer.add_sheet("it's")
            writer.write_sheet([[Hyperlink("#'its'!A1", text="go")]])
