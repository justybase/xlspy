"""Property-based round trips of structured references (PtgList) through their tokens.

The example suites cover the forms Excel writes. Here Hypothesis builds table and
column names that need escapes, and every allowed combination of row keywords, then
checks that text -> PtgList -> text keeps the text and that a second pass keeps the
same token bytes.
"""

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from xlspy._xlsb_formulas import (
    FormulaContext, compile_formula, decode_formula, read_parsed, structured_reference,
)

SMOKE = settings(max_examples=150, deadline=None, derandomize=True)

_TABLE = st.from_regex(r"[A-Za-z_][A-Za-z0-9_.]{0,7}", fullmatch=True)
_COLUMN = st.text(st.characters(blacklist_categories=("Cs", "Cc")), min_size=1, max_size=6)
_COLUMNS = st.lists(_COLUMN, min_size=1, max_size=4, unique_by=str.casefold)

# The row keyword combinations Excel accepts ([MS-XLSB] 2.5.97.52 and Excel itself).
_ROW_KEYWORDS = [
    (),
    ("#Headers",),
    ("#Data",),
    ("#Totals",),
    ("#Headers", "#Data"),
    ("#Data", "#Totals"),
    ("#All",),
    ("#This Row",),
]
_KEYWORD_BITS = {"#ALL": 0x04, "#HEADERS": 0x08, "#DATA": 0x10, "#TOTALS": 0x20, "#THIS ROW": 0x40}


@st.composite
def _references(draw):
    table = draw(_TABLE)
    columns = draw(_COLUMNS)
    keywords = draw(st.sampled_from(_ROW_KEYWORDS))
    # A column range needs two columns; a one-column table can only take a single column.
    mode = draw(st.sampled_from((0, 1) if len(columns) == 1 else (0, 1, 2)))
    first = last = 0
    if mode == 1:
        first = last = draw(st.integers(min_value=0, max_value=len(columns) - 1))
    elif mode == 2:
        first = draw(st.integers(min_value=0, max_value=len(columns) - 2))
        last = draw(st.integers(min_value=first + 1, max_value=len(columns) - 1))
    flags = mode | sum(_KEYWORD_BITS[keyword.upper()] for keyword in keywords)
    return table, columns, flags, first, last


def _context(table, columns):
    return FormulaContext(["Data"], tables=lambda: {7: (table, columns, "Data")})


@SMOKE
@given(reference=_references())
def test_a_reference_written_from_its_flags_round_trips(reference):
    table, columns, flags, first, last = reference
    text = structured_reference(table, columns, flags, first, last)
    context = _context(table, columns)
    compiled = compile_formula("=" + text, context, sheet="Data")
    assert decode_formula(compiled, context, sheet="Data") == text
    again = compile_formula("=" + decode_formula(compiled, context, sheet="Data"), context, sheet="Data")
    assert read_parsed(again)[:2] == read_parsed(compiled)[:2]


@SMOKE
@given(reference=_references())
def test_the_name_of_the_table_and_its_columns_are_case_insensitive(reference):
    table, columns, flags, first, last = reference
    text = structured_reference(table, columns, flags, first, last)
    context = _context(table.upper() if table.isalpha() else table.swapcase(), columns)
    compiled = compile_formula("=" + text, context, sheet="Data")
    # The token names the table by id, so its text comes back with the catalog's spelling.
    assert decode_formula(compiled, context, sheet="Data").casefold() == text.casefold()


@pytest.mark.parametrize(("keywords", "flags"), [
    (("#Headers",), 0x0008),
    (("#Data",), 0x0010),
    (("#Totals",), 0x0020),
    (("#Headers", "#Data"), 0x0018),
    (("#Data", "#Totals"), 0x0030),
    (("#All",), 0x0004),
    (("#This Row",), 0x0040),
])
def test_the_keyword_sets_excel_accepts_are_written_and_read(keywords, flags):
    text = structured_reference("Tbl", ["qty"], flags, 0, 0)
    context = _context("Tbl", ["qty"])
    compiled = compile_formula("=" + text, context, sheet="Data")
    assert decode_formula(compiled, context, sheet="Data") == text
