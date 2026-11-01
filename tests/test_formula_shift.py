"""One rule shifts relative references, for shared formulas on read and for moved cells.

Both paths call ``refs.shift_references``; these tests pin its results to
openpyxl's ``Translator`` and to the cases that differed before the merge.
"""

import pytest
from openpyxl.formula.translate import Translator

from xlspy._xlsx_read import translate_formula as shared_formula_text
from xlspy.refs import column_label, shift_references
from xlspy.workbook._shift import translate_formula as moved_formula_text


@pytest.mark.parametrize(
    "formula,rows,cols,expected",
    [
        ("=A1+$B$2+B$3+$C4+Sheet2!D5+\"A1\"", 2, 1, "=B3+$B$2+C$3+$C6+Sheet2!E7+\"A1\""),
        ("=SUM(A1:B2)+LOG10(A1)", 1, 0, "=SUM(A2:B3)+LOG10(A2)"),
        ("=SUM(A:A)", 1, 1, "=SUM(B:B)"),
        ("=SUM($A:$A)", 1, 1, "=SUM($A:$A)"),
        ("=SUM(1:2)+A1", 1, 1, "=SUM(2:3)+B2"),
        ("=SUM($1:$2)", 1, 1, "=SUM($1:$2)"),
        ("=A1", -1, 0, "=#REF!"),
        ("=XFD1+A1", 0, 1, "=#REF!+B1"),
        ("=XFE1", 0, 0, "=XFE1"),  # a name, not a cell
        ("='My A1'!A1+B1", 1, 1, "='My A1'!B2+C2"),
        ('="A1"&B1', 1, 1, '="A1"&C2'),
        ("=Name1+A1", 1, 0, "=Name1+A2"),
    ],
)
def test_shift_references(formula, rows, cols, expected):
    assert shift_references(formula, rows, cols) == expected


@pytest.mark.parametrize(
    "formula,rows,cols",
    [
        ("=SUM(A:A)", 1, 1),
        ("=SUM(1:2)+A1", 1, 1),
        ("=$A:$A+B$3", 2, 2),
        ("='My A1'!A1+B1", 1, 1),
        ("=A1:B2*C3", 3, 1),
    ],
)
def test_matches_openpyxl_translator(formula, rows, cols):
    origin = "B2"
    destination = f"{column_label(1 + cols)}{2 + rows}"
    expected = Translator(formula, origin=origin).translate_formula(destination)
    assert shift_references(formula, rows, cols) == expected


def test_moved_cells_and_shared_formulas_share_the_rule():
    formula = "=SUM(A:A)+A1+\"A1\""
    assert moved_formula_text(formula, 1, 1) == shared_formula_text(formula, "B2", "C3")
    assert shared_formula_text(formula, "B2", "C3") == "=SUM(B:B)+B2+\"A1\""
