"""Tokenizer cases where openpyxl either refuses the formula or differs.

``test_compat_utils`` compares the common formulas to openpyxl. These cases pin
xlspy's documented fallback: a formula the reference tokenizer rejects is still
split into useful tokens instead of raising.
"""

from __future__ import annotations

from xlspy.formula import Tokenizer
from xlspy.formula.token import Token


def _tokens(formula):
    return [(token.value, token.type, token.subtype) for token in Tokenizer(formula).items]


def test_an_empty_formula_has_no_tokens():
    assert _tokens("") == []


def test_a_formula_without_an_equals_sign_is_a_literal():
    assert _tokens("A1") == [("A1", Token.LITERAL, "")]


def test_escaped_quotes_inside_a_sheet_name_stay_in_the_range():
    assert _tokens("='My ''Sheet''!A1") == [("'My ''Sheet''!A1", Token.OPERAND, Token.RANGE)]


def test_an_unknown_error_code_is_a_range_operand():
    assert _tokens("=#FOO") == [("#FOO", Token.OPERAND, Token.RANGE)]


def test_an_unmatched_closing_parenthesis_is_kept():
    assert _tokens("=A1)") == [
        ("A1", Token.OPERAND, Token.RANGE),
        (")", Token.PAREN, Token.CLOSE),
    ]


def test_empty_parentheses_are_two_parenthesis_tokens():
    assert _tokens("=()") == [
        ("(", Token.PAREN, Token.OPEN),
        (")", Token.PAREN, Token.CLOSE),
    ]


def test_a_prefix_operator_after_an_open_parenthesis():
    assert _tokens("=(-A1)") == [
        ("(", Token.PAREN, Token.OPEN),
        ("-", Token.OP_PRE, ""),
        ("A1", Token.OPERAND, Token.RANGE),
        (")", Token.PAREN, Token.CLOSE),
    ]


def test_a_prefix_operator_after_a_separator():
    assert _tokens("=SUM(,-1)") == [
        ("SUM(", Token.FUNC, Token.OPEN),
        (",", Token.SEP, Token.ARG),
        ("-", Token.OP_PRE, ""),
        ("1", Token.OPERAND, Token.NUMBER),
        (")", Token.FUNC, Token.CLOSE),
    ]
