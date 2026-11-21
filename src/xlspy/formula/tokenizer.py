"""A practical A1 formula tokenizer with openpyxl-compatible tokens."""

from __future__ import annotations

import re

from .token import Token

_ERROR_CODES = (
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
)
_NUMBER_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$")
_MANTISSA_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)[Ee]")
_OPERATORS = "+-*/^&=><"


class TokenizerError(Exception):
    """Raised when a formula cannot be tokenized."""


class Tokenizer:
    """Tokenize an Excel formula in A1 notation.

    The token type/subtype names match ``openpyxl.formula.tokenizer.Token``;
    the scanner covers literals, ranges, numbers, text, functions, operators,
    arrays and separators.
    """

    def __init__(self, formula: str):
        self.formula = formula or ""
        self.items: list[Token] = []
        self._stack: list[str] = []
        self._parse()

    def _parse(self) -> None:
        formula = self.formula
        if not formula:
            return
        if not formula.startswith("="):
            self.items.append(Token(formula, Token.LITERAL))
            return

        self._index = 1
        self._length = len(formula)
        self._buffer: list[str] = []
        self._previous: Token | None = None

        while self._index < self._length:
            if self._scan(formula):
                continue
            self._buffer.append(formula[self._index])
            self._index += 1

        self._flush()

    def _scan(self, formula: str) -> bool:
        char = formula[self._index]
        if char in " \n\r\t":
            self._whitespace(formula)
        elif char == '"':
            self._text(formula)
        elif char == "'":
            self._quoted_sheet(formula)
        elif char == "#":
            self._error(formula)
        elif char == "[":
            self._brackets(formula)
        elif char in "({":
            self._open(char)
        elif char in ")}":
            self._close(char)
        elif char in ",;":
            self._separator(char)
        elif self._exponent_sign(formula):
            pass
        elif char in _OPERATORS or char == "%":
            self._operator(char)
        else:
            return False
        return True

    def _emit(self, token: Token) -> None:
        self.items.append(token)
        self._previous = token

    def _flush(self) -> None:
        value = "".join(self._buffer)
        self._buffer.clear()
        if not value:
            return
        upper = value.upper()
        if upper in {"TRUE", "FALSE"}:
            token = Token(value, Token.OPERAND, Token.LOGICAL)
        elif _NUMBER_RE.match(value):
            token = Token(value, Token.OPERAND, Token.NUMBER)
        else:
            token = Token(value, Token.OPERAND, Token.RANGE)
        self._emit(token)

    def _whitespace(self, formula: str) -> None:
        self._flush()
        start = self._index
        while self._index < self._length and formula[self._index] in " \n\r\t":
            self._index += 1
        self._emit(Token(formula[start : self._index], Token.WSPACE))

    def _text(self, formula: str) -> None:
        self._flush()
        start = self._index
        self._index += 1
        while self._index < self._length:
            if formula[self._index] == '"':
                if self._index + 1 < self._length and formula[self._index + 1] == '"':
                    self._index += 2
                    continue
                self._index += 1
                break
            self._index += 1
        self._emit(Token(formula[start : self._index], Token.OPERAND, Token.TEXT))

    def _quoted_sheet(self, formula: str) -> None:
        self._flush()
        start = self._index
        self._index += 1
        while self._index < self._length:
            if formula[self._index] == "'":
                if self._index + 1 < self._length and formula[self._index + 1] == "'":
                    self._index += 2
                    continue
                self._index += 1
                break
            self._index += 1
        self._buffer.append(formula[start : self._index])

    def _error(self, formula: str) -> None:
        self._flush()
        for code in _ERROR_CODES:
            if formula.startswith(code, self._index):
                self._emit(Token(code, Token.OPERAND, Token.ERROR))
                self._index += len(code)
                break
        else:
            self._buffer.append("#")
            self._index += 1

    def _brackets(self, formula: str) -> None:
        start = self._index
        depth = 0
        while self._index < self._length:
            if formula[self._index] == "[":
                depth += 1
            elif formula[self._index] == "]":
                depth -= 1
                if depth == 0:
                    self._index += 1
                    break
            self._index += 1
        self._buffer.append(formula[start : self._index])

    def _open(self, char: str) -> None:
        self._flush()
        if char == "(":
            if self.items and self.items[-1].type == Token.OPERAND:
                name = self.items.pop()
                token = Token(name.value + "(", Token.FUNC, Token.OPEN)
                self._stack.append(Token.FUNC)
            else:
                token = Token("(", Token.PAREN, Token.OPEN)
                self._stack.append(Token.PAREN)
        else:
            token = Token("{", Token.ARRAY, Token.OPEN)
            self._stack.append(Token.ARRAY)
        self._emit(token)
        self._index += 1

    def _close(self, char: str) -> None:
        self._flush()
        kind = self._stack.pop() if self._stack else Token.PAREN
        if char == "}":
            token = Token("}", Token.ARRAY, Token.CLOSE)
        elif kind == Token.FUNC:
            token = Token(")", Token.FUNC, Token.CLOSE)
        else:
            token = Token(")", Token.PAREN, Token.CLOSE)
        self._emit(token)
        self._index += 1

    def _separator(self, char: str) -> None:
        self._flush()
        self._emit(Token(char, Token.SEP, Token.ARG if char == "," else Token.ROW))
        self._index += 1

    def _exponent_sign(self, formula: str) -> bool:
        # The sign of an exponent (1.5E+2), not an operator after a name (RATE+1).
        if (
            formula[self._index] in "+-"
            and _MANTISSA_RE.fullmatch("".join(self._buffer))
            and self._index + 1 < self._length
            and formula[self._index + 1].isdigit()
        ):
            self._buffer.append(formula[self._index])
            self._index += 1
            return True
        return False

    def _operator(self, char: str) -> None:
        self._flush()
        previous = self._previous
        if char == "%":
            token = Token(char, Token.OP_POST)
        elif (
            previous is None
            or previous.type in {Token.OP_IN, Token.OP_PRE, Token.SEP}
            or (previous.type in {Token.PAREN, Token.FUNC} and previous.subtype == Token.OPEN)
        ):
            token = Token(char, Token.OP_PRE)
        else:
            token = Token(char, Token.OP_IN)
        self._emit(token)
        self._index += 1
