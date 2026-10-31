"""A practical A1 formula tokenizer with openpyxl-compatible tokens."""

from __future__ import annotations

import re

from .token import Token

_ERROR_CODES = (
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?",
    "#NUM!", "#N/A", "#GETTING_DATA",
)
_NUMBER_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$")
_MANTISSA_RE = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)[Ee]")
_OPERATORS = "+-*/^&=><"
_OPERAND_END = set(',;}) +-*/^&=><%(')


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

        index = 1
        length = len(formula)
        buffer: list[str] = []
        previous: Token | None = None

        def flush() -> None:
            nonlocal previous
            value = "".join(buffer)
            buffer.clear()
            if not value:
                return
            upper = value.upper()
            if upper in {"TRUE", "FALSE"}:
                token = Token(value, Token.OPERAND, Token.LOGICAL)
            elif _NUMBER_RE.match(value):
                token = Token(value, Token.OPERAND, Token.NUMBER)
            else:
                token = Token(value, Token.OPERAND, Token.RANGE)
            self.items.append(token)
            previous = token

        while index < length:
            char = formula[index]
            if char in " \n\r\t":
                flush()
                start = index
                while index < length and formula[index] in " \n\r\t":
                    index += 1
                token = Token(formula[start:index], Token.WSPACE)
                self.items.append(token)
                previous = token
                continue

            if char == '"':
                flush()
                start = index
                index += 1
                while index < length:
                    if formula[index] == '"':
                        if index + 1 < length and formula[index + 1] == '"':
                            index += 2
                            continue
                        index += 1
                        break
                    index += 1
                token = Token(formula[start:index], Token.OPERAND, Token.TEXT)
                self.items.append(token)
                previous = token
                continue

            if char == "'":
                flush()
                start = index
                index += 1
                while index < length:
                    if formula[index] == "'":
                        if index + 1 < length and formula[index + 1] == "'":
                            index += 2
                            continue
                        index += 1
                        break
                    index += 1
                buffer.append(formula[start:index])
                continue

            if char == "#":
                flush()
                for code in _ERROR_CODES:
                    if formula.startswith(code, index):
                        token = Token(code, Token.OPERAND, Token.ERROR)
                        self.items.append(token)
                        previous = token
                        index += len(code)
                        break
                else:
                    buffer.append(char)
                    index += 1
                continue

            if char == "[":
                start = index
                depth = 0
                while index < length:
                    if formula[index] == "[":
                        depth += 1
                    elif formula[index] == "]":
                        depth -= 1
                        if depth == 0:
                            index += 1
                            break
                    index += 1
                buffer.append(formula[start:index])
                continue

            if char in "({":
                flush()
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
                self.items.append(token)
                previous = token
                index += 1
                continue

            if char in ")}":
                flush()
                kind = self._stack.pop() if self._stack else Token.PAREN
                if char == "}":
                    token = Token("}", Token.ARRAY, Token.CLOSE)
                elif kind == Token.FUNC:
                    token = Token(")", Token.FUNC, Token.CLOSE)
                else:
                    token = Token(")", Token.PAREN, Token.CLOSE)
                self.items.append(token)
                previous = token
                index += 1
                continue

            if char in ",;":
                flush()
                subtype = Token.ARG if char == "," else Token.ROW
                token = Token(char, Token.SEP, subtype)
                self.items.append(token)
                previous = token
                index += 1
                continue

            # The sign of an exponent (1.5E+2), not an operator after a name (RATE+1).
            if char in "+-" and _MANTISSA_RE.fullmatch("".join(buffer)) \
                    and index + 1 < length and formula[index + 1].isdigit():
                buffer.append(char)
                index += 1
                continue

            if char in _OPERATORS or char == "%":
                flush()
                if char == "%":
                    token = Token(char, Token.OP_POST)
                elif previous is None or previous.type in {
                    Token.OP_IN, Token.OP_PRE, Token.SEP
                } or (previous.type in {Token.PAREN, Token.FUNC}
                      and previous.subtype == Token.OPEN):
                    token = Token(char, Token.OP_PRE)
                else:
                    token = Token(char, Token.OP_IN)
                self.items.append(token)
                previous = token
                index += 1
                continue

            buffer.append(char)
            index += 1

        flush()
