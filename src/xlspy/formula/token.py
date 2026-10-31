"""Formula token model compatible with openpyxl's Tokenizer tokens."""

from __future__ import annotations


class Token:
    LITERAL = "LITERAL"
    OPERAND = "OPERAND"
    FUNC = "FUNC"
    PAREN = "PAREN"
    SEP = "SEP"
    OP_IN = "OPERATOR-INFIX"
    OP_PRE = "OPERATOR-PREFIX"
    OP_POST = "OPERATOR-POSTFIX"
    WSPACE = "WHITE-SPACE"

    RANGE = "RANGE"
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    LOGICAL = "LOGICAL"
    ERROR = "ERROR"
    ARRAY = "ARRAY"

    OPEN = "OPEN"
    CLOSE = "CLOSE"
    ARG = "ARG"
    ROW = "ROW"

    def __init__(self, value, type, subtype=""):
        self.value = value
        self.type = type
        self.subtype = subtype

    def __repr__(self):
        return f"<{self.type} {self.subtype}: {self.value!r}>"

    def __eq__(self, other):
        if not isinstance(other, Token):
            return NotImplemented
        return (self.value, self.type, self.subtype) == (
            other.value, other.type, other.subtype
        )

    def __hash__(self):
        return hash((self.value, self.type, self.subtype))
