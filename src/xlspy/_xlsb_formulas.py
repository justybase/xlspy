"""BIFF12 formula tokens and workbook name/reference context.

Used only for formulas and metadata; ordinary data encoders never call this
module. No Excel installation or third-party parser is needed at runtime.
"""

from __future__ import annotations

from .exceptions import UnsupportedFeatureError
from collections.abc import Iterable
from dataclasses import dataclass
import datetime
from decimal import Decimal
import math
import posixpath
import re
import struct
from typing import Any, cast

from ._metadata import DefinedName
from ._xlsb_functions import FUNCTIONS, BY_ID
from ._xlsb_metadata import wide_string, read_wide, records, replace_records
from .biff_utils import build_record
from .excel_dates import datetime_to_serial
from .formulas import _XLFN_FUNCTIONS, future_function_name
from .refs import parse_cell_range, column_label, quote_sheet_name
from ._safe_xml import fromstring as safe_fromstring

NAME, BEGIN_EXTERNALS, END_EXTERNALS, SUP_SELF, EXTERN_SHEET = 0x27, 0x161, 0x162, 0x165, 0x16A
SUP_BOOK_SRC, SUP_TABS = 0x163, 0x167
ERRORS = {
    0: "#NULL!",
    7: "#DIV/0!",
    15: "#VALUE!",
    23: "#REF!",
    29: "#NAME?",
    36: "#NUM!",
    42: "#N/A",
    43: "#GETTING_DATA",
}
_OPS = {
    3: "+",
    4: "-",
    5: "*",
    6: "/",
    7: "^",
    8: "&",
    9: "<",
    10: "<=",
    11: "=",
    12: ">=",
    13: ">",
    14: "<>",
    15: " ",
    16: ",",
    17: ":",
}
_OP_IDS = {value: key for key, value in _OPS.items()}
_PREC = {
    ",": 1,
    "=": 2,
    "<": 2,
    ">": 2,
    "<=": 2,
    ">=": 2,
    "<>": 2,
    "&": 3,
    "+": 4,
    "-": 4,
    "*": 5,
    "/": 5,
    "^": 6,
    " ": 9,
    ":": 10,
}
_ADDRESS = r"(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)"
_LEX = re.compile(
    r'\s+|"(?:[^"]|"")*"|\x27(?:[^\x27]|\x27\x27)*\x27'
    r"|#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)"
    + rf"|{_ADDRESS}(?![\w.])"
    + r"|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[^\W\d][\w.]*|\\[\w.]*|<>|<=|>=|[+\-*/^&%=<>(),!:]",
    re.UNICODE | re.I,
)


def parsed_formula(tokens: bytes, extra: bytes = b"") -> bytes:
    if not 0 < len(tokens) <= 16384:
        raise ValueError("XLSB formula token length exceeds the BIFF12 limit")
    return struct.pack("<I", len(tokens)) + tokens + struct.pack("<I", len(extra)) + extra


def read_parsed(data: bytes, offset: int = 0) -> tuple[bytes, bytes, int]:
    if offset + 4 > len(data):
        raise ValueError("Truncated XLSB formula length")
    length = struct.unpack_from("<I", data, offset)[0]
    end = offset + 4 + length
    if not 0 < length <= 16384 or end + 4 > len(data):
        raise ValueError("Invalid XLSB formula token length")
    extra_length = struct.unpack_from("<I", data, end)[0]
    if end + 4 + extra_length > len(data):
        raise ValueError("Truncated XLSB formula extra data")
    return data[offset + 4 : end], data[end + 4 : end + 4 + extra_length], end + 4 + extra_length


def _syntax_error(text: str, end: int) -> UnsupportedFeatureError:
    message = f"unsupported XLSB formula syntax near {text[end : end + 24]!r}"
    if "[" in text[max(end - 1, 0) : end + 24]:
        message += (
            " (a structured reference is written Table[Column], Table[[#Totals],[Column]] or "
            "Table[@Column]: a bare [Column] needs its table name)"
        )
    return UnsupportedFeatureError(message)


_TABLE_NAME = re.compile(r"[^\W\d][\w.]*")
# Characters a structured reference escapes with an apostrophe.
_ESCAPABLE = "'#[]@"
# The row keywords of a structured reference and their PtgList bits ([MS-XLSB] 2.5.97.52).
_KEYWORD_BITS = {"#ALL": 0x04, "#HEADERS": 0x08, "#DATA": 0x10, "#TOTALS": 0x20, "#THIS ROW": 0x40}


class _Structured:
    """A token ``Table[spec]``; ``spec`` is the text between the outer brackets."""

    def __init__(self, name: str, spec: str) -> None:
        self.name = name
        self.spec = spec

    def __repr__(self) -> str:
        return f"{self.name}[{self.spec}]"


def _scan_group(text: str, start: int) -> tuple[str, int]:
    """The text inside the brackets that open at ``start``, and the index after them.

    An escaped character (``'[``, ``''``) does not count as a bracket.
    """
    depth = 0
    index = start
    while index < len(text):
        char = text[index]
        if char == "'":
            index += 2
            continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                return text[start + 1 : index], index + 1
        index += 1
    raise _syntax_error(text, start)


class _Parser:
    def __init__(self, text: str) -> None:
        self.tokens: list[Any] = []
        position = 0
        while position < len(text):
            match = _LEX.match(text, position)
            if match is None:
                raise _syntax_error(text, position)
            token: Any = match[0]
            position = match.end()
            if token.isspace():
                continue
            if text[position : position + 1] == "[" and _TABLE_NAME.fullmatch(token):
                spec, position = _scan_group(text, position)
                token = _Structured(token, spec)
            self.tokens.append(token)
        self.tokens.append(None)
        self.index = 0

    def peek(self) -> Any:
        return self.tokens[self.index]

    def pop(self) -> Any:
        token = self.peek()
        if token is None:
            raise ValueError("unexpected end of XLSB formula")
        self.index += 1
        return token

    def expression(self, minimum: int = 0) -> Any:
        left = self._primary()
        while self.peek() == "(":
            self.pop()
            left = ("call", left, self._argument_list())
        while True:
            operator = self.peek()
            if operator == "%":
                self.pop()
                left = ("percent", left)
                continue
            precedence = _PREC.get(operator, -1)
            if precedence < minimum:
                break
            self.pop()
            right = self.expression(precedence + 1)
            left = ("binary", operator, left, right)
        return left

    def _primary(self) -> Any:
        token = self.pop()
        if isinstance(token, _Structured):
            return ("table", token.name, token.spec)
        if token in ("+", "-"):
            return ("unary", token, self.expression(7))
        if token == "(":
            left = ("paren", self.expression())
            if self.pop() != ")":
                raise ValueError("missing closing parenthesis")
            return left
        if token.startswith('"'):
            return ("string", token[1:-1].replace('""', '"'))
        if token.startswith("#"):
            return ("error", token.upper())
        return self._name_or_value(token)

    def _name_or_value(self, token: str) -> Any:
        sheet = None
        if self.peek() == "!":
            self.pop()
            sheet = token[1:-1].replace("''", "'") if token.startswith("'") else token
            token = self.pop()
        if self.peek() == "(" and sheet is None:
            self.pop()
            return ("function", token.upper(), self._argument_list())
        if re.fullmatch(_ADDRESS, token):
            return ("reference", token, sheet)
        if re.fullmatch(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?", token):
            return ("number", token)
        if token.upper() in ("TRUE", "FALSE") and sheet is None:
            return ("bool", token.upper() == "TRUE")
        return ("name", token, sheet)

    def _argument_list(self) -> list[Any]:
        args = []
        if self.peek() != ")":
            while True:
                args.append(("missing",) if self.peek() in (",", ")") else self.expression(2))
                if self.peek() != ",":
                    break
                self.pop()
        if self.pop() != ")":
            raise ValueError("invalid function argument list")
        return args


def parse_expression(text: str) -> Any:
    parser = _Parser(text.removeprefix("="))
    tree = parser.expression()
    if parser.peek() is not None:
        raise UnsupportedFeatureError(f"unsupported XLSB formula token {parser.peek()!r}")
    return tree


def _location(ref: str, *, relative_offsets: bool = False, origin: tuple[int, int] = (0, 0)) -> tuple[int, int]:
    match = re.fullmatch(r"(\$?)([A-Za-z]+)(\$?)(\d+)", ref)
    if match is None:
        raise ValueError(f"invalid cell reference {ref!r}")
    col, row, _, _ = parse_cell_range(ref)
    col_relative, row_relative = not match[1], not match[3]
    if relative_offsets:
        if col_relative:
            col = (col - origin[0]) & 0x3FFF
        if row_relative:
            row = (row - origin[1]) & 0xFFFFF
    flags = (int(col_relative) << 14) | (int(row_relative) << 15)
    return row, col | flags


# Functions Excel recalculates on every change; formulas using them start
# with PtgAttrSemi in Excel-written files.
_VOLATILE_FUNCTIONS = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"})


def _is_relative(ref: str) -> bool:
    """True when a column or row of an A1 reference lacks its ``$``."""
    for part in ref.split(":"):
        match = re.fullmatch(r"(\$?)([A-Za-z]*)(\$?)(\d*)", part)
        if match is None:
            raise ValueError(f"not an A1 reference: {ref!r}")
        if (match[2] and not match[1]) or (match[4] and not match[3]):
            return True
    return False


def encode_reference(
    ref: str,
    sheet: str | None,
    context: Any,
    operand: str,
    origin: tuple[int, int],
    name_formula: bool,
    cell_relative: bool = False,
) -> bytes:
    kind = 0x20 if operand == "R" else 0x60 if operand == "A" else 0x40
    # Validation and conditional-format formulas store same-sheet references
    # with a relative part as offsets from the range's top-left cell
    # (PtgRefN, PtgAreaN); fully absolute ones stay PtgRef/PtgArea, like Excel.
    relative_token = cell_relative and sheet is None and _is_relative(ref)
    offsets = name_formula or relative_token
    if ":" in ref:
        a, b = ref.split(":")
        if not re.search(r"\d", a):
            a, b = a + "$1", b + "$1048576"
        elif not re.search("[A-Za-z]", a):
            a, b = "$A" + a, "$XFD" + b
        r1, c1 = _location(a, relative_offsets=offsets, origin=origin)
        r2, c2 = _location(b, relative_offsets=offsets, origin=origin)
        token = bytes([kind | (0x1B if sheet is not None else 0x0D if relative_token else 5)])
        payload = struct.pack("<IIHH", r1, r2, c1, c2)
    else:
        row, col = _location(ref, relative_offsets=offsets, origin=origin)
        token = bytes([kind | (0x1A if sheet is not None else 0x0C if relative_token else 4)])
        payload = struct.pack("<IH", row, col)
    if sheet is not None:
        token += struct.pack("<H", context.sheet_reference(sheet))
    return token + payload


class _Compiler:
    """Encode one formula expression into XLSB tokens."""

    def __init__(
        self,
        context: Any,
        *,
        sheet: str | None,
        origin: tuple[int, int],
        name_formula: bool,
        cell_relative: bool,
    ) -> None:
        self.context = context
        self.sheet = sheet
        self.origin = origin
        self.name_formula = name_formula
        self.cell_relative = cell_relative
        self.references = 0
        self.volatile = False

    def compile(self, text: str, operand: str | None, stats: dict | None) -> bytes:
        tree = parse_expression(text)
        top = operand or ("R" if self.name_formula else "V")
        tokens = self.emit(tree, top)
        if self.volatile:
            # PtgAttrSemi opens every formula with a volatile function.
            tokens = b"\x19\x01\x00\x00" + tokens
        if stats is not None:
            stats["references"] = self.references
        return cast(bytes, parsed_formula(tokens))

    def table_token(self, table: str, spec: Any, operand_class: str) -> bytes:
        self.references += 1
        if self.name_formula:
            raise UnsupportedFeatureError(
                "a defined name cannot refer to a table: Excel does not accept a structured reference in a name"
            )
        return cast(bytes, encode_table(table, spec, self.context, operand_class))

    def emit(self, node: tuple, operand: str = "V") -> bytes:
        kind = node[0]
        if kind == "number":
            return self._emit_number(node[1])
        if kind == "string":
            return self._emit_string(node[1])
        if kind == "bool":
            return b"\x1d" + bytes([node[1]])
        if kind == "error":
            return self._emit_error(node[1])
        if kind == "missing":
            return b"\x16"
        if kind == "reference":
            return self._emit_reference(node, operand)
        if kind == "table":
            return self.table_token(node[1], node[2], operand)
        if kind == "name":
            return self._emit_name(node, operand)
        if kind == "paren":
            return self.emit(node[1], operand) + b"\x15"
        if kind == "percent":
            return self.emit(node[1]) + b"\x14"
        if kind == "unary":
            return self.emit(node[2]) + (b"\x12" if node[1] == "+" else b"\x13")
        if kind == "binary":
            arg_type = "R" if node[1] in (",", " ", ":") else "V"
            return self.emit(node[2], arg_type) + self.emit(node[3], arg_type) + bytes([_OP_IDS[node[1]]])
        if kind == "function":
            return self._emit_function(node, operand)
        if kind == "call":
            return self._emit_call(node, operand)
        raise UnsupportedFeatureError(f"unsupported XLSB formula node {kind}")

    @staticmethod
    def _emit_number(text: str) -> bytes:
        value = float(text)
        if not math.isfinite(value):
            raise ValueError("XLSB formula numeric constants must be finite")
        return (
            b"\x1e" + struct.pack("<H", int(value))
            if value.is_integer() and 0 <= value <= 65535
            else b"\x1f" + struct.pack("<d", value)
        )

    @staticmethod
    def _emit_string(text: str) -> bytes:
        data = text.encode("utf-16le")
        if len(data) // 2 > 255:
            raise ValueError("XLSB formula string literals are limited to 255 UTF-16 units")
        return b"\x17" + struct.pack("<H", len(data) // 2) + data

    @staticmethod
    def _emit_error(label: str) -> bytes:
        return b"\x1c" + bytes([next(code for code, name in ERRORS.items() if name == label)])

    def _emit_reference(self, node: tuple, operand: str) -> bytes:
        self.references += 1
        return encode_reference(
            node[1],
            node[2],
            self.context,
            operand,
            self.origin,
            self.name_formula,
            self.cell_relative,
        )

    def _emit_name(self, node: tuple, operand: str) -> bytes:
        if node[2] is None and self.context.has_table(node[1]):
            return self.table_token(node[1], None, operand)  # a table alone is the whole table
        if node[2] is None and node[1].startswith("_xlpm."):
            index = self.context.lambda_name(node[1])
        else:
            index = self.context.name_reference(
                node[1],
                node[2] or self.sheet,
                qualified=node[2] is not None,
            )
        return bytes([0x23 if operand == "R" else 0x63 if operand == "A" else 0x43]) + struct.pack("<I", index + 1)

    @staticmethod
    def _result_class(operand: str) -> int:
        return 0x20 if operand == "R" else 0x60 if operand == "A" else 0x40

    def _emit_function(self, node: tuple, operand: str) -> bytes:
        function, args = node[1], node[2]
        base = function.removeprefix("_XLFN.")
        info = FUNCTIONS.get(base)
        result_class = self._result_class(operand)
        if info is None:
            return self._emit_future_function(function, base, args, result_class)
        identifier, minimum, maximum, fixed, classes = info
        if base in _VOLATILE_FUNCTIONS:
            self.volatile = True
        if not minimum <= len(args) <= maximum:
            raise ValueError(f"{function} expects {minimum}..{maximum} arguments")
        encoded = [self.emit(arg, self._argument_class(identifier, index, classes)) for index, arg in enumerate(args)]
        ending = (
            bytes([result_class | 1]) + struct.pack("<H", identifier)
            if fixed is not None
            else bytes([result_class | 2, len(args)]) + struct.pack("<H", identifier)
        )
        if identifier == 1:  # IF, including its short-circuit jump records.
            return self._emit_if(encoded, ending)
        if identifier == 480:  # IFERROR
            return self._emit_iferror(encoded, ending)
        if identifier == 100:  # CHOOSE jump table and per-branch exits.
            return self._emit_choose(encoded, ending)
        return b"".join(encoded) + ending

    def _emit_future_function(self, function: str, base: str, args: list, result_class: int) -> bytes:
        if base not in _XLFN_FUNCTIONS:
            raise UnsupportedFeatureError(f"unsupported XLSB function {function!r}")
        index = self.context.future_function(function)
        if len(args) > 254:
            raise ValueError("too many XLSB function arguments")
        return (
            b"\x23"
            + struct.pack("<I", index + 1)
            + b"".join(self.emit(arg, "R") for arg in args)
            + bytes([result_class | 2, len(args) + 1])
            + struct.pack("<H", 255)
        )

    @staticmethod
    def _argument_class(identifier: int, index: int, classes: Any) -> str:
        if identifier == 481:
            return "R" if index % 2 == 0 else "V"
        if identifier in (482, 484):
            return "R" if index == 0 or index % 2 else "V"
        return classes[min(index, len(classes) - 1)] if classes else "V"

    @staticmethod
    def _emit_if(encoded: list[bytes], ending: bytes) -> bytes:
        yes = encoded[1]
        no = encoded[2] if len(encoded) == 3 else b""
        first_goto = b"\x19\x08" + struct.pack("<H", len(no) + (4 if no else 0) + len(ending) - 1)
        last_goto = b"\x19\x08" + struct.pack("<H", len(ending) - 1) if no else b""
        return encoded[0] + b"\x19\x02" + struct.pack("<H", len(yes) + 4) + yes + first_goto + no + last_goto + ending

    @staticmethod
    def _emit_iferror(encoded: list[bytes], ending: bytes) -> bytes:
        return (
            encoded[0]
            + b"\x19\x80"
            + struct.pack("<H", len(encoded[1]))
            + encoded[1]
            + b"\x19\x08"
            + struct.pack("<H", len(ending) - 1)
            + ending
        )

    @staticmethod
    def _emit_choose(encoded: list[bytes], ending: bytes) -> bytes:
        alternatives = encoded[1:]
        offsets = [2 * (len(alternatives) + 1)]
        for branch in alternatives:
            offsets.append(offsets[-1] + len(branch) + 4)
        chunks = [
            encoded[0],
            b"\x19\x04"
            + struct.pack("<H", len(alternatives))
            + b"".join(struct.pack("<H", offset) for offset in offsets),
        ]
        for index, branch in enumerate(alternatives):
            remaining = sum(len(value) + 4 for value in alternatives[index + 1 :]) + len(ending) - 1
            chunks.append(branch + b"\x19\x08" + struct.pack("<H", remaining))
        return b"".join(chunks) + ending

    def _emit_call(self, node: tuple, operand: str) -> bytes:
        # An immediate invocation: LAMBDA(...)(args). Excel stores the
        # callee and the call arguments, then a PtgFuncVar with id 255
        # whose first stack item is the callee.
        callee, call_args = node[1], node[2]
        result_class = self._result_class(operand)
        if len(call_args) + 1 > 254:
            raise ValueError("too many XLSB function arguments")
        encoded = [self.emit(callee, "R")] + [self.emit(arg, "R") for arg in call_args]
        return b"".join(encoded) + bytes([result_class | 2, len(call_args) + 1]) + struct.pack("<H", 255)


def compile_formula(
    text: str,
    context: Any,
    *,
    sheet: str | None = None,
    origin: tuple[int, int] = (0, 0),
    name_formula: bool = False,
    cell_relative: bool = False,
    operand: str | None = None,
    stats: dict | None = None,
) -> bytes:
    """Encode formula text as a CellParsedFormula.

    ``stats``, when given, is a dict that receives the number of reference
    tokens under ``"references"``.
    """
    if not isinstance(text, str) or not text or len(text) > 8192:
        raise ValueError("invalid XLSB formula text or length")
    compiler = _Compiler(
        context,
        sheet=sheet,
        origin=origin,
        name_formula=name_formula,
        cell_relative=cell_relative,
    )
    return compiler.compile(text, operand, stats)


def _decode_ref(row: int, column: int, origin: tuple[int, int], relative_offsets: bool) -> str:
    col = column & 0x3FFF
    col_relative, row_relative = bool(column & 0x4000), bool(column & 0x8000)
    if relative_offsets:
        if col_relative:
            col = (origin[0] + (col - 0x4000 if col & 0x2000 else col)) % 16384
        if row_relative:
            row = (origin[1] + (row - 0x100000 if row & 0x80000 else row)) % 1048576
    if row > 1048575:
        return "#REF!"
    return ("" if col_relative else "$") + column_label(col) + ("" if row_relative else "$") + str(row + 1)


class _Decoder:
    """Turn XLSB formula tokens back into formula text."""

    def __init__(
        self,
        data: bytes,
        context: Any,
        *,
        sheet: str | None = None,
        origin: tuple[int, int] = (0, 0),
        name_formula: bool = False,
    ) -> None:
        self.tokens: bytes = read_parsed(data)[0]
        self.context = context
        self.sheet = sheet
        self.origin = origin
        self.name_formula = name_formula
        self.stack: list[tuple[str, int]] = []
        self.position = 0

    def decode(self) -> str:
        while self.position < len(self.tokens):
            self._step()
        if len(self.stack) != 1:
            raise ValueError("invalid XLSB formula stack result")
        return self.stack[0][0]

    def push(self, text: str, precedence: int = 100) -> None:
        self.stack.append((text, precedence))

    def pop(self) -> tuple[str, int]:
        if not self.stack:
            raise ValueError("invalid XLSB formula operand stack")
        return self.stack.pop()

    def take(self, size: int) -> bytes:
        end = self.position + size
        if end > len(self.tokens):
            raise ValueError("truncated XLSB formula token")
        value = self.tokens[self.position : end]
        self.position = end
        return value

    def _step(self) -> None:
        token = self.take(1)[0]
        if token in _OPS:
            self._binary(token)
        elif token < 0x18:
            self._small(token)
        elif token in (0x18, 0x19):
            self._extended(token)
        elif token <= 0x1F:
            self._constant(token)
        else:
            self._reference(token)

    def _binary(self, token: int) -> None:
        right, rp = self.pop()
        left, lp = self.pop()
        op = _OPS[token]
        prec = _PREC[op]
        if lp < prec:
            left = "(" + left + ")"
        if rp < prec or (rp == prec and op in ("-", "/", "^", "<", ">", "=", "<=", ">=", "<>")):
            right = "(" + right + ")"
        self.push(left + op + right, prec)

    def _small(self, token: int) -> None:
        if token in (0x12, 0x13):
            text, prec = self.pop()
            self.push(("+" if token == 0x12 else "-") + ("(" + text + ")" if prec < 7 else text), 7)
        elif token == 0x14:
            text, prec = self.pop()
            self.push(("(" + text + ")" if prec < 8 else text) + "%", 8)
        elif token == 0x15:
            self.push("(" + self.pop()[0] + ")")
        elif token == 0x16:
            self.push("")
        elif token == 0x17:
            length = struct.unpack("<H", self.take(2))[0]
            self.push('"' + self.take(length * 2).decode("utf-16le").replace('"', '""') + '"')
        else:
            raise UnsupportedFeatureError(f"unsupported XLSB formula token 0x{token:02x}")

    def _constant(self, token: int) -> None:
        if token == 0x1C:
            self.push(ERRORS.get(self.take(1)[0], "#VALUE!"))
        elif token == 0x1D:
            self.push("TRUE" if self.take(1)[0] else "FALSE")
        elif token == 0x1E:
            self.push(str(struct.unpack("<H", self.take(2))[0]))
        elif token == 0x1F:
            self.push(format(struct.unpack("<d", self.take(8))[0], ".15g"))
        else:
            raise UnsupportedFeatureError(f"unsupported XLSB formula token 0x{token:02x}")

    def _extended(self, token: int) -> None:
        if token == 0x18:
            if self.take(1)[0] != 0x19:  # PtgList; the other extended tokens are not read
                raise UnsupportedFeatureError("unsupported XLSB formula token 0x18")
            _ixti, list_flags, ident, first, last = struct.unpack("<HHIHH", self.take(12))
            self.push(self.context.table_reference(ident, list_flags, first, last))
            return
        flags = self.take(1)[0]
        value = struct.unpack("<H", self.take(2))[0]
        if flags & 4:  # CHOOSE jump table.
            self.take(2 * (value + 1))
        elif flags & 0x10:
            self.push("SUM(" + self.pop()[0] + ")")

    def _reference(self, token: int) -> None:
        base = (token & 0x1F) | 0x20 if token >= 0x20 else token
        if base in (0x24, 0x2C, 0x3A):
            self._cell(base)
        elif base in (0x25, 0x2D, 0x3B):
            self._area(base)
        elif base == 0x23:
            self.push(self.context.name_label(struct.unpack("<I", self.take(4))[0] - 1, self.sheet))
        elif base in (0x21, 0x22):
            self._function(base)
        elif base in (0x2A, 0x2B, 0x3C, 0x3D):
            self.take((2 if base in (0x3C, 0x3D) else 0) + (6 if base in (0x2A, 0x3C) else 12))
            self.push("#REF!")
        elif base in (0x26, 0x27, 0x28):
            self.take(6)  # Memory/cache tokens do not change the expression stack.
        elif base == 0x29:
            self.take(2)
        else:
            raise UnsupportedFeatureError(f"unsupported XLSB formula token 0x{token:02x}")

    def _cell(self, base: int) -> None:
        qualifier = self.context.sheet_label(struct.unpack("<H", self.take(2))[0]) if base == 0x3A else ""
        row, col = struct.unpack("<IH", self.take(6))
        self.push(qualifier + _decode_ref(row, col, self.origin, self.name_formula or base == 0x2C))

    def _area(self, base: int) -> None:
        qualifier = self.context.sheet_label(struct.unpack("<H", self.take(2))[0]) if base == 0x3B else ""
        r1, r2, c1, c2 = struct.unpack("<IIHH", self.take(12))
        first = _decode_ref(r1, c1, self.origin, self.name_formula or base == 0x2D)
        last = _decode_ref(r2, c2, self.origin, self.name_formula or base == 0x2D)
        # Excel stores whole-column/row references as full worksheet ranges.
        if r1 == 0 and r2 == 1048575 and not (c1 & 0x8000 or c2 & 0x8000):
            first, last = re.sub(r"\$?\d+$", "", first), re.sub(r"\$?\d+$", "", last)
        elif c1 & 0x3FFF == 0 and c2 & 0x3FFF == 16383 and not (c1 & 0x4000 or c2 & 0x4000):
            first, last = re.sub(r"^\$?[A-Za-z]+", "", first), re.sub(r"^\$?[A-Za-z]+", "", last)
        self.push(qualifier + first + ":" + last)

    def _function(self, base: int) -> None:
        count = self.take(1)[0] & 0x7F if base == 0x22 else None
        identifier = struct.unpack("<H", self.take(2))[0]
        if identifier not in BY_ID and identifier != 255:
            raise UnsupportedFeatureError(f"unsupported XLSB function ID {identifier}")
        if identifier == 255:
            if count is None or count < 1:
                raise ValueError("invalid XLSB future function")
            args = [self.pop()[0] for _ in range(count)][::-1]
            function, args = args[0], args[1:]
        else:
            function, info = BY_ID[identifier]
            count = info[3] if count is None else count
            if count is None:
                raise ValueError("variable-argument function used as fixed")
            args = [self.pop()[0] for _ in range(count)][::-1]
        self.push(function + "(" + ",".join(args) + ")")


def decode_formula(
    data: bytes,
    context: Any,
    *,
    sheet: str | None = None,
    origin: tuple[int, int] = (0, 0),
    name_formula: bool = False,
) -> str:
    return _Decoder(data, context, sheet=sheet, origin=origin, name_formula=name_formula).decode()


# PtgList flags ([MS-XLSB] 2.5.97.52): bits 0-1 say how the columns are given, the others
# which rows. Excel lists the rows in this order.
_LIST_ROWS = ((0x08, "#Headers"), (0x10, "#Data"), (0x20, "#Totals"), (0x04, "#All"), (0x40, "#This Row"))
_LIST_ESCAPE = re.compile(r"(['#\[\]@])")


def structured_reference(table: str, columns: list[str], flags: int, first: int, last: int) -> str:
    """``Table[[#Data],[Col]]`` as Excel writes it to a file (``[@Col]`` is ``[[#This Row],[Col]]``).

    ``first`` and ``last`` are the 0-based positions of the columns in the table.
    """
    mode = flags & 3
    if mode == 3:
        raise ValueError("invalid XLSB structured reference")
    parts = [f"[{label}]" for bit, label in _LIST_ROWS if flags & bit]
    if mode:
        if not (0 <= first < len(columns) and first <= last < len(columns)):
            raise ValueError("invalid XLSB structured reference column")
        labels = [_LIST_ESCAPE.sub(r"'\1", columns[index]) for index in ((first,) if mode == 1 else (first, last))]
        parts.append("[" + labels[0] + "]" if mode == 1 else "[" + labels[0] + "]:[" + labels[1] + "]")
    if not parts:
        return table
    if len(parts) == 1:
        text = parts[0]
        # A single column is [Col]; a name that begins or ends with a space needs its own brackets.
        if mode == 1 and (text[1:-1] != text[1:-1].strip()):
            text = "[" + text + "]"
        elif mode == 2:
            text = "[" + text + "]"
        return table + text
    return table + "[" + ",".join(parts) + "]"


def _unescape_column(raw: str) -> str:
    """A column name as a structured reference writes it: ``'`` escapes the next character."""
    chars = []
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == "'":
            if index + 1 >= len(raw) or raw[index + 1] not in _ESCAPABLE:
                raise ValueError(f"invalid escape in the structured reference column {raw!r}")
            chars.append(raw[index + 1])
            index += 2
            continue
        if char in "[]":
            raise ValueError(f"a bracket in the structured reference column {raw!r} must be escaped")
        chars.append(char)
        index += 1
    if not chars:
        raise ValueError("an empty column name in a structured reference")
    return "".join(chars)


def _bracket_groups(text: str) -> list[list[str]]:
    """The groups of a bracketed spec: ``[a]`` or ``[a]:[b]`` (a range), separated by commas."""
    groups = []
    position = 0
    while True:
        if text[position : position + 1] != "[":
            raise ValueError(f"invalid structured reference [{text}]")
        item, position = _scan_group(text, position)
        group = [item]
        if text[position : position + 1] == ":":
            if text[position + 1 : position + 2] != "[":
                raise ValueError(f"invalid structured reference [{text}]")
            item, position = _scan_group(text, position + 1)
            group.append(item)
        groups.append(group)
        if position == len(text):
            return groups
        if text[position : position + 1] != ",":
            raise ValueError(f"invalid structured reference [{text}]")
        position += 1


def _parse_spec(spec: str) -> tuple[list[str], list[str] | None]:
    """``(row keywords, column names or None)`` of the text inside a structured reference's brackets.

    Takes ``#Data``, ``Col``, ``[Col1]:[Col2]``, ``[#Headers],[Col]`` and ``@Col`` as Excel writes
    them; the keywords come back in upper case, the names unescaped.
    """
    this_row = spec.startswith("@")
    if this_row:
        spec = spec[1:]
    keywords = ["#THIS ROW"] if this_row else []
    if spec == "":
        if this_row:
            return keywords, None  # Table[@]: the whole of this row
        raise ValueError("an empty structured reference")
    groups = _bracket_groups(spec) if spec.startswith("[") else [[spec]]
    columns = None
    for group in groups:
        if len(group) == 1 and group[0].startswith("#"):
            keyword = group[0].upper()
            if keyword not in _KEYWORD_BITS:
                raise ValueError(f"unknown keyword {group[0]!r} in a structured reference")
            if keyword not in keywords:
                keywords.append(keyword)
            continue
        if any(item.startswith("#") for item in group):
            raise ValueError(f"a keyword cannot be part of a range of columns: [{spec}]")
        if columns is not None:
            raise ValueError("a structured reference names one column, or one range of columns")
        columns = [_unescape_column(item) for item in group]
    _check_row_keywords(keywords, spec)
    return keywords, columns


def _check_row_keywords(keywords: list[str], spec: str) -> None:
    """Refuse the combinations of row keywords that Excel refuses.

    ``#All`` and ``#This Row`` stand alone; ``#Headers`` and ``#Totals`` meet
    only through ``#Data``. Duplicates are allowed, as Excel accepts and removes
    them.
    """
    unique = set(keywords)
    if "#ALL" in unique and len(unique) > 1:
        raise ValueError(f"#All cannot be combined with another row keyword: [{spec}]")
    if "#THIS ROW" in unique and len(unique) > 1:
        raise ValueError(f"#This Row cannot be combined with another row keyword: [{spec}]")
    if "#HEADERS" in unique and "#TOTALS" in unique:
        raise ValueError(f"#Headers and #Totals need #Data between them: [{spec}]")


def encode_table(name: str, spec: str | None, context: Any, operand: str) -> bytes:
    """The PtgExtend/PtgList token of ``name[spec]`` ([MS-XLSB] 2.5.97.52).

    The token names the table by its id, its sheet by an XTI index and its columns by
    their positions. ``spec`` is None for the table alone. The flags follow Excel: #This
    Row sets 0x0200, and a reference used as a value (not as the argument of a function
    that takes references) sets 0x0400.
    """
    ident, columns, sheet = context.table_for_formula(name)
    keywords, names = _parse_spec(spec) if spec is not None else ([], None)
    mode = first = last = 0
    if names is not None:
        positions = []
        for column in names:
            index = next((i for i, known in enumerate(columns) if known.casefold() == column.casefold()), None)
            if index is None:
                raise ValueError(f"table {name!r} has no column {column!r}")
            positions.append(index)
        if len(positions) == 1:
            mode, first, last = 1, positions[0], positions[0]
        else:
            if positions[0] > positions[1]:
                raise ValueError(f"the columns of {name}[{spec}] are out of order")
            mode, first, last = 2, positions[0], positions[1]
    flags = mode | sum(_KEYWORD_BITS[keyword] for keyword in keywords)
    if "#THIS ROW" in keywords:
        flags |= 0x0200
    if operand != "R":
        flags |= 0x0400
    return b"\x18\x19" + struct.pack("<HHIHH", context.sheet_reference(sheet), flags, ident, first, last)


# BrtName's fHidden flag ([MS-XLSB] 2.4.691).
_NAME_HIDDEN = 0x1


def _name_comment(raw: bytes, formula_offset: int) -> str | None:
    """The comment of a BrtName: the string after its formula, None when there is none."""
    try:
        end = read_parsed(raw, formula_offset)[2]
        return read_wide(raw, end, nullable=True)[0] or None
    except (ValueError, struct.error):
        return None


@dataclass
class _Name:
    name: str
    scope: int | str | None
    flags: int = 0
    raw: bytes | None = None
    definition: DefinedName | None = None


def external_sheet_names(part: bytes | None) -> list[str]:
    """The sheet names of an external workbook (BrtSupTabs of an externalLink part)."""
    if not part:
        return []
    for record, payload in records(part):
        if record.record_id == SUP_TABS and len(payload) >= 4:
            count, offset = struct.unpack_from("<I", payload)[0], 4
            names = []
            for _ in range(count):
                name, offset = read_wide(payload, offset)
                names.append(name)
            return names
    return []


class FormulaContext:
    """``parts`` (path -> bytes or None) lets references to other workbooks show their sheet names."""

    def __init__(
        self,
        sheets: Any,
        workbook: bytes | None = None,
        parts: Any = None,
        *,
        tables: Any = None,
        cache_tables: bool = True,
        removed_tables: Iterable[str] = (),
    ) -> None:
        self._sheets = sheets
        # ``tables()`` gives {table id: (name, column names, sheet title or None)}, for structured
        # references. A writer adds tables while it streams, so it asks for a fresh catalog each time.
        self._tables = tables
        self._cache_tables = cache_tables
        self._table_catalog: dict[int, tuple[str, list[str], str | None]] | None = None
        # Tables a save removes are not in the catalog, and a new formula cannot name them.
        self._removed_tables = frozenset(name.casefold() for name in removed_tables)
        self.names: list[_Name] = []
        self.links: list[int] = []
        self.xti: list[Any] = []
        self._external_block: bytes | None = None
        self.external_rids: list[str] = []
        self.external_names: list[list[str]] = []
        if workbook is not None:
            external_start = None
            for record, payload in records(workbook):
                kind = record.record_id
                if kind == NAME:
                    flags, _, scope = struct.unpack_from("<IBI", payload)
                    name, _ = read_wide(payload, 9)
                    self.names.append(_Name(name, None if scope == 0xFFFFFFFF else scope, flags, payload))
                elif kind == BEGIN_EXTERNALS:
                    external_start = record.header_start
                elif kind == END_EXTERNALS and external_start is not None:
                    self._external_block = workbook[external_start : record.data_end]
                elif kind in (0x163, 0x165, 0x166):
                    self.links.append(kind)
                    if kind == SUP_BOOK_SRC:
                        self.external_rids.append(read_wide(payload, 0)[0])
                elif kind == EXTERN_SHEET:
                    count = struct.unpack_from("<I", payload)[0]
                    if len(payload) != 4 + count * 12:
                        raise ValueError("invalid XLSB external sheet table")
                    self.xti = [struct.unpack_from("<Iii", payload, 4 + i * 12) for i in range(count)]
        if parts is not None and self.external_rids:
            self._load_external_names(parts)

    def _load_external_names(self, parts: Any) -> None:
        rels = parts("xl/_rels/workbook.bin.rels")
        targets = {}
        if rels:
            targets = {node.get("Id"): node.get("Target") for node in safe_fromstring(rels)}
        for rid in self.external_rids:
            target = targets.get(rid)
            path = None if not target else ("xl/" + target if not target.startswith("/") else target[1:])
            self.external_names.append(external_sheet_names(parts(posixpath.normpath(path)) if path else None))

    @property
    def sheets(self) -> list[str]:
        return list(self._sheets() if callable(self._sheets) else self._sheets)

    def sheet_reference(self, title: str) -> int:
        title = next((s for s in self.sheets if s.casefold() == title.casefold()), title)
        # Sheet titles may be declared after the formula is streamed.
        for index, entry in enumerate(self.xti):
            if entry == title:
                return index
            if (
                isinstance(entry, tuple)
                and entry[0] < len(self.links)
                and self.links[entry[0]] == SUP_SELF
                and entry[1] == entry[2]
                and 0 <= entry[1] < len(self.sheets)
                and self.sheets[entry[1]] == title
            ):
                return index
        self.xti.append(title)
        return len(self.xti) - 1

    def sheet_label(self, index: int) -> str:
        if not 0 <= index < len(self.xti):
            raise ValueError("invalid XLSB sheet reference index")
        entry = self.xti[index]
        if isinstance(entry, str):
            return quote_sheet_name(entry) + "!"
        link, first, last = entry
        if link < len(self.links) and self.links[link] == SUP_BOOK_SRC:
            return self._external_label(link, first, last)
        if link >= len(self.links) or self.links[link] != SUP_SELF:
            raise UnsupportedFeatureError("this kind of external reference is not supported")
        if first < 0 or last < 0:
            return "#REF!"
        title = self.sheets[first] if first == last else self.sheets[first] + ":" + self.sheets[last]
        return quote_sheet_name(title) + "!"

    def _external_label(self, link: int, first: int, last: int) -> str:
        """``'[1]Sheet'!`` for a sheet of the n-th external workbook."""
        number = sum(kind == SUP_BOOK_SRC for kind in self.links[: link + 1])
        names = self.external_names[number - 1] if number <= len(self.external_names) else []

        def name(index: int) -> str:
            return names[index] if 0 <= index < len(names) else f"Sheet{index + 1}"

        if first < 0 or last < 0:
            return "#REF!"
        title = name(first) if first == last else name(first) + ":" + name(last)
        return quote_sheet_name(f"[{number}]{title}") + "!"

    def _scope(self, entry: _Name) -> str | None:
        return self.sheets[entry.scope] if isinstance(entry.scope, int) else entry.scope

    def _catalog(self) -> Any:
        if self._tables is None:
            return {}
        if not self._cache_tables:
            return self._tables()
        if self._table_catalog is None:
            self._table_catalog = self._tables()
        return self._table_catalog

    def table_reference(self, ident: int, flags: int, first: int, last: int) -> str:
        """The text of a structured reference (PtgList) to the table with id ``ident``."""
        catalog = self._catalog()
        if ident not in catalog:
            raise UnsupportedFeatureError(f"a formula refers to a table (id {ident}) that the workbook does not hold")
        name, columns = catalog[ident][:2]
        return structured_reference(name, columns, flags, first, last)

    def has_table(self, name: str) -> bool:
        """True when a table of the workbook has this name (or the save removes it)."""
        folded = name.casefold()
        return folded in self._removed_tables or any(
            entry[0].casefold() == folded for entry in self._catalog().values()
        )

    def table_for_formula(self, name: str) -> tuple[int, list[str], str]:
        """``(id, column names, sheet title)`` of the table ``name``, for a formula that refers to it."""
        folded = name.casefold()
        if folded in self._removed_tables:
            raise UnsupportedFeatureError(f"table {name!r} is removed by this save, so a formula cannot refer to it")
        for ident, entry in self._catalog().items():
            if entry[0].casefold() == folded:
                if entry[2] is None:
                    raise UnsupportedFeatureError(f"the sheet of table {name!r} is not known")
                return ident, list(entry[1]), entry[2]
        raise UnsupportedFeatureError(f"table {name!r} is not in the workbook, or its header row is not written yet")

    def name_reference(self, title: str, scope: str | None = None, *, qualified: bool = False) -> int:
        if scope is not None:
            scope = next((s for s in self.sheets if s.casefold() == scope.casefold()), scope)
        for candidate_scope in (scope, None) if scope is not None else (None,):
            for index, entry in enumerate(self.names):
                if entry.name.casefold() == title.casefold() and self._scope(entry) == candidate_scope:
                    return index
        # A sheet-qualified name (Sheet!Name) defined later is a local name;
        # set_name() then fills this placeholder in place.
        self.names.append(_Name(title, scope if qualified else None))
        return len(self.names) - 1

    def name_label(self, index: int, sheet: str | None = None) -> str:
        if not 0 <= index < len(self.names):
            raise ValueError("invalid XLSB name reference index")
        entry = self.names[index]
        scope = self._scope(entry)
        return (quote_sheet_name(scope) + "!" if scope is not None and scope != sheet else "") + entry.name

    def future_function(self, title: str) -> int:
        title = future_function_name(title)
        index = self.name_reference(title)
        entry = self.names[index]
        if entry.raw is None:
            entry.flags = 0x2000B
            entry.raw = (
                struct.pack("<IBI", entry.flags, 0, 0xFFFFFFFF)
                + wide_string(title)
                + parsed_formula(b"\x1c\x1d")
                + b"\xff" * 20
            )
        return index

    def lambda_name(self, title: str) -> int:
        """A LET/LAMBDA parameter: the hidden ``_xlpm.`` name Excel writes."""
        index = self.name_reference(title)
        entry = self.names[index]
        if entry.raw is None:
            entry.flags = 0x20019
            entry.raw = (
                struct.pack("<IBI", entry.flags, 0, 0xFFFFFFFF)
                + wide_string(title)
                + parsed_formula(b"\x1c\x1d")
                + b"\xff" * 20
            )
        return index

    def set_name(self, definition: DefinedName) -> None:
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be DefinedName")
        scope = definition.scope
        if scope is None and definition.localSheetId is not None and definition.localSheetId < len(self.sheets):
            scope = list(self.sheets)[definition.localSheetId]
        match = next(
            (
                i
                for i, n in enumerate(self.names)
                if n.name.casefold() == definition.name.casefold() and self._scope(n) == scope
            ),
            None,
        )
        entry = _Name(definition.name, scope, flags=_NAME_HIDDEN if definition.hidden else 0, definition=definition)
        if match is None:
            self.names.append(entry)
        else:
            self.names[match] = entry

    def remove_name(self, title: str, scope: str | None) -> bool:
        """Remove a name; False when there was no such name."""
        DefinedName(title, "0", scope)
        for index, entry in enumerate(self.names):
            if entry.name.casefold() == title.casefold() and self._scope(entry) == scope:
                # PtgName uses ordinal IDs. Keep the slot so other formulas and
                # opaque parts cannot accidentally start referring to a new name.
                name = f"_xlspy_deleted_{index}_{len(self.names)}"
                entry.name, entry.flags, entry.definition = name, 1, None
                scope_id = 0xFFFFFFFF if scope is None else self.sheets.index(scope)
                entry.raw = (
                    struct.pack("<IBI", 1, 0, scope_id) + wide_string(name) + parsed_formula(b"\x1c\x1d") + b"\xff" * 4
                )
                return True
        return False

    def name_payloads(self) -> list[bytes]:
        payloads: list[bytes] = []
        index = 0
        while index < len(self.names):
            entry = self.names[index]
            if entry.definition is not None:
                definition = entry.definition.resolved(self.sheets)
                scope_id = 0xFFFFFFFF if definition.scope is None else definition.localSheetId
                formula = compile_formula(definition.formula, self, sheet=definition.scope, name_formula=True)
                # What follows the formula is the name's comment (a null string when there is none).
                comment = wide_string(definition.comment) if definition.comment else b"\xff" * 4
                entry.raw = (
                    struct.pack("<IBI", entry.flags, 0, scope_id) + wide_string(definition.name) + formula + comment
                )
            if entry.raw is None:
                raise ValueError(f"formula references undefined name {entry.name!r}")
            payloads.append(entry.raw)
            index += 1
        return payloads

    def defined_names(self) -> list[DefinedName]:
        result: list[DefinedName] = []
        for entry in self.names:
            if entry.flags & (0x20 | 0x20000) or entry.name.startswith("_xlspy_deleted_"):
                continue
            if entry.definition is not None:
                result.append(entry.definition.resolved(self.sheets))
            elif entry.raw is not None:
                _, offset = read_wide(entry.raw, 9)
                scope = self._scope(entry)
                result.append(
                    DefinedName(
                        entry.name,
                        decode_formula(entry.raw[offset:], self, sheet=scope, name_formula=True),
                        scope,
                        localSheetId=None if scope is None else list(self.sheets).index(scope),
                        hidden=True if entry.flags & _NAME_HIDDEN else None,
                        comment=_name_comment(entry.raw, offset),
                    )
                )
        return result

    def external_records(self) -> bytes:
        if not self.xti and self._external_block is None:
            return b""
        links = list(self.links)
        if any(isinstance(entry, str) for entry in self.xti) and SUP_SELF not in links:
            links.append(SUP_SELF)
        entries = []
        for entry in self.xti:
            if isinstance(entry, str):
                title = next((s for s in self.sheets if s.casefold() == entry.casefold()), None)
                if title is None:
                    raise ValueError(f"formula references unknown worksheet {entry!r}")
                index = self.sheets.index(title)
                entry = (links.index(SUP_SELF), index, index)
            entries.append(struct.pack("<Iii", *entry))
        table = build_record(EXTERN_SHEET, struct.pack("<I", len(entries)) + b"".join(entries))
        if self._external_block is None:
            return (
                build_record(BEGIN_EXTERNALS)
                + b"".join(build_record(kind) for kind in links)
                + table
                + build_record(END_EXTERNALS)
            )
        block = self._external_block
        if len(links) > len(self.links):
            block = replace_records(block, set(), build_record(SUP_SELF), before=(EXTERN_SHEET, END_EXTERNALS))
        return replace_records(block, {EXTERN_SHEET}, table, before=(END_EXTERNALS,))

    def write_workbook(self, original: bytes) -> bytes:
        payloads = self.name_payloads()
        externals = self.external_records()
        output = []
        external_start = False
        wrote_external = False
        name_index = 0
        wrote_names = False
        for record, payload in records(original):
            kind = record.record_id
            if kind == BEGIN_EXTERNALS:
                output.append(externals)
                external_start = True
                wrote_external = True
            elif external_start:
                if kind == END_EXTERNALS:
                    external_start = False
            elif kind == NAME:
                if not wrote_external:
                    output.append(externals)
                    wrote_external = True
                output.append(build_record(NAME, payloads[name_index]))
                name_index += 1
            else:
                if kind in (0x9D, 0x84) and not wrote_names:
                    if not wrote_external:
                        output.append(externals)
                        wrote_external = True
                    output.extend(build_record(NAME, data) for data in payloads[name_index:])
                    wrote_names = True
                output.append(original[record.header_start : record.data_end])
        return b"".join(output)


@dataclass(frozen=True)
class ArrayValue:
    source_row: int
    source_column: int
    value: object = None


def formula_record(formula: Any, column: int, style: int, parsed: bytes, value: Any = None) -> bytes:
    """A ``BrtFmla*`` record for a formula already compiled to ``parsed``."""
    header = struct.pack("<II", column, style)
    if value is None:
        value = getattr(formula, "value", None)
    if isinstance(value, (datetime.date, datetime.datetime)):
        # Array members may carry the cell value that was in the row.
        value = datetime_to_serial(value, False)
    if isinstance(value, str):
        kind, cache = 8, wide_string(value)
    elif isinstance(value, bool):
        kind, cache = 10, bytes([value])
    elif value is None or isinstance(value, (int, float, Decimal)):
        if not math.isfinite(value or 0):
            raise ValueError("XLSB formula cached values must be finite")
        kind, cache = 9, struct.pack("<d", float(value or 0))
    else:
        raise TypeError(f"unsupported XLSB formula cached value {type(value).__name__}")
    return build_record(kind, header + cache + b"\x01\0" + parsed)


def formula_cell(
    formula: Any, column: int, style: int, context: Any, *, sheet: str | None = None, row: int = 0
) -> bytes:
    """Encode a scalar formula; cached zero is used if no cache was supplied."""
    if getattr(formula, "spill_ref", None) is not None:
        raise UnsupportedFeatureError("a dynamic-array formula needs the workbook writer's spill metadata")
    array_definition = b""
    if isinstance(formula, ArrayValue):
        parsed = parsed_formula(
            b"\x01" + struct.pack("<I", formula.source_row), struct.pack("<I", formula.source_column)
        )
    elif formula.array_ref is not None:
        c1, r1, c2, r2 = parse_cell_range(formula.array_ref)
        if (column, row) != (c1, r1):
            raise ValueError("array formula must be written in its top-left cell")
        parsed = parsed_formula(b"\x01" + struct.pack("<I", row), struct.pack("<I", column))
        group = compile_formula(formula.text, context, sheet=sheet, origin=(column, row))
        array_definition = build_record(0x1AA, struct.pack("<IIII", r1, r2, c1, c2) + b"\x01" + group)
    else:
        parsed = compile_formula(formula.text, context, sheet=sheet, origin=(column, row))
    return formula_record(formula, column, style, parsed) + array_definition
