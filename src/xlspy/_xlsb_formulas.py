"""BIFF12 formula tokens and workbook name/reference context.

Used only for formulas and metadata; ordinary data encoders never call this
module. No Excel installation or third-party parser is needed at runtime.
"""

from __future__ import annotations

from dataclasses import dataclass
import datetime
from decimal import Decimal
import math
import re
import struct

from ._metadata import DefinedName
from ._xlsb_functions import FUNCTIONS, BY_ID
from ._xlsb_metadata import wide_string, read_wide, records, replace_records
from .biff_utils import build_record
from .excel_dates import datetime_to_serial
from .formulas import _XLFN_FUNCTIONS
from .refs import parse_cell_range, column_label, quote_sheet_name

NAME, BEGIN_EXTERNALS, END_EXTERNALS, SUP_SELF, EXTERN_SHEET = 0x27, 0x161, 0x162, 0x165, 0x16A
ERRORS = {0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA'}
_OPS = {3: '+', 4: '-', 5: '*', 6: '/', 7: '^', 8: '&', 9: '<', 10: '<=', 11: '=', 12: '>=', 13: '>', 14: '<>', 15: ' ', 16: ',', 17: ':'}
_OP_IDS = {value: key for key, value in _OPS.items()}
_PREC = {',': 1, '=': 2, '<': 2, '>': 2, '<=': 2, '>=': 2, '<>': 2,
         '&': 3, '+': 4, '-': 4, '*': 5, '/': 5, '^': 6, ' ': 9, ':': 10}
_ADDRESS = r'(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)'
_LEX = re.compile(
    r'\s+|"(?:[^"]|"")*"|\x27(?:[^\x27]|\x27\x27)*\x27'
    r'|#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)'
    + rf'|{_ADDRESS}(?![\w.])'
    + r'|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[^\W\d][\w.]*|\\[\w.]*|<>|<=|>=|[+\-*/^&%=<>(),!:]',
    re.UNICODE | re.I,
)


def parsed_formula(tokens, extra=b''):
    if not 0 < len(tokens) <= 16384:
        raise ValueError("XLSB formula token length exceeds the BIFF12 limit")
    return struct.pack('<I', len(tokens)) + tokens + struct.pack('<I', len(extra)) + extra


def read_parsed(data, offset=0):
    if offset + 4 > len(data):
        raise ValueError("Truncated XLSB formula length")
    length = struct.unpack_from('<I', data, offset)[0]
    end = offset + 4 + length
    if not 0 < length <= 16384 or end + 4 > len(data):
        raise ValueError("Invalid XLSB formula token length")
    extra_length = struct.unpack_from('<I', data, end)[0]
    if end + 4 + extra_length > len(data):
        raise ValueError("Truncated XLSB formula extra data")
    return data[offset + 4:end], data[end + 4:end + 4 + extra_length], end + 4 + extra_length


class _Parser:
    def __init__(self, text):
        self.tokens = []
        end = 0
        for match in _LEX.finditer(text):
            if match.start() != end:
                raise NotImplementedError(f"unsupported XLSB formula syntax near {text[end:end + 24]!r}")
            token = match[0]
            end = match.end()
            if not token.isspace():
                self.tokens.append(token)
        if end != len(text):
            raise NotImplementedError(f"unsupported XLSB formula syntax near {text[end:end + 24]!r}")
        self.tokens.append(None)
        self.index = 0

    def peek(self):
        return self.tokens[self.index]

    def pop(self):
        token = self.peek()
        if token is None:
            raise ValueError("unexpected end of XLSB formula")
        self.index += 1
        return token

    def expression(self, minimum=0):
        token = self.pop()
        if token in ('+', '-'):
            left = ('unary', token, self.expression(7))
        elif token == '(':
            left = ('paren', self.expression())
            if self.pop() != ')':
                raise ValueError("missing closing parenthesis")
        elif token.startswith('"'):
            left = ('string', token[1:-1].replace('""', '"'))
        elif token.startswith('#'):
            left = ('error', token.upper())
        else:
            sheet = None
            if self.peek() == '!':
                self.pop()
                sheet = token[1:-1].replace("''", "'") if token.startswith("'") else token
                token = self.pop()
            if self.peek() == '(' and sheet is None:
                self.pop()
                args = []
                if self.peek() != ')':
                    while True:
                        args.append(('missing',) if self.peek() in (',', ')') else self.expression(2))
                        if self.peek() != ',':
                            break
                        self.pop()
                if self.pop() != ')':
                    raise ValueError("invalid function argument list")
                left = ('function', token.upper(), args)
            elif re.fullmatch(_ADDRESS, token):
                left = ('reference', token, sheet)
            elif re.fullmatch(r'(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?', token):
                left = ('number', token)
            elif token.upper() in ('TRUE', 'FALSE') and sheet is None:
                left = ('bool', token.upper() == 'TRUE')
            else:
                left = ('name', token, sheet)
        while True:
            operator = self.peek()
            if operator == '%':
                self.pop()
                left = ('percent', left)
                continue
            precedence = _PREC.get(operator, -1)
            if precedence < minimum:
                break
            self.pop()
            right = self.expression(precedence + 1)
            left = ('binary', operator, left, right)
        return left


def parse_expression(text):
    parser = _Parser(text.removeprefix('='))
    tree = parser.expression()
    if parser.peek() is not None:
        raise NotImplementedError(f"unsupported XLSB formula token {parser.peek()!r}")
    return tree


def _location(ref, *, relative_offsets=False, origin=(0, 0)):
    match = re.fullmatch(r'(\$?)([A-Za-z]+)(\$?)(\d+)', ref)
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
_VOLATILE_FUNCTIONS = frozenset({'NOW', 'TODAY', 'RAND', 'OFFSET', 'INDIRECT', 'CELL', 'INFO'})


def _is_relative(ref):
    """True when a column or row of an A1 reference lacks its ``$``."""
    for part in ref.split(':'):
        match = re.fullmatch(r'(\$?)([A-Za-z]*)(\$?)(\d*)', part)
        if (match[2] and not match[1]) or (match[4] and not match[3]):
            return True
    return False


def encode_reference(ref, sheet, context, operand, origin, name_formula, cell_relative=False):
    kind = 0x20 if operand == 'R' else 0x60 if operand == 'A' else 0x40
    # Validation and conditional-format formulas store same-sheet references
    # with a relative part as offsets from the range's top-left cell
    # (PtgRefN, PtgAreaN); fully absolute ones stay PtgRef/PtgArea, like Excel.
    relative_token = cell_relative and sheet is None and _is_relative(ref)
    offsets = name_formula or relative_token
    if ':' in ref:
        a, b = ref.split(':')
        if not re.search(r'\d', a):
            a, b = a + '$1', b + '$1048576'
        elif not re.search('[A-Za-z]', a):
            a, b = '$A' + a, '$XFD' + b
        r1, c1 = _location(a, relative_offsets=offsets, origin=origin)
        r2, c2 = _location(b, relative_offsets=offsets, origin=origin)
        token = bytes([kind | (0x1B if sheet is not None else 0x0D if relative_token else 5)])
        payload = struct.pack('<IIHH', r1, r2, c1, c2)
    else:
        row, col = _location(ref, relative_offsets=offsets, origin=origin)
        token = bytes([kind | (0x1A if sheet is not None else 0x0C if relative_token else 4)])
        payload = struct.pack('<IH', row, col)
    if sheet is not None:
        token += struct.pack('<H', context.sheet_reference(sheet))
    return token + payload


def compile_formula(text, context, *, sheet=None, origin=(0, 0), name_formula=False,
                    cell_relative=False, operand=None, stats=None):
    """Encode formula text as a CellParsedFormula.

    ``stats``, when given, is a dict that receives the number of reference
    tokens under ``"references"``.
    """
    if not isinstance(text, str) or not text or len(text) > 8192:
        raise ValueError("invalid XLSB formula text or length")
    references = 0
    volatile = False

    def emit(node, operand='V'):
        kind = node[0]
        if kind == 'number':
            value = float(node[1])
            if not math.isfinite(value):
                raise ValueError("XLSB formula numeric constants must be finite")
            return b'\x1e' + struct.pack('<H', int(value)) if value.is_integer() and 0 <= value <= 65535 else b'\x1f' + struct.pack('<d', value)
        if kind == 'string':
            data = node[1].encode('utf-16le')
            if len(data) // 2 > 255:
                raise ValueError("XLSB formula string literals are limited to 255 UTF-16 units")
            return b'\x17' + struct.pack('<H', len(data) // 2) + data
        if kind == 'bool':
            return b'\x1d' + bytes([node[1]])
        if kind == 'error':
            return b'\x1c' + bytes([next(code for code, label in ERRORS.items() if label == node[1])])
        if kind == 'missing':
            return b'\x16'
        if kind == 'reference':
            nonlocal references
            references += 1
            return encode_reference(node[1], node[2], context, operand, origin, name_formula, cell_relative)
        if kind == 'name':
            index = context.name_reference(node[1], node[2] or sheet, qualified=node[2] is not None)
            return bytes([0x23 if operand == 'R' else 0x63 if operand == 'A' else 0x43]) + struct.pack('<I', index + 1)
        if kind == 'paren':
            return emit(node[1], operand) + b'\x15'
        if kind == 'percent':
            return emit(node[1]) + b'\x14'
        if kind == 'unary':
            return emit(node[2]) + (b'\x12' if node[1] == '+' else b'\x13')
        if kind == 'binary':
            arg_type = 'R' if node[1] in (',', ' ', ':') else 'V'
            return emit(node[2], arg_type) + emit(node[3], arg_type) + bytes([_OP_IDS[node[1]]])
        if kind == 'function':
            function, args = node[1], node[2]
            base = function.removeprefix('_XLFN.')
            info = FUNCTIONS.get(base)
            result_class = 0x20 if operand == 'R' else 0x60 if operand == 'A' else 0x40
            if info is None:
                if base not in _XLFN_FUNCTIONS:
                    raise NotImplementedError(f"unsupported XLSB function {function!r}")
                index = context.future_function(function)
                if len(args) > 254:
                    raise ValueError("too many XLSB function arguments")
                return b'\x23' + struct.pack('<I', index + 1) + b''.join(emit(arg, 'R') for arg in args) + bytes([result_class | 2, len(args) + 1]) + struct.pack('<H', 255)
            identifier, minimum, maximum, fixed, classes = info
            if base in _VOLATILE_FUNCTIONS:
                nonlocal volatile
                volatile = True
            if not minimum <= len(args) <= maximum:
                raise ValueError(f"{function} expects {minimum}..{maximum} arguments")
            def argument_class(index):
                if identifier == 481:
                    return 'R' if index % 2 == 0 else 'V'
                if identifier in (482, 484):
                    return 'R' if index == 0 or index % 2 else 'V'
                return classes[min(index, len(classes) - 1)] if classes else 'V'
            encoded = [emit(arg, argument_class(index)) for index, arg in enumerate(args)]
            ending = (bytes([result_class | 1]) + struct.pack('<H', identifier) if fixed is not None
                      else bytes([result_class | 2, len(args)]) + struct.pack('<H', identifier))
            if identifier == 1:  # IF, including its short-circuit jump records.
                yes = encoded[1]
                no = encoded[2] if len(encoded) == 3 else b''
                first_goto = b'\x19\x08' + struct.pack('<H', len(no) + (4 if no else 0) + len(ending) - 1)
                last_goto = b'\x19\x08' + struct.pack('<H', len(ending) - 1) if no else b''
                return encoded[0] + b'\x19\x02' + struct.pack('<H', len(yes) + 4) + yes + first_goto + no + last_goto + ending
            if identifier == 480:  # IFERROR
                return encoded[0] + b'\x19\x80' + struct.pack('<H', len(encoded[1])) + encoded[1] + b'\x19\x08' + struct.pack('<H', len(ending) - 1) + ending
            if identifier == 100:  # CHOOSE jump table and per-branch exits.
                alternatives = encoded[1:]
                offsets = [2 * (len(alternatives) + 1)]
                for branch in alternatives:
                    offsets.append(offsets[-1] + len(branch) + 4)
                chunks = [encoded[0], b'\x19\x04' + struct.pack('<H', len(alternatives))
                          + b''.join(struct.pack('<H', offset) for offset in offsets)]
                for index, branch in enumerate(alternatives):
                    remaining = sum(len(value) + 4 for value in alternatives[index + 1:]) + len(ending) - 1
                    chunks.append(branch + b'\x19\x08' + struct.pack('<H', remaining))
                return b''.join(chunks) + ending
            return b''.join(encoded) + ending
        raise NotImplementedError(f"unsupported XLSB formula node {kind}")

    top = operand or ('R' if name_formula else 'V')
    tokens = emit(parse_expression(text), top)
    if volatile:
        # PtgAttrSemi opens every formula with a volatile function.
        tokens = b'\x19\x01\x00\x00' + tokens
    if stats is not None:
        stats['references'] = references
    return parsed_formula(tokens)


def _decode_ref(row, column, origin, relative_offsets):
    col = column & 0x3FFF
    col_relative, row_relative = bool(column & 0x4000), bool(column & 0x8000)
    if relative_offsets:
        if col_relative:
            col = (origin[0] + (col - 0x4000 if col & 0x2000 else col)) % 16384
        if row_relative:
            row = (origin[1] + (row - 0x100000 if row & 0x80000 else row)) % 1048576
    if row > 1048575:
        return '#REF!'
    return ('' if col_relative else '$') + column_label(col) + ('' if row_relative else '$') + str(row + 1)


def decode_formula(data, context, *, sheet=None, origin=(0, 0), name_formula=False):
    tokens, extra, _ = read_parsed(data)
    stack = []
    position = 0
    def push(text, precedence=100):
        stack.append((text, precedence))
    def pop():
        if not stack:
            raise ValueError("invalid XLSB formula operand stack")
        return stack.pop()
    def take(size):
        nonlocal position
        end = position + size
        if end > len(tokens):
            raise ValueError("truncated XLSB formula token")
        value = tokens[position:end]
        position = end
        return value
    while position < len(tokens):
        token = take(1)[0]
        base = (token & 0x1F) | 0x20 if token >= 0x20 else token
        if token in _OPS:
            right, rp = pop()
            left, lp = pop()
            op = _OPS[token]
            prec = _PREC[op]
            if lp < prec:
                left = '(' + left + ')'
            if rp < prec or (rp == prec and op in ('-', '/', '^', '<', '>', '=', '<=', '>=', '<>')):
                right = '(' + right + ')'
            push(left + op + right, prec)
        elif token in (0x12, 0x13):
            text, prec = pop()
            push(('+' if token == 0x12 else '-') + ('(' + text + ')' if prec < 7 else text), 7)
        elif token == 0x14:
            text, prec = pop()
            push(('(' + text + ')' if prec < 8 else text) + '%', 8)
        elif token == 0x15:
            push('(' + pop()[0] + ')')
        elif token == 0x16:
            push('')
        elif token == 0x17:
            length = struct.unpack('<H', take(2))[0]
            push('"' + take(length * 2).decode('utf-16le').replace('"', '""') + '"')
        elif token == 0x1C:
            push(ERRORS.get(take(1)[0], '#VALUE!'))
        elif token == 0x1D:
            push('TRUE' if take(1)[0] else 'FALSE')
        elif token == 0x1E:
            push(str(struct.unpack('<H', take(2))[0]))
        elif token == 0x1F:
            push(format(struct.unpack('<d', take(8))[0], '.15g'))
        elif token == 0x19:
            flags = take(1)[0]
            value = struct.unpack('<H', take(2))[0]
            if flags & 4:  # CHOOSE jump table.
                take(2 * (value + 1))
            elif flags & 0x10:
                push('SUM(' + pop()[0] + ')')
        elif base in (0x24, 0x2C, 0x3A):
            qualifier = context.sheet_label(struct.unpack('<H', take(2))[0]) if base == 0x3A else ''
            row, col = struct.unpack('<IH', take(6))
            push(qualifier + _decode_ref(row, col, origin, name_formula or base == 0x2C))
        elif base in (0x25, 0x2D, 0x3B):
            qualifier = context.sheet_label(struct.unpack('<H', take(2))[0]) if base == 0x3B else ''
            r1, r2, c1, c2 = struct.unpack('<IIHH', take(12))
            a = _decode_ref(r1, c1, origin, name_formula or base == 0x2D)
            b = _decode_ref(r2, c2, origin, name_formula or base == 0x2D)
            # Excel stores whole-column/row references as full worksheet ranges.
            if r1 == 0 and r2 == 1048575 and not (c1 & 0x8000 or c2 & 0x8000):
                a, b = re.sub(r'\$?\d+$', '', a), re.sub(r'\$?\d+$', '', b)
            elif c1 & 0x3FFF == 0 and c2 & 0x3FFF == 16383 and not (c1 & 0x4000 or c2 & 0x4000):
                a, b = re.sub(r'^\$?[A-Za-z]+', '', a), re.sub(r'^\$?[A-Za-z]+', '', b)
            push(qualifier + a + ':' + b)
        elif base == 0x23:
            index = struct.unpack('<I', take(4))[0] - 1
            push(context.name_label(index, sheet))
        elif base in (0x21, 0x22):
            count = take(1)[0] & 0x7F if base == 0x22 else None
            identifier = struct.unpack('<H', take(2))[0]
            if identifier not in BY_ID and identifier != 255:
                raise NotImplementedError(f"unsupported XLSB function ID {identifier}")
            if identifier == 255:
                if count is None or count < 1:
                    raise ValueError("invalid XLSB future function")
                args = [pop()[0] for _ in range(count)][::-1]
                function, args = args[0], args[1:]
            else:
                function, info = BY_ID[identifier]
                count = info[3] if count is None else count
                if count is None:
                    raise ValueError("variable-argument function used as fixed")
                args = [pop()[0] for _ in range(count)][::-1]
            push(function + '(' + ','.join(args) + ')')
        elif base in (0x2A, 0x2B, 0x3C, 0x3D):
            take((2 if base in (0x3C, 0x3D) else 0) + (6 if base in (0x2A, 0x3C) else 12))
            push('#REF!')
        elif base in (0x26, 0x27, 0x28):
            take(6)  # Memory/cache tokens do not change the expression stack.
        elif base == 0x29:
            take(2)
        else:
            raise NotImplementedError(f"unsupported XLSB formula token 0x{token:02x}")
    if len(stack) != 1:
        raise ValueError("invalid XLSB formula stack result")
    return stack[0][0]


@dataclass
class _Name:
    name: str
    scope: int | str | None
    flags: int = 0
    raw: bytes | None = None
    definition: DefinedName | None = None


class FormulaContext:
    def __init__(self, sheets, workbook=None):
        self._sheets = sheets
        self.names = []
        self.links = []
        self.xti = []
        self._external_block = None
        if workbook is not None:
            external_start = None
            for record, payload in records(workbook):
                kind = record.record_id
                if kind == NAME:
                    flags, _, scope = struct.unpack_from('<IBI', payload)
                    name, _ = read_wide(payload, 9)
                    self.names.append(_Name(name, None if scope == 0xFFFFFFFF else scope, flags, payload))
                elif kind == BEGIN_EXTERNALS:
                    external_start = record.header_start
                elif kind == END_EXTERNALS and external_start is not None:
                    self._external_block = workbook[external_start:record.data_end]
                elif kind in (0x163, 0x165, 0x166):
                    self.links.append(kind)
                elif kind == EXTERN_SHEET:
                    count = struct.unpack_from('<I', payload)[0]
                    if len(payload) != 4 + count * 12:
                        raise ValueError("invalid XLSB external sheet table")
                    self.xti = [struct.unpack_from('<Iii', payload, 4 + i * 12) for i in range(count)]

    @property
    def sheets(self):
        return list(self._sheets() if callable(self._sheets) else self._sheets)

    def sheet_reference(self, title):
        title = next((s for s in self.sheets if s.casefold() == title.casefold()), title)
        # Sheet titles may be declared after the formula is streamed.
        for index, entry in enumerate(self.xti):
            if entry == title:
                return index
            if isinstance(entry, tuple) and entry[0] < len(self.links) and self.links[entry[0]] == SUP_SELF and entry[1] == entry[2] and 0 <= entry[1] < len(self.sheets) and self.sheets[entry[1]] == title:
                return index
        self.xti.append(title)
        return len(self.xti) - 1

    def sheet_label(self, index):
        if not 0 <= index < len(self.xti):
            raise ValueError("invalid XLSB sheet reference index")
        entry = self.xti[index]
        if isinstance(entry, str):
            return quote_sheet_name(entry) + '!'
        link, first, last = entry
        if link >= len(self.links) or self.links[link] != SUP_SELF:
            raise NotImplementedError("external-workbook formula references are not supported")
        if first < 0 or last < 0:
            return '#REF!'
        title = self.sheets[first] if first == last else self.sheets[first] + ':' + self.sheets[last]
        return quote_sheet_name(title) + '!'

    def _scope(self, entry):
        return self.sheets[entry.scope] if isinstance(entry.scope, int) else entry.scope

    def name_reference(self, title, scope=None, *, qualified=False):
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

    def name_label(self, index, sheet=None):
        if not 0 <= index < len(self.names):
            raise ValueError("invalid XLSB name reference index")
        entry = self.names[index]
        scope = self._scope(entry)
        return (quote_sheet_name(scope) + '!' if scope is not None and scope != sheet else '') + entry.name

    def future_function(self, title):
        title = '_xlfn.' + title.removeprefix('_XLFN.').removeprefix('_xlfn.')
        index = self.name_reference(title)
        entry = self.names[index]
        if entry.raw is None:
            entry.flags = 0x2000B
            entry.raw = struct.pack('<IBI', entry.flags, 0, 0xFFFFFFFF) + wide_string(title) + parsed_formula(b'\x1c\x1d') + b'\xff' * 20
        return index

    def set_name(self, definition):
        if not isinstance(definition, DefinedName):
            raise TypeError("definition must be DefinedName")
        match = next((i for i, n in enumerate(self.names) if n.name.casefold() == definition.name.casefold() and self._scope(n) == definition.scope), None)
        entry = _Name(definition.name, definition.scope, definition=definition)
        if match is None:
            self.names.append(entry)
        else:
            self.names[match] = entry

    def remove_name(self, title, scope):
        DefinedName(title, '0', scope)
        for index, entry in enumerate(self.names):
            if entry.name.casefold() == title.casefold() and self._scope(entry) == scope:
                # PtgName uses ordinal IDs. Keep the slot so other formulas and
                # opaque parts cannot accidentally start referring to a new name.
                name = f'_xlspy_deleted_{index}_{len(self.names)}'
                entry.name, entry.flags, entry.definition = name, 1, None
                scope_id = 0xFFFFFFFF if scope is None else self.sheets.index(scope)
                entry.raw = struct.pack('<IBI', 1, 0, scope_id) + wide_string(name) + parsed_formula(b'\x1c\x1d') + b'\xff' * 4
                return

    def name_payloads(self):
        index = 0
        while index < len(self.names):
            entry = self.names[index]
            if entry.definition is not None:
                definition = entry.definition
                scope_id = 0xFFFFFFFF if definition.scope is None else self.sheets.index(definition.scope)
                formula = compile_formula(definition.formula, self, sheet=definition.scope, name_formula=True)
                entry.raw = struct.pack('<IBI', entry.flags, 0, scope_id) + wide_string(definition.name) + formula + b'\xff' * 4
            if entry.raw is None:
                raise ValueError(f"formula references undefined name {entry.name!r}")
            index += 1
        return [entry.raw for entry in self.names]

    def defined_names(self):
        result = []
        for entry in self.names:
            if entry.flags & (0x20 | 0x20000) or entry.name.startswith('_xlspy_deleted_'):
                continue
            if entry.definition is not None:
                result.append(entry.definition)
            elif entry.raw is not None:
                _, offset = read_wide(entry.raw, 9)
                scope = self._scope(entry)
                result.append(DefinedName(entry.name, decode_formula(entry.raw[offset:], self, sheet=scope, name_formula=True), scope))
        return result

    def external_records(self):
        if not self.xti and self._external_block is None:
            return b''
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
            entries.append(struct.pack('<Iii', *entry))
        table = build_record(EXTERN_SHEET, struct.pack('<I', len(entries)) + b''.join(entries))
        if self._external_block is None:
            return build_record(BEGIN_EXTERNALS) + b''.join(build_record(kind) for kind in links) + table + build_record(END_EXTERNALS)
        block = self._external_block
        if len(links) > len(self.links):
            block = replace_records(block, set(), build_record(SUP_SELF), before=(EXTERN_SHEET, END_EXTERNALS))
        return replace_records(block, {EXTERN_SHEET}, table, before=(END_EXTERNALS,))

    def write_workbook(self, original):
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
                output.append(original[record.header_start:record.data_end])
        return b''.join(output)


@dataclass(frozen=True)
class ArrayValue:
    source_row: int
    source_column: int
    value: object = None


def formula_cell(formula, column, style, context, *, sheet=None, row=0):
    """Encode a scalar formula; cached zero is used if no cache was supplied."""
    array_definition = b''
    if isinstance(formula, ArrayValue):
        parsed = parsed_formula(b'\x01' + struct.pack('<I', formula.source_row), struct.pack('<I', formula.source_column))
    elif formula.array_ref is not None:
        c1, r1, c2, r2 = parse_cell_range(formula.array_ref)
        if (column, row) != (c1, r1):
            raise ValueError("array formula must be written in its top-left cell")
        parsed = parsed_formula(b'\x01' + struct.pack('<I', row), struct.pack('<I', column))
        group = compile_formula(formula.text, context, sheet=sheet, origin=(column, row))
        array_definition = build_record(0x1AA, struct.pack('<IIII', r1, r2, c1, c2) + b'\x01' + group)
    else:
        parsed = compile_formula(formula.text, context, sheet=sheet, origin=(column, row))
    header = struct.pack('<II', column, style)
    value = formula.value
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
        kind, cache = 9, struct.pack('<d', float(value or 0))
    else:
        raise TypeError(f"unsupported XLSB formula cached value {type(value).__name__}")
    return build_record(kind, header + cache + b'\x01\0' + parsed) + array_definition
