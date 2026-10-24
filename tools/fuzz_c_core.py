"""Fuzz the C extension's readers and row encoders.

Three phases, all deterministic for a given ``--seed``:

  reader   Corrupts valid XLSB worksheet and shared-strings parts (flipped
           bytes, oversized length fields, truncation, inserted garbage) and
           feeds them to the C readers. A Python exception is an acceptable
           answer for malformed input; a crash or a sanitizer report is not.

  encoder  Builds rows from awkward values (lone surrogates, text over the
           cell limit, out-of-range numbers and dates, objects whose hash or
           str() raises) and feeds them to the C row encoders. After every
           call the shared-string table must match the counters the caller
           holds: a call that fails or falls back must leave it unchanged.
           Successful XLSX rows must also parse as XML.

  xlsx     Generates worksheet XML (attribute order, duplicate attributes,
           missing and malformed references, every cell type, entities,
           rich text, prefixes, odd numbers and dates) and reads it through
           the C scanner, the Python regex path and ElementTree. Rows and
           values must agree; a row the scanner declines must surface as raw
           bytes. For malformed input both paths may raise, but one raising
           while the other returns rows is a violation.

Usage:
    python tools/fuzz_c_core.py --iterations 20000 --seed 1

Run it against a sanitizer build to catch memory errors; see docs/validation.md.
The process exits 1 at the first violated check and prints the offending case.
"""

import argparse
import datetime
import io
import pathlib
import random
import sys
import tempfile
import xml.etree.ElementTree as ElementTree
import zipfile
from collections import Counter
from decimal import Decimal
from types import SimpleNamespace

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from xlspy import ExcelReader, Formula, Hyperlink, XlsbWriter, XlsxWriter  # noqa: E402
from xlspy import _accel, _xlsx_read  # noqa: E402
from xlspy.reader import _get_date_num_fmt_ids  # noqa: E402

if not _accel.HAVE_C_EXT:
    sys.exit("the C extension is not built; nothing to fuzz")


class BadHash(str):
    """A str whose hash raises, so dict lookups fail mid-row."""

    def __hash__(self):
        raise RuntimeError("hash failed on purpose")


class BadStr:
    """An object whose str() raises, so the fallback path fails."""

    def __str__(self):
        raise RuntimeError("str failed on purpose")


AWKWARD_VALUES = [
    None, True, False, 0, -1, 2**31, -2**31 - 1, 2**63 - 1, -2**63, 2**63, 2**70, -2**70,
    0.0, -0.0, 1.5, float("nan"), float("inf"), -float("inf"), 1e308,
    "", "a", "\ud800", "x\udfff", "\U0001D11E", "\U0001D11E" * 16384, "a" * 32767, "a" * 32768,
    "a\x00b", "\x0b", "\r", "_x0041_", "_x005F_x", "￾", "￿", " ", "\x7f", "<&>",
    "tab\tnl\n",
    Decimal("1.5"), Decimal("NaN"), Decimal("1e400"), Decimal("-0"),
    datetime.date(1899, 12, 31), datetime.date(1900, 1, 1), datetime.date(9999, 12, 31),
    datetime.datetime(1, 1, 1), datetime.datetime(2024, 2, 29, 23, 59, 59, 999999),
    BadHash("hashy"), BadStr(), Formula("A1+1"), Hyperlink("https://example.com", text="link"),
    ("value", 3), [1, 2], object(),
]


def sample_workbook(shared):
    """A small XLSB workbook covering the common cell types."""
    buffer = io.BytesIO()
    rows = [["Name", "Qty", "Price", "Flag", "Day", "Note"]]
    for i in range(1, 40):
        rows.append([
            f"item {i % 7}", i, i / 7, i % 2 == 0,
            datetime.date(2024, 1, 1) + datetime.timedelta(days=i),
            f"né {i} \U0001D11E",
        ])
    with XlsbWriter(buffer, useSharedStrings=shared) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows)
    return buffer.getvalue()


def workbook_parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = archive.namelist()
        sheet = next(n for n in names if n.startswith("xl/worksheets/sheet"))
        sheet_bytes = archive.read(sheet)
        sst_bytes = archive.read("xl/sharedStrings.bin") if "xl/sharedStrings.bin" in names else b""
    return sheet_bytes, sst_bytes


def reader_state(data):
    """The shared strings, cell styles and date formats a reader passes to C."""
    with tempfile.TemporaryDirectory() as tmp:
        path = pathlib.Path(tmp) / "state.xlsb"
        path.write_bytes(data)
        with ExcelReader(str(path)) as reader:
            return (list(reader._shared_string_array), list(reader._styles_cell_xfs_array),
                    _get_date_num_fmt_ids(reader._number_formats_type_dict), reader._date_1904)


def mutate(rng, data):
    buf = bytearray(data)
    for _ in range(rng.randint(1, 6)):
        if not buf:
            buf += rng.randbytes(8)
            continue
        pos = rng.randrange(len(buf))
        op = rng.randrange(6)
        if op == 0:
            buf[pos] = rng.randrange(256)                       # flip one byte
        elif op == 1:
            buf[pos:pos + 4] = b"\xff\xff\xff\xff"              # maximal length field
        elif op == 2:
            buf[pos:pos + 4] = rng.randrange(2**32).to_bytes(4, "little")
        elif op == 3:
            del buf[pos:pos + rng.randint(1, 64)]               # drop a run of bytes
        elif op == 4:
            buf[pos:pos] = rng.randbytes(rng.randint(1, 16))    # insert garbage
        else:
            buf[pos] = 0xFF                                     # unterminated varint
    if rng.random() < 0.1:
        del buf[rng.randrange(len(buf) + 1):]                  # truncate
    return bytes(buf)


def reader_phase(rng, iterations, state, sheets, ssts, counts):
    shared_strings, styles, date_fmts, date_1904 = state
    for _ in range(iterations):
        counts["reader calls"] += 1
        try:
            rows = _accel.read_xlsb_worksheet(mutate(rng, rng.choice(sheets)), shared_strings,
                                              styles, date_fmts, date_1904)
            # The C reader is lazy: parsing happens while the rows are consumed.
            for _ in rows or ():
                counts["reader rows"] += 1
            counts["reader ok"] += 1
        except Exception as exc:  # malformed input may raise; that is allowed
            counts[f"reader raised {type(exc).__name__}"] += 1
        try:
            _accel.read_xlsb_shared_strings(mutate(rng, rng.choice(ssts)))
            counts["sst ok"] += 1
        except Exception as exc:
            counts[f"sst raised {type(exc).__name__}"] += 1


def random_row(rng):
    return [rng.choice(AWKWARD_VALUES) for _ in range(rng.randint(0, 6))]


def _table_consistent(ss_dict, ss_list, uniq):
    if len(ss_list) != uniq or len(ss_dict) != uniq:
        return f"table has {len(ss_list)} strings and {len(ss_dict)} keys, counter says {uniq}"
    for index, text in enumerate(ss_list):
        if ss_dict.get(text) != index:
            return f"string at index {index} maps to {ss_dict.get(text)!r}"
    return None


def encoder_phase(rng, iterations, counts):
    """Every encode either succeeds, or fails or falls back leaving the table as it was."""
    ss_dict, ss_list, uniq, all_ = {}, [], 0, 0
    for n in range(iterations):
        counts["encoder calls"] += 1
        if n % 50 == 0:  # restart the table now and then, keeping checks cheap
            ss_dict, ss_list, uniq, all_ = {}, [], 0, 0
        row = random_row(rng)
        style_ids = rng.choice([None, 0, 5, [1, 2, 3], ((0, 4),)])

        # Shared-string XLSB row. On None or an error the caller's counters are
        # still the old ones, so the table must still match them.
        try:
            out = _accel.encode_xlsb_row(row, ss_dict, ss_list, uniq, all_, n, style_ids)
        except Exception as exc:
            out = None
            counts[f"xlsb shared raised {type(exc).__name__}"] += 1
        if out is None:
            counts["xlsb shared fallback or error"] += 1
        else:
            _, uniq, all_ = out
            counts["xlsb shared ok"] += 1
        if problem := _table_consistent(ss_dict, ss_list, uniq):
            return _violation("xlsb shared", row, problem)

        # Inline XLSB row: takes no table.
        try:
            _accel.encode_xlsb_row_inline(row, n, style_ids)
            counts["xlsb inline ok"] += 1
        except Exception:
            counts["xlsb inline raised"] += 1

        # XLSX row, shared and inline.
        for use_shared in (True, False):
            try:
                out = _accel.encode_xlsx_row(row, ss_dict, ss_list, uniq, all_, n, use_shared)
            except Exception:
                out = None
                counts["xlsx raised"] += 1
            if out is None:
                counts["xlsx fallback"] += 1
            else:
                xml, uniq, all_ = out
                counts["xlsx ok"] += 1
                try:
                    ElementTree.fromstring(xml)
                except ElementTree.ParseError as exc:
                    return _violation(f"xlsx shared={use_shared}", row, f"row XML does not parse: {exc}")
            if problem := _table_consistent(ss_dict, ss_list, uniq):
                return _violation(f"xlsx shared={use_shared}", row, problem)

        # Column widths over the same rows.
        try:
            _accel.calc_column_widths([row], len(row))
            counts["widths ok"] += 1
        except Exception:
            counts["widths raised"] += 1

        # Row pre-scan and sharedStrings.bin writer. The table holds only strings
        # that were encoded successfully, so the writer's input is always valid.
        try:
            _accel.xlsb_row_has_formatted(row)
            counts["row formatted ok"] += 1
        except Exception:
            counts["row formatted raised"] += 1
        try:
            _accel.encode_xlsb_sst(ss_list, all_, uniq)
            counts["sst encode ok"] += 1
        except Exception:
            counts["sst encode raised"] += 1
    return None


def _violation(encoder, row, problem):
    return f"{encoder}: {problem}\n  row = {row!r}"[:2000]


class ChoppyStream:
    """A stream whose read() returns at most ``size`` bytes per call."""

    def __init__(self, data, size):
        self.data = data
        self.size = size
        self.pos = 0

    def read(self, size):
        end = min(len(self.data), self.pos + min(size, self.size))
        chunk = self.data[self.pos:end]
        self.pos = end
        return chunk


XLSX_NUMBERS = ["1", "0", "-3", "0.25", "1e3", "1E+3", "1e999", " 5", "1_0",
                "12345678901234567890", "abc", "", "NaN", "1.5e-7", "+7"]
XLSX_SHARED_INDEXES = ["0", "1", "2", "99", "-1", "", "abc"]
XLSX_REFS = ["A{r}", "B{r}", "AA{r}", "ZZZ{r}", "a{r}", "A{r}B", "A", "1", "", "{r}"]
XLSX_ISO = ["2024-02-29T12:30:00", "2024-01-01", "nope", ""]
XLSX_PREFIX = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _xlsx_attrs(rng, ref, style, kind):
    pairs = []
    if kind is not None:
        pairs.append(f't="{kind}"')
    if ref is not None:
        pairs.append(f'r="{ref}"')
    if style:
        pairs.append(f's="{style}"')
    rng.shuffle(pairs)
    text = "".join(" " + pair for pair in pairs)
    # `_CELL` allows trailing whitespace before `/>` and `>`; the scanner must
    # stay on the fast path for those forms instead of delegating the row.
    return text + (" " if rng.random() < 0.25 else "")


def _xlsx_cell(rng, row):
    ref = rng.choice(XLSX_REFS).format(r=row)
    style = rng.choice(["", "", "0", "1", "2", "99", "-1", "-7", "abc"])
    kind = rng.randrange(10)
    if kind == 0:
        return f'<c{_xlsx_attrs(rng, ref, style, None)}><v>{rng.choice(XLSX_NUMBERS)}</v></c>'
    if kind == 1:
        return (f'<c{_xlsx_attrs(rng, ref, style, "s")}>'
                f'<v>{rng.choice(XLSX_SHARED_INDEXES)}</v></c>')
    if kind == 2:
        return (f'<c{_xlsx_attrs(rng, ref, style, "b")}>'
                f'<v>{rng.choice(["1", "0", "", "true"])}</v></c>')
    if kind == 3:
        # Error values are fixed codes; `_fast_value` keeps entities raw for
        # t="e", which predates the C scanner, so do not generate `&` here.
        return (f'<c{_xlsx_attrs(rng, ref, style, "e")}>'
                f'<v>{rng.choice(["#DIV/0!", "#N/A", ""])}</v></c>')
    if kind == 4:
        text = rng.choice(["plain", "a &lt; b", "a&amp;b", "_x0041_", "line\rbreak", ""])
        return f'<c{_xlsx_attrs(rng, ref, style, "str")}><v>{text}</v></c>'
    if kind == 5:
        inline = rng.choice([
            "<is><t>plain</t></is>",
            '<is><t xml:space="preserve"> spaced </t></is>',
            "<is><t></t></is>",
            "<is><t>a &amp; b</t></is>",
            "<is><r><t>a</t></r><r><t>b</t></r></is>",
            "<is><t>x_x000D_y</t></is>",
            "<is><t>a\rb</t></is>",
            "<is> <t>x</t> </is>",
        ])
        return f'<c{_xlsx_attrs(rng, ref, style, "inlineStr")}>{inline}</c>'
    if kind == 6:
        return f'<c{_xlsx_attrs(rng, ref, style, "d")}><v>{rng.choice(XLSX_ISO)}</v></c>'
    if kind == 7:
        formula = rng.choice([
            "<f>1+1</f><v>2</v>",
            '<f t="shared" ref="A1:A9" si="0">B1*2</f><v>4</v>',
            '<f t="shared" si="0"/><v>6</v>',
            '<f t="array" ref="A1:A2">1</f><v>1</v>',
            "<f>1+1</f>",
        ])
        return f'<c{_xlsx_attrs(rng, ref, style, None)}>{formula}</c>'
    if kind == 8:
        extra = rng.choice(["<extLst/>", "<v>1</v><v>2</v>", "<is><t>a</t></is>"])
        return f'<c{_xlsx_attrs(rng, ref, style, None)}><v>1</v>{extra}</c>'
    return f'<c{_xlsx_attrs(rng, ref, style, None)}/>'


def _xlsx_row(rng, index):
    number = rng.choice([str(index), str(index), "", f"{index} ", f"0{index}", f"+{index}"])
    attrs = f' r="{number}"' if number else ""
    if rng.random() < 0.3:
        attrs += ' spans="1:3"'
    if rng.random() < 0.1:
        attrs += ' note="é"'
    cells = "".join(_xlsx_cell(rng, index) for _ in range(rng.randrange(0, 4)))
    if not cells:
        return f"<row{attrs}/>" if rng.random() < 0.5 else f"<row{attrs}></row>"
    return f"<row{attrs}>{cells}</row>"


def _xlsx_sheet(rng):
    body = "".join(_xlsx_row(rng, index) for index in range(1, rng.randrange(1, 6)))
    if rng.random() < 0.2:
        body += rng.choice([
            '<row r="9"><c r="A9"><v>1</v></c><x:c r="B9" xmlns:x="' + XLSX_PREFIX +
            '"><x:v>2</x:v></x:c></row>',
            '<row r="9"><c r="A9"><v>1</v></c></row><extLst/>',
        ])
    return ('<?xml version="1.0" encoding="UTF-8"?>'
            f'<worksheet xmlns="{XLSX_PREFIX}">'
            f"<sheetData>{body}</sheetData></worksheet>").encode("utf-8")


def xlsx_reader_state():
    """Shared strings, styles, formats and flags a reader passes to the scanner."""
    with tempfile.TemporaryDirectory() as tmp:
        path = pathlib.Path(tmp) / "state.xlsx"
        with XlsxWriter(path) as writer:
            writer.write_sheet([["text", 1, 2.5, True, datetime.datetime(2024, 1, 5)]])
        with ExcelReader(str(path)) as reader:
            return (list(reader._shared_string_array),
                    list(reader._styles_cell_xfs_array),
                    dict(reader._number_formats_type_dict),
                    reader._date_1904,
                    _xlsx_read._date_style_flags(reader))


def _xlsx_stub(state):
    shared, styles, formats, date_1904, _flags = state
    return SimpleNamespace(_shared_string_array=shared, _styles_cell_xfs_array=styles,
                           _number_formats_type_dict=formats, _date_1904=date_1904)


class _FallbackCounter:
    """Wraps the C iterator to count rows it declines, keeping resume()."""

    def __init__(self, rows):
        self._rows = rows
        self.raw = 0

    def __iter__(self):
        for item in self._rows:
            if isinstance(item, bytes):
                self.raw += 1
            yield item

    def resume(self, row_number):
        self._rows.resume(row_number)


def _xlsx_c_rows(state, head, rest, size, probe):
    shared, _styles, _formats, date_1904, flags = state
    rows = _accel.read_xlsx_worksheet(ChoppyStream(rest, size), head, shared, flags, date_1904)
    counter = _FallbackCounter(rows)
    resolved = list(_xlsx_read._c_rows(_xlsx_stub(state), counter, probe[2], probe[3]))
    return resolved, counter.raw


def _xlsx_python_rows(state, rest, probe, size):
    return list(_xlsx_read._iter_fast(_xlsx_stub(state), ChoppyStream(rest, size), probe))


def _xlsx_tree_rows(state, data):
    shared, styles, formats, date_1904, _flags = state
    reader = SimpleNamespace(_shared_string_array=shared, _styles_cell_xfs_array=styles,
                             _number_formats_type_dict=formats, _date_1904=date_1904)
    root = ElementTree.fromstring(data)
    from xlspy._metadata import namespace_of
    ns = namespace_of(root.tag) or XLSX_PREFIX
    rows = []
    next_row = 1
    for node in root.findall(f"{{{ns}}}sheetData/{{{ns}}}row"):
        number, cells = _xlsx_read._row_from_element(reader, node, ns, next_row, True, None)
        next_row = number + 1
        rows.append((number, cells))
    return rows


def _xlsx_same(left, right):
    if type(left) is not type(right):
        return f"{type(left).__name__} != {type(right).__name__}: {left!r} / {right!r}"
    if isinstance(left, list):
        if len(left) != len(right):
            return f"length {len(left)} != {len(right)}"
        for item_left, item_right in zip(left, right):
            problem = _xlsx_same(item_left, item_right)
            if problem:
                return problem
        return None
    if isinstance(left, tuple):
        if len(left) != len(right):
            return f"tuple length {len(left)} != {len(right)}"
        for item_left, item_right in zip(left, right):
            problem = _xlsx_same(item_left, item_right)
            if problem:
                return problem
        return None
    return None if left == right else f"{left!r} != {right!r}"


def xlsx_reader_phase(rng, iterations, state, counts):
    for _ in range(iterations):
        counts["xlsx calls"] += 1
        data = _xlsx_sheet(rng)
        probe_stream = ChoppyStream(data, rng.choice([1, 2, 5, 13, 64, 1 << 18]))
        probe = _xlsx_read._probe_fast(probe_stream)
        if probe is None:
            counts["xlsx tree only"] += 1
            continue
        head = probe[4]
        rest = data[probe_stream.pos:]

        expected, got = None, None
        try:
            expected = _xlsx_python_rows(state, rest, probe, rng.choice([1, 7, 4096]))
            counts["xlsx python ok"] += 1
        except Exception as exc:
            expected = exc
            counts[f"xlsx python raised {type(exc).__name__}"] += 1
        try:
            got, raw = _xlsx_c_rows(state, head, rest, rng.choice([1, 3, 64]), probe)
            counts[f"xlsx fallback rows {raw > 0}"] += 1
            counts["xlsx c ok"] += 1
        except Exception as exc:
            got = exc
            counts[f"xlsx c raised {type(exc).__name__}"] += 1

        if isinstance(expected, Exception) or isinstance(got, Exception):
            if isinstance(expected, Exception) and isinstance(got, Exception):
                # Both reject the sheet; the C scanner reports the first cell
                # it cannot read while the regex path may raise from a later
                # matched cell, so the exception types need not match.
                if type(expected) is type(got):
                    counts["xlsx both raised same type"] += 1
                else:
                    counts["xlsx both raised, types differ"] += 1
                continue
            return f"xlsx: Python {expected!r}, C {got!r}\n  sheet = {data!r}"[:2000]
        if problem := _xlsx_same(got, expected):
            return f"xlsx: {problem}\n  sheet = {data!r}"[:2000]
        if rng.random() < 0.2:
            try:
                tree = _xlsx_tree_rows(state, data)
            except Exception:
                # ElementTree rejects some references the fast path accepts
                # (e.g. a column past Excel's limit); nothing to compare.
                counts["xlsx tree raised"] += 1
                continue
            if problem := _xlsx_same(got, tree):
                return f"xlsx vs tree: {problem}\n  sheet = {data!r}"[:2000]
            counts["xlsx tree compared"] += 1
    return None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--iterations", type=int, default=2000,
                        help="mutations per reader input and rows per encoder phase (default 2000)")
    parser.add_argument("--seed", type=int, default=1)
    args = parser.parse_args(argv)

    rng = random.Random(args.seed)
    sheets, ssts = [], []
    for shared in (True, False):
        sheet, sst = workbook_parts(sample_workbook(shared))
        sheets.append(sheet)
        if sst:
            ssts.append(sst)
    state = reader_state(sample_workbook(True))
    counts = Counter()

    reader_phase(rng, args.iterations, state, sheets, ssts, counts)
    failure = encoder_phase(rng, args.iterations, counts)
    if failure is None:
        failure = xlsx_reader_phase(rng, args.iterations, xlsx_reader_state(), counts)

    for name in sorted(counts):
        print(f"{name:<40}{counts[name]:>10}")
    if failure:
        print(f"\nFAILED (seed {args.seed}): {failure}")
        return 1
    print(f"\nall checks held (seed {args.seed}, {args.iterations} iterations per phase)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
