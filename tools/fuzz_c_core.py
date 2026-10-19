"""Fuzz the C extension's readers and row encoders.

Two phases, both deterministic for a given ``--seed``:

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

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from xlspy import ExcelReader, Formula, Hyperlink, XlsbWriter  # noqa: E402
from xlspy import _accel  # noqa: E402
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

    for name in sorted(counts):
        print(f"{name:<40}{counts[name]:>10}")
    if failure:
        print(f"\nFAILED (seed {args.seed}): {failure}")
        return 1
    print(f"\nall checks held (seed {args.seed}, {args.iterations} iterations per phase)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
