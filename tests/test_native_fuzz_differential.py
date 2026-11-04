"""Differential, structure-aware fuzzing: the native readers against the Python ones.

``test_native_fuzz.py`` corrupts random bytes and only checks that a damaged
part fails cleanly. Most such cases stop at the first record or tag. Here the
mutations keep the part's structure (whole XLSB records, XLSX cells, rows and
attributes), so they reach the cell decoders, and every case is read twice:
through the native module and through the pure-Python readers, in-process. Both
must refuse it (InvalidFileError, no panic) or both must return the same rows,
with the same types.

``XLSPY_FUZZ_ITERATIONS`` and ``XLSPY_FUZZ_SEED`` change the budget and the seed
(as in ``test_native_fuzz.py``); a failure message names the case's seed.
"""

import datetime as dt
import math
import os
import random
import re
import struct
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter, _accel
from xlspy import reader as reader_module
from xlspy.exceptions import InvalidFileError

pytestmark = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")

ITERATIONS = int(os.environ.get("XLSPY_FUZZ_ITERATIONS", "300"))
SEED = int(os.environ.get("XLSPY_FUZZ_SEED", "7"))

ROWS = [
    ["Name", "Value", "Flag", "When", "Note"],
    ["alpha", 1.5, True, dt.datetime(2024, 2, 29, 13, 30), "  padded  "],
    [None, 2, False, dt.date(1900, 3, 1), "line\nbreak"],
    ["gamma", -3, None, None, "zażółć 😀"],
    ["", 1e300, 0, dt.datetime(9999, 12, 31), "x" * 300],
    [None, None, None, None, None],
    ["delta", 123456789012, "TRUE", 0.5, "<&>\"'"],
]


def _write_fixture(path, kind, shared_strings):
    writer = XlsbWriter if kind == "xlsb" else XlsxWriter
    with writer(str(path), use_shared_strings=shared_strings) as workbook:
        workbook.add_sheet("Data")
        workbook.write_sheet(ROWS)


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {info.filename: archive.read(info.filename) for info in archive.infolist()}


def _save(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


# --- reading through either backend ------------------------------------------


def _read(path):
    """("ok", rows per sheet) or ("refused", None); a panic fails the test."""
    try:
        with ExcelReader(str(path)) as reader:
            return "ok", [reader.read_all(name) for name in reader.get_sheet_names()]
    except InvalidFileError as exc:
        assert type(exc.__cause__).__name__ != "PanicException", f"the native reader panicked: {exc}"
        return "refused", None


def _read_python(path, monkeypatch):
    with monkeypatch.context() as patch:
        patch.setattr(reader_module, "HAVE_NATIVE", False)
        patch.setattr(_accel, "_native_read_xlsx_worksheet", None)
        patch.setattr(_accel, "_native_read_xlsb_sst", None)
        patch.setattr(_accel, "_native_read_xlsx_sst", None)
        return _read(path)


def _normal(value):
    """A comparable form: NaN equals NaN, and the type is part of the value."""
    if isinstance(value, float) and math.isnan(value):
        return ("nan",)
    return (type(value).__name__, value)


def _grid(result):
    status, sheets = result
    if sheets is None:
        return status, None
    return status, [[[_normal(v) for v in row] for row in rows] for rows in sheets]


def _check_same(path, monkeypatch, case):
    native = _grid(_read(path))
    python = _grid(_read_python(path, monkeypatch))
    if native != python:
        detail = _first_difference(native, python)
        pytest.fail(f"native and Python readers disagree on {case}: {detail}")


def _first_difference(native, python):
    if native[0] != python[0]:
        return f"native {native[0]}, Python {python[0]}"
    for s, (left, right) in enumerate(zip(native[1], python[1])):
        if len(left) != len(right):
            return f"sheet {s}: native {len(left)} rows, Python {len(right)}"
        for r, (lrow, rrow) in enumerate(zip(left, right)):
            if lrow != rrow:
                return f"sheet {s} row {r}: native {lrow!r}, Python {rrow!r}"
    return f"sheet counts {len(native[1])} vs {len(python[1])}"


# --- XLSB: record-level mutations --------------------------------------------


def _vint(value):
    out = bytearray()
    while True:
        byte = value & 0x7F
        value >>= 7
        out.append(byte | 0x80 if value else byte)
        if not value:
            return bytes(out)


def _split_records(data):
    records, pos = [], 0
    while pos < len(data):
        ids = []
        for limit in (2, 4):
            value = shift = 0
            for _ in range(limit):
                byte = data[pos]
                pos += 1
                value |= (byte & 0x7F) << shift
                shift += 7
                if not byte & 0x80:
                    break
            ids.append(value)
        rec_id, length = ids
        records.append([rec_id, data[pos:pos + length]])
        pos += length
    return records


def _join_records(records):
    return b"".join(_vint(rec_id) + _vint(len(body)) + bytes(body) for rec_id, body in records)


# Cell records (blank, RK, error, bool, real, string, SST) and their formula
# forms, the row header, and the sheet-data frame.
_CELL_IDS = [0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B]
_ROW_HDR, _BEGIN_DATA, _END_DATA = 0x00, 0x91, 0x92
_EDGE_U32 = [0, 1, 2, 0x3FFF, 0x4000, 16383, 16384, 1048575, 1048576, 0x7FFFFFFF, 0xFFFFFFFF]


def _mutate_record(rec, rng):
    rec_id, body = rec[0], bytearray(rec[1])
    choice = rng.randrange(7)
    if choice == 0 and body:
        for _ in range(rng.randint(1, 4)):
            body[rng.randrange(len(body))] = rng.randrange(256)
    elif choice == 1:
        # A short or a long body for the record's type.
        if body and rng.random() < 0.6:
            del body[rng.randrange(len(body)):]
        else:
            body += bytes(rng.randrange(256) for _ in range(rng.randint(1, 12)))
    elif choice == 2 and len(body) >= 4:
        # The first u32: a cell's column, a row header's row index.
        struct.pack_into("<I", body, 0, rng.choice(_EDGE_U32))
    elif choice == 3 and len(body) >= 8:
        # The second u32: a cell's style index.
        struct.pack_into("<I", body, 4, rng.choice(_EDGE_U32 + [rng.randrange(64)]))
    elif choice == 4 and rec_id in _CELL_IDS:
        rec_id = rng.choice(_CELL_IDS)
    elif choice == 5 and len(body) >= 12:
        # The value: an RK, a string length, an SST index or the start of a double.
        if rng.random() < 0.5:
            struct.pack_into("<I", body, 8, rng.choice(_EDGE_U32 + [rng.randrange(16)]))
        elif len(body) >= 16:
            value = rng.choice([math.nan, math.inf, -math.inf, -0.0, 5e-324, 1.7976931348623157e308,
                                2958465.9999999, 2958466.0, 60.0, -1.0])
            struct.pack_into("<d", body, 8, value)
    elif len(body) >= 12 and rec_id in (0x06, 0x08):
        # A string's character count past (or inside) its body.
        struct.pack_into("<I", body, 8, rng.choice([0, 1, len(body), 0x7FFFFFFF, 0xFFFFFFFF]))
    return [rec_id, bytes(body)]


def _mutate_xlsb_sheet(data, rng):
    records = _split_records(data)
    for _ in range(rng.randint(1, 3)):
        operation = rng.randrange(6)
        cells = [i for i, (rec_id, _b) in enumerate(records) if rec_id in _CELL_IDS or rec_id == _ROW_HDR]
        if operation <= 2 and cells:
            i = rng.choice(cells)
            records[i] = _mutate_record(records[i], rng)
        elif operation == 3 and cells:
            del records[rng.choice(cells)]
        elif operation == 4 and cells:
            i = rng.choice(cells)
            records.insert(rng.randrange(len(records) + 1), list(records[i]))
        elif cells:
            # Records out of order, or a cell outside the sheet-data frame.
            i, j = rng.choice(cells), rng.randrange(len(records))
            records[i], records[j] = records[j], records[i]
    return _join_records(records)


def _mutate_xlsb_sst(data, rng):
    records = _split_records(data)
    for _ in range(rng.randint(1, 2)):
        i = rng.randrange(len(records))
        records[i] = _mutate_record(records[i], rng)
        if rng.random() < 0.3:
            del records[rng.randrange(len(records))]
    return _join_records(records)


# --- XLSX: cell, row and attribute mutations ---------------------------------

_REFS = ["A1", "B2", "XFD1", "XFE1", "A0", "A1048576", "A1048577", "a1", "1A", "", "AAAA1", "C99999999999"]
_TYPES = ["s", "str", "inlineStr", "b", "e", "n", "d", "", "x"]
_VALUES = ["0", "1", "-1", "2", "1e308", "1e999", "nan", "inf", "", "  7 ", "0x10", "#N/A", "TRUE",
           "99999999", "4294967296", "-0", "1.5E-3", "2024-02-29T13:30:00", "abc", "&amp;", "&#65;",
           "&#x0;", "&bogus;", "<![CDATA[c]]>", "é", "\U0001F600"]


def _mutate_xlsx_cell(cell, rng):
    choice = rng.randrange(6)
    if choice == 0:
        ref = rng.choice(_REFS)
        cell = re.sub(r' r="[^"]*"', f' r="{ref}"', cell, count=1) if ' r="' in cell else cell
    elif choice == 1:
        new = rng.choice(_TYPES)
        if ' t="' in cell:
            cell = re.sub(r' t="[^"]*"', f' t="{new}"' if new else "", cell, count=1)
        else:
            cell = cell.replace("<c ", f'<c t="{new}" ', 1)
    elif choice == 2:
        style = rng.choice(["0", "1", "2", "999", "-1", "4294967296", "x"])
        if ' s="' in cell:
            cell = re.sub(r' s="[^"]*"', f' s="{style}"', cell, count=1)
        else:
            cell = cell.replace("<c ", f'<c s="{style}" ', 1)
    elif choice == 3 and "<v>" in cell:
        cell = re.sub(r"<v>[^<]*</v>", f"<v>{rng.choice(_VALUES)}</v>", cell, count=1)
    elif choice == 4:
        # An inline string, or a value next to one.
        cell = re.sub(r"(<c [^>]*?)/?>.*$", r"\1>", cell, flags=re.S)
        cell += rng.choice(["<is><t>in</t></is></c>", "<is><r><t>a</t></r><r><t>b</t></r></is></c>",
                            "<v>1</v><is><t>x</t></is></c>", "<f>1+1</f><v>2</v></c>", "</c>",
                            "<is><t xml:space=\"preserve\"> s </t></is></c>"])
    else:
        # Cut the cell anywhere: a broken tag.
        cell = cell[: rng.randrange(len(cell) + 1)]
    return cell


_CELL = re.compile(r"<c [^>]*?(?:/>|>.*?</c>)", re.S)
_ROW = re.compile(r"<row [^>]*?(?:/>|>.*?</row>)", re.S)


def _mutate_xlsx_sheet(data, rng):
    text = data.decode("utf-8")
    for _ in range(rng.randint(1, 3)):
        operation = rng.randrange(5)
        cells = list(_CELL.finditer(text))
        rows = list(_ROW.finditer(text))
        if operation <= 2 and cells:
            m = rng.choice(cells)
            text = text[:m.start()] + _mutate_xlsx_cell(m.group(0), rng) + text[m.end():]
        elif operation == 3 and rows:
            m = rng.choice(rows)
            row = m.group(0)
            choice = rng.randrange(3)
            if choice == 0:
                row = re.sub(r' r="\d+"', f' r="{rng.choice(["0", "1", "3", "1048577", "x", ""])}"', row, count=1)
            elif choice == 1:
                row = row + row  # a repeated row
            else:
                row = re.sub(r' spans="[^"]*"', f' spans="{rng.choice(["1:1", "5:2", "x", "1:20000"])}"', row)
            text = text[:m.start()] + row + text[m.end():]
        elif rows and len(rows) > 1:
            # Rows out of order.
            a, b = rng.sample(rows, 2)
            first, second = sorted((a, b), key=lambda m: m.start())
            text = (text[:first.start()] + second.group(0) + text[first.end():second.start()]
                    + first.group(0) + text[second.end():])
    return text.encode("utf-8")


_SI = re.compile(r"<si>.*?</si>", re.S)


def _mutate_xlsx_sst(data, rng):
    text = data.decode("utf-8")
    items = list(_SI.finditer(text))
    m = rng.choice(items)
    choice = rng.randrange(5)
    if choice == 0:
        new = f"<si><t>{rng.choice(_VALUES)}</t></si>"
    elif choice == 1:
        new = "<si><r><rPr><b/></rPr><t>bo</t></r><r><t xml:space=\"preserve\"> ld </t></r></si>"
    elif choice == 2:
        new = "<si><t>base</t><rPh sb=\"0\" eb=\"1\"><t>ph</t></rPh></si>"
    elif choice == 3:
        new = ""  # one string fewer than the count says
    else:
        new = m.group(0)[: rng.randrange(len(m.group(0)) + 1)]
    text = text[:m.start()] + new + text[m.end():]
    if rng.random() < 0.3:
        text = re.sub(r' (count|uniqueCount)="\d+"', lambda g: f' {g.group(1)}="{rng.choice([0, 1, 99999999])}"', text)
    return text.encode("utf-8")


# --- the tests ---------------------------------------------------------------

_SHEET = {"xlsb": "xl/worksheets/sheet1.bin", "xlsx": "xl/worksheets/sheet1.xml"}
_SST = {"xlsb": "xl/sharedStrings.bin", "xlsx": "xl/sharedStrings.xml"}
_MUTATORS = {
    ("xlsb", "sheet"): _mutate_xlsb_sheet,
    ("xlsx", "sheet"): _mutate_xlsx_sheet,
    ("xlsb", "sst"): _mutate_xlsb_sst,
    ("xlsx", "sst"): _mutate_xlsx_sst,
}


@pytest.mark.parametrize("kind", ["xlsb", "xlsx"])
@pytest.mark.parametrize("shared", [True, False], ids=["sst", "inline"])
def test_unmutated_fixture_reads_the_same(tmp_path, monkeypatch, kind, shared):
    path = tmp_path / f"seed.{kind}"
    _write_fixture(path, kind, shared)
    native = _read(path)
    assert native[0] == "ok"
    _check_same(path, monkeypatch, "the unmutated fixture")


@pytest.mark.parametrize("kind,part", [("xlsb", "sheet"), ("xlsx", "sheet"), ("xlsb", "sst"), ("xlsx", "sst")])
def test_mutated_parts_read_the_same_on_both_backends(tmp_path, monkeypatch, kind, part):
    source = tmp_path / f"seed.{kind}"
    _write_fixture(source, kind, shared_strings=True)
    parts = _parts(source)
    name = (_SHEET if part == "sheet" else _SST)[kind]
    original = parts[name]
    mutate = _MUTATORS[kind, part]
    target = tmp_path / f"case.{kind}"
    for case in range(ITERATIONS):
        rng = random.Random(f"{SEED}:{kind}:{part}:{case}")
        parts[name] = mutate(original, rng)
        _save(target, parts)
        _check_same(target, monkeypatch, f"{kind} {part} case {case} (seed {SEED})")
