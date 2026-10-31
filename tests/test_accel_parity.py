"""Differential tests: the Rust backend (``xlspy._rs_core``) against the C one.

Every case runs through both modules and must give the same outcome: an
equal result of the same type, or an exception of the same class.  Functions
not ported to Rust yet are skipped.  Run with more examples for a deeper
check: ``pytest tests/test_accel_parity.py --hypothesis-seed=N``.
"""

import datetime
from decimal import Decimal

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

import xlspy._accel  # noqa: F401  (registers CellError and Style on both backends)
from xlspy import Style
from xlspy._errors import CellError

c_core = pytest.importorskip("xlspy._c_core")
rs_core = pytest.importorskip("xlspy._rs_core")

for _module in (c_core, rs_core):
    if hasattr(_module, "set_error_type"):
        _module.set_error_type(CellError)
        _module.set_style_type(Style)


def ported(name):
    return pytest.mark.skipif(not hasattr(rs_core, name), reason=f"{name} not ported to Rust yet")


def outcome(func, *args):
    try:
        result = func(*args)
    except Exception as exc:  # noqa: BLE001 - the exception class is the outcome
        return ("raises", type(exc))
    return ("returns", type(result), result)


def same(name, *args, fresh=None):
    """Both backends agree on ``name(*args)``; returns the shared outcome.
    ``fresh`` builds the arguments anew for each backend (iterators, mutated
    tables)."""
    expected = outcome(getattr(c_core, name), *(fresh() if fresh else args))
    actual = outcome(getattr(rs_core, name), *(fresh() if fresh else args))
    assert actual == expected, f"{name}{args!r}: C {expected!r} != Rust {actual!r}"
    return expected


class BadStr:
    def __str__(self):
        raise RuntimeError("str failed on purpose")


class BadLenTuple(tuple):
    def __len__(self):
        raise RuntimeError("len failed on purpose")


class BadItemTuple(tuple):
    def __getitem__(self, index):
        raise RuntimeError("getitem failed on purpose")


class Timestampish(datetime.datetime):
    pass


TEXTS = [
    "", "a", "plain ascii", "zażółć gęślą jaźń", "日本語", "emoji \U0001F600 here",
    "line one\nline two", "\x00\x07 controls", "a" * 32767, "\U0001F600" * 16383,
]


# -- shared strings ---------------------------------------------------------

@ported("encode_xlsb_sst")
@pytest.mark.parametrize("strings", [[], TEXTS, ["x"] * 3, ["\ud800"], ["ok", "bad \udfff"],
                                     ["😀"], ["a" * 32768], ["\U0001F600" * 16384],
                                     ["ok", 1], [None], ("tuple", "not list")])
@pytest.mark.parametrize("counts", [(0, 0), (5, 3), (2**32 - 1, 7)])
def test_encode_sst(strings, counts):
    same("encode_xlsb_sst", strings, *counts)


@ported("encode_xlsb_sst")
def test_encode_sst_counts_wrap_to_32_bits():
    # C's "I" format wraps out-of-range counts (from Python 3.15 with a
    # DeprecationWarning); Rust wraps without one.
    assert rs_core.encode_xlsb_sst([], 2**32 + 7, -1)[3:11] == bytes([7, 0, 0, 0, 255, 255, 255, 255])


@ported("encode_xlsb_sst")
def test_encode_sst_bad_counts():
    same("encode_xlsb_sst", ["a"], "1", 1)
    same("encode_xlsb_sst", ["a"], 1.5, 1)


def _sst_part(strings):
    return c_core.encode_xlsb_sst(strings, len(strings), len(strings))


@ported("read_xlsb_shared_strings")
@pytest.mark.parametrize("data", [
    b"", b"\x9f", _sst_part([]), _sst_part(TEXTS), _sst_part(["only"]),
    # No BrtBeginSst: the first item is dropped, like the Python reader.
    _sst_part(["a", "b", "c"])[11:],
    # Truncated BrtBeginSst, truncated item, odd string bytes, overrunning length.
    b"\x9f\x01\x02\x00\x00", b"\x9f\x01\x08" + b"\0" * 8 + b"\x13\x02\x00\x00",
    b"\x9f\x01\x08" + b"\0" * 8 + b"\x13\x08\x00\x02\x00\x00\x00a\x00b",
    b"\x9f\x01\x08" + b"\0" * 8 + b"\x13\x07\x00\xff\xff\xff\xffa\x00",
    # Lone surrogate (surrogatepass) and a 5-byte varint record length.
    b"\x9f\x01\x08" + b"\0" * 8 + b"\x13\x07\x00\x01\x00\x00\x00\x00\xd8",
    b"\x9f\x01\x08" + b"\0" * 8 + b"\x13\x80\x80\x80\x80\x01",
])
def test_read_sst(data):
    same("read_xlsb_shared_strings", data)
    same("read_xlsb_shared_strings", bytearray(data))
    same("read_xlsb_shared_strings", memoryview(data))


@ported("read_xlsb_shared_strings")
def test_read_sst_rejects_non_buffers():
    same("read_xlsb_shared_strings", "text")
    same("read_xlsb_shared_strings", None)


@ported("read_xlsb_shared_strings")
@settings(max_examples=400, deadline=None)
@given(st.lists(st.text(max_size=20), max_size=8), st.data())
def test_read_sst_corrupted(strings, data):
    part = bytearray(_sst_part(strings))
    for _ in range(data.draw(st.integers(0, 4))):
        if not part:
            break
        index = data.draw(st.integers(0, len(part) - 1))
        part[index] = data.draw(st.integers(0, 255))
    cut = data.draw(st.integers(0, len(part)))
    same("read_xlsb_shared_strings", bytes(part[:cut]))


# -- column widths and the formatted-cell scan ------------------------------

STYLE = Style(bold=True)

CELLS = [
    None, 0, -12, 3.5, float("nan"), float("inf"), True, Decimal("1.10"), Decimal("NaN"),
    Decimal("sNaN"), datetime.datetime(2024, 1, 2, 3, 4), Timestampish(2020, 5, 6),
    datetime.date(2024, 1, 2), datetime.time(1, 2), datetime.timedelta(days=2),
    *TEXTS[:8], "x" * 300, (5, "0.00"), (None, "0.00"), ("text", STYLE), (1, 2), (1, 2, 3),
    (Decimal("2.5"), "0.0"), BadStr(), (BadStr(), "@"), BadLenTuple((1, "@")),
    BadItemTuple((1, "@")), CellError("error:#N/A"), b"bytes", [1, 2],
]


@ported("calc_column_widths")
@pytest.mark.parametrize("max_cols", [0, 1, 3, 10])
def test_column_widths_cases(max_cols):
    rows = [CELLS[i:i + 4] for i in range(0, len(CELLS), 4)]
    rows += [tuple(CELLS[:3]), "not a row", None, []]
    same("calc_column_widths", rows, max_cols)


@ported("calc_column_widths")
def test_column_widths_bad_arguments():
    same("calc_column_widths", ((1,),), 1)
    same("calc_column_widths", [[1]], -1)
    same("calc_column_widths", [[1]], 2**31)
    same("calc_column_widths", [[1]], "2")


@ported("calc_column_widths")
@settings(max_examples=300, deadline=None)
@given(st.lists(st.lists(st.sampled_from(CELLS), max_size=6), max_size=6), st.integers(0, 8))
def test_column_widths_random(rows, max_cols):
    same("calc_column_widths", rows, max_cols)


@ported("xlsb_row_has_formatted")
@pytest.mark.parametrize("row", [
    [], [1, "a", None], [(1, "0.0")], ((1, STYLE),), [(1, 2)], [None, (1, 2, 3), ("a", "@")],
    [BadLenTuple((1, "@"))], [BadItemTuple((1, "@"))], [1, BadLenTuple((1, "@"))],
    iter([(1, "@")]), (x for x in [1, 2]), 5, None, "a string row",
])
def test_row_has_formatted_cases(row):
    if hasattr(row, "__next__"):
        items = list(row)
        same("xlsb_row_has_formatted", fresh=lambda: (iter(items),))
    else:
        same("xlsb_row_has_formatted", row)


@ported("xlsb_row_has_formatted")
@settings(max_examples=300, deadline=None)
@given(st.lists(st.sampled_from(CELLS), max_size=10))
def test_row_has_formatted_random(row):
    same("xlsb_row_has_formatted", row)
    same("xlsb_row_has_formatted", tuple(row))


# -- XLSB row encoders ------------------------------------------------------

class BadHash(str):
    def __hash__(self):
        raise RuntimeError("hash failed on purpose")


class BadFloatDecimal(Decimal):
    def __float__(self):
        raise RuntimeError("float failed on purpose")


class Dateish(datetime.date):
    pass


class FloatSub(float):
    pass


class IntSub(int):
    pass


ROW_CELLS = [
    None, "", "text", "zażółć", "\U0001F600 emoji", "\ud800 lone", "a" * 32767, "a" * 32768,
    "\U0001F600" * 16384, 0, 1, -1, 536870911, 536870912, -536870912, -536870913, 2**62, 2**63, -(2**63) - 1,
    10**400, True, False, 1.5, -0.0, float("nan"), float("inf"), FloatSub(2.5), IntSub(7), Decimal("1.25"),
    Decimal("NaN"), BadFloatDecimal("2"), datetime.datetime(2024, 2, 29, 23, 59, 59, 999999),
    datetime.datetime(1900, 1, 1), datetime.datetime(1900, 2, 28, 12), datetime.datetime(1900, 3, 1),
    datetime.datetime(1899, 12, 31), datetime.datetime(10, 1, 1), datetime.date(2024, 1, 2),
    datetime.date(1800, 1, 1), Dateish(2020, 1, 1), Timestampish(2021, 6, 7, 8), datetime.time(1, 2),
    (5, "0.00"), ("x", "@"), (None, "0.00"), ((1, 2), "0.00"), (datetime.date(2024, 1, 1), "yyyy"),
    (3, STYLE), (1, 2), BadItemTuple((1, "@")), BadStr(), CellError("error:#N/A"), b"bytes", object(),
]

FMT_XF = {"0.00": 11, "@": 12, "yyyy": 13, id(STYLE): 14, (5, "0.00"): 15, (5, id(STYLE)): 16,
          ("date", 5): 17, ("datetime", 5): 18, ("date", 7): -1}


def _table():
    strings = ["text", "shared", "zażółć"]
    return {s: i for i, s in enumerate(strings)}, list(strings)


def _xlsb_outcomes(row, row_idx=1, *extra, inline=False):
    results = []
    for module in (c_core, rs_core):
        if inline:
            results.append(outcome(module.encode_xlsb_row_inline, list(row), row_idx, *extra))
        else:
            ss_dict, ss_list = _table()
            res = outcome(module.encode_xlsb_row, list(row), ss_dict, ss_list, 3, 9, row_idx, *extra)
            results.append((res, ss_dict, ss_list))
    return results


def assert_same_xlsb(row, row_idx=1, *extra):
    for inline in (False, True):
        c_result, rs_result = _xlsb_outcomes(row, row_idx, *extra, inline=inline)
        assert rs_result == c_result, f"inline={inline} row={row!r} extra={extra!r}"


@ported("encode_xlsb_row")
@pytest.mark.parametrize("cell", ROW_CELLS, ids=lambda c: type(c).__name__)
@pytest.mark.parametrize("extra", [(), (None, FMT_XF), (5, FMT_XF), ([0, 5, None, 7], FMT_XF), (((1, 5),), FMT_XF),
                                   (5, None), (None, FMT_XF, b"\x00\x19" + b"\x01" * 25)])
def test_xlsb_row_single_cell(cell, extra):
    assert_same_xlsb(["first", cell, "text"], 1, *extra)
    assert_same_xlsb([cell], 0, *extra)


@ported("encode_xlsb_row")
@pytest.mark.parametrize("extra", [
    ("bad",), (-1,), ([1, "x", -3, 2**40],), (((0, 1), (2, 3), (2, 4)),), (((0, 1), "bad"),), (((0, "x"),),),
    (((0, -1),),), (None, ["not", "dict"]), (None, None, "not bytes"), (None, None, b"x" * 257), (None, None, b""),
    (True,), (2**32 - 1,),
])
def test_xlsb_row_style_arguments(extra):
    assert_same_xlsb([None, 1, "a", None, 2.5], 7, *extra)


@ported("encode_xlsb_row")
@pytest.mark.parametrize("style_ids", [2**32 + 3, [2**32], ((0, 2**32),)])
def test_xlsb_row_style_ids_past_32_bits(style_ids):
    # C's unsigned long is 32 bits on Windows (OverflowError) and 64 on Linux
    # (silently truncated); Rust raises on every platform.
    expected = ("returns", bytes, None) if isinstance(style_ids, list) else ("raises", OverflowError)
    result = outcome(rs_core.encode_xlsb_row_inline, [1, 2], 1, style_ids)
    assert result[:2] == expected[:2]


@ported("encode_xlsb_row")
def test_xlsb_row_bad_arguments():
    same("encode_xlsb_row", "not a list", {}, [], 0, 0, 1)
    same("encode_xlsb_row", [1], [], [], 0, 0, 1)
    same("encode_xlsb_row", [1], {}, (), 0, 0, 1)
    same("encode_xlsb_row", [1], {}, [], "0", 0, 1)
    same("encode_xlsb_row_inline", "not a list", 1)
    same("encode_xlsb_row_inline", [1], "1")


@ported("encode_xlsb_row")
def test_xlsb_row_failure_rolls_back_the_table():
    for row in (["new one", "new two", BadHash("x")], ["new one", float("nan")], ["new", object()]):
        c_result, rs_result = _xlsb_outcomes(row)
        assert rs_result == c_result
        assert rs_result[1:] == _table()


@ported("encode_xlsb_row")
@settings(max_examples=500, deadline=None)
@given(st.lists(st.sampled_from(ROW_CELLS), max_size=12), st.integers(0, 3),
       st.sampled_from([(), (None, FMT_XF), (5, FMT_XF), ([0, 5, 7, None, 5], FMT_XF), (((1, 5), (3, 7)), FMT_XF)]))
def test_xlsb_row_random(row, row_idx, extra):
    assert_same_xlsb(row, row_idx, *extra)


# -- XLSX row encoder -------------------------------------------------------

class NanFloat(float):
    pass


XLSX_CELLS = ROW_CELLS + [
    "a & b <c> d", "_x0041_ literal", "_x00G1_", "\ufffe", "\uffff end", " leading", "\tlead tab", "x\x01y",
    "tab\tnew\nline", "\r carriage", 1e16, 1e-7, 123456789.125, -5e300, 2.0**53, NanFloat("nan"), NanFloat("inf"),
    datetime.datetime(9999, 12, 31, 23, 59, 59), datetime.date(1900, 2, 28), datetime.date(1900, 3, 1),
]


def _xlsx_outcomes(row, row_idx, use_sst, *extra, make_row=list):
    results = []
    for module in (c_core, rs_core):
        ss_dict, ss_list = _table()
        res = outcome(module.encode_xlsx_row, make_row(row), ss_dict, ss_list, 3, 9, row_idx, use_sst, *extra)
        results.append((res, ss_dict, ss_list))
    return results


def assert_same_xlsx(row, row_idx=1, *extra, make_row=list):
    for use_sst in (True, False):
        c_result, rs_result = _xlsx_outcomes(row, row_idx, use_sst, *extra, make_row=make_row)
        assert rs_result == c_result, f"use_sst={use_sst} row={row!r} extra={extra!r}"


@ported("encode_xlsx_row")
@pytest.mark.parametrize("cell", XLSX_CELLS, ids=lambda c: type(c).__name__)
@pytest.mark.parametrize("extra", [(), (None, FMT_XF), (5, FMT_XF), ([0, 5, None, 7], FMT_XF), (((1, 5),), FMT_XF),
                                   (5, None), (None, FMT_XF, ' ht="20" customHeight="1"')])
def test_xlsx_row_single_cell(cell, extra):
    assert_same_xlsx(["first", cell, "text"], 1, *extra)
    assert_same_xlsx([cell], 0, *extra)


@ported("encode_xlsx_row")
@pytest.mark.parametrize("extra", [
    ("bad",), (-1,), ([1, "x", None],), ([2**40],), (((0, 1), (2, 3), (2, 4)),), (((0, 1), "bad"),),
    (None, ["not", "dict"]), (None, None, 5), (None, None, "\ud800"), (True,), (2**32 - 1,),
])
def test_xlsx_row_style_arguments(extra):
    assert_same_xlsx([None, 1, "a", None, 2.5], 7, *extra)


@ported("encode_xlsx_row")
@pytest.mark.parametrize("make_row", [list, tuple, iter, lambda r: (x for x in r)])
def test_xlsx_row_sequences(make_row):
    assert_same_xlsx(["a", 1, None, (5, "0.00")], 2, None, FMT_XF, make_row=make_row)


@ported("encode_xlsx_row")
def test_xlsx_row_bad_arguments():
    same("encode_xlsx_row", 5, {}, [], 0, 0, 1, True)
    same("encode_xlsx_row", [1], [], [], 0, 0, 1, True)
    same("encode_xlsx_row", [1], {}, (), 0, 0, 1, False)
    same("encode_xlsx_row", [1], {}, [], "0", 0, 1, True)
    same("encode_xlsx_row", [1], {}, [], 0, 0, 1, BadStr())
    same("encode_xlsx_row", [1], {}, [], 0, 0, 2**32 - 1, True)


@ported("encode_xlsx_row")
def test_xlsx_row_failure_rolls_back_the_table():
    for row in (["new one", "new two", BadHash("x")], ["new one", float("nan")], ["new", object()]):
        c_result, rs_result = _xlsx_outcomes(row, 1, True)
        assert rs_result == c_result
        assert rs_result[1:] == _table()


@ported("encode_xlsx_row")
@settings(max_examples=500, deadline=None)
@given(st.lists(st.sampled_from(XLSX_CELLS), max_size=12), st.integers(0, 3),
       st.sampled_from([(), (None, FMT_XF), (5, FMT_XF), ([0, 5, 7, None, 5], FMT_XF), (((1, 5), (3, 7)), FMT_XF)]))
def test_xlsx_row_random(row, row_idx, extra):
    assert_same_xlsx(row, row_idx, *extra)


@ported("encode_xlsx_row")
@settings(max_examples=300, deadline=None)
@given(st.lists(st.text(max_size=12), max_size=6))
def test_xlsx_row_random_text(row):
    assert_same_xlsx(row, 1)


# -- XLSB worksheet reader --------------------------------------------------

def _drain(module, *args):
    """Every row the reader yields, then how it ended (StopIteration or an
    exception class)."""
    try:
        rows = module.read_xlsb_worksheet(*args)
    except Exception as exc:  # noqa: BLE001
        return ("raises", type(exc))
    if rows is None:
        return ("none",)
    seen = []
    try:
        for row in rows:
            seen.append(row)
            if len(seen) > 5000:
                break
    except Exception as exc:  # noqa: BLE001
        return ("rows", seen, type(exc))
    return ("rows", seen, None)


def same_rows(*args):
    expected, actual = _drain(c_core, *args), _drain(rs_core, *args)
    assert actual == expected


def _rec(rec_id, body=b""):
    out = bytearray()
    for value in (rec_id, len(body)):
        while True:
            byte = value & 0x7F
            value >>= 7
            out.append(byte | (0x80 if value else 0))
            if not value:
                break
    return bytes(out) + body


def _row(n):
    return _rec(0x00, n.to_bytes(4, "little") + b"\0" * 21)


def _cell(rec_id, col, payload, xf=0):
    return _rec(rec_id, col.to_bytes(4, "little") + xf.to_bytes(4, "little") + payload)


def _rk_int(n):
    return ((n << 2) | 2).to_bytes(4, "little", signed=n < 0)


SYNTH_SHARED = ["zero", "one", "two"]
SYNTH_STYLES = [{"numFmtId": 0}, {"numFmtId": 14}, {"numFmtId": 22}, {"other": 1}, "not a dict"]
SYNTH_DATES = {14, 22}

def _synth_sheets():
    import struct
    text = "zażółć \U0001F600".encode("utf-16-le")
    cells = [
        _cell(0x01, 0, b""), _cell(0x02, 1, _rk_int(42)), _cell(0x02, 2, ((314 << 2) | 3).to_bytes(4, "little")),
        _cell(0x03, 3, b"\x07"), _cell(0x0B, 4, b"\x2a"), _cell(0x03, 5, b""),
        _cell(0x04, 6, b"\x01"), _cell(0x04, 7, b"\x02"), _cell(0x05, 8, struct.pack("<d", 2.5)),
        _cell(0x06, 9, (len(text) // 2).to_bytes(4, "little") + text), _cell(0x07, 10, (1).to_bytes(4, "little")),
        _cell(0x07, 11, (99).to_bytes(4, "little")), _cell(0x05, 12, struct.pack("<d", 45000.25), xf=1),
        _cell(0x02, 13, _rk_int(60), xf=2), _cell(0x05, 16, struct.pack("<d", 1e300)),
        _cell(0x05, 17, struct.pack("<d", -0.0)), _cell(0x05, 18, struct.pack("<d", 5e-324)),
    ]
    body = b"".join(cells)
    sheets = [
        _row(0) + body,
        _row(0) + body[:-3],                                 # truncated last record
        body + _row(2) + _cell(0x02, 0, _rk_int(1)),          # cells before any Row record
        _row(2) + _row(5) + _cell(0x02, 3, _rk_int(1)) + _cell(0x02, 4, _rk_int(2)) + _row(9) + _cell(0x02, 7, _rk_int(3)),
        _row(1) + _cell(0x02, 4, _rk_int(1)) + _cell(0x02, 2, _rk_int(2)) + _row(2) + _cell(0x02, 3, _rk_int(3)),
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _row(1) + _cell(0x02, 30, _rk_int(2)) + _cell(0x02, 20000, _rk_int(3)),
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _cell(0x02, 2_000_000, _rk_int(2)),
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _row(2_000_000),
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _rec(0x00, b"\xff\xff\xff\xff") + _cell(0x02, 0, _rk_int(2)),
        _row(0) + _cell(0x06, 0, b"\xff\xff\xff\x7f") + _cell(0x02, 1, b"\x01"),
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _rec(0x92) + _row(1) + _cell(0x02, 0, _rk_int(2)),
        _row(4),
        # Dates the reference cannot convert end the iteration with its error.
        _row(0) + _cell(0x02, 0, _rk_int(1)) + _row(1) + _cell(0x05, 0, struct.pack("<d", 1e12), xf=1),
        _row(0) + _cell(0x05, 0, struct.pack("<d", float("nan")), xf=1),
        _row(0) + _cell(0x05, 0, struct.pack("<d", -700000.0), xf=1),
        _row(0) + _cell(0x05, 0, struct.pack("<d", 2958466.0), xf=2),
    ]
    return sheets


@pytest.fixture(scope="module")
def written_sheets(tmp_path_factory):
    """Sheets written by XlsbWriter, with the styles and shared strings an
    ExcelReader loads for them."""
    import zipfile
    from xlspy import ExcelReader, XlsbWriter
    from xlspy.reader import _get_date_num_fmt_ids
    path = tmp_path_factory.mktemp("xlsb") / "parity.xlsb"
    rows = [["text", 1, 2.5, True, None, datetime.datetime(2024, 5, 6, 7, 8, 9), datetime.date(2020, 1, 1),
             (3.25, "0.00"), Decimal("1.5"), -12345678901, CellError("error:#DIV/0!")]]
    rows += [[f"r{i}", i, i / 3, None, i % 2 == 0] for i in range(200)]
    rows += [[], [None, None, "late"], ["x" * 300, "\U0001F600"]]
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(rows)
    with zipfile.ZipFile(path) as zf:
        blob = zf.read("xl/worksheets/sheet1.bin")
    with ExcelReader(str(path)) as reader:
        return [(blob, list(reader._shared_string_array), list(reader._styles_cell_xfs_array),
                 set(_get_date_num_fmt_ids(reader._number_formats_type_dict)))]


@ported("read_xlsb_worksheet")
@pytest.mark.parametrize("index", range(len(_synth_sheets())))
@pytest.mark.parametrize("date_1904", [False, True])
def test_read_xlsb_synthetic(index, date_1904):
    same_rows(_synth_sheets()[index], SYNTH_SHARED, SYNTH_STYLES, SYNTH_DATES, date_1904)


@ported("read_xlsb_worksheet")
def test_read_xlsb_written(written_sheets):
    for blob, shared, styles, dates in written_sheets:
        same_rows(blob, shared, styles, dates, False)
        same_rows(bytearray(blob), shared, styles, dates, True)


@ported("read_xlsb_worksheet")
def test_read_xlsb_bad_arguments():
    same_rows("text", [], [], set())
    same_rows(b"", (), [], set())
    same_rows(b"", [], (), set())
    same_rows(b"", [], [], frozenset())
    same_rows(b"", [], [], set(), BadStr())


@ported("read_xlsb_worksheet")
@settings(max_examples=400, deadline=None)
@given(st.sampled_from(range(len(_synth_sheets()))), st.data())
def test_read_xlsb_corrupted(index, data):
    sheet = bytearray(_synth_sheets()[index] or b"\0")
    for _ in range(data.draw(st.integers(0, 6))):
        sheet[data.draw(st.integers(0, len(sheet) - 1))] = data.draw(st.integers(0, 255))
    cut = data.draw(st.integers(0, len(sheet)))
    same_rows(bytes(sheet[:cut]), SYNTH_SHARED, SYNTH_STYLES, SYNTH_DATES, data.draw(st.booleans()))
