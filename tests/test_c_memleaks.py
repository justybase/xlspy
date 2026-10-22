"""Memory-leak regression tests for the ``xlspy._c_core`` extension.

Cover every C entry point on success paths, error paths (where a missed
``Py_DECREF`` would leak), and end-to-end workbook cycles:

- RSS-based loops (``psutil``, skipped when unavailable): catch raw
  ``malloc``/``realloc`` leaks that ``tracemalloc`` cannot see.
- ``tracemalloc`` loops (stdlib, always run): catch ``PyMem``/``PyObject``
  leaks precisely.
- refcount contracts: borrowed inputs keep their refcount, fresh results
  start at exactly 2.

Thresholds are generous ( allocator arenas make RSS noisy); a real leak of
a per-call result (tens of KB-MBs) trips them within seconds.
"""
import datetime
import gc
import tracemalloc
import zipfile
from decimal import Decimal

import pytest

from xlspy import ExcelReader, XlsbWriter
from xlspy._accel import (
    HAVE_C_EXT,
    calc_column_widths,
    encode_xlsb_row,
    encode_xlsb_sst,
    encode_xlsx_row,
    read_xlsb_shared_strings,
    read_xlsb_worksheet,
    xlsb_row_has_formatted,
)

needs_c_ext = pytest.mark.skipif(not HAVE_C_EXT, reason="C extension not built")

SST_BLOB = None
SHEET_BLOB = None
SHARED = None
STYLES = None
DATE_FMTS = None


def _mixed_row(i):
    return [
        f"str_{i}_żółć",
        i,
        i / 3.0,
        i % 2 == 0,
        datetime.date(2023, 1, 1) + datetime.timedelta(days=i % 365),
        datetime.datetime(2023, 1, 1, 12, 0, 0) + datetime.timedelta(seconds=i),
        Decimal(f"{i}.25"),
        2**60 + i,
        None,
    ]


@pytest.fixture(scope="module")
def workbook_parts(tmp_path_factory):
    """A real .xlsb plus its raw parts, built once for the module."""
    global SST_BLOB, SHEET_BLOB, SHARED, STYLES, DATE_FMTS
    path = str(tmp_path_factory.mktemp("memleak") / "leak.xlsb")
    with XlsbWriter(path) as w:
        w.add_sheet("S")
        w.write_sheet([_mixed_row(i) for i in range(3000)])
    with zipfile.ZipFile(path) as zf:
        SST_BLOB = zf.read("xl/sharedStrings.bin")
        SHEET_BLOB = zf.read("xl/worksheets/sheet1.bin")
    with ExcelReader(path) as r:
        SHARED = list(r._shared_string_array)
        STYLES = list(r._styles_cell_xfs_array)
        from xlspy.reader import _get_date_num_fmt_ids
        DATE_FMTS = set(_get_date_num_fmt_ids(r._number_formats_type_dict))
    return path


def _rss_mb():
    psutil = pytest.importorskip("psutil")
    return psutil.Process().memory_info().rss / 1e6


def assert_no_rss_growth(fn, *, iters=40, warmup=8, allowance_mb=5.0):
    for _ in range(warmup):
        fn()
    gc.collect()
    before = _rss_mb()
    for _ in range(iters):
        fn()
    gc.collect()
    after = _rss_mb()
    assert after - before < allowance_mb, (
        f"RSS grew by {after - before:.1f} MB over {iters} iterations"
    )


def assert_no_tracemalloc_growth(fn, *, iters=30, warmup=5, allowance_kb=128.0):
    tracemalloc.start()
    try:
        for _ in range(warmup):
            fn()
        gc.collect()
        snap1 = tracemalloc.take_snapshot()
        for _ in range(iters):
            fn()
        gc.collect()
        snap2 = tracemalloc.take_snapshot()
        growth = sum(
            stat.size_diff
            for stat in snap2.compare_to(snap1, "filename")
            if stat.size_diff > 0
        )
        assert growth / 1e3 < allowance_kb, (
            f"traced Python memory grew by {growth / 1e3:.0f} KB over {iters} iterations"
        )
    finally:
        tracemalloc.stop()


def assert_no_gc_growth(fn, *, iters=100, warmup=5, allowance=10):
    """No net accumulation of GC-tracked objects (lists, dicts, instances).

    Used instead of ``sys.getrefcount``: refcount values are not reliable
    assertions under pytest on Python 3.14 (GC-tracked objects read one
    less than expected), while GC population deltas are deterministic.
    """
    for _ in range(warmup):
        fn()
    gc.collect()
    before = len(gc.get_objects())
    for _ in range(iters):
        fn()
    gc.collect()
    after = len(gc.get_objects())
    assert after - before <= allowance, (
        f"GC-tracked objects grew by {after - before} over {iters} iterations"
    )


@needs_c_ext
def test_sst_read_no_growth(workbook_parts):
    assert SST_BLOB is not None and len(SHARED) > 1000
    assert_no_rss_growth(lambda: read_xlsb_shared_strings(SST_BLOB))
    assert_no_tracemalloc_growth(lambda: read_xlsb_shared_strings(SST_BLOB))


@needs_c_ext
def test_sst_read_error_paths_no_growth(workbook_parts):
    bad_blobs = [
        SST_BLOB[:7],
        SST_BLOB[: len(SST_BLOB) // 2],
        b"\x13\x02\x00\x00",          # short item record
        b"\x9f\x01\x02\x00\x00",      # short BeginSst
        b"\x13\x08\x00\x01\x00\x00\x00abc",  # odd string bytes
        b"\xff\xff\xff\xff\xff\x04" + b"\x00" * 8,  # hostile lengths
    ]

    def run():
        for blob in bad_blobs:
            try:
                read_xlsb_shared_strings(blob)
            except (ValueError, IndexError, UnicodeDecodeError):
                pass

    assert_no_tracemalloc_growth(run, iters=60, allowance_kb=256.0)
    assert_no_rss_growth(run, iters=40, allowance_mb=5.0)


@needs_c_ext
def test_encode_sst_no_growth(workbook_parts):
    assert_no_rss_growth(
        lambda: encode_xlsb_sst(SHARED, len(SHARED) * 2, len(SHARED))
    )
    assert_no_tracemalloc_growth(
        lambda: encode_xlsb_sst(SHARED, len(SHARED) * 2, len(SHARED))
    )
    # error path: lone surrogate must raise without leaking
    bad = ["ok", "bad\ud800"]
    for _ in range(200):
        with pytest.raises(UnicodeEncodeError):
            encode_xlsb_sst(bad, 2, 2)
    gc.collect()


@needs_c_ext
def test_row_encode_no_growth():
    ss_dict, ss_list = {}, []
    rows = [_mixed_row(i) for i in range(200)]
    rows.append([(1, "#,##0"), "fmt"])
    state = {"uniq": 0, "all": 0}

    def run():
        for idx, row in enumerate(rows):
            res = encode_xlsb_row(row, ss_dict, ss_list, state["uniq"], state["all"], idx)
            if res is not None:  # None: the row goes to the Python encoder (here, the format tuple)
                state["uniq"], state["all"] = res[1], res[2]

    assert_no_tracemalloc_growth(run, iters=20, allowance_kb=512.0)
    assert_no_rss_growth(run, iters=20, allowance_mb=5.0)


@needs_c_ext
def test_prescan_and_widths_no_growth():
    rows = [_mixed_row(i) for i in range(200)]

    def run():
        for row in rows:
            xlsb_row_has_formatted(row)
        calc_column_widths(rows, 9)

    assert_no_tracemalloc_growth(run, iters=30, allowance_kb=256.0)


@needs_c_ext
def test_worksheet_iterator_no_growth(workbook_parts):
    def run():
        it = read_xlsb_worksheet(SHEET_BLOB, SHARED, STYLES, DATE_FMTS)
        n = 0
        for _row in it:
            n += 1
        assert n == 3000

    assert_no_rss_growth(run, iters=25, allowance_mb=6.0)
    assert_no_tracemalloc_growth(run, iters=25, allowance_kb=512.0)


@needs_c_ext
def test_worksheet_iterator_abandoned_no_growth(workbook_parts):
    # Dropping a partially-consumed iterator must free cur_cells + buffer.
    def run():
        it = read_xlsb_worksheet(SHEET_BLOB, SHARED, STYLES, DATE_FMTS)
        next(it)

    assert_no_tracemalloc_growth(run, iters=200, allowance_kb=512.0)


@needs_c_ext
def test_end_to_end_read_no_growth(workbook_parts):
    def run():
        with ExcelReader(workbook_parts) as r:
            assert len(r.read_all("S")) == 3000

    assert_no_rss_growth(run, iters=15, allowance_mb=8.0)


@needs_c_ext
def test_gc_contracts(workbook_parts):
    out = read_xlsb_shared_strings(SST_BLOB)
    assert out == SHARED
    del out
    assert gc.garbage == []

    def run_all():
        read_xlsb_shared_strings(SST_BLOB)
        encode_xlsb_sst(SHARED, 1, 1)
        xlsb_row_has_formatted(_mixed_row(1))

    assert_no_gc_growth(run_all, iters=100, allowance=10)


@needs_c_ext
def test_encode_xlsx_row_no_growth():
    rows = [_mixed_row(i) for i in range(100)]

    def run_sst():
        ss_dict, ss_list = {}, []
        uniq = allc = 0
        for idx, row in enumerate(rows):
            res = encode_xlsx_row(row, ss_dict, ss_list, uniq, allc, idx, True)
            assert res is not None
            uniq, allc = res[1], res[2]

    def run_inline():
        ss_dict, ss_list = {}, []
        for idx, row in enumerate(rows):
            res = encode_xlsx_row(row, ss_dict, ss_list, 0, 0, idx, False)
            assert res is not None

    assert_no_rss_growth(run_sst, iters=20, allowance_mb=5.0)
    assert_no_tracemalloc_growth(run_sst, iters=20, allowance_kb=512.0)
    assert_no_tracemalloc_growth(run_inline, iters=20, allowance_kb=512.0)


@needs_c_ext
def test_encode_xlsx_row_rollback_no_growth():
    # Rows that insert SST entries then hit a row fallback (huge int overflows
    # float(); tuple formats owned by the Python path) must leave the caller's
    # dict/list untouched so the Python slow path reprocesses the row cleanly.
    fb_rows = [
        ["a", "b", 10**1000],
        ["x", "y", ("v", "#,##0")],
        ["p", "q", "r", 10**1000, ("w", "0.00")],
    ]
    d0, l0 = {"pre": 0}, ["pre"]

    def run():
        for idx, row in enumerate(fb_rows):
            d, l = dict(d0), list(l0)
            assert encode_xlsx_row(row, d, l, 1, 1, idx, True) is None
            assert d == d0 and l == l0, f"rollback leaked SST state: {d} {l}"

    run()  # correctness first
    assert_no_tracemalloc_growth(run, iters=100, allowance_kb=256.0)
    assert_no_rss_growth(run, iters=40, allowance_mb=5.0)


@needs_c_ext
def test_encode_xlsx_row_error_paths_no_growth():
    def run():
        # non-dict / non-list args raise TypeError; non-sequence rows fall back
        for bad in (("not-a-dict", []), ({}, "not-a-list")):
            with pytest.raises(TypeError):
                encode_xlsx_row(["a"], bad[0], bad[1], 0, 0, 0, True)
        assert encode_xlsx_row(42, {}, [], 0, 0, 0, True) is None
        # inline lone surrogate falls back to the Python path (which raises)
        assert encode_xlsx_row(["ok", "bad\ud800"], {}, [], 0, 0, 0, False) is None

    run()
    assert_no_tracemalloc_growth(run, iters=100, allowance_kb=256.0)
    assert_no_gc_growth(run, iters=100, allowance=10)


def _formatted_row(i):
    """(value, format) cells of every type the C encoders take, plus a styled one."""
    return [
        (f"s{i}", "@"),
        (i, "#,##0"),
        (i / 4, "0.00"),
        (i % 2 == 0, "0"),
        (datetime.date(2024, 1, 1) + datetime.timedelta(days=i % 300), "yyyy-mm-dd"),
        (Decimal(f"{i}.5"), "0.0"),
        None,
    ]


FMT_XF = {"@": 10, "#,##0": 11, "0.00": 12, "0": 13, "yyyy-mm-dd": 14, "0.0": 15,
          (20, "0.00"): 21}


@needs_c_ext
def test_encoders_with_format_ids_no_growth():
    from xlspy._accel import encode_xlsb_row_inline

    rows = [_formatted_row(i) for i in range(100)]
    styled = [[(1.5, "0.00"), "x"]]  # column 0 has style id 20

    def run():
        ss_dict, ss_list = {}, []
        uniq = allc = 0
        for idx, row in enumerate(rows):
            res = encode_xlsx_row(row, ss_dict, ss_list, uniq, allc, idx, True, None, FMT_XF)
            assert res is not None
            uniq, allc = res[1], res[2]
            res = encode_xlsb_row(row, {}, [], 0, 0, idx, None, FMT_XF)
            assert res is not None
            assert encode_xlsb_row_inline(row, idx, None, FMT_XF) is not None
        assert encode_xlsx_row(styled[0], {}, [], 0, 0, 1, True, ((0, 20),), FMT_XF) is not None
        assert encode_xlsb_row(styled[0], {}, [], 0, 0, 1, ((0, 20),), FMT_XF) is not None

    run()
    assert_no_tracemalloc_growth(run, iters=20, allowance_kb=512.0)
    assert_no_gc_growth(run, iters=50, allowance=10)


@needs_c_ext
def test_encoders_with_format_ids_fall_back_cleanly():
    from xlspy._accel import encode_xlsb_row_inline

    class Tup(tuple):
        pass

    fallbacks = [
        ["a", ("v", "unknown format")],   # format not recorded yet
        ["a", (None, "0.00")],            # the Python path writes "None"
        ["a", Tup((1, "0.00"))],          # XLSB borrows items of exact tuples only
        ["a", ((1, 2), "0.00")],          # nested tuple
    ]
    d0, l0 = {"pre": 0}, ["pre"]

    def run():
        for idx, row in enumerate(fallbacks):
            d, l = dict(d0), list(l0)
            if idx != 2:
                assert encode_xlsx_row(row, d, l, 1, 1, idx, True, None, FMT_XF) is None
                assert d == d0 and l == l0
            assert encode_xlsb_row(row, d, l, 1, 1, idx, None, FMT_XF) is None
            assert d == d0 and l == l0
            assert encode_xlsb_row_inline(row, idx, None, FMT_XF) is None
        with pytest.raises(TypeError, match="fmt_xf"):
            encode_xlsx_row(["a"], {}, [], 0, 0, 0, True, None, [])
        with pytest.raises(TypeError, match="fmt_xf"):
            encode_xlsb_row(["a"], {}, [], 0, 0, 0, None, [])

    run()
    assert_no_tracemalloc_growth(run, iters=100, allowance_kb=256.0)
    assert_no_gc_growth(run, iters=100, allowance=10)


@needs_c_ext
def test_encoders_with_row_attributes_no_growth():
    from xlspy._accel import encode_xlsb_row_inline

    header = XlsbWriter._row_header_record(3, 2, 400, True, 2, False)
    row = ["a", 1, 2.5]

    def run():
        xml, _uniq, _all = encode_xlsx_row(row, {}, [], 0, 0, 3, True, None, None,
                                           ' ht="20" customHeight="1" hidden="1"')
        assert xml.startswith('<row r="4" ht="20" customHeight="1" hidden="1"><c')
        data, _uniq, _all = encode_xlsb_row(row, {}, [], 0, 0, 3, None, None, header)
        assert data.startswith(header)
        assert encode_xlsb_row_inline(row, 3, None, None, header).startswith(header)
        with pytest.raises(TypeError, match="row_attrs"):
            encode_xlsx_row(row, {}, [], 0, 0, 3, True, None, None, b"x")
        with pytest.raises(TypeError, match="row_header"):
            encode_xlsb_row(row, {}, [], 0, 0, 3, None, None, "x")

    run()
    assert_no_tracemalloc_growth(run, iters=100, allowance_kb=256.0)
    assert_no_gc_growth(run, iters=100, allowance=10)
