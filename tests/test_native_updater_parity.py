"""The updaters' native row encoders write the bytes their Python code writes.

``XlsxUpdater`` and ``XlsbUpdater.replace_sheet_data`` (and ``_stream``) encode
the common cells in Rust (``update_rows.rs``) and hand any other cell to the
Python code, one cell at a time. Every variant here is written both ways and
the saved packages must be equal, part for part; errors must be the same.
"""

import datetime as dt
import random
import zipfile
from decimal import Decimal

import pytest

from xlspy import Font, Formula, Style, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter
from xlspy import _accel, xlsb_updater, xlsx_updater

pytestmark = pytest.mark.skipif(_accel.encode_update_xlsx_row is None, reason="native extension not built")

UPDATERS = [(XlsxUpdater, XlsxWriter, "xlsx"), (XlsbUpdater, XlsbWriter, "xlsb")]


class _Opaque:
    def __str__(self):
        return "opaque"


class _Text(str):
    pass


def _values(xlsb: bool):
    rng = random.Random(11)
    values = [
        0, 1, -1, 536870911, 536870912, -536870913, 2**53 + 1, 2**63 - 1, -(2**63), 2**64, 10**30,
        0.0, -0.0, 0.1 + 0.2, 1e-300, 5e-324, 1.7976931348623157e308, 123456789.123456789, 1e15, 1e16,
        True, False, None, "", "text", " lead", "trail ", "tab\there", "line\nbreak", "bell\x07",
        "zażółć", "😀 astral", "<&>\"'", _Text("sub"), Decimal("1.25"), _Opaque(), (5, "0.00"),
        (None, "0.00"), ("wrapped", "@"), (1, 2), [3],
        dt.date(2024, 2, 29), dt.datetime(2024, 5, 6, 7, 8, 9, 123456), dt.date(1900, 1, 1),
        dt.date(1900, 2, 28), dt.date(1900, 3, 1), dt.datetime(1899, 12, 31, 12), dt.date(1800, 1, 1),
        dt.date(9999, 12, 31), dt.datetime(2024, 1, 1, tzinfo=dt.timezone(dt.timedelta(hours=2))),
        dt.time(3, 4), dt.timedelta(days=1),
    ]
    values += [rng.uniform(-1e12, 1e12) for _ in range(40)] + [rng.random() * 10 ** rng.randint(-20, 20) for _ in range(40)]
    values += [rng.randint(-(2**62), 2**62) for _ in range(20)]
    if not xlsb:
        values += [Formula("1+1"), Formula("A1*2", 3.5), Formula('"a"&"b"', "ab"), (Formula("2*2"), "0.00")]
    try:
        import numpy as np
    except ImportError:
        pass
    else:
        values += [np.int64(7), np.float32(1.5), np.bool_(True), np.float64(2.5)]
    return values


def _rows(xlsb: bool, width: int = 9):
    values = _values(xlsb)
    return [[values[(r * width + c) % len(values)] for c in range(width)] for r in range(len(values))]


def _source(path, writer, shared):
    styles = {"B:B": Style(font=Font(bold=True)), "C:C": Style(number_format="0.000"),
              "E:E": Style(number_format="yyyy-mm-dd")}
    with writer(path, use_shared_strings=shared) as book:
        book.add_sheet("S")
        book.write_sheet(
            [["h1", "h2", "h3", "h4", "h5"], [1, "x", 2.5, dt.date(2024, 1, 1), dt.date(2024, 1, 2)]],
            styles=styles, header_style=Style(font=Font(italic=True)),
        )


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in sorted(archive.namelist())}


def _replace(tmp_path, monkeypatch, updater, writer, suffix, *, native, shared, stream, fallback,
             headers, date_1904, rows):
    source = tmp_path / f"source-{shared}.{suffix}"
    if not source.exists():
        _source(source, writer, shared)
    target = tmp_path / f"{native}-{shared}-{stream}-{fallback}-{headers}-{date_1904}.{suffix}"
    with monkeypatch.context() as patch:
        if not native:
            patch.setattr(xlsx_updater, "encode_update_xlsx_row", None)
            patch.setattr(xlsb_updater, "encode_update_xlsb_row", None)
            patch.setattr(xlsb_updater, "patch_xlsb_row_headers", None)
        book = updater(source)
        try:
            book._date_1904 = date_1904
            replace = book.replace_sheet_data_stream if stream else book.replace_sheet_data
            replace("S", iter(rows) if stream else rows,
                    headers=["a", "b", None, 4, "e"] if headers else None, style_fallback=fallback)
            book.save(target)
        finally:
            book.close()
    return _parts(target)


@pytest.mark.parametrize("updater, writer, suffix", UPDATERS)
@pytest.mark.parametrize("shared", [True, False])
@pytest.mark.parametrize("stream", [False, True])
@pytest.mark.parametrize("fallback", ["inherit", "general"])
def test_native_rows_match_python_rows(tmp_path, monkeypatch, updater, writer, suffix, shared, stream, fallback):
    rows = _rows(suffix == "xlsb")
    for headers, date_1904 in ((True, False), (False, True)):
        options = {"shared": shared, "stream": stream, "fallback": fallback, "headers": headers, "date_1904": date_1904, "rows": rows}
        native = _replace(tmp_path, monkeypatch, updater, writer, suffix, native=True, **options)
        python = _replace(tmp_path, monkeypatch, updater, writer, suffix, native=False, **options)
        assert native.keys() == python.keys()
        for name in native:
            assert native[name] == python[name], name


@pytest.mark.parametrize("updater, writer, suffix", UPDATERS)
@pytest.mark.parametrize("bad", [float("nan"), float("inf"), "x" * 40000], ids=["nan", "inf", "long-text"])
def test_errors_are_the_python_errors(tmp_path, monkeypatch, updater, writer, suffix, bad):
    errors = []
    for native in (True, False):
        with pytest.raises(Exception) as caught:
            _replace(tmp_path, monkeypatch, updater, writer, suffix, native=native, shared=True, stream=False,
                     fallback="inherit", headers=False, date_1904=False, rows=[[1, "a"], ["b", bad]])
        errors.append((type(caught.value), str(caught.value)))
    assert errors[0] == errors[1]


def test_xlsb_formula_rows_match_the_python_encoder(tmp_path, monkeypatch):
    rows = [[1, "a"], ["b", Formula("1+1", 2)], ["c", Formula("B2&\"x\"")]]
    options = {"shared": True, "stream": False, "fallback": "inherit", "headers": False,
               "date_1904": False, "rows": rows}
    native = _replace(tmp_path, monkeypatch, XlsbUpdater, XlsbWriter, "xlsb", native=True, **options)
    python = _replace(tmp_path, monkeypatch, XlsbUpdater, XlsbWriter, "xlsb", native=False, **options)
    assert native == python
