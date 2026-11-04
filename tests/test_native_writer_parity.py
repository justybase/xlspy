"""The native row encoders write the same worksheet bytes as the Python encoders."""

import datetime as dt
import random
import zipfile

import pytest

from xlspy import (
    CellRichText, ExcelReader, Font, Formula, Hyperlink, InlineFont, Style, TextBlock, XlsbWriter, XlsxWriter,
    _accel, _xlsb_writer_bin, _xlsx_writer_sheet,
)

pytestmark = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")


def _rows():
    rng = random.Random(5)
    rows = [[0, -0.0, 0.0, 1e16, -1e16, 1e16 - 2, 2**53 + 1, -2**63, 2**64, 10**20, 0.5,
             5e-324, 1.7976931348623157e308, 123456789.0, True]]
    for _ in range(500):
        rows.append([
            rng.randrange(-10**12, 10**12),
            float(rng.randrange(-10**17, 10**17)),
            rng.uniform(-1e6, 1e6),
            dt.date(1900, 1, 1) + dt.timedelta(days=rng.randrange(0, 2_900_000)),
            # Microseconds exercise the exact serial; years past ~2185 the row fallback.
            dt.datetime(1900, 1, 1) + dt.timedelta(
                seconds=rng.randrange(0, 2**37), microseconds=rng.randrange(10**6)),
            dt.date(1900, rng.choice([1, 2, 3]), rng.randrange(1, 29)),
            dt.datetime(rng.choice([1899, 1900, 2024, 9999]), 2, 28, 23, 59, 59, 999999),
        ])
    return rows


def _sheet(path):
    with zipfile.ZipFile(path) as archive:
        name = next(n for n in archive.namelist() if n.startswith("xl/worksheets/sheet"))
        return archive.read(name)


@pytest.mark.parametrize("cls, suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
@pytest.mark.parametrize("shared", [True, False])
def test_native_rows_match_python_rows(tmp_path, monkeypatch, cls, suffix, shared):
    rows = _rows()
    outputs = []
    for native in (True, False):
        with monkeypatch.context() as patch:
            if not native:
                none = lambda *args, **kwargs: None
                patch.setattr(_xlsx_writer_sheet, "accel_encode_xlsx_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row_inline", none)
            path = tmp_path / f"{native}.{suffix}"
            with cls(path, use_shared_strings=shared) as book:
                book.add_sheet("S")
                book.write_sheet(rows)
            outputs.append(_sheet(path))
    assert outputs[0] == outputs[1]


def _write_both(tmp_path, monkeypatch, cls, rows, shared=True, **options):
    """The sheet bytes written natively and by the Python encoders, and how many
    rows the native writer handed to the Python encoder."""
    suffix = "xlsx" if cls is XlsxWriter else "xlsb"
    outputs, python_rows = [], []
    for native in (True, False):
        with monkeypatch.context() as patch:
            if not native:
                none = lambda *args, **kwargs: None
                patch.setattr(_xlsx_writer_sheet, "accel_encode_xlsx_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row_inline", none)
            calls = []
            for owner, name in ((_xlsb_writer_bin._XlsbBinMixin, "_write_python_row"),
                                (_xlsx_writer_sheet._XlsxSheetMixin, "_write_row_cells")):
                original = getattr(owner, name)

                def counted(self, *args, _original=original, _calls=calls, **kwargs):
                    _calls.append(1)
                    return _original(self, *args, **kwargs)

                patch.setattr(owner, name, counted)
            path = tmp_path / f"{native}.{suffix}"
            with cls(path, use_shared_strings=shared) as book:
                book.add_sheet("S")
                book.write_sheet(rows() if callable(rows) else rows, **options)
            outputs.append(_sheet(path))
            python_rows.append(len(calls))
    return outputs[0], outputs[1], python_rows[0]


WRITERS = pytest.mark.parametrize("cls", [XlsxWriter, XlsbWriter])
SHARED = pytest.mark.parametrize("shared", [True, False])


@WRITERS
@SHARED
def test_styled_date_columns_stay_on_the_native_encoder(tmp_path, monkeypatch, cls, shared):
    # A column style on dates needs the style's date variant, which the first
    # row records; from then on the native encoder writes the rows. XLSB once
    # sent every such row to the Python encoder (ten times slower).
    styles = {"A:A": Style(font=Font(name="Arial", italic=True)), "B:C": Style(font=Font(bold=True))}
    rows = [["date", "datetime", "mixed"]] + [
        [dt.date(2024, 1, 1) + dt.timedelta(days=i), dt.datetime(2024, 5, 6, 7, 8, i % 60),
         (dt.date(2024, 2, 3), "yyyy-mm-dd") if i % 3 == 0 else i]
        for i in range(300)
    ]
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows, shared, styles=styles)
    assert native == python
    assert python_rows <= 2


@WRITERS
@SHARED
def test_formatted_cells_in_a_styled_column_stay_native(tmp_path, monkeypatch, cls, shared):
    # Every cell of the column is a (value, format) overlay on the column style:
    # XLSB never registered the column style itself, so no row could go native.
    rows = [[i, (i * 1.5, "0.00"), (dt.date(2024, 1, 1), "yyyy-mm-dd")] for i in range(200)]
    styles = {"B:C": Style(font=Font(bold=True))}
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows, shared, styles=styles)
    assert native == python
    assert python_rows <= 1


@WRITERS
def test_rows_after_an_array_formula_stay_native(tmp_path, monkeypatch, cls):
    rows = [[i, "text", Formula("SUM(A1:A2*2)", array_ref="C6:C8") if i == 5 else None] for i in range(200)]
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows)
    assert native == python
    # The anchor row and its two member rows; XLSB wrote every later row in Python.
    assert python_rows <= 3


@WRITERS
@SHARED
def test_tuple_and_numpy_rows_write_as_lists_do(tmp_path, monkeypatch, cls, shared):
    np = pytest.importorskip("numpy")
    lists = [[i, i * 0.5, "t", dt.date(2024, 1, 1)] for i in range(50)]
    expected, _, _ = _write_both(tmp_path, monkeypatch, cls, lists, shared)
    native, python, _ = _write_both(tmp_path, monkeypatch, cls, [tuple(r) for r in lists], shared)
    assert native == python == expected
    numbers = np.arange(60, dtype=np.float64).reshape(20, 3)
    native, python, _ = _write_both(tmp_path, monkeypatch, cls, numbers, shared)
    assert native == python
    listed, _, _ = _write_both(tmp_path, monkeypatch, cls, numbers.tolist(), shared)
    assert native == listed


@WRITERS
@SHARED
def test_ints_past_the_float_range_are_written_as_text(tmp_path, monkeypatch, cls, shared):
    rows = [["big", 10**400], ["negative", -(10**400)], ["fits", 2**70]]
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows, shared)
    assert native == python
    assert python_rows == 0
    path = tmp_path / ("t." + ("xlsx" if cls is XlsxWriter else "xlsb"))
    with cls(path, use_shared_strings=shared) as book:
        book.add_sheet("S")
        book.write_sheet(rows)
    with ExcelReader(path) as reader:
        values = reader.read_all("S")
    assert values[0][1] == str(10**400) and values[1][1] == str(-(10**400))
    assert values[2][1] == float(2**70)


class _Opaque:
    """An object the writers know nothing about: written as str()."""

    def __str__(self):
        return "opaque"


def _other_cells():
    np = pytest.importorskip("numpy")
    cells = [
        np.int64(-5), np.int32(7), np.uint8(200), np.bool_(True), np.float32(1.5), np.float16(0.25),
        np.datetime64("2024-01-02"), np.str_("numpy text"),
        # Where the rule for a numpy scalar (integer, unsigned, bool, float16/32) ends.
        np.uint64(2**63 + 5), np.int8(-3), np.array(7), np.array(2.5, dtype=np.float32),
        np.array([1, 2]), np.longdouble(1.5), np.complex64(1 + 2j), np.float64(2.5),
        dt.time(3, 4, 5), dt.timedelta(hours=5), b"bytes", _Opaque(), (1, 2, 3), [4, 5],
        (None, "0.00"), (np.int64(3), "0.00"),
    ]
    try:
        import pandas as pd
    except ImportError:
        pass
    else:
        cells += [pd.NaT, pd.NA, pd.Timestamp("2024-01-02 03:04:05"), (pd.NaT, "yyyy-mm-dd")]
    return cells


@WRITERS
@SHARED
@pytest.mark.parametrize("styled", [False, True])
def test_other_values_are_written_natively_as_python_writes_them(tmp_path, monkeypatch, cls, shared, styled):
    cells = _other_cells()
    rows = [["head"] * len(cells)] + [list(cells) for _ in range(30)]
    styles = {"A:ZZ": Style(font=Font(italic=True))} if styled else {}
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows, shared, styles=styles)
    assert native == python
    # The first data row records the (None, format) and (value, format) xfs.
    assert python_rows <= 1


@pytest.mark.parametrize("shared", [True, False])
@pytest.mark.parametrize("styled", [False, True])
def test_xlsb_formulas_are_written_natively(tmp_path, monkeypatch, shared, styled):
    def rows():
        out = [["n", "double", "text", "flag", "formatted"]]
        for i in range(1, 120):
            out.append([
                i, Formula(f"A{i + 1}*2", i * 2), Formula(f'"x"&A{i + 1}', f"x{i}"),
                Formula(f"A{i + 1}>5"), (Formula(f"A{i + 1}/3"), "0.00"),
            ])
        return out
    styles = {"B:C": Style(font=Font(bold=True))} if styled else {}
    native, python, python_rows = _write_both(tmp_path, monkeypatch, XlsbWriter, rows, shared, styles=styles)
    assert native == python
    assert python_rows <= 2


def test_xlsb_formula_errors_are_raised_as_before(tmp_path):
    with pytest.raises(Exception) as native_error:
        with XlsbWriter(tmp_path / "bad.xlsb") as book:
            book.add_sheet("S")
            book.write_sheet([["ok", 1], ["bad", Formula("SUM(")]])
    with pytest.raises(type(native_error.value)):
        with XlsbWriter(tmp_path / "bad2.xlsb") as book:
            book.add_sheet("S")
            book.write_sheet([["bad", Formula("SUM(")]])


def _links_part(path):
    with zipfile.ZipFile(path) as archive:
        return {n: archive.read(n) for n in archive.namelist() if "_rels/sheet" in n}


@WRITERS
@SHARED
@pytest.mark.parametrize("styled", [False, True])
def test_links_are_written_natively(tmp_path, monkeypatch, cls, shared, styled):
    def rows():
        out = [["n", "web", "inside", "text"]]
        for i in range(1, 80):
            out.append([i, Hyperlink(f"https://example.com/{i % 7}", f"site {i}"),
                        Hyperlink(f"#S!A{i}", f"go {i}"), "t"])
        return out
    styles = {"B:C": Style(font=Font(bold=True))} if styled else {}
    native, python, python_rows = _write_both(tmp_path, monkeypatch, cls, rows, shared, styles=styles)
    # The first link row registers the link style (and its variant for the range style).
    assert python_rows <= 1
    if not styled:
        assert native == python
    paths = []
    for name, patch_native in (("native", False), ("python", True)):
        path = tmp_path / f"{name}.{'xlsx' if cls is XlsxWriter else 'xlsb'}"
        with monkeypatch.context() as patch:
            if patch_native:
                none = lambda *args, **kwargs: None
                patch.setattr(_xlsx_writer_sheet, "accel_encode_xlsx_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row_inline", none)
            with cls(path, use_shared_strings=shared) as book:
                book.add_sheet("S")
                book.write_sheet(rows(), styles=styles)
        paths.append(path)
    assert _links_part(paths[0]) == _links_part(paths[1])
    with ExcelReader(paths[0]) as first, ExcelReader(paths[1]) as second:
        assert first.read_all("S") == second.read_all("S")


def _package(path):
    """Every part of a package but the core properties (they hold the time)."""
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


def _rich_rows():
    bold, italic = InlineFont(b=True), InlineFont(i=True, sz=14)
    rich = CellRichText(["plain ", TextBlock(bold, "bold"), " tail"])
    return [
        ["head", rich, CellRichText(["only plain"])],                # the first row has the header style
        [1, CellRichText([TextBlock(italic, "italic")]), "plain"],
        [2, rich, CellRichText(["only plain"])],                      # the same strings again: one entry each
        [3, CellRichText(["plain ", TextBlock(bold, "bold"), " tail"]), None],  # equal, not identical
        [4, (rich, "0.00"), (CellRichText(["x", TextBlock(italic, "y")]), Style(font=Font(bold=True)))],
        [5, "plain ", CellRichText([TextBlock(bold, "")])],           # the text of the rich cell is a plain cell's
    ]


@pytest.mark.parametrize("styled", [False, True])
def test_xlsb_rich_text_is_written_natively_as_python_writes_it(tmp_path, monkeypatch, styled):
    rows = _rich_rows()
    styles = {"A:ZZ": Style(font=Font(italic=True))} if styled else {}
    packages, python_rows = [], []
    for native in (True, False):
        with monkeypatch.context() as patch:
            if not native:
                none = lambda *args, **kwargs: None
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row", none)
            calls = []
            original = _xlsb_writer_bin._XlsbBinMixin._write_python_row

            def counted(self, *args, _original=original, _calls=calls, **kwargs):
                _calls.append(1)
                return _original(self, *args, **kwargs)

            patch.setattr(_xlsb_writer_bin._XlsbBinMixin, "_write_python_row", counted)
            path = tmp_path / f"{native}.xlsb"
            with XlsbWriter(path) as book:
                book.add_sheet("S")
                book.write_sheet(rows, styles=styles)
            packages.append(_package(path))
            python_rows.append(len(calls))
    assert packages[0] == packages[1]
    # Only a (rich text, Style) cell may need the Python encoder once, to register its xf.
    assert python_rows[0] <= 1 < python_rows[1]


def test_xlsb_rich_text_with_inline_strings_still_goes_to_the_table(tmp_path):
    path = tmp_path / "a.xlsb"
    with XlsbWriter(path, use_shared_strings=False) as book:
        book.add_sheet("S")
        book.write_sheet(_rich_rows())
    assert "xl/sharedStrings.bin" in _package(path)
