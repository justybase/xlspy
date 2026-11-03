"""The native row encoders write the same worksheet bytes as the Python encoders."""

import datetime as dt
import random
import zipfile

import pytest

from xlspy import ExcelReader, Font, Formula, Style, XlsbWriter, XlsxWriter, _accel, _xlsb_writer_bin, _xlsx_writer_sheet

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

                def counted(self, *args, _original=original, **kwargs):
                    calls.append(1)
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
