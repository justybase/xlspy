"""The native row encoders write the same worksheet bytes as the Python encoders."""

import datetime as dt
import random
import zipfile

import pytest

from xlspy import XlsbWriter, XlsxWriter, _accel, _xlsb_writer_bin, _xlsb_writer_parts, _xlsx_writer_sheet, writer, xlsx_writer

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
    for use_c in (True, False):
        with monkeypatch.context() as patch:
            if not use_c:
                none = lambda *args, **kwargs: None
                patch.setattr(_xlsx_writer_sheet, "accel_encode_xlsx_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row", none)
                patch.setattr(_xlsb_writer_bin, "encode_xlsb_row_inline", none)
            path = tmp_path / f"{use_c}.{suffix}"
            with cls(path, useSharedStrings=shared) as book:
                book.add_sheet("S")
                book.write_sheet(rows)
            outputs.append(_sheet(path))
    assert outputs[0] == outputs[1]
