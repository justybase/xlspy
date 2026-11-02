"""Hidden rows and columns: argument rules and output, in both writers."""
from __future__ import annotations

import struct
import zipfile

import pytest

from xlspy import XlsbWriter, XlsxWriter
from xlspy.biff_utils import iter_records

openpyxl = pytest.importorskip("openpyxl")


def _xlsb_row_flags(path, part="xl/worksheets/sheet1.bin"):
    """Map 0-based row index to the BrtRowHdr flags byte."""
    with zipfile.ZipFile(path) as archive:
        blob = archive.read(part)
    records = list(iter_records(blob))
    assert records[-1].data_end == len(blob), "sheet records must tile the buffer"
    return {
        struct.unpack_from("<I", blob, r.data_start)[0]: blob[r.data_start + 11]
        for r in records if r.record_id == 0x0000
    }


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
@pytest.mark.parametrize("bad", [-1, 16384, "XFE"])
def test_hidden_columns_out_of_range_rejected_by_both_writers(tmp_path, writer_cls, suffix, bad):
    with writer_cls(tmp_path / f"bad.{suffix}") as writer:
        with pytest.raises(ValueError, match="out of range"):
            writer.write_sheet([["x"]], hidden_columns=[bad])
        writer.write_sheet([["x"]], hidden_columns=["XFD", 0])


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
@pytest.mark.parametrize("bad,error", [([-1], TypeError), (["2"], TypeError), ([True], TypeError), (5, TypeError)])
def test_hidden_rows_rules_match_between_writers(tmp_path, writer_cls, suffix, bad, error):
    with writer_cls(tmp_path / f"bad.{suffix}") as writer:
        with pytest.raises(error, match="hidden_rows"):
            writer.write_sheet([["x"]], hidden_rows=bad)


@pytest.mark.parametrize("writer_cls,suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_hidden_rows_and_columns_written_by_both_writers(tmp_path, writer_cls, suffix):
    path = tmp_path / f"hidden.{suffix}"
    with writer_cls(path) as writer:
        writer.write_sheet([["a", "b", "c"], [1, 2, 3], [4, 5, 6]], hidden_rows={1}, hidden_columns=["B"])
    if suffix == "xlsx":
        sheet = openpyxl.load_workbook(path).active
        assert sheet.row_dimensions[2].hidden
        assert sheet.column_dimensions["B"].hidden
    else:
        assert _xlsb_row_flags(path)[1] & 0x10
