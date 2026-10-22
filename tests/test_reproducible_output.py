"""The same workbook gives the same bytes on every save."""

import io
import zipfile

import pytest

from xlspy import Chart, ChartSeries, ExcelTable, XlsbWriter, XlsxWriter


def _save(writer_class):
    stream = io.BytesIO()
    with writer_class(stream) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Region", "Sales"], ["North", 10], ["South", 30]],
            tables=[ExcelTable("Sales", "A1:B3")],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")])],
        )
    return stream.getvalue()


@pytest.mark.parametrize("writer_class", [XlsxWriter, XlsbWriter])
def test_two_saves_are_identical(writer_class, monkeypatch):
    first = _save(writer_class)
    # A later clock must not change the archive.
    monkeypatch.setattr("time.time", lambda: 2_000_000_000.0)
    assert _save(writer_class) == first


@pytest.mark.parametrize("writer_class", [XlsxWriter, XlsbWriter])
def test_members_carry_excel_timestamp_and_compression(writer_class):
    with zipfile.ZipFile(io.BytesIO(_save(writer_class))) as archive:
        infos = archive.infolist()
        assert {info.date_time for info in infos} == {(1980, 1, 1, 0, 0, 0)}
        assert {info.compress_type for info in infos} == {zipfile.ZIP_DEFLATED}
        assert archive.testzip() is None


def test_xlsb_table_ids_differ_per_table():
    stream = io.BytesIO()
    with XlsbWriter(stream) as writer:
        writer.add_sheet("A")
        writer.write_sheet([["a", "b"], [1, 2]], tables=[ExcelTable("First", "A1:B2")])
        writer.add_sheet("B")
        writer.write_sheet([["a", "b"], [1, 2]], tables=[ExcelTable("Second", "A1:B2")])
    with zipfile.ZipFile(stream) as archive:
        first, second = (archive.read(f"xl/tables/table{i}.bin") for i in (1, 2))
    # FRTBegin (0x0C00, written as 0x80 0x18) holds the revision GUID.
    marker = b"\x80\x18\x10"
    guids = {part[part.index(marker) + 3:part.index(marker) + 19] for part in (first, second)}
    assert len(guids) == 2
