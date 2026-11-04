"""Resource limits on untrusted package input: declared part size and row width."""

import time
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, _accel, reader as reader_module
from xlspy import _untrusted
from xlspy._untrusted import SafeZipFile, check_column_span
from xlspy.biff_utils import build_record, iter_record_payloads


def test_check_column_span_accepts_a_full_worksheet_width():
    check_column_span(0, _untrusted.MAX_COLUMNS - 1)  # A..XFD is allowed


def test_check_column_span_refuses_a_row_wider_than_excel():
    with pytest.raises(ValueError, match="wider than"):
        check_column_span(0, _untrusted.MAX_COLUMNS)


def test_safe_zip_reads_an_ordinary_member(tmp_path):
    path = tmp_path / "plain.zip"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", b"<workbook/>")
    with SafeZipFile(path) as archive:
        assert archive.read("xl/workbook.xml") == b"<workbook/>"


def test_safe_zip_refuses_a_member_declared_over_the_limit(tmp_path, monkeypatch):
    path = tmp_path / "big.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/worksheets/sheet1.xml", b"x" * 4096)
    monkeypatch.setattr(_untrusted, "MAX_PART_BYTES", 1024)
    with SafeZipFile(path) as archive:
        with pytest.raises(ValueError, match="over the 1024-byte limit"):
            archive.read("xl/worksheets/sheet1.xml")


def test_safe_zip_still_raises_key_error_for_a_missing_member(tmp_path):
    # read_part() relies on KeyError for a package without the part.
    path = tmp_path / "empty.zip"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("a.txt", b"a")
    with SafeZipFile(path) as archive:
        with pytest.raises(KeyError):
            archive.read("missing.xml")


@pytest.mark.parametrize("backend", ["native", "python"])
def test_xlsb_cell_past_column_xfd_is_refused_at_once(tmp_path, monkeypatch, backend):
    # A cell in column 0x7FFFFFFF after a normal first row: the Python reader
    # used to pad the row up to it (minutes and gigabytes) and the native one
    # dropped it. Both refuse it now, as load_workbook does.
    if backend == "native" and not _accel.HAVE_NATIVE:
        pytest.skip("native extension not built")
    source = tmp_path / "source.xlsb"
    with XlsbWriter(str(source)) as workbook:
        workbook.add_sheet("S")
        workbook.write_sheet([["a", "b"], [1, 2]])
    part = "xl/worksheets/sheet1.bin"
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    records = list(iter_record_payloads(parts[part]))
    rows = [i for i, (rec_id, _body) in enumerate(records) if rec_id == 0x00]
    # The last cell of the second row gets column 0x7FFFFFFF.
    last = max(i for i, (rec_id, _body) in enumerate(records) if 0x01 <= rec_id <= 0x0B and i > rows[1])
    rec_id, body = records[last]
    records[last] = (rec_id, (0x7FFFFFFF).to_bytes(4, "little") + bytes(body[4:]))
    parts[part] = b"".join(build_record(rec_id, body) for rec_id, body in records)
    target = tmp_path / "far-column.xlsb"
    with zipfile.ZipFile(target, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    if backend == "python":
        monkeypatch.setattr(reader_module, "HAVE_NATIVE", False)
    started = time.perf_counter()
    with ExcelReader(str(target)) as reader:
        with pytest.raises(ValueError, match="column exceeds Excel limits"):
            list(reader.get_rows("S"))
    assert time.perf_counter() - started < 5


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="the buffer cap belongs to the native XLSX reader")
def test_xlsx_row_that_never_closes_is_refused_at_the_buffer_cap(tmp_path):
    # A row whose closing tag never arrives must not grow the native buffer
    # without bound. The body compresses to almost nothing, like a bomb would.
    from xlspy import XlsxWriter

    fixture = tmp_path / "fixture.xlsx"
    with XlsxWriter(str(fixture)) as workbook:
        workbook.add_sheet("S")
        workbook.write_sheet([["a"]])
    body = b"x" * ((64 << 20) + 4096)
    sheet = (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>' + body
    )
    path = tmp_path / "open_row.xlsx"
    with zipfile.ZipFile(fixture) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = sheet if info.filename == "xl/worksheets/sheet1.xml" else src.read(info.filename)
            dst.writestr(info.filename, data)
    with ExcelReader(str(path)) as reader:
        with pytest.raises(ValueError, match="without closing a row"):
            list(reader.get_rows("S"))
