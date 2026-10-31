"""Resource limits on untrusted package input: declared part size and row width."""

import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, _accel, reader as reader_module
from xlspy import _untrusted
from xlspy._untrusted import SafeZipFile, check_column_span


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


def test_python_xlsb_reader_refuses_a_row_wider_than_the_limit(tmp_path, monkeypatch):
    path = tmp_path / "wide.xlsb"
    with XlsbWriter(str(path)) as workbook:
        workbook.add_sheet("S")
        workbook.write_sheet([["a", None, "c"]])
    # Shrink the limit so a three-column row is over it; the native path is
    # bypassed, so the Python reader has to make the decision.
    monkeypatch.setattr(_untrusted, "MAX_COLUMNS", 2)
    monkeypatch.setattr(reader_module, "HAVE_NATIVE", False)
    with ExcelReader(str(path)) as reader:
        with pytest.raises(ValueError, match="wider than"):
            list(reader.get_rows("S"))


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
