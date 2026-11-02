"""Corrupt workbooks raise InvalidFileError from every public read path (ExcelReader,
load_workbook, the updaters), with the decoder's own error chained as __cause__.

InvalidFileError is a ValueError, so ``except ValueError`` keeps working."""
from __future__ import annotations

import random
import struct
import zipfile
import zlib
from pathlib import Path
from xml.etree.ElementTree import ParseError
from xml.parsers.expat import ExpatError

import pytest

from xlspy import ExcelReader, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, _accel, load_workbook
from xlspy.exceptions import InvalidFileError

ROWS = [["Name", "Value"], ["alpha", 1.5], ["beta", 2]]
RAW_DECODER_ERRORS = (ParseError, ExpatError, ValueError, struct.error, IndexError, zlib.error, EOFError)


def _workbook(tmp_path: Path, kind: str) -> Path:
    path = tmp_path / f"source.{kind}"
    with (XlsbWriter if kind == "xlsb" else XlsxWriter)(str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS)
    return path


def _replace_part(source: Path, target: Path, part: str, data: bytes) -> Path:
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info.filename, data if info.filename == part else src.read(info.filename))
    return target


def _read_rows(path: Path) -> None:
    with ExcelReader(str(path)) as reader:
        for name in reader.get_sheet_names():
            list(reader.get_rows(name))


def _read_rows_streaming(path: Path) -> None:
    with ExcelReader(str(path)) as reader:
        for name in reader.get_sheet_names():
            list(reader.get_rows(name, streaming=True))


def _open_updater(path: Path) -> None:
    (XlsbUpdater if path.suffix == ".xlsb" else XlsxUpdater)(str(path)).close()


READERS = [_read_rows, _read_rows_streaming, load_workbook, _open_updater]


@pytest.mark.parametrize("read", READERS)
@pytest.mark.parametrize("part", ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml"])
def test_malformed_xml_part(tmp_path: Path, read, part: str) -> None:
    path = _replace_part(_workbook(tmp_path, "xlsx"), tmp_path / "bad.xlsx", part, b"<worksheet><row>")
    if read is _open_updater and part != "xl/workbook.xml":
        pytest.skip("opening an updater reads only the workbook part")
    with pytest.raises(InvalidFileError, match="bad.xlsx") as caught:
        read(path)
    assert isinstance(caught.value.__cause__, (ParseError, ExpatError))


def test_undecodable_text_in_a_sheet(tmp_path: Path) -> None:
    sheet = b'<?xml version="1.0" encoding="UTF-8"?><worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>\xff\xfe</t></is></c></row></sheetData></worksheet>'
    path = _replace_part(_workbook(tmp_path, "xlsx"), tmp_path / "bad.xlsx", "xl/worksheets/sheet1.xml", sheet)
    for read in (_read_rows, load_workbook):
        with pytest.raises(InvalidFileError):
            read(path)


def _varint(data: bytes, pos: int) -> tuple[int, int]:
    value = shift = 0
    while True:
        byte = data[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        shift += 7
        if not byte & 0x80:
            return value, pos


def _cut_sheet_files(tmp_path: Path) -> list[tuple[Path, bool]]:
    """The sheet part cut after every byte inside its records, each with whether the
    cut falls in the cell data (up to BrtEndSheetData). A cut between two records
    is a clean end of the data, so those are left out."""
    source = _workbook(tmp_path, "xlsb")
    with zipfile.ZipFile(source) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
    boundaries, pos, cells_end = {0}, 0, 0
    while pos < len(sheet):
        record_id, pos = _varint(sheet, pos)
        length, pos = _varint(sheet, pos)
        pos += length
        boundaries.add(pos)
        if record_id == 0x92:  # BrtEndSheetData
            cells_end = pos
    assert cells_end
    return [
        (_replace_part(source, tmp_path / f"cut{cut}.xlsb", "xl/worksheets/sheet1.bin", sheet[:cut]),
         cut < cells_end)
        for cut in range(1, len(sheet)) if cut not in boundaries
    ]


@pytest.mark.parametrize("backend", ["native", "python"])
def test_truncated_xlsb_record(tmp_path: Path, monkeypatch, backend: str) -> None:
    # A record cut short is a damaged file: no reader returns the rows before the cut
    # as if they were the whole sheet. ExcelReader stops reading at the end of the
    # cell data, so a cut after it, which loses no cell, is checked by the model only.
    if backend == "native" and not _accel.HAVE_NATIVE:
        pytest.skip("native extension not built")
    if backend == "python":
        import xlspy.reader as reader_module

        monkeypatch.setattr(reader_module, "_accel_read_xlsb_worksheet", lambda *args, **kwargs: None)
        monkeypatch.setattr(_accel, "xlsb_load_cells", None)
    files = _cut_sheet_files(tmp_path)
    assert len(files) > 20
    for path, in_cells in files:
        for read in (_read_rows, _read_rows_streaming, load_workbook) if in_cells else (load_workbook,):
            with pytest.raises(InvalidFileError):
                read(path)


def test_invalid_file_error_is_still_a_value_error(tmp_path: Path) -> None:
    path = _replace_part(_workbook(tmp_path, "xlsx"), tmp_path / "bad.xlsx", "xl/workbook.xml", b"<")
    with pytest.raises(ValueError):
        load_workbook(path)


def test_unknown_sheet_name_is_not_an_invalid_file(tmp_path: Path) -> None:
    with ExcelReader(str(_workbook(tmp_path, "xlsx"))) as reader:
        with pytest.raises(KeyError):
            list(reader.get_rows("Missing"))


def _mutate(data: bytes, rng: random.Random) -> bytes:
    out = bytearray(data)
    operation = rng.randrange(3)
    if operation == 0 and out:
        for _ in range(rng.randint(1, 8)):
            out[rng.randrange(len(out))] = rng.randrange(256)
    elif operation == 1:
        del out[rng.randrange(len(out) + 1):]
    else:
        at = rng.randrange(len(out) + 1)
        out[at:at] = bytes(rng.randrange(256) for _ in range(rng.randint(1, 16)))
    return bytes(out)


@pytest.mark.parametrize("kind", ["xlsx", "xlsb"])
def test_no_raw_decoder_error_escapes(tmp_path: Path, kind: str) -> None:
    source = _workbook(tmp_path, kind)
    with zipfile.ZipFile(source) as archive:
        blobs = {name: archive.read(name) for name in archive.namelist()}
    parts = [name for name in blobs if name.startswith("xl/") and not name.endswith(".rels")]
    rng = random.Random(20261102)
    for case in range(60):
        part = rng.choice(parts)
        path = _replace_part(source, tmp_path / f"case{case}.{kind}", part, _mutate(blobs[part], rng))
        for read in (_read_rows, load_workbook):
            try:
                read(path)
            except InvalidFileError:
                pass
            except RAW_DECODER_ERRORS as exc:
                pytest.fail(f"{read.__name__} on corrupt {part} (case {case}) raised {exc!r}")
            except (KeyError, TypeError):
                pass  # a sheet listed without its part, or a damaged value of the wrong type


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
def test_a_panic_in_a_native_reader_is_an_invalid_file(tmp_path: Path, monkeypatch) -> None:
    import xlspy._rs_core as native
    from xlspy import reader as reader_module

    def panic(*_args: object) -> None:
        raise native.PanicException("index out of bounds")

    monkeypatch.setattr(reader_module, "_accel_read_xlsb_worksheet", panic)
    with pytest.raises(InvalidFileError, match="PanicException") as caught:
        _read_rows(_workbook(tmp_path, "xlsb"))
    assert isinstance(caught.value.__cause__, native.PanicException)


def test_interrupts_pass_through(tmp_path: Path, monkeypatch) -> None:
    from xlspy.exceptions import reading

    with pytest.raises(KeyboardInterrupt):
        with reading("book.xlsx"):
            raise KeyboardInterrupt
