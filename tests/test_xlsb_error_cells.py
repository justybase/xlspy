"""XLSB error cells (BrtCellError) read as error values on every path."""

import zipfile
from pathlib import Path

import pytest

from xlspy import ExcelReader, XlsbWriter
from xlspy import reader as reader_module
from xlspy._errors import CellError
from xlspy.biff_utils import build_record

SHEET_PART = "xl/worksheets/sheet1.bin"
BRT_ROW_HDR = 0x00
BRT_CELL_ERROR = 0x03
BRT_FMLA_ERROR = 0x0B
VALUE_CELLS = (0x02, 0x05)  # BrtCellRk, BrtCellReal
NA = 0x2A  # the error code for #N/A


def _read_vint(data, pos):
    value = shift = 0
    while True:
        byte = data[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        if not byte & 0x80:
            return value, pos
        shift += 7


def _records(data):
    pos = 0
    while pos < len(data):
        rec_id, pos = _read_vint(data, pos)
        size, pos = _read_vint(data, pos)
        yield rec_id, data[pos:pos + size]
        pos += size


def _with_error_cell(payload, *, row, code, record=BRT_CELL_ERROR, length=9):
    """Turn the first value cell of ``row`` into an error cell of ``record`` and ``length`` bytes with ``code``."""
    out = []
    current = -1
    replaced = False
    for rec_id, body in _records(payload):
        if rec_id == BRT_ROW_HDR:
            current += 1
        elif current == row and not replaced and rec_id in VALUE_CELLS and len(body) >= 8:
            # The cell header (column and style) stays; the error code follows it.
            rec_id, body = record, body[:8] + bytes([code])[:length - 8]
            replaced = True
        out.append(build_record(rec_id, body))
    assert replaced, "no value cell to replace in the generated sheet"
    return b"".join(out)


@pytest.fixture
def error_workbook(tmp_path):
    source = tmp_path / "numbers.xlsb"
    with XlsbWriter(str(source)) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[10.0, 20.0, 30.0], [40.0, 50.0, 60.0]])
    path = tmp_path / "errors.xlsb"
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            payload = zin.read(name)
            if name == SHEET_PART:
                payload = _with_error_cell(payload, row=1, code=NA)
            zout.writestr(name, payload)
    return path


@pytest.mark.parametrize("native", [True, False])
def test_bulk_read_returns_error_values(error_workbook, monkeypatch, native):
    if native and not reader_module.HAVE_NATIVE:
        pytest.skip("native extension not built")
    monkeypatch.setattr(reader_module, "HAVE_NATIVE", native)
    with ExcelReader(str(error_workbook)) as reader:
        rows = reader.read_all("S")
    assert rows[0] == [10, 20, 30]
    assert rows[1][1:] == [50, 60]
    assert isinstance(rows[1][0], CellError) and rows[1][0] == "error:#N/A"


@pytest.mark.parametrize("native", [True, False])
def test_openpyxl_mode_returns_bare_error_text(error_workbook, monkeypatch, native):
    if native and not reader_module.HAVE_NATIVE:
        pytest.skip("native extension not built")
    monkeypatch.setattr(reader_module, "HAVE_NATIVE", native)
    with ExcelReader(str(error_workbook), errors="openpyxl") as reader:
        rows = reader.read_all("S")
    assert rows[1][0] == "#N/A"
    assert type(rows[1][0]) is str


def test_streaming_read_returns_error_values(error_workbook):
    with ExcelReader(str(error_workbook)) as reader:
        rows = reader.read_all("S", streaming=True)
    assert isinstance(rows[1][0], CellError) and rows[1][0] == "error:#N/A"


def test_reader_accepts_path_objects(error_workbook):
    with ExcelReader(error_workbook) as reader:
        assert reader.get_sheet_names() == ["S"]
    reader = ExcelReader()
    try:
        reader.open(Path(error_workbook))
        assert reader.read_all("S")[1][0] == CellError("error:#N/A")
    finally:
        reader.close()


@pytest.mark.parametrize("length", [8, 9])
def test_a_formula_error_cell_is_read_the_same_by_both_readers(tmp_path, monkeypatch, length):
    """A BrtFmlaError is a cell only with its 9-byte header and error code, as in the Python reader."""
    # One value per row, so the replaced cell has no compact neighbours in its row.
    source = tmp_path / "numbers.xlsb"
    with XlsbWriter(str(source)) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[10.0], [40.0], [70.0]])
    path = tmp_path / "formula-errors.xlsb"
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            payload = zin.read(name)
            if name == SHEET_PART:
                payload = _with_error_cell(payload, row=1, code=NA, record=BRT_FMLA_ERROR, length=length)
            zout.writestr(name, payload)
    results = {}
    for native in (True, False):
        if native and not reader_module.HAVE_NATIVE:
            pytest.skip("native extension not built")
        monkeypatch.setattr(reader_module, "HAVE_NATIVE", native)
        with ExcelReader(str(path)) as reader:
            results[native] = reader.read_all("S")
    assert results[True] == results[False]
    assert results[False][0] == [10]
    assert results[False][2] == [70]
    if length == 8:
        assert not any(isinstance(value, CellError) for value in results[False][1])
    else:
        assert results[False][1] == [CellError("error:#N/A")]
