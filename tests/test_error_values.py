"""Error cells stay apart from text that starts with ``error:``."""

import pytest

openpyxl = pytest.importorskip("openpyxl")

from xlspy import ExcelReader, Workbook, load_workbook  # noqa: E402
from xlspy._errors import CellError  # noqa: E402
from xlspy.workbook import _xlsb_reader  # noqa: E402

TEXT = "error: disk full"


@pytest.fixture
def xlsx(tmp_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "S"
    sheet["A1"] = TEXT
    sheet["A2"] = "#N/A"
    sheet["A2"].data_type = "e"
    path = tmp_path / "errors.xlsx"
    workbook.save(path)
    return path


@pytest.mark.parametrize("streaming", [False, True])
def test_reader_keeps_error_like_text(xlsx, streaming):
    with ExcelReader(str(xlsx)) as reader:
        legacy = [value for row in reader.get_rows("S", streaming=streaming) for value in row]
    with ExcelReader(str(xlsx), errors="openpyxl") as reader:
        bare = [value for row in reader.get_rows("S", streaming=streaming) for value in row]
    assert legacy == [TEXT, "error:#N/A"]
    assert type(legacy[0]) is str and isinstance(legacy[1], CellError)
    assert bare == [TEXT, "#N/A"]
    assert type(bare[1]) is str


def test_model_keeps_error_like_text(xlsx):
    workbook = load_workbook(xlsx)
    try:
        sheet = workbook.active
        assert (sheet["A1"].value, sheet["A1"].data_type) == (TEXT, "s")
        assert (sheet["A2"].value, sheet["A2"].data_type) == ("#N/A", "e")
    finally:
        workbook.close()


@pytest.mark.parametrize("c_loader", [True, False])
def test_xlsb_model_keeps_error_like_text(tmp_path, monkeypatch, c_loader):
    workbook = Workbook()
    workbook.active["A1"] = TEXT
    path = tmp_path / "text.xlsb"
    workbook.save(path)
    if not c_loader:
        monkeypatch.setattr(_xlsb_reader, "xlsb_load_cells", None)
    loaded = load_workbook(path)
    try:
        cell = loaded.active["A1"]
        assert (cell.value, cell.data_type) == (TEXT, "s")
    finally:
        loaded.close()
