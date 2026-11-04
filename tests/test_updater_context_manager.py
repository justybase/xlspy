"""XlsxUpdater and XlsbUpdater as context managers: leaving the block releases the updater, and only save() keeps changes."""

import pytest

from xlspy import Workbook, XlsbWriter, XlsbUpdater, XlsxUpdater, load_workbook


def _xlsx(path):
    workbook = Workbook()
    workbook.active["A1"] = "original"
    workbook.save(path)
    workbook.close()
    return path


def _xlsb(path):
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet([["original"]])
    return path


@pytest.fixture(params=[
    pytest.param((XlsxUpdater, _xlsx, "xlsx"), id="xlsx"),
    pytest.param((XlsbUpdater, _xlsb, "xlsb"), id="xlsb"),
])
def updater_case(request):
    return request.param


@pytest.fixture
def closes(monkeypatch):
    """Counts the close() calls of both updater classes."""
    calls = []
    for cls in (XlsxUpdater, XlsbUpdater):
        original = cls.close

        def spy(self, _original=original):
            calls.append(type(self).__name__)
            _original(self)

        monkeypatch.setattr(cls, "close", spy)
    return calls


def _value(path):
    workbook = load_workbook(path)
    try:
        return workbook.active["A1"].value
    finally:
        workbook.close()


def test_saving_inside_the_block_keeps_the_changes_and_releases_the_updater(tmp_path, updater_case, closes):
    cls, make, suffix = updater_case
    source = make(tmp_path / f"source.{suffix}")
    out = tmp_path / f"out.{suffix}"
    with cls(source) as updater:
        updater.replace_sheet_data(updater.get_sheet_names()[0], [["changed"]])
        updater.save(out)
    assert closes == [cls.__name__]
    assert _value(out) == "changed"


def test_leaving_the_block_without_saving_discards_the_changes(tmp_path, updater_case, closes):
    cls, make, suffix = updater_case
    source = make(tmp_path / f"source.{suffix}")
    before = source.read_bytes()
    with cls(source) as updater:
        updater.replace_sheet_data(updater.get_sheet_names()[0], [["changed"]])
    assert closes == [cls.__name__]
    assert source.read_bytes() == before


def test_an_exception_in_the_block_releases_the_updater_and_propagates(tmp_path, updater_case, closes):
    cls, make, suffix = updater_case
    source = make(tmp_path / f"source.{suffix}")
    out = tmp_path / f"out.{suffix}"
    with pytest.raises(RuntimeError, match="stop"):
        with cls(source) as updater:
            updater.replace_sheet_data(updater.get_sheet_names()[0], [["changed"]])
            raise RuntimeError("stop")
    assert closes == [cls.__name__]
    assert not out.exists()


def test_the_updater_enters_as_itself(tmp_path, updater_case):
    cls, make, suffix = updater_case
    updater = cls(make(tmp_path / f"source.{suffix}"))
    with updater as entered:
        assert entered is updater
