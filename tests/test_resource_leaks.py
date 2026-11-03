"""Resources the public API must give back: objects, threads, staging files.

Writers, readers and updaters are used the normal way, abandoned half-way
and broken by an exception; afterwards nothing of theirs may stay alive, no
worker thread may keep running and no temporary file may be left behind.
Runs with the native extension and with ``XLSPY_ACCEL=python`` alike.
"""

from __future__ import annotations

import gc
import os
import tempfile
import threading
import warnings

import pytest

from xlspy import ExcelReader, Formula, Hyperlink, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook

ROWS = [["Name", "Amount", "Rate"]] + [[f"name {i}", i, (i / 7, "0.00%")] for i in range(200)]


def _xlspy_objects() -> int:
    count = 0
    for obj in gc.get_objects():
        module = getattr(type(obj), "__module__", None)
        if isinstance(module, str) and module.startswith("xlspy"):
            count += 1
    return count


def _temp_entries(*dirs) -> set[str]:
    entries = set()
    for directory in dirs:
        entries |= {os.path.join(directory, name) for name in os.listdir(directory)}
    return entries


def assert_released(run, directory, rounds: int = 3) -> None:
    """``run()`` leaves no xlspy object, thread, file or ResourceWarning behind."""
    run()  # warm-up: module-level caches fill here
    gc.collect()
    objects, threads = _xlspy_objects(), threading.active_count()
    entries = _temp_entries(tempfile.gettempdir(), directory)
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", ResourceWarning)
        for _ in range(rounds):
            run()
        gc.collect()
    assert not [w for w in caught if issubclass(w.category, ResourceWarning)]
    assert threading.active_count() == threads
    assert _temp_entries(tempfile.gettempdir(), directory) <= entries
    assert _xlspy_objects() <= objects


WRITERS = pytest.mark.parametrize("cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
UPDATERS = pytest.mark.parametrize("cls,ext", [(XlsxUpdater, "xlsx"), (XlsbUpdater, "xlsb")])


def _write(cls, path, rows=ROWS):
    with cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, autofit_rows=None, auto_filter=True)
        writer.add_sheet("Links")
        writer.write_sheet([[Hyperlink("https://example.com", "x"), Formula("=1+1", 2)]] * 20)


@WRITERS
def test_writer_releases_everything(tmp_path, cls, ext):
    assert_released(lambda: _write(cls, tmp_path / f"out.{ext}"), tmp_path)


@WRITERS
def test_writer_failing_mid_sheet_releases_everything(tmp_path, cls, ext):
    def rows():
        yield ["a"]
        for i in range(300):
            if i == 150:
                raise KeyError("source failed")
            yield [i]

    def run():
        with pytest.raises(KeyError):
            _write(cls, tmp_path / f"broken.{ext}", rows())

    assert_released(run, tmp_path)
    assert not (tmp_path / f"broken.{ext}").exists()


@WRITERS
def test_abandoned_writer_releases_everything(tmp_path, cls, ext):
    def run():
        writer = cls(tmp_path / f"abandoned.{ext}")
        writer.__enter__()
        writer.add_sheet("Data")
        writer.write_sheet(ROWS)
        del writer

    assert_released(run, tmp_path)
    assert not (tmp_path / f"abandoned.{ext}").exists()


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
@pytest.mark.parametrize("streaming", [False, True])
def test_abandoned_row_iterator_releases_everything(tmp_path, ext, streaming):
    path = tmp_path / f"source.{ext}"
    _write(XlsbWriter if ext == "xlsb" else XlsxWriter, path)

    def run():
        reader = ExcelReader(str(path))
        reader.open(str(path))
        rows = reader.get_rows("Data", streaming=streaming)
        next(rows)
        del rows, reader

    assert_released(run, tmp_path)


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_abandoned_read_only_workbook_releases_everything(tmp_path, ext):
    path = tmp_path / f"source.{ext}"
    _write(XlsbWriter if ext == "xlsb" else XlsxWriter, path)

    def run():
        workbook = load_workbook(path, read_only=True)
        rows = workbook.worksheets[0].iter_rows(values_only=True)
        next(rows)
        del rows, workbook

    assert_released(run, tmp_path)


@UPDATERS
@pytest.mark.parametrize("finish", ["save", "abandon", "fail"])
def test_updater_releases_everything(tmp_path, cls, ext, finish):
    path = tmp_path / f"template.{ext}"
    _write(XlsbWriter if ext == "xlsb" else XlsxWriter, path)

    def rows():
        for i in range(100):
            if finish == "fail" and i == 50:
                raise KeyError("source failed")
            yield [f"n{i}", i]

    def run():
        updater = cls(path)
        try:
            updater.replace_sheet_data_stream("Data", rows(), headers=["Name", "Amount"])
            if finish == "save":
                updater.save(tmp_path / f"updated.{ext}")
        except KeyError:
            pass
        if finish != "abandon":
            updater.close()
        del updater

    assert_released(run, tmp_path)


# A failed open must give its resources back at once, not when the exception
# (whose traceback still references the half-built object) is collected: on
# Windows an open handle keeps the file locked.

def _without_part(source, target, part):
    import zipfile

    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w") as zout:
        for info in zin.infolist():
            if info.filename != part:
                zout.writestr(info, zin.read(info))


def _open_handles(path) -> list[str]:
    psutil = pytest.importorskip("psutil")
    name = os.path.normcase(os.path.basename(path))
    return [f.path for f in psutil.Process().open_files() if os.path.normcase(os.path.basename(f.path)) == name]


def _updater_dirs() -> set[str]:
    root = tempfile.gettempdir()
    return {name for name in os.listdir(root) if name.startswith("xlspy-updater-")}


def test_failed_read_only_load_closes_the_archive(tmp_path):
    source, broken = tmp_path / "source.xlsb", tmp_path / "broken.xlsb"
    _write(XlsbWriter, source)
    _without_part(source, broken, "xl/worksheets/sheet1.bin")
    with pytest.raises(ValueError) as caught:
        load_workbook(broken, read_only=True)
    assert caught.value is not None  # the traceback is still alive here
    assert _open_handles(broken) == []


@pytest.mark.parametrize("cls,wrong", [(XlsxUpdater, XlsbWriter), (XlsbUpdater, XlsxWriter)])
def test_updater_refusing_the_format_removes_its_staging_dir(tmp_path, cls, wrong):
    path = tmp_path / ("source.xlsb" if wrong is XlsbWriter else "source.xlsx")
    _write(wrong, path)
    before = _updater_dirs()
    with pytest.raises(ValueError) as caught:
        cls(path)
    assert caught.value is not None
    assert _updater_dirs() <= before
    assert _open_handles(path) == []


def test_failed_xlsb_model_load_releases_the_updater(tmp_path, monkeypatch):
    from xlspy import reader as reader_module

    path = tmp_path / "source.xlsb"
    _write(XlsbWriter, path)

    def broken_open(self, *args, **kwargs):
        raise ValueError("reader failed")

    monkeypatch.setattr(reader_module.ExcelReader, "open", broken_open)
    before = _updater_dirs()
    with pytest.raises(ValueError, match="reader failed") as caught:
        load_workbook(path)
    assert caught.value is not None
    assert _updater_dirs() <= before
    assert _open_handles(path) == []
