"""Members compressed on the worker thread hold the same bytes as sequential ones."""

import io
import threading
import zipfile

import pytest

from xlspy import XlsbWriter, XlsxWriter, _zip_threads


def _rows():
    for i in range(3000):
        yield [f"text {i % 50}", i, i / 7, i % 2 == 0, f"unique {i}"]


def _members(data: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert archive.testzip() is None
        return {name: archive.read(name) for name in archive.namelist()}


def _write(cls, level, shared):
    target = io.BytesIO()
    with cls(target, compression_level=level, use_shared_strings=shared) as writer:
        writer.add_sheet("S")
        writer.write_sheet(_rows(), header_style=None, freeze_panes="B2")
        writer.add_sheet("Empty")
        writer.write_sheet([])
    return target.getvalue()


@pytest.mark.parametrize("cls", [XlsxWriter, XlsbWriter])
@pytest.mark.parametrize("level", [0, 1, 3, 6, 9])
@pytest.mark.parametrize("shared", [True, False])
def test_threaded_members_match_sequential(monkeypatch, cls, level, shared):
    threaded = _members(_write(cls, level, shared))
    monkeypatch.setattr(_zip_threads, "AVAILABLE", False)
    sequential = _members(_write(cls, level, shared))
    assert threaded == sequential


@pytest.mark.parametrize("cls", [XlsxWriter, XlsbWriter])
def test_threaded_output_is_deterministic(cls):
    # A sheet larger than the row buffer is compressed in segments (its head,
    # which holds the dimension, is known only after the rows), so the archive
    # bytes differ from a sequential write; the members do not (see above), and
    # the same input always gives the same bytes.
    assert _write(cls, 3, True) == _write(cls, 3, True)


def test_failed_write_stops_the_worker():
    def failing():
        yield from _rows()
        raise RuntimeError("source failed")

    for cls in (XlsxWriter, XlsbWriter):
        with pytest.raises(RuntimeError, match="source failed"):
            with cls(io.BytesIO()) as writer:
                writer.add_sheet("S")
                writer.write_sheet(failing())
    assert not [t for t in threading.enumerate() if t.name.startswith("xlspy-deflate")]
