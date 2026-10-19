"""Workbook output is published only after a successful save.

Path targets are staged in a temporary file next to the target and moved into
place; file-object targets are staged in a spool and copied in on success. A
failure at any stage leaves the target as it was.
"""

import io
import tempfile
import warnings
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter

WRITERS = [
    pytest.param(XlsxWriter, "xlsx", id="xlsx"),
    pytest.param(XlsbWriter, "xlsb", id="xlsb"),
]
TARGETS = ["path", "bytesio", "write_only_stream"]


class WriteOnlyStream(io.RawIOBase):
    """A writable binary stream that cannot seek, like a socket or a pipe."""

    def __init__(self):
        super().__init__()
        self.chunks = []

    def writable(self):
        return True

    def write(self, data):
        self.chunks.append(bytes(data))
        return len(data)

    @property
    def data(self):
        return b"".join(self.chunks)


class FailingStream(WriteOnlyStream):
    """Accepts ``limit`` bytes, then fails like a full disk."""

    def __init__(self, limit):
        super().__init__()
        self.limit = limit

    def write(self, data):
        if sum(map(len, self.chunks)) + len(data) > self.limit:
            raise OSError("no space left on device")
        return super().write(data)


def _target(kind, tmp_path, ext):
    if kind == "path":
        return tmp_path / f"out.{ext}"
    if kind == "bytesio":
        return io.BytesIO()
    return WriteOnlyStream()


def _published(target):
    """The bytes the target received, or None when nothing was published."""
    if isinstance(target, io.BytesIO):
        return target.getvalue() or None
    if isinstance(target, WriteOnlyStream):
        return target.data or None
    return target.read_bytes() if target.exists() else None


def _rows_of(data, ext, tmp_path, sheet):
    path = tmp_path / f"check.{ext}"
    path.write_bytes(data)
    with ExcelReader(str(path)) as reader:
        return list(reader.get_rows(sheet))


def _leftovers(directory):
    return sorted(p.name for p in directory.iterdir() if p.name.endswith(".tmp"))


def _exploding_rows():
    yield ["a", "b"]
    yield [1, 2]
    raise RuntimeError("source failed mid-export")


def _two_sheets(writer):
    writer.add_sheet("First")
    writer.write_sheet([["name", "qty"], ["apple", 3]])
    writer.add_sheet("Second")
    writer.write_sheet([["x"], [1.5]])


# ------------------------------------------------------------- success


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("kind", TARGETS)
def test_successful_write_publishes_a_readable_workbook(tmp_path, writer_cls, ext, kind):
    target = _target(kind, tmp_path, ext)
    with writer_cls(target) as writer:
        _two_sheets(writer)
    data = _published(target)
    assert data is not None
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert archive.testzip() is None
    assert _rows_of(data, ext, tmp_path, "First") == [["name", "qty"], ["apple", 3]]
    assert _rows_of(data, ext, tmp_path, "Second") == [["x"], [1.5]]
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_workbook_is_appended_after_existing_stream_content(tmp_path, writer_cls, ext):
    stream = io.BytesIO()
    stream.write(b"PREFIX")
    with writer_cls(stream) as writer:
        _two_sheets(writer)
    data = stream.getvalue()
    assert data.startswith(b"PREFIX")
    assert _rows_of(data[len(b"PREFIX"):], ext, tmp_path, "First")[1] == ["apple", 3]


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_successful_write_replaces_an_existing_file(tmp_path, writer_cls, ext):
    path = tmp_path / f"out.{ext}"
    path.write_bytes(b"old contents")
    with writer_cls(path) as writer:
        _two_sheets(writer)
    assert _rows_of(path.read_bytes(), ext, tmp_path, "First")[1] == ["apple", 3]
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_explicit_save_inside_the_block_publishes_once(tmp_path, writer_cls, ext):
    path = tmp_path / f"out.{ext}"
    with warnings.catch_warnings():
        warnings.simplefilter("error")  # zipfile warns on duplicate members
        with writer_cls(path) as writer:
            _two_sheets(writer)
            writer.save()
    assert _rows_of(path.read_bytes(), ext, tmp_path, "Second") == [["x"], [1.5]]


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_manual_enter_save_close_publishes(tmp_path, writer_cls, ext):
    buffer = io.BytesIO()
    writer = writer_cls(buffer).__enter__()
    _two_sheets(writer)
    writer.save()
    assert buffer.getvalue() == b""  # nothing is published before close()
    writer.close()
    assert _rows_of(buffer.getvalue(), ext, tmp_path, "First")[1] == ["apple", 3]


# ------------------------------------------------------------- failures


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("kind", TARGETS)
def test_generator_failure_publishes_nothing(tmp_path, writer_cls, ext, kind):
    target = _target(kind, tmp_path, ext)
    with pytest.raises(RuntimeError, match="source failed"):
        with writer_cls(target) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(_exploding_rows())
    assert _published(target) is None
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_failure_keeps_the_existing_file(tmp_path, writer_cls, ext):
    path = tmp_path / f"out.{ext}"
    with writer_cls(path) as writer:
        _two_sheets(writer)
    before = path.read_bytes()
    with pytest.raises(RuntimeError, match="source failed"):
        with writer_cls(path) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(_exploding_rows())
    assert path.read_bytes() == before
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_invalid_cell_value_publishes_nothing(tmp_path, writer_cls, ext):
    buffer = io.BytesIO()
    with pytest.raises(ValueError, match="32767"):
        with writer_cls(buffer) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([["ok"], ["x" * 40_000]])
    assert buffer.getvalue() == b""


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("kind", ["path", "bytesio"])
def test_failure_in_a_later_sheet_discards_the_whole_workbook(tmp_path, writer_cls, ext, kind):
    target = _target(kind, tmp_path, ext)
    with pytest.raises(RuntimeError, match="source failed"):
        with writer_cls(target) as writer:
            writer.add_sheet("Good")
            writer.write_sheet([["fine"], [1]])
            writer.add_sheet("Bad")
            writer.write_sheet(_exploding_rows())
    assert _published(target) is None
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext,method", [
    pytest.param(XlsxWriter, "xlsx", "_sheet_tail", id="xlsx"),
    pytest.param(XlsbWriter, "xlsb", "_create_worksheet_rels", id="xlsb"),
])
def test_serialization_error_publishes_nothing(tmp_path, monkeypatch, writer_cls, ext, method):
    def broken(*args, **kwargs):
        raise RuntimeError("serializer failed")

    monkeypatch.setattr(writer_cls, method, broken)
    buffer = io.BytesIO()
    with pytest.raises(RuntimeError, match="serializer failed"):
        with writer_cls(buffer) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([["a"], [1]])
    assert buffer.getvalue() == b""


@pytest.mark.parametrize("writer_cls,ext,method", [
    pytest.param(XlsxWriter, "xlsx", "_create_styles_xml", id="xlsx"),
    pytest.param(XlsbWriter, "xlsb", "_build_styles_bin", id="xlsb"),
])
@pytest.mark.parametrize("kind", ["path", "bytesio"])
def test_error_while_saving_publishes_nothing(tmp_path, monkeypatch, writer_cls, ext, method, kind):
    def broken(*args, **kwargs):
        raise OSError("save failed")

    monkeypatch.setattr(writer_cls, method, broken)
    target = _target(kind, tmp_path, ext)
    with pytest.raises(OSError, match="save failed"):
        with writer_cls(target) as writer:
            _two_sheets(writer)
    assert _published(target) is None
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("kind", ["path", "bytesio"])
def test_error_while_closing_the_archive_publishes_nothing(tmp_path, monkeypatch, writer_cls, ext, kind):
    def broken(self, *args, **kwargs):
        raise OSError("central directory write failed")

    monkeypatch.setattr(zipfile.ZipFile, "_write_end_record", broken)
    target = _target(kind, tmp_path, ext)
    with pytest.raises(OSError, match="central directory"):
        with writer_cls(target) as writer:
            _two_sheets(writer)
    monkeypatch.undo()
    assert _published(target) is None
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_target_stream_failure_is_reported(tmp_path, writer_cls, ext):
    # Copying into the caller's stream is the one step that cannot be undone:
    # the error reaches the caller, who must discard the partial stream.
    stream = FailingStream(limit=100)
    with pytest.raises(OSError, match="no space left"):
        with writer_cls(stream) as writer:
            _two_sheets(writer)
    assert len(stream.data) <= 100


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_exception_after_explicit_save_publishes_nothing(tmp_path, writer_cls, ext):
    path = tmp_path / f"out.{ext}"
    with pytest.raises(KeyError):
        with writer_cls(path) as writer:
            _two_sheets(writer)
            writer.save()
            raise KeyError("user code failed")
    assert not path.exists()


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_close_without_save_publishes_nothing(tmp_path, writer_cls, ext):
    buffer = io.BytesIO()
    writer = writer_cls(buffer).__enter__()
    _two_sheets(writer)
    writer.close()
    assert buffer.getvalue() == b""


# --------------------------------------------------------- writer states


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_writes_and_save_after_a_failure_are_refused(tmp_path, writer_cls, ext):
    path = tmp_path / f"out.{ext}"
    with pytest.raises(RuntimeError, match="incomplete and was not saved"):
        with writer_cls(path) as writer:
            writer.add_sheet("Sheet1")
            with pytest.raises(RuntimeError, match="source failed"):
                writer.write_sheet(_exploding_rows())
            with pytest.raises(RuntimeError, match="incomplete and was not saved"):
                writer.write_sheet([["more"]])
            writer.save()
    assert not path.exists()
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_writes_after_save_are_refused(tmp_path, writer_cls, ext):
    with writer_cls(tmp_path / f"out.{ext}") as writer:
        _two_sheets(writer)
        writer.save()
        with pytest.raises(RuntimeError, match="already saved"):
            writer.save()
        with pytest.raises(RuntimeError, match="already saved"):
            writer.write_sheet([["late"]])
    assert (tmp_path / f"out.{ext}").exists()


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_writer_must_be_open(tmp_path, writer_cls, ext):
    writer = writer_cls(tmp_path / f"out.{ext}")
    with pytest.raises(RuntimeError, match="not open"):
        writer.write_sheet([["x"]])
    with pytest.raises(RuntimeError, match="not open"):
        writer.save()
    writer.close()  # closing a writer that never opened is a no-op
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_unusable_paths_fail_before_writing(tmp_path, writer_cls, ext):
    with pytest.raises(IsADirectoryError):
        writer_cls(tmp_path).__enter__()
    with pytest.raises(FileNotFoundError) as excinfo:
        writer_cls(tmp_path / "missing" / f"out.{ext}").__enter__()
    assert "missing" in str(excinfo.value.filename)
    assert _leftovers(tmp_path) == []


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
def test_abandoned_writer_leaves_no_staging_file(tmp_path, writer_cls, ext):
    import gc

    writer = writer_cls(tmp_path / f"out.{ext}").__enter__()
    writer.write_sheet([["x"]])
    assert len(_leftovers(tmp_path)) == 1  # the staging file exists while writing
    del writer
    gc.collect()
    assert list(tmp_path.iterdir()) == []


def test_non_writable_target_is_rejected():
    with pytest.raises(TypeError, match="writable binary file object"):
        XlsxWriter(object()).__enter__()


# ------------------------------------------------------ temporary files


@pytest.mark.parametrize("writer_cls,ext", WRITERS)
@pytest.mark.parametrize("fail", [False, True], ids=["success", "failure"])
def test_stream_staging_leaves_no_temporary_files(tmp_path, monkeypatch, writer_cls, ext, fail):
    from xlspy import _output

    spool_dir = tmp_path / "spool"
    spool_dir.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(spool_dir))
    # A one-byte staging spool rolls the archive over to disk at once.
    # XlsxWriter stages with its own spool_max_bytes; XlsbWriter uses the default.
    monkeypatch.setattr(_output, "DEFAULT_STAGING_BYTES", 1)
    options = {"spool_max_bytes": 1} if writer_cls is XlsxWriter else {}
    rolled = []
    original_rollover = tempfile.SpooledTemporaryFile.rollover

    def recording_rollover(self):
        rolled.append(True)
        original_rollover(self)

    monkeypatch.setattr(tempfile.SpooledTemporaryFile, "rollover", recording_rollover)
    buffer = io.BytesIO()

    def source():
        for i in range(5000):
            yield [f"row {i}", i, i * 0.5]
        if fail:
            raise RuntimeError("source failed mid-export")

    if fail:
        with pytest.raises(RuntimeError, match="source failed"):
            with writer_cls(buffer, **options) as writer:
                writer.write_sheet(source())
        assert buffer.getvalue() == b""
    else:
        with writer_cls(buffer, **options) as writer:
            writer.write_sheet(source())
        assert _rows_of(buffer.getvalue(), ext, tmp_path, "Sheet1")[4999] == ["row 4999", 4999, 2499.5]
    assert rolled
    assert list(spool_dir.iterdir()) == []
