"""Remote workbook paths through fsspec: read, write and update."""

import io

import pytest

fsspec = pytest.importorskip("fsspec")

from xlspy import ExcelReader, Workbook, XlsbWriter, XlsxWriter, load_workbook, read_dataframe
from xlspy._remote import is_remote_uri, uri_extension


def memory_uri(name):
    return f"memory://xlspy/{name}"


def read_bytes(uri):
    with fsspec.open(uri, "rb") as stream:
        return stream.read()


def write_book(uri, writer=XlsxWriter):
    with writer(uri) as book:
        book.add_sheet("Data")
        book.write_sheet([["name", "value"], ["a", 1], ["b", 2]])


def test_uri_detection():
    assert is_remote_uri("memory://book.xlsx")
    assert is_remote_uri("s3://bucket/book.xlsb")
    assert not is_remote_uri("report.xlsx")
    assert not is_remote_uri(r"C:\reports\book.xlsx")
    assert not is_remote_uri("C:/reports/book.xlsx")


def test_uri_extension_ignores_query_strings():
    assert uri_extension("s3://bucket/book.xlsb?versionId=1") == "xlsb"
    assert uri_extension("memory://bucket/folder/book") == "xlsx"


def test_xlsxwriter_writes_a_remote_package():
    uri = memory_uri("writer.xlsx")
    write_book(uri)
    assert read_bytes(uri).startswith(b"PK")
    with ExcelReader(uri) as reader:
        assert reader.get_sheet_names() == ["Data"]
        assert reader.read_all("Data")[1] == ["a", 1.0]


def test_xlsbwriter_writes_a_remote_package():
    uri = memory_uri("writer.xlsb")
    write_book(uri, XlsbWriter)
    with ExcelReader(uri) as reader:
        assert reader.read_all("Data")[2] == ["b", 2.0]


def test_workbook_model_saves_and_loads_a_remote_file():
    uri = memory_uri("model.xlsx")
    workbook = Workbook()
    workbook.active["A1"] = "remote"
    workbook.active["A2"] = 7
    workbook.save(uri)
    workbook.close()

    loaded = load_workbook(uri)
    try:
        assert loaded["Sheet"]["A1"].value == "remote"
        loaded["Sheet"]["B1"] = "edited"
        loaded.save(memory_uri("edited.xlsb"))
    finally:
        loaded.close()

    with ExcelReader(memory_uri("edited.xlsb")) as reader:
        assert reader.read_all("Sheet") == [["remote", "edited"], [7.0, None]]


def test_read_dataframe_accepts_a_uri():
    uri = memory_uri("frame.xlsx")
    write_book(uri)
    frame = read_dataframe(uri)
    assert frame.to_dict("records") == [{"name": "a", "value": 1.0}, {"name": "b", "value": 2.0}]


def test_pointing_at_a_missing_uri_raises():
    with pytest.raises(FileNotFoundError):
        load_workbook(memory_uri("missing.xlsx"))


def test_a_failed_write_leaves_the_remote_file_unchanged():
    uri = memory_uri("unchanged.xlsx")
    write_book(uri)
    before = read_bytes(uri)
    with pytest.raises(RuntimeError):
        with XlsxWriter(uri) as book:
            book.add_sheet("Data")
            book.write_sheet([["replacement"]])
            raise RuntimeError("stop")
    assert read_bytes(uri) == before


def test_file_uri_works_like_a_local_path(tmp_path):
    path = tmp_path / "file-uri.xlsx"
    uri = path.as_uri()
    write_book(uri)
    assert path.exists()
    with ExcelReader(uri) as reader:
        assert reader.read_all("Data")[0] == ["name", "value"]


def test_local_paths_and_streams_still_work(tmp_path):
    path = tmp_path / "local.xlsx"
    write_book(path)
    assert path.exists()
    stream = io.BytesIO()
    write_book(stream)
    assert stream.getvalue().startswith(b"PK")
