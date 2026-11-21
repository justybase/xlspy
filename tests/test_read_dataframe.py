"""read_dataframe: streaming a worksheet into pandas or polars."""

import datetime

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter, read_dataframe

pandas = pytest.importorskip("pandas")
polars = pytest.importorskip("polars")

ROWS = [
    ["name", "age", "when"],
    ["Alice", 25, datetime.date(2026, 1, 2)],
    ["Bob", 30, datetime.date(2026, 2, 3)],
]


def make_book(path, rows=ROWS, writer=XlsxWriter):
    with writer(path) as book:
        book.add_sheet("Data")
        book.write_sheet(rows)
        book.add_sheet("Other")
        book.write_sheet([["x"], [1]])
    return path


@pytest.mark.parametrize("engine", ["pandas", "polars"])
def test_basic_frame(tmp_path, engine):
    path = make_book(tmp_path / "data.xlsx")
    frame = read_dataframe(path, engine=engine)
    assert frame.shape == (2, 3)
    assert list(frame.columns) == ["name", "age", "when"]
    assert frame["name"][0] == "Alice"
    assert float(frame["age"][1]) == 30.0


@pytest.mark.parametrize("engine", ["pandas", "polars"])
def test_xlsb(tmp_path, engine):
    path = make_book(tmp_path / "data.xlsb", writer=XlsbWriter)
    frame = read_dataframe(path, engine=engine)
    assert frame.shape == (2, 3)


def test_sheet_name_and_index(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    assert read_dataframe(path, "Other", engine="pandas").shape == (1, 1)
    assert read_dataframe(path, 1, engine="pandas").shape == (1, 1)
    assert set(read_dataframe(path, None, engine="pandas")) == {"Data", "Other"}


def test_header_false_and_supplied(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    frame = read_dataframe(path, header=False, engine="pandas")
    assert list(frame.columns) == ["column_0", "column_1", "column_2"]
    assert frame.shape == (3, 3)

    frame = read_dataframe(path, header=["a", "b", "c"], engine="pandas")
    assert list(frame.columns) == ["a", "b", "c"]
    assert frame.shape == (3, 3)


def test_usecols_by_name_and_index(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    assert list(read_dataframe(path, usecols=["age"], engine="pandas").columns) == ["age"]
    assert list(read_dataframe(path, usecols=[2, 0], engine="pandas").columns) == ["when", "name"]
    with pytest.raises(ValueError, match="not found"):
        read_dataframe(path, usecols=["missing"], engine="pandas")
    with pytest.raises(ValueError, match="out of range"):
        read_dataframe(path, usecols=[9], engine="pandas")


def test_nrows_and_dtype(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    frame = read_dataframe(path, nrows=1, dtype={"age": "int64"}, engine="pandas")
    assert frame.shape == (1, 3)
    assert str(frame["age"].dtype) == "int64"


def test_duplicate_headers_are_suffixed(tmp_path):
    path = make_book(
        tmp_path / "dupes.xlsx",
        rows=[["a", "a", "a"], [1, 2, 3]],
    )
    assert list(read_dataframe(path, engine="pandas").columns) == ["a", "a.1", "a.2"]


def test_ragged_rows_are_padded(tmp_path):
    path = make_book(
        tmp_path / "ragged.xlsx",
        rows=[["a", "b"], [1], [2, 3, 4]],
    )
    frame = read_dataframe(path, engine="pandas")
    assert list(frame.columns) == ["a", "b", "column_2"]
    first = frame.iloc[0].tolist()
    assert first[0] == 1 and pandas.isna(first[1]) and pandas.isna(first[2])
    assert frame.iloc[1].tolist() == [2, 3, 4]


def test_empty_sheet(tmp_path):
    path = make_book(tmp_path / "empty.xlsx", rows=[])
    assert read_dataframe(path, engine="pandas").shape == (0, 0)
    assert read_dataframe(path, engine="polars").shape == (0, 0)


def test_reader_methods(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    with ExcelReader(path) as reader:
        assert reader.to_pandas("Data").shape == (2, 3)
        assert reader.to_polars("Data").shape == (2, 3)


def test_engine_auto_uses_an_installed_engine(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    assert read_dataframe(path).shape == (2, 3)


def test_argument_checks(tmp_path):
    path = make_book(tmp_path / "data.xlsx")
    with pytest.raises(ValueError, match="engine"):
        read_dataframe(path, engine="excel")
    with pytest.raises(ValueError, match="nrows"):
        read_dataframe(path, nrows=-1)
    with pytest.raises(TypeError, match="sheet_name"):
        read_dataframe(path, sheet_name=True)
    with pytest.raises(KeyError, match="missing"):
        read_dataframe(path, "missing")
    with pytest.raises(IndexError, match="out of range"):
        read_dataframe(path, 99)
