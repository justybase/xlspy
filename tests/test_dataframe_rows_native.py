"""The native conversion of a DataFrame row gives what the Python loop gives."""

import datetime as dt
from decimal import Decimal

import pytest

from xlspy import _accel
from xlspy import _dataframe_rows as frames

pytestmark = pytest.mark.skipif(_accel.convert_frame_row is None, reason="native extension not built")


def _wide_frame():
    np = pytest.importorskip("numpy")
    pd = pytest.importorskip("pandas")
    n = 40
    return pd.DataFrame({
        "int": np.arange(n, dtype="int64"),
        "nullable_int": pd.array([None if i % 4 == 0 else i for i in range(n)], dtype="Int64"),
        "float": [float("nan") if i % 5 == 0 else i / 3 for i in range(n)],
        "float32": np.array([np.nan if i % 6 == 0 else i / 3 for i in range(n)], dtype="float32"),
        "bool": [i % 2 == 0 for i in range(n)],
        "nullable_bool": pd.array([None if i % 3 == 0 else i % 2 == 0 for i in range(n)], dtype="boolean"),
        "text": [f"t{i % 7}" if i % 8 else None for i in range(n)],
        "moment": pd.date_range("2024-01-01", periods=n, freq="7min").where(np.arange(n) % 9 != 0),
        "span": pd.to_timedelta(np.arange(n), unit="h").where(np.arange(n) % 10 != 0),
        "category": pd.Categorical([["a", "b", None][i % 3] for i in range(n)]),
        "mixed": pd.Series(
            [
                [None, Decimal("1.5"), dt.date(2024, 1, 2), dt.time(3, 4), dt.timedelta(hours=5), np.int64(7),
                 np.float32("nan"), pd.Timestamp("2024-02-03 04:05:06"), pd.NaT, pd.NA, "x", 1.5, True, b"raw",
                 np.datetime64("NaT", "us"), np.datetime64("2024-01-02"), (1, 2), float("nan"), 2**70, "é\n"][i % 20]
                for i in range(n)
            ],
            dtype=object,
        ),
    })


@pytest.mark.parametrize("na_rep", [None, "", 0, "n/a"])
@pytest.mark.parametrize("formats", [None, {"int": "0000", "float": "0.00", "text": "@", "moment": "yyyy-mm-dd", "mixed": "General"}])
@pytest.mark.parametrize("include_index", [False, True])
def test_native_rows_equal_python_rows(monkeypatch, na_rep, formats, include_index):
    frame = _wide_frame()
    native = repr(list(frames.dataframe_to_rows(frame, include_index=include_index, na_rep=na_rep, formats=formats)))
    monkeypatch.setattr(frames._accel, "convert_frame_row", None)
    python = repr(list(frames.dataframe_to_rows(frame, include_index=include_index, na_rep=na_rep, formats=formats)))
    assert native == python


def test_multiindex_and_polars_rows_equal(monkeypatch):
    pd = pytest.importorskip("pandas")
    frame = _wide_frame()[["int", "float", "text"]]
    frame.index = pd.MultiIndex.from_product([["a", "b"], range(20)])
    native = repr(list(frames.dataframe_to_rows(frame, include_index=True)))
    monkeypatch.setattr(frames._accel, "convert_frame_row", None)
    assert native == repr(list(frames.dataframe_to_rows(frame, include_index=True)))


def test_polars_rows_equal(monkeypatch):
    pl = pytest.importorskip("polars")
    frame = pl.DataFrame({"i": [1, None, 3], "f": [1.5, float("nan"), None], "s": ["a", None, "c"],
                          "d": [dt.date(2024, 1, 2), None, dt.date(2024, 1, 4)], "t": [dt.time(1, 2), None, dt.time(3, 4)]})
    native = repr(list(frames.dataframe_to_rows(frame, na_rep="-", formats={"f": "0.0"})))
    monkeypatch.setattr(frames._accel, "convert_frame_row", None)
    assert native == repr(list(frames.dataframe_to_rows(frame, na_rep="-", formats={"f": "0.0"})))


def test_a_timestamp_with_nanoseconds_warns_as_before(monkeypatch):
    pd = pytest.importorskip("pandas")
    frame = pd.DataFrame({"t": [pd.Timestamp("2024-01-02 03:04:05.123456789")]})
    with pytest.warns(UserWarning):
        native = list(frames.dataframe_to_rows(frame))
    monkeypatch.setattr(frames._accel, "convert_frame_row", None)
    with pytest.warns(UserWarning):
        python = list(frames.dataframe_to_rows(frame))
    assert native == python
