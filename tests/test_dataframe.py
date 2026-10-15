"""Tests for the optional pandas / polars DataFrame integration.

pandas and polars are optional dependencies: tests that need them are skipped
when the library is not installed.
"""
import datetime
import os
import subprocess
import sys
import tempfile
from decimal import Decimal

import pytest

from xlspy import (  # noqa: E402
    ExcelReader,
    F,
    XlsbWriter,
    XlsxWriter,
    dataframe_to_rows,
    to_xlsb,
    to_xlsx,
)
from conftest import PROJECT_TMP  # noqa: E402

EXAMPLES = [
    "dataframe_to_xlsx.py",
    "dataframe_to_xlsb.py",
    "dataframe_writer_methods.py",
    "dataframe_formats_and_index.py",
    "dataframe_roundtrip.py",
    "polars_to_xlsx_xlsb.py",
]


def _pd():
    return pytest.importorskip("pandas")


def _pl():
    return pytest.importorskip("polars")


def _read(path, sheet="Sheet1"):
    with ExcelReader(path) as reader:
        return reader.read_all(sheet)


@pytest.fixture()
def tmp_dir():
    os.makedirs(PROJECT_TMP, exist_ok=True)
    directory = tempfile.mkdtemp(prefix="pyxlsb_df_", dir=PROJECT_TMP)
    yield directory
    import shutil

    shutil.rmtree(directory, ignore_errors=True)


@pytest.fixture()
def sample_df():
    pd = _pd()
    return pd.DataFrame(
        {
            "name": ["Alice", "Bob", None],
            "age": [25, 30, 40],
            "score": [1.5, float("nan"), 3.0],
            "active": [True, False, True],
            "joined": [pd.Timestamp("2024-01-15 10:30"), pd.NaT, pd.Timestamp("2024-02-01")],
            "amount": [Decimal("3.14"), Decimal("2.50"), Decimal("10.00")],
        }
    )


# ── pandas: shortcuts ───────────────────────────────────────────────────
@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_shortcuts_roundtrip(tmp_dir, sample_df, ext, shortcut):
    path = os.path.join(tmp_dir, f"shortcut.{ext}")
    returned = shortcut(sample_df, path, formats={"joined": F.DATE_ISO})
    assert returned == path
    assert os.path.exists(path)

    rows = _read(path)
    assert rows[0] == ["name", "age", "score", "active", "joined", "amount"]
    assert rows[1][0] == "Alice"
    assert rows[1][1] == 25
    assert rows[1][2] == 1.5
    assert rows[1][3] is True
    assert rows[1][4] == datetime.datetime(2024, 1, 15, 10, 30)
    # Excel stores numbers as doubles, so Decimal comes back as float.
    assert rows[1][5] == 3.14


@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_missing_values_become_empty_cells(tmp_dir, sample_df, ext, shortcut):
    path = os.path.join(tmp_dir, f"missing.{ext}")
    shortcut(sample_df, path, formats={"joined": F.DATE_ISO})

    rows = _read(path)
    # row 2: name None, score NaN, joined NaT -> all empty
    assert rows[2][0] == "Bob"
    assert rows[2][2] is None
    assert rows[2][4] is None
    # row 3: name None
    assert rows[3][0] is None


def test_na_rep_writes_placeholder(tmp_dir):
    pd = _pd()
    path = os.path.join(tmp_dir, "narep.xlsx")
    to_xlsx(pd.DataFrame({"a": [1, None, 3]}), path, na_rep="N/A")
    rows = _read(path)
    assert rows[1] == [1]
    assert rows[2] == ["N/A"]
    assert rows[3] == [3]


def test_include_index(tmp_dir):
    pd = _pd()
    df = pd.DataFrame({"a": [1, 2]}, index=pd.Index([10, 20], name="id"))
    path = os.path.join(tmp_dir, "index.xlsx")
    to_xlsx(df, path, include_index=True)
    rows = _read(path)
    assert rows[0] == ["id", "a"]
    assert rows[1] == [10, 1]
    assert rows[2] == [20, 2]


def test_multiindex_columns_and_index(tmp_dir):
    pd = _pd()
    columns = pd.MultiIndex.from_arrays([["A", "A", "B"], ["x", "y", "z"]])
    index = pd.MultiIndex.from_tuples([("g", 1), ("h", 2)], names=["grp", "n"])
    df = pd.DataFrame([[1, 2, 3], [4, 5, 6]], columns=columns, index=index)

    path = os.path.join(tmp_dir, "multiindex.xlsx")
    to_xlsx(df, path, include_index=True)
    rows = _read(path)
    assert rows[0] == ["grp", "n", "A x", "A y", "B z"]
    assert rows[1] == ["g", 1, 1, 2, 3]
    assert rows[2] == ["h", 2, 4, 5, 6]


def test_numpy_scalar_types(tmp_dir):
    pd = _pd()
    np = pytest.importorskip("numpy")
    df = pd.DataFrame(
        {
            "i": np.array([1, 2], dtype=np.int64),
            "f": np.array([1.5, 2.5], dtype=np.float32),
            "b": np.array([True, False], dtype=np.bool_),
        }
    )
    path = os.path.join(tmp_dir, "numpy.xlsb")
    to_xlsb(df, path)
    rows = _read(path)
    assert [type(v).__name__ for v in rows[1]] == ["int", "float", "bool"]
    assert rows[1] == [1, 1.5, True]
    assert rows[2] == [2, 2.5, False]


@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_numpy_float_nan_in_object_column_is_empty(tmp_dir, ext, shortcut):
    # np.float16 / np.float32 do not subclass float and used to be written
    # as the literal string "NaN" (xlsx) or a NaN double (xlsb).
    pd = _pd()
    np = pytest.importorskip("numpy")
    df = pd.DataFrame(
        {
            "v": pd.Series([np.float32(1.0), np.float32("nan"), np.float16("nan")], dtype=object),
            "tag": ["a", "b", "c"],
        }
    )
    path = os.path.join(tmp_dir, f"np_nan.{ext}")
    shortcut(df, path)
    rows = _read(path)
    assert rows[1][0] == 1.0
    assert rows[2][0] is None
    assert rows[3][0] is None


def test_numpy_timedelta_becomes_day_fraction(tmp_dir):
    # np.timedelta64 used to be passed through as datetime.timedelta and
    # written as text ("12:00:00") instead of a day fraction.
    pd = _pd()
    np = pytest.importorskip("numpy")
    df = pd.DataFrame(
        {"dur": pd.Series([np.timedelta64(12, "h")], dtype=object), "tag": ["a"]}
    )
    path = os.path.join(tmp_dir, "np_dur.xlsx")
    to_xlsx(df, path)
    rows = _read(path)
    assert rows[1][0] == 0.5


@pytest.mark.filterwarnings("ignore::DeprecationWarning")  # pandas internals warn on NaT->object
@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_numpy_nat_with_format_is_empty(tmp_dir, ext, shortcut):
    # NaT used to be written as the literal string "None" when a format was set.
    pd = _pd()
    np = pytest.importorskip("numpy")
    df = pd.DataFrame(
        {
            "when": pd.Series([np.datetime64("2024-01-15"), np.datetime64("NaT")], dtype=object),
            "dur": pd.Series([np.timedelta64(1, "h"), np.timedelta64("NaT")], dtype=object),
            "tag": ["a", "b"],
        }
    )
    path = os.path.join(tmp_dir, f"np_nat.{ext}")
    shortcut(df, path, formats={"when": F.DATE_ISO, "dur": F.TIME_HH_MM})
    rows = _read(path)
    assert rows[1][0] is not None
    assert rows[2][0] is None
    assert rows[2][1] is None


def test_time_column_becomes_day_fraction(tmp_dir):
    pd = _pd()
    df = pd.DataFrame({"t": [datetime.time(12, 0), datetime.time(6, 0)]})
    path = os.path.join(tmp_dir, "time.xlsx")
    to_xlsx(df, path)
    rows = _read(path)
    assert rows[1][0] == 0.5
    assert rows[2][0] == 0.25


def test_formats_dict_applied(tmp_dir):
    pd = _pd()
    df = pd.DataFrame(
        {
            "amount": [100000.0],
            "when": [pd.Timestamp("2024-01-15")],
        }
    )
    path = os.path.join(tmp_dir, "formats.xlsx")
    to_xlsx(df, path, formats={"amount": F.TWO_DECIMALS, "when": F.DATE_ISO})

    import openpyxl

    worksheet = openpyxl.load_workbook(path).active
    assert worksheet["A2"].number_format == "#,##0.00"
    assert worksheet["B2"].number_format == "yyyy-mm-dd"


def test_xlsb_format_reads_back_as_datetime(tmp_dir):
    pd = _pd()
    df = pd.DataFrame({"when": [pd.Timestamp("2024-01-15 08:00")]})
    path = os.path.join(tmp_dir, "formats.xlsb")
    to_xlsb(df, path, formats={"when": F.DATETIME_SHORT})
    rows = _read(path)
    assert rows[1] == [datetime.datetime(2024, 1, 15, 8, 0)]


# ── pandas: writer methods ──────────────────────────────────────────────
@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_writer_methods_multiple_sheets(tmp_dir, writer_cls, ext):
    pd = _pd()
    path = os.path.join(tmp_dir, f"methods.{ext}")
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_dataframe(pd.DataFrame({"a": [1, 2], "b": ["x", "y"]}))
        writer.add_sheet("Raw")
        writer.write_sheet([["manual"], [42]])

    data = _read(path, "Data")
    assert data[0] == ["a", "b"]
    assert data[1] == [1, "x"]
    raw = _read(path, "Raw")
    assert raw[0] == ["manual"]
    assert raw[1] == [42]


def test_write_dataframe_creates_default_sheet(tmp_dir):
    pd = _pd()
    path = os.path.join(tmp_dir, "default.xlsx")
    with XlsxWriter(path) as writer:
        writer.write_dataframe(pd.DataFrame({"a": [1]}))
    assert _read(path, "Sheet1") == [["a"], [1]]


@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_empty_dataframe_shapes_roundtrip(tmp_dir, ext, shortcut):
    pd = _pd()
    cases = [
        (pd.DataFrame(), [[None]] if ext == "xlsb" else [[]]),
        (pd.DataFrame(columns=["a", "b"]), [["a", "b"]]),
        (pd.DataFrame({"a": pd.Series(dtype="int64")}), [["a"]]),
    ]
    for index, (frame, expected) in enumerate(cases):
        path = os.path.join(tmp_dir, f"empty_{index}.{ext}")
        shortcut(frame, path)
        assert _read(path) == expected


@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_dataframe_matches_write_sheet_rows(tmp_dir, ext, shortcut):
    pd = _pd()
    frame = pd.DataFrame(
        {
            "label": ["x", "y"],
            "when": [pd.Timestamp("1900-01-01"), pd.Timestamp("2024-02-03 12:30")],
            "value": [1.25, None],
        }
    )
    formats = {"when": F.DATE_ISO}
    rows = list(dataframe_to_rows(frame, formats=formats))
    shortcut_path = os.path.join(tmp_dir, f"shortcut_equiv.{ext}")
    rows_path = os.path.join(tmp_dir, f"rows_equiv.{ext}")
    shortcut(frame, shortcut_path, formats=formats)
    writer_cls = XlsbWriter if ext == "xlsb" else XlsxWriter
    with writer_cls(rows_path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(rows)
    assert _read(shortcut_path) == _read(rows_path)


@pytest.mark.parametrize(
    "ext,shortcut,writer_cls",
    [("xlsx", to_xlsx, XlsxWriter), ("xlsb", to_xlsb, XlsbWriter)],
)
def test_dataframe_shortcut_removes_partial_output_on_failure(
    tmp_dir, monkeypatch, ext, shortcut, writer_cls
):
    pd = _pd()

    def fail_during_conversion(*args, **kwargs):
        raise RuntimeError("intentional conversion failure")

    monkeypatch.setattr(writer_cls, "write_dataframe", fail_during_conversion)

    path = os.path.join(tmp_dir, f"atomic.{ext}")
    frame = pd.DataFrame({"value": ["data"]})
    with pytest.raises(RuntimeError, match="intentional conversion failure"):
        shortcut(frame, path)
    assert not os.path.exists(path)


def test_openpyxl_reads_dataframe_xlsx_output(tmp_dir):
    pd = _pd()
    pytest.importorskip("openpyxl")
    frame = pd.DataFrame(
        {
            "name": ["Ada", "Lin"],
            "count": [3, 7],
            "when": [pd.Timestamp("1900-01-01"), pd.Timestamp("2024-02-03")],
            "active": [True, False],
        }
    )
    path = os.path.join(tmp_dir, "pandas_oracle.xlsx")
    to_xlsx(frame, path, formats={"when": F.DATE_ISO})
    actual = pd.read_excel(path, engine="openpyxl")
    expected = frame.copy()
    expected["when"] = pd.to_datetime(expected["when"])
    pd.testing.assert_frame_equal(actual, expected, check_dtype=False)


# ── polars ──────────────────────────────────────────────────────────────
@pytest.mark.parametrize("ext,shortcut", [("xlsx", to_xlsx), ("xlsb", to_xlsb)])
def test_polars_roundtrip(tmp_dir, ext, shortcut):
    pl = _pl()
    df = pl.DataFrame(
        {
            "name": ["Alice", "Bob"],
            "age": [25, None],
            "score": [1.5, 2.5],
            "active": [True, False],
            "joined": [datetime.datetime(2024, 1, 15, 10, 30), None],
        }
    )
    path = os.path.join(tmp_dir, f"polars.{ext}")
    shortcut(df, path, formats={"joined": F.DATETIME_SHORT})

    rows = _read(path)
    assert rows[0] == ["name", "age", "score", "active", "joined"]
    assert rows[1] == ["Alice", 25, 1.5, True, datetime.datetime(2024, 1, 15, 10, 30)]
    assert rows[2][1] is None
    assert rows[2][4] is None


def test_polars_duration_becomes_day_fraction(tmp_dir):
    pl = _pl()
    df = pl.DataFrame({"dur": [datetime.timedelta(hours=12)]})
    path = os.path.join(tmp_dir, "duration.xlsx")
    to_xlsx(df, path)
    rows = _read(path)
    # 12 hours -> 0.5 of a day; apply F.TIME_HH_MM to display it as a clock time.
    assert rows[1][0] == 0.5


def test_polars_include_index_raises(tmp_dir):
    pl = _pl()
    path = os.path.join(tmp_dir, "polars_index.xlsx")
    with pytest.raises(ValueError, match="no index"):
        to_xlsx(pl.DataFrame({"a": [1]}), path, include_index=True)
    # a failed conversion must not leave a partial file behind
    assert not os.path.exists(path)


# ── generic behaviour ───────────────────────────────────────────────────
def test_non_dataframe_raises_type_error():
    with pytest.raises(TypeError, match="pandas or polars"):
        list(dataframe_to_rows([["a"], [1]]))


def test_import_xlspy_does_not_import_optional_dependencies():
    code = (
        "import sys; import xlspy; "
        "leaked = [m for m in ('pandas', 'polars', 'numpy') if m in sys.modules]; "
        "print(','.join(leaked))"
    )
    env = dict(os.environ)
    env["PYTHONPATH"] = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
    result = subprocess.run(
        [sys.executable, "-c", code], capture_output=True, text=True, env=env
    )
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "", f"xlspy imported optional deps: {result.stdout}"


# ── bundled examples stay runnable ──────────────────────────────────────
@pytest.mark.parametrize("script", EXAMPLES)
def test_example_script_runs(script, tmp_path):
    if script.startswith("polars"):
        _pl()
    else:
        _pd()

    script_path = os.path.join(os.path.dirname(__file__), "..", "examples", script)
    result = subprocess.run(
        [sys.executable, script_path, str(tmp_path)],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, f"{script} failed:\n{result.stderr}"
    assert any(tmp_path.iterdir()), f"{script} produced no output files"
