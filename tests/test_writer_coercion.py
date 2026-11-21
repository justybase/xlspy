"""Writer input options: strings_to_* and nan_inf_to_errors."""

import io

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter, load_workbook

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


def write(stream, writer, rows, **options):
    with writer(stream, **options) as book:
        book.add_sheet("Data")
        book.write_sheet(rows)


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_nan_and_inf_become_errors(writer, suffix):
    stream = io.BytesIO()
    write(stream, writer, [[float("nan"), float("inf"), float("-inf")]], nan_inf_to_errors=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data")) == [["error:#NUM!", "error:#DIV/0!", "error:#DIV/0!"]]


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_finite_numbers_are_untouched(writer, suffix):
    stream = io.BytesIO()
    write(stream, writer, [[1.5, float("-2.25"), 3]], nan_inf_to_errors=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data")) == [[1.5, -2.25, 3]]


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_nan_still_raises_by_default(writer, suffix):
    with pytest.raises(ValueError, match="finite"):
        write(io.BytesIO(), writer, [[float("nan")]])


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_strings_to_formulas(writer, suffix):
    stream = io.BytesIO()
    write(stream, writer, [["=1+1", "text"]], strings_to_formulas=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data", data_only=False)) == [["=1+1", "text"]]


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_formula_strings_stay_text_by_default(writer, suffix):
    stream = io.BytesIO()
    write(stream, writer, [["=1+1"]])
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data", data_only=False)) == [["=1+1"]]


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_strings_to_numbers(writer, suffix):
    stream = io.BytesIO()
    write(stream, writer, [["42", "-1.5", "1e3", "nan", "0x10", "1_000"]], strings_to_numbers=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data")) == [[42, -1.5, 1000.0, "nan", "0x10", "1_000"]]


@pytest.mark.parametrize("writer, suffix", WRITERS)
def test_strings_to_urls(writer, suffix):
    stream = io.BytesIO()
    write(
        stream,
        writer,
        [["https://example.com", "mailto:a@example.com", "not a url"]],
        strings_to_urls=True,
    )
    stream.seek(0)
    book = load_workbook(stream)
    sheet = book["Data"]
    assert sheet["A1"].hyperlink.target == "https://example.com"
    assert sheet["B1"].hyperlink.target == "mailto:a@example.com"
    assert sheet["C1"].hyperlink is None
    book.close()


def test_formatted_cell_keeps_its_format_when_coerced():
    stream = io.BytesIO()
    write(stream, XlsxWriter, [[(float("nan"), "0.00")]], nan_inf_to_errors=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data")) == [["error:#NUM!"]]


def test_options_must_be_bools():
    with pytest.raises(TypeError, match="strings_to_formulas"):
        XlsxWriter(io.BytesIO(), strings_to_formulas=1)
    with pytest.raises(TypeError, match="nan_inf_to_errors"):
        XlsbWriter(io.BytesIO(), nan_inf_to_errors="yes")


def test_finite_check_is_not_fooled_by_numpy_float32():
    np = pytest.importorskip("numpy")
    stream = io.BytesIO()
    write(stream, XlsxWriter, [[np.float32("nan")]], nan_inf_to_errors=True)
    stream.seek(0)
    with ExcelReader(stream) as reader:
        assert list(reader.get_rows("Data")) == [["error:#NUM!"]]
