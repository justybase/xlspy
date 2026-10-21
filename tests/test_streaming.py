"""Streaming writes: one-pass input, bounded memory, spooling, and failure handling."""

import io
import re
import struct
import tempfile
import tracemalloc
import zipfile

import pytest

from xlspy import Chart, ChartSeries, ExcelReader, Hyperlink, Formula, XlsbWriter, XlsxWriter

ROWS = 120_000
COLS = 5


def _rows(count=ROWS, cols=COLS):
    yield [f"H{j}" for j in range(cols)]
    for i in range(1, count):
        yield [f"row {i} text {i}", i, i * 0.5, f"more {i}", bool(i % 2)][:cols]


class OnePass:
    """A generator-like iterable that refuses a second pass."""

    def __init__(self, rows):
        self._iterator = iter(rows)
        self.passes = 0

    def __iter__(self):
        self.passes += 1
        if self.passes > 1:
            raise AssertionError("the writer iterated the rows twice")
        return self._iterator


def _sheet_bytes(path, member):
    with zipfile.ZipFile(path) as archive:
        return archive.read(member)


# -------------------------------------------------------- correctness


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_one_pass_generator_is_written_in_full(tmp_path, writer_cls, ext):
    path = tmp_path / f"stream.{ext}"
    source = OnePass(_rows(5000))
    with writer_cls(path) as writer:
        writer.add_sheet("Stream")
        writer.write_sheet(source)
    assert source.passes == 1
    with ExcelReader(str(path)) as reader:
        rows = list(reader.get_rows("Stream"))
    assert len(rows) == 5000
    assert rows[1] == ["row 1 text 1", 1, 0.5, "more 1", True]


def test_spooled_body_matches_in_memory_output(tmp_path):
    """Rolling the spool over to disk must not change a single byte of the sheet."""
    data = [["Name", "Note"]] + [[f"name {i}", f"note {i % 7}"] for i in range(3000)]
    in_memory = tmp_path / "memory.xlsx"
    on_disk = tmp_path / "disk.xlsx"
    with XlsxWriter(in_memory) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)
    with XlsxWriter(on_disk, spool_max_bytes=1) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(data)
    assert _sheet_bytes(in_memory, "xl/worksheets/sheet1.xml") == _sheet_bytes(on_disk, "xl/worksheets/sheet1.xml")


def test_spooled_styled_sheet_matches_in_memory_output(tmp_path):
    from xlspy import Style

    data = [["A", "B"]] + [[i, f"x{i}"] for i in range(2000)]
    kwargs = dict(
        header_style=Style(bold=True, fill="#183C52", font_color="#FFFFFF"),
        styles={"A2:A2001": Style(number_format="#,##0")},
        auto_filter=True,
        freeze_panes="A2",
    )
    paths = []
    for spool in (None, 1):
        path = tmp_path / f"styled_{spool}.xlsx"
        paths.append(path)
        options = {} if spool is None else {"spool_max_bytes": spool}
        with XlsxWriter(path, **options) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(data, **kwargs)
    assert _sheet_bytes(paths[0], "xl/worksheets/sheet1.xml") == _sheet_bytes(paths[1], "xl/worksheets/sheet1.xml")


def test_dimension_and_autofilter_reflect_the_streamed_rows(tmp_path):
    path = tmp_path / "dims.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(OnePass(_rows(777, cols=4)))
    xml = _sheet_bytes(path, "xl/worksheets/sheet1.xml").decode()
    assert '<dimension ref="A1:D777"/>' in xml
    assert '<autoFilter ref="A1:D777"/>' in xml


def test_empty_input_and_sheet_without_data(tmp_path):
    path = tmp_path / "empty.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Filled")
        writer.write_sheet([])
        writer.add_sheet("Never written")
    names = zipfile.ZipFile(path).namelist()
    assert "xl/worksheets/sheet2.xml" in names
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Filled")) == []


def test_chart_caches_come_from_streamed_rows(tmp_path):
    rows = [["Month", "Sales", "Other"]] + [[f"M{i}", i * 10, i] for i in range(1, 4001)]
    path = tmp_path / "chart_stream.xlsb"
    source = OnePass(rows)
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            source,
            charts=[Chart("line", [ChartSeries("B2:B4001", categories="A2:A4001", title_cell="B1")],
                          anchor="F2")],
        )
    assert source.passes == 1
    with zipfile.ZipFile(path) as archive:
        chart = archive.read("xl/charts/chart1.xml").decode()
    # Numeric caches hold the streamed values; text caches (titles, categories) are not digits.
    values = [int(v) for v in re.findall(r"<pt idx=\"\d+\"><v>(\d+)</v></pt>", chart)]
    assert values[:3] == [10, 20, 30]
    assert len(values) == 4000
    assert "<f>'Data'!$B$2:$B$4001</f>" in chart
    assert "<v>M3999</v>" in chart  # category text came from the stream too


# ----------------------------------------------------- bounded memory


def test_xlsx_inline_streaming_peak_memory_is_bounded(tmp_path):
    path = tmp_path / "bounded.xlsx"
    tracemalloc.start()
    try:
        with XlsxWriter(path, useSharedStrings=False) as writer:
            writer.add_sheet("Stream")
            writer.write_sheet(OnePass(_rows()))
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    # 120k rows x 5 columns: the shared-string table alone grows to ~40 MB
    # in shared mode; streaming with inline strings stays near the spool limit.
    assert peak < 25 * 2**20, f"peak {peak / 2**20:.1f} MB"


def test_xlsb_inline_streaming_peak_memory_is_bounded(tmp_path):
    path = tmp_path / "bounded.xlsb"
    tracemalloc.start()
    try:
        with XlsbWriter(path, useSharedStrings=False) as writer:
            writer.add_sheet("Stream")
            writer.write_sheet(OnePass(_rows()))
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert peak < 8 * 2**20, f"peak {peak / 2**20:.1f} MB"
    with ExcelReader(str(path)) as reader:
        assert len(list(reader.get_rows("Stream"))) == ROWS


# --------------------------------------------------- failure handling


def _exploding_rows():
    yield ["a", "b"]
    yield [1, 2]
    raise RuntimeError("source failed mid-export")


def test_generator_error_discards_xlsx_output(tmp_path):
    path = tmp_path / "failed.xlsx"
    with pytest.raises(RuntimeError, match="source failed"):
        with XlsxWriter(path) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(_exploding_rows())
    assert not path.exists()


def test_generator_error_discards_xlsb_output(tmp_path):
    path = tmp_path / "failed.xlsb"
    with pytest.raises(RuntimeError, match="source failed"):
        with XlsbWriter(path) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(_exploding_rows(), charts=[Chart("bar", [ChartSeries("A1:A2")])])
    assert not path.exists()


def test_exception_in_with_block_discards_output(tmp_path):
    path = tmp_path / "interrupted.xlsx"
    with pytest.raises(KeyError):
        with XlsxWriter(path) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet([["x"]])
            raise KeyError("user code failed")
    assert not path.exists()


def test_failed_sheet_blocks_save_for_file_objects():
    buffer = io.BytesIO()
    with pytest.raises(RuntimeError, match="source failed"):
        with XlsxWriter(buffer) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(_exploding_rows())


def test_failed_spooled_write_leaves_no_temporary_files(tmp_path, monkeypatch):
    # spool_max_bytes=1 forces the encoded rows onto disk before the failure.
    spool_dir = tmp_path / "spool"
    spool_dir.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(spool_dir))

    def rows():
        for i in range(30_000):
            yield [f"row {i} text", i, i * 0.5]
        raise RuntimeError("source failed mid-export")

    path = tmp_path / "failed.xlsx"
    with pytest.raises(RuntimeError, match="source failed"):
        with XlsxWriter(path, spool_max_bytes=1) as writer:
            writer.add_sheet("Sheet1")
            writer.write_sheet(rows())
    assert not path.exists()
    assert list(spool_dir.iterdir()) == []


def test_failed_write_refuses_save_when_caught_inside_block(tmp_path):
    path = tmp_path / "caught.xlsx"
    with pytest.raises(RuntimeError, match="incomplete and was not saved"):
        with XlsxWriter(path) as writer:
            writer.add_sheet("Sheet1")
            try:
                writer.write_sheet(_exploding_rows())
            except RuntimeError:
                pass
    assert not path.exists()


def test_argument_errors_do_not_abort_the_workbook(tmp_path):
    path = tmp_path / "args.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Sheet1")
        with pytest.raises(ValueError, match="freeze_panes must be a single cell"):
            writer.write_sheet([["x"]], freeze_panes="A1:B2")
        writer.write_sheet([["recovered"]])
    assert path.exists()
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Sheet1")) == [["recovered"]]


# ------------------------------------------------ XLSB inline strings


def test_xlsb_inline_and_shared_strings_round_trip(tmp_path):
    data = [["Name", "Note"], ["Zażółć", "ünï"], ["", "x" * 300], ["Sum", None]]
    for shared in (True, False):
        path = tmp_path / f"strings_{shared}.xlsb"
        with XlsbWriter(path, useSharedStrings=shared) as writer:
            writer.add_sheet("Data")
            writer.write_sheet(data)
        with ExcelReader(str(path)) as reader:
            assert list(reader.get_rows("Data")) == data


def test_xlsb_inline_text_limit_is_enforced(tmp_path):
    path = tmp_path / "limit.xlsb"
    with pytest.raises(ValueError, match="32767"):
        with XlsbWriter(path, useSharedStrings=False) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([["x" * 32768]])
    assert not path.exists()


# -------------------------------------------- XLSB refuses XLSX-only cells


def test_xlsb_rejects_hyperlink_cells(tmp_path):
    for index, cell in enumerate([Hyperlink("https://example.com")]):
        path = tmp_path / f"rich_{index}.xlsb"
        with pytest.raises(TypeError, match="does not support Hyperlink"):
            with XlsbWriter(path) as writer:
                writer.add_sheet("Data")
                writer.write_sheet([["x", cell]])
        assert not path.exists()


@pytest.mark.parametrize("chart", [
    Chart("barh", [ChartSeries("B2:B3")]),
    Chart("scatter", [ChartSeries("B2:B3", x_values="A2:A3")]),
    Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], show_data_labels=True),
    Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")], title_size=14),
])
def test_xlsb_rejects_extended_chart_options(tmp_path, chart):
    with XlsbWriter(tmp_path / "chart.xlsb") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="XlsbWriter supports only the original chart options"):
            writer.write_sheet([["a", 1], ["b", 2], ["c", 3]], charts=[chart])


def test_xlsb_original_chart_forms_still_write(tmp_path):
    path = tmp_path / "original.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Q", "A"], ["Q1", 10], ["Q2", 20]],
            charts=[Chart("bar", [ChartSeries("B2:B3", categories="A2:A3", title="A")], anchor="D2"),
                    Chart("pie", [ChartSeries("B2:B3", title="A")], anchor="D20")],
        )
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data"))[1] == ["Q1", 10]


# -------------------------------- XLSB inline text: UTF-16 units and parity


def _biff_records(data):
    """Yield (record_type, payload) pairs from a BIFF12 part."""
    pos = 0
    while pos < len(data):
        rec_type, pos = _biff_varint(data, pos)
        size, pos = _biff_varint(data, pos)
        yield rec_type, data[pos:pos + size]
        pos += size


def _biff_varint(data, pos):
    value = shift = 0
    while True:
        byte = data[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        if not byte & 0x80:
            return value, pos
        shift += 7


def test_xlsb_inline_counts_supplementary_characters_as_two_units(tmp_path):
    # BrtCellSt stores cch in UTF-16 code units; each astral character is two.
    # xlspy's own reader ignores cch, so the raw record is checked directly:
    # a count in code points would still round-trip here but corrupt the cell in Excel.
    text = "\U0001D11E" + "x" * 3
    path = tmp_path / "astral.xlsb"
    with XlsbWriter(path, useSharedStrings=False) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[text, "after"]])
    with zipfile.ZipFile(path) as archive:
        sheet = archive.read("xl/worksheets/sheet1.bin")
    cells = [payload for rec_type, payload in _biff_records(sheet) if rec_type == 0x06]
    assert len(cells) == 2
    _col, _style, cch = struct.unpack_from("<III", cells[0], 0)
    units = len(text.encode("utf-16-le")) // 2
    assert cch == units == 5
    assert len(cells[0]) == 12 + 2 * cch
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [[text, "after"]]


def test_xlsb_inline_text_limit_counts_utf16_units(tmp_path):
    # 16383 astral characters are 32766 units (allowed); 16384 are 32768 (refused).
    ok = "\U0001D11E" * 16383
    path = tmp_path / "limit_ok.xlsb"
    with XlsbWriter(path, useSharedStrings=False) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[ok]])
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [[ok]]
    with pytest.raises(ValueError, match="32767"):
        with XlsbWriter(tmp_path / "limit_bad.xlsb", useSharedStrings=False) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([["\U0001D11E" * 16384]])


def test_xlsb_inline_lone_surrogate_is_refused(tmp_path):
    path = tmp_path / "surrogate.xlsb"
    with pytest.raises(UnicodeEncodeError):
        with XlsbWriter(path, useSharedStrings=False) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([["bad\ud800"]])
    assert not path.exists()


def _accel_available():
    from xlspy import _accel

    return _accel.HAVE_C_EXT


_PARITY_SCRIPT = r'''
import datetime
import decimal
import sys

from xlspy import XlsbWriter
from xlspy import _accel


class Odd:
    def __str__(self):
        return "odd \u00e9 \U0001D11E"


rows = [
    ["text", "\U0001D11E astral", "\u00e9t\u00e9", ""],
    [1, -1, 2 ** 40, 2 ** 70, 1.5, -0.0, float("nan"), 536870911, -536870912],
    [True, False, None, decimal.Decimal("1.25"), datetime.date(2024, 1, 2),
     datetime.datetime(2024, 1, 2, 3, 4, 5)],
    [Odd(), "x" * 300, 7],
]
print("c_row=" + str(_accel.encode_xlsb_row_inline(["a"], 0) is not None))
with XlsbWriter(sys.argv[1], useSharedStrings=False) as writer:
    writer.add_sheet("Data")
    writer.write_sheet(rows)
'''


@pytest.mark.skipif(not _accel_available(), reason="C extension not built")
def test_xlsb_inline_c_and_python_encoders_write_identical_bytes(tmp_path):
    import os
    import subprocess
    import sys

    import xlspy

    src_dir = os.path.dirname(os.path.dirname(xlspy.__file__))
    script = tmp_path / "write_inline.py"
    script.write_text(_PARITY_SCRIPT, encoding="utf-8")
    members = {}
    for label, disable in (("c", "0"), ("python", "1")):
        out = tmp_path / f"{label}.xlsb"
        env = dict(os.environ, PYTHONPATH=src_dir, XLSPY_DISABLE_C_EXT=disable)
        proc = subprocess.run(
            [sys.executable, str(script), str(out)],
            env=env, capture_output=True, text=True, encoding="utf-8", check=False,
        )
        assert proc.returncode == 0, proc.stderr
        expected = "c_row=True" if disable == "0" else "c_row=False"
        assert expected in proc.stdout
        with zipfile.ZipFile(out) as archive:
            members[label] = archive.read("xl/worksheets/sheet1.bin")
    assert members["c"] == members["python"]
