"""Regression tests for output that Excel would repair or that held wrong data.

Each test reproduces one defect that was confirmed by writing a file and
inspecting its parts: an XLSB sheet dimension written into the selection record,
a chart formula with an unescaped ampersand, a calcChain left behind after its
formulas were removed, and similar. The checks read the package XML directly, so
they do not need Microsoft Excel.
"""

import re
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import pytest

from xlspy import (
    Chart,
    ChartSeries,
    ExcelTable,
    Font,
    Trendline,
    XlsbUpdater,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)
from xlspy.exceptions import UnsupportedFeatureError
from xlspy.formulas import Formula, normalize_formula
from xlspy.styles.numbers import BUILTIN_FORMATS

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
CHART = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
EXTERNAL_DATA = Path(__file__).parent / "data" / "external_ref.xlsx"


def parts(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def xml_root(data: bytes):
    return ET.fromstring(data)


def xlsb_records(data: bytes):
    """Yield ``(record_id, payload)`` from a BIFF12 stream."""

    def varint(i):
        value = shift = 0
        while True:
            byte = data[i]
            i += 1
            value |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                return value, i

    i = 0
    while i < len(data):
        record_id, i = varint(i)
        size, i = varint(i)
        yield record_id, data[i:i + size]
        i += size


def xlsb_sheet(path, number=1):
    return parts(path)[f"xl/worksheets/sheet{number}.bin"]


ROWS = [["h1", "h2", "h3"], [1, 2, 3], [4, 5, 6]]


# --- XLSB sheet dimension -------------------------------------------------

def _xlsb_base(tmp_path):
    base = tmp_path / "base.xlsb"
    with XlsbWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet(ROWS)
    return base


def _dimension_and_selection(data):
    dimension = selection = None
    for record_id, payload in xlsb_records(data):
        if record_id == 0x94:
            dimension = struct.unpack("<iiii", payload[:16])
        if record_id == 0x98:
            selection = payload
    return dimension, selection


@pytest.mark.parametrize("streaming", [False, True])
def test_xlsb_replace_writes_dimension_and_leaves_selection(tmp_path, streaming):
    base = _xlsb_base(tmp_path)
    _, selection_before = _dimension_and_selection(xlsb_sheet(base))
    new_rows = [[r * 10 + c for c in range(6)] for r in range(10)]
    out = tmp_path / "out.xlsb"
    updater = XlsbUpdater(base)
    if streaming:
        updater.replace_sheet_data_stream("S", new_rows)
    else:
        updater.replace_sheet_data("S", new_rows)
    updater.save(out)

    dimension, selection_after = _dimension_and_selection(xlsb_sheet(out))
    assert dimension == (0, 9, 0, 5)
    assert selection_after == selection_before


# --- chart formulas -------------------------------------------------------

def _chart_formulas(root):
    return [node.text for node in root.iter(f"{CHART}f")]


@pytest.mark.parametrize("ext", ["xlsx", "xlsb"])
def test_chart_formula_escapes_sheet_name(tmp_path, ext):
    writer = XlsxWriter if ext == "xlsx" else XlsbWriter
    path = tmp_path / f"amp.{ext}"
    with writer(path) as wb:
        wb.add_sheet("Q&A")
        wb.write_sheet(
            [["Q", "A", "B"], ["Q1", 10, 20], ["Q2", 30, 40]],
            charts=[Chart("bar", [ChartSeries("B2:B3", title_cell="B1", categories="A2:A3")])],
        )
    chart = xml_root(parts(path)["xl/charts/chart1.xml"])  # raises if not well-formed
    assert "'Q&A'!$B$2:$B$3" in _chart_formulas(chart)


def test_radar_trendline_is_rejected():
    with pytest.raises(ValueError, match="radar"):
        Chart("radar", [ChartSeries("B2:B4", trendline=Trendline("linear"))])


def test_chart_title_size_is_written_as_an_integer(tmp_path):
    path = tmp_path / "size.xlsx"
    with XlsxWriter(path) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["a"], [1]], charts=[Chart("bar", [ChartSeries("A2")], title="T", title_size=12.345)])
    chart = parts(path)["xl/charts/chart1.xml"].decode()
    sizes = re.findall(r'sz="([^"]+)"', chart)
    assert sizes and all(size.isdigit() for size in sizes)


def test_control_character_in_chart_title_is_rejected():
    with pytest.raises(ValueError, match="control character"):
        Chart("bar", [ChartSeries("A2")], title="bad\x01title")


# --- sheet names -----------------------------------------------------------

def test_sheet_name_with_double_quote_round_trips_in_xlsx(tmp_path):
    path = tmp_path / "quote.xlsx"
    with XlsxWriter(path) as wb:
        wb.add_sheet('Say "hi"')
        wb.write_sheet([["a"]])
    assert openpyxl.load_workbook(path).sheetnames == ['Say "hi"']


def test_sheet_name_with_control_character_is_rejected(tmp_path):
    with pytest.raises(ValueError, match="control character"):
        with XlsxWriter(tmp_path / "ctl.xlsx") as wb:
            wb.add_sheet("ctl\x01x")


# --- calcChain ---------------------------------------------------------------

CALC_CHAIN = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<c r="B2" i="1"/><c r="B3" i="1"/></calcChain>'
)


def _with_calc_chain(source, target):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/'
                    b'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
                )
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/calcChain.xml" ContentType="application/'
                    b'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
                )
            zout.writestr(item.filename, data)
        zout.writestr("xl/calcChain.xml", CALC_CHAIN)


@pytest.mark.parametrize("streaming", [False, True])
def test_replacing_formulas_drops_the_calc_chain(tmp_path, streaming):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["h1", "h2"], [1, Formula("A2*2")], [2, Formula("A3*2")]])
    chained = tmp_path / "chained.xlsx"
    _with_calc_chain(base, chained)

    out = tmp_path / "out.xlsx"
    updater = XlsxUpdater(chained)
    if streaming:
        updater.replace_sheet_data_stream("S", [["h1", "h2"], [5, 6]])
    else:
        updater.replace_sheet_data("S", [["h1", "h2"], [5, 6]])
    updater.save(out)

    package = parts(out)
    assert "xl/calcChain.xml" not in package
    assert b"calcChain" not in package["xl/_rels/workbook.xml.rels"]
    assert b"<f>" not in package["xl/worksheets/sheet1.xml"]


# --- formulas in replace_sheet_data ------------------------------------------

@pytest.mark.parametrize("streaming", [False, True])
def test_xlsx_replace_writes_formula_with_cached_value(tmp_path, streaming):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["h1", "h2"], [1, 2]])
    out = tmp_path / "out.xlsx"
    updater = XlsxUpdater(base)
    rows = [["h1", "h2"], [7, Formula("A2*3", value=21)], [8, Formula('"x"&"y"', value="xy")]]
    if streaming:
        updater.replace_sheet_data_stream("S", rows)
    else:
        updater.replace_sheet_data("S", rows)
    updater.save(out)

    formulas = openpyxl.load_workbook(out)["S"]
    cached = openpyxl.load_workbook(out, data_only=True)["S"]
    assert formulas["B2"].value == "=A2*3"
    assert cached["B2"].value == 21
    assert formulas["B3"].value == '="x"&"y"'
    assert cached["B3"].value == "xy"


def test_xlsb_replace_refuses_formulas_instead_of_writing_their_text(tmp_path):
    base = _xlsb_base(tmp_path)
    updater = XlsbUpdater(base)
    with pytest.raises(UnsupportedFeatureError, match="update_cells"):
        updater.replace_sheet_data("S", [["a", Formula("A1+1")]])


def test_formula_names_get_the_prefix_excel_needs():
    assert normalize_formula("=COT(1)") == "_xlfn.COT(1)"
    assert normalize_formula("=SORT(A1:A3)") == "_xlfn._xlws.SORT(A1:A3)"
    assert normalize_formula("=FILTER(A1:A3,B1:B3>1)") == "_xlfn._xlws.FILTER(A1:A3,B1:B3>1)"
    assert normalize_formula("=SUM(1,2)") == "SUM(1,2)"


# --- writers ---------------------------------------------------------------

def test_xlsb_writes_a_part_for_a_sheet_never_written(tmp_path):
    path = tmp_path / "empty.xlsb"
    with XlsbWriter(path) as wb:
        wb.add_sheet("First")
        wb.write_sheet([["a"]])
        wb.add_sheet("Never written")
    package = parts(path)
    assert "xl/worksheets/sheet2.bin" in package
    assert load_workbook(path).sheetnames == ["First", "Never written"]


@pytest.mark.parametrize("writer", [XlsxWriter, XlsbWriter])
def test_second_write_sheet_on_one_sheet_is_rejected(tmp_path, writer):
    path = tmp_path / "twice"
    with pytest.raises(ValueError, match="already has its data"):
        with writer(path) as wb:
            wb.add_sheet("S")
            wb.write_sheet([["a"]])
            wb.write_sheet([["b"]])


@pytest.mark.parametrize("writer", [XlsxWriter, XlsbWriter])
def test_table_names_differing_only_in_case_are_rejected(tmp_path, writer):
    rows = [["h", "h2", None, "d1", "d2"], [1, 2, None, 3, 4], [5, 6, None, 7, 8]]
    with pytest.raises(ValueError, match="Excel ignores case"):
        with writer(tmp_path / "tables") as wb:
            wb.add_sheet("S")
            wb.write_sheet(
                rows,
                tables=[ExcelTable("Sales", "A1:B3"), ExcelTable("sales", "D1:E3")],
            )


def test_xlsb_autofilter_and_widths_cover_columns_past_the_first_rows(tmp_path):
    rows = [["x", "x"] for _ in range(150)]
    rows[120] = ["x"] * 6  # six columns, first seen after the first 100 rows
    path = tmp_path / "wide.xlsb"
    with XlsbWriter(path) as wb:
        wb.add_sheet("S")
        wb.write_sheet(rows, column_widths={5: 40}, auto_filter=True)
    ws = load_workbook(path)["S"]
    assert ws.auto_filter.ref == "A1:F150"
    assert ws.column_dimensions["F"].width == pytest.approx(40)


# --- loaded workbooks ------------------------------------------------------

def test_builtin_accounting_format_survives_a_font_edit(tmp_path):
    fmt = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    base = tmp_path / "base.xlsx"
    book = openpyxl.Workbook()
    book.active["A1"] = 1234.5
    book.active["A1"].number_format = fmt
    book.save(base)
    # Re-point the custom format at built-in id 44, as Excel writes accounting formats.
    builtin = tmp_path / "builtin.xlsx"
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(builtin, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = re.sub(rb"<numFmts[^>]*>.*?</numFmts>", b"", data, flags=re.S)
                data = text.replace(b'numFmtId="164"', b'numFmtId="44"')
            zout.writestr(item.filename, data)

    # The built-in spelling of id 44 (openpyxl's), not the custom string written above.
    builtin_44 = BUILTIN_FORMATS[44]
    loaded = load_workbook(builtin)
    assert loaded.active["A1"].number_format == builtin_44
    loaded.active["A1"].font = Font(bold=True)
    out = tmp_path / "out.xlsx"
    loaded.save(out)
    assert openpyxl.load_workbook(out).active["A1"].number_format == builtin_44


def test_rows_of_a_self_closing_sheetdata_go_inside_it(tmp_path):
    path = tmp_path / "blank.xlsx"
    with XlsxWriter(path) as wb:
        wb.add_sheet("Blank")  # never written: <sheetData/>
    wb = load_workbook(path)
    wb["Blank"].row_dimensions[2].height = 30
    out = tmp_path / "blank_out.xlsx"
    wb.save(out)
    root = xml_root(parts(out)["xl/worksheets/sheet1.xml"])
    sheet_data = root.find(f"{MAIN}sheetData")
    assert [row.get("r") for row in sheet_data] == ["2"]


def test_dimension_follows_sheetpr_when_the_updater_adds_it(tmp_path):
    base = tmp_path / "base.xlsx"
    with XlsxWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([["a", "b"], [1, 2]], tab_color="#FF0000")
    stripped = tmp_path / "no_dimension.xlsx"
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<dimension[^>]*/>", b"", data)
            zout.writestr(item.filename, data)

    for method in ("replace_sheet_data", "replace_sheet_data_stream"):
        out = tmp_path / f"{method}.xlsx"
        updater = XlsxUpdater(stripped)
        getattr(updater, method)("S", [["a", "b"], [3, 4]])
        updater.save(out)
        names = [child.tag.replace(MAIN, "") for child in xml_root(parts(out)["xl/worksheets/sheet1.xml"])]
        assert names.index("sheetPr") < names.index("dimension") < names.index("sheetData"), method


# --- external links ---------------------------------------------------------

def test_unresolved_external_link_keeps_later_numbering(tmp_path):
    dst = tmp_path / "two_links.xlsx"
    with zipfile.ZipFile(EXTERNAL_DATA) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r"(<(?:[\w.-]+:)?externalReferences[^>]*>)",
                    r'\1<externalReference xmlns:r="http://schemas.openxmlformats.org/'
                    r'officeDocument/2006/relationships" r:id="rIdMissing"/>',
                    text,
                    count=1,
                )
                data = text.encode("utf-8")
            zout.writestr(item.filename, data)

    links = load_workbook(dst).external_links
    # The first reference has no relationship; the real link is still [2] in formulas.
    assert [link.index for link in links] == [1, 2]
    assert links[0].part is None
    assert links[1].target == "ext.xlsx"


# --- table headers in replace_sheet_data -----------------------------------

FORMATS = pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
STREAMING = pytest.mark.parametrize("streaming", [False, True])
TABLE_ROWS = [["Name", "Qty"], ["a", 1], ["b", 2]]


def _table_base(tmp_path, suffix, rows=TABLE_ROWS, ref="A1:B3"):
    base = tmp_path / f"base.{suffix}"
    writer = XlsxWriter if suffix == "xlsx" else XlsbWriter
    with writer(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet(rows, tables=[ExcelTable("Tbl", ref)])
    return base


def _updater(path, suffix):
    return XlsxUpdater(path) if suffix == "xlsx" else XlsbUpdater(path)


def _replace(updater, rows, streaming, headers=None):
    if streaming:
        updater.replace_sheet_data_stream("S", rows, headers=headers)
    else:
        updater.replace_sheet_data("S", rows, headers=headers)


@FORMATS
@STREAMING
def test_table_header_mismatch_is_refused_and_package_unchanged(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    before = parts(base)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell A1"):
        _replace(updater, [["Nome", "Qty"], ["c", 3]], streaming)
    out = tmp_path / f"out.{suffix}"
    updater.save(out)
    assert parts(out) == before


@FORMATS
@STREAMING
def test_table_header_mismatch_with_headers_argument_is_refused(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell B1"):
        _replace(updater, [["c", 3]], streaming, headers=["Name", "Amount"])


@FORMATS
@STREAMING
def test_table_header_mismatch_in_the_middle_of_the_data_is_refused(tmp_path, suffix, streaming):
    # The table's header is row 4; the new data changes it on the way down.
    rows = [["intro", None], [None, None], [None, None], ["Name", "Qty"], ["c", 3]]
    base = _table_base(tmp_path, suffix, rows=rows, ref="A4:B5")
    updater = _updater(base, suffix)
    new_rows = [["x"], ["y"], ["z"], ["Name", "Other"], ["c", 3]]
    with pytest.raises(UnsupportedFeatureError, match="header cell B4"):
        _replace(updater, new_rows, streaming)


@FORMATS
@STREAMING
def test_data_too_short_to_reach_the_table_header_is_refused(tmp_path, suffix, streaming):
    rows = [["intro", None], [None, None], [None, None], ["Name", "Qty"], ["c", 3]]
    base = _table_base(tmp_path, suffix, rows=rows, ref="A4:B5")
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError, match="header cell A4"):
        _replace(updater, [["only", 1]], streaming)


@FORMATS
@STREAMING
def test_matching_headers_are_accepted_ignoring_case_and_length(tmp_path, suffix, streaming):
    base = _table_base(tmp_path, suffix)
    out = tmp_path / f"out.{suffix}"
    updater = _updater(base, suffix)
    # Excel only needs the header cells to match; a shorter range is not corruption.
    _replace(updater, [["NAME", "qty"], ["c", 3]], streaming)
    updater.save(out)

    sheet = load_workbook(out)["S"]
    assert [sheet.cell(row=2, column=c).value for c in (1, 2)] == ["c", 3]


def _truncate_sheet_after_last_cell(source, target):
    """Copy an XLSB whose first sheet stream ends right after its last cell record."""

    def read_varint(data, i):
        value = shift = 0
        while True:
            byte = data[i]
            i += 1
            value |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                return value, i

    with zipfile.ZipFile(source) as zin:
        names = zin.namelist()
        members = {name: zin.read(name) for name in names}
    sheet = members["xl/worksheets/sheet1.bin"]
    end_of_last_cell = None
    i = 0
    while i < len(sheet):
        record_id, i = read_varint(sheet, i)
        size, i = read_varint(sheet, i)
        if record_id in range(0x01, 0x0C):
            end_of_last_cell = i + size
        i += size
    members["xl/worksheets/sheet1.bin"] = sheet[:end_of_last_cell]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, members[name])


def test_pure_python_xlsb_reader_stops_at_eof_after_a_cell(tmp_path, monkeypatch):
    """Ignoring read_worksheet's result at EOF repeated the last cell forever."""
    from xlspy import biff_reader
    from xlspy import reader as reader_module
    from xlspy.reader import ExcelReader

    base = tmp_path / "base.xlsb"
    with XlsbWriter(base) as wb:
        wb.add_sheet("S")
        wb.write_sheet([[1, 2], [3, 4]])
    truncated = tmp_path / "eof.xlsb"
    _truncate_sheet_after_last_cell(base, truncated)

    # The native reader does not reach the Python loop; force the fallback.
    monkeypatch.setattr(reader_module, "HAVE_NATIVE", False)
    original = biff_reader.BiffReader.read_worksheet
    calls = []

    def bounded(self):
        calls.append(None)
        if len(calls) > 10_000:
            raise AssertionError("read_worksheet kept returning the last cell past the end of the stream")
        return original(self)

    monkeypatch.setattr(biff_reader.BiffReader, "read_worksheet", bounded)
    with ExcelReader(truncated) as reader:
        assert list(reader.get_rows("S")) == [[1, 2], [3, 4]]


@FORMATS
def test_mismatch_after_refused_write_leaves_later_valid_write_working(tmp_path, suffix):
    base = _table_base(tmp_path, suffix)
    updater = _updater(base, suffix)
    with pytest.raises(UnsupportedFeatureError):
        updater.replace_sheet_data("S", [["Other", "Qty"], ["c", 3]])
    updater.replace_sheet_data("S", [["Name", "Qty"], ["c", 3]])
    out = tmp_path / f"out.{suffix}"
    updater.save(out)
    assert load_workbook(out)["S"].cell(row=2, column=1).value == "c"
