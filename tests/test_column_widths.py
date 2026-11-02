"""Automatic column widths: displayed text, per-character widths, writers."""

import datetime
import re
import struct
import zipfile

import pytest

from xlspy import Alignment, ExcelTable, Font, Style, XlsbWriter, XlsxWriter
from xlspy._widths import (
    DEFAULT_WIDTH, FILTER_BUTTON, cell_text, cell_width, estimate_column_widths, file_width,
    text_pixels, text_width,
)
from xlspy.biff_utils import iter_records
from xlspy import load_workbook


@pytest.mark.parametrize("value,number_format,expected", [
    (0.1 + 0.2, None, "0.3"),
    (1 / 3, None, "0.333333333"),
    (1234567890123, None, "1.23457E+12"),
    (2.5e15, None, "2.5E+15"),
    (1e-7, None, "0.0000001"),
    (-98765.4321, None, "-98765.4321"),
    (42, None, "42"),
    (1234567.891, "#,##0.00", "1,234,567.89"),
    (1234567.891, "#,##0", "1,234,568"),
    (0.4567, "0%", "46%"),
    (0.4567, "0.00%", "45.67%"),
    (12345, "000000000", "000012345"),
    (1234.5678, "0.00E+00", "1.23E+03"),
    (-1234.5, "#,##0.00;[Red]-#,##0.00", "-1,234.50"),
    (1234.5, "[$€-407] #,##0.00", "€ 1,234.50"),
    (True, None, "PRAWDA"),
    ("text", None, "text"),
    (datetime.date(2024, 9, 30), None, "12/31/2024"),
    (datetime.datetime(2024, 9, 30, 13, 45), None, "12/31/2024 23:45"),
    (datetime.date(2024, 9, 30), "yyyy-mm-dd", "2024-12-30"),
])
def test_cell_text(value, number_format, expected):
    assert cell_text(value, number_format) == expected


def test_long_date_names_are_generous():
    # Month and day names come from the reader's language: take long ones.
    text = cell_text(datetime.date(2024, 9, 30), "dddd, d mmmm yyyy")
    assert text_width(text) >= text_width("poniedziałek, 30 wrzesień 2024")
    assert "zł" in cell_text(1234.5, '#,##0.00 "zł"')


def test_character_widths():
    assert text_width("WWWW") > text_width("abcd") > text_width("iiii")
    assert text_width("中文") == pytest.approx(32 / 7)  # 16 pixels each
    assert text_width("0123456789") == 10.0
    assert text_width("WWWW", bold=True) > text_width("WWWW")


# Excel's AutoFit ColumnWidth (regular, bold) for Calibri 11 on a display at
# 100% scaling (1920x1080, 96 DPI), measured with tools/calibrate_column_widths.py.
# Excel measures text in GDI pixels at that size; an earlier table, measured at
# 175% scaling, came out up to 7% narrow ("i" is 4/7 of a digit at 96 DPI but
# 0.46 at 175%, and long text gets 3.6% more room).
EXCEL_AUTOFIT_100 = [
    ("Nazwa klienta", 12.86, 12.86),
    ("Customer name", 14.43, 14.43),
    ("Product", 7.14, 7.14),
    ("Quantity", 8.0, 8.0),
    ("ID", 2.14, 2.14),
    ("Jan Kowalski", 11.43, 11.43),
    ("john.smith@example.com", 24.14, 24.14),
    ("https://example.com/products/12345", 34.43, 34.43),
    ("Wrocław, ul. Świdnicka 12/4", 25.43, 25.57),
    ("Zażółć gęślą jaźń", 14.71, 14.86),
    ("ORDER-2024-000123", 18.0, 18.0),
    ("The quick brown fox jumps over the lazy dog", 40.43, 40.43),
    ("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor", 72.71, 72.71),
    ("Faktura VAT nr FV/2024/11/0042", 29.14, 29.14),
    ("Opis pozycji zamówienia", 22.43, 22.43),
    ("Kwota brutto", 11.86, 11.86),
    ("PL61109010140000071219812874", 29.0, 29.0),
    ("+48 600 123 456", 13.71, 13.71),
    ("iiiiiiiiii", 6.0, 6.0),
    ("aaaaaaaaaa", 10.29, 10.29),
    ("mmmmmmmmmm", 17.71, 17.71),
    ("WWWWWW", 11.43, 12.43),
    ("0", 1.29, 1.29),
    ("00", 2.29, 2.29),
    ("1234567890", 10.29, 10.29),
    ("x" * 60, 62.14, 62.14),
    ("illicit little lilies fill the hill", 24.14, 24.14),
    ("ŁÓDŹ ŚWIĘTOKRZYSKIE", 20.43, 20.86),
]


# Excel's AutoFit is matched to the pixel; xlspy adds one spare pixel.
SPARE = 1 / 7


@pytest.mark.parametrize("text,regular,bold", EXCEL_AUTOFIT_100)
def test_cell_width_matches_excel_autofit(text, regular, bold):
    for expected, is_bold in ((regular, False), (bold, True)):
        assert cell_width(text, bold=is_bold) == pytest.approx(expected + SPARE, abs=0.006), (text, is_bold)


# (value, number_format, font size, bold, Excel's AutoFit ColumnWidth) at 100%
# scaling, sampled from 11,786 measured cells (all matched): other font sizes,
# and General numbers, whose room Excel reserves differently from the text
# (decimal point as wide as a digit, E-05 as wide as E+05).
EXCEL_AUTOFIT_SAMPLES = [
    (-816699358, None, 8, False, 8.43),  # -816699358
    (-3945296.002, None, 10, False, 11.86),  # -3945296.002
    (-7858293.27, None, 32, True, 35.57),  # -7858293.27
    (-825273.6, None, 28, False, 24.86),  # -825273.6
    (13.654381472, None, None, False, 11.29),  # 13.65438147
    (-726404.2261447909, None, None, False, 12.0),  # -726404.2261
    (-4678636.32, None, 24, True, 25.71),  # -4678636.32
    (-801.01346055244, None, None, False, 12.0),  # -801.0134606
    (787486084089.6, None, None, True, 11.29),  # 7.87486E+11
    (585681337, None, 18, False, 16.14),  # 585681337
    (6383925.2928, None, 10, True, 11.29),  # 6383925.293
    (405947.71218181145, None, None, False, 11.29),  # 405947.7122
    (-784849349, None, 11, False, 10.0),  # -784849349
    (5140376111136.4, None, None, False, 11.29),  # 5.14038E+12
    (-586911916, None, 14, True, 14.43),  # -586911916
    (-476644.3945326506, None, None, False, 12.0),  # -476644.3945
    (-839563782869.5443, None, None, False, 12.0),  # -8.39564E+11
    (-3.702533000906521e-06, None, 18, True, 20.71),  # -3.70253E-06
    (0.006981919770114465, None, 10.5, True, 10.29),  # 0.00698192
    (177159.85303789494, None, None, False, 10.29),  # 177159.853
    (-510446531, None, 28, False, 27.71),  # -510446531
    (72540591640, None, None, False, 11.29),  # 72540591640
    (-429166.1224131858, None, None, False, 12.0),  # -429166.1224
    (261428754, None, 20, False, 19.14),  # 261428754
    (-481734.6021240736, None, None, False, 12.0),  # -481734.6021
    (3845473.1188, None, 11, True, 11.29),  # 3845473.119
    (0.0, None, None, True, 1.29),  # 0
    (-7928748.93, None, 9, False, 9.43),  # -7928748.93
    (-331790572, None, 14, True, 14.43),  # -331790572
    (2921250.36, None, 9, False, 8.86),  # 2921250.36
    (224632.58968747337, '0.00%', None, False, 12.57),  # 22463258.97%
    (80914419563, '0.00%', None, False, 17.71),  # 8091441956300.00%
    (-1315.95, '0.00E+00', None, False, 8.57),  # -1.32E+03
    (97803292, '0.00%', 14, True, 20.71),  # 9780329200.00%
    (0.0, '0.00%', None, False, 5.43),  # 0.00%
    (-56961.28905593057, '#,##0', None, True, 6.57),  # -56,961
    (-8.4811915, '#,##0.00', None, False, 4.57),  # -8.48
    (0.0461637240908247, '0.00%', 18, True, 9.14),  # 4.62%
    (-49152890211280, '0.000', None, False, 18.86),  # -49152890211280.000
    (110461.59134260612, '000000', None, False, 6.29),  # 110462
    (-104216408723248, '#,##0.00;[Red]-#,##0.00', None, True, 21.29),  # -104,216,408,723,248.00
    (-8083569990691.486, '0.00%', 24, False, 46.86),  # -808356999069149.00%
    ('ufłŃo0%FÓOłR:-ŹźWX(ŃŹZąME!H', None, 24, False, 64.86),  # ufłŃo0%FÓOłR:-ŹźWX(ŃŹZąME!H
    ('m?/)_tŁęxŃlF*V-ńD5tÓtMS#Óy=Ńoo', None, 20, False, 59.71),  # m?/)_tŁęxŃlF*V-ńD5tÓtMS#Óy=Ńoo
    ('SbwĄkóprłEsŚ2Ń]nźźu3O*łYr', None, 15, True, 36.0),  # SbwĄkóprłEsŚ2Ń]nźźu3O*łYr
    ('brutto of klient as', None, 16, False, 22.71),  # brutto of klient as
    ('netto zamówienie product customer', None, 10, False, 29.0),  # netto zamówienie product customer
    ('wartość klient', None, 28, False, 31.71),  # wartość klient
    ('-ĘyHBV7G2qjYAkŃ3e&D', None, 24, False, 47.43),  # -ĘyHBV7G2qjYAkŃ3e&D
    ('VŁx0jj#1*K[9C3Eb', None, 13, False, 18.57),  # VŁx0jj#1*K[9C3Eb
    ('zamówienie', None, 20, True, 20.71),  # zamówienie
    ('of with kwota the', None, 20, False, 28.71),  # of with kwota the
    ('netto was of amount', None, 24, True, 41.43),  # netto was of amount
    ('by kwota brutto of are netto order with', None, 26, False, 83.71),  # by kwota brutto of are netto order with
    ('aOżKgDxĆG 0#9Ń#ć[OqKr', None, 9, True, 19.0),  # aOżKgDxĆG 0#9Ń#ć[OqKr
    ('h63?Ń!qMUfx', None, 10.5, True, 11.86),  # h63?Ń!qMUfx
    ('total the', None, 10.5, False, 7.29),  # total the
    ('by are', None, 22, False, 11.29),  # by are
    ('QbKtć.Z_Wgną9Q_r-2xKFzN', None, 13, False, 28.57),  # QbKtć.Z_Wgną9Q_r-2xKFzN
    ('are wartość was', None, 22, True, 29.71),  # are wartość was
    ('amount invoice', None, 15, True, 19.43),  # amount invoice
    (':O0VPI0vĄj9+0kqbnuP;ós].', None, 11, True, 24.14),  # :O0VPI0vĄj9+0kqbnuP;ós].
    ('BZv-XóŚT*eTRXĘ=SgŻrćR)?S(;7dUę/!źĆpÓ', None, 11, False, 37.0),  # BZv-XóŚT*eTRXĘ=SgŻrćR)?S(;7dUę/!źĆpÓ
    ('net faktura brutto the', None, 12, True, 22.57),  # net faktura brutto the
    ('of order from Wrocław and of kwota', None, 18, False, 53.0),  # of order from Wrocław and of kwota
    ('kwota zamówienie Łódź', None, 13, True, 25.43),  # kwota zamówienie Łódź
    ('brutto zamówienie Łódź', None, 24, True, 47.86),  # brutto zamówienie Łódź
    ('ęłe', None, 36, False, 10.0),  # ęłe
    ('as at report order on with by report', None, 10.5, True, 31.29),  # as at report order on with by report
    ('6Żucg', None, 26, True, 13.57),  # 6Żucg
    ('and netto report brutto', None, 24, False, 46.29),  # and netto report brutto
    ('Eó]7-GDJDrń4.(źa&bńDeÓGŹ', None, 24, True, 57.0),  # Eó]7-GDJDrń4.(źa&bńDeÓGŹ
]


@pytest.mark.parametrize("value,number_format,size,bold,expected", EXCEL_AUTOFIT_SAMPLES)
def test_cell_width_matches_excel_samples(value, number_format, size, bold, expected):
    got = cell_width(value, number_format, bold=bold, size=size)
    assert got == pytest.approx(expected + SPARE, abs=0.006)


@pytest.mark.parametrize("value,expected", [
    (-6619407.0022, "-6619407.002"),       # the sign is not one of the 11 characters
    (49062358930.42197, "49062358930"),    # whole-number zeros are kept
    (-9526850.4905, "-9526850.491"),       # rounded half up, as a decimal
    (4.995176312908987e-06, "4.99518E-06"),  # fixed would keep 4 digits, scientific 6
    (-6.077e-06, "-0.000006077"),          # fixed shows it exactly
    (0.00115916, "0.00115916"),
    (1e-7, "0.0000001"),
])
def test_general_number_text(value, expected):
    assert cell_text(value) == expected


def test_size_pixels_follow_gdi_hinting():
    # GDI's widths (Windows): Calibri 9 pt is 12 pixels high, where hinting
    # draws "a" as wide as at 11 pt and wider than the digit 0.
    assert text_pixels("aaaaaaaaaa", size=9) == 70
    assert text_pixels("0000000000", size=9) == 60
    assert text_pixels("aaaaaaaaaa", size=11) == 70
    assert text_pixels("aaaaaaaaaa", size=16) == 100


def test_cell_width_options():
    plain = cell_width("WWWWWW")
    assert cell_width("WWWWWW", bold=True) > plain
    assert cell_width("WWWWWW", size=22) == pytest.approx(2 * plain, rel=0.05)
    two_lines = "short\na much longer second line"
    assert cell_width(two_lines, wrap=True) < cell_width(two_lines)
    assert cell_width(None) == 0.0


def test_estimate_column_widths():
    rows = [["ID", "Customer name", None], [1, "x", None], [2, "y", (1234567.891, "#,##0.00")]]
    widths = estimate_column_widths(rows, 3)
    assert widths[0] == DEFAULT_WIDTH  # never narrower than Excel's default
    assert widths[1] > DEFAULT_WIDTH
    assert widths[2] == pytest.approx(cell_width(1234567.891, "#,##0.00"))
    with_filter = estimate_column_widths(rows, 3, filter_cells=[(0, 0, 2)])
    assert with_filter[1] == pytest.approx(widths[1] + FILTER_BUTTON)
    styled = estimate_column_widths(
        rows, 3, lambda row_idx, length: {1: Style(font=Font(size=20))} if row_idx else {},
    )
    assert styled[1] == widths[1]  # the big font only applies to the short cells
    wrapped = estimate_column_widths(
        [["a fairly long heading\nx"]], 1, lambda r, n: {0: Style(alignment=Alignment(wrap_text=True))},
    )
    assert wrapped[0] < estimate_column_widths([["a fairly long heading\nx"]], 1)[0]


def _xlsx_widths(path):
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode()
    return {int(col) - 1: float(width) for col, width in
            re.findall(r'<col min="(\d+)" max="\d+" width="([0-9.]+)"', xml)}


def _xlsb_widths(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    widths = {}
    for record in iter_records(data):
        if record.record_id == 0x003C:
            first, _last, coldx = struct.unpack_from("<III", data, record.data_start)
            widths[first] = coldx / 256
    return widths


ROWS = [
    ["Order number", "Amount", "Note"],
    ["ORDER-2024-000123", (1234567.891, "#,##0.00"), "WWWWWWWWWWWW"],
    ["x", 2.5, "中文字符测试中文字符测试"],
]


@pytest.mark.parametrize("writer_cls,ext,read", [
    (XlsxWriter, "xlsx", _xlsx_widths), (XlsbWriter, "xlsb", _xlsb_widths),
])
def test_writers_use_the_estimate(tmp_path, writer_cls, ext, read):
    path = tmp_path / f"widths.{ext}"
    with writer_cls(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(ROWS)
    widths = read(path)
    expected = estimate_column_widths(ROWS, 3, filter_cells=[(0, 0, 2)])
    for col in range(3):
        # The file stores the width with Excel's cell padding.
        assert widths[col] == pytest.approx(file_width(expected[col]), abs=0.01)
    assert widths[2] > 20  # the CJK text needs about two digits per character
    assert file_width(DEFAULT_WIDTH) == pytest.approx(9.140625, abs=0.005)  # Excel's own default


def test_xlsb_counts_columns_after_empty_cells(tmp_path):
    path = tmp_path / "gaps.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[None, None, "a long text in column C"]], auto_filter=False)
    widths = _xlsb_widths(path)
    assert widths[2] > 15


def test_table_header_gets_room_for_its_button(tmp_path):
    rows = [["Customer name", "Value"], ["a", 1], ["b", 2]]
    path = tmp_path / "table.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, tables=[ExcelTable(ref="A1:B3", display_name="Orders")])
    with_button = estimate_column_widths(rows, 2, filter_cells=[(0, 0, 1)])
    assert _xlsx_widths(path)[0] == pytest.approx(file_width(with_button[0]), abs=0.01)
    assert with_button[0] > estimate_column_widths(rows, 2)[0]


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


# --- autofit_rows / autofit_padding ----------------------------------------

def _write(writer_cls, path, rows, **kwargs):
    with writer_cls(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet(rows, auto_filter=False, **kwargs)
    return (_xlsx_widths if writer_cls is XlsxWriter else _xlsb_widths)(path)


def _long_late_rows(late_row):
    rows = [["id", "text"]] + [[i, "short"] for i in range(1, 3000)]
    rows[late_row] = [late_row, "a much longer text that comes late in the sheet"]
    return rows


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_autofit_measures_the_first_1000_rows_by_default(tmp_path, writer_cls, ext):
    long_text = file_width(cell_width("a much longer text that comes late in the sheet"))
    inside = _write(writer_cls, tmp_path / f"in.{ext}", _long_late_rows(999))
    assert inside[1] == pytest.approx(long_text, abs=0.01)
    outside = _write(writer_cls, tmp_path / f"out.{ext}", _long_late_rows(1000))
    assert outside.get(1, file_width(DEFAULT_WIDTH)) < long_text


@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_autofit_rows_and_padding(tmp_path, writer_cls, ext):
    rows = _long_late_rows(2500)
    long_text = file_width(cell_width("a much longer text that comes late in the sheet"))
    every_row = _write(writer_cls, tmp_path / f"all.{ext}", rows, autofit_rows=None)
    assert every_row[1] == pytest.approx(long_text, abs=0.01)
    few = _write(writer_cls, tmp_path / f"few.{ext}", rows, autofit_rows=3000)
    assert few[1] == pytest.approx(long_text, abs=0.01)
    padded = _write(writer_cls, tmp_path / f"pad.{ext}", rows, autofit_rows=None, autofit_padding=2.5)
    assert padded[1] == pytest.approx(long_text + 2.5, abs=0.01)
    # 0 rows: no measurement, Excel's default width.
    off = _write(writer_cls, tmp_path / f"off.{ext}", rows, autofit_rows=0, autofit_padding=3)
    assert off.get(1, file_width(DEFAULT_WIDTH)) == pytest.approx(file_width(DEFAULT_WIDTH), abs=0.01)
    # Explicit widths win over both.
    fixed = _write(writer_cls, tmp_path / f"fixed.{ext}", rows, column_widths={1: 20}, autofit_padding=5)
    assert fixed[1] == pytest.approx(20, abs=0.01)


def test_autofit_options_in_sheet_options(tmp_path):
    from xlspy import SheetOptions
    path = tmp_path / "opts.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet(_long_late_rows(2500), options=SheetOptions(autofit_rows=None, autofit_padding=1))
    long_text = file_width(cell_width("a much longer text that comes late in the sheet"))
    assert _xlsx_widths(path)[1] == pytest.approx(long_text + 1, abs=0.01)


@pytest.mark.parametrize("kwargs,error", [
    ({"autofit_rows": -1}, ValueError),
    ({"autofit_rows": 1.5}, TypeError),
    ({"autofit_rows": True}, TypeError),
    ({"autofit_padding": -1}, ValueError),
    ({"autofit_padding": "2"}, TypeError),
])
@pytest.mark.parametrize("writer_cls,ext", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_autofit_options_are_checked(tmp_path, writer_cls, ext, kwargs, error):
    with writer_cls(tmp_path / f"bad.{ext}") as writer:
        writer.add_sheet("S")
        with pytest.raises(error):
            writer.write_sheet([["a"]], **kwargs)
        writer.write_sheet([["a"]])


def test_estimate_matches_cell_width_with_filters():
    rows = [["Header", "Amount"], ["x", 1234567.25], ["longer text here", -0.5]]
    widths = estimate_column_widths(rows, 2, filter_cells=[(0, 0, 1)])
    header = cell_width("Header", bold=True) + FILTER_BUTTON
    assert widths[0] == pytest.approx(max(DEFAULT_WIDTH, header, cell_width("longer text here")))
    assert widths[1] == pytest.approx(max(DEFAULT_WIDTH, cell_width("Amount", bold=True) + FILTER_BUTTON,
                                          cell_width(1234567.25)))
