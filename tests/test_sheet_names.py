"""Sheet names in the workbook model and in formulas."""

import pytest

from xlspy import Workbook
from xlspy.formula.sheets import rename_sheet_references
from xlspy import XlsbWriter
from xlspy import XlsxWriter
import openpyxl


@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=SUM(Jan:Mar!B2)", "=SUM(January:Mar!B2)"),
        ("=SUM(Dec:jan!B2)", "=SUM(Dec:January!B2)"),
        ("=SUM('Jan:Mar'!B2)", "=SUM(January:Mar!B2)"),
        ("=SUM(Sheet1!A1:Jan!B2)", "=SUM(Sheet1!A1:January!B2)"),
        ('=Jan!A1&"Jan!A1"', '=January!A1&"Jan!A1"'),
        ("=[1]Jan!A1", "=[1]Jan!A1"),
        ("=SUM(Feb:Mar!B2)", "=SUM(Feb:Mar!B2)"),
    ],
)
def test_rename_reaches_both_ends_of_3d_references(formula, expected):
    assert rename_sheet_references(formula, "Jan", "January") == expected


def test_rename_quotes_a_3d_span_when_one_end_needs_it():
    assert rename_sheet_references("=SUM(Jan:Mar!B2)", "Jan", "My Jan") == "=SUM('My Jan:Mar'!B2)"


def test_renaming_a_sheet_updates_3d_formulas():
    workbook = Workbook()
    workbook.active.title = "Jan"
    workbook.create_sheet("Mar")
    summary = workbook.create_sheet("Summary")
    summary["A1"] = "=SUM(Jan:Mar!B2)"
    workbook["Jan"].title = "January"
    assert summary["A1"].value == "=SUM(January:Mar!B2)"


@pytest.mark.parametrize("title", ["data", "DATA", "Chart1"])
def test_sheet_names_are_unique_without_case_across_chartsheets(title):
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.create_chartsheet("Chart1")
    with pytest.raises(ValueError, match="already exists"):
        workbook.create_sheet(title)
    with pytest.raises(ValueError, match="already exists"):
        workbook.create_chartsheet(title)
    other = workbook.create_sheet("Other")
    with pytest.raises(ValueError, match="already exists"):
        other.title = title


@pytest.mark.parametrize("title", ["a/b", "x" * 32, "History", "'quoted'", ""])
def test_chartsheet_names_follow_excel_rules(title):
    workbook = Workbook()
    with pytest.raises(ValueError):
        workbook.create_chartsheet(title)
    chartsheet = workbook.create_chartsheet("Chart")
    with pytest.raises(ValueError):
        chartsheet.title = title


def test_a_sheet_can_change_the_case_of_its_own_name():
    workbook = Workbook()
    workbook.active.title = "data"
    workbook.active.title = "Data"
    chartsheet = workbook.create_chartsheet("chart")
    chartsheet.title = "Chart"
    assert workbook.sheetnames == ["Data"]
    assert [sheet.title for sheet in workbook.chartsheets] == ["Chart"]


def test_default_chartsheet_names_skip_taken_names_in_any_case():
    workbook = Workbook()
    workbook.active.title = "chart1"
    assert workbook.create_chartsheet().title == "Chart2"


def test_renaming_a_sheet_updates_conditional_formats_and_internal_links():
    from xlspy import Hyperlink
    from xlspy.styles import FormulaRule

    workbook = Workbook()
    workbook.active.title = "Data"
    report = workbook.create_sheet("Report")
    report.conditional_formatting.add("A1:A3", FormulaRule("A1:A3", "Data!$A$1>0"))
    report["B1"].hyperlink = Hyperlink("#Data!A1")
    report["B2"].hyperlink = Hyperlink("https://example.com/Data!A1")
    workbook["Data"].title = "Input Data"
    assert report.conditional_formatting[0].rules[0].formula == ("'Input Data'!$A$1>0",)
    assert report["B1"].hyperlink.target == "#'Input Data'!A1"
    assert report["B2"].hyperlink.target == "https://example.com/Data!A1"


@pytest.mark.parametrize("extension", ["xlsx", "xlsb"])
def test_renaming_a_sheet_updates_loaded_charts(tmp_path, extension):
    import zipfile

    from xlspy import Chart, ChartSeries, load_workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in (["name", "value"], ["a", 1], ["b", 2]):
        sheet.append(row)
    sheet.add_chart(Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")]), "D2")
    source = tmp_path / f"source.{extension}"
    workbook.save(source)
    out = tmp_path / f"out.{extension}"
    loaded = load_workbook(source)
    try:
        loaded["Data"].title = "Sales 2026"
        loaded.save(out)
    finally:
        loaded.close()
    with zipfile.ZipFile(out) as archive:
        charts = [archive.read(name).decode() for name in archive.namelist() if name.startswith("xl/charts/chart")]
    assert charts and all("'Sales 2026'!$B$2:$B$3" in xml and "Data!" not in xml for xml in charts)


@pytest.mark.parametrize("writer_cls", [XlsxWriter, XlsbWriter])
@pytest.mark.parametrize(
    "name,message",
    [
        ("", "1 to 31 characters"),
        ("x" * 32, "1 to 31 characters"),
        ("'x", "apostrophe"),
        ("x'", "apostrophe"),
        ("a:b", "forbids"),
        ("a/b", "forbids"),
        (r"a\b", "forbids"),
        ("a?b", "forbids"),
        ("a*b", "forbids"),
        ("[a]", "forbids"),
        ("History", "reserves"),
        ("history", "reserves"),
    ],
)
def test_sheet_names_excel_refuses_are_rejected(tmp_path, writer_cls, name, message):
    with writer_cls(tmp_path / "names.bin") as writer:
        with pytest.raises(ValueError, match=message):
            writer.add_sheet(name)
        writer.add_sheet("Ok")


@pytest.mark.parametrize("writer_cls", [XlsxWriter, XlsbWriter])
def test_duplicate_sheet_names_ignore_case(tmp_path, writer_cls):
    with writer_cls(tmp_path / "dupes.bin") as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="duplicate sheet name"):
            writer.add_sheet("DATA")


@pytest.mark.parametrize("writer_cls", [XlsxWriter, XlsbWriter])
def test_sheet_names_excel_accepts_are_kept(tmp_path, writer_cls):
    names = ["it's", "x" * 31, "Q1 & Q2 <draft>", "Żółć", "History 2024", "a.b,c;d"]
    with writer_cls(tmp_path / "ok.bin") as writer:
        for name in names:
            writer.add_sheet(name)
            writer.write_sheet([[name]])


def test_sheet_name_must_be_a_string(tmp_path):
    with XlsxWriter(tmp_path / "x.xlsx") as writer:
        with pytest.raises(TypeError, match="sheet name must be a string"):
            writer.add_sheet(5)
        writer.add_sheet("Ok")


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
