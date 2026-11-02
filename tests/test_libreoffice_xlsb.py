"""LibreOffice Calc as a second consumer of the .xlsb files xlspy writes.

Excel tolerates records other readers do not (PR #1: with the wrong column-info container ids,
LibreOffice kept the sheets but dropped every cell). Each workbook below is converted to .xlsx by
headless LibreOffice in one run, and the result is read back with openpyxl, an independent reader.
Skipped when LibreOffice (soffice) is not installed.
"""
import datetime
import os
import shutil
import subprocess

import pytest

from xlspy import (
    CellRichText, Comment, DataValidation, DefinedName, ExcelTable, Font, Formula, Group, Hyperlink,
    InlineFont, Style, TextBlock, XlsbUpdater, XlsbWriter, load_workbook,
)
from xlspy.styles import Border, CellIsRule, DifferentialStyle, Side

openpyxl = pytest.importorskip("openpyxl")

pytestmark = pytest.mark.libreoffice


def _find_soffice():
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for root in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if root:
            candidate = os.path.join(root, "LibreOffice", "program", "soffice.exe")
            if os.path.isfile(candidate):
                return candidate
    return None


SOFFICE = _find_soffice()

TYPES_ROWS = [
    ["text", 42, 3.25, True, datetime.date(2026, 11, 2), datetime.datetime(2026, 11, 2, 13, 45, 30),
     -7, 10**12, "zażółć ✓", False],
    [(0.125, "0.00%"), (1234.5, "#,##0.00"), "", None, 0, -0.5],
]
BIG_ROWS = 30000


def _rich_workbook(path, *, shared_strings=True):
    """One sheet per feature; together they cover what a report download uses."""
    with XlsbWriter(path, use_shared_strings=shared_strings) as writer:
        writer.add_sheet("Types")
        writer.write_sheet(TYPES_ROWS)
        writer.add_sheet("Formulas")
        writer.write_sheet([
            [2, 3, Formula("A1*B1", 6), Formula("SUM(A1:B1)", 5), Formula('IF(A1>1,"big","small")', "big"),
             Formula("Types!B1+1", 43), Formula("1/0", "#DIV/0!"), Formula("MyRange", 42)],
            # The second value is the cached result of the array's other cell.
            [Formula("A1:B1*2", 4, array_ref="A2:B2"), 6],
        ])
        writer.add_sheet("Layout")
        writer.write_sheet(
            [["a", "b", "c", "d"], [1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12], ["merged", None, 13, 14]],
            column_widths={"A": 25, "D": 8}, hidden_columns=["C"], row_heights={2: 30}, hidden_rows=[3],
            merges=["A5:B5"], freeze_panes="B2",
        )
        writer.add_sheet("Outline")
        # Group indexes are 0-based: rows 4-6 and column C.
        writer.write_sheet([[f"r{i}", i, i * 2] for i in range(8)],
                           row_groups=[Group(3, 5)], column_groups=[Group(2, 2)])
        writer.add_sheet("Sparse")
        writer.write_sheet([[], [None, None], [None, None, "C3"], [], [None, None, None, None, 5]])
        writer.add_sheet("Styles")
        writer.write_sheet(
            # G1 and H1: formatted cells without a value.
            [["bold", "fill", "red", "border", "center", 0.5, (None, "0.00"), (None, Style(bold=True))]],
            styles={
                "A1": Style(bold=True, italic=True), "B1": Style(fill="#FFFF00"), "C1": Style(font_color="#FF0000"),
                "D1": Style(border=Border(bottom=Side("thin", "FF000000"))),
                "E1": Style(horizontal="center", wrap=True), "F1": Style(number_format="0.0%", bold=True),
            },
        )
        writer.add_sheet("Links")
        writer.write_sheet([[Hyperlink("https://example.com/", "site"), Hyperlink("#Types!A1", "internal"),
                             CellRichText(TextBlock(InlineFont(b=True), "bold "), "plain")]])
        writer.add_sheet("Extras")
        writer.write_sheet(
            [["h1", "h2"], [1, "x"], [7, "y"]],
            comments={"A1": Comment("a note", "Tester")},
            data_validations=[DataValidation("list", "B2:B3", formula1='"x,y,z"')],
            conditional_formats=[CellIsRule("A2:A3", "greaterThan", ["5"],
                                            dxf=DifferentialStyle(font=Font(color="FF9C0006", bold=True)))],
            tables=[ExcelTable("Tbl", "A1:B3")],
        )
        writer.add_sheet("View")
        writer.write_sheet([["v"]], tab_color="#00AA00", zoom=150, show_grid_lines=False)
        writer.add_sheet("Big")
        writer.write_sheet([f"row {i}", i, i * 0.5] for i in range(BIG_ROWS))
        writer.add_sheet("Hidden", hidden=True)
        writer.write_sheet([["secret"]])
        writer.set_defined_name(DefinedName("MyRange", "Types!$B$1"))


def _updated_workbook(source, path):
    with XlsbWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["name", "value"], ["a", 1]], column_widths={"A": 20}, hidden_columns=["C"])
    updater = XlsbUpdater(source)
    try:
        updater.update_cells("Data", {"B2": 10, "A3": "b", "B3": 20, "D6": "far"})
        updater.save(path)
    finally:
        updater.close()


def _model_workbook(source, path):
    with XlsbWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["name", "value"], ["a", 1]])
    workbook = load_workbook(source)
    try:
        sheet = workbook["Data"]
        sheet["B2"] = 11
        sheet["C4"] = "added"
        sheet.column_dimensions["A"].width = 30
        sheet.column_dimensions["B"].hidden = True
        workbook.create_sheet("New")["D5"] = 5
        workbook.save(path)
    finally:
        workbook.close()


@pytest.fixture(scope="module")
def calc(tmp_path_factory):
    """name -> (formulas workbook, cached-values workbook) as LibreOffice saved them."""
    if SOFFICE is None:
        pytest.skip("LibreOffice (soffice) not found")
    root = tmp_path_factory.mktemp("libreoffice")
    sources = {
        "rich": root / "rich.xlsb",
        "inline": root / "inline.xlsb",
        "updated": root / "updated.xlsb",
        "model": root / "model.xlsb",
    }
    _rich_workbook(sources["rich"])
    _rich_workbook(sources["inline"], shared_strings=False)
    _updated_workbook(root / "updated_source.xlsb", sources["updated"])
    _model_workbook(root / "model_source.xlsb", sources["model"])
    outdir = root / "out"
    # A private profile: a running LibreOffice would otherwise take the job and ignore --headless.
    profile = (root / "profile").resolve().as_uri()
    subprocess.run(
        [SOFFICE, "--headless", "--norestore", f"-env:UserInstallation={profile}",
         "--convert-to", "xlsx", "--outdir", str(outdir), *map(str, sources.values())],
        check=True, capture_output=True, timeout=300,
    )
    books = {}
    for name in sources:
        converted = outdir / f"{name}.xlsx"
        assert converted.is_file(), f"LibreOffice did not convert {name}.xlsb"
        books[name] = (openpyxl.load_workbook(converted), openpyxl.load_workbook(converted, data_only=True))
    yield sources, books
    for formulas, values in books.values():
        formulas.close()
        values.close()


def _values(sheet):
    return [[cell.value for cell in row] for row in sheet.iter_rows()]


@pytest.fixture(params=["rich", "inline"])
def book(request, calc):
    """The feature workbook, with shared strings and with inline strings."""
    return calc[1][request.param]


def test_sheets_and_visibility(book):
    formulas, _ = book
    assert formulas.sheetnames == ["Types", "Formulas", "Layout", "Outline", "Sparse", "Styles", "Links",
                                   "Extras", "View", "Big", "Hidden"]
    assert [sheet.sheet_state for sheet in formulas] == ["visible"] * 10 + ["hidden"]


def test_defined_name(book):
    formulas, _ = book
    assert formulas.defined_names["MyRange"].attr_text == "Types!$B$1"


def test_value_types(book):
    _, values = book
    first, second = _values(values["Types"])
    assert first == ["text", 42, 3.25, True, datetime.datetime(2026, 11, 2), datetime.datetime(2026, 11, 2, 13, 45, 30),
                     -7, 10**12, "zażółć ✓", False]
    # An empty string cell reads back as no value in LibreOffice.
    assert second[:6] == [0.125, 1234.5, None, None, 0, -0.5]


def test_number_formats(book):
    formulas, _ = book
    sheet = formulas["Types"]
    assert sheet["A2"].number_format == "0.00%"
    assert sheet["B2"].number_format == "#,##0.00"
    assert sheet["E1"].is_date and sheet["F1"].is_date


def test_formulas_and_cached_values(book):
    formulas, values = book
    assert [cell.value for cell in formulas["Formulas"][1]][2:8] == [
        "=A1*B1", "=SUM(A1:B1)", '=IF(A1>1,"big","small")', "=Types!B1+1", "=1/0", "=MyRange"]
    assert _values(values["Formulas"]) == [[2, 3, 6, 5, "big", 43, "#DIV/0!", 42],
                                           [4, 6, None, None, None, None, None, None]]
    array = formulas["Formulas"]["A2"].value
    assert array.ref == "A2:B2" and array.text == "=A1:B1*2"


def test_layout(book):
    """Column widths and hidden columns are the column-info records of PR #1."""
    formulas, values = book
    sheet = formulas["Layout"]
    assert _values(values["Layout"]) == [["a", "b", "c", "d"], [1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12],
                                         ["merged", None, 13, 14]]
    assert sheet.column_dimensions["A"].width == pytest.approx(25, abs=0.5)
    assert sheet.column_dimensions["C"].hidden
    assert not sheet.column_dimensions["B"].hidden
    assert sheet.row_dimensions[3].height == pytest.approx(30)
    assert sheet.row_dimensions[4].hidden
    assert [str(merged) for merged in sheet.merged_cells.ranges] == ["A5:B5"]
    assert sheet.freeze_panes == "B2"


def test_outline(book):
    formulas, values = book
    sheet = formulas["Outline"]
    assert _values(values["Outline"]) == [[f"r{i}", i, i * 2] for i in range(8)]
    assert [sheet.row_dimensions[row].outlineLevel for row in range(1, 9)] == [0, 0, 0, 1, 1, 1, 0, 0]
    assert sheet.column_dimensions["C"].outlineLevel == 1
    assert not sheet.column_dimensions["B"].outlineLevel


def test_sparse_cells_keep_their_positions(book):
    _, values = book
    sheet = values["Sparse"]
    assert sheet["C3"].value == "C3" and sheet["E5"].value == 5
    assert sheet.max_row == 5 and sheet.max_column == 5


def test_styles(book):
    formulas, _ = book
    sheet = formulas["Styles"]
    assert sheet["A1"].font.b and sheet["A1"].font.i
    assert sheet["B1"].fill.fgColor.rgb == "FFFFFF00"
    assert sheet["C1"].font.color.rgb == "FFFF0000"
    assert sheet["D1"].border.bottom.style == "thin"
    assert sheet["E1"].alignment.horizontal == "center" and sheet["E1"].alignment.wrap_text
    assert sheet["F1"].number_format == "0.0%" and sheet["F1"].font.b
    assert sheet["G1"].value is None and sheet["G1"].number_format == "0.00"
    assert sheet["H1"].value is None and sheet["H1"].font.b


def test_hyperlinks_and_rich_text(book):
    formulas, _ = book
    sheet = formulas["Links"]
    assert sheet["A1"].value == "site" and sheet["A1"].hyperlink.target == "https://example.com/"
    assert sheet["B1"].value == "internal" and sheet["B1"].hyperlink.location == "Types!A1"
    assert sheet["C1"].value == "bold plain"


def test_comment_validation_conditional_format_and_table(book):
    formulas, _ = book
    sheet = formulas["Extras"]
    assert sheet["A1"].comment.text.strip() == "a note"
    (validation,) = sheet.data_validations.dataValidation
    assert (validation.type, validation.formula1, str(validation.sqref)) == ("list", '"x,y,z"', "B2:B3")
    (rules,) = list(sheet.conditional_formatting)
    assert str(rules.sqref) == "A2:A3" and [rule.operator for rule in rules.rules] == ["greaterThan"]
    assert list(sheet.tables) == ["Tbl"] and sheet.tables["Tbl"].ref == "A1:B3"


def test_sheet_view(book):
    formulas, _ = book
    sheet = formulas["View"]
    assert sheet.sheet_properties.tabColor.rgb == "FF00AA00"
    assert sheet.sheet_view.zoomScale == 150
    assert sheet.sheet_view.showGridLines is False


def test_sheet_larger_than_the_row_buffer(book):
    _, values = book
    sheet = values["Big"]
    assert sheet.max_row == BIG_ROWS
    assert [cell.value for cell in sheet[BIG_ROWS]] == [f"row {BIG_ROWS - 1}", BIG_ROWS - 1, (BIG_ROWS - 1) * 0.5]


def _normalized(rows):
    """Rows padded to one width; LibreOffice drops empty-string cells, so "" reads as None."""
    width = max((len(row) for row in rows), default=0)
    return [[None if value == "" else value for value in row] + [None] * (width - len(row)) for row in rows]


@pytest.mark.parametrize("name", ["rich", "updated", "model"])
def test_libreoffice_reads_what_xlspy_reads(calc, name):
    """Sheet by sheet, the values LibreOffice saw are the ones xlspy reads back."""
    sources, books = calc
    _, values = books[name]
    workbook = load_workbook(sources[name], data_only=True)
    try:
        assert workbook.sheetnames == values.sheetnames
        for sheet in workbook:
            ours = [[cell.value for cell in row] for row in sheet.iter_rows()]
            assert _normalized(ours) == _normalized(_values(values[sheet.title])), sheet.title
    finally:
        workbook.close()


def test_updater_output(calc):
    _, books = calc
    formulas, values = books["updated"]
    assert _values(values["Data"]) == [["name", "value", None, None], ["a", 10, None, None],
                                       ["b", 20, None, None], [None] * 4, [None] * 4, [None, None, None, "far"]]
    assert formulas["Data"].column_dimensions["A"].width == pytest.approx(20, abs=0.5)
    assert formulas["Data"].column_dimensions["C"].hidden


def test_model_output(calc):
    _, books = calc
    formulas, values = books["model"]
    assert values.sheetnames == ["Data", "New"]
    assert _values(values["Data"]) == [["name", "value", None], ["a", 11, None], [None] * 3, [None, None, "added"]]
    assert formulas["Data"].column_dimensions["A"].width == pytest.approx(30, abs=0.5)
    assert formulas["Data"].column_dimensions["B"].hidden
    assert values["New"]["D5"].value == 5
