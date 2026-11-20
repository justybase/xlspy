"""XLSX and XLSB updaters: one operation, the same observable result.

The two formats encode cells differently, so the comparison is made after a
round trip through ``load_workbook``: values, styles, notes, links, names,
protection and the used region must agree.  Array and spill ranges are compared
in the package itself (``<dimension>`` / BrtWsDim and the auto-filter range),
because the object model sizes a sheet from its cells while XLSB also stores
the array's member cells.
"""

from __future__ import annotations

import datetime as dt
import re
import struct
import zipfile
from pathlib import Path

import pytest

from xlspy import (
    ArrayFormula,
    Chart,
    ChartSeries,
    Comment,
    DataTableFormula,
    DataValidation,
    DefinedName,
    ExcelReader,
    ExcelTable,
    Font,
    Formula,
    Hyperlink,
    SheetProtection,
    Style,
    UnsupportedFeatureError,
    Workbook,
    WorkbookProtection,
    XlsbUpdater,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)
from xlspy.biff_utils import iter_records
from xlspy.refs import column_label

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}

BASE_ROWS = [
    [1, "alpha", dt.date(2024, 1, 2), 1.5, True],
    [2, "beta", dt.datetime(2024, 3, 4, 5, 6, 7), -2, False],
]
BASE_HEADERS = ["number", "text", "date", "amount", "flag"]


def _source(directory: Path, ext: str, *, shared: bool = True, table: bool = False) -> Path:
    path = directory / f"source.{ext}"
    writer = FORMATS[ext][0]
    styles = {
        "A:A": Style(number_format="0.00"),
        "C:C": Style(number_format="yyyy-mm-dd"),
    }
    with writer(str(path), use_shared_strings=shared) as book:
        book.add_sheet("Data")
        options = {
            "styles": styles,
            "header_style": Style(font=Font(bold=True)),
            "auto_filter": not table,
        }
        if table:
            options["tables"] = [ExcelTable("Tbl", "A1:C3")]
        book.write_sheet(
            [["h1", "h2", "h3"], [1, "a", dt.date(2024, 1, 2)],
             [2, "b", dt.date(2024, 1, 3)]],
            **options,
        )
        book.add_sheet("Other")
        book.write_sheet([[9]], auto_filter=False)
    return path


def _apply(path: Path, ext: str, change) -> None:
    updater = FORMATS[ext][1](str(path))
    try:
        change(updater)
        updater.save(str(path))
    finally:
        updater.close()


def _replace(updater, stream: bool, rows, headers=None, **options) -> None:
    method = updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
    method("Data", iter(rows) if stream else rows, headers=headers, **options)


def _value_key(value):
    if isinstance(value, ArrayFormula):
        return ("ArrayFormula", value.ref, _name_formula_key(value.text))
    if isinstance(value, Formula):
        return ("Formula", _name_formula_key(value.text), value.value, value.array_ref, value.spill_ref)
    if isinstance(value, dt.datetime):
        return ("datetime", value.isoformat())
    if isinstance(value, dt.date):
        return ("date", value.isoformat())
    if isinstance(value, str) and value.startswith("="):
        return _name_formula_key(value)
    return value


def _font_key(font):
    if font is None:
        return None
    return (
        font.bold,
        font.italic,
        font.underline,
        font.size or 11.0,
        font.name,
        font.color.rgb if font.color is not None else None,
    )


def _fill_key(fill):
    if fill is None:
        return None
    return (
        fill.fill_type,
        fill.start_color.rgb if fill.start_color is not None else None,
    )


def _name_formula_key(formula: str) -> str:
    """A defined name's formula without optional quotes around a simple sheet name."""

    return re.sub(r"'([^'!]+)'!", r"\1!", formula)


def _protection_key(protection):
    return (
        protection.sheet,
        protection.selectLockedCells,
        protection.selectUnlockedCells,
        protection.password,
    )


def snapshot(path: Path) -> dict:
    workbook = load_workbook(str(path))
    try:
        result = {
            "sheets": workbook.get_sheet_names(),
            "active": workbook.index(workbook.active.title),
            "names": [
                (
                    name.name,
                    _name_formula_key(name.formula),
                    name.scope,
                    name.hidden,
                )
                for name in workbook.defined_names
            ],
            "sheets_data": {},
        }
        for name in result["sheets"]:
            sheet = workbook[name]
            cells = {}
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None and not cell.has_style:
                        continue
                    cells[cell.coordinate] = (
                        _value_key(cell.value),
                        cell.number_format,
                        _font_key(cell.font),
                        _fill_key(cell.fill),
                        cell.comment.text if cell.comment is not None else None,
                        cell.comment.author if cell.comment is not None else None,
                        (
                            _name_formula_key(cell.hyperlink.target)
                            if cell.hyperlink is not None else None
                        ),
                        cell.hyperlink.tooltip if cell.hyperlink is not None else None,
                    )
            result["sheets_data"][name] = {
                "dimensions": sheet.dimensions,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "cells": cells,
                "merged": sorted(str(value) for value in sheet.merged_cells.ranges),
                "tables": [
                    (table.displayName, table.ref, table.headerRowCount)
                    for table in sheet.tables
                ],
                "auto_filter": (
                    sheet.auto_filter.ref if sheet.auto_filter is not None else None
                ),
                "state": sheet.sheet_state,
                "protection": _protection_key(sheet.protection),
                "validations": [
                    (
                        rule.type,
                        str(rule.sqref),
                        _name_formula_key(str(rule.formula1)),
                        (
                            None if rule.formula2 is None
                            else _name_formula_key(str(rule.formula2))
                        ),
                        rule.operator,
                        rule.allow_blank,
                        rule.show_dropdown,
                        rule.show_input_message,
                        rule.prompt_title,
                        rule.prompt,
                        rule.show_error_message,
                        rule.error_title,
                        rule.error,
                        rule.error_style,
                        rule.ime_mode,
                    )
                    for rule in sheet.data_validations.dataValidation
                ],
            }
        return result
    finally:
        workbook.close()


def _region(path: Path, ext: str) -> tuple[str | None, str | None]:
    """The sheet's dimension and auto-filter ref, read from the package."""

    with zipfile.ZipFile(path) as archive:
        if ext == "xlsx":
            xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            dimension = re.search(
                r'<(?:\w+:)?dimension\b[^>]*\bref="([^"]*)"', xml
            )
            auto_filter = re.search(
                r'<(?:\w+:)?autoFilter\b[^>]*\bref="([^"]*)"', xml
            )
            return (
                dimension.group(1) if dimension else None,
                auto_filter.group(1) if auto_filter else None,
            )
        data = archive.read("xl/worksheets/sheet1.bin")

    dimension = auto_filter = None
    for record in iter_records(data):
        if record.record_id == 0x94 and dimension is None and record.length >= 16:
            first_row, last_row, first_col, last_col = struct.unpack_from(
                "<iiii", data, record.data_start
            )
            dimension = (
                f"{column_label(first_col)}{first_row + 1}:"
                f"{column_label(last_col)}{last_row + 1}"
            )
        elif record.record_id == 0xA1 and auto_filter is None and record.length >= 16:
            first_row, last_row, first_col, last_col = struct.unpack_from(
                "<iiii", data, record.data_start
            )
            auto_filter = (
                f"{column_label(first_col)}{first_row + 1}:"
                f"{column_label(last_col)}{last_row + 1}"
            )
    return dimension, auto_filter


def _run_both(tmp_path: Path, *changes, shared: bool = True, table: bool = False):
    """Build one file per format, apply ``changes`` to each, return snapshots."""

    results = {}
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext, shared=shared, table=table)
        for change in changes:
            _apply(path, ext, change)
        results[ext] = snapshot(path)
    return results["xlsx"], results["xlsb"]


def _run_both_regions(tmp_path: Path, *changes):
    results = {}
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        for change in changes:
            _apply(path, ext, change)
        results[ext] = (_region(path, ext), path)
    return results["xlsx"], results["xlsb"]


# --- replace_sheet_data ------------------------------------------------------


@pytest.mark.parametrize("stream", [False, True])
@pytest.mark.parametrize("shared", [False, True])
@pytest.mark.parametrize("headers", [True, False])
def test_replacing_data_has_the_same_result(tmp_path, stream, shared, headers):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: _replace(
            updater, stream, BASE_ROWS, BASE_HEADERS if headers else None
        ),
        shared=shared,
    )
    assert xlsx == xlsb


@pytest.mark.parametrize("stream", [False, True])
def test_a_none_header_becomes_the_same_empty_cell(tmp_path, stream):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: _replace(
            updater, stream, BASE_ROWS, ["a", None, "c", "d", "e"]
        ),
    )
    assert xlsx == xlsb
    header = xlsx["sheets_data"]["Data"]["cells"]["B1"]
    assert header[0] == ""


@pytest.mark.parametrize("stream", [False, True])
@pytest.mark.parametrize(
    "rows",
    [
        [],
        [[None], [None]],
        [[None, None, None]],
        BASE_ROWS + [[None, None, None], [None]],
        [[1, 2, 3], [None, None, None], [4, 5, 6]],
    ],
    ids=["empty", "trailing-rows", "blank-row", "trailing-blanks", "blank-middle"],
)
def test_empty_and_blank_rows_have_the_same_result(tmp_path, stream, rows):
    xlsx, xlsb = _run_both(
        tmp_path, lambda updater: _replace(updater, stream, rows)
    )
    assert xlsx == xlsb


def test_table_headers_are_guarded_the_same(tmp_path):
    messages = {}
    rows = [[1, "a", dt.date(2024, 1, 2)], [2, "b", dt.date(2024, 1, 3)]]
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext, table=True)
        with pytest.raises(UnsupportedFeatureError) as caught:
            _apply(
                path,
                ext,
                lambda updater: _replace(
                    updater, False, rows, ["wrong", "h2", "h3"]
                ),
            )
        messages[ext] = str(caught.value)
    assert "table 'Tbl'" in messages["xlsx"]
    assert messages["xlsx"] == messages["xlsb"]


# --- formulas ----------------------------------------------------------------


@pytest.mark.parametrize("stream", [False, True])
def test_scalar_formulas_and_cached_values_match(tmp_path, stream):
    rows = [
        [1, Formula("A2*2", 2), '"x"&"y"'],
        [2, Formula("A3*2", 4), Formula('"x"&"y"', "xy")],
    ]
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: _replace(
            updater, stream, rows, ["n", "double", "text"]
        ),
    )
    assert xlsx == xlsb
    with ExcelReader(xlsx_path(tmp_path)) as reader:
        formulas = reader.read_all("Data", data_only=False)
        values = reader.read_all("Data", data_only=True)
    assert formulas[1][:2] == [1, "=A2*2"]
    assert formulas[2][:2] == [2, "=A3*2"]
    assert formulas[1][2].lstrip("=") == '"x"&"y"'
    assert formulas[2][2].lstrip("=") == '"x"&"y"'
    assert values[1][1] == 2 and values[2][1] == 4


def xlsx_path(tmp_path: Path) -> Path:
    return tmp_path / "xlsx" / "source.xlsx"


@pytest.mark.parametrize("stream", [False, True])
@pytest.mark.parametrize(
    "formula, expected",
    [
        (Formula("SEQUENCE(2)", spill_ref="A2:A3"), "A1:C3"),
        (Formula("ROW(A2:A3)", array_ref="A2:A3"), "A1:C3"),
        (ArrayFormula("A2:B4", "ROW(A2:B4)"), "A1:C4"),
        (Formula("SEQUENCE(3,2)", spill_ref="A2:B4"), "A1:C4"),
    ],
    ids=["spill-one-column", "array-one-column", "array-two-columns", "spill-wide"],
)
def test_array_and_spill_ranges_expand_both_regions(
    tmp_path, stream, formula, expected
):
    xlsx, xlsb = _run_both_regions(
        tmp_path,
        lambda updater: _replace(
            updater, stream, [[formula, 1, 2]], ["a", "b", "c"]
        ),
    )
    assert xlsx[0] == xlsb[0] == (expected, expected)
    for _region_value, path in (xlsx, xlsb):
        with ExcelReader(str(path)) as reader:
            formulas = reader.read_all("Data", data_only=False)
        assert formulas[1][0] == f"={formula.text}"


@pytest.mark.parametrize("stream", [False, True])
@pytest.mark.parametrize(
    "formula",
    [
        Formula("1", array_ref="B2:B3"),
        ArrayFormula("B2:B3", "1"),
        Formula("1", spill_ref="B2:B3"),
    ],
    ids=["scalar-array", "array-formula", "spill"],
)
def test_a_formula_outside_its_top_left_cell_is_refused(tmp_path, stream, formula):
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with pytest.raises(ValueError, match="top-left cell"):
            _apply(
                path,
                ext,
                lambda updater: _replace(updater, stream, [[formula]], ["value"]),
            )


# --- notes, links, names, protection ----------------------------------------


def test_notes_and_links_match_across_formats(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_comments(
            "Data", {"A1": Comment("note", "Jan")}
        ),
        lambda updater: updater.set_hyperlinks(
            "Data", {"B1": Hyperlink("https://example.com/")}
        ),
    )
    assert xlsx == xlsb
    data = xlsx["sheets_data"]["Data"]["cells"]
    assert data["A1"][4:] == ("note", "Jan", None, None)
    assert data["B1"][6:] == ("https://example.com/", None)


def test_notes_and_links_can_be_removed_again(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_comments(
            "Data", {"A1": Comment("note", "Jan")}
        ),
        lambda updater: updater.set_hyperlinks(
            "Data", {"B1": Hyperlink("https://example.com/")}
        ),
        lambda updater: updater.set_comments("Data", {"A1": None}),
        lambda updater: updater.set_hyperlinks("Data", {"B1": None}),
    )
    assert xlsx == xlsb
    data = xlsx["sheets_data"]["Data"]["cells"]
    assert data["A1"][4] is None and data["A1"][6] is None
    assert data["B1"][4] is None and data["B1"][6] is None


def test_defined_names_match_across_formats(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_defined_name(
            DefinedName("Abc", "Data!$A$1", localSheetId=0)
        ),
        lambda updater: updater.set_defined_name(
            DefinedName("Global", "Data!$B$2", hidden=True)
        ),
    )
    assert xlsx == xlsb
    assert set(xlsx["names"]) == {
        ("Abc", "Data!$A$1", "Data", None),
        ("Global", "Data!$B$2", None, True),
    }


def test_defined_names_can_be_removed_again(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_defined_name(
            DefinedName("Abc", "Data!$A$1", localSheetId=0)
        ),
        lambda updater: updater.set_defined_name(
            DefinedName("Global", "Data!$B$2")
        ),
        lambda updater: updater.remove_defined_names(
            [("Abc", "Data"), ("Global", None)]
        ),
    )
    assert xlsx == xlsb
    assert xlsx["names"] == []


def _has_workbook_protection(path: Path, ext: str) -> bool:
    with zipfile.ZipFile(path) as archive:
        if ext == "xlsx":
            return b"<workbookProtection" in archive.read("xl/workbook.xml")
        data = archive.read("xl/workbook.bin")
    return any(record.record_id == 0x216 for record in iter_records(data))


def test_sheet_and_workbook_protection_match(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_sheet_protection(
            "Data",
            SheetProtection(
                enabled=True,
                select_locked_cells=False,
                password="abc",
            ),
        ),
        lambda updater: updater.set_workbook_protection(
            WorkbookProtection(password="abc")
        ),
    )
    assert xlsx == xlsb
    protection = xlsx["sheets_data"]["Data"]["protection"]
    assert protection[0] is True and protection[3] is not None
    for ext in FORMATS:
        assert _has_workbook_protection(tmp_path / ext / f"source.{ext}", ext)


# --- update_cells ------------------------------------------------------------


def test_update_cells_values_match(tmp_path):
    cells = {
        "A1": 7,
        "B2": "z",
        "C3": None,
        "D4": True,
        "E5": dt.date(2024, 6, 7),
        "F6": Formula("1+1", 2),
        "Z9": "far",
    }
    xlsx, xlsb = _run_both(
        tmp_path, lambda updater: updater.update_cells("Data", cells)
    )
    assert xlsx == xlsb
    data = xlsx["sheets_data"]["Data"]
    assert data["max_row"] == 9 and data["max_column"] == 26
    assert data["cells"]["A1"][0] == 7
    assert data["cells"]["B2"][0] == "z"
    assert data["cells"]["D4"][0] is True
    assert data["cells"]["E5"][0] == 45450  # no date style on column E
    assert data["cells"]["F6"][0] == "=1+1"
    assert data["cells"]["Z9"][0] == "far"


def test_update_cells_styles_match(tmp_path):
    # XF 7 is the "0.00" style of column A in _source.
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.update_cells(
            "Data", {"D4": "styled"}, styles={"A2": 7, "D4": 7, "G10": 7}
        ),
    )
    assert xlsx == xlsb
    data = xlsx["sheets_data"]["Data"]
    assert data["cells"]["A2"][0] == 1
    assert data["cells"]["A2"][1] == "0.00"
    assert data["cells"]["D4"][0] == "styled"
    assert data["cells"]["D4"][1] == "0.00"
    assert data["cells"]["G10"][0] is None
    assert data["cells"]["G10"][1] == "0.00"
    assert data["max_row"] == 10 and data["max_column"] == 7


def test_an_invalid_cell_style_is_refused_in_both_formats(tmp_path):
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with pytest.raises(ValueError, match="invalid cell style index"):
            _apply(
                path,
                ext,
                lambda updater: updater.update_cells(
                    "Data", {}, styles={"A1": -1}
                ),
            )


# --- merged cells and data validations ---------------------------------------


def test_merged_ranges_match(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_merged_cells("Data", ["A1:B1", "C3:D4"]),
    )
    assert xlsx == xlsb
    assert xlsx["sheets_data"]["Data"]["merged"] == ["A1:B1", "C3:D4"]


def test_merged_ranges_can_be_removed_again(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_merged_cells("Data", ["A1:B1", "C3:D4"]),
        lambda updater: updater.set_merged_cells("Data", []),
    )
    assert xlsx == xlsb
    assert xlsx["sheets_data"]["Data"]["merged"] == []


def test_overlapping_merged_ranges_are_refused_in_both(tmp_path):
    messages = {}
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with pytest.raises(ValueError) as caught:
            _apply(
                path,
                ext,
                lambda updater: updater.set_merged_cells(
                    "Data", ["A1:B2", "B2:C3"]
                ),
            )
        messages[ext] = str(caught.value)
    assert messages["xlsx"] == messages["xlsb"]


def test_data_validations_match(tmp_path):
    rules = [
        DataValidation("list", "A2:A9", formula1='"x,y"', allow_blank=False),
        DataValidation(
            "whole",
            "B2:B9",
            formula1=1,
            formula2=10,
            operator="not_between",
            error_style="warning",
            error_title="Out of range",
            error="Enter 1..10",
        ),
        DataValidation(
            "date",
            "C2:C9",
            formula1=dt.date(2024, 1, 1),
            operator="greater_than",
            show_input_message=True,
            prompt_title="Date",
            prompt="Enter a date",
        ),
    ]
    xlsx, xlsb = _run_both(
        tmp_path, lambda updater: updater.set_data_validations("Data", rules)
    )
    assert xlsx == xlsb
    assert len(xlsx["sheets_data"]["Data"]["validations"]) == 3


def test_data_validations_can_be_removed_again(tmp_path):
    rule = DataValidation("list", "A2:A9", formula1='"x,y"')
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_data_validations("Data", [rule]),
        lambda updater: updater.set_data_validations("Data", []),
    )
    assert xlsx == xlsb
    assert xlsx["sheets_data"]["Data"]["validations"] == []


def test_data_validations_are_type_checked_in_both(tmp_path):
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with pytest.raises(TypeError, match="DataValidation"):
            _apply(
                path,
                ext,
                lambda updater: updater.set_data_validations(
                    "Data", [object()]  # type: ignore[list-item]
                ),
            )


# --- preservation and repetition --------------------------------------------


@pytest.mark.parametrize("stream", [False, True])
def test_the_other_sheet_and_package_members_are_preserved(tmp_path, stream):
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with zipfile.ZipFile(path) as archive:
            before_members = set(archive.namelist())
        before_other = snapshot(path)["sheets_data"]["Other"]
        _apply(path, ext, lambda updater: _replace(updater, stream, BASE_ROWS))
        with zipfile.ZipFile(path) as archive:
            after_members = set(archive.namelist())
        assert after_members == before_members
        assert snapshot(path)["sheets_data"]["Other"] == before_other


def test_repeated_replacements_match(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: _replace(
            updater, False, BASE_ROWS, ["a", "b", "c", "d", "e"]
        ),
        lambda updater: _replace(
            updater, True, BASE_ROWS[:1], ["a", "b", "c", "d", "e"]
        ),
    )
    assert xlsx == xlsb
    data = xlsx["sheets_data"]["Data"]
    assert data["dimensions"] == "A1:E2"
    assert data["cells"]["A2"][0] == 1


# --- errors ------------------------------------------------------------------


def _error(tmp_path: Path, ext: str, change) -> tuple[type[BaseException], str]:
    directory = tmp_path / f"error-{ext}"
    directory.mkdir()
    path = _source(directory, ext)
    with pytest.raises(BaseException) as caught:
        _apply(path, ext, change)
    return type(caught.value), str(caught.value)


@pytest.mark.parametrize(
    "change, error",
    [
        (lambda u: u.replace_sheet_data("Nope", []), KeyError),
        (
            lambda u: u.replace_sheet_data("Data", [], style_fallback="wrong"),
            ValueError,
        ),
        (lambda u: u.replace_sheet_data("Data", [[float("nan")]]), ValueError),
        (lambda u: u.replace_sheet_data("Data", [[float("inf")]]), ValueError),
        (lambda u: u.replace_sheet_data("Data", ["text"]), TypeError),
        (lambda u: u.replace_sheet_data("Data", [["x" * 40000]]), ValueError),
    ],
)
def test_the_same_errors_have_the_same_type(tmp_path, change, error):
    for ext in FORMATS:
        kind, _message = _error(tmp_path, ext, change)
        assert kind is error


@pytest.mark.parametrize("stream", [False, True])
def test_a_value_past_column_xfd_is_refused_in_both_formats(tmp_path, stream):
    for ext in FORMATS:
        directory = tmp_path / ext
        directory.mkdir()
        path = _source(directory, ext)
        with pytest.raises(ValueError, match="16384"):
            _apply(
                path,
                ext,
                lambda updater: _replace(
                    updater, stream, [[None] * 16384 + [1]]
                ),
            )


# --- updater API parity ------------------------------------------------------


def _xlsx_with_chartsheet(tmp_path: Path) -> Path:
    path = tmp_path / "charts.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Arkusz1"
    sheet.append(["a"])
    sheet.append([1])
    chart = workbook.create_chartsheet("ChartOnly")
    chart.add_chart(Chart("bar", [ChartSeries("A1:A2")]), source=sheet)
    workbook.save(str(path))
    workbook.close()
    return path


def test_removed_sheets_match(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.add_sheet("Temp", 1),
        lambda updater: updater.remove_sheet("Temp"),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["Data", "Other"]


def test_remove_the_active_sheet_selects_the_next_visible_one(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_sheet_state("Other", "hidden"),
        lambda updater: updater.set_active_sheet(0),
        lambda updater: updater.remove_sheet("Data"),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["Other"]
    assert xlsx["active"] == 0


def test_local_names_of_a_removed_sheet_are_dropped(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_defined_name(
            DefinedName("LocalData", "Data!$A$1", localSheetId=0)
        ),
        lambda updater: updater.set_defined_name(
            DefinedName("LocalOther", "Other!$A$1", localSheetId=1)
        ),
        lambda updater: updater.remove_sheet("Data"),
    )
    assert xlsx == xlsb
    assert [name for name, *_rest in xlsx["names"]] == ["LocalOther"]


@pytest.mark.parametrize(
    "change, error",
    [
        (lambda updater: updater.remove_sheet("Nope"), KeyError),
        (
            lambda updater: (
                updater.remove_sheet("Data"),
                updater.remove_sheet("Other"),
            ),
            ValueError,
        ),
    ],
)
def test_remove_sheet_errors_match(tmp_path, change, error):
    for ext in FORMATS:
        kind, _message = _error(tmp_path, ext, change)
        assert kind is error


def test_renamed_sheet_rewrites_formulas_names_validation_and_links(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_defined_name(
            DefinedName("RefData", "Data!$A$1")
        ),
        lambda updater: updater.update_cells(
            "Other", {"A2": Formula("Data!A1", value=1)}
        ),
        lambda updater: updater.set_data_validations(
            "Other",
            [DataValidation("list", "B2", formula1="Data!$A$1")],
        ),
        lambda updater: updater.update_cells("Other", {"C2": "link"}),
        lambda updater: updater.set_hyperlinks(
            "Other", {"C2": Hyperlink("#Data!A1")}
        ),
        lambda updater: updater.rename_sheet("Data", "Renamed"),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["Renamed", "Other"]
    assert xlsx["names"][0][0] == "RefData"
    assert xlsx["names"][0][1] == "Renamed!$A$1"
    other = xlsx["sheets_data"]["Other"]
    assert other["cells"]["A2"][0] == "=Renamed!A1"
    assert other["validations"][0][2] == "Renamed!$A$1"
    assert other["cells"]["C2"][6] == "#Renamed!A1"


def test_rename_missing_sheet_errors_match(tmp_path):
    for ext in FORMATS:
        kind, _message = _error(
            tmp_path, ext, lambda updater: updater.rename_sheet("Nope", "New")
        )
        assert kind is KeyError


def test_rename_to_an_existing_sheet_errors_match(tmp_path):
    for ext in FORMATS:
        kind, _message = _error(
            tmp_path, ext, lambda updater: updater.rename_sheet("Data", "Other")
        )
        assert kind is ValueError


def test_added_sheets_land_in_the_same_order(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.add_sheet("First", 0),
        lambda updater: updater.add_sheet("Last"),
        lambda updater: updater.add_sheet("Middle", 1, state="hidden"),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["First", "Middle", "Data", "Other", "Last"]
    assert xlsx["sheets_data"]["Middle"]["state"] == "hidden"


@pytest.mark.parametrize(
    "change, error",
    [
        (lambda updater: updater.add_sheet("Data"), ValueError),
        (lambda updater: updater.add_sheet("bad/name"), ValueError),
        (lambda updater: updater.add_sheet("Hidden", state="bogus"), ValueError),
        (lambda updater: updater.add_sheet("History"), ValueError),
    ],
)
def test_add_sheet_errors_match(tmp_path, change, error):
    for ext in FORMATS:
        kind, _message = _error(tmp_path, ext, change)
        assert kind is error


def test_added_sheet_is_active_after_an_insertion_before_it(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_active_sheet(1),
        lambda updater: updater.add_sheet("Zero", 0),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["Zero", "Data", "Other"]
    assert xlsx["active"] == 2


def test_sheet_states_and_active_survive_a_reorder(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_sheet_state("Other", "hidden"),
        lambda updater: updater.set_active_sheet(0),
        lambda updater: updater.reorder_sheets(["Other", "Data"]),
    )
    assert xlsx == xlsb
    assert xlsx["sheets"] == ["Other", "Data"]
    assert xlsx["active"] == 1
    assert xlsx["sheets_data"]["Other"]["state"] == "hidden"


def test_reorder_leaves_a_visible_active_sheet(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_sheet_state("Data", "hidden"),
        lambda updater: updater.set_active_sheet(0),
        lambda updater: updater.reorder_sheets(["Other", "Data"]),
    )
    assert xlsx == xlsb
    assert xlsx["active"] == 0
    assert xlsx["sheets"] == ["Other", "Data"]


def test_local_names_follow_a_reorder(tmp_path):
    xlsx, xlsb = _run_both(
        tmp_path,
        lambda updater: updater.set_defined_name(
            DefinedName("LocalData", "Data!$A$1", localSheetId=0)
        ),
        lambda updater: updater.set_defined_name(
            DefinedName("LocalOther", "Other!$A$1", localSheetId=1)
        ),
        lambda updater: updater.reorder_sheets(["Other", "Data"]),
    )
    assert xlsx == xlsb
    scopes = {name: scope for name, _formula, scope, _hidden in xlsx["names"]}
    assert scopes == {"LocalData": "Data", "LocalOther": "Other"}


@pytest.mark.parametrize(
    "change, error",
    [
        (lambda updater: updater.set_sheet_state("Nope", "hidden"), KeyError),
        (lambda updater: updater.set_sheet_state("Data", "bogus"), ValueError),
        (lambda updater: updater.reorder_sheets(["Data"]), ValueError),
        (lambda updater: updater.reorder_sheets(["Data", "Data"]), ValueError),
        (lambda updater: updater.set_active_sheet(5), IndexError),
    ],
)
def test_sheet_management_errors_match(tmp_path, change, error):
    for ext in FORMATS:
        kind, _message = _error(tmp_path, ext, change)
        assert kind is error


def test_xlsx_sheet_management_includes_chartsheets(tmp_path):
    path = _xlsx_with_chartsheet(tmp_path)
    updater = XlsxUpdater(str(path))
    try:
        assert updater.get_sheet_names() == ["Arkusz1", "ChartOnly"]
        assert updater.get_worksheet_names() == ["Arkusz1"]
        assert updater.get_chartsheet_names() == ["ChartOnly"]
        assert updater.get_sheet_states() == {
            "Arkusz1": "visible",
            "ChartOnly": "visible",
        }
        updater.set_active_sheet(1)
        updater.reorder_sheets(["ChartOnly", "Arkusz1"])
        updater.set_sheet_state("ChartOnly", "hidden")
        updater.set_active_sheet(1)
        assert updater.get_sheet_names() == ["ChartOnly", "Arkusz1"]
        assert updater.get_active_sheet() == 1
        assert updater.get_sheet_states() == {
            "ChartOnly": "hidden",
            "Arkusz1": "visible",
        }
        updater.save(str(path))
    finally:
        updater.close()

    reopened = XlsxUpdater(str(path))
    try:
        assert reopened.get_sheet_names() == ["ChartOnly", "Arkusz1"]
        assert reopened.get_active_sheet() == 1
        assert reopened.get_chartsheet_names() == ["ChartOnly"]
        assert reopened.get_sheet_states()["ChartOnly"] == "hidden"
    finally:
        reopened.close()


def test_formula_range_bounds_understands_every_multi_cell_formula():
    from xlspy.updater_utils import formula_range_bounds

    assert formula_range_bounds(ArrayFormula("B2:C3", "1")) == (1, 1, 2, 2)
    assert formula_range_bounds(DataTableFormula("B2:C3")) == (1, 1, 2, 2)
    assert formula_range_bounds(Formula("1", array_ref="B2:C3")) == (1, 1, 2, 2)
    assert formula_range_bounds(Formula("1", spill_ref="B2:C3")) == (1, 1, 2, 2)
    assert formula_range_bounds(Formula("1")) is None
    assert formula_range_bounds("plain") is None


def test_the_shared_updater_hooks_are_abstract():
    from xlspy._updater_base import _UpdaterBase

    base = _UpdaterBase()
    with pytest.raises(NotImplementedError):
        base._ensure_stream_shared_strings()
    with pytest.raises(NotImplementedError):
        base._replace_sheet_data_stream_impl("S", [])
    with pytest.raises(NotImplementedError):
        base._replace_loaded_rows("S", "sheet.xml", [], None, "inherit")
    with pytest.raises(NotImplementedError):
        base._read_table_header(b"")
    with pytest.raises(NotImplementedError):
        base._cell_xf_number_formats()


def test_get_worksheet_and_chartsheet_names_match_the_xlsb_updater(tmp_path):
    path = _xlsx_with_chartsheet(tmp_path)
    updater = XlsxUpdater(str(path))
    try:
        assert updater.get_sheet_names() == ["Arkusz1", "ChartOnly"]
        assert updater.get_worksheet_names() == ["Arkusz1"]
        assert updater.get_chartsheet_names() == ["ChartOnly"]
    finally:
        updater.close()

    fixture = Path(__file__).parent / "data" / "chartsheet.xlsb"
    updater = XlsbUpdater(str(fixture))
    try:
        assert updater.get_sheet_names() == ["ChartOnly", "Arkusz1"]
        assert updater.get_worksheet_names() == ["Arkusz1"]
        assert updater.get_chartsheet_names() == ["ChartOnly"]
    finally:
        updater.close()
