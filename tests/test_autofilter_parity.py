"""The default autofilter range is the same in XlsxWriter and XlsbWriter.

The rule: an explicit ``auto_filter_ref`` is written as given (unless ``auto_filter=False``).
Otherwise a filter is added only when data follows the header; its buttons sit on the last header
row, it spans the columns up to the last one holding a value (a trailing ``None`` or a formatted
cell without a value does not widen it) and the rows down to the last data row, or to the last row
of an array formula that reaches past the data.

The writers used to differ: XLSX counted the row length and ignored array rows, XLSB filtered a
header with no data and a sheet with no values.
"""
import pytest

from xlspy import ExcelTable, Formula, Group, Style, XlsbWriter, XlsxWriter, load_workbook

# name -> (rows, write_sheet options, autofilter range or None)
CASES = {
    "plain": ([["h1", "h2"], [1, 2]], {}, "A1:B2"),
    "trailing_none": ([["h1", "h2"], [1, 2, None]], {}, "A1:B2"),
    "trailing_formatted_blank": ([["h1", "h2"], [1, 2, (None, Style(bold=True))]], {}, "A1:B2"),
    "styled_empty_header_cell": ([["h1", "h2", None], [1, 2]], {"header_style": Style(italic=True)}, "A1:B2"),
    "wider_later_row": ([["h1", "h2"], [1, 2], [1, 2, 3]], {}, "A1:C3"),
    "wider_past_the_width_sample": ([["h1", "h2"]] + [[1, 2]] * 150 + [[1, 2, 3]], {}, "A1:C152"),
    "value_far_right": ([["h1"], [None, None, None, None, "x"]], {}, "A1:E2"),
    "trailing_empty_rows": ([["h1", "h2"], [1, 2], [], [None]], {}, "A1:B4"),
    "array_past_the_data": ([["h1", "h2"], [Formula("=1", 1, array_ref="A2:B4")]], {}, "A1:B4"),
    "outline_past_the_data": ([["h1", "h2"], [1, 2]], {"row_groups": [Group(4, 5)]}, "A1:B2"),
    "two_header_rows": ([["title"], ["h1", "h2"], [1, 2]], {"header_rows": 2}, "A2:B3"),
    "header_only": ([["h1", "h2"]], {}, None),
    "two_header_rows_only": ([["title"], ["h1", "h2"]], {"header_rows": 2}, None),
    "no_values": ([[None], [None, None]], {}, None),
    "empty": ([], {}, None),
    "explicit_range_on_a_header": ([["h1", "h2"]], {"auto_filter_ref": "A1:B1"}, "A1:B1"),
    "explicit_range_on_an_empty_sheet": ([], {"auto_filter_ref": "A1:C5"}, "A1:C5"),
    "explicit_range_turned_off": ([["h"], [1]], {"auto_filter_ref": "A1:A2", "auto_filter": False}, None),
    "table_has_its_own_filter": ([["a", "b"], [1, 2]], {"tables": [ExcelTable("T", "A1:B2")]}, None),
}

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


def _write(path, writer_class, rows, options):
    options = {"auto_filter": True, **options}
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, **options)


def _filter_ref(path):
    workbook = load_workbook(path)
    try:
        ref = workbook["Data"].auto_filter.ref
        return None if ref is None else str(ref)
    finally:
        workbook.close()


@pytest.mark.parametrize("case", sorted(CASES))
@pytest.mark.parametrize("writer_class, extension", WRITERS, ids=["xlsx", "xlsb"])
def test_default_filter_range(tmp_path, writer_class, extension, case):
    rows, options, expected = CASES[case]
    path = tmp_path / f"{case}.{extension}"
    _write(path, writer_class, rows, options)
    assert _filter_ref(path) == expected


@pytest.mark.com
@pytest.mark.parametrize("case", ["trailing_formatted_blank", "array_past_the_data", "two_header_rows",
                                  "table_has_its_own_filter", "header_only", "value_far_right"])
@pytest.mark.parametrize("writer_class, extension", WRITERS, ids=["xlsx", "xlsb"])
def test_excel_shows_the_same_filter(tmp_path, open_workbook, writer_class, extension, case):
    rows, options, expected = CASES[case]
    path = tmp_path / f"{case}_com.{extension}"
    _write(path, writer_class, rows, options)
    sheet = open_workbook(path).Worksheets("Data")
    if expected is None:
        assert not sheet.AutoFilterMode
    else:
        assert sheet.AutoFilterMode
        assert sheet.AutoFilter.Range.Address.replace("$", "") == expected
