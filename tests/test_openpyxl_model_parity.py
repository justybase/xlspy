"""openpyxl behaviours the model gained: aliases, indexing, titles and streams.

Every check here compares with openpyxl itself where openpyxl can produce the
same state, so the test records the compatible behaviour rather than xlspy's
current one.
"""

import io

import pytest

from xlspy import (
    Alignment,
    Border,
    DefinedName,
    ExcelReader,
    Font,
    GradientFill,
    NamedStyle,
    PatternFill,
    Protection,
    Side,
    Stop,
    Workbook,
    load_workbook,
)
from xlspy.utils.exceptions import IllegalCharacterError


def _filled_workbook():
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = 1
    sheet["B3"] = 2
    return workbook, sheet


def test_style_aliases_match_openpyxl_keywords():
    pytest.importorskip("openpyxl")
    from openpyxl.styles import Alignment as OFAlignment, Border as OFBorder, Font as OFFont
    from openpyxl.styles import PatternFill as OFPatternFill, Side as OFSide

    font = Font(sz=12, b=True, i=True, u="single", strike=True)
    reference = OFFont(size=12, bold=True, italic=True, underline="single", strike=True)
    assert (font.size, font.bold, font.italic) == (reference.size, reference.bold, reference.italic)
    assert font.sz == font.size and font.b == font.bold and font.u == font.underline

    fill = PatternFill("solid", fgColor="FF0000", bgColor="0000FF")
    assert fill.fill_type == "solid"
    assert fill.start_color.rgb == "FFFF0000"
    assert fill.end_color.rgb == "FF0000FF"
    assert fill.patternType == "solid"
    reference_fill = OFPatternFill("solid", fgColor="FF0000", bgColor="0000FF")
    assert reference_fill.patternType == fill.patternType

    side = Side(border_style="thin", color="00FF00")
    assert side.style == "thin"
    assert OFSide(style="thin").style == "thin"

    border = Border(diagonalUp=True, diagonalDown=True, left=side)
    assert border.diagonal_up is True and border.diagonalDown is True
    assert OFBorder(diagonalUp=True).diagonalUp is True

    alignment = Alignment(
        wrapText=True,
        textRotation=45,
        shrinkToFit=True,
        readingOrder=2,
    )
    assert alignment.wrap_text is True
    assert alignment.textRotation == 45
    reference_alignment = OFAlignment(wrapText=True, textRotation=45)
    assert reference_alignment.wrap_text is True and reference_alignment.textRotation == 45


def test_style_copy_takes_openpyxl_keywords():
    font = Font(size=10, bold=True)
    changed = font.copy(sz=14, b=False)
    assert changed.size == 14 and changed.bold is False
    assert font.size == 10  # the original is unchanged

    fill = PatternFill("solid", fgColor="FF0000")
    assert fill.copy(fgColor="00FF00").start_color.rgb == "FF00FF00"

    gradient = GradientFill(stop=[Stop("FF0000", 0), Stop("0000FF", 1)])
    assert gradient.copy(type="linear").type == "linear"
    assert GradientFill(stop=[Stop("FFFFFF", 0)], type="path").fill_type == "path"

    with pytest.raises(TypeError):
        Font(size=10, sz=12)
    with pytest.raises(TypeError):
        font.copy(colour="FF0000")


def test_protection_alias_keywords():
    protection = Protection(locked=False, hidden=True)
    assert protection.locked is False
    assert protection.hidden is True
    assert protection.copy(locked=True).locked is True


def test_worksheet_indexing_matches_openpyxl():
    openpyxl = pytest.importorskip("openpyxl")

    workbook, sheet = _filled_workbook()
    reference_book = openpyxl.Workbook()
    reference = reference_book.active
    reference["A1"] = 1
    reference["B3"] = 2

    for key in ("A1", "A1:B3", "2", "A", "A:A", "1:1", "A:C", "1:3"):
        mine = sheet[key]
        theirs = reference[key]
        if isinstance(mine, tuple) and mine and isinstance(mine[0], tuple):
            assert len(mine) == len(theirs)
            assert [len(row) for row in mine] == [len(row) for row in theirs]
        elif isinstance(mine, tuple):
            assert len(mine) == len(theirs)
        else:
            assert mine.coordinate == theirs.coordinate

    for key in (2, 3):
        assert len(sheet[key]) == len(reference[key])
    for key in (slice(1, 3), slice("A", "C")):
        assert len(sheet[key]) == len(reference[key])

    with pytest.raises(IndexError):
        sheet[slice(None, 3)]
    with pytest.raises(TypeError):
        sheet[1.5]
    with pytest.raises(ValueError):
        sheet["A:1"]


def test_worksheet_iterates_like_openpyxl():
    openpyxl = pytest.importorskip("openpyxl")

    workbook, sheet = _filled_workbook()
    reference_book = openpyxl.Workbook()
    reference = reference_book.active
    reference["A1"] = 1
    reference["B3"] = 2

    rows = list(sheet)
    reference_rows = list(reference)
    assert [len(row) for row in rows] == [len(row) for row in reference_rows]
    assert [[cell.value for cell in row] for row in rows] == [[cell.value for cell in row] for row in reference_rows]


def test_assignment_of_none_clears_a_cell():
    workbook, sheet = _filled_workbook()
    sheet["A1"] = None
    assert sheet["A1"].value is None
    # cell() keeps openpyxl's meaning: no value given.
    assert sheet.cell(1, 1).value is None


def test_worksheet_object_lists_use_openpyxl_names():
    workbook = Workbook()
    sheet = workbook.active
    assert sheet.data_validations.dataValidation == []
    sheet.add_data_validation(__import__("xlspy").DataValidation(type="list", formula1='"a,b"', sqref="A1"))
    assert len(sheet.data_validations.dataValidation) == 1
    # openpyxl's idiom appends through the property; the rule must survive a save.
    sheet.data_validations.dataValidation.append(
        __import__("xlspy").DataValidation(type="whole", operator="equal", formula1="1", sqref="B1")
    )
    assert len(sheet.data_validations) == 2
    assert len(sheet.data_validations.dataValidation) == 2
    assert sheet.conditional_formatting is not None
    assert list(sheet.tables) == []


def test_workbook_membership_index_and_default_titles():
    workbook = Workbook()
    assert "Sheet" in workbook
    assert workbook.index("Sheet") == 0
    assert workbook.get_index("Sheet") == 0
    with pytest.raises(ValueError):
        workbook.index("Missing")
    with pytest.raises(ValueError):
        workbook.get_index("Missing")

    workbook.create_sheet()
    workbook.create_sheet()
    assert workbook.sheetnames == ["Sheet", "Sheet1", "Sheet2"]
    workbook.remove(workbook["Sheet1"])
    assert workbook.create_sheet().title == "Sheet3"


def test_copy_worksheet_uses_openpyxl_titles():
    workbook = Workbook()
    assert workbook.copy_worksheet(workbook.active).title == "Sheet Copy"
    assert workbook.copy_worksheet(workbook.active).title == "Sheet Copy1"


def test_freeze_panes_round_trip():
    workbook, sheet = _filled_workbook()
    sheet.freeze_panes = "B2"
    assert sheet.freeze_panes == "B2"
    sheet.freeze_panes = "A1"
    assert sheet.freeze_panes is None
    with pytest.raises(ValueError):
        sheet.freeze_panes = "A1:B2"


def test_unstyled_cells_report_openpyxl_defaults():
    workbook = Workbook()
    cell = workbook.active["A1"]
    assert cell.border.left.style is None
    assert cell.alignment.horizontal is None
    assert cell.protection.locked is True
    assert cell.protection.hidden is False
    assert cell.fill.fill_type is None

    # Reading the defaults must not store a style on the cell.
    assert cell.has_style is False

    cell.border = Border(top=Side(style="thin"))
    assert cell.border.top.style == "thin"
    assert cell.has_style is True


def test_views_scenarios_and_legacy_views():
    workbook = Workbook()
    sheet = workbook.active
    assert sheet.views.active is sheet.sheet_view
    assert sheet.views.sheetView == [sheet.sheet_view]
    sheet.sheet_view.zoomScale = 150
    assert sheet.views.active.zoomScale == 150
    assert sheet.scenarios is None
    assert sheet.legacy_drawing is None


def test_workbook_package_views_are_lists():
    workbook = Workbook()
    assert workbook.rels == []
    assert workbook.shared_strings == []
    assert workbook.loaded_theme is None


def test_named_style_and_lists_are_dict_like():
    workbook = Workbook()
    style = NamedStyle("Fancy", font=Font(bold=True), number_format="0.00")
    workbook.add_named_style(style)
    assert workbook.named_styles.names[:1] == ["Normal"]
    assert workbook.named_styles.get("Fancy") is style
    assert workbook.named_styles.keys() == workbook.named_styles.names
    assert [name for name, _style in workbook.named_styles.items()] == workbook.named_styles.names

    array = style.as_tuple()
    assert len(array) == 9
    assert array.fontId == array[0] == 0  # ids are assigned by the writer
    assert array.numFmtId == array[3]
    font_id, fill_id, border_id, num_fmt_id, protection_id, alignment_id, pivot_button, quote_prefix, xf_id = array
    assert font_id == fill_id == border_id == 0

    definition = DefinedName("Rate", "'Sheet'!$A$1")
    workbook.defined_names.add(definition)
    assert "Rate" in workbook.defined_names
    assert workbook.defined_names.get("Rate") is definition
    assert workbook.defined_names.get("Missing", "fallback") == "fallback"
    assert "Rate" in list(workbook.defined_names.keys())


def test_excel_reader_reads_a_binary_stream():
    from xlspy import XlsbWriter, XlsxWriter

    for writer in (XlsxWriter, XlsbWriter):
        stream = io.BytesIO()
        with writer(stream) as book:
            book.add_sheet("S")
            book.write_sheet([["a", 1]])
        stream.seek(0)
        with ExcelReader(stream) as reader:
            assert reader.get_sheet_names() == ["S"]
            assert reader.read_all("S") == [["a", 1]]
        # The caller's stream stays open.
        assert stream.closed is False


def test_load_workbook_reads_a_binary_stream(tmp_path):
    from xlspy import XlsbWriter, XlsxWriter

    for writer in (XlsxWriter, XlsbWriter):
        stream = io.BytesIO()
        with writer(stream) as book:
            book.add_sheet("S")
            book.write_sheet([["a", 1], ["b", 2]])
        stream.seek(0)

        workbook = load_workbook(stream)
        try:
            assert workbook.sheetnames == ["S"]
            assert workbook["S"]["A2"].value == "b"
            workbook["S"]["A2"] = "z"
            output = io.BytesIO()
            workbook.save(output)
            assert output.getvalue()
        finally:
            workbook.close()
        assert stream.closed is False

        stream.seek(0)
        streamed = load_workbook(stream, read_only=True)
        try:
            assert streamed.read_only is True
            assert streamed["S"].max_row == 2
            assert list(streamed["S"].iter_rows(values_only=True)) == [
                ("a", 1),
                ("b", 2),
            ]
        finally:
            streamed.close()
        assert stream.closed is False


def test_load_workbook_reads_a_stream_of_either_format_by_content():
    from xlspy import XlsxWriter

    stream = io.BytesIO()
    with XlsxWriter(stream) as book:
        book.add_sheet("S")
        book.write_sheet([[1]])
    stream.seek(0)
    workbook = load_workbook(stream)
    try:
        assert workbook.sheetnames == ["S"]
        assert workbook["S"]["A1"].value == 1
    finally:
        workbook.close()


def test_chart_ranges_may_name_another_sheet(tmp_path):
    from xlspy import Chart, ChartSeries

    source = tmp_path / "cross.xlsx"
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data.append(["name", "value"])
    data.append(["a", 1])
    data.append(["b", 2])
    chart_sheet = workbook.create_sheet("Chart")
    chart_sheet["A1"] = "x"
    chart_sheet.add_chart(
        Chart(
            "bar",
            [
                ChartSeries(
                    "Data!B2:B3",
                    categories="Data!A2:A3",
                    title_cell="Data!B1",
                )
            ],
        ),
        "C2",
    )
    workbook.save(source)
    workbook.close()

    import zipfile

    with zipfile.ZipFile(source) as archive:
        xml = archive.read("xl/charts/chart1.xml").decode()
    assert "'Data'!$B$2:$B$3" in xml
    assert "'Data'!$A$2:$A$3" in xml
    assert "'Data'!B1" in xml

    loaded = load_workbook(source)
    try:
        assert len(loaded["Chart"].charts) == 1
    finally:
        loaded.close()

    with pytest.raises(ValueError, match="A1 range"):
        ChartSeries("B2:B4", categories="Data!")


def test_write_only_cells_keep_their_style(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from xlspy import WriteOnlyCell

    path = tmp_path / "write_only_styles.xlsx"
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Data")
    plain = WriteOnlyCell(sheet, value="plain")
    styled = WriteOnlyCell(sheet, value="styled")
    styled.font = Font(bold=True)
    styled.fill = PatternFill("solid", start_color="FFFF00")
    styled.number_format = "0.00"
    sheet.append([plain, styled])
    workbook.save(path)
    workbook.close()

    reference = openpyxl.load_workbook(path)["Data"]
    assert reference["A1"].font.bold is not True
    assert reference["B1"].font.bold is True
    assert reference["B1"].fill.patternType == "solid"
    assert reference["B1"].number_format == "0.00"


def test_load_workbook_reads_a_stream_left_at_eof():
    """Saving to a BytesIO leaves it at EOF; passing it straight back is the
    normal hand-over, and openpyxl reads it from the start."""
    stream = io.BytesIO()
    workbook = Workbook()
    workbook.active["A1"] = "x"
    workbook.save(stream)
    assert stream.tell() > 0

    loaded = load_workbook(stream)
    try:
        assert loaded["Sheet"]["A1"].value == "x"
    finally:
        loaded.close()
    assert stream.closed is False


def test_copy_worksheet_copies_its_live_view():
    workbook = Workbook()
    workbook.active.sheet_view.zoomScale = 150
    workbook.active.freeze_panes = "B2"

    copy = workbook.copy_worksheet(workbook.active)
    assert copy.views.active is copy.sheet_view
    assert copy.views.active.zoomScale == 150
    assert copy.views.active.pane is not None
    assert copy.views.active.pane.topLeftCell == "B2"


def test_worksheet_defined_names_write_through(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")

    path = tmp_path / "names.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "S"
    workbook.active.defined_names["LOCAL"] = openpyxl.workbook.defined_name.DefinedName("LOCAL", attr_text="S!$A$1")
    workbook.save(path)

    loaded = load_workbook(path)
    try:
        assert list(loaded["S"].defined_names) == ["LOCAL"]
        loaded["S"].defined_names["NEW"] = DefinedName("NEW", "S!$C$1")
        del loaded["S"].defined_names["LOCAL"]
        assert list(loaded["S"].defined_names) == ["NEW"]
        output = tmp_path / "edited.xlsx"
        loaded.save(output)
    finally:
        loaded.close()

    result = openpyxl.load_workbook(output)
    assert list(result["S"].defined_names) == ["NEW"]


def test_named_style_as_tuple_uses_openpyxl_style_array():
    style = NamedStyle("Fancy", font=Font(bold=True))
    array = style.as_tuple()
    assert len(array) == 9
    assert array.fontId == array[0] == 0  # the writer assigns the ids
    assert array.numFmtId == array[3]
    font_id, fill_id, border_id, num_fmt_id, protection_id, alignment_id, pivot_button, quote_prefix, xf_id = array
    assert font_id == num_fmt_id == 0


def test_check_string_matches_openpyxl_truncation_and_error():
    cell = Workbook().active["A1"]
    # Truncation happens before validation, as openpyxl does it.
    assert len(cell.check_string("a" * 32767 + "\x00")) == 32767
    with pytest.raises(IllegalCharacterError):
        cell.check_string("a\x00b")
    # Characters openpyxl's regex allows pass; the writers escape them.
    assert cell.check_string("ok\uffff") == "ok\uffff"


def test_worksheet_defined_names_add(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")

    path = tmp_path / "added_names.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "S"
    sheet.defined_names.add(DefinedName("ADDED", "S!$A$1"))
    assert list(sheet.defined_names) == ["ADDED"]
    workbook.save(path)
    workbook.close()

    result = openpyxl.load_workbook(path)
    assert list(result["S"].defined_names) == ["ADDED"]


def test_bytes_paths_raise_the_openpyxl_compatible_error():
    from xlspy.exceptions import InvalidFileError
    from xlspy.utils.exceptions import InvalidFileException

    with pytest.raises(InvalidFileException):
        load_workbook(b"missing.xlsx")
    with pytest.raises(InvalidFileError):
        ExcelReader(b"missing.xlsx")


def test_invalid_slices_raise_index_error_as_openpyxl_does():
    sheet = Workbook().active
    for key in (slice(0, 3), slice(1, 0), slice(None, 3), slice(1, None)):
        with pytest.raises(IndexError):
            sheet[key]

    stream = io.BytesIO()
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.save(stream)
    stream.seek(0)
    streamed = load_workbook(stream, read_only=True)
    try:
        for key in (slice(0, 3), slice(1, 0)):
            with pytest.raises(IndexError):
                streamed["Sheet"][key]
    finally:
        streamed.close()


def test_style_copy_rejects_both_spellings():
    with pytest.raises(TypeError, match="both"):
        Font(size=12).copy(size=20, sz=14)
    assert Font(size=12).copy(sz=20).size == 20
    assert Font(size=12).copy(size=None, sz=20).size == 20


def test_chart_ranges_reject_malformed_sheet_names():
    from xlspy import ChartSeries

    for reference in ("!A2:A3", "A1:B2!B2:B3", "'Data!A2:A3"):
        with pytest.raises(ValueError, match="sheet"):
            ChartSeries("B2:B4", categories=reference)
    # A valid quoted sheet name still passes.
    ChartSeries("B2:B4", categories="'My Data'!A2:A3")


def test_style_array_is_public():
    import xlspy
    from xlspy.styles import StyleArray

    assert xlspy.StyleArray is StyleArray
    assert isinstance(NamedStyle("Fancy", font=Font(bold=True)).as_tuple(), StyleArray)


def test_cell_format_flags_are_read_and_kept(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")

    path = tmp_path / "flags.xlsx"
    reference = openpyxl.Workbook()
    reference.active["A1"] = "x"
    reference.active["A1"].quotePrefix = True
    reference.active["B1"] = "y"
    reference.active["B1"].pivotButton = True
    reference.save(path)

    loaded = load_workbook(path)
    try:
        assert loaded["Sheet"]["A1"].quotePrefix is True
        assert loaded["Sheet"]["B1"].pivotButton is True
        # Replacing another style property keeps the flags on the cell format.
        loaded["Sheet"]["A1"].font = Font(bold=True)
        output = tmp_path / "flags_edited.xlsx"
        loaded.save(output)
    finally:
        loaded.close()

    result = openpyxl.load_workbook(output)["Sheet"]
    assert result["A1"].quotePrefix is True
    assert result["A1"].font.bold is True
    assert result["B1"].pivotButton is True
