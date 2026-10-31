"""Phase 5: reading an XLSB file into the editable workbook model."""

from __future__ import annotations

import datetime as dt

from xlspy import Comment, DefinedName, ExcelReader, Formula, Hyperlink, XlsbWriter, load_workbook

from ._util import close_quietly, diff_parts, zip_parts


def _source(path):
    with XlsbWriter(str(path)) as writer:
        writer.add_sheet("Data")
        writer.set_defined_name(DefinedName("MyName", "'Data'!$A$1"))
        writer.write_sheet(
            [
                ["name", "value", "link"],
                ["alpha", 1, Hyperlink("https://example.com", text="go")],
                ["beta", Formula("B2*2", value=2), None],
                [dt.datetime(2024, 1, 2), 3.5, None],
            ],
            merges=["A5:B5"],
            comments={"A1": Comment("note", "author")},
        )
    return path


def test_xlsb_model_reads_values_and_objects(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    workbook = load_workbook(source)
    try:
        assert workbook._format == "xlsb"
        sheet = workbook["Data"]
        assert sheet["A2"].value == "alpha"
        assert sheet["B2"].value == 1
        assert sheet["B3"].value == "=B2*2"
        assert sheet["A4"].value == dt.datetime(2024, 1, 2)
        assert str(sheet.merged_cells) == "A5:B5"
        assert sheet["C2"].hyperlink is not None
        assert sheet["C2"].hyperlink.target == "https://example.com"
        assert sheet["A1"].comment.text == "note"
        assert workbook.defined_names["MyName"].formula == "'Data'!$A$1"
    finally:
        close_quietly(workbook)


def test_xlsb_model_noop_save_preserves_parts(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    output = tmp_path / "copy.xlsb"
    workbook = load_workbook(source)
    try:
        workbook.save(output)
    finally:
        close_quietly(workbook)
    assert diff_parts(zip_parts(source), zip_parts(output)) == []


def test_xlsb_model_edit_cell_and_save(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    workbook = load_workbook(source)
    try:
        workbook["Data"]["B2"] = 9
        workbook.save(source)
    finally:
        close_quietly(workbook)
    with ExcelReader(str(source)) as reader:
        rows = reader.read_all("Data", data_only=False)
    assert rows[1][1] == 9


def test_xlsb_model_edit_defined_names(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    workbook = load_workbook(source)
    try:
        workbook.defined_names.remove("MyName")
        workbook.defined_names.add(DefinedName("AddedName", "'Data'!$B$2"))
        workbook.save(source)
    finally:
        close_quietly(workbook)

    reloaded = load_workbook(source)
    try:
        assert "MyName" not in reloaded.defined_names
        assert reloaded.defined_names["AddedName"].formula == "'Data'!$B$2"
    finally:
        close_quietly(reloaded)


def test_xlsb_model_edit_comment(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    workbook = load_workbook(source)
    try:
        workbook["Data"]["A1"].comment = Comment("updated", "editor")
        workbook.save(source)
    finally:
        close_quietly(workbook)

    reloaded = load_workbook(source)
    try:
        assert reloaded["Data"]["A1"].comment.text == "updated"
        assert reloaded["Data"]["A1"].comment.author == "editor"
    finally:
        close_quietly(reloaded)


def test_xlsb_read_only_still_streams(tmp_path):
    source = _source(tmp_path / "source.xlsb")
    workbook = load_workbook(source, read_only=True)
    try:
        assert workbook["Data"]["B2"].value == 1
    finally:
        close_quietly(workbook)
