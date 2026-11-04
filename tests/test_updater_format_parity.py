"""XlsxUpdater and XlsbUpdater: the same call means the same thing in both formats."""

import zipfile

import pytest

from xlspy import Comment, ExcelReader, Hyperlink, XlsbWriter, XlsxWriter, load_workbook
from xlspy._metadata import DefinedName
from xlspy.biff_utils import build_record, iter_record_payloads, iter_xlsb_sheets, xlsb_date_1904
from xlspy.exceptions import InvalidFileError
from xlspy.xlsb_updater import XlsbUpdater
from xlspy.xlsx_updater import XlsxUpdater

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}
both = pytest.mark.parametrize("ext", FORMATS)
ONE = Hyperlink("https://one.example/")
TWO = Hyperlink("https://two.example/")


def _source(tmp_path, ext, rows=None):
    path = tmp_path / f"source.{ext}"
    with FORMATS[ext][0](str(path)) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows or [["a", "b"], [ONE, TWO], [1, 2]], auto_filter=False)
        writer.add_sheet("Other")
        writer.write_sheet([[1]], auto_filter=False)
    return path


def _edit(path, ext, change, target=None):
    updater = FORMATS[ext][1](str(path))
    try:
        result = change(updater)
        updater.save(str(target or path))
    finally:
        updater.close()
    return result


def _links(path):
    workbook = load_workbook(str(path))
    try:
        return {
            cell.coordinate: cell.hyperlink.target
            for row in workbook["Data"]._cells.values() for cell in row.values() if cell.hyperlink is not None
        }
    finally:
        workbook.close()


def _names(path, ext):
    updater = FORMATS[ext][1](str(path))
    try:
        return {(n.name, n.scope): n for n in updater.get_defined_names()}
    finally:
        updater.close()


# --- set_hyperlinks ----------------------------------------------------------


@both
def test_set_hyperlinks_changes_the_given_cells_and_keeps_the_others(tmp_path, ext):
    path = _source(tmp_path, ext)
    assert _links(path) == {"A2": ONE.target, "B2": TWO.target}
    _edit(path, ext, lambda u: u.set_hyperlinks("Data", {"A3": Hyperlink("https://new.example/")}))
    assert _links(path) == {"A2": ONE.target, "B2": TWO.target, "A3": "https://new.example/"}


@both
def test_set_hyperlinks_none_removes_one_link_and_its_relationship(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_hyperlinks("Data", {"A2": None}))
    assert _links(path) == {"B2": TWO.target}
    with zipfile.ZipFile(path) as archive:
        rels = next(n for n in archive.namelist() if n.startswith("xl/worksheets/_rels/"))
        text = archive.read(rels).decode()
    assert "one.example" not in text and "two.example" in text


@both
def test_set_hyperlinks_removing_the_last_link_leaves_no_relationship(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_hyperlinks("Data", {"A2": None, "B2": None}))
    assert _links(path) == {}
    with zipfile.ZipFile(path) as archive:
        rels = [n for n in archive.namelist() if n.startswith("xl/worksheets/_rels/sheet1")]
        assert not any(b"/hyperlink" in archive.read(name) for name in rels)


@both
def test_set_hyperlinks_replaces_a_cell_s_link_and_a_missing_one_is_no_change(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_hyperlinks("Data", {"A2": Hyperlink("#Other!A1"), "C9": None}))
    assert _links(path) == {"A2": "#Other!A1", "B2": TWO.target}


@both
def test_set_hyperlinks_arguments(tmp_path, ext):
    path = _source(tmp_path, ext)
    with pytest.raises(KeyError, match="Worksheet not found"):
        _edit(path, ext, lambda u: u.set_hyperlinks("Nope", {}))
    with pytest.raises(TypeError, match="Hyperlink objects or None"):
        _edit(path, ext, lambda u: u.set_hyperlinks("Data", {"A1": "https://x.example/"}))
    _edit(path, ext, lambda u: u.set_hyperlinks("Data", {}))
    assert _links(path) == {"A2": ONE.target, "B2": TWO.target}


@both
def test_saving_a_loaded_workbook_removes_only_the_link_that_was_removed(tmp_path, ext):
    path = _source(tmp_path, ext)
    workbook = load_workbook(str(path))
    workbook["Data"]["A2"].hyperlink = None
    workbook.save(str(path))
    workbook.close()
    assert _links(path) == {"B2": TWO.target}


def test_xlsb_link_over_a_range_survives_an_edit_of_another_link(tmp_path):
    path = _source(tmp_path, "xlsb")
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet = "xl/worksheets/sheet1.bin"
    records = []
    for kind, payload in iter_record_payloads(parts[sheet]):
        payload = bytes(payload)
        if kind == 0x1EE and payload[:4] == (1).to_bytes(4, "little") and payload[8:12] == bytes(4):
            payload = payload[:4] + (2).to_bytes(4, "little") + payload[8:]  # A2 -> A2:A3
        records.append(build_record(kind, payload))
    parts[sheet] = b"".join(records)
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    _edit(path, "xlsb", lambda u: u.set_hyperlinks("Data", {"B2": None}))
    with zipfile.ZipFile(path) as archive:
        kept = [bytes(p) for k, p in iter_record_payloads(archive.read(sheet)) if k == 0x1EE]
    assert len(kept) == 1 and kept[0][:8] == (1).to_bytes(4, "little") + (2).to_bytes(4, "little")


# --- threaded comments -------------------------------------------------------


def _with_threaded(path, ext):
    """Add a threaded comment on D4 to the first sheet (no note there)."""
    rels = "xl/worksheets/_rels/sheet1." + ("xml" if ext == "xlsx" else "bin") + ".rels"
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    relationship = (
        b'<Relationship Id="threaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"'
        b' Target="../threadedComments/threadedComment1.xml"/>'
    )
    parts[rels] = parts[rels].replace(b"</Relationships>", relationship + b"</Relationships>")
    parts["xl/threadedComments/threadedComment1.xml"] = (
        b'<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
        b'<threadedComment ref="D4" personId="p" id="t"><text>x</text></threadedComment></ThreadedComments>'
    )
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


@both
def test_threaded_comment_refuses_a_new_note_but_not_a_removal(tmp_path, ext):
    path = _source(tmp_path, ext)
    _with_threaded(path, ext)
    with pytest.raises(ValueError, match="cannot replace a threaded comment"):
        _edit(path, ext, lambda u: u.set_comments("Data", {"D4": Comment("new", "Jan")}))
    _edit(path, ext, lambda u: u.set_comments("Data", {"D4": None, "A1": Comment("kept", "Jan")}))
    assert set(_edit(path, ext, lambda u: u.get_comments("Data"))) == {"A1"}


# --- defined names -----------------------------------------------------------


@both
def test_a_name_may_give_its_sheet_by_number(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("Abc", "Data!$A$1", localSheetId=1)))
    name = _names(path, ext)["Abc", "Other"]
    assert (name.scope, name.localSheetId) == ("Other", 1)
    # The same name and sheet: replaced, not a second entry.
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("ABC", "Data!$B$1", localSheetId=1)))
    assert [key for key in _names(path, ext) if key[0].casefold() == "abc"] == [("ABC", "Other")]


@both
def test_a_scoped_name_reports_its_sheet_number_too(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("Abc", "Data!$A$1", "Other")))
    assert _names(path, ext)["Abc", "Other"].localSheetId == 1


@both
@pytest.mark.parametrize("definition,error", [
    (DefinedName("Abc", "1", "Nope"), KeyError),
    (DefinedName("Abc", "1", localSheetId=9), KeyError),
    (DefinedName("Abc", "1", "Data", localSheetId=1), ValueError),
])
def test_a_name_on_a_sheet_that_is_not_there_is_refused(tmp_path, ext, definition, error):
    path = _source(tmp_path, ext)
    with pytest.raises(error):
        _edit(path, ext, lambda u: u.set_defined_name(definition))


@both
def test_hidden_names_stay_hidden(tmp_path, ext):
    path = _source(tmp_path, ext)
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("Abc", "Data!$A$1", hidden=True)))
    assert _names(path, ext)["Abc", None].hidden is True
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("Abc", "Data!$A$1")))
    assert not _names(path, ext)["Abc", None].hidden


def _workbook_part(path, ext):
    with zipfile.ZipFile(path) as archive:
        return archive.read(f"xl/workbook.{'xml' if ext == 'xlsx' else 'bin'}")


@both
def test_removing_a_name_that_is_not_there_changes_nothing(tmp_path, ext):
    path = _source(tmp_path, ext)
    before = _workbook_part(path, ext)
    _edit(path, ext, lambda u: u.remove_defined_name("Nope"))
    assert _workbook_part(path, ext) == before


def _recalculates_on_load(path, ext):
    part = _workbook_part(path, ext)
    if ext == "xlsx":
        return b'fullCalcOnLoad="1"' in part
    # BrtCalcProp's flags, as request_recalculation sets them.
    calc = next(bytes(p) for k, p in iter_record_payloads(part) if k == 0x9D)
    return int.from_bytes(calc[24:26], "little") & 0x11 == 0x11


def _without_recalculation(path, ext):
    if ext != "xlsx":
        return  # the edit sets the flag; checked after it
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b' fullCalcOnLoad="1"', b"")
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


@both
def test_changing_a_name_asks_for_a_recalculation(tmp_path, ext):
    path = _source(tmp_path, ext)
    _without_recalculation(path, ext)
    _edit(path, ext, lambda u: u.set_defined_name(DefinedName("Abc", "Data!$A$1")))
    assert _recalculates_on_load(path, ext)
    _without_recalculation(path, ext)
    _edit(path, ext, lambda u: u.remove_defined_name("Abc"))
    assert _recalculates_on_load(path, ext)


# --- XLSB: reproducible comments, one parser of workbook.bin -----------------


def _comment_part(path):
    with zipfile.ZipFile(path) as archive:
        return {n: archive.read(n) for n in archive.namelist() if n.startswith("xl/comments/")}


def test_xlsb_comment_ids_are_reproducible_and_differ_between_notes(tmp_path):
    parts = []
    for folder in ("first", "second"):
        directory = tmp_path / folder
        directory.mkdir()
        path = directory / "notes.xlsb"
        with XlsbWriter(str(path)) as writer:
            for title in ("A", "B"):
                writer.add_sheet(title)
                writer.write_sheet([[1, 2], [3, 4]], auto_filter=False,
                                   comments={"A1": Comment("same", "Jan"), "B2": Comment("same", "Jan")})
        parts.append(_comment_part(path))
    assert parts[0] == parts[1]
    guids = [bytes(p)[20:36] for data in parts[0].values() for k, p in iter_record_payloads(data) if k == 0x27B]
    assert len(guids) == 4 and len(set(guids)) == 4


def test_xlsb_updater_comment_ids_are_reproducible(tmp_path):
    results = []
    for folder in ("first", "second"):
        directory = tmp_path / folder
        directory.mkdir()
        path = _source(directory, "xlsb")
        _edit(path, "xlsb", lambda u: u.set_comments("Data", {"A1": Comment("note", "Jan")}))
        results.append(_comment_part(path))
    assert results[0] == results[1]


def test_workbook_bin_has_one_reading_of_its_sheets(tmp_path):
    path = _source(tmp_path, "xlsb")
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/workbook.bin")
    assert [name for _rid, name in iter_xlsb_sheets(data)] == ["Data", "Other"]
    assert xlsb_date_1904(data) is False
    updater = XlsbUpdater(str(path))
    try:
        names = list(updater._sheets)
    finally:
        updater.close()
    with ExcelReader(str(path)) as reader:
        assert reader.get_sheet_names() == names == ["Data", "Other"]


@pytest.mark.parametrize("body", [b"\0" * 11, b"\0" * 8 + (1).to_bytes(4, "little"), b"\0" * 12])
def test_a_cut_short_sheet_record_is_an_invalid_file(body):
    with pytest.raises(InvalidFileError):
        list(iter_xlsb_sheets(build_record(0x9C, body)))
