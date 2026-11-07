"""Excel opens what both updaters write for links, notes and names, and reads it the same."""

import pytest

from xlspy import Comment, Formula, Hyperlink, XlsbWriter, XlsxWriter, load_workbook
from xlspy._metadata import DefinedName
from xlspy.xlsb_updater import XlsbUpdater
from xlspy.xlsx_updater import XlsxUpdater

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

FORMATS = {"xlsx": (XlsxWriter, XlsxUpdater), "xlsb": (XlsbWriter, XlsbUpdater)}


def _links(sheet):
    return {link.Range.Address.replace("$", ""): link.Address or link.SubAddress for link in sheet.Hyperlinks}


@pytest.mark.parametrize("ext", FORMATS)
def test_edited_links_notes_and_names_open_in_excel(tmp_path, ext, open_workbook):
    writer_cls, updater_cls = FORMATS[ext]
    path = tmp_path / f"edited.{ext}"
    with writer_cls(str(path)) as writer:
        for title in ("Data", "Other"):
            writer.add_sheet(title)
            writer.write_sheet([
                [Hyperlink("https://one.example/", "one"), Hyperlink("https://two.example/", "two")],
                [Formula("Rate*2"), 21],
            ], auto_filter=False, comments={"A1": Comment("same text", "Jan"), "B2": Comment("same text", "Jan")})
        writer.set_defined_name(DefinedName("Rate", "Data!$B$2"))

    updater = updater_cls(str(path))
    try:
        updater.set_hyperlinks("Data", {"A1": None, "A2": Hyperlink("https://new.example/")})
        updater.set_comments("Data", {"A1": None, "B1": Comment("added", "Ola")})
        updater.set_defined_name(DefinedName("Secret", "Data!$A$2", hidden=True))
        updater.set_defined_name(DefinedName("Local", "Data!$B$2", localSheetId=1))
        updater.save(str(path))
    finally:
        updater.close()

    workbook = open_workbook(path)
    data, other = workbook.Worksheets("Data"), workbook.Worksheets("Other")
    assert _links(data) == {"B1": "https://two.example/", "A2": "https://new.example/"}
    assert _links(other) == {"A1": "https://one.example/", "B1": "https://two.example/"}
    assert data.Comments.Count == 2 and other.Comments.Count == 2
    assert data.Range("B1").Comment.Text().endswith("added")
    assert data.Range("B2").Comment.Text().endswith("same text")
    assert other.Range("A1").Comment.Text().endswith("same text")
    names = {name.Name: name for name in workbook.Names}
    assert names["Secret"].Visible is False
    assert names["Rate"].Visible is True
    assert "Other!Local" in names
    assert data.Range("A2").Value == 42


@pytest.mark.parametrize("ext", FORMATS)
def test_batched_defined_names_open_in_excel(tmp_path, ext, open_workbook):
    """set/remove of many names in one updater call reaches Excel as-is."""
    writer_cls, updater_cls = FORMATS[ext]
    path = tmp_path / f"batched.{ext}"
    with writer_cls(str(path)) as writer:
        for title in ("Data", "Other"):
            writer.add_sheet(title)
            writer.write_sheet([[1], [2]])
        writer.set_defined_name(DefinedName("Rate", "Data!$A$2"))

    updater = updater_cls(str(path))
    try:
        updater.set_defined_names([
            DefinedName("Rate", "Data!$A$2"),
            DefinedName("Markup", "Data!$B$2"),
            DefinedName("Secret", "Data!$A$2", hidden=True),
            DefinedName("Local", "Data!$B$2", localSheetId=0),
        ])
        updater.remove_defined_names([("Markup", None)])
        updater.save(str(path))
    finally:
        updater.close()

    workbook = open_workbook(path)
    # Excel exposes the hidden built-in names its own filters create (the
    # writer turns on a filter with the header), and removed names keep a
    # hidden placeholder slot for formula ordinals.
    names = {
        name.Name: name
        for name in workbook.Names
        if not name.Name.endswith("_FilterDatabase")
        and not name.Name.startswith("_xlspy_deleted_")
    }
    assert set(names) == {"Rate", "Secret", "Data!Local"}
    assert names["Secret"].Visible is False
    assert names["Rate"].Visible is True


@pytest.mark.parametrize("ext", FORMATS)
def test_staged_defined_names_open_in_excel(tmp_path, ext, open_workbook):
    """Names staged on a loaded workbook keep their sheet scope in Excel."""
    writer_cls, _updater_cls = FORMATS[ext]
    path = tmp_path / f"staged.{ext}"
    with writer_cls(str(path)) as writer:
        for title in ("Data", "Other"):
            writer.add_sheet(title)
            writer.write_sheet([[1], [2]])
        writer.set_defined_name(DefinedName("Rate", "Data!$A$2"))
        writer.set_defined_name(DefinedName("Local", "Data!$B$2", localSheetId=0))

    book = load_workbook(str(path))
    try:
        book.defined_names.add(DefinedName("Moved", "Other!$A$1", "Other"))
        book.move_sheet("Other", -1)
        book.save(str(path))
    finally:
        book.close()

    workbook = open_workbook(path)
    # Built-in _FilterDatabase names appear for sheets with the writer's filter.
    names = {
        name.Name
        for name in workbook.Names
        if not name.Name.endswith("_FilterDatabase")
    }
    # Data kept its local name after Other moved ahead of it.
    assert names == {"Rate", "Data!Local", "Other!Moved"}
