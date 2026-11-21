"""Threaded (Excel 365) comments: reading, writing and editing conversations."""

import datetime
import io
import zipfile
from xml.etree import ElementTree as ET

import pytest

from xlspy import (
    Comment,
    ExcelReader,
    ThreadedComment,
    UnsupportedFeatureError,
    XlsbUpdater,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)
from xlspy.threaded_comments import THREADED_NS


def write_book(path, **options):
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["Header"], [10]], **options)


def parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def conversation(text="Root", author="Jan"):
    thread = ThreadedComment(text, author)
    thread.reply("A reply", "Ola")
    return thread


def test_reply_tree_and_validation():
    thread = conversation()
    assert [message.text for message in thread.messages] == ["Root", "A reply"]
    assert thread.replies[0].parent_id is None

    with pytest.raises(TypeError):
        ThreadedComment(1, "Jan")
    with pytest.raises(TypeError):
        ThreadedComment("text", "Jan", created="yesterday")

    nested = ThreadedComment("reply", "Ola")
    nested.reply("another", "Piotr")
    with pytest.raises(ValueError):
        ThreadedComment("text", "Jan", replies=[nested])


def test_writer_writes_conversation_parts_and_companion(tmp_path):
    path = tmp_path / "threads.xlsx"
    write_book(
        path,
        comments={"B1": Comment("classic", "Ann")},
        threaded_comments={"A2": conversation()},
    )
    archive = parts(path)
    assert "xl/threadedComments/threadedComment1.xml" in archive
    assert "xl/persons/person.xml" in archive
    assert "xl/comments/comment1.xml" in archive
    assert b"threadedcomments+xml" in archive["[Content_Types].xml"]
    assert b"person+xml" in archive["[Content_Types].xml"]
    assert b"threadedComment" in archive["xl/worksheets/_rels/sheet1.xml.rels"]
    assert b"relationships/person" in archive["xl/_rels/workbook.xml.rels"]

    threaded = ET.fromstring(archive["xl/threadedComments/threadedComment1.xml"])
    messages = threaded.findall(f"{{{THREADED_NS}}}threadedComment")
    assert [message.get("ref") for message in messages] == ["A2", "A2"]
    assert messages[1].get("parentId") == messages[0].get("id")
    assert messages[0].findtext(f"{{{THREADED_NS}}}text") == "Root"

    with ExcelReader(str(path)) as reader:
        assert reader.get_comments("Data") == {"B1": Comment("classic", "Ann")}
        read = reader.get_threaded_comments("Data")["A2"]
        assert read.text == "Root"
        assert read.replies[0].text == "A reply"
        assert read.replies[0].author == "Ola"

    workbook = load_workbook(path)
    sheet = workbook["Data"]
    assert sheet["A2"].threaded_comment.text == "Root"
    assert sheet["A2"].threaded_comment.replies[0].text == "A reply"
    assert sheet["A2"].comment is None
    assert sheet["B1"].comment.text == "classic"


def test_writer_keeps_timestamps_and_resolution(tmp_path):
    path = tmp_path / "resolved.xlsx"
    created = datetime.datetime(2026, 5, 1, 12, 30, 15, 250000)
    thread = ThreadedComment("Done", "Jan", created=created, resolved=True)
    write_book(path, threaded_comments={"A1": thread})
    workbook = load_workbook(path)
    read = workbook["Data"]["A1"].threaded_comment
    assert read.resolved is True
    assert read.created == created


def test_writer_refuses_note_and_thread_on_one_cell(tmp_path):
    path = tmp_path / "both.xlsx"
    with pytest.raises(ValueError, match="both a note and a threaded comment"):
        write_book(
            path,
            comments={"A1": Comment("classic", "Ann")},
            threaded_comments={"A1": conversation()},
        )


def test_xlsb_writer_refuses_threaded_comments(tmp_path):
    path = tmp_path / "threads.xlsb"
    with pytest.raises(UnsupportedFeatureError, match="threaded comments"):
        with XlsbWriter(path) as writer:
            writer.add_sheet("Data")
            writer.write_sheet([["Header"]], threaded_comments={"A1": conversation()})


def test_updater_edits_adds_and_removes_conversations(tmp_path):
    path = tmp_path / "edit.xlsx"
    write_book(path, comments={"B1": Comment("classic", "Ann")}, threaded_comments={"A2": conversation()})

    with XlsxUpdater(path) as updater:
        current = updater.get_threaded_comments("Data")["A2"]
        edited = ThreadedComment("Root edited", "Jan", id=current.id)
        edited.reply("Reply edited", "Ola", id=current.replies[0].id)
        updater.set_threaded_comments("Data", {"A2": edited, "A3": ThreadedComment("New", "Piotr")})
        assert updater.get_threaded_comments("Data")["A2"].text == "Root edited"
        assert updater.get_comments("Data") == {"B1": Comment("classic", "Ann")}
        updater.save()

    workbook = load_workbook(path)
    sheet = workbook["Data"]
    assert sheet["A2"].threaded_comment.text == "Root edited"
    assert sheet["A2"].threaded_comment.replies[0].text == "Reply edited"
    assert sheet["A3"].threaded_comment.text == "New"
    assert sheet["B1"].comment.text == "classic"
    assert sheet["A2"].comment is None

    with XlsxUpdater(path) as updater:
        updater.set_threaded_comments("Data", {"A2": None, "A3": None})
        updater.save()

    archive = parts(path)
    assert not [name for name in archive if "threadedComments" in name or "persons/" in name]
    workbook = load_workbook(path)
    assert workbook["Data"]["A2"].threaded_comment is None
    assert workbook["Data"]["B1"].comment.text == "classic"
    with ExcelReader(str(path)) as reader:
        assert reader.get_threaded_comments("Data") == {}


def test_updater_removes_threaded_and_comment_parts_when_empty(tmp_path):
    path = tmp_path / "empty.xlsx"
    write_book(path, threaded_comments={"A1": conversation()})
    with XlsxUpdater(path) as updater:
        updater.set_threaded_comments("Data", {"A1": None})
        updater.save()
    archive = parts(path)
    assert not [name for name in archive if "threadedComments" in name]
    assert not [name for name in archive if "persons/" in name]
    assert "xl/comments/comment1.xml" not in archive
    assert "xl/drawings/commentsDrawing1.vml" not in archive


def test_updater_threaded_comment_cannot_replace_classic_note(tmp_path):
    path = tmp_path / "classic.xlsx"
    write_book(path, comments={"A1": Comment("classic", "Ann")})
    with XlsxUpdater(path) as updater:
        with pytest.raises(ValueError, match="classic note"):
            updater.set_threaded_comments("Data", {"A1": conversation()})


def test_editing_classic_notes_keeps_threaded_companions(tmp_path):
    path = tmp_path / "mix.xlsx"
    write_book(path, comments={"B1": Comment("classic", "Ann")}, threaded_comments={"A1": conversation()})
    with XlsxUpdater(path) as updater:
        updater.set_comments("Data", {"B1": Comment("edited", "Ann")})
        updater.save()
    with ExcelReader(str(path)) as reader:
        assert reader.get_comments("Data") == {"B1": Comment("edited", "Ann")}
        assert reader.get_threaded_comments("Data")["A1"].text == "Root"


def test_loaded_workbook_saves_threaded_edits(tmp_path):
    path = tmp_path / "loaded.xlsx"
    write_book(path, threaded_comments={"A2": conversation()})

    workbook = load_workbook(path)
    sheet = workbook["Data"]
    thread = sheet["A2"].threaded_comment
    thread.text = "Changed"
    thread.replies[0].text = "Changed reply"
    sheet["A2"].threaded_comment = thread
    sheet["A4"].threaded_comment = ThreadedComment("Added", "Ewa")
    workbook.save(path)

    workbook = load_workbook(path)
    sheet = workbook["Data"]
    assert sheet["A2"].threaded_comment.text == "Changed"
    assert sheet["A2"].threaded_comment.replies[0].text == "Changed reply"
    assert sheet["A4"].threaded_comment.text == "Added"


def test_threaded_comments_move_with_cells_and_copy(tmp_path):
    path = tmp_path / "moves.xlsx"
    write_book(path, threaded_comments={"A2": conversation()})
    workbook = load_workbook(path)
    sheet = workbook["Data"]
    sheet.insert_rows(1)
    assert sheet["A2"].threaded_comment is None
    assert sheet["A3"].threaded_comment.text == "Root"
    copy = workbook.copy_worksheet(sheet)
    assert copy["A3"].threaded_comment.text == "Root"
    copy["A3"].threaded_comment.text = "Copy"
    assert sheet["A3"].threaded_comment.text == "Root"


def test_two_saves_are_identical_with_threaded_comments(tmp_path):
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    write_book(first, threaded_comments={"A1": conversation()})
    write_book(second, threaded_comments={"A1": conversation()})
    assert first.read_bytes() == second.read_bytes()


def test_xlsb_updater_reads_but_refuses_to_edit(tmp_path):
    path = tmp_path / "plain.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["Header"], [10]])
    with XlsbUpdater(path) as updater:
        assert updater.get_threaded_comments("Data") == {}
        with pytest.raises(UnsupportedFeatureError, match="XLSB"):
            updater.set_threaded_comments("Data", {"A1": conversation()})

    buffer = io.BytesIO()
    with XlsxWriter(buffer) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["Header"]], threaded_comments={"A1": conversation()})
    assert buffer.getvalue()
