"""Excel reads xlspy's threaded comments, and xlspy reads and edits Excel's."""

import pytest

from xlspy import ThreadedComment, XlsxWriter, load_workbook

pytestmark = pytest.mark.com

XL_XLSX = 51
XL_XLSB = 50


def test_excel_reads_the_thread_xlspy_wrote(tmp_path, open_workbook):
    path = tmp_path / "threads.xlsx"
    thread = ThreadedComment("Confirm the amount", "Jan", resolved=True)
    thread.reply("Confirmed", "Ola")
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["Amount"], [120]], threaded_comments={"A2": thread})

    book = open_workbook(path)
    cell = book.Worksheets("Data").Range("A2")
    assert cell.CommentThreaded.Text() == "Confirm the amount"
    assert cell.CommentThreaded.Author.Name == "Jan"
    assert cell.CommentThreaded.Replies.Count == 1
    assert cell.CommentThreaded.Replies(1).Text() == "Confirmed"
    assert cell.CommentThreaded.Resolved is True
    assert cell.Comment is None


def test_xlspy_reads_and_edits_the_thread_excel_wrote(tmp_path, excel):
    path = tmp_path / "excel.xlsx"
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        root = sheet.Range("A1").AddCommentThreaded("Written by Excel")
        root.AddReply("A reply")
        book.SaveAs(str(path), XL_XLSX)
    finally:
        book.Close(False)

    loaded = load_workbook(path)
    try:
        thread = loaded["Data"]["A1"].threaded_comment
        assert thread.text == "Written by Excel"
        assert thread.replies[0].text == "A reply"
        thread.text = "Edited by xlspy"
        loaded["Data"]["A1"].threaded_comment = thread
        loaded["Data"]["B2"].threaded_comment = ThreadedComment("Added by xlspy", "Ola")
        loaded.save(path)
    finally:
        loaded.close()

    book = excel.Workbooks.Open(str(path))
    try:
        sheet = book.Worksheets("Data")
        assert sheet.Range("A1").CommentThreaded.Text() == "Edited by xlspy"
        assert sheet.Range("B2").CommentThreaded.Text() == "Added by xlspy"
    finally:
        book.Close(False)


def test_xlspy_reads_the_xlsb_thread_excel_wrote(tmp_path, excel):
    path = tmp_path / "excel.xlsb"
    book = excel.Workbooks.Add()
    try:
        sheet = book.Worksheets(1)
        sheet.Name = "Data"
        root = sheet.Range("A1").AddCommentThreaded("XLSB thread")
        root.AddReply("XLSB reply")
        book.SaveAs(str(path), XL_XLSB)
    finally:
        book.Close(False)

    loaded = load_workbook(path)
    try:
        thread = loaded["Data"]["A1"].threaded_comment
        assert thread.text == "XLSB thread"
        assert thread.replies[0].text == "XLSB reply"
    finally:
        loaded.close()
