"""The xlspy command line: inspect, sheets, validate and convert."""

import json

import pytest

from xlspy import Comment, ThreadedComment, XlsbWriter, XlsxWriter, load_workbook
from xlspy.cli import main


def make_book(path, writer=XlsxWriter):
    thread = ThreadedComment("thread", "Jan")
    threaded = {"A2": thread} if writer is XlsxWriter else None
    with writer(path) as book:
        book.add_sheet("Data")
        book.write_sheet(
            [["header"], [1]],
            comments={"B1": Comment("note", "Ann")},
            threaded_comments=threaded,
            merges=["C1:D1"],
        )
        book.add_sheet("Other")
        book.write_sheet([["x"]])
    return path


def test_sheets(tmp_path, capsys):
    path = make_book(tmp_path / "book.xlsx")
    assert main(["sheets", str(path)]) == 0
    assert capsys.readouterr().out.split() == ["Data", "Other"]


def test_sheets_json(tmp_path, capsys):
    path = make_book(tmp_path / "book.xlsx")
    assert main(["sheets", str(path), "--json"]) == 0
    assert json.loads(capsys.readouterr().out) == ["Data", "Other"]


def test_inspect(tmp_path, capsys):
    path = make_book(tmp_path / "book.xlsx")
    assert main(["inspect", str(path)]) == 0
    out = capsys.readouterr().out
    assert "Path:" in out and "Data (2x4)" in out and "threaded_comments=1" in out


def test_inspect_json(tmp_path, capsys):
    path = make_book(tmp_path / "book.xlsx")
    assert main(["inspect", str(path), "--json"]) == 0
    info = json.loads(capsys.readouterr().out)
    assert info["format"] == "xlsx"
    assert [sheet["name"] for sheet in info["sheets"]] == ["Data", "Other"]
    data = info["sheets"][0]
    assert data["threaded_comments"] == 1
    assert data["merged_ranges"] == 1


def test_validate(tmp_path, capsys):
    path = make_book(tmp_path / "book.xlsb", writer=XlsbWriter)
    assert main(["validate", str(path)]) == 0
    assert capsys.readouterr().out.startswith("OK:")
    assert main(["validate", str(path), "--quick"]) == 0


def test_validate_bad_file(tmp_path, capsys):
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"not a zip")
    assert main(["validate", str(path)]) == 1
    assert "xlspy:" in capsys.readouterr().err


def test_convert_round_trip(tmp_path, capsys):
    source = make_book(tmp_path / "book.xlsx")
    binary = tmp_path / "book.xlsb"
    back = tmp_path / "back.xlsx"
    assert main(["convert", str(source), str(binary)]) == 0
    assert main(["convert", str(binary), str(back)]) == 0
    workbook = load_workbook(back)
    try:
        assert workbook["Data"]["A2"].value == 1
        assert workbook["Other"]["A1"].value == "x"
    finally:
        workbook.close()
    assert "Converted" in capsys.readouterr().out


def test_convert_rejects_bad_destination(tmp_path, capsys):
    source = make_book(tmp_path / "book.xlsx")
    assert main(["convert", str(source), str(tmp_path / "book.csv")]) == 1
    assert "unsupported destination" in capsys.readouterr().err
    assert main(["convert", str(source), str(source)]) == 1
    assert "same file" in capsys.readouterr().err


def test_missing_command_is_a_usage_error():
    with pytest.raises(SystemExit) as error:
        main([])
    assert error.value.code == 2
