"""Ignored error ranges are read, edited and written."""

import zipfile

import pytest

from xlspy import IgnoredError, Workbook, load_workbook
from xlspy._xlsb_ignored_errors import ignored_errors_bin, ignored_errors_from_records
from xlspy.biff_utils import iter_records


def _sheet_xml(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode()


def _sheet_records(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    return [
        (record.record_id, data[record.data_start:record.data_end])
        for record in iter_records(data)
    ]


@pytest.fixture
def source(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["12", "13"])
    sheet.append([1, "=A1+B1"])
    sheet.ignored_errors.append(IgnoredError("A1:B1", numberStoredAsText=True))
    sheet.ignored_errors.append(IgnoredError("B2", formula=True, unlockedFormula=False))
    path = tmp_path / "source.xlsx"
    workbook.save(path)
    workbook.close()
    return path


def test_new_workbook_writes_ignored_errors(source):
    xml = _sheet_xml(source)
    assert '<ignoredErrors><ignoredError sqref="A1:B1" numberStoredAsText="1"/>' in xml
    assert '<ignoredError sqref="B2" formula="1" unlockedFormula="0"/>' in xml
    assert xml.index("<pageMargins") < xml.index("<ignoredErrors>")


def test_ignored_errors_are_read(source):
    workbook = load_workbook(source)
    try:
        items = workbook.active.ignored_errors
        assert [(item.sqref, item.numberStoredAsText, item.formula, item.unlockedFormula) for item in items] == [
            ("A1:B1", True, None, None), ("B2", None, True, False),
        ]
    finally:
        workbook.close()


def test_ignored_errors_can_be_edited_and_removed(source, tmp_path):
    out = tmp_path / "out.xlsx"
    workbook = load_workbook(source)
    try:
        sheet = workbook.active
        sheet.ignored_errors.pop(0)
        sheet.ignored_errors.append(IgnoredError("C1 D4:E5", formulaRange=True))
        workbook.save(out)
    finally:
        workbook.close()
    reloaded = load_workbook(out)
    try:
        assert [item.sqref for item in reloaded.active.ignored_errors] == ["B2", "C1 D4:E5"]
        reloaded.active.ignored_errors.clear()
        reloaded.save(tmp_path / "none.xlsx")
    finally:
        reloaded.close()
    assert "ignoredErrors" not in _sheet_xml(tmp_path / "none.xlsx")


def test_untouched_ignored_errors_leave_the_file_alone(source, tmp_path):
    out = tmp_path / "same.xlsx"
    workbook = load_workbook(source)
    try:
        workbook.active["A5"] = 1
        workbook.save(out)
    finally:
        workbook.close()
    assert '<ignoredError sqref="A1:B1" numberStoredAsText="1"/>' in _sheet_xml(out)


def test_invalid_range_is_refused(tmp_path):
    with pytest.raises(ValueError):
        IgnoredError("")
    with pytest.raises(ValueError):
        IgnoredError("not a range")


def test_xlsb_records_match_excels_bytes():
    items = [
        IgnoredError("A1:B2", numberStoredAsText=True),
        IgnoredError("D1", emptyCellReference=True, formula=False),
    ]
    # Captured from an XLSB saved by Excel 16: flags, one range and its bounds.
    assert ignored_errors_bin(items).hex() == (
        "880500"
        "890518" "04000000" "01000000" "00000000" "01000000" "00000000" "01000000"
        "890518" "02000000" "01000000" "00000000" "00000000" "03000000" "03000000"
        "8a0500"
    )


def test_xlsb_writer_writes_ignored_errors(tmp_path):
    path = tmp_path / "errors.xlsb"
    workbook = Workbook()
    try:
        sheet = workbook.active
        sheet.append(["12", "13"])
        sheet.ignored_errors.append(IgnoredError("A1:B1", numberStoredAsText=True))
        sheet.ignored_errors.append(IgnoredError("B2", formula=True, unlockedFormula=False))
        workbook.save(path)
    finally:
        workbook.close()
    items = ignored_errors_from_records(_sheet_records(path))
    assert [(item.sqref, item.numberStoredAsText, item.formula, item.unlockedFormula) for item in items] == [
        ("A1:B1", True, None, None), ("B2", None, True, None),
    ]


def test_xlsb_writer_option_writes_ignored_errors(tmp_path):
    from xlspy import XlsbWriter

    path = tmp_path / "option.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["12", "13"]],
            ignored_errors=[IgnoredError("A1", numberStoredAsText=True, formulaRange=True)],
        )
    items = ignored_errors_from_records(_sheet_records(path))
    assert [(item.sqref, item.numberStoredAsText, item.formulaRange) for item in items] == [
        ("A1", True, True)
    ]


def test_xlsb_ignored_errors_are_read_and_saved(tmp_path):
    from xlspy import XlsbWriter

    source = tmp_path / "source.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["12", "13"]],
            ignored_errors=[IgnoredError("A1", numberStoredAsText=True)],
        )

    out = tmp_path / "out.xlsb"
    workbook = load_workbook(source)
    try:
        assert [
            (item.sqref, item.numberStoredAsText) for item in workbook.active.ignored_errors
        ] == [("A1", True)]
        workbook.active.ignored_errors.append(IgnoredError("B1", formula=True))
        workbook.save(out)
    finally:
        workbook.close()

    assert [
        (item.sqref, item.numberStoredAsText, item.formula)
        for item in ignored_errors_from_records(_sheet_records(out))
    ] == [("A1", True, None), ("B1", None, True)]


def test_xlsb_ignored_errors_can_be_cleared(tmp_path):
    from xlspy import XlsbWriter

    source = tmp_path / "source.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["12"]], ignored_errors=[IgnoredError("A1", numberStoredAsText=True)]
        )
    out = tmp_path / "none.xlsb"
    workbook = load_workbook(source)
    try:
        workbook.active.ignored_errors.clear()
        workbook.save(out)
    finally:
        workbook.close()
    assert ignored_errors_from_records(_sheet_records(out)) == []


def test_copy_worksheet_copies_ignored_errors():
    workbook = Workbook()
    workbook.active.ignored_errors.append(IgnoredError("A1", formula=True))
    copy = workbook.copy_worksheet(workbook.active)
    copy.ignored_errors[0].formula = False
    assert workbook.active.ignored_errors[0].formula is True
