"""Ignored error ranges are read, edited and written."""

import zipfile

import pytest

from xlspy import IgnoredError, Workbook, load_workbook


def _sheet_xml(path):
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/worksheets/sheet1.xml").decode()


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


def test_invalid_range_and_xlsb_are_refused(tmp_path):
    with pytest.raises(ValueError):
        IgnoredError("")
    with pytest.raises(ValueError):
        IgnoredError("not a range")
    workbook = Workbook()
    workbook.active.ignored_errors.append(IgnoredError("A1", formula=True))
    with pytest.raises(NotImplementedError):
        workbook.save(tmp_path / "x.xlsb")
    workbook.close()


def test_copy_worksheet_copies_ignored_errors():
    workbook = Workbook()
    workbook.active.ignored_errors.append(IgnoredError("A1", formula=True))
    copy = workbook.copy_worksheet(workbook.active)
    copy.ignored_errors[0].formula = False
    assert workbook.active.ignored_errors[0].formula is True
