"""Excel COM checks of ignored errors in XLSX and XLSB.

Set XLSPY_RUN_EXCEL_COM=1 to run. Excel reads, for the cells of each range, the
``Ignore`` flag of every error check; the XLSB records are the block Excel
itself writes (see ``_xlsb_ignored_errors``).
"""

import pytest

from xlspy import IgnoredError, XlsbWriter, XlsxWriter, load_workbook

pytestmark = pytest.mark.com  # skipped unless XLSPY_RUN_EXCEL_COM=1 (conftest.py)

# XlErrorChecks: number stored as text, inconsistent formula, empty cell reference.
XL_NUMBER_AS_TEXT = 3
XL_INCONSISTENT_FORMULA = 4
XL_EMPTY_CELL_REFERENCES = 7
WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


def _ignored(sheet, reference):
    """The XlErrorChecks indexes Excel shows as ignored on a cell."""
    errors = sheet.Range(reference).Errors
    return [
        index for index in range(1, 10)
        if errors.Item(index).Ignore
    ]


@pytest.mark.parametrize("writer_class,extension", WRITERS)
def test_excel_reports_ignored_errors(tmp_path, open_workbook, writer_class, extension):
    path = tmp_path / f"ignored.{extension}"
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["12", "13"], ["x", "=1+2"]],
            ignored_errors=[
                IgnoredError("A1:B1", numberStoredAsText=True),
                IgnoredError("B2", formula=True),
                IgnoredError("D1:D3", emptyCellReference=True),
            ],
        )
    sheet = open_workbook(path).Worksheets("Data")
    assert _ignored(sheet, "A1") == [XL_NUMBER_AS_TEXT]
    assert _ignored(sheet, "B1") == [XL_NUMBER_AS_TEXT]
    assert _ignored(sheet, "B2") == [XL_INCONSISTENT_FORMULA]
    assert _ignored(sheet, "D2") == [XL_EMPTY_CELL_REFERENCES]
    assert _ignored(sheet, "C3") == []


def test_loaded_xlsb_keeps_ignored_errors(tmp_path, open_workbook):
    source = tmp_path / "source.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["12", "13"], ["x", "=1+2"]],
            ignored_errors=[IgnoredError("A1:B1", numberStoredAsText=True)],
        )

    target = tmp_path / "target.xlsb"
    loaded = load_workbook(source)
    try:
        title = loaded.active.title
        loaded.active.ignored_errors.append(IgnoredError("B2", formula=True))
        loaded.save(target)
    finally:
        loaded.close()

    sheet = open_workbook(target).Worksheets(title)
    assert _ignored(sheet, "A1") == [XL_NUMBER_AS_TEXT]
    assert _ignored(sheet, "B1") == [XL_NUMBER_AS_TEXT]
    assert _ignored(sheet, "B2") == [XL_INCONSISTENT_FORMULA]
