"""XlsbWriter output read by consumers other than Excel (PR #1).

Two records of the worksheet part were wrong in a way Excel tolerates but other readers do not:

* the column-info container carried ``0x86`` / ``0x87`` (``BrtEndWsViews`` / ``BrtBeginBookViews``)
  instead of ``BrtBeginColInfos`` / ``BrtEndColInfos`` (``0x186`` / ``0x187``): LibreOffice then drops
  every cell of the sheet. Fixed on main in 5112d17; the LibreOffice test guards it.
* ``BrtWsDim`` is written as ``(0, 0, 0, 0)`` whatever the data: readers that size rows from it
  (pyxlsb) fail past column A. The expected dimensions below are the ones Excel writes when it
  resaves the same data (see test_excel_com_xlsb_dimension.py), so the tests fail until the writer
  records the used range.
"""
import os
import shutil
import struct
import subprocess
import zipfile

import pytest

from xlspy import Formula, Style, XlsbWriter, load_workbook
from xlspy.biff_utils import iter_records

BRT_COL_INFO = 0x003C
BRT_END_WS_VIEWS = 0x0086
BRT_BEGIN_SHEET_DATA = 0x0091
BRT_WS_DIM = 0x0094
BRT_BEGIN_COL_INFOS = 0x0186
BRT_END_COL_INFOS = 0x0187

# (rows, write_sheet options, BrtWsDim Excel writes for them: rwFirst, rwLast, colFirst, colLast).
DIMENSION_CASES = {
    "sparse": (
        [
            [None, None, None],
            [None, "Particulars", "Q1", None],
            [None, "Sales", 10.5],
            [None, None],
            [None, "Total", Formula("=C3", value=10.5), None, None],
        ],
        {},
        (1, 4, 1, 2),
    ),
    "formula": ([[1, 2, Formula("A1*2")], [4]], {}, (0, 1, 0, 2)),
    "styled_blanks": ([["a", None, None]], {"styles": {"A1:C1": Style(bold=True)}}, (0, 0, 0, 2)),
    "empty": ([], {}, (0, 0, 0, 0)),
}


def _write(path, rows, **options):
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(rows, **options)


def _sheet_records(path):
    with zipfile.ZipFile(path) as archive:
        data = archive.read("xl/worksheets/sheet1.bin")
    return data, list(iter_records(data))


def sheet_dimension(path):
    data, records = _sheet_records(path)
    (record,) = [record for record in records if record.record_id == BRT_WS_DIM]
    return struct.unpack_from("<IIII", data, record.data_start)


def test_column_info_container_uses_the_biff12_ids(tmp_path):
    path = tmp_path / "cols.xlsb"
    _write(path, [["a", "b", "c"], [1, 2, 3]], hidden_columns=["B"], column_widths={"A": 20})
    _, records = _sheet_records(path)
    ids = [record.record_id for record in records]
    begin, end = ids.index(BRT_BEGIN_COL_INFOS), ids.index(BRT_END_COL_INFOS)
    assert ids.count(BRT_BEGIN_COL_INFOS) == ids.count(BRT_END_COL_INFOS) == 1
    assert begin < end < ids.index(BRT_BEGIN_SHEET_DATA)
    assert all(record_id == BRT_COL_INFO for record_id in ids[begin + 1:end])
    # 0x86 is BrtEndWsViews: once, closing the view block, before the columns.
    assert ids.count(BRT_END_WS_VIEWS) == 1 and ids.index(BRT_END_WS_VIEWS) < begin


@pytest.mark.xfail(strict=True, reason="XlsbWriter writes an all-zero BrtWsDim (PR #1)")
@pytest.mark.parametrize("case", [name for name in DIMENSION_CASES if name != "empty"])
def test_dimension_record_holds_the_used_range(tmp_path, case):
    rows, options, expected = DIMENSION_CASES[case]
    path = tmp_path / f"{case}.xlsb"
    _write(path, rows, **options)
    assert sheet_dimension(path) == expected


def test_empty_sheet_keeps_the_zero_dimension(tmp_path):
    rows, options, expected = DIMENSION_CASES["empty"]
    path = tmp_path / "empty.xlsb"
    _write(path, rows, **options)
    assert sheet_dimension(path) == expected


@pytest.mark.xfail(strict=True, reason="XlsbWriter writes an all-zero BrtWsDim (PR #1)")
def test_read_only_load_sizes_the_sheet_from_the_dimension(tmp_path):
    path = tmp_path / "dim.xlsb"
    rows, options, _ = DIMENSION_CASES["formula"]
    _write(path, rows, **options)
    workbook = load_workbook(path, read_only=True)
    try:
        assert workbook["Data"].dimensions == "A1:C2"
        assert workbook["Data"].max_row == 2
    finally:
        workbook.close()


# LibreOffice ---------------------------------------------------------------------------------

def _find_soffice():
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for root in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if root:
            candidate = os.path.join(root, "LibreOffice", "program", "soffice.exe")
            if os.path.isfile(candidate):
                return candidate
    return None


SOFFICE = _find_soffice()


@pytest.fixture
def libreoffice_xlsx(tmp_path):
    """Convert an .xlsb to .xlsx with headless LibreOffice and return the new path."""
    if SOFFICE is None:
        pytest.skip("LibreOffice (soffice) not found")
    # A private profile: a running LibreOffice instance would otherwise take the job.
    profile = (tmp_path / "lo_profile").resolve().as_uri()

    def convert(path):
        outdir = tmp_path / "lo_out"
        outdir.mkdir(exist_ok=True)
        subprocess.run(
            [SOFFICE, "--headless", "--norestore", f"-env:UserInstallation={profile}",
             "--convert-to", "xlsx", "--outdir", str(outdir), str(path)],
            check=True, capture_output=True, timeout=180,
        )
        converted = outdir / f"{path.stem}.xlsx"
        assert converted.is_file(), "LibreOffice produced no output"
        return converted

    return convert


@pytest.mark.libreoffice
def test_libreoffice_reads_a_sheet_with_column_info(tmp_path, libreoffice_xlsx):
    # With the old 0x86/0x87 container ids LibreOffice keeps the sheet but drops every cell.
    path = tmp_path / "cols.xlsb"
    _write(path, [["a", 1, 2], [None, "x", 3.5]], hidden_columns=["B"], column_widths={"A": 20})
    workbook = load_workbook(libreoffice_xlsx(path))
    try:
        sheet = workbook["Data"]
        assert [[cell.value for cell in row] for row in sheet.iter_rows()] == [["a", 1, 2], [None, "x", 3.5]]
        assert sheet.column_dimensions["B"].hidden
        assert sheet.column_dimensions["A"].width == pytest.approx(20, abs=1)
    finally:
        workbook.close()
