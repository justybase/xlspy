"""ExcelReader and read-only load_workbook on workbooks with chartsheets: a chartsheet
is listed with the sheets but has no cells, and is not an error."""
import re
import zipfile
from pathlib import Path

import pytest

from xlspy import Chart, ChartSeries, ExcelReader, Workbook, load_workbook
from xlspy.exceptions import InvalidFileError

XLSB_FIXTURE = Path(__file__).parent / "data" / "chartsheet.xlsb"  # ChartOnly (chartsheet), Arkusz1 (Excel)


@pytest.fixture(params=["xlsx", "xlsb"])
def chartsheet_file(request, tmp_path):
    """(path, worksheet name, chartsheet name)."""
    if request.param == "xlsb":
        return XLSB_FIXTURE, "Arkusz1", "ChartOnly"
    path = tmp_path / "chartsheet.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["a", 1])
    sheet.append(["b", 2])
    workbook.create_chartsheet("Chart1").add_chart(Chart("bar", [ChartSeries("B1:B2")]), source=sheet)
    workbook.save(path)
    return path, "Data", "Chart1"


def test_chartsheet_is_listed_and_has_no_rows(chartsheet_file):
    path, worksheet, chartsheet = chartsheet_file
    with ExcelReader(str(path)) as reader:
        assert set(reader.get_sheet_names()) == {worksheet, chartsheet}
        assert reader.get_chartsheet_names() == [chartsheet]
        for options in ({}, {"streaming": True}, {"data_only": False}):
            assert reader.read_all(chartsheet, **options) == []
            assert reader.read_all(worksheet, **options)
        assert reader.get_comments(chartsheet) == {}
        with pytest.raises(KeyError):
            reader.read_all("Missing")


def test_read_only_workbook_leaves_chartsheets_out_like_a_full_load(chartsheet_file):
    path, worksheet, _ = chartsheet_file
    for read_only in (True, False):
        workbook = load_workbook(path, read_only=read_only)
        try:
            assert workbook.sheetnames == [worksheet]
        finally:
            workbook.close()


@pytest.mark.parametrize("drop", ["part", "relationship"])
def test_worksheet_without_its_part_is_still_a_damaged_file(tmp_path, drop):
    with zipfile.ZipFile(XLSB_FIXTURE) as source:
        sheet_part = next(name for name in source.namelist() if name.startswith("xl/worksheets/sheet"))
        damaged = tmp_path / "damaged.xlsb"
        with zipfile.ZipFile(damaged, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if drop == "part" and info.filename == sheet_part:
                    continue
                if drop == "relationship" and info.filename == "xl/_rels/workbook.bin.rels":
                    data = re.sub(rb'<Relationship [^>]*worksheets/sheet[^>]*/>', b"", data)
                target.writestr(info, data)
    with ExcelReader(str(damaged)) as reader:
        with pytest.raises(InvalidFileError):
            reader.read_all("Arkusz1")
