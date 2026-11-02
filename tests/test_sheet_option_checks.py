"""Both writers check write_sheet options the same way, before anything is recorded:
a rejected call leaves the writer usable (xlspy._writer_common.WriterFrontEnd._check_sheet_options)."""
from __future__ import annotations

import pytest

from xlspy import ExcelReader, ExcelTable, XlsbWriter, XlsxWriter
from xlspy import Chart
from xlspy import ChartSeries
from xlspy import IconSetRule
import zipfile

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


@pytest.mark.parametrize(("writer_class", "suffix"), WRITERS)
@pytest.mark.parametrize(("ref", "error"), [(5, TypeError), ("not a range", ValueError)])
def test_bad_auto_filter_ref_is_refused_before_the_rows_stream(tmp_path, writer_class, suffix, ref, error) -> None:
    path = tmp_path / f"book.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(error):
            writer.write_sheet([["a", "b"], [1, 2]], auto_filter_ref=ref)
        writer.write_sheet([["a", "b"], [1, 2]], auto_filter_ref="A1:B2")
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["a", "b"], [1, 2]]


@pytest.mark.parametrize(("writer_class", "suffix"), WRITERS)
def test_rejected_call_claims_no_table_name(tmp_path, writer_class, suffix) -> None:
    path = tmp_path / f"book.{suffix}"
    table = ExcelTable("Sales", "A1:B2")
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError):
            writer.write_sheet([["a", "b"], [1, 2]], tables=[table], zoom="big")
        writer.write_sheet([["a", "b"], [1, 2]], tables=[table])


openpyxl = pytest.importorskip("openpyxl")


def test_xlsb_rejected_sheet_leaves_no_charts_tables_or_priorities(tmp_path):
    path = tmp_path / "retry.xlsb"
    data = [["k", "v"], ["a", 1], ["b", 2]]
    chart = Chart("bar", [ChartSeries("B2:B3", categories="A2:A3")])
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(ValueError, match="duplicate table display name"):
            writer.write_sheet(
                data,
                charts=[chart],
                tables=[ExcelTable("Sales", "A1:B3"), ExcelTable("Sales", "D1:E3")],
            )
        cf = [IconSetRule("B2:B3", "3TrafficLights1", [0, 33, 67], priority=5)]
        with pytest.raises(ValueError, match="duplicate conditional formatting priority 5"):
            writer.write_sheet(data, conditional_formats=cf + cf)
        # The same name, chart-free and with priority 5, is accepted afterwards.
        writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], conditional_formats=cf)

    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
    assert not any(name.startswith(("xl/charts/", "xl/drawings/")) for name in names)
    assert "xl/tables/table1.bin" in names and "xl/tables/table2.bin" not in names
    with ExcelReader(str(path)) as reader:
        assert [row[:2] for row in reader.get_rows("Data")][:3] == data


def test_xlsx_rejected_sheet_does_not_reserve_table_names_or_priorities(tmp_path):
    path = tmp_path / "retry.xlsx"
    data = [["k", "v"], ["a", 1], ["b", 2]]
    cf = [IconSetRule("B2:B3", "3TrafficLights1", [0, 33, 67], priority=3)]
    with XlsxWriter(path) as writer:
        with pytest.raises(TypeError, match="images must be Image instances"):
            writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], images=["logo.png"])
        with pytest.raises(ValueError, match="duplicate conditional formatting priority 3"):
            writer.write_sheet(data, conditional_formats=cf + cf)
        writer.write_sheet(data, tables=[ExcelTable("Sales", "A1:B3")], conditional_formats=cf)

    sheet = openpyxl.load_workbook(path).active
    assert list(sheet.tables) == ["Sales"]
