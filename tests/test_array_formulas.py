"""Array and data-table formulas in the openpyxl-like model and the writers."""

import zipfile

import pytest

from xlspy import (
    ArrayFormula,
    DataTableFormula,
    Font,
    Formula,
    Workbook,
    XlsbWriter,
    XlsxUpdater,
    XlsxWriter,
    load_workbook,
)
from xlspy.exceptions import UnsupportedFeatureError


def _sheet_xml(path, name="xl/worksheets/sheet1.xml"):
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


def test_array_formula_round_trip(tmp_path):
    path = tmp_path / "array.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"], sheet["A2"] = 1, 2
    sheet["B1"] = ArrayFormula("B1:B2", "A1:A2*2")
    workbook.save(path)

    xml = _sheet_xml(path)
    assert '<f t="array" ref="B1:B2">A1:A2*2</f>' in xml

    loaded = load_workbook(path)
    try:
        cell = loaded.active["B1"]
        assert isinstance(cell.value, ArrayFormula)
        assert (cell.value.ref, cell.value.text) == ("B1:B2", "A1:A2*2")
        assert cell.data_type == "f"
        assert loaded.active.array_formulae == {"B1": "B1:B2"}
    finally:
        loaded.close()


def test_data_table_formula_round_trip(tmp_path):
    path = tmp_path / "table.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = DataTableFormula(
        "A1:C3", dt2D=True, r1="E1", r2="E2", ca=True,
    )
    workbook.save(path)

    xml = _sheet_xml(path)
    assert 't="dataTable"' in xml
    assert 'ref="A1:C3"' in xml
    assert 'dt2D="1"' in xml
    assert 'r1="E1"' in xml

    loaded = load_workbook(path)
    try:
        value = loaded.active["A1"].value
        assert isinstance(value, DataTableFormula)
        assert (value.ref, value.dt2D, value.r1, value.r2, value.ca) == (
            "A1:C3", True, "E1", "E2", True,
        )
    finally:
        loaded.close()


def test_array_formula_saves_into_a_loaded_workbook(tmp_path):
    path = tmp_path / "loaded.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.save(path)

    workbook = load_workbook(path)
    try:
        workbook.active["B1"] = ArrayFormula("B1:B2", "A1:A2+1")
        workbook.save(path)
    finally:
        workbook.close()

    loaded = load_workbook(path)
    try:
        assert isinstance(loaded.active["B1"].value, ArrayFormula)
    finally:
        loaded.close()


def test_replacing_an_array_anchor_keeps_or_changes_the_range(tmp_path):
    path = tmp_path / "replace.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = ArrayFormula("A1:A2", "1")
    workbook.save(path)

    workbook = load_workbook(path)
    try:
        workbook.active["A1"] = ArrayFormula("A1:A2", "2")
        workbook.save(path)
        workbook.active["A1"] = ArrayFormula("A1:B1", "3")
        with pytest.raises(ValueError, match="array/data-table"):
            workbook.save(path)
    finally:
        workbook.close()


def test_xlsx_writer_takes_array_and_data_table_formulas(tmp_path):
    path = tmp_path / "writer.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([
            [ArrayFormula("A1:A2", "ROW(A1:A2)"), DataTableFormula("B1:C2", r1="D1")],
        ])

    xml = _sheet_xml(path)
    assert 't="array" ref="A1:A2"' in xml
    assert 't="dataTable" ref="B1:C2"' in xml


def test_xlsx_updater_takes_array_and_data_table_formulas(tmp_path):
    path = tmp_path / "update.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1, 2]])

    updater = XlsxUpdater(path)
    try:
        updater.update_cells("Data", {
            "A1": ArrayFormula("A1:A2", "ROW(A1:A2)"),
            "B1": DataTableFormula("B1:C2", r1="D1"),
        })
        updater.save(path)
    finally:
        updater.close()

    loaded = load_workbook(path)
    try:
        assert isinstance(loaded["Data"]["A1"].value, ArrayFormula)
        assert isinstance(loaded["Data"]["B1"].value, DataTableFormula)
    finally:
        loaded.close()


def test_xlsb_writer_converts_array_formulas_and_refuses_data_tables(tmp_path):
    path = tmp_path / "binary.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1, ArrayFormula("B1:B2", "A1:A2*2")], [2, None]])

    from xlspy import ExcelReader
    with ExcelReader(path) as reader:
        assert reader.read_all("Data", data_only=False) == [[1, "=A1:A2*2"], [2, 0]]

    with pytest.raises(UnsupportedFeatureError):
        with XlsbWriter(tmp_path / "refused.xlsb") as writer:
            writer.add_sheet("Data")
            writer.write_sheet([[DataTableFormula("A1:B2")]])


def test_xlsb_updater_takes_array_formulas(tmp_path):
    from xlspy import ExcelReader, XlsbUpdater

    path = tmp_path / "updater.xlsb"
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1, 2], [3, 4]])

    updater = XlsbUpdater(path)
    try:
        updater.update_cells("Data", {
            "C1": ArrayFormula("C1:C2", "A1:A2+B1:B2"),
        })
        updater.save(path)
    finally:
        updater.close()

    with ExcelReader(path) as reader:
        assert reader.read_all("Data", data_only=False)[0][2] == "=A1:A2+B1:B2"


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_replace_sheet_data_takes_array_formulas(tmp_path, suffix):
    from xlspy import ExcelReader, XlsbUpdater

    path = tmp_path / f"replace.{suffix}"
    writer_class = XlsbWriter if suffix == "xlsb" else XlsxWriter
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a", "b"], [1, 2], [3, 4]])

    updater = XlsbUpdater(path) if suffix == "xlsb" else XlsxUpdater(path)
    try:
        updater.replace_sheet_data("Data", [
            [ArrayFormula("A1:A2", "ROW(A1:A2)")],
            [],
        ])
        updater.save(path)
    finally:
        updater.close()

    if suffix == "xlsb":
        with ExcelReader(path) as reader:
            assert reader.read_all("Data", data_only=False)[0][0] == "=ROW(A1:A2)"
    else:
        loaded = load_workbook(path)
        try:
            assert isinstance(loaded["Data"]["A1"].value, ArrayFormula)
        finally:
            loaded.close()


def test_array_formula_moves_with_move_range(tmp_path):
    path = tmp_path / "move.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["B2"] = ArrayFormula("B2:B3", "A2:A3")
    sheet.move_range("B2:B3", rows=1, cols=1, translate=True)
    workbook.save(path)

    loaded = load_workbook(path)
    try:
        value = loaded.active["C3"].value
        assert isinstance(value, ArrayFormula)
        assert value.ref == "C3:C4"
        assert value.text == "B3:B4"
    finally:
        loaded.close()


def test_openpyxl_model_aliases():
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = 1
    sheet["A1"].font = Font(bold=True)
    sheet.merge_cells("B1:C1")
    sheet.column_dimensions.group("D", "F", outline_level=1)

    assert sheet.parent is workbook
    assert sheet.encoding == "utf-8"
    assert sheet.mime_type.endswith("worksheet+xml")
    assert sheet.path == "/xl/worksheets/sheet1.xml"
    assert sheet.merged_cell_ranges == ["B1:C1"]
    assert sheet.column_groups == ["D:F"]
    assert sheet["A1"].col_idx == 1
    assert sheet["A1"].has_style is True
    assert isinstance(sheet["A1"].style_id, int)

    assert workbook.excel_base_date == workbook.epoch
    assert workbook.data_only is False
    assert workbook.get_index(sheet) == 0
    assert workbook.get_sheet_names() == workbook.sheetnames
    assert workbook.style_names == ["Normal"]
    assert workbook.path == "/xl/workbook.xml"
    assert workbook.mime_type.endswith("sheet.main+xml")
    workbook.remove_sheet(sheet)
    assert workbook.sheetnames == []


def test_loaded_workbook_reports_data_only(tmp_path):
    path = tmp_path / "data.xlsx"
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[Formula("1+1", 2)]])

    loaded = load_workbook(path)
    assert loaded.data_only is False
    loaded.close()
    loaded = load_workbook(path, data_only=True)
    try:
        assert loaded.data_only is True
    finally:
        loaded.close()
