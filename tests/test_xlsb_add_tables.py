"""Tables can be added to a loaded XLSB workbook."""

import zipfile

import pytest

from xlspy import ExcelTable, Workbook, load_workbook


@pytest.fixture(params=[False, True], ids=["plain", "with_table"])
def source(request, tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    for row in (["name", "value", "x", "y"], ["a", 1, "p", 5], ["b", 2, "q", 6]):
        sheet.append(row)
    if request.param:
        sheet.add_table(ExcelTable("First", "A1:B3"))
    path = tmp_path / "source.xlsb"
    workbook.save(path)
    workbook.close()
    return path, request.param


def _save(path, tmp_path, edit):
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(path)
    try:
        edit(workbook)
        workbook.save(out)
    finally:
        workbook.close()
    return out


def test_table_added_to_loaded_xlsb(source, tmp_path):
    path, has_table = source
    out = _save(path, tmp_path, lambda wb: wb.active.add_table(
        ExcelTable("Second", "C1:D3", style_name="TableStyleLight9"),
    ))
    with zipfile.ZipFile(out) as archive:
        assert archive.testzip() is None
        assert sum(n.startswith("xl/tables/") for n in archive.namelist()) == 1 + has_table
    reloaded = load_workbook(out)
    try:
        tables = {table.displayName: table for table in reloaded.active.tables}
        assert set(tables) == ({"First", "Second"} if has_table else {"Second"})
        assert tables["Second"].ref == "C1:D3"
        assert [c.name for c in tables["Second"].tableColumns] == ["x", "y"]
        assert reloaded.active["D3"].value == 6
    finally:
        reloaded.close()


def test_table_on_a_new_sheet(source, tmp_path):
    path, _ = source

    def edit(workbook):
        sheet = workbook.create_sheet("Extra")
        sheet.append(["k", "v"])
        sheet.append(["x", 1])
        sheet.add_table(ExcelTable("Fresh", "A1:B2"))

    reloaded = load_workbook(_save(path, tmp_path, edit))
    try:
        assert [t.displayName for t in reloaded["Extra"].tables] == ["Fresh"]
    finally:
        reloaded.close()


def test_duplicate_name_and_overlap_are_refused(source, tmp_path):
    path, has_table = source
    if has_table:
        with pytest.raises(ValueError, match="already exists"):
            _save(path, tmp_path, lambda wb: wb.active.add_table(ExcelTable("first", "C1:D3")))
        with pytest.raises(ValueError, match="overlaps"):
            _save(path, tmp_path, lambda wb: wb.active.add_table(ExcelTable("Other", "B2:C3")))
    with pytest.raises(ValueError, match="header|name"):
        _save(path, tmp_path, lambda wb: wb.active.add_table(
            ExcelTable("Bad", "C1:D3", columns=["nope", "y"]),
        ))


def test_added_table_is_not_added_twice(source, tmp_path):
    path, _ = source
    workbook = load_workbook(path)
    try:
        workbook.active.add_table(ExcelTable("Once", "C1:D3"))
        workbook.save(tmp_path / "a.xlsb")
        workbook.save(tmp_path / "b.xlsb")
    finally:
        workbook.close()
    reloaded = load_workbook(tmp_path / "b.xlsb")
    try:
        assert [t.displayName for t in reloaded.active.tables].count("Once") == 1
    finally:
        reloaded.close()
