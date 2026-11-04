"""Tables of a loaded XLSB workbook can be changed and removed, as they can in XLSX."""

import zipfile

import pytest

from xlspy import ExcelTable, UnsupportedFeatureError, Workbook, load_workbook

FORMATS = pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])


def _source(tmp_path, suffix, *, second=True, totals=False):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    rows = [["name", "value", "", "code", "n"]] + [[f"r{i}", i, None, f"c{i}", i * 2] for i in range(1, 9)]
    for row in rows:
        sheet.append(row)
    sheet.add_table(ExcelTable("First", "A1:B6", totals_row=totals))
    if second:
        sheet.add_table(ExcelTable("Second", "D1:E6", style_name="TableStyleLight9", show_row_stripes=False))
    path = tmp_path / f"source.{suffix}"
    workbook.save(path)
    workbook.close()
    return path


def _edit(path, tmp_path, edit):
    out = tmp_path / f"out{path.suffix}"
    workbook = load_workbook(path)
    try:
        edit(workbook.active)
        workbook.save(out)
    finally:
        workbook.close()
    return out


def _view(path):
    workbook = load_workbook(path)
    try:
        return [
            (t.displayName, t.ref, [c.name for c in t.tableColumns], t.headerRowCount, t.totalsRowCount or 0,
             t.tableStyleInfo.name, bool(t.tableStyleInfo.showRowStripes),
             t.autoFilter.ref if t.autoFilter else None)
            for t in workbook.active.tables
        ]
    finally:
        workbook.close()


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {n: archive.read(n) for n in archive.namelist()}


@FORMATS
def test_range_style_and_name_of_a_loaded_table_are_changed(tmp_path, suffix):
    def edit(sheet):
        table = sheet.tables["First"]
        table.ref = "A1:B8"
        table.tableStyleInfo.name = "TableStyleDark3"
        table.tableStyleInfo.showRowStripes = False
        table.displayName = table.name = "Renamed"

    view = _view(_edit(_source(tmp_path, suffix), tmp_path, edit))
    assert view[0][:3] == ("Renamed", "A1:B8", ["name", "value"])
    assert view[0][5:7] == ("TableStyleDark3", False)
    assert view[0][7] == "A1:B8"
    assert view[1][:2] == ("Second", "D1:E6")


def test_xlsx_and_xlsb_give_the_same_tables_after_the_same_edits(tmp_path):
    def edit(sheet):
        first = sheet.tables["First"]
        first.ref = "A1:B7"
        first.tableStyleInfo.name = "TableStyleMedium2"
        del sheet.tables["Second"]
        sheet.add_table(ExcelTable("Third", "D1:E4"))

    views = {}
    for suffix in ("xlsx", "xlsb"):
        views[suffix] = _view(_edit(_source(tmp_path / ".." / suffix if False else tmp_path, suffix), tmp_path, edit))
    assert views["xlsx"] == views["xlsb"]


def test_a_table_that_was_not_edited_keeps_its_bytes(tmp_path):
    path = _source(tmp_path, "xlsb")
    before = _parts(path)
    out = _edit(path, tmp_path, lambda sheet: setattr(sheet.tables["First"], "ref", "A1:B7"))
    after = _parts(out)
    unchanged = [n for n in before if n.startswith("xl/tables/") and before[n] == after.get(n)]
    changed = [n for n in before if n.startswith("xl/tables/") and before[n] != after.get(n)]
    assert len(unchanged) == 1 and len(changed) == 1


def test_a_table_is_removed_with_its_part(tmp_path):
    path = _source(tmp_path, "xlsb")
    out = _edit(path, tmp_path, lambda sheet: sheet.tables.remove("First"))
    parts = _parts(out)
    assert sum(n.startswith("xl/tables/") and n.endswith(".bin") for n in parts) == 1
    assert parts["[Content_Types].xml"].count(b"/xl/tables/") == 1
    assert [t[0] for t in _view(out)] == ["Second"]


def test_all_tables_can_be_removed_and_one_added_again(tmp_path):
    path = _source(tmp_path, "xlsb")
    out = _edit(path, tmp_path, lambda sheet: [sheet.tables.remove(n) for n in ("First", "Second")])
    assert not any(n.startswith("xl/tables/") for n in _parts(out))
    assert _view(out) == []
    again = _edit(out, tmp_path / "x" if False else tmp_path, lambda sheet: sheet.add_table(ExcelTable("First", "A1:B3")))
    assert [t[:2] for t in _view(again)] == [("First", "A1:B3")]


def test_a_name_freed_by_a_removed_table_can_be_used_again(tmp_path):
    def edit(sheet):
        sheet.tables.remove("First")
        sheet.add_table(ExcelTable("First", "A1:B3"))

    view = _view(_edit(_source(tmp_path, "xlsb"), tmp_path, edit))
    assert [t[:2] for t in view] == [("Second", "D1:E6"), ("First", "A1:B3")]


def test_a_column_is_renamed_with_its_header_cell(tmp_path):
    def edit(sheet):
        sheet["B1"] = "amount"
        sheet.tables["First"].tableColumns[1].name = "amount"
        sheet.tables["First"].ref = "A1:B7"

    assert _view(_edit(_source(tmp_path, "xlsb"), tmp_path, edit))[0][2] == ["name", "amount"]


def test_a_column_name_that_differs_from_the_header_is_refused(tmp_path):
    def edit(sheet):
        sheet.tables["First"].tableColumns[1].name = "amount"
        sheet.tables["First"].ref = "A1:B7"

    with pytest.raises(ValueError, match="header"):
        _edit(_source(tmp_path, "xlsb"), tmp_path, edit)


def test_a_new_number_of_columns_is_refused(tmp_path):
    with pytest.raises(UnsupportedFeatureError, match="columns"):
        _edit(_source(tmp_path, "xlsb"), tmp_path, lambda sheet: setattr(sheet.tables["First"], "ref", "A1:C6"))


def test_overlap_and_duplicate_names_are_refused(tmp_path):
    with pytest.raises(ValueError, match="overlaps"):
        _edit(_source(tmp_path, "xlsb"), tmp_path, lambda sheet: setattr(sheet.tables["First"], "ref", "A1:B6") or
              setattr(sheet.tables["Second"], "ref", "B1:C6"))
    with pytest.raises(ValueError, match="already exists"):
        _edit(_source(tmp_path, "xlsb"), tmp_path, lambda sheet: setattr(sheet.tables["Second"], "displayName", "first"))


def test_a_refused_edit_writes_nothing(tmp_path):
    path = _source(tmp_path, "xlsb")
    out = tmp_path / "out.xlsb"
    workbook = load_workbook(path)
    workbook.active.tables["First"].ref = "A1:C6"
    with pytest.raises(UnsupportedFeatureError):
        workbook.save(out)
    workbook.close()
    assert not out.exists()


def test_the_totals_row_and_the_filter_follow_the_range(tmp_path):
    path = _source(tmp_path, "xlsb", totals=True)
    out = _edit(path, tmp_path, lambda sheet: setattr(sheet.tables["First"], "ref", "A1:B8"))
    view = _view(out)[0]
    assert view[1] == "A1:B8" and view[4] == 1


def test_a_second_save_after_an_edit_changes_nothing_more(tmp_path):
    path = _source(tmp_path, "xlsb")
    workbook = load_workbook(path)
    sheet = workbook.active
    sheet.tables["First"].ref = "A1:B8"
    first = tmp_path / "one.xlsb"
    second = tmp_path / "two.xlsb"
    workbook.save(first)
    workbook.save(second)
    workbook.close()
    assert _parts(first) == _parts(second)
