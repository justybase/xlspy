"""Defined names staged while sheets are renamed, removed or reordered."""

import pytest

from xlspy import DefinedName, Workbook, load_workbook


def _source(path):
    workbook = Workbook()
    try:
        data = workbook.active
        data.title = "Data"
        data["A1"] = 1
        other = workbook.create_sheet("Other")
        other["A1"] = 2
        third = workbook.create_sheet("Third")
        third["A1"] = 3
        workbook.defined_names.add(DefinedName("LocalOther", "Other!$A$1", "Other"))
        workbook.defined_names.add(DefinedName("Global", "Data!$A$1"))
        workbook.save(path)
    finally:
        workbook.close()


def _names(path):
    workbook = load_workbook(path)
    try:
        return [
            (item.name, item.formula, item.scope, item.localSheetId)
            for item in workbook.defined_names
        ]
    finally:
        workbook.close()


def _save_edit(path, edit):
    workbook = load_workbook(path)
    try:
        edit(workbook)
        workbook.save(path)
    finally:
        workbook.close()


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_add_name_after_removing_a_scoped_sheet(tmp_path, suffix):
    path = tmp_path / f"removed.{suffix}"
    _source(path)
    _save_edit(
        path,
        lambda workbook: (
            workbook.remove(workbook["Other"]),
            workbook.defined_names.add(DefinedName("New", "Data!$A$1")),
        ),
    )

    names = {(name, scope) for name, _formula, scope, _sheet in _names(path)}
    assert names == {("Global", None), ("New", None)}


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_add_name_after_renaming_a_scoped_sheet(tmp_path, suffix):
    path = tmp_path / f"renamed.{suffix}"
    _source(path)
    _save_edit(
        path,
        lambda workbook: (
            setattr(workbook["Other"], "title", "Renamed"),
            workbook.defined_names.add(DefinedName("New", "Data!$A$1")),
        ),
    )

    names = _names(path)
    local = next(item for item in names if item[0] == "LocalOther")
    assert local[2:] == ("Renamed", 1)
    assert {item[0] for item in names} == {"LocalOther", "Global", "New"}


@pytest.mark.parametrize("suffix", ["xlsx", "xlsb"])
def test_add_name_after_reordering_sheets(tmp_path, suffix):
    path = tmp_path / f"reordered.{suffix}"
    _source(path)
    _save_edit(
        path,
        lambda workbook: (
            workbook.move_sheet("Third", -2),
            workbook.defined_names.add(DefinedName("New", "Data!$A$1")),
        ),
    )

    names = _names(path)
    local = next(item for item in names if item[0] == "LocalOther")
    # Third moved to the front, so Other is now the last sheet.
    assert local[2:] == ("Other", 2)
    assert {item[0] for item in names} == {"LocalOther", "Global", "New"}
