"""The object model reaches an updater's package only through PackageHandle."""
from __future__ import annotations

import ast
import pathlib

import pytest

import xlspy
from xlspy import XlsbWriter, XlsxWriter, load_workbook
from xlspy._package_handle import PackageHandle

WORKBOOK_DIR = pathlib.Path(xlspy.__file__).parent / "workbook"


def _private_updater_accesses(path: pathlib.Path) -> list[str]:
    """``<...updater>._name`` attribute reads: an updater's private state."""
    found = []
    for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
        if not isinstance(node, ast.Attribute) or not node.attr.startswith("_") or node.attr.startswith("__"):
            continue
        owner = node.value
        name = owner.id if isinstance(owner, ast.Name) else owner.attr if isinstance(owner, ast.Attribute) else ""
        if name.endswith("updater"):
            found.append(f"{path.name}:{node.lineno} {name}.{node.attr}")
    return found


def test_workbook_package_uses_no_private_updater_state() -> None:
    accesses = [hit for path in sorted(WORKBOOK_DIR.glob("*.py")) for hit in _private_updater_accesses(path)]
    assert not accesses, "use PackageHandle instead:\n" + "\n".join(accesses)


def test_guard_detects_a_private_access(tmp_path: pathlib.Path) -> None:
    sample = tmp_path / "sample.py"
    sample.write_text("def f(updater, workbook):\n    return updater._sheets, workbook._updater._package\n")
    assert len(_private_updater_accesses(sample)) == 2


def _xlsx(path: pathlib.Path) -> pathlib.Path:
    with XlsxWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a"], [1]])
    return path


def _xlsb(path: pathlib.Path) -> pathlib.Path:
    with XlsbWriter(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["a"], [1]])
    return path


def test_handle_views_a_loaded_xlsx(tmp_path: pathlib.Path) -> None:
    workbook = load_workbook(_xlsx(tmp_path / "book.xlsx"))
    handle = PackageHandle(workbook._updater)
    assert list(handle.sheets) == ["Data"]
    assert handle.date_1904 is False
    path, xml = handle.sheet_xml("Data")
    assert handle.part_or_none(path) == xml.encode("utf-8")
    assert handle.part_or_none("xl/missing.xml") is None
    assert handle.tag_prefix(xml, "sheetData") == ""
    with pytest.raises(TypeError, match="XLSB package operation"):
        handle.sheet_bin("Data")
    with pytest.raises(TypeError, match="XLSB package operation"):
        _ = handle.chartsheets
    workbook.close()


def test_handle_views_a_loaded_xlsb(tmp_path: pathlib.Path) -> None:
    workbook = load_workbook(_xlsb(tmp_path / "book.xlsb"))
    handle = PackageHandle(workbook._updater)
    assert list(handle.sheets) == ["Data"]
    assert handle.chartsheets == set()
    path, data = handle.sheet_bin("Data")
    assert handle.package.get(path) == data
    with pytest.raises(TypeError, match="XLSX package operation"):
        handle.recalculate({}, [])
    workbook.close()


def test_handle_reads_the_live_sheet_map(tmp_path: pathlib.Path) -> None:
    workbook = load_workbook(_xlsx(tmp_path / "book.xlsx"))
    handle = PackageHandle(workbook._updater)
    workbook._updater._sheets = {"Other": "xl/worksheets/sheet1.xml"}  # an updater may rebind the map
    assert list(handle.sheets) == ["Other"]
    workbook.close()
