"""Opt-in integration validation through the installed Microsoft Excel COM server."""

import shutil
import subprocess
from pathlib import Path

import pytest

from xlspy import XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter

pytestmark = pytest.mark.com


@pytest.mark.com
def test_updated_files_open_in_excel_com(tmp_path):
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if powershell is None:
        pytest.skip("Windows PowerShell is not available")

    xlsx_path = tmp_path / "updated.xlsx"
    xlsb_path = tmp_path / "updated.xlsb"
    with XlsxWriter(xlsx_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["old", 1], ["old2", 2]])
        writer.add_sheet("Notes")
        writer.write_sheet([["note"]])
    with XlsbWriter(xlsb_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["old", 1], ["old2", 2]])
        writer.add_sheet("Notes")
        writer.write_sheet([["note"]])

    xlsx_updater = XlsxUpdater(xlsx_path)
    xlsx_updater.replace_sheet_data("Data", [["new", 10], ["newer", 20]], headers=["Name", "Value"])
    xlsx_updater.set_sheet_state("Notes", "hidden")
    xlsx_updater.reorder_sheets(["Notes", "Data"])
    xlsx_updater.set_active_sheet(1)
    xlsx_updater.save()

    xlsb_updater = XlsbUpdater(xlsb_path)
    xlsb_updater.replace_sheet_data("Data", [["new", 10], ["newer", 20]], headers=["Name", "Value"])
    xlsb_updater.set_sheet_state("Notes", "hidden")
    xlsb_updater.reorder_sheets(["Notes", "Data"])
    xlsb_updater.set_active_sheet(1)
    xlsb_updater.save()

    script = Path(__file__).parents[1] / "tools" / "validate_excel_com.ps1"
    completed = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), str(xlsx_path), str(xlsb_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + "\n" + completed.stderr


@pytest.mark.parametrize(
    "suffix, writer_class, updater_class",
    [
        ("xlsx", XlsxWriter, XlsxUpdater),
        ("xlsb", XlsbWriter, XlsbUpdater),
    ],
)
def test_sheet_structure_updates_open_in_excel_com(tmp_path, open_workbook, suffix, writer_class, updater_class):
    path = tmp_path / f"structure.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["value", 1]])
        writer.add_sheet("Notes")
        writer.write_sheet([["note"]])
        writer.add_sheet("Temp")
        writer.write_sheet([["temp"]])

    updater = updater_class(path)
    try:
        updater.add_sheet("Added", 1)
        updater.rename_sheet("Notes", "Renamed")
        updater.remove_sheet("Temp")
        updater.save()
    finally:
        updater.close()

    book = open_workbook(path)
    assert [sheet.Name for sheet in book.Worksheets] == ["Data", "Added", "Renamed"]
    assert book.Worksheets("Data").Range("A1").Value == "value"
    assert book.Worksheets("Renamed").Range("A1").Value == "note"
    assert book.Worksheets("Added").Range("A1").Value in (None, "")
