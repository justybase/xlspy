"""Smoke-test an installed xlspy wheel using only the standard library.

Run this from a directory that does *not* contain the repository (for example
from a cibuildwheel test environment or a clean virtual environment) so the
wheel under test is the one that gets imported.  Exits with a non-zero status
on any failure.
"""

import pathlib
import sys
import tempfile

import xlspy
from xlspy import ExcelReader, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter, load_workbook
from xlspy._accel import ACCEL_BACKEND, HAVE_NATIVE


def _read_all(path):
    with ExcelReader(str(path)) as reader:
        return [list(row) for row in reader.read_all("Sheet1")]


def _roundtrip_and_update(directory, suffix, writer_cls, updater_cls):
    path = directory / f"smoke{suffix}"
    rows = [["alpha", 1, True, 2.5], [None, -2, False, None]]

    with writer_cls(str(path)) as writer:
        writer.add_sheet("Sheet1")
        writer.write_sheet(rows)
    written = _read_all(path)
    assert written == rows, f"{suffix}: writer/reader mismatch: {written!r}"

    updater = updater_cls(path)
    try:
        updater.replace_sheet_data(
            "Sheet1", [["beta", 7]], headers=["name", "value"]
        )
        updater.save()
    finally:
        updater.close()
    updated = _read_all(path)
    expected = [["name", "value"], ["beta", 7]]
    assert updated == expected, f"{suffix}: updater mismatch: {updated!r}"

    # The workbook model: the XLSB load builds its cells natively (model.rs).
    workbook = load_workbook(path)
    sheet = workbook["Sheet1"]
    assert sheet["A2"].value == "beta" and sheet["B2"].value == 7, f"{suffix}: model load mismatch"
    sheet["C2"] = 3.5
    saved = directory / f"model{suffix}"
    workbook.save(saved)
    reloaded = _read_all(saved)
    assert reloaded[1] == ["beta", 7, 3.5], f"{suffix}: model save mismatch: {reloaded!r}"


def _check_typing_files():
    # Type checkers read the stub and the py.typed marker from the installed package.
    package = pathlib.Path(xlspy.__file__).parent
    for name in ("_rs_core.pyi", "py.typed"):
        assert (package / name).is_file(), f"{name} is missing from the installed package"


def main():
    assert HAVE_NATIVE, "the native extension did not load from the wheel"
    assert ACCEL_BACKEND.get("xlsb_load_cells") == "rust", "the native model loader is missing"
    _check_typing_files()
    print(f"xlspy imported from: {xlspy.__file__}")
    with tempfile.TemporaryDirectory() as directory:
        directory = pathlib.Path(directory)
        _roundtrip_and_update(directory, ".xlsx", XlsxWriter, XlsxUpdater)
        _roundtrip_and_update(directory, ".xlsb", XlsbWriter, XlsbUpdater)
    print("wheel smoke test OK")


if __name__ == "__main__":
    sys.exit(main())
