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
from xlspy import ExcelReader, XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter
from xlspy._accel import HAVE_C_EXT


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


def main():
    assert HAVE_C_EXT, "the C extension did not load from the wheel"
    print(f"xlspy imported from: {xlspy.__file__}")
    with tempfile.TemporaryDirectory() as directory:
        directory = pathlib.Path(directory)
        _roundtrip_and_update(directory, ".xlsx", XlsxWriter, XlsxUpdater)
        _roundtrip_and_update(directory, ".xlsb", XlsbWriter, XlsbUpdater)
    print("wheel smoke test OK")


if __name__ == "__main__":
    sys.exit(main())
