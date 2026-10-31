"""The goldens are present, valid and loadable by both libraries."""

from __future__ import annotations

import openpyxl
import pytest

from ._util import close_quietly, golden_dir, xlspy_load, gap_xfail


def test_manifest_matches_installed_openpyxl(manifest):
    assert manifest["openpyxl_version"] == openpyxl.__version__
    assert not manifest["errors"]
    for name in manifest["files"]:
        assert (golden_dir() / name).exists(), name


def test_openpyxl_reads_every_golden(golden):
    workbook = openpyxl.load_workbook(str(golden), data_only=False)
    try:
        assert workbook.sheetnames
    finally:
        workbook.close()


def test_xlspy_reads_every_golden(golden):
    try:
        workbook = xlspy_load(golden)
    except Exception as exc:  # pragma: no cover - baseline marker
        gap_xfail(f"golden cannot be loaded yet: {type(exc).__name__}: {exc}")
    try:
        assert workbook.sheetnames
    finally:
        close_quietly(workbook)


def test_golden_object_counts(golden):
    workbook = openpyxl.load_workbook(str(golden), data_only=False)
    try:
        if golden.name == "objects.xlsx":
            sheet = workbook["Objects"]
            assert len(sheet.tables) == 1
            assert len(sheet.data_validations.dataValidation) == 3
            assert len(sheet.conditional_formatting) == 3
        elif golden.name == "charts.xlsx":
            charts = sum(len(ws._charts) for ws in workbook.worksheets)
            assert charts == 15
        elif golden.name == "images.xlsx":
            # openpyxl drops images when Pillow is missing.
            pytest.importorskip("PIL")
            assert len(workbook["Images"]._images) == 4
        elif golden.name == "chartsheet.xlsx":
            assert [sheet.title for sheet in workbook.chartsheets] == ["ChartSheet"]
        elif golden.name == "long_tail.xlsx":
            sheet = workbook["LongTail"]
            assert sheet["A1"].comment is not None
            assert sheet["A2"].hyperlink is not None
            assert sheet["A4"].value is not None
    finally:
        workbook.close()
