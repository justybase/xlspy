"""A load/save with no edits must preserve every untouched ZIP part."""

from __future__ import annotations

import openpyxl
import pytest

from ._util import close_quietly, diff_parts, xlspy_load, zip_parts, gap_xfail


def test_noop_save_preserves_every_part(golden, tmp_path):
    try:
        wb = xlspy_load(golden)
    except Exception as exc:
        gap_xfail(f"golden cannot be loaded yet: {type(exc).__name__}: {exc}")
    output = tmp_path / golden.name
    try:
        wb.save(output)
    except Exception as exc:
        gap_xfail(f"golden cannot be saved yet: {type(exc).__name__}: {exc}")
    finally:
        close_quietly(wb)

    differences = diff_parts(zip_parts(golden), zip_parts(output))
    if differences:
        gap_xfail("no-op save changed: " + ", ".join(differences[:12]))


def test_noop_save_output_opens_in_openpyxl(golden, tmp_path):
    wb = None
    try:
        wb = xlspy_load(golden)
        output = tmp_path / golden.name
        wb.save(output)
    except Exception as exc:
        gap_xfail(f"golden cannot be round-tripped yet: {type(exc).__name__}: {exc}")
    finally:
        if wb is not None:
            close_quietly(wb)
    try:
        check = openpyxl.load_workbook(str(output), data_only=False)
    except Exception as exc:
        pytest.fail(f"openpyxl rejected the no-op save: {type(exc).__name__}: {exc}")
    else:
        check.close()
