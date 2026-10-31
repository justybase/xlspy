"""Fixtures for the openpyxl compatibility harness."""

from __future__ import annotations

import json

import pytest

from ._util import golden_dir, golden_path


@pytest.fixture(scope="session")
def manifest():
    path = golden_dir() / "manifest.json"
    assert path.exists(), "run python tools/make_openpyxl_goldens.py first"
    return json.loads(path.read_text(encoding="utf-8"))


@pytest.fixture(params=["metadata.xlsx", "objects.xlsx", "charts.xlsx",
                        "images.xlsx", "chartsheet.xlsx", "long_tail.xlsx"])
def golden(request):
    return golden_path(request.param)
