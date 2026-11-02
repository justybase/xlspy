"""Limits on what a package may make the readers hold in memory."""
from __future__ import annotations

import zipfile

import pytest

from xlspy import _untrusted


def test_package_total_size_is_limited(tmp_path, monkeypatch):
    path = tmp_path / "many.zip"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("a.xml", b"x" * 1000)
        archive.writestr("b.xml", b"y" * 1000)

    monkeypatch.setattr(_untrusted, "MAX_PACKAGE_BYTES", 1500)
    with pytest.raises(ValueError, match="package declares"):
        _untrusted.SafeZipFile(path, "r")

    monkeypatch.setattr(_untrusted, "MAX_PACKAGE_BYTES", 4000)
    with _untrusted.SafeZipFile(path, "r") as archive:
        assert archive.read("a.xml") == b"x" * 1000
