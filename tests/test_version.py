"""The package version has one source: [project].version in pyproject.toml."""

import tomllib
from importlib.metadata import version
from pathlib import Path

import xlspy

PYPROJECT = Path(__file__).resolve().parents[1] / "pyproject.toml"


def test_version_comes_from_the_installed_metadata():
    assert xlspy.__version__ == version("xlspy")


def test_installed_metadata_matches_pyproject():
    # Fails after a version bump until the package is reinstalled (pip install -e .).
    with PYPROJECT.open("rb") as handle:
        expected = tomllib.load(handle)["project"]["version"]
    assert xlspy.__version__ == expected
