"""The package version comes from the installed metadata."""

from importlib.metadata import version

import xlspy


def test_version_comes_from_the_installed_metadata():
    assert xlspy.__version__ == version("xlspy")
