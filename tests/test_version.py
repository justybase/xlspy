"""The package version comes from installed metadata or identifies an uninstalled source tree."""

from importlib.metadata import PackageNotFoundError, version

import xlspy


def test_version_comes_from_installed_metadata_or_marks_source_tree():
    try:
        installed_version = version("xlspy")
    except PackageNotFoundError:
        assert xlspy.__version__ == "0+unknown"
    else:
        assert xlspy.__version__ == installed_version
