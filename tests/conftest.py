import os, shutil

import pytest

PROJECT_TMP = os.path.join(os.path.dirname(__file__), "..", ".test_tmp")


def pytest_collection_modifyitems(config, items):
    """Excel reads and sets PageSetup through the default printer: without
    one it waits, shows a printer prompt and fails, so these tests are opt-in."""
    if os.environ.get("XLSPY_RUN_EXCEL_PRINT") == "1":
        return
    skip = pytest.mark.skip(reason="Set XLSPY_RUN_EXCEL_PRINT=1 (needs a reachable default printer)")
    for item in items:
        if "printer" in item.keywords:
            item.add_marker(skip)


def pytest_sessionfinish(session):
    """Clean up .test_tmp after all tests finish (visible proof nothing leaks)."""
    if os.path.isdir(PROJECT_TMP):
        try:
            shutil.rmtree(PROJECT_TMP)
        except OSError:
            pass
