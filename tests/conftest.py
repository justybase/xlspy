import gc
import os
import shutil

import pytest

PROJECT_TMP = os.path.join(os.path.dirname(__file__), "..", ".test_tmp")


# Marker -> (environment variable that opts in, skip reason). One gate for every
# test that needs Microsoft Excel or a printer, so `-m com` selects all of them.
OPT_IN = {
    "com": ("XLSPY_RUN_EXCEL_COM", "Set XLSPY_RUN_EXCEL_COM=1 (needs Microsoft Excel on Windows)"),
    "printer": ("XLSPY_RUN_EXCEL_PRINT", "Set XLSPY_RUN_EXCEL_PRINT=1 (needs a reachable default printer)"),
}
COMPAT_DIR = os.path.join(os.path.dirname(__file__), "openpyxl_compat") + os.sep


def pytest_collection_modifyitems(config, items):
    """Excel reads and sets PageSetup through the default printer: without
    one it waits, shows a printer prompt and fails, so these tests are opt-in."""
    for item in items:
        if str(item.path).startswith(COMPAT_DIR):
            item.add_marker(pytest.mark.compat)
    for marker, (env, reason) in OPT_IN.items():
        if os.environ.get(env) == "1":
            continue
        skip = pytest.mark.skip(reason=reason)
        for item in items:
            if marker in item.keywords:
                item.add_marker(skip)


def pytest_sessionfinish(session):
    """Clean up .test_tmp after all tests finish (visible proof nothing leaks)."""
    if os.path.isdir(PROJECT_TMP):
        try:
            shutil.rmtree(PROJECT_TMP)
        except OSError:
            pass


# Microsoft Excel through COM, shared by the tests marked `com`.
@pytest.fixture(scope="module")
def excel():
    client = pytest.importorskip("win32com.client")
    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    yield app
    gc.collect()
    app.Quit()
    del app
    gc.collect()


@pytest.fixture
def open_workbook(excel):
    opened = []
    def open_file(path):
        workbook = excel.Workbooks.Open(str(path), 0, True)
        opened.append(workbook)
        return workbook
    yield open_file
    while opened:
        workbook = opened.pop()
        try:
            workbook.Close(False)
        finally:
            del workbook
    gc.collect()
