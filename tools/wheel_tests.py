"""Run a subset of the test suite against an installed xlspy wheel.

The repository's ``pyproject.toml`` puts ``src`` on ``sys.path`` for pytest, which
would test the source tree instead of the wheel. This script copies the chosen
test files (they need only the standard library, pytest and xlspy) into a
temporary directory and runs pytest there, so the installed package is what gets
imported. CI runs it in a clean venv and cibuildwheel runs it on every wheel
target, including those without a full test job (Linux aarch64, Windows ARM64,
Intel macOS).

Usage: ``python tools/wheel_tests.py [REPOSITORY_ROOT]`` (default: this script's
parent directory). Needs pytest; exits with pytest's status.
"""

import pathlib
import shutil
import subprocess
import sys
import tempfile

# Core reading, writing and updating, plus the native-versus-Python checks that
# would catch a platform-specific difference in the compiled extension.
TEST_FILES = (
    "test_reader.py",
    "test_updaters.py",
    "test_streaming_updaters.py",
    "test_output_integrity.py",
    "test_invalid_file_errors.py",
    "test_native_read_correctness.py",
    "test_native_writer_parity.py",
    "test_native_encoders.py",
    "test_native_formula.py",
    "test_native_hyperlinks.py",
    "test_xlsb_native_dates.py",
)
SUPPORT_FILES = ("conftest.py", "zip_utils.py")
# The repository's settings that matter here (pyproject.toml), without its pythonpath.
PYTEST_INI = """\
[pytest]
addopts = -ra --strict-markers
filterwarnings =
    error
markers =
    com: requires Microsoft Excel via COM
    printer: needs a reachable default printer
    libreoffice: converts output with headless LibreOffice
    compat: openpyxl compatibility harness
"""


def main(argv):
    root = pathlib.Path(argv[1]) if len(argv) > 1 else pathlib.Path(__file__).resolve().parent.parent
    tests = root / "tests"
    with tempfile.TemporaryDirectory(prefix="xlspy-wheel-tests-") as directory:
        for name in TEST_FILES + SUPPORT_FILES:
            shutil.copy2(tests / name, directory)
        # Not the repository's pytest configuration: it would import from src/.
        (pathlib.Path(directory) / "pytest.ini").write_text(PYTEST_INI)
        command = [sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", directory]
        return subprocess.call(command, cwd=directory)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
