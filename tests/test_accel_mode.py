"""XLSPY_ACCEL picks the backend at import; bad or removed settings fail loudly."""

import os
import subprocess
import sys
from pathlib import Path

import pytest

SRC = str(Path(__file__).resolve().parents[1] / "src")


def _import_accel(**env_overrides: str) -> subprocess.CompletedProcess:
    env = {k: v for k, v in os.environ.items() if k not in ("XLSPY_ACCEL", "XLSPY_DISABLE_C_EXT")}
    env.update(env_overrides, PYTHONPATH=SRC)
    return subprocess.run(
        [sys.executable, "-c", "from xlspy import _accel; print(_accel.ACCEL_MODE, _accel.HAVE_NATIVE)"],
        env=env, capture_output=True, text=True, encoding="utf-8", check=False,
    )


def test_python_mode_disables_the_native_module():
    proc = _import_accel(XLSPY_ACCEL="python")
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.split() == ["python", "False"]


def test_an_unknown_mode_is_an_import_error():
    proc = _import_accel(XLSPY_ACCEL="c")
    assert proc.returncode != 0
    assert "XLSPY_ACCEL must be one of" in proc.stderr


@pytest.mark.parametrize("value", ["0", "1"])
def test_the_removed_c_extension_switch_is_an_import_error(value):
    proc = _import_accel(XLSPY_DISABLE_C_EXT=value)
    assert proc.returncode != 0
    assert "XLSPY_DISABLE_C_EXT was removed; set XLSPY_ACCEL=python" in proc.stderr
