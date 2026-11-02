"""Bounded mutation fuzzing of the native worksheet readers.

Each case takes a small workbook written by xlspy, corrupts its worksheet part,
and reads the result back through the native path. Corrupt input may be refused,
and then only with InvalidFileError. It must never panic (ExcelReader reports a
Rust panic as InvalidFileError, so the test looks at the cause) and must never
take long. The run is deterministic per seed.

``XLSPY_FUZZ_ITERATIONS`` and ``XLSPY_FUZZ_SEED`` change the budget and the seed
for deeper local runs.
"""

import os
import random
import time
import zipfile

import pytest

from xlspy import ExcelReader, XlsbWriter, XlsxWriter, _accel
from xlspy.exceptions import InvalidFileError

pytestmark = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")

ITERATIONS = int(os.environ.get("XLSPY_FUZZ_ITERATIONS", "300"))
SEED = int(os.environ.get("XLSPY_FUZZ_SEED", "7"))
# A single corrupt case that takes longer than this counts as a hang.
CASE_SECONDS = 5.0

ROWS = [
    ["Name", "Value", "Flag"],
    ["alpha", 1.5, "x"],
    [None, 2, "beta"],
    ["gamma", -3, None],
]


def _write_fixture(path, kind):
    writer = XlsbWriter if kind == "xlsb" else XlsxWriter
    with writer(str(path)) as workbook:
        workbook.add_sheet("Data")
        workbook.write_sheet(ROWS)


def _sheet_part(path):
    with zipfile.ZipFile(path) as archive:
        return next(name for name in archive.namelist() if name.startswith("xl/worksheets/"))


def _mutate(data, rng):
    out = bytearray(data)
    operation = rng.randrange(5)
    if operation == 0 and out:
        for _ in range(rng.randint(1, 8)):
            out[rng.randrange(len(out))] = rng.randrange(256)
    elif operation == 1:
        del out[rng.randrange(len(out) + 1):]
    elif operation == 2:
        at = rng.randrange(len(out) + 1)
        out[at:at] = bytes(rng.randrange(256) for _ in range(rng.randint(1, 16)))
    elif operation == 3 and out:
        start = rng.randrange(len(out))
        end = min(len(out), start + rng.randint(1, 256))
        at = rng.randrange(len(out) + 1)
        out[at:at] = out[start:end]
    elif out:
        # Large counts and ids: a run of 0xFF.
        start = rng.randrange(len(out))
        end = min(len(out), start + rng.randint(1, 8))
        out[start:end] = b"\xff" * (end - start)
    return bytes(out)


def _rebuild(source, target, part, payload):
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = payload if info.filename == part else src.read(info.filename)
            dst.writestr(info.filename, data)


def _read_all(path):
    with ExcelReader(str(path)) as reader:
        sheet = reader.get_sheet_names()[0]
        return sum(1 for _ in reader.get_rows(sheet))


@pytest.mark.parametrize("kind", ["xlsb", "xlsx"])
def test_mutated_worksheets_fail_cleanly(tmp_path, kind):
    source = tmp_path / f"seed.{kind}"
    _write_fixture(source, kind)
    part = _sheet_part(source)
    original = zipfile.ZipFile(source).read(part)
    target = tmp_path / f"case.{kind}"
    rng = random.Random(f"{SEED}:{kind}")

    slowest = 0.0
    for _ in range(ITERATIONS):
        _rebuild(source, target, part, _mutate(original, rng))
        started = time.perf_counter()
        try:
            _read_all(target)
        except InvalidFileError as exc:
            cause = type(exc.__cause__).__name__
            assert cause != "PanicException", f"the native {kind} reader panicked: {exc}"
        slowest = max(slowest, time.perf_counter() - started)

    assert slowest < CASE_SECONDS, f"a corrupt {kind} worksheet took {slowest:.1f}s to read"
