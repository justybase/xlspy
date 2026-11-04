"""Line and region coverage of the native module (``rust/xlspy_core``).

Builds ``xlspy._rs_core`` with ``-Cinstrument-coverage`` (through
``cargo llvm-cov show-env``), runs the Rust unit tests and the chosen pytest
selection against it, and prints ``cargo llvm-cov report``.

The instrumented module is not built in place: the package is copied to the
output directory with the instrumented module next to it, and pytest runs with
that copy first on ``sys.path``. The release module in ``src/xlspy`` is left
alone, so the editable install keeps working (and keeps its speed) afterwards.

Needs ``cargo install cargo-llvm-cov`` and ``rustup component add
llvm-tools-preview``. Usage::

    python tools/rust_coverage.py                       # tests/ (whole suite)
    python tools/rust_coverage.py tests/test_native_fuzz.py --html
    python tools/rust_coverage.py --no-pytest           # cargo test only
"""
from __future__ import annotations

import argparse
import os
import shlex
import shutil
import subprocess
import sys
import sysconfig
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CRATE = ROOT / "rust" / "xlspy_core"
PACKAGE = ROOT / "src" / "xlspy"


def _cargo_env() -> dict[str, str]:
    """The environment ``cargo llvm-cov`` builds with, on top of ours."""
    out = subprocess.run(
        ["cargo", "llvm-cov", "show-env"], cwd=CRATE, check=True,
        capture_output=True, text=True,
    ).stdout
    env = dict(os.environ)
    for line in out.splitlines():
        key, sep, value = line.partition("=")
        if not sep or not key.isidentifier():
            continue
        # Values are shell-quoted ('...'); shlex keeps Windows backslashes in
        # single quotes.
        parts = shlex.split(value, posix=True)
        env[key] = parts[0] if parts else ""
    env["PYO3_PYTHON"] = sys.executable
    # Rust unit tests that embed Python load libpython / pythonXY.dll.
    base = sys.base_prefix
    if sys.platform == "win32":
        env["PATH"] = base + os.pathsep + env.get("PATH", "")
    else:
        libdir = sysconfig.get_config_var("LIBDIR") or os.path.join(base, "lib")
        key = "DYLD_LIBRARY_PATH" if sys.platform == "darwin" else "LD_LIBRARY_PATH"
        env[key] = libdir + os.pathsep + env.get(key, "")
    return env


def _built_module(target: Path) -> Path:
    release = target / "release"
    for name in ("_rs_core.dll", "lib_rs_core.so", "lib_rs_core.dylib"):
        if (release / name).exists():
            return release / name
    raise SystemExit(f"no instrumented module in {release}")


def _overlay(out: Path, module: Path) -> Path:
    """Copy the package to ``out/pkg`` with the instrumented module in it."""
    pkg_root = out / "pkg"
    if pkg_root.exists():
        shutil.rmtree(pkg_root)
    shutil.copytree(
        PACKAGE, pkg_root / "xlspy",
        ignore=shutil.ignore_patterns("__pycache__", "_rs_core*.pyd", "_rs_core*.so", "*.pdb"),
    )
    suffix = sysconfig.get_config_var("EXT_SUFFIX")
    shutil.copy2(module, pkg_root / "xlspy" / f"_rs_core{suffix}")
    return pkg_root


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("tests", nargs="*", default=["tests"], help="pytest arguments (default: tests); unknown options go to pytest too")
    parser.add_argument("--out", type=Path, default=ROOT / "build" / "rust-coverage")
    parser.add_argument("--html", action="store_true", help="also write an HTML report to OUT/html")
    parser.add_argument("--lcov", action="store_true", help="also write OUT/lcov.info")
    parser.add_argument("--no-pytest", action="store_true", help="only the Rust unit tests")
    parser.add_argument("--no-cargo-test", action="store_true", help="only the pytest run")
    args, extra = parser.parse_known_args()
    args.tests = [*args.tests, *extra]

    env = _cargo_env()
    target = Path(env.get("CARGO_LLVM_COV_TARGET_DIR", CRATE / "target"))
    args.out.mkdir(parents=True, exist_ok=True)

    def cargo(*cmd: str) -> None:
        subprocess.run(["cargo", *cmd], cwd=CRATE, env=env, check=True)

    cargo("llvm-cov", "clean", "--workspace")
    if not args.no_cargo_test:
        cargo("test", "--release", "--locked")
    if not args.no_pytest:
        cargo("build", "--release", "--locked", "--features", "pyo3/extension-module")
        pkg_root = _overlay(args.out, _built_module(target))
        py_env = dict(env)
        py_env["PYTHONPATH"] = str(pkg_root) + os.pathsep + env.get("PYTHONPATH", "")
        py_env.pop("XLSPY_ACCEL", None)
        check = subprocess.run(
            [sys.executable, "-c", "import xlspy, xlspy._accel as a; print(xlspy.__file__, a.HAVE_NATIVE)"],
            cwd=ROOT, env=py_env, check=True, capture_output=True, text=True,
        ).stdout.split()
        if not check[0].startswith(str(pkg_root)) or check[1] != "True":
            raise SystemExit(f"the overlay is not imported: {check}")
        # pyproject.toml puts src/ first (pythonpath = ["src"]): point it at the
        # copy instead, or pytest imports the release module. Failures are
        # reported but do not stop the report.
        subprocess.run([sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider",
                        "-o", f"pythonpath={pkg_root}", *args.tests],
                       cwd=ROOT, env=py_env)

    report = ["llvm-cov", "report", "--release"]
    cargo(*report)
    if args.html:
        cargo(*report, "--html", "--output-dir", str(args.out / "html"))
    if args.lcov:
        cargo(*report, "--lcov", "--output-path", str(args.out / "lcov.info"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
