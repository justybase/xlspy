"""Record a local Microsoft Excel validation run as a JSON artifact.

CI cannot run Excel, so the COM suite is opt-in on a Windows machine with
Excel installed.  This tool runs it, records the Excel build, the commit and
the test counts, and writes ``validation/excel-validation-<date>-<sha>.json``.
``--check`` verifies the newest artifact before a release.

Examples::

    python tools/excel_validation.py                 # pytest -m com
    python tools/excel_validation.py tests/test_excel_com_updaters.py
    python tools/excel_validation.py --include-printer
    python tools/excel_validation.py --check

The tool needs ``pywin32`` (``pip install pywin32``) and Microsoft Excel.  It
does not run in CI.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gc
import json
import os
import platform
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT = ROOT / "validation"


def _run(command: list[str], *, check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(command, cwd=ROOT, text=True, capture_output=True, check=check)


def _git_commit() -> str:
    return _run(["git", "rev-parse", "HEAD"]).stdout.strip()


def _git_dirty() -> bool:
    return bool(_run(["git", "status", "--porcelain"]).stdout.strip())


def _git_is_ancestor(older: str) -> bool:
    return _run(["git", "merge-base", "--is-ancestor", older, "HEAD"], check=False).returncode == 0


def _git_changed_since(older: str) -> list[str]:
    result = _run(["git", "diff", "--name-only", f"{older}..HEAD"], check=False)
    return [line for line in result.stdout.splitlines() if line]


def _excel_info() -> dict[str, str]:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SystemExit("pywin32 is required: pip install pywin32") from exc
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        return {
            "version": str(app.Version),
            "build": str(app.Build),
            "operating_system": str(app.OperatingSystem),
        }
    finally:
        try:
            app.Quit()
        finally:
            del app
            gc.collect()


def _junit_counts(path: Path) -> dict[str, int | float]:
    root = ET.parse(path).getroot()
    suite = root if root.tag == "testsuite" else root.find("testsuite")
    if suite is None:
        raise SystemExit(f"no testsuite element in {path}")
    return {
        "tests": int(suite.get("tests", 0)),
        "failures": int(suite.get("failures", 0)),
        "errors": int(suite.get("errors", 0)),
        "skipped": int(suite.get("skipped", 0)),
        "seconds": round(float(suite.get("time", 0.0)), 3),
    }


def _run_pytest(args: list[str]) -> tuple[int, dict[str, int | float]]:
    with tempfile.TemporaryDirectory(prefix="xlspy-excel-validation-") as tmp:
        junit = Path(tmp) / "junit.xml"
        command = [sys.executable, "-m", "pytest", "-q", f"--junitxml={junit}", *args]
        env = os.environ.copy()
        env["XLSPY_RUN_EXCEL_COM"] = "1"
        started = time.perf_counter()
        completed = subprocess.run(command, cwd=ROOT, text=True, env=env)
        seconds = round(time.perf_counter() - started, 3)
        counts = _junit_counts(junit) if junit.exists() else {
            "tests": 0, "failures": 0, "errors": 0, "skipped": 0, "seconds": seconds
        }
        counts["wall_seconds"] = seconds
        return completed.returncode, counts


def _artifact(commit: str, dirty: bool, excel: dict[str, str], args: list[str],
              result: dict, exit_code: int) -> dict:
    return {
        "tool": "tools/excel_validation.py",
        "created": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "excel": excel,
        "git": {"commit": commit, "dirty": dirty},
        "xlspy": _xlspy_version(),
        "python": platform.python_version(),
        "platform": sys.platform,
        "command": ["pytest", *args],
        "result": {**result, "exit_code": exit_code},
    }


def _xlspy_version() -> str:
    sys.path.insert(0, str(ROOT / "src"))
    try:
        import xlspy
        return xlspy.__version__
    finally:
        sys.path.pop(0)


def _write(out_dir: Path, artifact: dict) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    date = artifact["created"][:10]
    sha = artifact["git"]["commit"][:7]
    path = out_dir / f"excel-validation-{date}-{sha}.json"
    path.write_text(json.dumps(artifact, indent=2) + "\n", encoding="utf-8")
    return path


def _newest(out_dir: Path) -> Path | None:
    paths = sorted(out_dir.glob("excel-validation-*.json"))
    if not paths:
        return None
    return max(paths, key=lambda path: json.loads(path.read_text(encoding="utf-8"))["created"])


def _check(out_dir: Path, max_age_days: int) -> int:
    path = _newest(out_dir)
    if path is None:
        print(f"no validation artifact in {out_dir}", file=sys.stderr)
        return 1
    artifact = json.loads(path.read_text(encoding="utf-8"))
    created = dt.datetime.fromisoformat(artifact["created"])
    age = (dt.datetime.now(dt.UTC) - created).days
    problems = []
    if age > max_age_days:
        problems.append(f"artifact is {age} days old (limit {max_age_days})")
    if artifact["git"]["dirty"]:
        problems.append("artifact was recorded from a dirty working tree")
    if not _git_is_ancestor(artifact["git"]["commit"]):
        problems.append("artifact commit is not an ancestor of HEAD")
    else:
        changed = _git_changed_since(artifact["git"]["commit"])
        allowed = ("docs/", "validation/", "CHANGELOG.md")
        unexpected = [name for name in changed if not name.startswith(allowed)]
        if unexpected:
            problems.append("code changed since validation: " + ", ".join(unexpected))
    if artifact["result"]["exit_code"] != 0:
        problems.append(f"validation exited {artifact['result']['exit_code']}")
    if artifact["result"]["tests"] == 0:
        problems.append("the validation ran no tests")
    if artifact["result"]["failures"] or artifact["result"]["errors"]:
        problems.append(
            f"{artifact['result']['failures']} failures, "
            f"{artifact['result']['errors']} errors"
        )
    print(f"artifact: {path}")
    print(f"excel: {artifact['excel']['version']} ({artifact['excel']['build']})")
    print(f"commit: {artifact['git']['commit'][:12]} (age {age} days)")
    if problems:
        for problem in problems:
            print(f"FAIL: {problem}", file=sys.stderr)
        return 1
    print("OK")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="check the newest artifact")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help=f"artifact directory (default: {DEFAULT_OUT})")
    parser.add_argument("--max-age-days", type=int, default=90, help="--check age limit (default: 90)")
    parser.add_argument("--include-printer", action="store_true", help="also run tests marked printer")
    parser.add_argument("pytest_args", nargs=argparse.REMAINDER, help="pytest arguments (default: -m com)")
    args = parser.parse_args()

    if args.check:
        return _check(args.out, args.max_age_days)
    if sys.platform != "win32":
        print("Excel validation needs Windows and a desktop Excel installation", file=sys.stderr)
        return 1

    pytest_args = list(args.pytest_args)
    if pytest_args[:1] == ["--"]:
        pytest_args = pytest_args[1:]
    if not pytest_args:
        pytest_args = ["-m", "com"]
    if args.include_printer:
        os.environ["XLSPY_RUN_EXCEL_PRINT"] = "1"

    commit = _git_commit()
    dirty = _git_dirty()
    excel = _excel_info()
    exit_code, result = _run_pytest(pytest_args)
    artifact = _artifact(commit, dirty, excel, pytest_args, result, exit_code)
    path = _write(args.out, artifact)
    passed = result["tests"] - result["failures"] - result["errors"] - result["skipped"]
    print(
        f"{path.name}: Excel {excel['version']} ({excel['build']}), "
        f"{passed} passed, {result['failures']} failed, "
        f"{result['errors']} errors, {result['skipped']} skipped"
    )
    print(
        f"| {artifact['created'][:10]} | {excel['version']} ({excel['build']}) | "
        f"`{commit[:7]}` | {passed} passed, {result['failures']} failed, "
        f"{result['skipped']} skipped |"
    )
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
