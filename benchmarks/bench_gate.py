"""Performance gate for the streaming write and read paths.

    python benchmarks/bench_gate.py [--rows 100000] [--cols 5] [--repeat 3]

Runs ``bench_io.py`` twice on the same machine, once with the native
accelerator and once with ``XLSPY_ACCEL=python``, and checks three things:

* the native extension is actually loaded;
* each native case is no slower than its pure-Python counterpart (the
  ``--max-ratio`` budget, 10% by default). Comparing the two backends in one
  run cancels most machine-to-machine variation, so this gate catches native
  paths that silently fell back to Python;
* each native case stays below a generous absolute ceiling
  (``--max-seconds``), which catches order-of-magnitude regressions.

An optional ``--baseline`` JSON (as written by ``bench_io.py --out``) adds an
absolute comparison with ``--max-slowdown``; use a baseline measured on the
same runner class, because shared CI machines are not comparable across
operating systems.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
BENCH = os.path.join(ROOT, "benchmarks", "bench_io.py")


def run_benchmark(args: argparse.Namespace, accel: str | None, out: str) -> dict:
    env = dict(os.environ)
    env.pop("XLSPY_ACCEL", None)
    if accel is not None:
        env["XLSPY_ACCEL"] = accel
    completed = subprocess.run(
        [
            sys.executable,
            BENCH,
            "--rows",
            str(args.rows),
            "--cols",
            str(args.cols),
            "--repeat",
            str(args.repeat),
            "--only",
            ",".join(args.cases),
            "--out",
            out,
        ],
        capture_output=True,
        text=True,
        env=env,
    )
    if completed.returncode:
        sys.stderr.write(completed.stdout)
        sys.stderr.write(completed.stderr)
        raise SystemExit(f"bench_io.py failed with XLSPY_ACCEL={accel or 'native'}")
    with open(out, encoding="utf-8") as handle:
        return json.load(handle)


def load_baseline(path: str) -> dict:
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)["results"]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=100_000)
    parser.add_argument("--cols", type=int, default=5)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument(
        "--cases",
        default="write_xlsb,read_xlsb,read_xlsx",
        help="comma-separated exact bench_io.py case names",
    )
    parser.add_argument("--max-ratio", type=float, default=1.10)
    parser.add_argument("--max-seconds", type=float, default=60.0)
    parser.add_argument("--baseline")
    parser.add_argument("--max-slowdown", type=float, default=1.10)
    args = parser.parse_args()
    args.cases = [case.strip() for case in args.cases.split(",") if case.strip()]
    if not args.cases:
        parser.error("--cases must name at least one case")

    with tempfile.TemporaryDirectory(prefix="xlspy-bench-gate-") as directory:
        native = run_benchmark(args, None, os.path.join(directory, "native.json"))
        pure = run_benchmark(args, "python", os.path.join(directory, "python.json"))

    failures: list[str] = []
    if not native["accel"]["have_native"]:
        failures.append("the native extension was not loaded in the native run")
    if pure["accel"]["have_native"]:
        failures.append("XLSPY_ACCEL=python still loaded the native extension")
    baseline = load_baseline(args.baseline) if args.baseline else {}

    print(f"{'case':20} {'native':>9} {'python':>9} {'ratio':>7}  result")
    for case in args.cases:
        native_case = native["results"].get(case, {})
        pure_case = pure["results"].get(case, {})
        if "seconds" not in native_case or "seconds" not in pure_case:
            failures.append(f"{case}: missing timing from bench_io.py")
            print(f"{case:20} {'-':>9} {'-':>9} {'-':>7}  FAILED (no timing)")
            continue
        seconds = native_case["seconds"]
        ratio = seconds / pure_case["seconds"]
        problems = []
        if ratio > args.max_ratio:
            problems.append(f"slower than the fallback (>{args.max_ratio:.2f}x)")
        if seconds > args.max_seconds:
            problems.append(f"over the {args.max_seconds:.0f}s ceiling")
        if case in baseline:
            slowdown = seconds / baseline[case]["seconds"]
            if slowdown > args.max_slowdown:
                problems.append(f"{slowdown:.2f}x over the baseline")
        result = "ok" if not problems else "FAILED: " + "; ".join(problems)
        failures.extend(f"{case}: {problem}" for problem in problems)
        print(f"{case:20} {seconds:9.3f} {pure_case['seconds']:9.3f} {ratio:7.2f}x  {result}")

    if failures:
        print("\nperformance gate failed:", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure}", file=sys.stderr)
        return 1
    print("\nperformance gate passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
