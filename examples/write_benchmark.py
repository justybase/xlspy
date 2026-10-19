"""Reproducible write benchmark: time, peak memory and file size per writer.

Every measurement runs in a fresh worker process (this script re-invoked with
``--worker``), so allocations from one run cannot inflate the next. The
orchestrator repeats each writer ``--iterations`` times, interleaving writers
between iterations, and reports medians.

Writers (``--writers``, comma-separated, default ``all``):

  xlsx-shared, xlsx-inline, xlsb-shared, xlsb-inline
      xlspy, with the C extension on, off, or both (``--c-ext``).
  openpyxl
      openpyxl ``Workbook(write_only=True)``; skipped if not installed.
  xlsxwriter
      xlsxwriter ``Workbook(constant_memory=True)``; skipped if not installed.

What is measured, per run:

  elapsed_s       Wall time from opening the workbook to closing it. With
                  ``--input stream`` (default) the rows are generated while the
                  writer consumes them, so every writer pays the same generation
                  cost. With ``--input list`` the rows are built before the
                  clock starts, and their memory is part of the peak.
  peak_rss_mb     Peak resident set of the worker process (Windows: peak working
                  set; Linux: ru_maxrss). Includes the interpreter and imports;
                  ``baseline_rss_mb`` is the resident set just before the write.
  tracemalloc_mb  With ``--tracemalloc``, a separate run's peak of Python
                  allocations. It excludes memory the C extension allocates
                  directly, and tracing slows the write, so these runs are not
                  timed.
  file_mb         Size of the output file.

Results go to a JSON file (``--output``) with the environment, every run, and
the medians; ``--csv`` also writes one row per run.

Usage:
    python examples/write_benchmark.py --rows 50000 --cols 50 --iterations 3 \\
        --output bench.json
    python examples/write_benchmark.py --writers xlsx-inline,xlsb-inline --c-ext both
"""

import argparse
import csv
import datetime
import json
import os
import platform
import random
import statistics
import subprocess
import sys
import tempfile
import time
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SRC = os.path.normpath(os.path.join(HERE, "..", "src"))
XLSPY_WRITERS = ("xlsx-shared", "xlsx-inline", "xlsb-shared", "xlsb-inline")
OTHER_WRITERS = ("openpyxl", "xlsxwriter")
ALL_WRITERS = XLSPY_WRITERS + OTHER_WRITERS
COMPRESSION_LEVEL = 4  # xlspy's default; openpyxl and xlsxwriter use their own defaults


# ---------------------------------------------------------------- datasets


def mixed_rows(rows, cols, seed):
    """The ``performance_test.py`` layout: eight cell types in rotating columns.

    Deterministic; ``seed`` is not used. The first row is a header.
    """
    yield [f"Header_{j}" for j in range(cols)]
    day = datetime.date(2023, 1, 1)
    moment = datetime.datetime(2023, 1, 1, 12, 30, 0)
    for i in range(1, rows):
        row = []
        for j in range(cols):
            kind = j % 8
            if kind == 0:
                row.append(f"String_{i}_{j}")
            elif kind == 1:
                row.append(i * j)
            elif kind == 2:
                row.append(i / (j + 1))
            elif kind == 3:
                row.append((i + j) % 2 == 0)
            elif kind == 4:
                row.append(day + datetime.timedelta(days=i))
            elif kind == 5:
                row.append(moment + datetime.timedelta(seconds=i))
            elif kind == 6:
                row.append(Decimal(f"{i}.{j}"))
            else:
                row.append(2**60 + i)
        yield row


def random_rows(rows, cols, seed):
    """Seeded text-heavy rows: repeated categories, unique notes and numbers."""
    rng = random.Random(seed)
    categories = [f"Category {k}" for k in range(200)]
    yield [f"Col_{j}" for j in range(cols)]
    for i in range(1, rows):
        row = []
        for j in range(cols):
            kind = j % 4
            if kind == 0:
                row.append(rng.choice(categories))
            elif kind == 1:
                row.append(f"note {i}-{j} {rng.getrandbits(32):08x}")
            elif kind == 2:
                row.append(rng.randrange(1_000_000))
            else:
                row.append(round(rng.uniform(-1e6, 1e6), 4))
        yield row


DATASETS = {"mixed": mixed_rows, "random": random_rows}


# --------------------------------------------------------------- memory


def _peak_rss_bytes():
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        class Counters(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD),
                        ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
                        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
                        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                        ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t)]

        kernel32, psapi = ctypes.WinDLL("kernel32"), ctypes.WinDLL("psapi")
        # Explicit types: the process pseudo-handle is pointer-sized.
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        psapi.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.POINTER(Counters), wintypes.DWORD]
        psapi.GetProcessMemoryInfo.restype = wintypes.BOOL
        counters = Counters()
        counters.cb = ctypes.sizeof(counters)
        if not psapi.GetProcessMemoryInfo(kernel32.GetCurrentProcess(), ctypes.byref(counters), counters.cb):
            raise ctypes.WinError()
        return counters.PeakWorkingSetSize, counters.WorkingSetSize
    import resource

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    peak *= 1 if sys.platform == "darwin" else 1024  # bytes on macOS, KiB on Linux
    current = None
    try:
        with open("/proc/self/statm") as statm:
            current = int(statm.read().split()[1]) * os.sysconf("SC_PAGE_SIZE")
    except OSError:
        pass
    return peak, current


# --------------------------------------------------------------- worker


def _write(writer, path, rows):
    if writer in XLSPY_WRITERS:
        from xlspy import XlsbWriter, XlsxWriter

        fmt, storage = writer.split("-")
        cls = XlsxWriter if fmt == "xlsx" else XlsbWriter
        # Shared strings are the default; passing the option only for inline
        # mode also lets older versions (whose XlsbWriter lacks it) run shared.
        options = {} if storage == "shared" else {"useSharedStrings": False}
        with cls(path, compressionLevel=COMPRESSION_LEVEL, **options) as workbook:
            workbook.add_sheet("Bench")
            workbook.write_sheet(rows)
    elif writer == "openpyxl":
        import openpyxl

        workbook = openpyxl.Workbook(write_only=True)
        sheet = workbook.create_sheet("Bench")
        for row in rows:
            sheet.append(row)
        workbook.save(path)
    elif writer == "xlsxwriter":
        import xlsxwriter

        workbook = xlsxwriter.Workbook(path, {"constant_memory": True})
        sheet = workbook.add_worksheet("Bench")
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                if isinstance(value, Decimal):
                    value = float(value)  # xlsxwriter has no Decimal support
                if isinstance(value, (datetime.date, datetime.datetime)):
                    sheet.write_datetime(r, c, value)
                elif isinstance(value, int) and not isinstance(value, bool) and abs(value) >= 2**53:
                    sheet.write_number(r, c, float(value))
                else:
                    sheet.write(r, c, value)
        workbook.close()
    else:
        raise ValueError(f"unknown writer {writer!r}")


def run_worker(args):
    sys.path.insert(0, args.src)
    result = {"writer": args.writer, "rows": args.rows, "cols": args.cols, "dataset": args.dataset,
              "input": args.input, "target": args.target, "seed": args.seed, "tracemalloc": args.tracemalloc}
    if args.writer in XLSPY_WRITERS:
        from xlspy import _accel

        result["c_ext"] = bool(_accel.HAVE_C_EXT)
    else:
        __import__(args.writer)  # import before the baseline so it is not counted as write memory
    ext = "xlsb" if args.writer.startswith("xlsb") else "xlsx"
    path = os.path.join(args.outdir, f"bench_{os.getpid()}.{ext}")
    if args.target == "memory":
        import io

        path = io.BytesIO()
    rows = DATASETS[args.dataset](args.rows, args.cols, args.seed)
    if args.input == "list":
        rows = list(rows)
    import tracemalloc

    if args.tracemalloc:
        tracemalloc.start()
    _, baseline = _peak_rss_bytes()
    start = time.perf_counter()
    try:
        _write(args.writer, path, rows)
    except Exception as exc:  # reported, not raised: one failure must not stop the series
        result.update(status="error", error=f"{type(exc).__name__}: {exc}")
    else:
        result["status"] = "ok"
    elapsed = time.perf_counter() - start
    peak, _ = _peak_rss_bytes()
    if args.tracemalloc:
        result["tracemalloc_mb"] = tracemalloc.get_traced_memory()[1] / 2**20
        tracemalloc.stop()
    else:
        result["elapsed_s"] = elapsed
    result["peak_rss_mb"] = peak / 2**20
    result["baseline_rss_mb"] = baseline / 2**20 if baseline else None
    if not isinstance(path, str):
        result["file_mb"] = len(path.getvalue()) / 2**20
    elif os.path.exists(path):
        result["file_mb"] = os.path.getsize(path) / 2**20
        if not args.keep:
            os.remove(path)
    print(json.dumps(result))


# --------------------------------------------------------- orchestrator


def _version(module):
    try:
        return __import__(module).__version__
    except Exception:
        return None


def _cpu_name():
    if sys.platform.startswith("linux"):
        try:
            with open("/proc/cpuinfo") as info:
                for line in info:
                    if line.startswith("model name"):
                        return line.split(":", 1)[1].strip()
        except OSError:
            pass
    if sys.platform == "win32":
        try:
            import winreg

            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DESCRIPTION\System\CentralProcessor\0")
            return winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
        except OSError:
            pass
    return platform.processor() or None


def _git(src, *args):
    try:
        out = subprocess.run(["git", *args], cwd=src, capture_output=True, text=True, timeout=10)
        return out.stdout.strip() if out.returncode == 0 else None
    except (OSError, subprocess.SubprocessError):
        return None


def environment(args):
    sys.path.insert(0, args.src)
    import xlspy

    status = _git(args.src, "status", "--porcelain")
    return {
        "date": datetime.date.today().isoformat(),
        "platform": platform.platform(),
        "machine": platform.machine(),
        "cpu": _cpu_name(),
        "logical_cpus": os.cpu_count(),
        "python": sys.version.split()[0],
        "python_implementation": platform.python_implementation(),
        "xlspy": getattr(xlspy, "__version__", None),
        "xlspy_src": args.src,
        "git_commit": _git(args.src, "rev-parse", "HEAD"),
        "git_dirty": None if status is None else bool(status),
        "openpyxl": _version("openpyxl"),
        "xlsxwriter": _version("xlsxwriter"),
        "compression": {"xlspy": f"deflate level {COMPRESSION_LEVEL}",
                        "openpyxl": "library default", "xlsxwriter": "library default"},
        "peak_memory_source": "peak working set" if sys.platform == "win32" else "ru_maxrss",
    }


def _jobs(args):
    writers = ALL_WRITERS if args.writers == "all" else tuple(w.strip() for w in args.writers.split(","))
    unknown = set(writers) - set(ALL_WRITERS)
    if unknown:
        raise SystemExit(f"unknown writers: {', '.join(sorted(unknown))}")
    modes = {"on": ["on"], "off": ["off"], "both": ["on", "off"]}[args.c_ext]
    jobs = []
    for writer in writers:
        if writer in OTHER_WRITERS:
            if _version(writer) is None:
                print(f"skipping {writer}: not installed", file=sys.stderr)
                continue
            if args.target == "memory":
                print(f"skipping {writer}: --target memory covers the xlspy writers only", file=sys.stderr)
                continue
            jobs.append((writer, None))
        else:
            jobs.extend((writer, mode) for mode in modes)
    return jobs


def _run_job(args, writer, mode, tracemalloc, outdir):
    cmd = [sys.executable, os.path.abspath(__file__), "--worker", "--writer", writer, "--rows", str(args.rows),
           "--cols", str(args.cols), "--dataset", args.dataset, "--seed", str(args.seed), "--input", args.input,
           "--target", args.target, "--src", args.src, "--outdir", outdir]
    if tracemalloc:
        cmd.append("--tracemalloc")
    env = dict(os.environ)
    env.pop("XLSPY_DISABLE_C_EXT", None)
    if mode == "off":
        env["XLSPY_DISABLE_C_EXT"] = "1"
    proc = subprocess.run(cmd, env=env, capture_output=True, text=True)
    lines = [line for line in proc.stdout.splitlines() if line.startswith("{")]
    if proc.returncode != 0 or not lines:
        return {"writer": writer, "status": "error", "error": (proc.stderr or proc.stdout)[-500:]}
    result = json.loads(lines[-1])
    if mode is not None and result.get("c_ext") != (mode == "on"):
        result.update(status="error", error=f"C extension state {result.get('c_ext')} does not match --c-ext {mode}")
    return result


def _label(writer, mode):
    return writer if mode is None else f"{writer} (C {mode})"


def summarize(runs):
    summary = {}
    for run in runs:
        entry = summary.setdefault(run["label"], {"runs": 0, "errors": 0, "elapsed_s": [], "peak_rss_mb": [],
                                                  "tracemalloc_mb": [], "file_mb": None})
        if run.get("status") != "ok":
            entry["errors"] += 1
            entry["error"] = run.get("error")
            continue
        entry["runs"] += 1
        for key in ("elapsed_s", "tracemalloc_mb"):
            if run.get(key) is not None:
                entry[key].append(run[key])
        if run.get("elapsed_s") is not None:
            entry["peak_rss_mb"].append(run["peak_rss_mb"])
        entry["file_mb"] = run.get("file_mb", entry["file_mb"])
    for entry in summary.values():
        for key in ("elapsed_s", "peak_rss_mb", "tracemalloc_mb"):
            values = entry.pop(key)
            entry[f"median_{key}"] = statistics.median(values) if values else None
    return summary


def run_orchestrator(args):
    jobs = _jobs(args)
    report = {"environment": environment(args),
              "parameters": {k: getattr(args, k) for k in ("rows", "cols", "iterations", "dataset", "input",
                                                           "target", "seed", "c_ext", "tracemalloc")},
              "runs": []}
    outdir = args.outdir or tempfile.mkdtemp(prefix="xlspy_bench_")
    os.makedirs(outdir, exist_ok=True)
    passes = [False] + ([True] if args.tracemalloc else [])
    for traced in passes:
        for iteration in range(args.iterations if not traced else 1):
            for writer, mode in jobs:
                result = _run_job(args, writer, mode, traced, outdir)
                result.update(label=_label(writer, mode), iteration=iteration)
                report["runs"].append(result)
                shown = (f"{result['tracemalloc_mb']:.1f} MB traced" if traced and "tracemalloc_mb" in result
                         else f"{result.get('elapsed_s', 0):.2f} s, peak {result.get('peak_rss_mb', 0):.1f} MB")
                print(f"[{iteration + 1}/{args.iterations}] {result['label']}: "
                      f"{shown if result.get('status') == 'ok' else 'ERROR ' + str(result.get('error'))[:200]}",
                      file=sys.stderr)
    report["summary"] = summarize(report["runs"])
    if args.output:
        with open(args.output, "w", encoding="utf-8") as out:
            json.dump(report, out, indent=2)
    if args.csv:
        fields = ["label", "iteration", "writer", "c_ext", "status", "rows", "cols", "dataset", "input", "seed",
                  "elapsed_s", "peak_rss_mb", "baseline_rss_mb", "tracemalloc_mb", "file_mb", "error"]
        with open(args.csv, "w", newline="", encoding="utf-8") as out:
            writer = csv.DictWriter(out, fieldnames=fields, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(report["runs"])
    print(f"\n{args.rows:,} x {args.cols} ({args.dataset}, {args.input} input), "
          f"median of {args.iterations} run(s):")
    print(f"{'writer':<26}{'time (s)':>10}{'peak (MB)':>11}{'traced (MB)':>13}{'file (MiB)':>12}")
    for label, entry in report["summary"].items():
        def fmt(value, spec):
            return format(value, spec) if value is not None else "-"

        status = f"  ({entry['errors']} failed: {entry.get('error', '')[:60]})" if entry["errors"] else ""
        print(f"{label:<26}{fmt(entry['median_elapsed_s'], '.2f'):>10}{fmt(entry['median_peak_rss_mb'], '.1f'):>11}"
              f"{fmt(entry['median_tracemalloc_mb'], '.1f'):>13}{fmt(entry['file_mb'], '.2f'):>12}{status}")
    return 1 if any(entry["errors"] for entry in report["summary"].values()) else 0


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--rows", type=int, default=50_000, help="rows including the header (default 50000)")
    parser.add_argument("--cols", type=int, default=50)
    parser.add_argument("--iterations", type=int, default=3, help="timed runs per writer (default 3)")
    parser.add_argument("--writers", default="all", help=f"comma-separated: {', '.join(ALL_WRITERS)}, or all")
    parser.add_argument("--c-ext", choices=["on", "off", "both"], default="on",
                        help="xlspy with the C extension, without it, or both (default on)")
    parser.add_argument("--dataset", choices=sorted(DATASETS), default="mixed")
    parser.add_argument("--seed", type=int, default=0, help="seed for the random dataset (default 0)")
    parser.add_argument("--input", choices=["stream", "list"], default="stream")
    parser.add_argument("--target", choices=["file", "memory"], default="file",
                        help="write to a file path, or to an io.BytesIO kept in memory (default file)")
    parser.add_argument("--tracemalloc", action="store_true", help="add one traced, untimed run per writer")
    parser.add_argument("--outdir", help="directory for the workbooks (default: a new temporary directory)")
    parser.add_argument("--output", help="write the JSON report here")
    parser.add_argument("--csv", help="also write one CSV row per run here")
    parser.add_argument("--src", default=DEFAULT_SRC, help="xlspy source tree to import (default: this checkout)")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--writer", help=argparse.SUPPRESS)
    parser.add_argument("--keep", action="store_true", help="keep the workbooks (worker mode)")
    args = parser.parse_args(argv)
    args.src = os.path.abspath(args.src)
    if args.worker:
        if not args.outdir:
            parser.error("--worker needs --outdir")
        run_worker(args)
        return 0
    return run_orchestrator(args)


if __name__ == "__main__":
    sys.exit(main())
