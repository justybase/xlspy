# Contributing to xlspy

[Home](README.md) · [Guide](docs/usage.md) · [Architecture](docs/architecture.md) · [Tests](tests/README.md)

xlspy is a Python package with an optional native module, `xlspy._rs_core`, written in
Rust with PyO3. Every feature works without it; the module only speeds up the hot paths
(row encoding, worksheet and shared-string reading, column widths).

## Set up

You need Python 3.11 or later and a Rust toolchain (`rustup`, Rust 1.88 or later).

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[test]'      # Windows: .venv\Scripts\python
```

The editable install builds the native module with maturin, in the release profile
(`[tool.maturin]` in `pyproject.toml`). Check which backend loaded:

```bash
python -c "import xlspy._accel as a; print(a.HAVE_NATIVE, a.ACCEL_BACKEND)"
```

`XLSPY_ACCEL=python` turns the native module off; `auto` (the default) and `rust` use it
when it is built.

### Rebuilding the native module

After a change under `rust/xlspy_core`, rebuild it before running the tests:

```bash
python -m pip install maturin
PYO3_PYTHON="$(pwd)/.venv/bin/python" maturin develop --release
```

Things that go wrong, mostly on Windows:

- **The wrong Python.** Set `PYO3_PYTHON` for the build command only, never with
  `export` in a shell that later installs into another Python version: the module
  would link the wrong `python3XY.dll` and fail with "DLL load failed".
- **Stale modules.** Each Python version you build for leaves its own
  `src/xlspy/_rs_core.cp3XY-*.pyd` (or `.so`). Python loads the one for its own
  version, so after a Rust change rebuild for every version you test with, or delete
  the old files.
- **A locked module.** A running process that imported the module (a test run, a
  REPL, a fuzzer) locks the `.pyd`, and the rebuild fails. Close it first.
- **Long paths.** A build in a deep directory can fail with `LNK1104`; point
  `CARGO_TARGET_DIR` at a short path.
- **Debug builds.** `maturin develop` without `--release` makes a module several
  times slower. Never benchmark one.

## Run the checks

Run the tests on both backends; CI does too. `tests/README.md` describes the suites,
the Excel COM tests (`XLSPY_RUN_EXCEL_COM=1`, Windows with Excel), the openpyxl
goldens and coverage.

```bash
python -m pytest
XLSPY_ACCEL=python python -m pytest
```

Run one pytest session at a time: every session shares `.test_tmp` and deletes it when
it ends.

Lint, types and the native module's stub:

```bash
ruff check src tests tools benchmarks examples
mypy
python -m mypy.stubtest xlspy._rs_core        # after changing the Rust API or _rs_core.pyi
```

The Rust crate (from `rust/xlspy_core`):

```bash
cargo fmt --check
cargo clippy --release --locked --all-targets -- -D warnings
cargo test --release --locked
```

Every `unsafe` block needs a `// SAFETY:` comment that says why it is sound; the
crate's lints reject one without it.

Coverage of the Rust code, from its unit tests and the pytest suite (needs
`cargo install cargo-llvm-cov` and `rustup component add llvm-tools-preview`; CI's
`rust-coverage` job runs it too). It builds an instrumented module into a copy of the
package, so the installed one is left as it was:

```bash
python tools/rust_coverage.py                 # the whole suite; --html for a browsable report
python tools/rust_coverage.py tests/test_native_fuzz_differential.py
```

The fuzzers (`tests/test_native_fuzz.py`, and `tests/test_native_fuzz_differential.py`,
which reads each damaged part through both backends and wants the same rows) run 300
cases per part in CI. Before changing a reader, run them deeper:

```bash
XLSPY_FUZZ_ITERATIONS=5000 python -m pytest tests/test_native_fuzz*.py
```

The documentation site, when you change `docs/` or a docstring:

```bash
python -m pip install -e '.[docs]'
mkdocs build --strict        # or: mkdocs serve
```

## Rules the tests enforce

- **Docstrings.** Every exported class and function, and every public method and
  property, has a docstring (`tests/test_public_docstrings.py`). The API reference
  (`docs/api/`) is generated from them. Write plain sentences and use double
  backticks for code; Sphinx roles such as `:meth:` are not rendered.
- **Types.** A new module is fully annotated (`disallow_untyped_defs`).
- **Errors.** A damaged file raises `InvalidFileError` on every read path, with the
  original error as its `__cause__` (`tests/test_invalid_file_errors.py`).
- **Layers.** Writers, readers and updaters do not import `xlspy.workbook`; the object
  model reaches the updaters only through `PackageHandle`. See
  [Architecture](docs/architecture.md).
- **One meaning per name.** No two modules define the same top-level name
  (`tests/test_no_duplicate_top_level_names.py`).
- **Links.** Every relative link and anchor in the Markdown files resolves
  (`tests/test_docs_links.py`).

## Refactoring and performance

A change that should not alter behavior must leave the files xlspy writes identical
byte for byte; compare the ZIP entries before and after. For a change to a hot path
(row encoding, the readers, `load_workbook`), compare the benchmarks:

```bash
python benchmarks/bench_io.py --out before.json
# ... make the change, rebuild ...
python benchmarks/bench_io.py --compare before.json --max-slowdown 1.03
```

## Commits and the changelog

- One topic per commit. The subject says what changed in a short sentence
  ("Split the Python XLSX and XLSB row readers"); the body says why, and what was
  checked.
- Add user-visible changes to `[Unreleased]` in `CHANGELOG.md`, under Breaking
  changes, Added, Changed, Removed, Fixed, Security or Internal. A breaking change gets a row in the
  migration table.
