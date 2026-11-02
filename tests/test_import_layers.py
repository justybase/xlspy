"""Import layering of the package.

Writers, readers and updaters are the lower layer; the object model
(``xlspy.workbook``) is built on them and loaded only when it is used. Imports
inside functions are kept for real reasons (optional or lazy parts, the model's
own cycles) and are counted so new ones are a deliberate choice.
"""
from __future__ import annotations

import ast
import pathlib
import subprocess
import sys

import xlspy

ROOT = pathlib.Path(xlspy.__file__).parent

# Raise only with a reason in the import's comment; lower it when one goes away.
MAX_FUNCTION_LEVEL_IMPORTS = 21


def _module(path: pathlib.Path) -> str:
    parts = list(path.relative_to(ROOT.parent).with_suffix("").parts)
    if parts[-1] == "__init__":
        parts.pop()
    return ".".join(parts)


def _imports(path: pathlib.Path) -> list[tuple[str, int, bool]]:
    """``(target module, line, inside a function)`` for every runtime import."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    parents = {child: node for node in ast.walk(tree) for child in ast.iter_child_nodes(node)}
    module = _module(path)
    package = module.split(".") if path.name == "__init__.py" else module.split(".")[:-1]
    found = []
    for node in ast.walk(tree):
        if not isinstance(node, (ast.Import, ast.ImportFrom)):
            continue
        in_function = type_checking = False
        parent = parents.get(node)
        while parent is not None:
            in_function |= isinstance(parent, (ast.FunctionDef, ast.AsyncFunctionDef))
            type_checking |= isinstance(parent, ast.If) and "TYPE_CHECKING" in ast.unparse(parent.test)
            parent = parents.get(parent)
        if type_checking:
            continue
        if isinstance(node, ast.Import):
            targets = [alias.name for alias in node.names]
        elif node.level:
            base = ".".join(package[: len(package) - node.level + 1] + ([node.module] if node.module else []))
            targets = [base] + [f"{base}.{alias.name}" for alias in node.names]
        else:
            targets = [node.module or ""]
        found.extend((target, node.lineno, in_function) for target in targets)
    return found


def _sources() -> list[pathlib.Path]:
    return sorted(ROOT.rglob("*.py"))


def test_lower_layers_do_not_import_the_object_model() -> None:
    offenders = [
        f"{_module(path)}:{line} imports {target}"
        for path in _sources()
        if not _module(path).startswith("xlspy.workbook") and _module(path) != "xlspy"
        for target, line, _ in _imports(path)
        if target == "xlspy.workbook" or target.startswith("xlspy.workbook.")
    ]
    assert not offenders, "\n".join(offenders)


def test_function_level_imports_stay_few() -> None:
    local = sorted(
        f"{_module(path)}:{line} {target}"
        for path in _sources()
        for target, line, in_function in _imports(path)
        if in_function and not target.startswith(f"{_module(path)}.")
    )
    # One import statement can name several targets; count statements.
    statements = {entry.rsplit(" ", 1)[0] for entry in local}
    assert len(statements) <= MAX_FUNCTION_LEVEL_IMPORTS, "\n".join(local)


def test_import_xlspy_leaves_the_object_model_unloaded() -> None:
    code = (
        "import sys, xlspy\n"
        "loaded = sorted(m for m in sys.modules if m.startswith(('xlspy.workbook', 'xlspy.formula.')) or m == 'xlspy.formula')\n"
        "print(','.join(loaded))\n"
    )
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == ""


def test_updaters_do_not_load_the_object_model() -> None:
    code = (
        "import sys, xlspy.xlsx_updater, xlspy.xlsb_updater, xlspy._xlsb_cf, xlspy._xlsb_table_part\n"
        "print(any(m.startswith('xlspy.workbook') for m in sys.modules))\n"
    )
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    assert result.stdout.strip() == "False"
