"""No two modules of the package define the same top-level name.

A name defined in two modules is two different things to a reader, and the
copies drift apart. Shared helpers are imported from one owner module instead.
Imports do not count: re-exporting a name (``from .a import X as X``) is fine.
There are no exceptions: the model classes behind ``ws.protection`` and
``ws.page_margins`` are ``WorksheetProtection`` and ``WorksheetPageMargins``.
"""

import ast
import pathlib

PACKAGE = pathlib.Path(__file__).resolve().parents[1] / "src" / "xlspy"

ALLOWED_PER_MODULE = {"__all__"}  # every module has its own export list


def _top_level_names(tree: ast.Module) -> list[str]:
    names = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.append(node.name)
        elif isinstance(node, ast.Assign):
            names.extend(t.id for t in node.targets if isinstance(t, ast.Name))
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.append(node.target.id)
    return names


def test_no_top_level_name_is_defined_in_two_modules():
    owners: dict[str, set[str]] = {}
    for path in sorted(PACKAGE.rglob("*.py")):
        if "__pycache__" in path.parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        rel = path.relative_to(PACKAGE).as_posix()
        for name in _top_level_names(tree):
            if name in ALLOWED_PER_MODULE:
                continue
            owners.setdefault(name, set()).add(rel)

    duplicates = {
        name: sorted(files)
        for name, files in owners.items()
        if len(files) > 1
    }
    assert duplicates == {}, f"defined in more than one module: {duplicates}"
