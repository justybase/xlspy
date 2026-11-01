"""No two modules of the package define the same top-level name.

A name defined in two modules is two different things to a reader, and the
copies drift apart. Shared helpers are imported from one owner module instead.
Imports do not count: re-exporting a name (``from .a import X as X``) is fine.
"""

import ast
import pathlib

import pytest

PACKAGE = pathlib.Path(__file__).resolve().parents[1] / "src" / "xlspy"

# Public names that two models define on purpose. Each is exported from the
# package under the same name, so renaming either one is an API decision.
PUBLIC_DUPLICATES = {
    # xlspy's own protection model (``xlspy.SheetProtection``) and the
    # openpyxl-shaped one on ``Worksheet.protection``.
    "SheetProtection",
    # ``xlspy.PageMargins`` (the historical model) and ``ws.page_margins``.
    "PageMargins",
    # ``xlspy.outline.Outline`` (the writer's outline index) and the
    # openpyxl-shaped ``outlinePr`` on sheet properties.
    "Outline",
}
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
        if len(files) > 1 and name not in PUBLIC_DUPLICATES
    }
    assert duplicates == {}, f"defined in more than one module: {duplicates}"


@pytest.mark.parametrize("name", sorted(PUBLIC_DUPLICATES))
def test_public_duplicates_are_still_two_objects(name):
    """The allow-list must not hide a duplicate that has since become a re-export."""
    owners = []
    for path in sorted(PACKAGE.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        if any(
            isinstance(node, (ast.ClassDef, ast.FunctionDef)) and node.name == name
            for node in tree.body
        ):
            owners.append(path.relative_to(PACKAGE).as_posix())
    assert len(owners) == 2, f"{name} is now defined in {owners}; drop it from PUBLIC_DUPLICATES"
