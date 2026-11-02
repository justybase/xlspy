"""Every public callable exported by ``xlspy`` and ``xlspy.workbook`` is fully annotated.

mypy only checks callers against annotations that exist; an unannotated public
method is silently ``Any`` for users. This guard keeps new API typed.
"""
from __future__ import annotations

import inspect
from collections.abc import Callable, Iterator
from typing import Any

import xlspy
import xlspy.workbook

# Private names are skipped, except the protocol methods users call through syntax.
_DUNDERS = frozenset({
    "__init__", "__enter__", "__exit__", "__iter__", "__getitem__", "__setitem__",
    "__len__", "__contains__",
})


def _public_functions() -> Iterator[tuple[str, Callable[..., Any]]]:
    for module in (xlspy, xlspy.workbook):
        for export in sorted(module.__all__):
            obj = getattr(module, export)
            if inspect.isclass(obj):
                # Walk the MRO: writer methods live on mixin bases.
                members = {
                    member: value
                    for base in reversed(obj.__mro__)
                    if base.__module__.startswith("xlspy")
                    for member, value in vars(base).items()
                }
                for member, value in members.items():
                    if member.startswith("_") and member not in _DUNDERS:
                        continue
                    if isinstance(value, (staticmethod, classmethod)):
                        value = value.__func__
                    if isinstance(value, property):
                        for accessor in (value.fget, value.fset):
                            if accessor is not None:
                                yield f"{module.__name__}.{export}.{member}", accessor
                    elif callable(value):
                        yield f"{module.__name__}.{export}.{member}", value
            elif inspect.isfunction(obj):
                yield f"{module.__name__}.{export}", obj


def _missing_annotations(function: Callable[..., Any]) -> list[str]:
    function = inspect.unwrap(function)
    if not inspect.isfunction(function) or not function.__module__.startswith("xlspy"):
        return []
    annotations = function.__annotations__
    missing = [
        name for name in inspect.signature(function).parameters
        if name not in ("self", "cls") and name not in annotations
    ]
    if "return" not in annotations:
        missing.append("return")
    return missing


def test_public_api_is_fully_annotated() -> None:
    problems = sorted({
        f"{name}: {', '.join(missing)}"
        for name, function in _public_functions()
        if (missing := _missing_annotations(function))
    })
    assert not problems, "unannotated public API:\n" + "\n".join(problems)


def test_guard_sees_the_public_api() -> None:
    names = {name for name, _ in _public_functions()}
    assert "xlspy.XlsxWriter.write_sheet" in names
    assert "xlspy.workbook.Worksheet.iter_rows" in names
    assert "xlspy.load_workbook" in names
