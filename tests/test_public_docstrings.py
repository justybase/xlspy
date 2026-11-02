"""The public API is documented: every exported class and function, and every public
method and property, has a docstring. The API reference is built from them.

``MINIMUM_SHARE`` is the floor for methods and properties; lower it only on purpose.
"""
from __future__ import annotations

import inspect

import xlspy
import xlspy.workbook

MINIMUM_SHARE = 1.0


def _exported() -> dict[str, object]:
    names = {}
    for module in (xlspy, xlspy.workbook):
        for name in module.__all__:
            obj = getattr(module, name)
            if inspect.isclass(obj) or inspect.isfunction(obj):
                names[f"{module.__name__}.{name}"] = obj
    return names


def _members(cls: type) -> dict[str, object]:
    """Public methods and properties that xlspy defines on ``cls`` or its xlspy bases."""
    found = {}
    for klass in cls.__mro__:
        if not klass.__module__.startswith("xlspy"):
            continue
        for name, value in vars(klass).items():
            if name.startswith("_") or name in found:
                continue
            if isinstance(value, (staticmethod, classmethod)):
                value = value.__func__
            if inspect.isfunction(value) or isinstance(value, property):
                found[name] = value
    return found


def _undocumented() -> tuple[list[str], list[str], int]:
    missing_objects, missing_members, members = [], [], 0
    seen: set[int] = set()
    seen_members: set[int] = set()  # a base-class method counts once, not once per subclass
    for qualname, obj in sorted(_exported().items()):
        if id(obj) in seen:
            continue
        seen.add(id(obj))
        if not inspect.getdoc(obj):
            missing_objects.append(qualname)
        if inspect.isclass(obj):
            for name, member in sorted(_members(obj).items()):
                if id(member) in seen_members:
                    continue
                seen_members.add(id(member))
                members += 1
                doc = member.__doc__ if not isinstance(member, property) else (member.fget and member.fget.__doc__)
                if not doc:
                    missing_members.append(f"{qualname}.{name}")
    return missing_objects, missing_members, members


def test_every_exported_class_and_function_has_a_docstring() -> None:
    missing, _, _ = _undocumented()
    assert not missing, "no docstring: " + ", ".join(missing)


def test_public_methods_are_documented() -> None:
    _, missing, total = _undocumented()
    share = 1 - len(missing) / total
    assert share >= MINIMUM_SHARE, (
        f"{share:.0%} of {total} public methods and properties have a docstring; "
        f"undocumented: {', '.join(missing[:40])}{' ...' if len(missing) > 40 else ''}"
    )
