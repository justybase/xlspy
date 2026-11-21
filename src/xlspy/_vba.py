"""Reading a VBA project for the writers and the workbook model."""

from __future__ import annotations

import os
from typing import Any


def read_vba_project(project: Any) -> bytes:
    """Return the bytes of a ``vbaProject.bin`` given as a path, bytes or stream."""
    if isinstance(project, (str, os.PathLike)):
        with open(project, "rb") as source:
            data = source.read()
    elif isinstance(project, (bytes, bytearray, memoryview)):
        data = bytes(project)
    elif callable(getattr(project, "read", None)):
        data = project.read()
        if not isinstance(data, (bytes, bytearray, memoryview)):
            raise TypeError("a VBA project stream must be opened in binary mode")
        data = bytes(data)
    else:
        raise TypeError("a VBA project must be a path, bytes or a binary file object")
    if not data:
        raise ValueError("the VBA project is empty")
    return data


def check_vba_name(name: Any) -> str:
    """Validate the VBA project name used as the workbook code name."""
    if not isinstance(name, str):
        raise TypeError("the VBA project name must be a string")
    if not name:
        raise ValueError("the VBA project name must not be empty")
    if any(ord(char) < 32 for char in name):
        raise ValueError("the VBA project name must not contain control characters")
    return name
