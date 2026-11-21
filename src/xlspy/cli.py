"""The ``xlspy`` command line: inspect, validate and convert workbooks."""

from __future__ import annotations

import argparse
import json
import os
import sys
import zipfile
from typing import Any

from .exceptions import XlspyError
from .reader import ExcelReader
from .workbook import load_workbook

_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb")


def _object_counts(sheet: Any) -> dict[str, int]:
    return {
        "tables": len(getattr(sheet, "tables", [])),
        "charts": len(getattr(sheet, "charts", [])),
        "images": len(getattr(sheet, "images", [])),
        "comments": len(getattr(sheet, "_comments", {})),
        "threaded_comments": len(getattr(sheet, "_threaded_comments", {})),
        "data_validations": len(getattr(sheet, "data_validations", [])),
        "conditional_formats": len(getattr(sheet, "conditional_formatting", [])),
        "merged_ranges": len(getattr(getattr(sheet, "merged_cells", None), "ranges", [])),
        "ignored_errors": len(getattr(sheet, "ignored_errors", [])),
    }


def _inspect(path: str) -> dict[str, Any]:
    workbook = load_workbook(path)
    try:
        sheets = []
        for sheet in workbook._sheets:
            sheets.append(
                {
                    "name": sheet.title,
                    "state": sheet.sheet_state,
                    "max_row": sheet.max_row,
                    "max_column": sheet.max_column,
                    **_object_counts(sheet),
                }
            )
        return {
            "path": os.fspath(path),
            "format": getattr(workbook, "_format", None) or _format_from_name(path),
            "sheets": sheets,
            "defined_names": len(workbook.defined_names),
        }
    finally:
        workbook.close()


def _format_from_name(path: str) -> str:
    return os.path.splitext(os.fspath(path))[1].lower().lstrip(".") or "unknown"


def _print_inspect(info: dict[str, Any]) -> None:
    print(f"Path: {info['path']}")
    print(f"Format: {info['format']}")
    print(f"Sheets: {len(info['sheets'])}")
    for sheet in info["sheets"]:
        objects = ", ".join(
            f"{name}={sheet[name]}"
            for name in (
                "tables",
                "charts",
                "images",
                "comments",
                "threaded_comments",
                "data_validations",
                "conditional_formats",
                "merged_ranges",
            )
            if sheet[name]
        )
        size = f"{sheet['max_row']}x{sheet['max_column']}"
        state = "" if sheet["state"] == "visible" else f", {sheet['state']}"
        print(f"  {sheet['name']} ({size}{state})" + (f": {objects}" if objects else ""))
    if info["defined_names"]:
        print(f"Defined names: {info['defined_names']}")


def _command_inspect(args: argparse.Namespace) -> int:
    info = _inspect(args.path)
    if args.json:
        print(json.dumps(info, indent=2))
    else:
        _print_inspect(info)
    return 0


def _command_sheets(args: argparse.Namespace) -> int:
    with ExcelReader(args.path) as reader:
        names = reader.get_sheet_names()
    if args.json:
        print(json.dumps(names, indent=2))
    else:
        print("\n".join(names))
    return 0


def _command_validate(args: argparse.Namespace) -> int:
    with ExcelReader(args.path) as reader:
        names = reader.get_sheet_names()
        if not args.quick:
            for name in names:
                for _row in reader.get_rows(name):
                    pass
    workbook = load_workbook(args.path)
    try:
        sheet_count = len(workbook._sheets)
    finally:
        workbook.close()
    print(f"OK: {args.path} ({sheet_count} sheets, {len(names)} readable)")
    return 0


def _command_convert(args: argparse.Namespace) -> int:
    destination = os.fspath(args.destination)
    extension = os.path.splitext(destination)[1].lower()
    if extension not in _EXTENSIONS:
        raise ValueError(f"unsupported destination extension {extension!r}; use one of {', '.join(_EXTENSIONS)}")
    if os.path.abspath(os.fspath(args.source)) == os.path.abspath(destination):
        raise ValueError("source and destination are the same file")
    workbook = load_workbook(args.source)
    try:
        workbook.save(destination)
    finally:
        workbook.close()
    print(f"Converted {args.source} -> {destination}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlspy",
        description="Inspect, validate and convert XLSX/XLSM/XLSB workbooks.",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    inspect = commands.add_parser("inspect", help="show sheets, sizes and object counts")
    inspect.add_argument("path")
    inspect.add_argument("--json", action="store_true", help="print machine-readable JSON")
    inspect.set_defaults(handler=_command_inspect)

    sheets = commands.add_parser("sheets", help="list worksheet names")
    sheets.add_argument("path")
    sheets.add_argument("--json", action="store_true", help="print machine-readable JSON")
    sheets.set_defaults(handler=_command_sheets)

    validate = commands.add_parser("validate", help="parse every sheet and report errors")
    validate.add_argument("path")
    validate.add_argument("--quick", action="store_true", help="open the package without reading rows")
    validate.set_defaults(handler=_command_validate)

    convert = commands.add_parser("convert", help="rewrite a workbook as another format")
    convert.add_argument("source")
    convert.add_argument("destination")
    convert.set_defaults(handler=_command_convert)
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the command line; return the process exit code."""
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.handler(args))
    except (OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, XlspyError) as error:
        print(f"xlspy: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
