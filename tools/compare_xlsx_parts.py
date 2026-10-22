"""Compare the zip parts of XLSX/XLSB files in two directories, byte by byte.

Usage:
    python tools/compare_xlsx_parts.py BEFORE_DIR AFTER_DIR

Files are matched by relative path. Parts whose bytes differ are listed;
``docProps/core.xml`` is ignored because it carries the creation time.
The exit status is 1 when any file or part differs.
"""

import os
import sys
import zipfile

IGNORED = {"docProps/core.xml"}


def _files(root):
    for directory, _, names in os.walk(root):
        for name in names:
            if name.endswith((".xlsx", ".xlsb")):
                path = os.path.join(directory, name)
                yield os.path.relpath(path, root), path


def compare(before_root, after_root):
    before = dict(_files(before_root))
    after = dict(_files(after_root))
    differences = 0
    for rel in sorted(set(before) | set(after)):
        if rel not in before or rel not in after:
            print(f"{rel}: only in {'before' if rel in before else 'after'}")
            differences += 1
            continue
        with zipfile.ZipFile(before[rel]) as old, zipfile.ZipFile(after[rel]) as new:
            old_names, new_names = set(old.namelist()), set(new.namelist())
            changed = [name for name in sorted(old_names ^ new_names)]
            changed += [
                name for name in sorted(old_names & new_names)
                if name not in IGNORED and old.read(name) != new.read(name)
            ]
        if changed:
            differences += 1
            print(f"{rel}: {', '.join(changed)}")
        else:
            print(f"{rel}: identical")
    return differences


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.exit(1 if compare(sys.argv[1], sys.argv[2]) else 0)
