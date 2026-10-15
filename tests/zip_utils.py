"""Helpers for tests that edit members of an existing zip-based workbook.

Appending a member with ``ZipFile(path, "a")`` that already exists creates a
duplicate entry (and a ``UserWarning``), which is not what these tests intend.
The helper below replaces existing members in place and only appends genuinely
new ones, so every produced archive keeps unique entry names.
"""

import zipfile
from pathlib import Path


def rewrite_workbook_members(path, members):
    """Set each ``members[name] = payload`` in the archive at ``path``.

    Existing members are replaced; new members are appended.  The archive is
    rewritten through a temporary file so it never contains duplicate names.
    """
    path = Path(path)
    temporary = path.with_name(path.name + ".tmp")
    remaining = dict(members)

    with zipfile.ZipFile(path) as source:
        entries = [(info, source.read(info)) for info in source.infolist()]

    with zipfile.ZipFile(temporary, "w") as target:
        for info, payload in entries:
            payload = remaining.pop(info.filename, payload)
            target.writestr(info, payload)
        for name, payload in remaining.items():
            target.writestr(name, payload)

    temporary.replace(path)
