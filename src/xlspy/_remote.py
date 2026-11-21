"""Workbook paths that live on a remote filesystem, through ``fsspec``.

A string such as ``s3://bucket/report.xlsx`` or ``memory://report.xlsx`` is a
URI, not a local path. ``fsspec`` maps its scheme to a filesystem; the package
is downloaded once on open and uploaded once on save. Local paths keep their
staged, atomic replacement.
"""

from __future__ import annotations

import re
import shutil
import tempfile
from typing import IO, Any, TypeGuard
from urllib.parse import urlsplit

_COPY_CHUNK = 1 << 20
_SCHEME = re.compile(r"^([A-Za-z][A-Za-z0-9+.\-]*)://")
# "C:\..." and "C:/..." are local paths, not "C" URIs.
_WINDOWS_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_STAGING_BYTES = 8 << 20


def is_remote_uri(value: Any) -> TypeGuard[str]:
    """True for ``scheme://...`` paths that ``fsspec`` should open."""
    if not isinstance(value, str) or _WINDOWS_DRIVE.match(value):
        return False
    return _SCHEME.match(value) is not None


def uri_extension(uri: str, default: str = "xlsx") -> str:
    """The lower-case extension of a URI's path, without the dot."""
    suffix = urlsplit(uri).path.rsplit("/", 1)[-1]
    return suffix.rsplit(".", 1)[-1].lower() if "." in suffix else default


def _fsspec() -> Any:
    try:
        import fsspec
    except ImportError as exc:  # pragma: no cover - depends on the environment
        raise ImportError(
            'reading or writing a workbook URI needs fsspec; install it with "pip install xlspy[remote]"'
        ) from exc
    return fsspec


def open_remote(uri: str) -> IO[bytes]:
    """Download ``uri`` into a seekable binary stream the caller must close.

    The stream stays in memory up to a few megabytes and spills to a temporary
    file after that, so a large workbook does not have to fit twice in memory.
    """
    stream = tempfile.SpooledTemporaryFile(max_size=_STAGING_BYTES, mode="w+b")
    try:
        with _fsspec().open(uri, "rb") as data:
            shutil.copyfileobj(data, stream, _COPY_CHUNK)
        stream.seek(0)
    except BaseException:
        stream.close()
        raise
    return stream


def download_remote(uri: str, destination: str) -> None:
    """Copy the remote package at ``uri`` to the local file ``destination``."""
    with _fsspec().open(uri, "rb") as source, open(destination, "wb") as target:
        shutil.copyfileobj(source, target, _COPY_CHUNK)


def upload_stream(source: IO[bytes], uri: str) -> None:
    """Copy an open binary stream to the remote target ``uri``."""
    with _fsspec().open(uri, "wb") as target:
        shutil.copyfileobj(source, target, _COPY_CHUNK)


def upload_remote(source: str, uri: str) -> None:
    """Copy the local file ``source`` to the remote target ``uri``."""
    with open(source, "rb") as data:
        upload_stream(data, uri)
