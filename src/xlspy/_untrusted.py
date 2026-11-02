"""Resource limits for untrusted package input.

Workbooks come from outside the process, so a crafted file can claim a huge
decompressed size or a row far wider than Excel allows. This module holds the
limits and the ZIP wrapper that enforces the size one. It imports only the
exception module, which has no imports of its own, so any reader can depend on it.
"""

from __future__ import annotations
from typing import IO, Any, Literal

import zipfile

from .exceptions import InvalidFileError

#: Widest row Excel can hold (columns A..XFD). Wider spans are refused.
MAX_COLUMNS = 16384

#: Largest uncompressed size accepted for one package part (2 GiB). A part is
#: read into memory whole, so this bounds the memory one member can take.
MAX_PART_BYTES = 2 << 30

#: Largest sum of the uncompressed sizes of all parts in one package (4 GiB).
#: Each part is capped by MAX_PART_BYTES, but a package with many large parts
#: could still exhaust memory once they are read one after another.
MAX_PACKAGE_BYTES = 4 << 30


def check_column_span(first_column: int, last_column: int) -> None:
    """Raise InvalidFileError when columns ``first_column``..``last_column`` span more than Excel's width."""
    if last_column - first_column >= MAX_COLUMNS:
        raise InvalidFileError(
            f"row spans columns {first_column}..{last_column}, wider than the "
            f"{MAX_COLUMNS} columns a worksheet can hold"
        )


class SafeZipFile(zipfile.ZipFile):
    """ZipFile that refuses to read a member whose declared size is over MAX_PART_BYTES.

    ``ZipFile.read`` and ``ZipFile.open`` both go through ``open``, so every read
    path is covered. The reader stops at the declared size, so a member that lies
    about its size is cut off there and fails its CRC check.
    """

    def __init__(self, file: Any, mode: Literal["r", "w", "x", "a"] = "r", *args: Any,
                 **kwargs: Any) -> None:
        # Report a non-ZIP input as InvalidFileError, which is also a BadZipFile.
        try:
            super().__init__(file, mode, *args, **kwargs)
        except zipfile.BadZipFile as exc:
            raise InvalidFileError(f"not a valid Excel package: {exc}") from exc
        if mode == "r":
            total = sum(info.file_size for info in self.infolist())
            if total > MAX_PACKAGE_BYTES:
                self.close()
                raise ValueError(
                    f"package declares {total} bytes uncompressed, over the "
                    f"{MAX_PACKAGE_BYTES}-byte limit"
                )

    def open(self, name: str | zipfile.ZipInfo, mode: Literal["r", "w"] = "r",
             pwd: bytes | None = None, *, force_zip64: bool = False) -> IO[bytes]:
        if mode == "r":
            info = name if isinstance(name, zipfile.ZipInfo) else self.getinfo(name)
            if info.file_size > MAX_PART_BYTES:
                raise ValueError(
                    f"package part {info.filename!r} is {info.file_size} bytes uncompressed, "
                    f"over the {MAX_PART_BYTES}-byte limit"
                )
        return super().open(name, mode, pwd, force_zip64=force_zip64)
