"""Pictures placed on an XLSX worksheet (PNG and JPEG).

The file bytes are embedded unchanged. Pixel dimensions are read from the
image headers, so Pillow is not needed.
"""

from __future__ import annotations

import hashlib
import os
import struct
from dataclasses import dataclass, field

from .refs import parse_cell_range

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
_JPEG_START_OF_FRAME = frozenset(
    {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}
)
_MAX_PIXELS = 16384


def _png_size(data: bytes) -> tuple[int, int]:
    if len(data) < 24 or data[12:16] != b"IHDR":
        raise ValueError("PNG file has no IHDR header")
    width, height = struct.unpack(">II", data[16:24])
    return width, height


def _jpeg_size(data: bytes) -> tuple[int, int]:
    index = 2
    while index + 9 <= len(data):
        if data[index] != 0xFF:
            break
        marker = data[index + 1]
        if marker == 0xFF:  # fill byte
            index += 1
            continue
        if marker in _JPEG_START_OF_FRAME:
            height, width = struct.unpack(">HH", data[index + 5:index + 9])
            return width, height
        (segment_length,) = struct.unpack(">H", data[index + 2:index + 4])
        index += 2 + segment_length
    raise ValueError("JPEG file has no frame header with its dimensions")


@dataclass(frozen=True, slots=True)
class Image:
    """A PNG or JPEG picture anchored to a cell.

    ``width`` and ``height`` are in pixels. When only one is given the other
    keeps the image's aspect ratio; when neither is given the image's own size
    is used. ``description`` is the alternative text.
    """

    path: str | os.PathLike
    anchor: str = "A1"
    width: int | None = None
    height: int | None = None
    description: str | None = None
    data: bytes = field(default=b"", repr=False, compare=False)
    extension: str = ""
    pixel_width: int = 0
    pixel_height: int = 0

    def __post_init__(self):
        if not isinstance(self.path, (str, os.PathLike)):
            raise TypeError("image path must be a string or path-like object")
        if not isinstance(self.anchor, str):
            raise TypeError("image anchor must be an A1 cell such as 'A1'")
        first_col, first_row, last_col, last_row = parse_cell_range(self.anchor)
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"image anchor must be a single cell, got {self.anchor!r}")
        with open(self.path, "rb") as handle:
            data = handle.read()
        extension, natural_width, natural_height = _identify(data, self.path)
        width = _pixels(self.width, "width")
        height = _pixels(self.height, "height")
        if width is None and height is None:
            width, height = natural_width, natural_height
        elif width is None:
            width = max(1, round(natural_width * height / natural_height))
        elif height is None:
            height = max(1, round(natural_height * width / natural_width))
        if self.description is not None and not isinstance(self.description, str):
            raise TypeError("image description must be a string or None")
        object.__setattr__(self, "data", data)
        object.__setattr__(self, "extension", extension)
        object.__setattr__(self, "pixel_width", width)
        object.__setattr__(self, "pixel_height", height)

    @property
    def digest(self) -> str:
        return hashlib.sha256(self.data).hexdigest()


def _pixels(value, name: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"image {name} must be an int number of pixels or None")
    if not 1 <= value <= _MAX_PIXELS:
        raise ValueError(f"image {name} must be between 1 and {_MAX_PIXELS} pixels")
    return value


def _identify(data: bytes, path) -> tuple[str, int, int]:
    """Return ``(extension, width, height)`` for PNG or JPEG bytes."""
    try:
        if data.startswith(_PNG_SIGNATURE):
            width, height = _png_size(data)
            extension = "png"
        elif data.startswith(b"\xff\xd8"):
            width, height = _jpeg_size(data)
            extension = "jpeg"
        else:
            raise ValueError(
                f"{os.fspath(path)!r} is not a PNG or JPEG file; other image formats are not supported"
            )
    except struct.error as exc:
        raise ValueError(f"{os.fspath(path)!r} has a truncated image header") from exc
    if width < 1 or height < 1:
        raise ValueError(f"{os.fspath(path)!r} reports an empty image size")
    return extension, width, height
