"""Deflate ZIP members on a worker thread while the writer encodes more data.

``zlib`` releases the GIL while it compresses, so a single worker overlaps
compression with row encoding (C or Python) in the calling thread.

Two ways in:

- ``ThreadedDeflate`` replaces the compressor of an open member: the same
  ``compressobj`` gets the same bytes in the same order, in blocks of
  ``_BLOCK`` bytes. (At levels 3-9 the member is byte-identical to a
  sequential write; at levels 0-1 zlib's block boundaries follow the input
  chunks, so the compressed bytes can differ. Output stays deterministic.)
- ``DeflateSegment`` compresses data that must be stored *after* a header that
  is only known later (an XLSX sheet: ``<dimension>`` depends on the rows).  The
  segment ends with a sync flush, so ``replay_member`` can write
  header + segment + trailer as one valid deflate stream, as pigz does.

Both reach into ``zipfile._ZipWriteFile._compressor``; ``AVAILABLE`` is False
when that attribute is not there, and callers then write sequentially.
"""

from __future__ import annotations

from typing import BinaryIO
import shutil
import zipfile
import zlib
from collections import deque
from concurrent.futures import Future, ThreadPoolExecutor

AVAILABLE = "_compressor" in getattr(
    getattr(zipfile, "_ZipWriteFile", None), "__init__", lambda: None
).__code__.co_names

# Input is handed to the worker in blocks of this size: large enough that the
# per-task overhead is negligible, small enough to keep the worker busy.
_BLOCK = 256 * 1024
# Blocks queued before the caller waits for the worker (bounds memory).
_MAX_PENDING = 8


def new_deflater(level: int | None) -> zlib._Compress:
    """The compressor ``zipfile`` would use for a deflated member."""
    return zlib.compressobj(
        zlib.Z_DEFAULT_COMPRESSION if level is None else level, zlib.DEFLATED, -15
    )


class _Pipeline:
    """Feed one ``compressobj`` from the caller and run it on ``executor``."""

    def __init__(self, executor: ThreadPoolExecutor, level: int | None):
        self._executor = executor
        self._deflater = new_deflater(level)
        self._buffer: list[bytes] = []
        self._buffered = 0
        self._pending: deque[Future] = deque()

    def _feed(self, data) -> list[bytes]:
        """Queue ``data``; return the compressed output that is ready."""
        if data:
            self._buffer.append(bytes(data))
            self._buffered += len(data)
            if self._buffered >= _BLOCK:
                self._submit()
        out = []
        pending = self._pending
        while pending and (pending[0].done() or len(pending) > _MAX_PENDING):
            out.append(pending.popleft().result())
        return out

    def _submit(self) -> None:
        block = b"".join(self._buffer)
        self._buffer.clear()
        self._buffered = 0
        self._pending.append(self._executor.submit(self._deflater.compress, block))

    def _finish(self, mode: int) -> list[bytes]:
        if self._buffer:
            self._submit()
        out = [future.result() for future in self._pending]
        self._pending.clear()
        out.append(self._deflater.flush(mode))
        return out

    def cancel(self) -> None:
        for future in self._pending:
            future.cancel()
        self._pending.clear()
        self._buffer.clear()


class ThreadedDeflate(_Pipeline):
    """A drop-in ``compressobj`` for an open ZIP member."""

    def compress(self, data) -> bytes:
        return b"".join(self._feed(data))

    def flush(self, mode: int = zlib.Z_FINISH) -> bytes:
        return b"".join(self._finish(mode))


class DeflateSegment(_Pipeline):
    """Compress a stream into ``out`` (a binary file) as a mid-stream segment."""

    def __init__(self, executor, level, out):
        super().__init__(executor, level)
        self._out = out

    def write(self, data) -> int:
        for chunk in self._feed(data):
            self._out.write(chunk)
        return len(data)

    def finish(self) -> None:
        for chunk in self._finish(zlib.Z_SYNC_FLUSH):
            self._out.write(chunk)


class _Replay:
    """Compressor that returns precompressed bytes instead of compressing."""

    def __init__(self, head: bytes, body: BinaryIO, tail: bytes):
        self._head = head
        self._body = body
        self._tail = tail

    def compress(self, data) -> bytes:
        out = self._head + self._body.read(1 << 20)
        self._head = b""
        return out

    def flush(self) -> bytes:
        out = self._head + self._body.read() + self._tail
        self._head = self._tail = b""
        return out


def _segment(level, data: bytes, mode: int) -> bytes:
    deflater = new_deflater(level)
    return deflater.compress(data) + deflater.flush(mode)


def replay_member(zf, name: str, head: bytes, raw_body, compressed_body, tail: bytes) -> None:
    """Write ``head + raw_body + tail`` as member ``name`` of ``zf``.

    ``compressed_body`` holds ``raw_body`` compressed by a finished
    ``DeflateSegment`` at the archive's level; both files are read from their
    current position. The raw bytes still pass through the member for its
    CRC and size.
    """
    level = zf.compresslevel
    with zf.open(name, "w") as member:
        member._compressor = _Replay(
            _segment(level, head, zlib.Z_SYNC_FLUSH), compressed_body,
            _segment(level, tail, zlib.Z_FINISH),
        )
        member.write(head)
        shutil.copyfileobj(raw_body, member, 1 << 20)
        member.write(tail)
