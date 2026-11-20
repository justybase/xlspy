"""The ZIP package and disk-backed shared-string index used by the updaters.

The readers and writers never share a package: this module owns the lazy
read-on-demand view, the single write-back transaction, and the SQLite-backed
string index that keeps a streamed replacement's memory bounded.
"""

from __future__ import annotations

import copy
import os
import shutil
import sqlite3
import struct
import tempfile
import weakref
import zipfile
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
from typing import IO, Any
from collections.abc import Iterable, Iterator, Mapping

from ._untrusted import SafeZipFile


STREAM_BUFFER_SIZE = 1024 * 1024


def copy_stream(source: IO[bytes], target: IO[bytes], *, length: int = STREAM_BUFFER_SIZE) -> None:
    """Copy a binary stream with a bounded buffer."""

    shutil.copyfileobj(source, target, length=length)


def copy_file_range(source_path: str | os.PathLike[str], target: IO[bytes], start: int, end: int) -> None:
    """Copy a byte range from a seekable file to an open binary stream."""

    with open(source_path, "rb") as source:
        source.seek(start)
        remaining = end - start
        while remaining > 0:
            chunk = source.read(min(STREAM_BUFFER_SIZE, remaining))
            if not chunk:
                raise ValueError("Unexpected end of temporary file")
            target.write(chunk)
            remaining -= len(chunk)


class DiskStringIndex:
    """Disk-backed shared-string lookup with an append-only pending file.

    The indexes of the strings looked up last are kept in memory as well (at
    most ``_CACHE_LIMIT``, so memory stays bounded): an index never changes
    once given, except by ``add_existing`` and a rollback, which drop them.
    """

    _CACHE_LIMIT = 65536

    def __init__(self, directory: str | os.PathLike[str], *, total_count: int = 0):
        directory = Path(directory)
        self._database = sqlite3.connect(directory / "shared-strings.sqlite3")
        self._database.execute(
            "CREATE TABLE IF NOT EXISTS strings (value TEXT PRIMARY KEY, idx INTEGER NOT NULL)"
        )
        self._database.commit()
        self._pending_path = directory / "shared-strings.pending"
        self._pending = self._pending_path.open("ab+")
        self.total_count = total_count
        self.unique_count = 0
        self.dirty = False
        self._uncommitted_inserts = 0
        self._transaction: dict[str, Any] | None = None
        self._cache: dict[str, int] = {}

    def seed(self, values: Iterable[str]) -> None:
        for index, value in enumerate(values):
            self.add_existing(value, index)
        self._database.commit()

    def add_existing(self, value: str, index: int) -> None:
        """Add one original SST value using its stable zero-based index."""

        # The in-memory reference implementation uses a dict comprehension,
        # so duplicate shared strings resolve to their last index.
        self._cache.pop(value, None)
        self._database.execute(
            "INSERT OR REPLACE INTO strings(value, idx) VALUES (?, ?)",
            (value, index),
        )
        self.unique_count = max(self.unique_count, index + 1)

    def get_or_add(self, value: str) -> int:
        cached = self._cache.get(value)
        if cached is not None:
            self.total_count += 1
            self.dirty = True
            return cached
        result = self._database.execute(
            "SELECT idx FROM strings WHERE value = ?", (value,)
        ).fetchone()
        if result is None:
            index = self.unique_count
            self._database.execute(
                "INSERT INTO strings(value, idx) VALUES (?, ?)", (value, index)
            )
            self._uncommitted_inserts += 1
            if self._uncommitted_inserts >= 1024:
                self._database.commit()
                self._uncommitted_inserts = 0
            encoded = value.encode("utf-8")
            # Buffered: every reader of the file flushes it first.
            self._pending.write(struct.pack("<Q", len(encoded)))
            self._pending.write(encoded)
            self.unique_count += 1
        else:
            index = int(result[0])
        if len(self._cache) >= self._CACHE_LIMIT:
            self._cache.clear()
        self._cache[value] = index
        self.total_count += 1
        self.dirty = True
        return index

    def pending_values(self) -> Iterator[str]:
        self._pending.flush()
        with self._pending_path.open("rb") as stream:
            while True:
                header = stream.read(8)
                if not header:
                    return
                if len(header) != 8:
                    raise ValueError("Malformed pending shared-string file")
                length = struct.unpack("<Q", header)[0]
                encoded = stream.read(length)
                if len(encoded) != length:
                    raise ValueError("Truncated pending shared-string file")
                yield encoded.decode("utf-8")

    def clear_pending(self) -> None:
        self._database.commit()
        self._uncommitted_inserts = 0
        self._pending.seek(0)
        self._pending.truncate()
        self._pending.flush()
        self.dirty = False

    def flush(self) -> None:
        """Commit pending lookup-index inserts without clearing new values."""

        if self._uncommitted_inserts:
            self._database.commit()
            self._uncommitted_inserts = 0

    def begin_transaction(self) -> None:
        """Start a disk-backed transaction for one streamed replacement."""

        if self._transaction is not None:
            raise RuntimeError("A shared-string transaction is already active")

        self.flush()
        self._pending.flush()
        descriptor, backup_path = tempfile.mkstemp(
            prefix="shared-strings-", suffix=".backup", dir=self._pending_path.parent
        )
        os.close(descriptor)
        backup = Path(backup_path)
        try:
            with self._pending_path.open("rb") as source, backup.open("wb") as target:
                copy_stream(source, target)
        except BaseException:
            try:
                backup.unlink()
            except FileNotFoundError:
                pass
            raise

        self._transaction = {
            "total_count": self.total_count,
            "unique_count": self.unique_count,
            "dirty": self.dirty,
            "uncommitted_inserts": self._uncommitted_inserts,
            "pending_backup": backup,
        }

    def commit_transaction(self) -> None:
        """Commit the current streamed replacement transaction."""

        transaction = self._transaction
        if transaction is None:
            return
        self._database.commit()
        backup = transaction["pending_backup"]
        try:
            backup.unlink()
        except FileNotFoundError:
            pass
        self._transaction = None

    def rollback_transaction(self) -> None:
        """Restore the index and pending values to the transaction snapshot."""

        transaction = self._transaction
        if transaction is None:
            return

        backup = transaction["pending_backup"]
        self._cache.clear()
        try:
            # New values always receive indexes at or above the old unique
            # count, so no in-memory list of inserted strings is needed.
            self._database.execute(
                "DELETE FROM strings WHERE idx >= ?",
                (transaction["unique_count"],),
            )
            self._database.commit()
            self._uncommitted_inserts = 0

            self._pending.seek(0)
            self._pending.truncate()
            with backup.open("rb") as source:
                copy_stream(source, self._pending)
            self._pending.flush()

            self.total_count = transaction["total_count"]
            self.unique_count = transaction["unique_count"]
            self.dirty = transaction["dirty"]
        finally:
            try:
                backup.unlink()
            except FileNotFoundError:
                pass
            self._transaction = None

    def close(self) -> None:
        if self._transaction is not None:
            self.rollback_transaction()
        self._pending.close()
        self._database.close()

    def __del__(self) -> None:
        # Close the pending file (and the sqlite handle) explicitly: relying on
        # garbage collection alone leaks the file handle and makes Python 3.12+
        # emit a ResourceWarning.
        try:
            self.close()
        except Exception:
            pass


class ZipPackage:
    """Lazy ZIP package used by both the regular and streaming updaters.

    The original updater API still exposes byte-oriented ``get``/``set`` and
    ``to_bytes`` methods.  Entries are now read on demand, while the optional
    file-backed replacements are used by the streaming updater so a workbook
    is not materialised in memory.
    """

    def __init__(self, path: str | os.PathLike[str]):
        self.path = Path(path)
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        if not self.path.is_file():
            raise ValueError(f"Expected a file: {self.path}")

        self._entries: list[zipfile.ZipInfo] = []
        self._indices: dict[str, int] = {}
        with SafeZipFile(self.path, "r") as archive:
            for info in archive.infolist():
                copied_info = copy.copy(info)
                self._entries.append(copied_info)
                self._indices[info.filename] = len(self._entries) - 1
        self._overrides: dict[int, bytes | Path] = {}
        self._temporary_directory = tempfile.mkdtemp(prefix="xlspy-updater-")
        # A weakref.finalize rather than TemporaryDirectory's own finalizer: in a
        # reference cycle CPython clears weak references before __del__ runs, so
        # TemporaryDirectory warned ("Implicitly cleaning up") when collected.
        self._temporary_directory_finalizer = weakref.finalize(
            self, shutil.rmtree, self._temporary_directory, ignore_errors=True
        )
        self._package_transaction: tuple[
            list[zipfile.ZipInfo],
            dict[str, int],
            dict[int, bytes | Path],
        ] | None = None
        self._superseded_paths: set[Path] = set()

    def close(self) -> None:
        """Remove the temporary directory holding staged package parts."""

        if self._temporary_directory_finalizer.detach() is not None:
            shutil.rmtree(self._temporary_directory)

    @property
    def in_transaction(self) -> bool:
        """Whether a transaction is open, so writes are staged directly."""
        return self._package_transaction is not None

    def has(self, name: str) -> bool:
        return name in self._indices

    def names(self) -> list[str]:
        """Return current ZIP member names in archive order."""

        return [info.filename for info in self._entries]

    def _index_for(self, name: str) -> int:
        try:
            return self._indices[name]
        except KeyError as exc:
            raise KeyError(f"ZIP member not found: {name}") from exc

    def get(self, name: str) -> bytes:
        index = self._index_for(name)
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            return override
        if isinstance(override, Path):
            return override.read_bytes()
        with SafeZipFile(self.path, "r") as archive:
            return archive.read(self._entries[index])

    @contextmanager
    def open_entry(self, name: str) -> Iterator[IO[bytes]]:
        """Open one entry for bounded-memory reading."""
        index = self._index_for(name)
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            stream: IO[bytes] = BytesIO(override)
            try:
                yield stream
            finally:
                stream.close()
            return
        if isinstance(override, Path):
            with override.open("rb") as file_stream:
                yield file_stream
            return

        with SafeZipFile(self.path, "r") as archive:
            with archive.open(self._entries[index], "r") as stream:
                yield stream

    def temporary_path(self, *, suffix: str = "") -> Path:
        """Create a named temporary file owned by this package."""

        descriptor, path = tempfile.mkstemp(
            prefix="part-", suffix=suffix, dir=self._temporary_directory
        )
        os.close(descriptor)
        return Path(path)

    @property
    def temporary_directory(self) -> Path:
        """Return the private directory used for staged package parts."""

        return Path(self._temporary_directory)

    def begin_transaction(self) -> None:
        """Start a transaction for staged package replacements."""

        if self._package_transaction is not None:
            raise RuntimeError("A ZIP package transaction is already active")
        self._package_transaction = (
            self._entries.copy(),
            self._indices.copy(),
            self._overrides.copy(),
        )
        self._superseded_paths = set()

    def _is_owned_staged_path(self, path: Path) -> bool:
        try:
            path.resolve().relative_to(self.temporary_directory.resolve())
        except (OSError, ValueError):
            return False
        return True

    def _delete_staged_path(self, path: Path) -> None:
        if not self._is_owned_staged_path(path):
            return
        try:
            path.unlink()
        except FileNotFoundError:
            pass

    def _set_override(self, index: int, replacement: bytes | Path) -> None:
        previous = self._overrides.get(index)
        if isinstance(previous, Path) and previous != replacement:
            if self._package_transaction is None:
                self._delete_staged_path(previous)
            else:
                self._superseded_paths.add(previous)
        self._overrides[index] = replacement

    def commit_transaction(self) -> None:
        """Commit staged replacements and remove superseded temporary files."""

        if self._package_transaction is None:
            return
        current_paths = {
            override
            for override in self._overrides.values()
            if isinstance(override, Path)
        }
        for path in self._superseded_paths:
            if path not in current_paths:
                self._delete_staged_path(path)
        self._package_transaction = None
        self._superseded_paths = set()

    def rollback_transaction(self) -> None:
        """Discard staged changes made since the transaction began."""

        transaction = self._package_transaction
        if transaction is None:
            return
        _, _, original_overrides = transaction
        original_paths = {
            override
            for override in original_overrides.values()
            if isinstance(override, Path)
        }
        for override in self._overrides.values():
            if isinstance(override, Path) and override not in original_paths:
                self._delete_staged_path(override)
        self._entries, self._indices, self._overrides = transaction
        self._package_transaction = None
        self._superseded_paths = set()

    def apply_parts(self, parts: Mapping[str, str | bytes], removed: Iterable[str] = ()) -> None:
        """Write ``parts`` (str or bytes) and remove ``removed`` as one transaction.

        The package owns this write-back: an updater or the object model that
        has pending parts hands them here. A failure rolls everything back, so
        the package is never left half-updated.
        """

        self.begin_transaction()
        try:
            for path, data in parts.items():
                self.set(path, data.encode("utf-8") if isinstance(data, str) else data)
            for path in removed:
                self.remove(path)
        except BaseException:
            self.rollback_transaction()
            raise
        else:
            self.commit_transaction()

    def set(self, name: str, data: bytes) -> None:
        if name in self._indices:
            index = self._indices[name]
            self._set_override(index, bytes(data))
            return

        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        self._indices[name] = len(self._entries)
        self._entries.append(info)
        self._set_override(len(self._entries) - 1, bytes(data))

    def remove(self, name: str) -> None:
        """Remove a member, including staged replacements; supports rollback."""
        if name not in self._indices:
            return
        removed_index = self._indices[name]
        previous = self._overrides.get(removed_index)
        if isinstance(previous, Path):
            if self._package_transaction is None:
                self._delete_staged_path(previous)
            else:
                self._superseded_paths.add(previous)
        self._entries = [info for index, info in enumerate(self._entries) if index != removed_index]
        self._indices = {info.filename: index for index, info in enumerate(self._entries)}
        self._overrides = {
            index - (index > removed_index): value
            for index, value in self._overrides.items() if index != removed_index
        }

    def set_file(self, name: str, path: str | os.PathLike[str]) -> None:
        """Use a file as the replacement payload for one ZIP member."""

        replacement = Path(path)
        if not replacement.is_file():
            raise ValueError(f"Expected a replacement file: {replacement}")
        if name in self._indices:
            self._set_override(self._indices[name], replacement)
            return

        info = zipfile.ZipInfo(name)
        info.compress_type = zipfile.ZIP_DEFLATED
        self._indices[name] = len(self._entries)
        self._entries.append(info)
        self._set_override(len(self._entries) - 1, replacement)

    @contextmanager
    def _replacement_stream(self, index: int) -> Iterator[IO[bytes]]:
        override = self._overrides.get(index)
        if isinstance(override, bytes):
            with BytesIO(override) as stream:
                yield stream
            return
        if isinstance(override, Path):
            with override.open("rb") as stream:
                yield stream
            return
        with SafeZipFile(self.path, "r") as archive:
            with archive.open(self._entries[index], "r") as stream:
                yield stream

    def to_bytes(self) -> bytes:
        result = BytesIO()
        with SafeZipFile(result, "w") as archive:
            for index, original_info in enumerate(self._entries):
                with self._replacement_stream(index) as stream:
                    with archive.open(copy.copy(original_info), "w") as target:
                        shutil.copyfileobj(stream, target, length=1024 * 1024)
        return result.getvalue()

    def save(self, output_path: str | os.PathLike[str] | None = None) -> None:
        """Write the package, using a same-directory temporary file and replace."""

        target = Path(output_path) if output_path is not None else self.path
        target = target.resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary_path: str | None = None
        descriptor, temporary_path = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp", dir=target.parent
        )
        os.close(descriptor)
        try:
            with SafeZipFile(temporary_path, "w") as archive:
                for index, original_info in enumerate(self._entries):
                    with self._replacement_stream(index) as stream:
                        with archive.open(copy.copy(original_info), "w") as target_stream:
                            shutil.copyfileobj(stream, target_stream, length=1024 * 1024)
                archive.comment = b""
            with open(temporary_path, "rb+") as stream:
                os.fsync(stream.fileno())
            os.replace(temporary_path, target)
            temporary_path = None
            if target == self.path.resolve():
                self._reload_entries(target)
        finally:
            if temporary_path is not None:
                try:
                    os.unlink(temporary_path)
                except FileNotFoundError:
                    pass

    def _reload_entries(self, source_path: str | os.PathLike[str]) -> None:
        """Refresh ZIP metadata after an archive has been replaced in place."""

        entries: list[zipfile.ZipInfo] = []
        indices: dict[str, int] = {}
        with SafeZipFile(source_path, "r") as archive:
            for info in archive.infolist():
                entries.append(copy.copy(info))
                indices[info.filename] = len(entries) - 1
        self._entries = entries
        self._indices = indices
