"""Staged output: a workbook reaches its target only after a successful save.

A writer builds its ZIP archive in a staging file and publishes it only when
the workbook was saved and the archive was closed without error:

* A path target is staged in a temporary file in the same directory and moved
  into place with ``os.replace``. A failed write creates nothing and leaves an
  existing file at that path unchanged.
* A file-object target is staged in a ``SpooledTemporaryFile`` (in memory up
  to ``spool_max_bytes``, then on disk) and copied into the stream on success.
  A failed write leaves the stream unchanged.

The copy into a caller's stream is the one step that cannot be rolled back:
if the stream itself fails part-way through that copy, the bytes it already
accepted stay there.
"""

import errno
import os
import secrets
import shutil
import tempfile
import zipfile

DEFAULT_STAGING_BYTES = 8 << 20
_COPY_CHUNK = 1 << 20


class StagedOutput:
    """The staging file for one workbook and its eventual publication."""

    def __init__(self, target, spool_max_bytes: int | None = None):
        if spool_max_bytes is None:
            spool_max_bytes = DEFAULT_STAGING_BYTES
        self._target = target
        self._temp_path = None
        self._final_path = None
        self._finished = False
        if isinstance(target, (str, bytes, os.PathLike)):
            self._final_path = _final_path(target)
            self._temp_path, self.stream = _create_sibling(self._final_path)
        elif callable(getattr(target, "write", None)):
            self.stream = tempfile.SpooledTemporaryFile(max_size=spool_max_bytes, mode="w+b")
        else:
            raise TypeError("file must be a path or a writable binary file object")

    def publish(self) -> None:
        """Move the finished archive to the target. Call once, after the ZIP is closed."""
        if self._finished:
            raise RuntimeError("the output was already published or discarded")
        self._finished = True
        try:
            if self._final_path is not None:
                self.stream.close()
                try:
                    shutil.copymode(self._final_path, self._temp_path)
                except OSError:
                    pass  # no existing file, or its mode cannot be read
                os.replace(self._temp_path, self._final_path)
                self._temp_path = None
            else:
                self.stream.seek(0)
                shutil.copyfileobj(self.stream, self._target, _COPY_CHUNK)
                flush = getattr(self._target, "flush", None)
                if flush is not None:
                    flush()
        finally:
            self._cleanup()

    def discard(self) -> None:
        """Drop the staged archive; the target is left as it was."""
        if self._finished:
            return
        self._finished = True
        self._cleanup()

    def _cleanup(self) -> None:
        try:
            self.stream.close()
        except OSError:
            pass
        if self._temp_path is not None:
            try:
                os.remove(self._temp_path)
            except OSError:
                pass
            self._temp_path = None



class StagedWriter:
    """The open/save/close lifecycle shared by ``XlsxWriter`` and ``XlsbWriter``.

    A subclass sets ``_file``, ``_compressionLevel``, ``_zf = None``,
    ``_output = None``, ``_aborted = False`` and ``_saved = False`` in its
    ``__init__``, writes its parts in ``_write_workbook_parts``, and sets
    ``_aborted`` when a sheet write fails part-way.
    """

    def _staging_bytes(self):
        """In-memory size of the staging spool for file-object targets."""
        return None

    def __enter__(self):
        if self._zf is None:
            output = StagedOutput(self._file, self._staging_bytes())
            try:
                self._zf = zipfile.ZipFile(output.stream, 'w', zipfile.ZIP_DEFLATED,
                                           compresslevel=self._compressionLevel)
            except BaseException:
                output.discard()
                raise
            self._output = output
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        publish = exc_type is None
        try:
            if publish and self._zf is not None and not self._saved:
                self.save()
        except BaseException:
            publish = False
            raise
        finally:
            self._finish(publish)
        return False

    def _require_writable(self):
        if self._zf is None:
            raise RuntimeError("the writer is not open; use it as a context manager")
        if self._aborted:
            raise RuntimeError(
                "a previous write failed; the workbook is incomplete and was not saved"
            )
        if self._saved:
            raise RuntimeError("the workbook was already saved")

    def save(self):
        """Write the workbook parts. The file is published when the writer closes."""
        self._require_writable()
        try:
            self._write_workbook_parts()
        except BaseException:
            self._aborted = True
            raise
        self._saved = True

    def close(self):
        """Close the writer, publishing the workbook if ``save()`` succeeded."""
        self._finish(publish=True)

    def __del__(self):
        # A writer abandoned without close() publishes nothing and leaves no
        # staging file. The archive is closed before its staging file goes.
        try:
            if getattr(self, "_output", None) is not None:
                self._finish(publish=False)
        except Exception:
            pass  # interpreter shutdown: module globals may already be gone

    def _finish(self, publish: bool):
        zf, output = self._zf, self._output
        self._zf = self._output = None
        if output is None:
            return
        if publish and self._saved and not self._aborted:
            try:
                zf.close()
            except BaseException:
                output.discard()
                raise
            output.publish()
        else:
            try:
                zf.close()
            except Exception:
                pass  # the archive is being dropped; a failure here changes nothing
            output.discard()


def _final_path(target) -> str:
    path = os.fsdecode(os.fspath(target))
    # Replace the file a symlink points to, not the link itself.
    if os.path.islink(path):
        path = os.path.realpath(path)
    if os.path.isdir(path):
        raise IsADirectoryError(errno.EISDIR, os.strerror(errno.EISDIR), path)
    # Fail now, as opening the file would have, rather than after the whole write.
    if os.path.exists(path) and not os.access(path, os.W_OK):
        raise PermissionError(errno.EACCES, os.strerror(errno.EACCES), path)
    return path


def _create_sibling(path: str):
    """Create a new temporary file next to ``path``; return its name and handle."""
    directory, name = os.path.split(os.path.abspath(path))
    for _ in range(100):
        temp_path = os.path.join(directory, f".{name}.{secrets.token_hex(4)}.tmp")
        try:
            return temp_path, open(temp_path, "x+b")
        except FileExistsError:
            continue
        except OSError as exc:
            # Report the target the caller named, not the temporary name.
            raise type(exc)(exc.errno, exc.strerror, path) from None
    raise FileExistsError(errno.EEXIST, "could not create a temporary file", path)
