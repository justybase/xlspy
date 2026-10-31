"""The exception hierarchy: what xlspy raises and which built-in types still match."""

import struct
import zipfile

import pytest

import xlspy
from xlspy import (
    ExcelReader,
    InvalidFileError,
    UnsupportedFeatureError,
    XlspyError,
    load_workbook,
)
from xlspy import _accel
from xlspy.exceptions import InvalidFileError as ModuleInvalidFileError
from xlspy.utils.exceptions import InvalidFileException
from xlspy.workbook import _xlsb_loaded


def test_hierarchy_keeps_the_builtin_bases():
    assert issubclass(InvalidFileError, XlspyError)
    assert issubclass(InvalidFileError, ValueError)
    assert issubclass(InvalidFileError, zipfile.BadZipFile)
    assert issubclass(UnsupportedFeatureError, XlspyError)
    assert issubclass(UnsupportedFeatureError, NotImplementedError)
    assert InvalidFileError is ModuleInvalidFileError


def test_public_names_are_exported():
    for name in ("XlspyError", "InvalidFileError", "UnsupportedFeatureError"):
        assert name in xlspy.__all__
        assert getattr(xlspy, name) is not None


def test_openpyxl_invalid_file_exception_joins_the_hierarchy():
    assert issubclass(InvalidFileException, InvalidFileError)
    assert issubclass(InvalidFileException, XlspyError)


@pytest.mark.parametrize("opener", [
    lambda path: ExcelReader(str(path)).__enter__(),
    lambda path: load_workbook(str(path)),
], ids=["ExcelReader", "load_workbook"])
def test_a_non_zip_file_raises_invalid_file_error(tmp_path, opener):
    path = tmp_path / "not-a-workbook.xlsx"
    path.write_bytes(b"this is plain text, not a ZIP package")

    with pytest.raises(InvalidFileError, match="not a valid Excel package") as info:
        opener(path)
    # Code written against the old built-in types keeps catching this error.
    assert isinstance(info.value, zipfile.BadZipFile)
    assert isinstance(info.value, ValueError)
    assert isinstance(info.value, XlspyError)


def test_unsupported_feature_helper_is_both_xlspy_and_not_implemented():
    error = _xlsb_loaded._unsupported("style changes")
    assert isinstance(error, UnsupportedFeatureError)
    assert isinstance(error, NotImplementedError)
    assert "style changes" in str(error)


@pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")
def test_native_reader_raises_invalid_file_error_while_iterating():
    # A valid row and cell, then a record whose header runs past the end of the
    # part. The native reader raises from its iterator, not from the call.
    row = b"\x00\x04" + struct.pack("<I", 0)
    cell = b"\x01\x08" + struct.pack("<II", 0, 0)
    data = b"\x91\x00" + row + cell + b"\x00\x10"

    rows = _accel.read_xlsb_worksheet(data, [], [], set(), False)
    assert rows is not None
    with pytest.raises(InvalidFileError) as info:
        list(rows)
    assert isinstance(info.value, ValueError)
