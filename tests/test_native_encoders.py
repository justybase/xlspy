"""Native row encoders: failed and fallback rows leave the shared-string table alone.

These checks pin the table contract: a row encoder that raises or declines a
row must leave the caller's shared-string table and counters exactly as they
were, and the next successful row must still index correctly.
"""

import pytest

from xlspy import Formula, _accel

pytestmark = pytest.mark.skipif(not _accel.HAVE_NATIVE, reason="native extension not built")


class BadHash(str):
    """A str whose hash raises, so the table lookup fails after earlier cells were added."""

    def __hash__(self):
        raise RuntimeError("hash failed on purpose")


def test_xlsb_row_that_raises_leaves_the_table_as_it_was():
    ss_dict, ss_list = {}, []
    with pytest.raises(RuntimeError, match="hash failed"):
        _accel.encode_xlsb_row(["first new text", BadHash("bad")], ss_dict, ss_list, 0, 0, 1)
    assert ss_list == []
    assert ss_dict == {}


def test_xlsx_row_that_raises_leaves_the_table_as_it_was():
    ss_dict, ss_list = {}, []
    with pytest.raises(RuntimeError, match="hash failed"):
        _accel.encode_xlsx_row(["first new text", BadHash("bad")], ss_dict, ss_list, 0, 0, 1, True)
    assert ss_list == []
    assert ss_dict == {}


def test_xlsb_fallback_row_leaves_the_table_as_it_was():
    # A formula cell is refused by the native encoder, so the whole row goes to Python.
    ss_dict, ss_list = {}, []
    assert _accel.encode_xlsb_row(["new", Formula("A1+1")], ss_dict, ss_list, 0, 0, 1) is None
    assert ss_list == []
    assert ss_dict == {}


def test_table_indexes_stay_correct_after_a_failed_row():
    ss_dict, ss_list = {}, []
    with pytest.raises(RuntimeError):
        _accel.encode_xlsb_row(["x", BadHash("bad")], ss_dict, ss_list, 0, 0, 1)
    _, uniq, all_ = _accel.encode_xlsb_row(["a", "b"], ss_dict, ss_list, 0, 0, 1)
    assert (uniq, all_) == (2, 2)
    assert ss_list == ["a", "b"]
    assert ss_dict == {"a": 0, "b": 1}


@pytest.mark.parametrize("index", ["bogus", -1, 2**32])
def test_corrupted_table_index_raises_instead_of_writing_a_bad_index(index):
    # The table maps text to the indices the writer handed out; any other value
    # must fail loudly rather than reach the file as index 0xFFFFFFFF.
    with pytest.raises((TypeError, OverflowError)):
        _accel.encode_xlsb_row(["a"], {"a": index}, ["a"], 1, 1, 1)
    with pytest.raises((TypeError, OverflowError)):
        _accel.encode_xlsx_row(["a"], {"a": index}, ["a"], 1, 1, 1, True)
