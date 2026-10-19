"""C row encoders: failed and fallback rows leave the shared-string table alone.

The full fuzzer is ``tools/fuzz_c_core.py``; run it with more iterations, or
against a sanitizer build, for deeper checks. The short run here keeps the
harness itself working in the ordinary suite.
"""

import importlib.util
import pathlib

import pytest

from xlspy import Formula, _accel

pytestmark = pytest.mark.skipif(not _accel.HAVE_C_EXT, reason="C extension not built")

ROOT = pathlib.Path(__file__).resolve().parents[1]


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
    # A formula cell is refused by the C encoder, so the whole row goes to Python.
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


def test_fuzz_harness_short_run_holds_every_check():
    spec = importlib.util.spec_from_file_location("fuzz_c_core", ROOT / "tools" / "fuzz_c_core.py")
    fuzz = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fuzz)
    assert fuzz.main(["--iterations", "300", "--seed", "7"]) == 0
