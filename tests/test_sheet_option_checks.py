"""Both writers check write_sheet options the same way, before anything is recorded:
a rejected call leaves the writer usable (xlspy._writer_common.WriterFrontEnd._check_sheet_options)."""
from __future__ import annotations

import pytest

from xlspy import ExcelReader, ExcelTable, XlsbWriter, XlsxWriter

WRITERS = [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")]


@pytest.mark.parametrize(("writer_class", "suffix"), WRITERS)
@pytest.mark.parametrize(("ref", "error"), [(5, TypeError), ("not a range", ValueError)])
def test_bad_auto_filter_ref_is_refused_before_the_rows_stream(tmp_path, writer_class, suffix, ref, error) -> None:
    path = tmp_path / f"book.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(error):
            writer.write_sheet([["a", "b"], [1, 2]], auto_filter_ref=ref)
        writer.write_sheet([["a", "b"], [1, 2]], auto_filter_ref="A1:B2")
    with ExcelReader(str(path)) as reader:
        assert list(reader.get_rows("Data")) == [["a", "b"], [1, 2]]


@pytest.mark.parametrize(("writer_class", "suffix"), WRITERS)
def test_rejected_call_claims_no_table_name(tmp_path, writer_class, suffix) -> None:
    path = tmp_path / f"book.{suffix}"
    table = ExcelTable("Sales", "A1:B2")
    with writer_class(path) as writer:
        writer.add_sheet("Data")
        with pytest.raises(TypeError):
            writer.write_sheet([["a", "b"], [1, 2]], tables=[table], zoom="big")
        writer.write_sheet([["a", "b"], [1, 2]], tables=[table])
