"""LET, LAMBDA and dynamic-array spill formulas in XLSX and XLSB."""

import zipfile

import pytest

from xlspy import (
    ExcelReader, Formula, UnsupportedFeatureError, XlsbUpdater, XlsbWriter,
    XlsxUpdater, XlsxWriter, load_workbook,
)


def _parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _formula_text(path, suffix):
    with ExcelReader(path) as reader:
        return reader.read_all("S", data_only=False)


def test_let_and_lambda_text_is_normalized():
    assert Formula("LET(x,1,x+1)").text == "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"
    assert Formula("LET(x,1,y,2,x+y)").text == "_xlfn.LET(_xlpm.x,1,_xlpm.y,2,_xlpm.x+_xlpm.y)"
    assert Formula("LAMBDA(x,x*2)(3)").text == "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3)"
    assert Formula("MAP(A1:A3,LAMBDA(v,v*2))").text == (
        "_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))"
    )
    assert Formula("LET(x,LET(y,1,y+1),x+y)").text == (
        "_xlfn.LET(_xlpm.x,_xlfn.LET(_xlpm.y,1,_xlpm.y+1),_xlpm.x+y)"
    )


def test_structured_references_are_not_lambda_parameters():
    assert Formula("LAMBDA(x,Table1[x])").text == "_xlfn.LAMBDA(_xlpm.x,Table1[x])"
    assert Formula("LET(a,1,Table1[[#This Row],[a]])").text == (
        "_xlfn.LET(_xlpm.a,1,Table1[[#This Row],[a]])"
    )
    assert Formula("LAMBDA(x,Sheet1!x)").text == "_xlfn.LAMBDA(_xlpm.x,Sheet1!x)"
    assert Formula("LET(x,1,Sheet1!x+x)").text == (
        "_xlfn.LET(_xlpm.x,1,Sheet1!x+_xlpm.x)"
    )


def test_spill_ref_and_array_ref_are_exclusive():
    with pytest.raises(ValueError, match="array_ref or spill_ref"):
        Formula("SEQUENCE(2)", array_ref="A1:A2", spill_ref="A1:A2")
    with pytest.raises(ValueError, match="invalid A1|A1 range"):
        Formula("SEQUENCE(2)", spill_ref="nope")


@pytest.mark.parametrize("writer_class, suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_spill_writes_its_metadata(tmp_path, writer_class, suffix):
    path = tmp_path / f"spill.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[Formula("SEQUENCE(3,2)", spill_ref="A1:B3")]])
    parts = _parts(path)
    if suffix == "xlsx":
        assert "xl/metadata.xml" in parts
        sheet = parts["xl/worksheets/sheet1.xml"].decode()
        assert 'cm="1"' in sheet
        assert 't="array" ref="A1:B3"' in sheet
        assert b"sheetMetadata" in parts["xl/_rels/workbook.xml.rels"]
        assert b"/xl/metadata.xml" in parts["[Content_Types].xml"]
    else:
        assert "xl/metadata.bin" in parts
        sheet = parts["xl/worksheets/sheet1.bin"]
        from xlspy.biff_utils import iter_records

        record_ids = {record.record_id for record in iter_records(sheet)}
        assert 0x1AA in record_ids  # the array formula group
        assert 0x31 in record_ids  # the cell metadata record
        assert b"sheetMetadata" in parts["xl/_rels/workbook.bin.rels"]
        assert b"/xl/metadata.bin" in parts["[Content_Types].xml"]
    # The formula text reads back through both readers.
    with ExcelReader(path) as reader:
        rows = reader.read_all("S", data_only=False)
    assert rows[0][0] == "=_xlfn.SEQUENCE(3,2)"


@pytest.mark.parametrize("writer_class, suffix", [(XlsxWriter, "xlsx"), (XlsbWriter, "xlsb")])
def test_let_and_lambda_read_back(tmp_path, writer_class, suffix):
    path = tmp_path / f"let.{suffix}"
    with writer_class(path) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[
            Formula("LET(x,1,x+1)"),
            None,
            Formula("MAP(B1:B3,LAMBDA(v,v*2))", spill_ref="C1:C3"),
        ], [1], [2], [3]])
    rows = _formula_text(path, suffix)
    assert rows[0][0] == "=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"
    assert rows[0][2] == "=_xlfn.MAP(B1:B3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))"
    # A loaded workbook saves the same file.
    loaded = load_workbook(path, data_only=False)
    try:
        out = tmp_path / f"out.{suffix}"
        loaded.save(out)
    finally:
        loaded.close()
    assert _parts(out) == _parts(path)


@pytest.mark.parametrize(
    "writer_class, updater_class, suffix",
    [(XlsxWriter, XlsxUpdater, "xlsx"), (XlsbWriter, XlsbUpdater, "xlsb")],
)
def test_updaters_refuse_new_spill_formulas(
    tmp_path, writer_class, updater_class, suffix
):
    source = tmp_path / f"plain.{suffix}"
    with writer_class(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x"], [1]])
    updater = updater_class(source)
    try:
        with pytest.raises(UnsupportedFeatureError, match="dynamic"):
            updater.update_cells(
                "S", {"A2": Formula("SEQUENCE(3)", spill_ref="A2:A4")}
            )
    finally:
        updater.close()


@pytest.mark.parametrize("stream", [False, True])
def test_xlsb_replace_sheet_data_writes_spill_members(tmp_path, stream):
    source = tmp_path / "source.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x", "y"]])
    updater = XlsbUpdater(source)
    try:
        replace = updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
        replace(
            "S",
            [
                [Formula("SEQUENCE(2,2)", spill_ref="A2:B3"), 99],
                [7, 8],
            ],
            headers=["x", "y"],
        )
        updater.save(source)
    finally:
        updater.close()

    with ExcelReader(source) as reader:
        rows = reader.read_all("S", data_only=False)
    # B2 and the whole second row are array members, not ordinary cells.
    assert rows[1][1] == 99
    assert rows[2] == [7, 8]
    parts = _parts(source)
    from xlspy.biff_utils import iter_records

    record_ids = {record.record_id for record in iter_records(parts["xl/worksheets/sheet1.bin"])}
    assert 0x1AA in record_ids  # the array formula group
    assert 0x31 in record_ids  # dynamic-array cell metadata


@pytest.mark.parametrize("stream", [False, True])
def test_xlsb_replace_sheet_data_writes_spill_rows_past_the_data(tmp_path, stream):
    source = tmp_path / "source.xlsb"
    with XlsbWriter(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x", "y"]])
    updater = XlsbUpdater(source)
    try:
        replace = updater.replace_sheet_data_stream if stream else updater.replace_sheet_data
        replace(
            "S",
            [[Formula("SEQUENCE(3)", spill_ref="A2:A4")]],
            headers=["x", "y"],
        )
        updater.save(source)
    finally:
        updater.close()

    with ExcelReader(source) as reader:
        rows = reader.read_all("S", data_only=False)
    assert len(rows) == 4
    assert rows[1][0] == "=_xlfn.SEQUENCE(3)"
    assert rows[2][0] == 0 and rows[3][0] == 0


@pytest.mark.parametrize(
    "writer_class, updater_class, suffix",
    [(XlsxWriter, XlsxUpdater, "xlsx"), (XlsbWriter, XlsbUpdater, "xlsb")],
)
def test_replace_sheet_data_writes_spill_formulas(
    tmp_path, writer_class, updater_class, suffix
):
    source = tmp_path / f"plain.{suffix}"
    with writer_class(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([["x"], [1]])
    updater = updater_class(source)
    try:
        updater.replace_sheet_data(
            "S", [[Formula("SEQUENCE(3)", spill_ref="A2:A4")]], headers=["x"]
        )
        updater.save(source)
    finally:
        updater.close()
    parts = _parts(source)
    assert ("xl/metadata.xml" if suffix == "xlsx" else "xl/metadata.bin") in parts
    with ExcelReader(source) as reader:
        rows = reader.read_all("S", data_only=False)
    assert rows[1][0] == "=_xlfn.SEQUENCE(3)"


@pytest.mark.parametrize(
    "writer_class, updater_class, suffix",
    [(XlsxWriter, XlsxUpdater, "xlsx"), (XlsbWriter, XlsbUpdater, "xlsb")],
)
def test_updater_edits_beside_a_spill_and_refuses_inside(
    tmp_path, writer_class, updater_class, suffix
):
    source = tmp_path / f"spill.{suffix}"
    with writer_class(source) as writer:
        writer.add_sheet("S")
        writer.write_sheet([[Formula("SEQUENCE(3,2)", spill_ref="A1:B3")]])

    updater = updater_class(source)
    try:
        # A cell in the anchor's row but outside the spill range is safe.
        updater.update_cells("S", {"D1": 5})
        with pytest.raises(ValueError, match="array"):
            updater.update_cells("S", {"A2": 5})
        updater.save(source)
    finally:
        updater.close()

    parts = _parts(source)
    assert ("xl/metadata.xml" if suffix == "xlsx" else "xl/metadata.bin") in parts
    with ExcelReader(source) as reader:
        rows = reader.read_all("S", data_only=False)
    assert rows[0][0] == "=_xlfn.SEQUENCE(3,2)"
    assert rows[0][3] == 5
