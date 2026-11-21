"""VBA projects: embedding them in XLSX/XLSM output."""

import io
import zipfile

import pytest

from xlspy import Workbook, XlsxWriter, load_workbook


def parts(stream):
    with zipfile.ZipFile(io.BytesIO(stream.getvalue())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def workbook_with_vba(**options):
    stream = io.BytesIO()
    with XlsxWriter(stream, **options) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
        writer.add_vba_project(b"fake vba project")
        writer.set_vba_name("MyProject")
    return stream


def test_xlsxwriter_writes_the_vba_parts():
    archive = parts(workbook_with_vba())
    assert archive["xl/vbaProject.bin"] == b"fake vba project"
    assert b"macroEnabled.main+xml" in archive["[Content_Types].xml"]
    assert b'ContentType="application/vnd.ms-office.vbaProject"' in archive["[Content_Types].xml"]
    assert b"relationships/vbaProject" in archive["xl/_rels/workbook.xml.rels"]
    assert b"vbaProject.bin" in archive["xl/_rels/workbook.xml.rels"]
    assert b'codeName="MyProject"' in archive["xl/workbook.xml"]


def test_xlsxwriter_without_vba_is_plain():
    stream = io.BytesIO()
    with XlsxWriter(stream) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
    archive = parts(stream)
    assert "xl/vbaProject.bin" not in archive
    assert b"sheet.main+xml" in archive["[Content_Types].xml"]
    assert b"relationships/vbaProject" not in archive["xl/_rels/workbook.xml.rels"]


def test_vba_project_accepts_a_path_and_a_stream(tmp_path):
    path = tmp_path / "vbaProject.bin"
    path.write_bytes(b"from path")
    stream = io.BytesIO()
    with XlsxWriter(stream) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
        writer.add_vba_project(path)
    assert parts(stream)["xl/vbaProject.bin"] == b"from path"

    stream = io.BytesIO()
    with XlsxWriter(stream) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
        writer.add_vba_project(io.BytesIO(b"from stream"))
    assert parts(stream)["xl/vbaProject.bin"] == b"from stream"


def test_vba_input_is_checked():
    with XlsxWriter(io.BytesIO()) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([[1]])
        with pytest.raises(ValueError, match="empty"):
            writer.add_vba_project(b"")
        with pytest.raises(TypeError, match="binary"):
            writer.add_vba_project(io.StringIO("text"))
        with pytest.raises(TypeError, match="path, bytes"):
            writer.add_vba_project(42)
        with pytest.raises(ValueError, match="empty"):
            writer.set_vba_name("")
        with pytest.raises(TypeError, match="string"):
            writer.set_vba_name(1)


def test_workbook_model_saves_a_macro_workbook(tmp_path):
    path = tmp_path / "macro.xlsm"
    workbook = Workbook()
    workbook.active["A1"] = 1
    workbook.add_vba_project(b"model vba")
    workbook.set_vba_name("ModelProject")
    workbook.save(path)

    with zipfile.ZipFile(path) as archive:
        assert archive.read("xl/vbaProject.bin") == b"model vba"
        assert b"macroEnabled.main+xml" in archive.read("[Content_Types].xml")
        assert b'codeName="ModelProject"' in archive.read("xl/workbook.xml")
        assert b"relationships/vbaProject" in archive.read("xl/_rels/workbook.xml.rels")

    loaded = load_workbook(path, keep_vba=True)
    try:
        assert loaded.vba_archive is not None
        assert loaded["Sheet"]["A1"].value == 1
    finally:
        loaded.close()


def test_workbook_model_refuses_to_drop_vba(tmp_path):
    workbook = Workbook()
    workbook.add_vba_project(b"model vba")
    with pytest.raises(ValueError, match="xlsm"):
        workbook.save(tmp_path / "plain.xlsx")
