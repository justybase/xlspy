from __future__ import annotations

import datetime
import os
import re
import zipfile

from ._untrusted import SafeZipFile
from .exceptions import InvalidFileError, reading, reading_rows
from typing import Any, TYPE_CHECKING, Literal
from collections.abc import Generator
from types import TracebackType
from xml.parsers.expat import ParserCreate
from ._safe_xml import harden_expat, fromstring as safe_fromstring

from ._errors import CellError
from .biff_reader import BiffReader, CellType
from .biff_utils import iter_xlsb_sheets, xlsb_date_1904
from ._accel import HAVE_NATIVE, read_xlsb_worksheet as _accel_read_xlsb_worksheet
from ._accel import read_xlsb_shared_strings as _accel_read_sst
from ._accel import read_xlsx_shared_strings as _accel_read_xlsx_sst
from ._text import xstring_unescape
from ._xlsb_formulas import ERRORS as _XLSB_ERRORS, FormulaContext
from ._xlsb_table_part import table_catalog
from ._xlsb_read import _date_or_number
from ._xlsb_metadata import parse_comments
from .comments import find_comment_parts, read_comments
from ._metadata import read_defined_names, MAIN, namespace_of, relationship_namespace, relationship_kind

if TYPE_CHECKING:
    from ._metadata import DefinedName
    from .comments import Comment

_SST_ITEM = re.compile(r"<si>(.*?)</si>", re.DOTALL)
_SST_TEXT = re.compile(r"<t(?:\s[^>]*)?>(.*?)</t>", re.DOTALL)
_SST_UNSAFE = re.compile(r"<!\[CDATA\[|<!--|&#|<\w+:|<si\s|<si/>|<t\b[^>]*/>")


def _parse_shared_strings_fast(content: str) -> list[str] | None:
    """Shared strings by regex, or None when the markup needs the XML parser
    (CDATA, comments, numeric references, prefixed or self-closing tags)."""
    if _SST_UNSAFE.search(content) or content.count("<si>") != content.count("</si>"):
        return None
    if "\r" in content:
        # XML end-of-line handling, as the parser applies it (&#13; and
        # _x000D_ keep a carriage return; the first goes to the parser above).
        content = content.replace("\r\n", "\n").replace("\r", "\n")
    strings: list[str] = []
    append = strings.append
    for item in _SST_ITEM.finditer(content):
        parts = _SST_TEXT.findall(item.group(1))
        text = parts[0] if len(parts) == 1 else "".join(parts)
        if "&" in text:
            text = (text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
                    .replace("&apos;", "'").replace("&amp;", "&"))
        append(xstring_unescape(text))
    return strings


_date_excel_masks: set[str] = {
    '[$-F800]dddd,\\ mmmm\\ dd,\\ yyyy',
    'd\\-mm;@',
    'yy\\-mm\\-dd;@',
    '[$-415]d\\ mmm;@',
    '[$-415]d\\ mmm\\ yy;@',
    '[$-415]dd\\ mmm\\ yy;@',
    '[$-415]mmm\\ yy;@',
    '[$-415]mmmm\\ yy;@',
    '[$-415]d\\ mmmm\\ yyyy;@',
    'yyyy\\-mm\\-dd\\ hh:mm',
    'yyyy\\-mm\\-dd\\ hh:mm:ss',
    'yyyy\\-mm\\-dd;@',
    'yyyy\\-mm\\-dd',
    '[$-409]dd\\-mm\\-yy\\ h:mm\\ AM/PM;@',
    'dd\\-mm\\-yy\\ h:mm;@',
    '[$-415]mmmmm;@',
    '[$-415]mmmmm\\.yy;@',
    '\\-m\\-yyyy;@',
    '[$-415]d\\-mmm\\-yyyy;@',
    'd\\-m\\-yyyy;@',
}


def _looks_like_date_format(fmt_code: str) -> bool:
    if not fmt_code:
        return False
    fmt_lower = fmt_code.lower()
    has_y = 'y' in fmt_lower
    has_m = 'm' in fmt_lower
    has_d = 'd' in fmt_lower
    has_h = 'h' in fmt_lower
    has_s = 's' in fmt_lower
    date_indicators = sum([has_y, has_m and not has_h and not has_s, has_d])
    return date_indicators >= 2


_default_number_formats_type_dict: dict[int, type] = {
    0: str,
    1: float,
    2: float,
    3: float,
    4: float,
    5: float,
    6: float,
    7: float,
    9: float,
    14: datetime.datetime,
    15: datetime.datetime,
    16: datetime.datetime,
    17: datetime.datetime,
    18: datetime.datetime,
    19: datetime.datetime,
    20: datetime.datetime,
    21: datetime.datetime,
    22: datetime.datetime,
    44: float,
}


def _get_date_num_fmt_ids(number_formats_type_dict: dict[int, type]) -> set[int]:
    return {k for k, v in number_formats_type_dict.items() if v is datetime.datetime}

_letters = [chr(65 + i) for i in range(26)]
for i in range(26):
    for j in range(26):
        _letters.append(chr(65 + i) + chr(65 + j))
for i in range(24):
    for j in range(26):
        for k in range(26):
            if i == 23 and j > 24:
                break
            _letters.append(chr(65 + i) + chr(65 + j) + chr(65 + k))

_letter_to_column_num: dict[str, int] = {letter: idx for idx, letter in enumerate(_letters)}


class ExcelReader:
    """A fast reader of cell values from XLSX, XLSM and XLSB files.

    Use it as a context manager::

        with ExcelReader("report.xlsb") as reader:
            rows = reader.read_all(reader.get_sheet_names()[0])

    Error cells come back as CellError strings (``"error:#DIV/0!"``); with
    ``errors="openpyxl"`` they are plain text such as ``"#DIV/0!"``. For styles,
    coordinates or editing use ``load_workbook``.
    """
    _open_xml_info_string = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def __init__(
        self,
        path: str | None = None,
        use_memory_stream_for_xlsb: bool = True,
        errors: str = "legacy",
    ) -> None:
        if errors not in {"legacy", "openpyxl"}:
            raise ValueError("errors must be 'legacy' or 'openpyxl'")
        self._errors = errors
        self._archive: zipfile.ZipFile | None = None
        self._is_xlsb = False
        self._worksheet_id_to_name: dict[str, str] = {}
        self._worksheet_name_to_id: dict[str, str] = {}
        self._worksheet_id_to_location: dict[str, str] = {}
        # Relationship ids of chartsheets and dialog sheets: sheets without cells.
        self._cellless_sheet_ids: set[str] = set()
        self._shared_string_array: list[str] = []
        self._styles_cell_xfs_array: list[dict] = []
        self._number_formats_type_dict: dict[int, type] = dict(_default_number_formats_type_dict)
        self._shared_strings_location: str | None = None
        self._styles_location: str | None = None
        self._date_1904 = False
        self._path: str | None = None if path is None else os.fspath(path)
        # The file last opened, named in InvalidFileError messages.
        self._opened_path = self._path or ""
        self.use_memory_stream_for_xlsb = use_memory_stream_for_xlsb

    def open(self, path: str | os.PathLike[str]) -> None:
        """Open ``path`` and read its sheet list, shared strings and styles.

        A name ending in ``.xlsb`` is read as XLSB, any other as XLSX/XLSM. Raises
        InvalidFileError for a damaged file.
        """
        # Re-opening the same instance must not leak the previous archive nor
        # accumulate worksheet/style/shared-string state from the previous file.
        path = os.fspath(path)
        self.close()
        self._opened_path = path
        self._worksheet_id_to_name = {}
        self._worksheet_name_to_id = {}
        self._worksheet_id_to_location = {}
        self._cellless_sheet_ids = set()
        self._shared_string_array = []
        self._styles_cell_xfs_array = []
        self._number_formats_type_dict = dict(_default_number_formats_type_dict)
        self._shared_strings_location = None
        self._styles_location = None
        self._date_1904 = False

        self._is_xlsb = path.lower().endswith('xlsb')
        try:
            with reading(path):
                if self._is_xlsb:
                    self._open_xlsb(path)
                else:
                    self._open_xlsx(path)
        except BaseException:
            # A file that fails to parse must not keep the archive (and, on
            # Windows, the file) open: `with` never reaches __exit__ here.
            self.close()
            raise

    def close(self) -> None:
        """Close the file; calling it again does nothing."""
        if self._archive is not None:
            self._archive.close()
            self._archive = None

    def _zip(self) -> zipfile.ZipFile:
        """The open archive. Reading after close() raises a clear error."""
        if self._archive is None:
            raise ValueError("ExcelReader is closed")
        return self._archive

    def __enter__(self) -> ExcelReader:
        if self._path is not None:
            self.open(self._path)
        return self

    def __exit__(self, exc_type: type[BaseException] | None, exc_val: BaseException | None,
                 exc_tb: TracebackType | None) -> Literal[False]:
        self.close()
        return False

    def get_sheet_names(self) -> list[str]:
        """The sheet names, in workbook order, chartsheets included.

        A chartsheet has no cells: ``get_rows`` yields nothing for it. See
        ``get_chartsheet_names``.
        """
        return list(self._worksheet_name_to_id.keys())

    def get_chartsheet_names(self) -> list[str]:
        """The names of the sheets without cells (chartsheets and dialog sheets), in workbook order."""
        return [name for name, rid in self._worksheet_name_to_id.items() if rid in self._cellless_sheet_ids]

    def _sheet_part(self, sheet_name: str) -> str | None:
        """The package path of a sheet's cells, or None for a sheet without cells.

        An unknown name raises KeyError; a worksheet whose part is missing from the
        package raises InvalidFileError.
        """
        rid = self._worksheet_name_to_id[sheet_name]
        location = self._worksheet_id_to_location.get(rid)
        if location is not None:
            path = self._resolve_path(location)
            try:
                self._zip().getinfo(path)
            except KeyError:
                raise InvalidFileError(f"sheet {sheet_name!r}: the package has no part {path!r}") from None
            return path
        if rid in self._cellless_sheet_ids:
            return None
        raise InvalidFileError(f"sheet {sheet_name!r} has no worksheet part")

    def read_part(self, path: str) -> bytes | None:
        """The bytes of a package part, or None when the package has no such part."""
        try:
            return self._zip().read(path)
        except KeyError:
            return None

    def get_sheet_index(self, sheet_name: str) -> int:
        """Return the 0-based position of ``sheet_name``; raises ValueError when absent."""
        names = self.get_sheet_names()
        for i, n in enumerate(names):
            if n == sheet_name:
                return i
        raise ValueError(f"Sheet '{sheet_name}' not found")

    def get_rows(self, sheet_name: str, *, data_only: bool = True, streaming: bool = False) -> Generator[list[Any], None, None]:
        """Yield values or formula text; streaming=True opts into incremental XLSB."""
        if not isinstance(data_only, bool):
            raise TypeError("data_only must be bool")
        if not isinstance(streaming, bool):
            raise TypeError("streaming must be bool")
        if self._sheet_part(sheet_name) is None:
            return  # a chartsheet: no cells
        if self._is_xlsb:
            if not data_only or streaming:
                from ._xlsb_read import read_xlsb_rows
                rows = read_xlsb_rows(self, sheet_name, data_only=data_only)
            else:
                rows = self._get_rows_xlsb(sheet_name)
        else:
            rows = self._get_rows_xlsx(sheet_name, data_only=data_only)
        rows = reading_rows(self._opened_path, rows)
        if self._errors == "legacy":
            yield from rows
        else:
            for row in rows:
                yield [
                    str(value[6:]) if isinstance(value, CellError) else value
                    for value in row
                ]

    def read_all(self, sheet_name: str, *, data_only: bool = True, streaming: bool = False) -> list[list[Any]]:
        """Return all rows of a sheet as a list; see ``get_rows``."""
        return list(self.get_rows(sheet_name, data_only=data_only, streaming=streaming))

    def get_comments(self, sheet_name: str) -> dict[str, Comment]:
        """Return classic note text and authors, without reading worksheet rows."""
        path = self._sheet_part(sheet_name)
        if path is None:
            return {}
        if self._is_xlsb:
            _, _, parts = find_comment_parts(self._zip().read, self._zip().namelist(), path)
            return parse_comments(self._zip().read(parts['comments'][0]))[0] if 'comments' in parts else {}
        _, _, parts = find_comment_parts(self._zip().read, self._zip().namelist(), path)
        return read_comments(self._zip().read(parts["comments"][0])) if "comments" in parts else {}

    def get_defined_names(self) -> list[DefinedName]:
        """Return global and worksheet-local user names (not built-in names)."""
        if self._is_xlsb:
            return FormulaContext(
                self.get_sheet_names(), self._zip().read('xl/workbook.bin'), self.read_part,
                tables=lambda: table_catalog(self._zip().namelist(), self.read_part),
            ).defined_names()
        return read_defined_names(self._zip().read("xl/workbook.xml"), self.get_sheet_names())

    def _open_xlsx(self, path: str) -> None:
        self._archive = SafeZipFile(path, 'r')

        with self._zip().open('xl/workbook.xml') as f:
            content = f.read().decode('utf-8')
            self._parse_workbook_xml(content)

        self._fill_rels('xml')

        if self._styles_location is not None:
            self._fill_styles()

        if self._shared_strings_location is not None:
            self._fill_shared_strings()

    def _open_xlsb(self, path: str) -> None:
        self._archive = SafeZipFile(path, 'r')

        with self._zip().open('xl/workbook.bin') as f:
            raw = f.read()
        for relationship, name in iter_xlsb_sheets(raw):
            # A sheet without a relationship is listed, and has no part to read.
            self._worksheet_id_to_name[relationship or ""] = name
            self._worksheet_name_to_id[name] = relationship or ""
        self._date_1904 = xlsb_date_1904(raw)

        self._fill_rels('bin')

        if self._styles_location is not None:
            self._fill_bin_styles()

        if self._shared_strings_location is not None:
            self._fill_bin_shared_strings()

    def _parse_workbook_xml(self, content: str) -> None:
        root = safe_fromstring(content)
        main = namespace_of(root.tag) or MAIN  # Transitional or Strict
        rel = relationship_namespace(main)
        for node in root.findall(f"{{{main}}}sheets/{{{main}}}sheet"):
            name, rid = node.get('name'), node.get(f'{{{rel}}}id')
            if name is None or rid is None:
                raise InvalidFileError("workbook.xml: a <sheet> lacks its name or r:id")
            self._worksheet_id_to_name[rid] = name
            self._worksheet_name_to_id[name] = rid
        properties = root.find(f"{{{main}}}workbookPr")
        if properties is not None:
            self._date_1904 = properties.get('date1904', '').strip().lower() in {'1', 'true'}

    def _fill_rels(self, ext: str) -> None:
        rels_path = f'xl/_rels/workbook.{ext}.rels'
        if rels_path not in self._zip().namelist():
            return

        with self._zip().open(rels_path) as f:
            content = f.read().decode('utf-8')

        parser = harden_expat(ParserCreate(namespace_separator='}'))
        parser.StartElementHandler = lambda tag, attrs: self._rels_start_elem(tag, attrs)
        parser.Parse(content, True)

    def _rels_start_elem(self, tag: str, attrs: dict[str, str]) -> None:
        local = tag.rsplit('}', 1)[-1].rsplit(':', 1)[-1]
        if local == 'Relationship':
            target = attrs.get('Target', '')
            rel_type = attrs.get('Type', '')
            r_id = attrs.get('Id', '')

            kind = relationship_kind(rel_type)
            if kind == 'worksheet':
                self._worksheet_id_to_location[r_id] = target
            elif kind in ('chartsheet', 'dialogsheet'):
                self._cellless_sheet_ids.add(r_id)
            elif kind == 'sharedStrings':
                self._shared_strings_location = target
            elif kind == 'styles':
                self._styles_location = target

    def _fill_shared_strings(self) -> None:
        if self._shared_strings_location is None:
            return
        entry_path = self._resolve_path(self._shared_strings_location)
        with self._zip().open(entry_path) as f:
            raw = f.read()
        native = _accel_read_xlsx_sst(raw)
        if native is not None:
            self._shared_string_array = native
            return
        content = raw.decode('utf-8')

        fast = _parse_shared_strings_fast(content)
        if fast is not None:
            self._shared_string_array = fast
            return

        self._shared_string_array = []

        parser = harden_expat(ParserCreate(namespace_separator='}'))
        parser.StartElementHandler = lambda tag, attrs: None
        parser.EndElementHandler = lambda tag: None
        parser.CharacterDataHandler = lambda data: None

        class SstHandler:
            def __init__(self, strings_list: list[str]) -> None:
                self.strings = strings_list
                self.in_si = False
                self.in_t = False
                self.buf = ""

            def start(self, tag: str, attrs: dict[str, str]) -> None:
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'si':
                    self.in_si = True
                    self.buf = ""
                elif local == 't' and self.in_si:
                    self.in_t = True

            def end(self, tag: str) -> None:
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 't' and self.in_t:
                    self.in_t = False
                elif local == 'si':
                    if self.in_si:
                        self.strings.append(xstring_unescape(self.buf))
                        self.in_si = False
                        self.buf = ""

            def chars(self, data: str) -> None:
                if self.in_t:
                    self.buf += data

        h = SstHandler(self._shared_string_array)
        parser.StartElementHandler = h.start
        parser.EndElementHandler = h.end
        parser.CharacterDataHandler = h.chars
        parser.Parse(content, True)

    def _fill_bin_shared_strings(self) -> None:
        if self._shared_strings_location is None:
            return
        entry_path = self._resolve_path(self._shared_strings_location)
        with self._zip().open(entry_path) as f:
            raw = f.read()
        sst = _accel_read_sst(raw)
        if sst is not None:
            self._shared_string_array = sst
            return
        stream = _MemoryStream(raw)
        reader = BiffReader(stream)
        reader.read_shared_strings()
        # The unique count in BrtBeginSst is not trusted to size the list (a
        # damaged count would allocate gigabytes); the items are what counts.
        strings: list[str] = []
        while reader.read_shared_strings():
            val = reader.shared_string_value
            if val is not None:
                strings.append(val)
        self._shared_string_array = strings
        reader.close()

    def _fill_styles(self) -> None:
        self._number_formats_type_dict = dict(_default_number_formats_type_dict)
        if self._styles_location is None:
            return
        entry_path = self._resolve_path(self._styles_location)
        with self._zip().open(entry_path) as f:
            content = f.read().decode('utf-8')

        self._styles_cell_xfs_array = []
        fmt_types = self._number_formats_type_dict

        class StylesHandler:
            def __init__(self, xfs_list: list[dict[str, int]]) -> None:
                self.xfs = xfs_list
                self.in_cell_xfs = False
                self.in_num_fmts = False

            def start(self, tag: str, attrs: dict[str, str]) -> None:
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'cellXfs':
                    self.in_cell_xfs = True
                elif local == 'numFmts':
                    self.in_num_fmts = True
                elif local == 'xf' and self.in_cell_xfs:
                    xf_id = int(attrs.get('xfId', '0'))
                    num_fmt_id = int(attrs.get('numFmtId', '0'))
                    self.xfs.append({'xfId': xf_id, 'numFmtId': num_fmt_id})
                elif local == 'numFmt' and self.in_num_fmts:
                    num_fmt_id = int(attrs.get('numFmtId', '0'))
                    fmt_code = attrs.get('formatCode', '')
                    if num_fmt_id not in fmt_types:
                        if fmt_code in _date_excel_masks or _looks_like_date_format(fmt_code):
                            fmt_types[num_fmt_id] = datetime.datetime
                        else:
                            fmt_types[num_fmt_id] = str

            def end(self, tag: str) -> None:
                local = tag.split('}')[-1] if '}' in tag else tag
                if local == 'cellXfs':
                    self.in_cell_xfs = False
                elif local == 'numFmts':
                    self.in_num_fmts = False

            def chars(self, data: str) -> None:
                pass

        h = StylesHandler(self._styles_cell_xfs_array)
        parser = harden_expat(ParserCreate(namespace_separator='}'))
        parser.StartElementHandler = h.start
        parser.EndElementHandler = h.end
        parser.CharacterDataHandler = h.chars
        parser.Parse(content, True)

    def _fill_bin_styles(self) -> None:
        self._number_formats_type_dict = dict(_default_number_formats_type_dict)
        if self._styles_location is None:
            return
        entry_path = self._resolve_path(self._styles_location)
        with self._zip().open(entry_path) as f:
            raw = f.read()
        stream = _MemoryStream(raw)
        reader = BiffReader(stream)

        self._styles_cell_xfs_array = []
        styles_first_time = True
        format_first_time = True

        while reader.read_styles():
            if reader.in_cell_xf:
                if styles_first_time:
                    reader.read_styles()
                    styles_first_time = False
                num_fmt_id = reader.number_format_index
                xf_id = reader.parent_cell_style_xf
                self._styles_cell_xfs_array.append({'xfId': xf_id, 'numFmtId': num_fmt_id})
            elif reader.in_number_format:
                if format_first_time:
                    reader.read_styles()
                    format_first_time = False
                fmt_code = reader.format_string
                num_fmt_id = reader.format
                if num_fmt_id not in self._number_formats_type_dict:
                    if fmt_code in _date_excel_masks or _looks_like_date_format(fmt_code):
                        self._number_formats_type_dict[num_fmt_id] = datetime.datetime
                    else:
                        self._number_formats_type_dict[num_fmt_id] = str

        reader.close()

    @staticmethod
    def _resolve_path(target: str) -> str:
        if target.startswith('/'):
            return target.lstrip('/')
        if target.startswith('xl/'):
            return target
        return f'xl/{target}'

    def _get_rows_xlsx(self, sheet_name: str, *, data_only: bool = True) -> Generator[list[Any], None, None]:
        from ._xlsx_read import read_xlsx_rows
        yield from read_xlsx_rows(self, sheet_name, data_only=data_only)

    def _get_rows_xlsb(self, sheet_name: str) -> Generator[list[Any], None, None]:
        entry_path = self._sheet_part(sheet_name)
        if entry_path is None:
            return

        with self._zip().open(entry_path) as f:
            raw = f.read()

        if HAVE_NATIVE:
            rows = _accel_read_xlsb_worksheet(
                raw,
                self._shared_string_array,
                self._styles_cell_xfs_array,
                _get_date_num_fmt_ids(self._number_formats_type_dict),
                self._date_1904,
            )
            if rows is not None:
                yield from rows
                return
        yield from self._decode_rows_xlsb(raw)

    def _decode_rows_xlsb(self, raw: bytes) -> Generator[list[Any], None, None]:
        """The Python path of _get_rows_xlsb: decode the worksheet records of
        ``raw`` into rows, padding empty rows and columns with None.

        The twin of the native ``XlsbReader`` (rust/xlspy_core/src/xlsb_read.rs),
        rule for rule, so both backends read even a damaged sheet the same way
        (tests/test_native_fuzz_differential.py checks it):

        - Each cell is at its own column and each row at its own place, as
          ``_xlsx_read.rows_in_place`` has it for XLSX: the rows before the first
          row record are empty lists, the first row's width is the least any
          later row has, and a row is as long as its last cell needs.
        - A cell goes to its column's index in the row, wherever its record is;
          a duplicate replaces the earlier value.
        - The rows between two row records (a gap) are empty rows; a row record
          with a lower index than the one before starts a row without gap rows.
        - Cells before any row record form a row of their own.
        - The data ends at BrtEndSheetData.
        """
        first_width: int | None = None
        reader = BiffReader(_MemoryStream(raw))
        strings = self._shared_string_array
        styles = self._styles_cell_xfs_array
        formats = self._number_formats_type_dict
        date_1904 = self._date_1904

        def finished(row: list[Any]) -> list[Any]:
            nonlocal first_width
            row.extend([None] * ((first_width or 0) - len(row)))
            if first_width is None and row:
                first_width = len(row)
            return row

        def value() -> Any:
            kind = reader.cell_type
            if kind == CellType.sharedString:
                idx = reader.int_value
                return strings[idx] if 0 <= idx < len(strings) else None
            if kind == CellType.stringVal:
                return reader.string_value
            if kind == CellType.boolVal:
                return reader.bool_value
            if kind == CellType.errorVal:
                return CellError("error:" + _XLSB_ERRORS.get(reader.error_code, "#VALUE!"))
            if kind == CellType.doubleVal:
                number = reader.double_val
                xf = reader.xf_index
                if (xf != 0 and xf < len(styles)
                        and formats.get(styles[xf]['numFmtId'], str) is datetime.datetime):
                    return _date_or_number(number, date_1904)
                return int(number) if number.is_integer() else number
            return None

        try:
            # Skip the metadata to the first row record or the first cell.
            cur_row = -1
            seen_row = pending_cell = False
            while reader.read_worksheet():
                if reader.record_id == BiffReader._row:
                    cur_row, seen_row = reader.row_index, True
                    for _ in range(cur_row):
                        yield []
                    break
                if reader.read_cell:
                    pending_cell = True
                    break
            else:
                return

            cells: list[Any] = []
            while pending_cell or reader.read_worksheet():
                pending_cell = False
                if reader.record_id == BiffReader._row:
                    yield finished(cells)
                    new_row = reader.row_index
                    gap = new_row - cur_row - 1 if seen_row else 0
                    cur_row, seen_row, cells = new_row, True, []
                    for _ in range(gap):
                        yield [None] * (first_width or 0)
                    continue
                if not reader.read_cell:
                    continue
                col = reader.column_num
                if col >= len(cells):
                    cells.extend([None] * (col + 1 - len(cells)))
                cells[col] = value()

            if cells or cur_row >= 0:
                yield finished(cells)
        finally:
            reader.close()


class _MemoryStream:
    def __init__(self, data: bytes):
        self._data = data
        self._pos = 0

    def read(self, size: int = -1) -> bytes:
        if size < 0:
            result = self._data[self._pos:]
            self._pos = len(self._data)
            return result
        if self._pos + size > len(self._data):
            result = self._data[self._pos:]
            self._pos = len(self._data)
            return result
        result = self._data[self._pos:self._pos + size]
        self._pos += size
        return result

    def readinto(self, buf: bytearray) -> int:
        size = len(buf)
        if self._pos + size > len(self._data):
            size = len(self._data) - self._pos
            if size <= 0:
                return 0
            buf[:size] = self._data[self._pos:self._pos + size]
            self._pos = len(self._data)
            return size
        buf[:size] = self._data[self._pos:self._pos + size]
        self._pos += size
        return size

    def close(self) -> None:
        pass

    @property
    def length(self) -> int:
        return len(self._data)
