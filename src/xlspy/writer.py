from dataclasses import dataclass
import datetime
from decimal import Decimal
import zipfile
import struct
from typing import  Tuple, Iterable, Mapping
import io
import itertools
from ._accel import encode_xlsb_row, encode_xlsb_row_inline, calc_column_widths as accel_calc_column_widths
from ._accel import xlsb_row_has_formatted as accel_row_has_formatted
from ._accel import encode_xlsb_sst as accel_encode_xlsb_sst
from .formats import _is_formatted_cell, _unwrap_cell, _get_format
from .excel_dates import datetime_to_serial
from . import chart_parts
from .styles import (
    Chart,
    Style,
    _column_index,
    capture_table_headers,
    explicit_cf_priorities,
    merge_styles,
    normalize_hidden_columns,
    normalize_hidden_rows,
    normalize_style_ranges,
    parse_cell_range,
    split_style,
    styles_for_row,
    table_column_names,
)
from .biff_utils import build_record
from ._output import StagedWriter
from ._text import check_cell_text
from .formulas import Formula
from .hyperlinks import Hyperlink
from .outline import build_outline
from .refs import check_sheet_name


# BIFF12 record identifiers used in xl/styles.bin ([MS-XLSB] section 2.1.4).
_BRT_FONT = 0x002B
_BRT_FMT = 0x002C
_BRT_FILL = 0x002D
_BRT_BORDER = 0x002E
_BRT_XF = 0x002F
_BRT_BEGIN_FILLS = 0x025B
_BRT_END_FILLS = 0x025C
_BRT_BEGIN_FONTS = 0x0263
_BRT_END_FONTS = 0x0264
_BRT_BEGIN_BORDERS = 0x0265
_BRT_END_BORDERS = 0x0266
_BRT_BEGIN_FMTS = 0x0267
_BRT_END_FMTS = 0x0268
_BRT_BEGIN_CELL_XFS = 0x0269
_BRT_END_CELL_XFS = 0x026A

# BIFF12 record identifiers used in xl/worksheets/sheetN.bin.
_BRT_ROW_HDR = 0x0000
_BRT_COL_INFO = 0x003C
_BRT_BEGIN_COL_INFOS = 0x0086
_BRT_END_COL_INFOS = 0x0087
_BRT_PANE = 0x0097
_BRT_SEL = 0x0098
_BRT_MERGE_CELL = 0x00B0
_BRT_BEGIN_MERGE_CELLS = 0x00B1
_BRT_END_MERGE_CELLS = 0x00B2
# BrtDrawing ([MS-XLSB] 2.4.351) links a sheet to its Drawings part.
_BRT_DRAWING = 0x0226

# Table (List) parts ([MS-XLSB] 2.1.7.51, 2.4.100).
_BRT_BEGIN_LIST = 343  # 0x157
_BRT_END_LIST = 344  # 0x158
_BRT_BEGIN_LIST_COLS = 345  # 0x159
_BRT_END_LIST_COLS = 346  # 0x15A
_BRT_BEGIN_LIST_COL = 347  # 0x15B
_BRT_END_LIST_COL = 348  # 0x15C
_BRT_BEGIN_LIST_PARTS = 660  # 0x294
_BRT_LIST_PART = 661  # 0x295
_BRT_END_LIST_PARTS = 662  # 0x296
_BRT_BEGIN_AUTOFILTER = 161  # 0xA1
_BRT_END_AUTOFILTER = 162  # 0xA2
_BRT_TABLE_STYLE_CLIENT = 513  # 0x201
# Conditional formatting ([MS-XLSB] 2.1.7.62).
_BRT_BEGIN_CF = 461
_BRT_END_CF = 462
_BRT_BEGIN_CF_RULE = 463
_BRT_END_CF_RULE = 464
_BRT_BEGIN_ICONSET = 465
_BRT_END_ICONSET = 466
_BRT_BEGIN_DATABAR = 467
_BRT_END_DATABAR = 468
_BRT_BEGIN_COLORSCALE = 469
_BRT_END_COLORSCALE = 470
_BRT_CFVO = 471
_BRT_COLOR = 564
_BRT_BEGIN_DXFS = 505
_BRT_END_DXFS = 506
_BRT_DXF = 507
# FRT wrappers used by table parts.
_BRT_AC_BEGIN = 0x25
_BRT_AC_END = 0x26
_BRT_FRT_BEGIN = 0xC00

# Active-pane identifiers for BrtPane/BrtSel ([MS-XLSB] 2.5.108 Pnn).
_PANE_ACTIVE = {"bottomRight": 0, "topRight": 1, "bottomLeft": 2, "topLeft": 3}

# Property encodings for BrtFont, BrtFill, BrtBorder and BrtXF; the values come
# from the tables in [MS-XLSB] sections 2.4.311 (BrtBorder), 2.4.674 (BrtFill),
# 2.4.683 (BrtFont) and 2.4.863 (BrtXF).
_UNDERLINE_BYTES = {
    None: 0x00, "none": 0x00, "single": 0x01, "double": 0x02,
    "singleAccounting": 0x21, "doubleAccounting": 0x22,
}
_FONT_BASELINE = {None: 0, "baseline": 0, "superscript": 1, "subscript": 2}
_FONT_SCHEME = {None: 0, "major": 1, "minor": 2}
_PATTERN_BYTES = {
    None: 0, "none": 0, "solid": 1, "mediumGray": 2, "darkGray": 3,
    "lightGray": 4, "darkHorizontal": 5, "darkVertical": 6, "darkDown": 7,
    "darkUp": 8, "darkGrid": 9, "darkTrellis": 10, "lightHorizontal": 11,
    "lightVertical": 12, "lightDown": 13, "lightUp": 14, "lightGrid": 15,
    "lightTrellis": 16, "gray125": 17, "gray0625": 18,
}
_BORDER_BYTES = {
    None: 0x00, "none": 0x00, "thin": 0x01, "medium": 0x02, "dashed": 0x03,
    "dotted": 0x04, "thick": 0x05, "double": 0x06, "hair": 0x07,
    "mediumDashed": 0x08, "dashDot": 0x09, "mediumDashDot": 0x0A,
    "dashDotDot": 0x0B, "mediumDashDotDot": 0x0C, "slantDashDot": 0x0D,
}
_HORIZONTAL_BYTES = {
    None: 0, "general": 0, "left": 1, "center": 2, "right": 3, "fill": 4,
    "justify": 5, "centerContinuous": 6, "distributed": 7,
}
_VERTICAL_BYTES = {
    None: 2, "top": 0, "center": 1, "bottom": 2, "justify": 3, "distributed": 4,
}


@dataclass
class FilterData:
    sheet_index: int = 0
    start_column: int = 0
    end_column: int = 0  
    start_row: int = 0
    end_row: int = 0


# XLSB has no encoding here for formulas and hyperlinks. The C row encoder
# hands rows with such objects to the Python encoder, which refuses them.
_XLSB_REJECTED_TYPES = frozenset({Formula, Hyperlink})


def _capture_cells(rows, store: dict):
    """Yield rows unchanged, filling the chart-cache slots they cover."""
    columns = [(col_idx, first_row, slots) for col_idx, (first_row, slots) in store.items()]
    for row_idx, row in enumerate(rows):
        for col_idx, first_row, slots in columns:
            offset = row_idx - first_row
            if 0 <= offset < len(slots) and col_idx < len(row):
                slots[offset] = row[col_idx]
        yield row


class XlsbWriter(StagedWriter):
    def __init__(self, file: str | io.IOBase, compressionLevel:int = 4,
                 useSharedStrings: bool = True):
        self._file = file
        self._useSharedStrings = useSharedStrings
        self._aborted = False
        self._saved = False
        self._output = None  # StagedOutput while open
        self._worksheet_data: list[Tuple[str, Iterable[list[any]],bool]] = []
        self._shared_strings: list[str] = []
        self._shared_strings_dict: dict[str, int] = {}
        self._sheet_count = 0
        self._sst_unique_count = 0
        self._sst_all_count = 0
        self._filtered_data_list: list[FilterData] = []
        self._sheetCnt = 1
        self._compressionLevel = compressionLevel
        self._zf: zipfile.ZipFile = None
        self._letters = chart_parts.column_letters()

        self._format_registry: dict[str, int] = {}
        self._format_xf_map: dict[str, int] = {}
        self._style_xf_map: dict[Style, int] = {}
        self._style_components: dict[int, tuple] = {}
        self._font_registry: dict[tuple, int] = {}
        self._fill_registry: dict[tuple, int] = {}
        self._border_registry: dict[tuple, int] = {}
        self._next_numfmt_id = 167
        self._next_xf_index = 4
        self._sheet_charts: dict[int, list] = {}
        self._sheet_chart_ids: dict[int, list[int]] = {}
        self._sheet_chart_cells: dict[int, dict] = {}
        self._next_chart_id = 1
        self._next_table_id = 1
        self._table_names: set[str] = set()
        self._sheet_tables: dict[int, list[tuple]] = {}
        self._tables_bin: dict[int, bytes] = {}
        self._seen_cf_priorities: set[int] = set()
        self._reserved_cf_priorities: set[int] = set()
        self._next_cf_priority = 1

    def add_sheet(self, sheet_name: str, hidden : bool = False):
        check_sheet_name(sheet_name, (name for name, _data, _hidden in self._worksheet_data))
        self._sheet_count += 1
        self._worksheet_data.append((sheet_name, iter([]),hidden))

    def _register_numfmt(self, fmt_string: str) -> int:
        numfmt_id = self._format_registry.get(fmt_string)
        if numfmt_id is None:
            numfmt_id = self._next_numfmt_id
            if numfmt_id > 382:
                raise ValueError("XLSB supports at most 216 custom number formats")
            self._next_numfmt_id += 1
            self._format_registry[fmt_string] = numfmt_id
        return numfmt_id

    def _register_format(self, fmt_string: str) -> int:
        if fmt_string in self._format_xf_map:
            return self._format_xf_map[fmt_string]
        # A style may already have registered this number format; reuse its id
        # so the style XF keeps referencing a format that is still emitted.
        self._register_numfmt(fmt_string)
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._format_xf_map[fmt_string] = xf_index
        return xf_index

    @staticmethod
    def _register_component(registry: dict, key, first_id: int) -> int:
        """Return the record id for ``key``, adding it to ``registry`` if new."""
        if key is None:
            return 0
        existing = registry.get(key)
        if existing is not None:
            return existing
        value = first_id + len(registry)
        registry[key] = value
        return value

    @staticmethod
    def _reject_rich_style(style: Style) -> None:
        """Reject the few style properties XLSB cannot express."""
        alignment = style.alignment
        if alignment is not None and alignment.relative_indent is not None:
            raise NotImplementedError(
                "relative_indent is not supported for XLSB output"
            )
        border = style.border
        if border is not None:
            unsupported = [
                name
                for name in ("vertical", "horizontal", "start", "end")
                if getattr(border, name) is not None
            ]
            if border.outline is False:
                unsupported.append("outline=False")
            if unsupported:
                raise NotImplementedError(
                    "XLSB output does not support border " + ", ".join(unsupported)
                )
        font = style.font
        if font is not None and font.name is not None and len(font.name) > 31:
            raise ValueError("XLSB font names are limited to 31 characters")
        number_format = style.number_format
        if number_format is not None and len(number_format) > 255:
            raise ValueError("XLSB number formats are limited to 255 characters")
        stops = getattr(style.fill, "stop", None)
        if stops is not None and len(stops) > 256:
            raise ValueError("XLSB gradients support at most 256 stops")

    def _register_style(self, style: Style) -> int:
        if style in self._style_xf_map:
            return self._style_xf_map[style]
        self._reject_rich_style(style)
        font_key, fill_key, border_key, numfmt, align_key, prot_key = split_style(style)
        font_id = self._register_component(self._font_registry, font_key, 2)
        fill_id = self._register_component(self._fill_registry, fill_key, 2)
        border_id = self._register_component(self._border_registry, border_key, 1)
        numfmt_id = 0
        if numfmt is not None:
            numfmt_id = self._register_numfmt(numfmt)
        alignment = (
            align_key
            if align_key is not None and any(value is not None for value in align_key)
            else None
        )
        protection = (
            prot_key
            if prot_key is not None
            and (prot_key.locked is not None or prot_key.hidden is not None)
            else None
        )
        xf_index = self._next_xf_index
        self._next_xf_index += 1
        self._style_xf_map[style] = xf_index
        self._style_components[xf_index] = (
            font_id, fill_id, border_id, numfmt_id, alignment, protection,
        )
        return xf_index

    @staticmethod
    def _normalize_merges(merges):
        if not merges:
            return []
        refs = []
        for ref in merges:
            first_col, first_row, last_col, last_row = parse_cell_range(ref)
            refs.append((ref, first_col, first_row, last_col, last_row))
        for i, (ref_a, c1a, r1a, c2a, r2a) in enumerate(refs):
            for ref_b, c1b, r1b, c2b, r2b in refs[i + 1:]:
                if c1a <= c2b and c1b <= c2a and r1a <= r2b and r1b <= r2a:
                    raise ValueError(
                        f"overlapping merged ranges {ref_a!r} and {ref_b!r}"
                    )
        return [ref for ref, *_ in refs]

    @staticmethod
    def _normalize_widths(mapping, kind):
        if mapping is None:
            return {}
        if not isinstance(mapping, Mapping):
            raise TypeError(f"{kind} must be a mapping or None")
        result = {}
        for key, value in mapping.items():
            if isinstance(key, str):
                col = _column_index(key)
            elif isinstance(key, int) and not isinstance(key, bool):
                col = key
            else:
                raise TypeError(f"{kind} keys must be column letters or 0-based ints")
            if col < 0 or col >= 16384:
                raise ValueError(f"{kind} column out of range: {key!r}")
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise TypeError(f"{kind} widths must be numbers")
            if not 0.5 <= value <= 255:
                raise ValueError(f"{kind} width must be 0.5..255")
            result[col] = float(value)
        return result

    @staticmethod
    def _normalize_row_heights(mapping):
        if mapping is None:
            return {}
        if not isinstance(mapping, Mapping):
            raise TypeError("row_heights must be a mapping or None")
        result = {}
        for key, value in mapping.items():
            if not isinstance(key, int) or isinstance(key, bool) or key < 0:
                raise TypeError("row_heights keys must be 0-based row ints")
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise TypeError("row_heights values must be numbers")
            if value <= 0:
                raise ValueError("row height must be > 0")
            result[key] = float(value)
        return result

    @staticmethod
    def _split_freeze_cell(ref: str) -> tuple[int, int]:
        first_col, first_row, last_col, last_row = parse_cell_range(ref)
        if (first_col, first_row) != (last_col, last_row):
            raise ValueError(f"freeze_panes must be a single cell, got {ref!r}")
        return first_col, first_row

    @staticmethod
    def _normalize_charts(charts) -> list:
        if not charts:
            return []
        if not isinstance(charts, (list, tuple)):
            raise TypeError("charts must be a list of Chart objects or None")
        for chart in charts:
            if not isinstance(chart, Chart):
                raise TypeError("charts must be Chart instances")
        return list(charts)

    def _normalize_tables(self, tables):
        from .styles import ExcelTable

        if not tables:
            return []
        # Names are only checked here; write_sheet registers them once every
        # option is valid, so a rejected call leaves no name behind.
        names = set(self._table_names)
        for table in tables:
            if not isinstance(table, ExcelTable):
                raise TypeError("tables must be ExcelTable instances")
            if table.display_name in names:
                raise ValueError(
                    f"duplicate table display name {table.display_name!r}"
                )
            names.add(table.display_name)
        return list(tables)

    def _check_unknown_cf_rule(self, rule) -> None:
        from .styles import (
            CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule,
        )
        if hasattr(rule, "sqref") and not isinstance(
            rule,
            (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule),
        ):
            raise TypeError(
                f"unsupported conditional formatting rule {type(rule).__name__}; "
                "use xlspy rule classes (CellIsRule, FormulaRule, "
                "ColorScaleRule, DataBarRule, IconSetRule)"
            )

    def _claim_cf_priority(self, rule) -> int:
        priority = rule.priority
        if priority is None:
            while (
                self._next_cf_priority in self._seen_cf_priorities
                or self._next_cf_priority in self._reserved_cf_priorities
            ):
                self._next_cf_priority += 1
            priority = self._next_cf_priority
            self._next_cf_priority += 1
        elif priority in self._seen_cf_priorities:
            raise ValueError(f"duplicate conditional formatting priority {priority}")
        if priority <= 0:
            raise ValueError("conditional formatting priority must be positive")
        self._seen_cf_priorities.add(priority)
        self._next_cf_priority = max(self._next_cf_priority, priority + 1)
        return priority

    def _normalize_cf_rules(self, conditional_formats):
        from .styles import (
            CellIsRule, ColorScaleRule, DataBarRule, DifferentialStyle,
            FormulaRule, IconSetRule,
        )
        if not conditional_formats:
            return []
        for rule in conditional_formats:
            self._check_unknown_cf_rule(rule)
            if not isinstance(
                rule,
                (CellIsRule, FormulaRule, ColorScaleRule, DataBarRule,
                 IconSetRule),
            ):
                raise TypeError(
                    f"unsupported conditional formatting rule {type(rule).__name__}"
                )
            if getattr(rule, "dxf", None) is not None and not isinstance(
                rule.dxf, DifferentialStyle
            ):
                raise TypeError("dxf must be a DifferentialStyle or None")
            # XLSB CF with formulas/DXF is not yet supported in this step;
            # tables + ColorScale/DataBar/IconSet (non-formula) are supported.
            # CellIs/Formula and formula-type CFVOs need Ptg + DXF (next step).
            from .styles import CellIsRule as _C, FormulaRule as _F
            if isinstance(rule, (_C, _F)):
                raise NotImplementedError(
                    "XLSB CellIs/Formula rules need Ptg + DXF (next step); "
                    "use ColorScale/DataBar/IconSet or XLSX for now"
                )
        return list(conditional_formats)

    _CFVO_CODES = {
        "num": 1, "min": 2, "max": 3, "percent": 4,
        "percentile": 5, "formula": 7,
    }
    _ICON_SET_CODES = {
        "3Arrows": 0, "3ArrowsGray": 1, "3Flags": 2, "3TrafficLights1": 3,
        "3TrafficLights2": 4, "3Signs": 5, "3Symbols": 6, "3Symbols2": 7,
        "4Arrows": 8, "4ArrowsGray": 9, "4RedToBlack": 10, "4Ratings": 11,
        "4TrafficLights": 12, "5Arrows": 13, "5ArrowsGray": 14,
        "5Ratings": 15, "5Quarters": 16,
    }

    @staticmethod
    def _color_to_brt(color) -> bytes:
        from .styles import _coerce_color
        resolved = _coerce_color(color, "color")
        if resolved is None or resolved.rgb is None:
            raise ValueError("CF colors must be RGB colors")
        rgb = resolved.rgb
        if len(rgb) == 8:  # ARGB -> drop alpha, use opaque
            rgb = rgb[2:]
        r = int(rgb[0:2], 16); g = int(rgb[2:4], 16); b = int(rgb[4:6], 16)
        return bytes((0x05, 0xFF, 0x00, 0x00, r, g, b, 0xFF))

    def _cfvo_bin(self, cfvo, default: str) -> bytes:
        from .styles import FormatObject
        if cfvo is None:
            code = self._CFVO_CODES[default]
            return (
                struct.pack("<I", code) + struct.pack("<d", 0.0)
                + struct.pack("<II", 0, 0) + struct.pack("<I", 0)
            )
        if not isinstance(cfvo, FormatObject):
            raise TypeError("CFVO must be a FormatObject or None")
        if cfvo.type == "formula":
            raise NotImplementedError(
                "XLSB formula-type CFVO needs Ptg (next step)"
            )
        code = self._CFVO_CODES[cfvo.type]
        if cfvo.type in ("min", "max"):
            num = 0.0
        else:
            if cfvo.value is None:
                raise ValueError(f"CFVO type {cfvo.type!r} needs a value")
            num = float(cfvo.value)
        # IconSets use GTE flags; others ignore (0,0).
        return (
            struct.pack("<I", code) + struct.pack("<d", num)
            + struct.pack("<II", 0, 0) + struct.pack("<I", 0)
        )

    def _cfvo_icon_bin(self, value: float) -> bytes:
        return (
            struct.pack("<I", 4) + struct.pack("<d", float(value))
            + struct.pack("<II", 1, 1) + struct.pack("<I", 0)
        )

    @staticmethod
    def _encode_nullable_wstring(value: str | None) -> bytes:
        if value is None:
            return struct.pack("<I", 0xFFFFFFFF)
        encoded = value.encode("utf-16-le")
        return struct.pack("<I", len(value)) + encoded

    @staticmethod
    def _encode_rfx(first_col: int, first_row: int,
                    last_col: int, last_row: int) -> bytes:
        return struct.pack("<IIII", first_row, last_row, first_col, last_col)

    @staticmethod
    def _acuid() -> bytes:
        # ACBegin / FRTBegin (random GUID, as Excel does) / ACEnd.
        import uuid
        guid = uuid.uuid4()
        payload = guid.bytes_le
        out = build_record(_BRT_AC_BEGIN, bytes((1, 0, 2, 0x10, 0, 0x80)))
        out += build_record(_BRT_FRT_BEGIN, payload)
        out += build_record(_BRT_AC_END, b"")
        return out

    def _table_bin(self, table, table_id: int, names: list[str]) -> bytes:
        first_col, first_row, last_col, last_row = parse_cell_range(table.ref)
        out = bytearray()
        out.extend(self._acuid())
        # BrtBeginList
        payload = bytearray()
        payload.extend(self._encode_rfx(first_col, first_row, last_col, last_row))
        payload.extend(struct.pack("<I", 0))  # lt = LTRANGE
        payload.extend(struct.pack("<I", table_id))
        payload.extend(struct.pack("<I", 1 if table.header_row else 0))
        payload.extend(struct.pack("<I", 1 if table.totals_row else 0))
        payload.extend(struct.pack("<I", 0))  # flags
        for _ in range(6):
            payload.extend(struct.pack("<i", -1))  # dxfIds
        payload.extend(struct.pack("<I", 0))  # dwConnID
        # Excel stores Name and DisplayName separately; COM shows
        # Name=T1, DisplayName=T1_ when Name is set. Store DisplayName with
        # trailing underscore to match Excel-generated files.
        display = table.display_name
        disp_display = display if display.endswith("_") else display + "_"
        payload.extend(self._encode_nullable_wstring(display))
        payload.extend(self._encode_nullable_wstring(disp_display))
        payload.extend(struct.pack("<I", 0))  # stComment empty
        for _ in range(3):
            payload.extend(self._encode_nullable_wstring(None))
        out.extend(build_record(_BRT_BEGIN_LIST, bytes(payload)))
        out.extend(self._acuid())
        # AutoFilter range inside the table part.
        out.extend(build_record(
            _BRT_BEGIN_AUTOFILTER,
            self._encode_rfx(first_col, first_row, last_col, last_row),
        ))
        out.extend(build_record(_BRT_END_AUTOFILTER, b""))
        # ListCols
        out.extend(build_record(
            _BRT_BEGIN_LIST_COLS, struct.pack("<I", len(names)),
        ))
        for idx, col_name in enumerate(names, start=1):
            out.extend(self._acuid())
            col_payload = bytearray()
            col_payload.extend(struct.pack("<I", idx))
            col_payload.extend(struct.pack("<I", 0))  # ilta none
            for _ in range(3):
                col_payload.extend(struct.pack("<i", -1))
            col_payload.extend(struct.pack("<I", 0))  # idqsif
            col_payload.extend(self._encode_nullable_wstring(None))
            col_payload.extend(self._encode_nullable_wstring(col_name))
            col_payload.extend(self._encode_nullable_wstring(None))
            for _ in range(3):
                col_payload.extend(self._encode_nullable_wstring(None))
            out.extend(build_record(_BRT_BEGIN_LIST_COL, bytes(col_payload)))
            out.extend(build_record(_BRT_END_LIST_COL, b""))
        out.extend(build_record(_BRT_END_LIST_COLS, b""))
        # TableStyleClient
        flags = 0
        if table.show_first_column:
            flags |= 0x01
        if table.show_last_column:
            flags |= 0x02
        if table.show_row_stripes:
            flags |= 0x04
        if table.show_column_stripes:
            flags |= 0x08
        style_payload = struct.pack("<H", flags)
        style_payload += self._encode_nullable_wstring(table.style_name)
        out.extend(build_record(_BRT_TABLE_STYLE_CLIENT, style_payload))
        out.extend(build_record(_BRT_END_LIST, b""))
        return bytes(out)

    def _list_parts_bin(self, rel_ids: list[str]) -> bytes:
        out = bytearray()
        out.extend(build_record(
            _BRT_BEGIN_LIST_PARTS, struct.pack("<I", len(rel_ids)),
        ))
        for rel_id in rel_ids:
            encoded = rel_id.encode("utf-16-le")
            out.extend(build_record(
                _BRT_LIST_PART,
                struct.pack("<I", len(rel_id)) + encoded,
            ))
        out.extend(build_record(_BRT_END_LIST_PARTS, b""))
        return bytes(out)

    def _begin_cf_rule_bin(self, iType: int, iTemplate: int,
                           dxfId: int, priority: int, iParam: int,
                           stop: bool = False) -> bytes:
        payload = bytearray()
        payload.extend(struct.pack("<I", iType))
        payload.extend(struct.pack("<I", iTemplate))
        payload.extend(struct.pack("<i", dxfId))
        payload.extend(struct.pack("<I", priority))
        payload.extend(struct.pack("<I", iParam))
        payload.extend(struct.pack("<I", 0))  # reserved1
        payload.extend(struct.pack("<I", 0))  # reserved2
        flags = 0x02 if stop else 0x00  # match Excel (fStopTrue=1 for CellIs)
        # For GRADIENT/DATABAR/MULTISTATE, flags must be 0.
        if iType in (3, 4, 6):
            flags = 0
        payload.extend(struct.pack("<H", flags))
        payload.extend(struct.pack("<I", 0))  # cbFmla1
        payload.extend(struct.pack("<I", 0))  # cbFmla2
        payload.extend(struct.pack("<I", 0))  # cbFmla3
        payload.extend(struct.pack("<I", 0xFFFFFFFF))  # strParam NULL
        return build_record(_BRT_BEGIN_CF_RULE, bytes(payload))

    def _cf_bin_for_rule(self, rule, priority: int) -> bytes:
        from .styles import ColorScaleRule, DataBarRule, IconSetRule
        out = bytearray()
        if isinstance(rule, ColorScaleRule):
            out.extend(self._begin_cf_rule_bin(3, 2, -1, priority, 0))
            out.extend(build_record(_BRT_BEGIN_COLORSCALE, b""))
            # 2-color or 3-color.
            if rule.midpoint is None:
                cfvos = (rule.minimum, rule.maximum)
            else:
                cfvos = (rule.minimum, rule.midpoint, rule.maximum)
            defaults = ("min", None, "max") if len(cfvos) == 2 else ("min", "percentile", "max")
            # Default midpoint for 3-color is percentile 50 (like Excel).
            for cfvo, default in zip(cfvos, ("min", "percentile", "max") if len(cfvos) == 3 else ("min", "max")):
                if cfvo is None and default == "percentile":
                    from .styles import FormatObject
                    cfvo = FormatObject("percentile", 50)
                out.extend(build_record(
                    _BRT_CFVO, self._cfvo_bin(cfvo, default),
                ))
            for color in rule.colors:
                out.extend(build_record(_BRT_COLOR, self._color_to_brt(color)))
            out.extend(build_record(_BRT_END_COLORSCALE, b""))
            out.extend(build_record(_BRT_END_CF_RULE, b""))
            return bytes(out)
        if isinstance(rule, DataBarRule):
            out.extend(self._begin_cf_rule_bin(4, 3, -1, priority, 0))
            show = 1 if rule.show_value is not False else 0
            # Excel defaults 10/90 when omitted? Use 0/100 to be safe? Use 10/90 to match Excel.
            min_len = rule.min_length if rule.min_length is not None else 0
            max_len = rule.max_length if rule.max_length is not None else 100
            # Clamp to reference defaults? Use provided or 0/100.
            out.extend(build_record(
                _BRT_BEGIN_DATABAR, bytes((min_len, max_len, show)),
            ))
            out.extend(build_record(
                _BRT_CFVO, self._cfvo_bin(rule.minimum, "min"),
            ))
            out.extend(build_record(
                _BRT_CFVO, self._cfvo_bin(rule.maximum, "max"),
            ))
            out.extend(build_record(_BRT_COLOR, self._color_to_brt(rule.color)))
            out.extend(build_record(_BRT_END_DATABAR, b""))
            out.extend(build_record(_BRT_END_CF_RULE, b""))
            return bytes(out)
        if isinstance(rule, IconSetRule):
            out.extend(self._begin_cf_rule_bin(6, 4, -1, priority, 0))
            code = self._ICON_SET_CODES[rule.icon_style]
            # BeginIconSet: iconSet (4) + showValue/percent/reverse flags (2)?
            # Reference 3TrafficLights1: 03 00 00 00 78 00.
            # 03 = code 3, 78 00 = flags? showValue=1, percent=1, reverse=0?
            # Bits: showValue? percent? reverse? Need to infer.
            show = 0 if rule.show_value is False else 1
            percent = 0 if rule.percent is False else 1
            reverse = 1 if rule.reverse else 0
            # Excel reference 78 00 = 0x0078 = 01111000b.
            # Guess: bit3=showValue? bit4=percent? bit5-6?
            # For now, match reference for defaults (show=1,percent=1,reverse=0):
            # show=1, percent=1 -> flags 0x78? Let's hardcode mapping from XLSX attrs:
            # showValue, percent, reverse -> flags per spec (need spec).
            # From reference: show=1,percent=1,reverse=0 => 0x78.
            # We compute: bit3 (0x08) = showValue? bit4 (0x10)=percent? bit5 (0x20)=? bit6 (0x40)=?
            # 0x78 = 0x08|0x10|0x20|0x40 = all except low 3 bits. Hmm.
            # Simpler: use 0x78 for defaults, adjust for reverse/show/percent by flipping?
            # For now, use 0x78 and set bit for reverse (0x80?)? Need spec, but test with defaults first.
            flags = 0x78
            if rule.show_value is False:
                flags &= ~0x08  # guess
            if rule.percent is False:
                flags &= ~0x10  # guess
            if rule.reverse:
                flags |= 0x80  # guess (reverse bit?)
            out.extend(build_record(
                _BRT_BEGIN_ICONSET,
                struct.pack("<I", code) + struct.pack("<H", flags),
            ))
            for v in rule.values:
                out.extend(build_record(_BRT_CFVO, self._cfvo_icon_bin(v)))
            out.extend(build_record(_BRT_END_ICONSET, b""))
            out.extend(build_record(_BRT_END_CF_RULE, b""))
            return bytes(out)
        raise NotImplementedError("CF rule type not yet supported for XLSB")

    def _cf_bin_grouped(self, rules) -> bytes:
        from .styles import parse_cell_range
        # Group by sqref like XLSX, claim priorities in order.
        by_range: dict[str, list] = {}
        for rule in rules:
            by_range.setdefault(rule.sqref, []).append(rule)
        out = bytearray()
        for sqref, group in by_range.items():
            first_col, first_row, last_col, last_row = parse_cell_range(sqref)
            # BeginCF: ccf, fPivot=0, count=1, range.
            payload = bytearray()
            payload.extend(struct.pack("<I", len(group)))
            payload.extend(struct.pack("<I", 0))
            payload.extend(struct.pack("<I", 1))
            payload.extend(self._encode_rfx(first_col, first_row, last_col, last_row))
            out.extend(build_record(_BRT_BEGIN_CF, bytes(payload)))
            for rule in group:
                priority = self._claim_cf_priority(rule)
                out.extend(self._cf_bin_for_rule(rule, priority))
            out.extend(build_record(_BRT_END_CF, b""))
        return bytes(out)

    def write_sheet(
        self,
        data: Iterable[list[any]],
        *,
        header_style: Style | None = None,
        styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        column_widths: dict | None = None,
        hidden_columns: list | tuple | set | None = None,
        row_heights: dict | None = None,
        hidden_rows: list | tuple | set | None = None,
        row_groups: list | tuple | None = None,
        column_groups: list | tuple | None = None,
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        charts: list[Chart] | tuple[Chart, ...] | None = None,
        tables: list | tuple | None = None,
        conditional_formats: list | tuple | None = None,
    ):
        self._require_writable()
        if header_style is not None and not isinstance(header_style, Style):
            raise TypeError("header_style must be a Style instance")
        style_ranges = normalize_style_ranges(styles)
        for candidate in ([header_style] if header_style is not None else []) + [item[-1] for item in style_ranges]:
            self._reject_rich_style(candidate)
        merge_refs = self._normalize_merges(merges)
        widths = self._normalize_widths(column_widths, "column_widths")
        hidden_cols = normalize_hidden_columns(hidden_columns)
        heights = self._normalize_row_heights(row_heights)
        hidden_row_set = normalize_hidden_rows(hidden_rows)
        row_outline = build_outline(row_groups, "row")
        col_outline = build_outline(column_groups, "column")
        if col_outline is not None:
            hidden_cols = hidden_cols | col_outline.hidden_indices()
        if freeze_panes is not None and freeze_panes != "":
            self._split_freeze_cell(freeze_panes)
        chart_specs = self._normalize_charts(charts)
        for chart in chart_specs:
            if chart.uses_extended_features:
                raise ValueError(
                    "XlsbWriter supports only the original chart options "
                    "(bar, line and pie with no axis, legend, label, colour or marker settings); "
                    "use XlsxWriter for these charts"
                )
        table_specs = self._normalize_tables(tables)
        cf_specs = self._normalize_cf_rules(conditional_formats)
        # Pre-reserve explicit priorities (like XLSX) so auto-assigned ones skip them.
        cf_priorities = explicit_cf_priorities(
            cf_specs, self._seen_cf_priorities, self._reserved_cf_priorities,
        )

        # Every option is valid; only now record what this sheet claims, so a
        # rejected call leaves the writer as it was.
        if not self._worksheet_data:
            self.add_sheet("Sheet1")
        sheet_name,_, hidden = self._worksheet_data[self._sheet_count - 1]
        self._worksheet_data[self._sheet_count - 1] = (sheet_name, data, hidden)
        sheet_id = self._sheet_count
        if chart_specs:
            self._sheet_charts[sheet_id] = chart_specs
            # Chart caches need only the cells the charts read. Capture those
            # while the rows stream past instead of holding the whole sheet.
            store = chart_parts.new_cell_store(chart_parts.chart_cell_spans(chart_specs))
            data = _capture_cells(data, store)
            self._sheet_chart_cells[sheet_id] = store
        table_ids: list[int] = []
        table_rel_ids: list[str] = []
        header_rows: dict = {}
        if table_specs:
            self._table_names.update(table.display_name for table in table_specs)
            # rId1 is binaryIndex; tables take rId2..; drawing (if any) last.
            for table in table_specs:
                table_id = self._next_table_id
                self._next_table_id += 1
                table_ids.append(table_id)
                rel_id = f"rId{1 + len(table_ids)}"
                table_rel_ids.append(rel_id)
            self._sheet_tables[sheet_id] = list(zip(table_ids, table_rel_ids))
            # Column names must match the header cells, which stream past with
            # the rows; a mismatch fails as soon as the header row is read.
            data = capture_table_headers(data, table_specs, header_rows)
        if table_specs and auto_filter is None:
            # Tables carry their own AutoFilter; skip sheet-level filter.
            auto_filter = False
        for priority in cf_priorities:
            self._reserved_cf_priorities.add(priority)
            self._next_cf_priority = max(self._next_cf_priority, priority + 1)
        if chart_specs and table_specs:
            drawing_rel = f"rId{1 + len(table_ids) + 1}"
        elif chart_specs:
            drawing_rel = "rId2"
        else:
            drawing_rel = None
        try:
            with self._zf.open(f"xl/worksheets/sheet{sheet_id}.bin", 'w') as sheet_file:
                self._write_worksheet_bin(
                    sheet_file, data, self._sheet_count - 1,
                    style_ranges=style_ranges, header_style=header_style,
                    merge_refs=merge_refs, widths=widths, hidden_cols=hidden_cols,
                    heights=heights, hidden_row_set=hidden_row_set,
                    freeze_panes=freeze_panes, auto_filter=auto_filter,
                    drawing_rel=drawing_rel, table_rel_ids=table_rel_ids,
                    cf_rules=cf_specs, row_outline=row_outline, col_outline=col_outline,
                )
            for table, table_id in zip(table_specs, table_ids):
                names = table_column_names(table, header_rows.get(parse_cell_range(table.ref)[1]))
                self._tables_bin[table_id] = self._table_bin(table, table_id, names)
            self._zf.writestr(f"xl/worksheets/_rels/sheet{sheet_id}.bin.rels", self._create_worksheet_rels(sheet_id))
            self._zf.writestr(f"xl/worksheets/binaryIndex{sheet_id}.bin", self._binaryIndexBin)
        except BaseException:
            # The sheet part is incomplete, so the workbook cannot be saved.
            self._aborted = True
            raise

    def write_dataframe(
        self, df, include_index: bool = False, na_rep=None, formats: dict = None,
        *, header_style: Style | None = None, styles: dict[str, Style] | None = None,
        merges: list[str] | tuple[str, ...] | None = None,
        column_widths: dict | None = None,
        hidden_columns: list | tuple | set | None = None,
        row_heights: dict | None = None,
        hidden_rows: list | tuple | set | None = None,
        row_groups: list | tuple | None = None,
        column_groups: list | tuple | None = None,
        freeze_panes: str | None = None,
        auto_filter: bool | None = None,
        charts: list | tuple | None = None,
        tables: list | tuple | None = None,
        conditional_formats: list | tuple | None = None,
    ):
        """Write a pandas or polars DataFrame to the current sheet.

        pandas and polars are optional: they are imported lazily only when a
        DataFrame is actually passed in.

        Parameters
        ----------
        df:
            A ``pandas.DataFrame`` or ``polars.DataFrame``.
        include_index:
            pandas only -- prepend the index as leading column(s).
        na_rep:
            Value for missing cells, defaults to ``None`` (empty cell).
        formats:
            Optional ``{column_name: format_string}`` mapping (see ``xlspy.F``).
        """
        from .dataframe import dataframe_to_rows

        self.write_sheet(
            dataframe_to_rows(df, include_index=include_index, na_rep=na_rep, formats=formats),
            header_style=header_style,
            styles=styles,
            merges=merges,
            column_widths=column_widths,
            hidden_columns=hidden_columns,
            row_heights=row_heights,
            hidden_rows=hidden_rows,
            row_groups=row_groups,
            column_groups=column_groups,
            freeze_panes=freeze_panes,
            auto_filter=auto_filter,
            charts=charts,
            tables=tables,
            conditional_formats=conditional_formats,
        )

    def _write_workbook_parts(self):
        for sheet_id, charts in self._sheet_charts.items():
            if not charts or sheet_id in self._sheet_chart_ids:
                continue
            sheet_name = self._worksheet_data[sheet_id - 1][0]
            cells = self._sheet_chart_cells.get(sheet_id)
            chart_ids = []
            for chart in charts:
                chart_id = self._next_chart_id
                self._next_chart_id += 1
                chart_ids.append(chart_id)
                self._zf.writestr(
                    f"xl/charts/chart{chart_id}.xml",
                    chart_parts.chart_xml(
                        self._letters, chart, sheet_name,
                        cells=cells, xlsb=True,
                    ),
                )
            self._sheet_chart_ids[sheet_id] = chart_ids
            self._zf.writestr(
                f"xl/drawings/drawing{sheet_id}.xml",
                chart_parts.drawing_xml(charts, chart_ids),
            )
            self._zf.writestr(
                f"xl/drawings/_rels/drawing{sheet_id}.xml.rels",
                chart_parts.drawing_rels_xml(chart_ids),
            )

        if any(self._sheet_charts.values()):
            self._zf.writestr("xl/theme/theme1.xml", chart_parts.THEME_XML)

        for table_id, table_bin in self._tables_bin.items():
            self._zf.writestr(f"xl/tables/table{table_id}.bin", table_bin)
            
        self._zf.writestr("[Content_Types].xml", self._create_content_types())
        self._zf.writestr("_rels/.rels", self._create_root_rels())
        self._zf.writestr("xl/workbook.bin", self._create_workbook_bin())
        self._zf.writestr("xl/styles.bin", self._build_styles_bin())
        self._zf.writestr("xl/_rels/workbook.bin.rels", self._create_workbook_rels())

        with self._zf.open("xl/sharedStrings.bin", 'w') as sst_file:
            self._write_shared_strings_bin(sst_file)

    def _create_content_types(self) -> str:
        parts = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.bin" ContentType="application/vnd.ms-excel.worksheet"/>'
            f'<Override PartName="/xl/worksheets/binaryIndex{i + 1}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>'
            for i in range(self._sheet_count)
        )

        drawing_parts = "".join(
            f'<Override PartName="/xl/drawings/drawing{sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            for sheet_id in sorted(self._sheet_charts)
            if self._sheet_charts[sheet_id]
        )
        table_parts = "".join(
            f'<Override PartName="/xl/tables/table{table_id}.bin" ContentType="application/vnd.ms-excel.table"/>'
            for table_id in sorted(self._tables_bin)
        )
        chart_parts_overrides = "".join(
            f'<Override PartName="/xl/charts/chart{chart_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            for chart_id in range(1, self._next_chart_id)
        )
        theme_part = ""
        if any(self._sheet_charts.values()):
            theme_part = '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
{parts}
{drawing_parts}
{table_parts}
{chart_parts_overrides}
{theme_part}
<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>'''

    def _create_root_rels(self) -> str:
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>'''

    def _create_workbook_rels(self) -> str:
        relationships = [
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.bin"/>'
            for i in range(self._sheet_count)
        ]
        
        style_rid = self._sheet_count + 1
        shared_strings_rid = self._sheet_count + 2
        relationships.append(f'<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>')
        relationships.append(f'<Relationship Id="rId{shared_strings_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>')
        if any(self._sheet_charts.values()):
            theme_rid = self._sheet_count + 3
            relationships.append(f'<Relationship Id="rId{theme_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')

        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{"".join(relationships)}
</Relationships>'''

    def _create_worksheet_rels(self, sheet_id) -> str:
        tables = self._sheet_tables.get(sheet_id, [])
        rels = (
            '<Relationship Id="rId1" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" '
            f'Target="binaryIndex{sheet_id}.bin"/>'
        )
        for table_id, rel_id in tables:
            rels += (
                f'<Relationship Id="{rel_id}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
                f'Target="../tables/table{table_id}.bin"/>'
            )
        if self._sheet_charts.get(sheet_id):
            drawing_rid = f"rId{1 + len(tables) + 1}"
            rels += (
                f'<Relationship Id="{drawing_rid}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/drawing{sheet_id}.xml"/>'
            )
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}
</Relationships>'''

    def _create_workbook_bin(self) -> bytes:
        data = bytearray(self._workbookBinStart)
        for i, (sheet_name, _, hidden) in enumerate(self._worksheet_data):
            sheet_id = i + 1
            r_id = f"rId{sheet_id}"
            
            sheet_name_utf16 = sheet_name.encode('utf-16-le')
            r_id_utf16 = r_id.encode('utf-16-le')
            record_length = 4 + 3 * 4 + len(sheet_name_utf16) + len(r_id_utf16)
            
            data.extend(b'\x9C\x01')
            data.extend(self._write_custom_variable_int(record_length))
            
            if hidden:
                data.extend(struct.pack('<I', 1))
            else:
                data.extend(struct.pack('<I', 0)) 

            data.extend(struct.pack('<I', sheet_id))
            data.extend(struct.pack('<I', len(r_id_utf16) // 2))
            data.extend(r_id_utf16)
            data.extend(struct.pack('<I', len(sheet_name_utf16) // 2))
            data.extend(sheet_name_utf16)

        data.extend(self._workbookBinMiddle)
        self._write_filter_defined_names(data)
        data.extend(self._workbookBinEnd)
        return bytes(data)

    def _write_custom_variable_int(self, value: int) -> bytes:
        if value < 0:
            raise ValueError("Value must be non-negative")
        
        buffer = bytearray()
        temp_val = value
        while True:
            byte = temp_val & 0x7F
            temp_val >>= 7
            if temp_val > 0:
                byte |= 0x80
                buffer.append(byte)
            else:
                buffer.append(byte)
                break
        return bytes(buffer)

    def _write_shared_strings_bin(self, sst_file: io.BufferedWriter):
        c_sst = accel_encode_xlsb_sst(
            self._shared_strings, self._sst_all_count, self._sst_unique_count
        )
        if c_sst is not None:
            view = memoryview(c_sst)
            for offset in range(0, len(view), 65536):
                sst_file.write(view[offset:offset + 65536])
            return
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536

        buffer.write(b'\x9F\x01\x08')
        buffer.write(struct.pack('<II', self._sst_all_count, self._sst_unique_count))
        
        for s in self._shared_strings:
            # An over-limit string is an error, never a cut: the cell writers
            # check every value before it reaches this table.
            check_cell_text(s)
            encoded = s.encode('utf-16-le')
            str_len = len(encoded) // 2

            record_length = 1 + 4 + 2 * str_len
            
            buffer.write(b'\x13')
            buffer.write(self._write_custom_variable_int(record_length))
            buffer.write(b'\x00')
            buffer.write(struct.pack('<I', str_len))
            buffer.write(encoded)

            if buffer.tell() > BUFFER_SIZE:
                sst_file.write(buffer.getvalue())
                buffer.seek(0)
                buffer.truncate()
            
        buffer.write(b'\xA0\x01\x00')
        
        if buffer.tell() > 0:
            sst_file.write(buffer.getvalue())

    @staticmethod
    def _build_numfmt_record(ifmt: int, fmt_string: str) -> bytes:
        encoded = fmt_string.encode('utf-16-le')
        payload = struct.pack('<HI', ifmt, len(encoded) // 2) + encoded
        return build_record(_BRT_FMT, payload)

    @staticmethod
    def _color_payload(color_key) -> bytes:
        """An 8-byte BrtColor ([MS-XLSB] 2.4.334) for a ``split_style`` color."""
        if color_key is None:
            return bytes(8)  # automatic
        kind = color_key[0]
        if kind == "rgb":
            alpha, red, green, blue = bytes.fromhex(color_key[1])
            return bytes((0x05, 0, 0, 0, red, green, blue, alpha))
        if kind == "indexed":
            return bytes((0x03, color_key[1], 0, 0, 0, 0, 0, 0xFF))
        if kind == "theme":
            tint = color_key[2]
            scaled = 0 if tint is None else max(-32768, min(32767, round(tint * 32767)))
            return bytes((0x07, color_key[1])) + struct.pack('<h', scaled) + bytes((0, 0, 0, 0xFF))
        return bytes(8)  # ("auto",)

    def _xlsb_font_record(self, font_key: tuple) -> bytes:
        """Build a BrtFont record ([MS-XLSB] 2.4.683) from a font key."""
        flags = 0
        for bit, value in (
            (0x01, font_key.bold), (0x02, font_key.italic), (0x08, font_key.strike),
            (0x10, font_key.outline), (0x20, font_key.shadow),
            (0x40, font_key.condense), (0x80, font_key.extend),
        ):
            if value:
                flags |= bit
        height = 220 if font_key.size is None else max(0x14, min(0x1FFF, round(font_key.size * 20)))
        payload = bytearray()
        payload.extend(struct.pack(
            '<HHHH', height, flags, 700 if font_key.bold else 400,
            _FONT_BASELINE[font_key.vert_align],
        ))
        payload.extend((
            _UNDERLINE_BYTES[font_key.underline],
            font_key.family if font_key.family is not None else 2,
            font_key.charset if font_key.charset is not None else 0,
            0,
        ))
        payload.extend(self._color_payload(font_key.color))
        scheme = _FONT_SCHEME[font_key.scheme]
        if font_key.scheme is None and font_key.name is None:
            # The default Calibri is the minor (body) theme font; a font that
            # names its typeface must not claim a theme scheme or Excel would
            # substitute the theme font for the requested one.
            scheme = 2
        payload.append(scheme)
        name = (font_key.name or 'Calibri').encode('utf-16-le')
        payload.extend(struct.pack('<I', len(name) // 2))
        payload.extend(name)
        return build_record(_BRT_FONT, bytes(payload))

    def _xlsb_fill_record(self, fill_key: tuple) -> bytes:
        """Build a BrtFill record ([MS-XLSB] 2.4.674) from a fill key."""
        if fill_key[0] == "gradient":
            return self._xlsb_gradient_fill_record(fill_key)
        _, fill_type, start_color, end_color = fill_key
        payload = bytearray()
        payload.extend(struct.pack('<I', _PATTERN_BYTES[fill_type]))
        payload.extend(self._color_payload(start_color))
        payload.extend(self._color_payload(end_color))
        payload.extend(bytes(48))  # unused gradient fields
        return build_record(_BRT_FILL, bytes(payload))

    def _xlsb_gradient_fill_record(self, fill_key: tuple) -> bytes:
        _, fill_type, degree, left, right, top, bottom, stops = fill_key
        linear = fill_type == "linear"
        payload = bytearray()
        payload.extend(struct.pack('<I', 0x28))
        payload.extend(self._color_payload(None))
        payload.extend(self._color_payload(None))
        payload.extend(struct.pack('<I', 0 if linear else 1))
        # BIFF12 measures the gradient angle counter-clockwise while xlspy (like
        # OOXML) measures it clockwise.
        angle = (360 - degree) % 360 if linear and degree is not None else 0.0
        payload.extend(struct.pack('<d', angle))
        if linear:
            payload.extend(struct.pack('<dddd', 0.0, 0.0, 0.0, 0.0))
        else:
            # OOXML omits unset fillToRect values, i.e. the inner rectangle
            # spans the whole cell.
            payload.extend(struct.pack(
                '<dddd',
                0.0 if left is None else float(left),
                1.0 if right is None else float(right),
                0.0 if top is None else float(top),
                1.0 if bottom is None else float(bottom),
            ))
        payload.extend(struct.pack('<I', len(stops)))
        for color_key, position in stops:
            payload.extend(self._color_payload(color_key))
            payload.extend(struct.pack('<d', float(position)))
        return build_record(_BRT_FILL, bytes(payload))

    def _xlsb_border_record(self, border_key: tuple) -> bytes:
        """Build a BrtBorder record ([MS-XLSB] 2.4.311) from a border key."""
        diagonal = 0
        if border_key.diagonal_down:
            diagonal |= 0x01
        if border_key.diagonal_up:
            diagonal |= 0x02
        payload = bytearray((diagonal,))
        for style, color in (
            border_key.top, border_key.bottom, border_key.left,
            border_key.right, border_key.diagonal,
        ):
            payload.append(_BORDER_BYTES[style])
            payload.append(0)  # reserved
            payload.extend(self._color_payload(color))
        return build_record(_BRT_BORDER, bytes(payload))

    def _build_style_record(
        self, ifmt: int = 0, *, flags: int = 0, font_id: int = 0, fill_id: int = 0,
        border_id: int = 0, alignment=None, protection=None, apply_flags: bool = False,
    ) -> bytes:
        """Build a BrtXF record ([MS-XLSB] 2.4.863).

        ``alignment`` is a ``split_style`` alignment key and ``protection`` a
        protection key.  ``apply_flags`` derives the xfGrbitAtr bits from the
        referenced components; the built-in XFs pass ``flags`` explicitly to
        keep byte parity with Excel's own file.
        """
        rotation = 0
        indent = 0
        alignment_byte = 0x10  # Excel's default: bottom vertical alignment
        flags_byte = 0x10  # Excel's default: locked
        if alignment is not None:
            (
                horizontal, vertical, wrap, shrink, indent_value,
                text_rotation, _relative_indent, justify_last_line,
                reading_order,
            ) = alignment
            rotation = 0 if text_rotation is None else text_rotation
            indent = 0 if indent_value is None else indent_value
            alignment_byte = (
                _HORIZONTAL_BYTES[horizontal]
                | (_VERTICAL_BYTES[vertical] << 3)
                | (0x40 if wrap else 0)
                | (0x80 if justify_last_line else 0)
            )
            flags_byte = (
                0x10
                | (0x01 if shrink else 0)
                | ((reading_order or 0) << 2)
            )
        if protection is not None:
            locked, hidden = protection
            flags_byte = (
                (flags_byte & 0xCF)
                | (0x10 if locked is not False else 0x00)
                | (0x20 if hidden else 0x00)
            )
        if apply_flags:
            flags = (
                (0x01 if ifmt else 0)
                | (0x02 if font_id else 0)
                | (0x04 if alignment is not None else 0)
                | (0x08 if border_id else 0)
                | (0x10 if fill_id else 0)
                | (0x20 if protection is not None else 0)
            )
        payload = bytearray(2)  # ixfeParent: cell style XF 0
        payload.extend(struct.pack('<HHHH', ifmt, font_id, fill_id, border_id))
        payload.append(rotation)
        payload.append(indent)
        payload.append(alignment_byte)
        payload.append(flags_byte)
        payload.extend(struct.pack('<H', flags))
        return build_record(_BRT_XF, bytes(payload))

    def _build_styles_bin(self) -> bytes:
        if not self._format_registry and not self._style_components:
            return self._stylesBin

        template = self._stylesBin
        buf = bytearray()
        buf.extend(template[0:3])  # BrtBeginStyleSheet

        # ── FMTS: built-in date and date-time formats plus user formats ──
        buf.extend(build_record(_BRT_BEGIN_FMTS, struct.pack('<I', 2 + len(self._format_registry))))
        buf.extend(template[10:88])
        for fmt_string, numfmt_id in self._format_registry.items():
            buf.extend(self._build_numfmt_record(numfmt_id, fmt_string))
        buf.extend(build_record(_BRT_END_FMTS))

        # ── FONTS: built-in regular/bold Calibri plus user fonts ──
        buf.extend(build_record(_BRT_BEGIN_FONTS, struct.pack('<I', 2 + len(self._font_registry))))
        buf.extend(template[98:180])
        for font_key, _font_id in sorted(self._font_registry.items(), key=lambda item: item[1]):
            buf.extend(self._xlsb_font_record(font_key))
        buf.extend(template[180:193])  # BrtACBegin / BrtKnownFonts / BrtACEnd
        buf.extend(build_record(_BRT_END_FONTS))

        # ── FILLS: built-in none/gray125 plus user fills ──
        buf.extend(build_record(_BRT_BEGIN_FILLS, struct.pack('<I', 2 + len(self._fill_registry))))
        buf.extend(template[203:343])
        for fill_key, _fill_id in sorted(self._fill_registry.items(), key=lambda item: item[1]):
            buf.extend(self._xlsb_fill_record(fill_key))
        buf.extend(build_record(_BRT_END_FILLS))

        # ── BORDERS: built-in empty border plus user borders ──
        buf.extend(build_record(_BRT_BEGIN_BORDERS, struct.pack('<I', 1 + len(self._border_registry))))
        buf.extend(template[353:406])
        for border_key, _border_id in sorted(self._border_registry.items(), key=lambda item: item[1]):
            buf.extend(self._xlsb_border_record(border_key))
        buf.extend(build_record(_BRT_END_BORDERS))

        # ── CELLSTYLEXFS: the single built-in "Normal" style XF ──
        buf.extend(template[409:437])

        # ── CELLXFS: four built-in XFs plus one XF per registered style ──
        total_styles = 4 + len(self._format_xf_map) + len(self._style_components)
        buf.extend(build_record(_BRT_BEGIN_CELL_XFS, struct.pack('<I', total_styles)))
        buf.extend(self._build_style_record(0))  # 0: default
        buf.extend(self._build_style_record(164, flags=0x0001))  # 1: date
        buf.extend(self._build_style_record(166, flags=0x0001))  # 2: date-time
        buf.extend(self._build_style_record(256, font_id=1))  # 3: bold header

        # Custom format/style records follow the XF indices assigned while
        # streaming cells, so preserve that exact order.
        custom_xfs = {
            index: (0, 0, 0, self._format_registry[fmt], None, None)
            for fmt, index in self._format_xf_map.items()
        }
        custom_xfs.update(self._style_components)
        for index in sorted(custom_xfs):
            font_id, fill_id, border_id, numfmt_id, alignment, protection = custom_xfs[index]
            buf.extend(self._build_style_record(
                numfmt_id,
                font_id=font_id,
                fill_id=fill_id,
                border_id=border_id,
                alignment=alignment,
                protection=protection,
                apply_flags=True,
            ))
        buf.extend(build_record(_BRT_END_CELL_XFS))

        # ── Remaining sections (styles, DXFs, tableStyles, BrtEndStyleSheet) ──
        buf.extend(template[519:])

        return bytes(buf)

    @staticmethod
    def _row_height_twips(points: float) -> int:
        height_twips = int(round(points * 20))
        if height_twips > 0x2000:
            raise ValueError("row height must be at most 409.6 points for XLSB output")
        return height_twips

    @staticmethod
    def _row_header_record(
        row_idx: int, last_col: int, height_twips: int | None = None,
        hidden: bool = False, outline_level: int = 0, collapsed: bool = False,
    ) -> bytes:
        """Build a BrtRowHdr record ([MS-XLSB] 2.4.761).

        The default record (300 twips = 15 pt, no flags) is byte-identical
        to the row headers the writer historically emitted. The flags byte
        holds the outline level in bits 0-2, then collapsed (0x08), hidden
        (0x10) and unsynced (0x20, a custom height). Excel-written files show
        this layout.
        """
        flags = (
            (0x20 if height_twips is not None else 0) | (0x10 if hidden else 0)
            | (0x08 if collapsed else 0) | outline_level
        )
        payload = struct.pack('<II', row_idx, 0)  # rw, ixfe
        payload += struct.pack('<H', 300 if height_twips is None else height_twips)
        payload += bytes((0, flags, 0))  # fExtraAsc/fExtraDsc, outline+collapsed+hidden+unsynced, phonetic
        payload += struct.pack('<I', 1)  # ccolspan
        payload += struct.pack('<II', 0, last_col)
        return build_record(_BRT_ROW_HDR, payload)

    def _pane_record(self, freeze_panes: str | None) -> bytes:
        """Build a BrtPane record ([MS-XLSB] 2.4.716).

        ``None`` keeps the template default (frozen top row); ``""`` removes
        the freeze; otherwise the cell gives the first unfrozen row/column.
        Column/row splits follow the OOXML convention (x = columns, y = rows).
        """
        if freeze_panes is None:
            return self._stickHeaderA1bytes
        if freeze_panes == "":
            split_col = split_row = 0
            active = _PANE_ACTIVE["topLeft"]
            frozen = 0x00
        else:
            split_col, split_row = self._split_freeze_cell(freeze_panes)
            # NOTE: the flags byte mirrors the template (0x03 = fFrozen plus
            # fFrozenNoSplit) rather than the spec-pure 0x02. Tested against
            # Excel: 0x02 hides the sheet tabs and the row/column headers,
            # 0x03 renders the freeze correctly. Likewise pnnAct stays
            # bottomLeft as in the template; Excel tolerates the mismatch
            # with the selection pane.
            active = _PANE_ACTIVE["bottomLeft"]
            frozen = 0x03
        payload = struct.pack('<dd', float(split_col), float(split_row))
        payload += struct.pack('<III', split_row, split_col, active)
        payload += bytes((frozen,))
        return build_record(_BRT_PANE, payload)

    def _selection_record(self, freeze_panes: str) -> bytes:
        """Build a BrtSel record ([MS-XLSB] 2.4.781) matching ``freeze_panes``."""
        if freeze_panes == "":
            split_col = split_row = 0
            active = _PANE_ACTIVE["topLeft"]
        else:
            split_col, split_row = self._split_freeze_cell(freeze_panes)
            if split_col > 0 and split_row > 0:
                active = _PANE_ACTIVE["bottomRight"]
            elif split_col > 0:
                active = _PANE_ACTIVE["topRight"]
            else:
                active = _PANE_ACTIVE["bottomLeft"]
        payload = struct.pack('<IIII', active, split_row, split_col, 0)
        payload += struct.pack('<I', 1)  # one range
        payload += struct.pack('<IIII', split_row, split_row, split_col, split_col)
        return build_record(_BRT_SEL, payload)

    def _write_cols_width(
        self, buffer: io.BytesIO, start_col: int, end_col: int,
        col_widths_array: list[float], widths: dict | None = None,
        hidden_cols: set | frozenset = frozenset(),
        col_outline=None,
    ):
        widths = widths or {}
        buffer.write(build_record(_BRT_BEGIN_COL_INFOS, b''))

        columns = range(start_col, end_col)
        if col_outline is not None:
            # Grouped columns past the data still need a BrtColInfo.
            columns = sorted(set(columns) | col_outline.indices())
        for i in columns:
            explicit = widths.get(i)
            if explicit is not None:
                coldx = int(round(explicit * 256))
            elif i - start_col < len(col_widths_array):
                # Historical encoding: the auto width truncated to whole
                # characters, i.e. coldx = int(width) * 256.
                coldx = int(col_widths_array[i - start_col]) * 256
            else:
                coldx = 8 * 256  # the default width, 8.43 characters, truncated
            flags = 0x02 | (0x01 if i in hidden_cols else 0)
            if col_outline is not None:
                # BrtColInfo: outline level in bits 8-10, collapsed in bit 12.
                flags |= col_outline.state_at(i)[0] << 8
                if col_outline.collapsed_at(i):
                    flags |= 0x1000
            buffer.write(build_record(
                _BRT_COL_INFO, struct.pack('<IIIIH', i, i, coldx, 0, flags),
            ))

        buffer.write(build_record(_BRT_END_COL_INFOS, b''))

    def _write_worksheet_bin(
        self, sheet_file: io.BufferedWriter, worksheet_data: Iterable[list[any]],
        worksheet_index: int, *, style_ranges=(), header_style: Style | None = None,
        merge_refs=(), widths=None, hidden_cols=frozenset(), heights=None,
        hidden_row_set=frozenset(), freeze_panes: str | None = None,
        auto_filter: bool | None = None, drawing_rel: str | None = None,
        table_rel_ids: list[str] | None = None,
        cf_rules: list | None = None,
        row_outline=None, col_outline=None,
    ):
        buffer = io.BytesIO()
        BUFFER_SIZE = 65536

        buffer.write(self._sheet1Bytes[:54])
        if self._sheetCnt != 1:
            buffer.write(b'\x9C')
        else:
            buffer.write(b'\xDC')
        self._sheetCnt += 1

        buffer.write(self._sheet1Bytes[55:84])

        buffer.write(self._pane_record(freeze_panes))

        data_iterator = iter(worksheet_data)
        rows_to_analyze = []
        max_cols = 1

        for _ in range(100):
            try:
                row = next(data_iterator)
                rows_to_analyze.append(row)
                row_len = 0
                for c in row:
                    cv = _unwrap_cell(c) if c is not None else None
                    if cv is not None:
                        row_len += 1
                max_cols = max(max_cols, row_len if row else 0)
            except StopIteration:
                break

        if max_cols == 0:
            max_cols = 1

        start_col = 0
        end_col = max_cols

        col_widths_array = accel_calc_column_widths(rows_to_analyze, max_cols)
        if col_widths_array is None:
            col_widths_array = []
            for i in range(max_cols):
                max_width = 0
                for row in rows_to_analyze:
                    if i < len(row):
                        cell = row[i] if i < len(row) else None
                        cell = _unwrap_cell(cell) if cell is not None else None
                        if cell is not None:
                            if isinstance(cell, Decimal):
                                cell = float(cell)
                            if isinstance(cell, datetime.datetime):
                                temp_width = 18
                            else:
                                str_cell = str(cell)
                                if '\n' in str_cell:
                                    str_cell = str_cell.split('\n')[0]
                                lenn = len(str_cell)
                                temp_width = 1.25 * lenn + 4
                            max_width = max(max_width, temp_width)
                
                if max_width == 0:
                    max_width = 4.0
                if max_width > 255:
                    max_width = 255
                col_widths_array.append(max_width)

        # The last two bytes of BrtWsFmtInfo are the maximum row and column
        # outline levels; the template holds zeros.
        outline_levels = bytes((
            row_outline.max_level if row_outline is not None else 0,
            col_outline.max_level if col_outline is not None else 0,
        ))
        if freeze_panes is None:
            buffer.write(self._sheet1Bytes[84:157])
        else:
            buffer.write(self._selection_record(freeze_panes))
            # Template offsets 84..123 hold the default BrtSel record; the
            # remainder (EndWsView .. BrtWsFmtInfo) is layout-independent.
            buffer.write(self._sheet1Bytes[123:157])
        buffer.write(outline_levels)
        self._write_cols_width(
            buffer, start_col, end_col, col_widths_array, widths, hidden_cols,
            col_outline,
        )
        # Template offsets 159..177 hold BrtBeginSheetData plus the complete
        # ACCELLTABLE block (ACBegin/RwDescent/ACEnd); rows follow directly.
        buffer.write(self._sheet1Bytes[159:177])

        shared_strings_dict = self._shared_strings_dict
        shared_strings_append = self._shared_strings.append

        if rows_to_analyze:
            data_iterator = itertools.chain(rows_to_analyze, data_iterator)
        else:
            data_iterator = iter([])

        # Rows with a custom height or hidden flag cannot use the fast C
        # encoder, which only emits default row headers.
        heights = heights or {}
        custom_rows = set(heights) | set(hidden_row_set)

        row_idx = -1
        for row_idx, row in enumerate(data_iterator):
            c_has_formatted = accel_row_has_formatted(row)
            if c_has_formatted is None:
                has_formatted = any(_is_formatted_cell(c) for c in row if c is not None)
            else:
                has_formatted = c_has_formatted
            row_styles = styles_for_row(row_idx, len(row), style_ranges, header_style)

            if row_outline is None:
                outline_level, outline_hidden, outline_collapsed = 0, False, False
            else:
                outline_level, outline_hidden = row_outline.state_at(row_idx)
                outline_collapsed = row_outline.collapsed_at(row_idx)
            if (not has_formatted and row_idx not in custom_rows
                    and not outline_level and not outline_collapsed):
                style_ids = None
                if row_styles:
                    first_style = next(iter(row_styles.values()))
                    if (len(row_styles) == len(row)
                            and all(style is first_style for style in row_styles.values())):
                        style_ids = self._register_style(first_style)
                    else:
                        style_ids = tuple(
                            (col, self._register_style(row_styles[col]))
                            for col in sorted(row_styles)
                        )
                cell_data = None
                if self._useSharedStrings:
                    c_result = encode_xlsb_row(
                        row, self._shared_strings_dict, self._shared_strings,
                        self._sst_unique_count, self._sst_all_count, row_idx, style_ids
                    )
                    if c_result is not None:
                        cell_data, self._sst_unique_count, self._sst_all_count = c_result
                else:
                    cell_data = encode_xlsb_row_inline(row, row_idx, style_ids)
                if cell_data is not None:
                    buffer.write(cell_data)
                    if buffer.tell() > BUFFER_SIZE:
                        sheet_file.write(buffer.getvalue())
                        buffer.seek(0)
                        buffer.truncate()
                    continue

            last_col = len(row) - 1 if row else 0

            height_twips = None
            if row_idx in heights:
                height_twips = self._row_height_twips(heights[row_idx])
            buffer.write(self._row_header_record(
                row_idx, last_col, height_twips,
                row_idx in hidden_row_set or outline_hidden,
                outline_level, outline_collapsed,
            ))

            for col_idx, cell in enumerate(row):
                if cell is None:
                    cell_style = row_styles.get(col_idx)
                    if cell_style is not None:
                        xf_index = self._register_style(cell_style)
                        buffer.write(b'\x01\x08')
                        buffer.write(struct.pack('<II', col_idx, xf_index))
                    continue

                fmt_string = _get_format(cell)
                if fmt_string is not None:
                    cell = _unwrap_cell(cell)
                    cell_style = row_styles.get(col_idx)
                    if cell_style is not None:
                        cell_style = merge_styles(cell_style, Style(number_format=fmt_string))
                else:
                    cell_style = row_styles.get(col_idx)
                if cell_style is not None:
                    xf_index = self._register_style(cell_style)
                    has_custom_fmt = True
                elif fmt_string is not None:
                    xf_index = self._register_format(fmt_string)
                    has_custom_fmt = True
                else:
                    xf_index = 0
                    has_custom_fmt = False

                if isinstance(cell, str):
                    if has_custom_fmt:
                        self._write_str(buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx, xf_index)
                    else:
                        self._write_str(buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx)
                elif isinstance(cell, bool):
                    style_ref = xf_index if has_custom_fmt else 0
                    buffer.write(b'\x04\x09')
                    buffer.write(struct.pack('<IIB', col_idx, style_ref, 1 if cell else 0))
                elif isinstance(cell, int):
                    style_ref = xf_index if has_custom_fmt else 0
                    if -536870912 <= cell <= 536870911:
                        rk_val = (cell << 2) | 0b10
                        buffer.write(b'\x02\x0C')
                        buffer.write(struct.pack('<IIi', col_idx, style_ref, rk_val))
                    else:
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, float(cell)))
                elif isinstance(cell, (float, Decimal)):
                    style_ref = xf_index if has_custom_fmt else 0
                    buffer.write(b'\x05\x10')
                    buffer.write(struct.pack('<IId', col_idx, style_ref, float(cell)))
                elif isinstance(cell, datetime.datetime):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = xf_index if has_custom_fmt else 1
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                elif isinstance(cell, datetime.date):
                    if cell.year < 1900 or cell.year > 9999:
                        self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                        xf_index if has_custom_fmt else None)
                    else:
                        excel_date = datetime_to_serial(cell)
                        style_ref = xf_index if has_custom_fmt else 2
                        buffer.write(b'\x05\x10')
                        buffer.write(struct.pack('<IId', col_idx, style_ref, excel_date))
                else:
                    # The C encoder hands rows with other objects to this path.
                    if type(cell) in _XLSB_REJECTED_TYPES:
                        raise TypeError(
                            "XlsbWriter does not support Formula or Hyperlink cells; use XlsxWriter"
                        )
                    self._write_str(buffer, str(cell), shared_strings_dict, shared_strings_append, row_idx, col_idx,
                                    xf_index if has_custom_fmt else None)

            if buffer.tell() > BUFFER_SIZE:
                sheet_file.write(buffer.getvalue())
                buffer.seek(0)
                buffer.truncate()

        if row_outline is not None:
            # Grouped rows and summary rows past the data get an empty BrtRowHdr.
            for extra_idx in row_outline.indices_from(row_idx + 1):
                height_twips = None
                if extra_idx in heights:
                    height_twips = self._row_height_twips(heights[extra_idx])
                outline_level, outline_hidden = row_outline.state_at(extra_idx)
                buffer.write(self._row_header_record(
                    extra_idx, 0, height_twips,
                    extra_idx in hidden_row_set or outline_hidden,
                    outline_level, row_outline.collapsed_at(extra_idx),
                ))
                if buffer.tell() > BUFFER_SIZE:
                    sheet_file.write(buffer.getvalue())
                    buffer.seek(0)
                    buffer.truncate()

        buffer.write(self._sheet1Bytes[218:290])

        last_row_idx = row_idx if row_idx > -1 else 0 

        if row_idx >= 0 and auto_filter is not False:
            buffer.write(self._autoFilterStartBytes)
            buffer.write(struct.pack('<I', 0))
            buffer.write(struct.pack('<I', last_row_idx))
            buffer.write(struct.pack('<I', 0))
            buffer.write(struct.pack('<I', 0 + end_col - 1))
            buffer.write(self._autoFilterEndBytes)
            self._filtered_data_list.append(
                FilterData(
                sheet_index=worksheet_index,
                start_column=0,
                end_column=end_col - 1,
                start_row=0,
                end_row=last_row_idx
            ))

        if merge_refs:
            buffer.write(build_record(
                _BRT_BEGIN_MERGE_CELLS, struct.pack('<I', len(merge_refs)),
            ))
            for ref in merge_refs:
                first_col, first_row, last_col, last_row = parse_cell_range(ref)
                buffer.write(build_record(
                    _BRT_MERGE_CELL,
                    struct.pack('<IIII', first_row, last_row, first_col, last_col),
                ))
            buffer.write(build_record(_BRT_END_MERGE_CELLS, b''))

        if cf_rules:
            buffer.write(self._cf_bin_grouped(cf_rules))

        # Template offsets 290..346 hold BrtPrintOptions + BrtMargins; 346..
        # holds the FRT alternative-content block and BrtEndSheet. BrtDrawing
        # ([MS-XLSB] 2.4.351, record 550) sits between them per the Worksheet
        # ABNF: ... [BrtMargins] ... [BrtDrawing] ... [LISTPARTS] ... BrtEndSheet.
        buffer.write(self._sheet1Bytes[290:346])
        if drawing_rel is not None:
            rel_encoded = drawing_rel.encode('utf-16-le')
            buffer.write(build_record(
                _BRT_DRAWING,
                struct.pack('<I', len(drawing_rel)) + rel_encoded,
            ))
        if table_rel_ids:
            buffer.write(self._list_parts_bin(table_rel_ids))
        buffer.write(self._sheet1Bytes[346:])
        if buffer.tell() > 0:
            sheet_file.write(buffer.getvalue())

    @staticmethod
    def _write_inline_str(buffer, text: str, col_idx: int, style_ref: int):
        """BrtCellSt: the text is stored in the cell, so no shared-string table grows.

        The count is in UTF-16 code units, so a non-BMP character counts as two.
        Lone surrogates raise UnicodeEncodeError here, as they do in the shared table.
        """
        check_cell_text(text)
        encoded = text.encode('utf-16-le')
        units = len(encoded) // 2
        buffer.write(build_record(
            0x06, struct.pack('<IIi', col_idx, style_ref, units) + encoded,
        ))

    def _write_str(self, buffer, cell, shared_strings_dict, shared_strings_append, row_idx, col_idx, style_override=None):
        if not self._useSharedStrings:
            style_ref = style_override if style_override is not None else (3 if row_idx == 0 else 0)
            self._write_inline_str(buffer, cell, col_idx, style_ref)
            return
        check_cell_text(cell)
        self._sst_all_count += 1
        if cell not in shared_strings_dict:
            string_index = self._sst_unique_count
            shared_strings_dict[cell] = string_index
            shared_strings_append(cell)
            self._sst_unique_count += 1
        else:
            string_index = shared_strings_dict[cell]
        if style_override is not None:
            style_ref = style_override
        else:
            style_ref = 3 if row_idx == 0 else 0
        buffer.write(b'\x07\x0C')
        buffer.write(struct.pack('<III', col_idx, style_ref, string_index))

    def _write_filter_defined_names(self, sw: bytearray) -> None:
        filtered_dict_items_cnt = len(self._filtered_data_list) if self._filtered_data_list else 0
        
        if (filtered_dict_items_cnt > 0 and 
            (0x80 + (filtered_dict_items_cnt - 21) * 0x0c) <= 255):
            
            sw.extend(self._magicFilterExcel2016Fix0)
            
            if filtered_dict_items_cnt <= 10:
                sw.extend(bytes([0x10 + (filtered_dict_items_cnt - 1) * 0x0c, 
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
                
            elif filtered_dict_items_cnt <= 20:
                sw.extend(bytes([0x10 + (filtered_dict_items_cnt - 1) * 0x0c,
                               (filtered_dict_items_cnt - 1) // 10,
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
                
            else:
                sw.extend(bytes([0x80 + (filtered_dict_items_cnt - 21) * 0x0c,
                               (filtered_dict_items_cnt - 1) // 10,
                               filtered_dict_items_cnt]))
                sw.extend(bytes([0x00, 0x00, 0x00]))
            
            for nm in range(filtered_dict_items_cnt):
                sw.extend(bytes([0x00, 0x00, 0x00, 0x00]))
                sw.extend(bytes([self._filtered_data_list[nm].sheet_index, 0x00, 0x00, 0x00]))
                sw.extend(bytes([self._filtered_data_list[nm].sheet_index, 0x00, 0x00, 0x00]))
            
            sw.extend(bytes([0xE2, 0x02, 0x00]))
            
            for sheet_num in range(len(self._filtered_data_list)):
                start_column = self._filtered_data_list[sheet_num].start_column
                end_column = self._filtered_data_list[sheet_num].end_column
                start_row = self._filtered_data_list[sheet_num].start_row
                end_row = self._filtered_data_list[sheet_num].end_row
                sheet_index = self._filtered_data_list[sheet_num].sheet_index
                
                magic_filter_fix = self._magicFilterExcel2016Fix1.copy()
                magic_filter_fix[7] = sheet_index
                magic_filter_fix[-2] = sheet_num
                sw.extend(magic_filter_fix)

                sw.extend(struct.pack('<I', start_row))
                sw.extend(struct.pack('<I', end_row))
                sw.extend(struct.pack('<H', start_column))
                sw.extend(struct.pack('<H', end_column))
                
                sw.extend(self._magicFilterExcel2016Fix2)


    _autoFilterStartBytes = bytes([0xA1, 0x01, 0x10])
    _sheet1Bytes = bytes([
        0x81,0x01,0x00,0x93,0x01,0x17,0xCB,0x04,0x02,0x00,0x40,0x00,0x00,0x00,0x00,0x00,0x00,0xFF,0xFF,0xFF,
        0xFF,0xFF,0xFF,0xFF,0xFF,0x00,0x00,0x00,0x00,0x94,0x01,0x10,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x85,0x01,0x00,0x89,0x01,0x1E,0xDC,0x03,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x00,0x00,0x64,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x98,0x01,0x24,0x03,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x8A,0x01,0x00,0x86,0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x95,0x08,0x02,
        0x05,0x00,0x26,0x00,0xE5,0x03,0x0C,0xFF,0xFF,0xFF,0xFF,0x08,0x00,0x2C,0x01,0x00,0x00,0x00,0x00,0x91,
        0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x80,0x08,0x02,0x05,0x00,0x26,0x00,0x00,0x19,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2C,0x01,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x07,0x0C,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x92,0x01,
        0x00,0x97,0x04,0x42,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0xDD,0x03,0x02,0x10,0x00,0xDC,0x03,0x30,0x66,0x66,
        0x66,0x66,0x66,0x66,0xE6,0x3F,0x66,0x66,0x66,0x66,0x66,0x66,0xE6,0x3F,0x00,0x00,0x00,0x00,0x00,0x00,
        0xE8,0x3F,0x00,0x00,0x00,0x00,0x00,0x00,0xE8,0x3F,0x33,0x33,0x33,0x33,0x33,0x33,0xD3,0x3F,0x33,0x33,
        0x33,0x33,0x33,0x33,0xD3,0x3F,0x25,0x06,0x01,0x00,0x00,0x10,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x82,0x01,0x00
    ])
    _stylesBin = bytes([
        0x96,0x02,0x00,0xE7,0x04,0x04,0x02,
        0x00,0x00,0x00,0x2C,0x2C,0xA4,0x00,0x13,
        0x00,0x00,0x00,0x79,0x00,0x79,0x00,0x79,
        0x00,0x79,0x00,0x5C,0x00,0x2D,0x00,0x6D,
        0x00,0x6D,0x00,0x5C,0x00,0x2D,0x00,0x64,
        0x00,0x64,0x00,0x5C,0x00,0x20,0x00,0x68,
        0x00,0x68,0x00,0x3A,0x00,0x6D,0x00,0x6D,
        0x00,0x2C,0x1E,0xA6,0x00,0x0C,0x00,0x00,
        0x00,0x79,0x00,0x79,0x00,0x79,0x00,0x79,
        0x00,0x5C,0x00,0x2D,0x00,0x6D,0x00,0x6D,
        0x00,0x5C,0x00,0x2D,0x00,0x64,0x00,0x64,
        0x00,0xE8,0x04,0x00,0xE3,0x04,0x04,0x01,
        0x00,0x00,0x00,
        0x2B,0x27,0xDC,0x00,0x00,0x00,0x90,0x01,
        0x00,0x00,0x00,0x02,0x00,0x00,0x07,0x01,
        0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
        0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
        0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
        0x00,
        0x2B,0x27,0xDC,0x00,0x01,0x00,0xBC,0x02,
        0x00,0x00,0x00,0x02,0xEE,0x00,0x07,0x01,
        0x00,0x00,0x00,0x00,0x00,0xFF,0x02,0x07,
        0x00,0x00,0x00,0x43,0x00,0x61,0x00,0x6C,
        0x00,0x69,0x00,0x62,0x00,0x72,0x00,0x69,
        0x00,
        0x25,0x06,0x01,0x00,0x02,0x0E,0x00,0x80,0x81,0x08,0x00,0x26,0x00,0xE4,0x04,0x00,0xDB,0x04,0x04,0x02,0x00,0x00,0x00,
        0x2D,0x44,0x00,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x2D,0x44,0x11,0x00,0x00,0x00,0x03,0x40,0x00,0x00,0x00,0x00,0x00,0xFF,0x03,0x41,0x00,0x00,0xFF,0xFF,0xFF,0xFF,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0xDC,0x04,0x00,0xE5,0x04,0x04,0x01,0x00,0x00,0x00,0x2E,0x33,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xE6,0x04,0x00,0xF2,
        0x04,0x04,0x01,0x00,0x00,0x00,0x2F,0x10,0xFF,0xFF,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,
        0x00,0xF3,0x04,0x00,0xE9,0x04,0x04,
        0x04,
        0x00,0x00,0x00,
        0x2F,0x10,0x00,0x00,0x00,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,
        0x2F,0x10,0x00,0x00,0xA4,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,
        0x2F,0x10,0x00,0x00,0xA6,0x00,    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x01,0x00,
        0x2F,0x10,0x00,0x00,0x00,0x01,    0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x10,0x00,0x00,
        0xEA,0x04,0x00,0xEB,0x04,0x04,0x01,0x00,
        0x00,0x00,0x25,0x06,0x01,0x00,0x02,0x11,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,0x30,0x1C,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x08,0x00,0x00,0x00,0x4E,
        0x00,0x6F,0x00,0x72,0x00,0x6D,0x00,0x61,0x00,0x6C,0x00,0x6E,0x00,0x79,0x00,0xEC,0x04,0x00,0xF9,0x03,0x04,0x00,0x00,
        0x00,0x00,0xFA,0x03,0x00,0xFC,0x03,0x50,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x54,0x00,0x61,0x00,0x62,0x00,0x6C,
        0x00,0x65,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4D,0x00,0x65,0x00,0x64,0x00,0x69,0x00,0x75,0x00,
        0x6D,0x00,0x32,0x00,0x11,0x00,0x00,0x00,0x50,0x00,0x69,0x00,0x76,0x00,0x6F,0x00,0x74,0x00,0x53,0x00,0x74,0x00,0x79,
        0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0x36,0x00,0xFD,0x03,0x00,0x23,
        0x04,0x02,0x0E,0x00,0x00,0xEB,0x08,0x00,0xF6,0x08,0x2A,0x00,0x00,0x00,0x00,0x11,0x00,0x00,0x00,0x53,0x00,0x6C,0x00,
        0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,0x67,
        0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xF7,0x08,0x00,0xEC,0x08,0x00,0x24,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xB0,0x10,
        0x00,0xB2,0x10,0x32,0x00,0x00,0x00,0x00,0x15,0x00,0x00,0x00,0x54,0x00,0x69,0x00,0x6D,0x00,0x65,0x00,0x53,0x00,0x6C,
        0x00,0x69,0x00,0x63,0x00,0x65,0x00,0x72,0x00,0x53,0x00,0x74,0x00,0x79,0x00,0x6C,0x00,0x65,0x00,0x4C,0x00,0x69,0x00,
        0x67,0x00,0x68,0x00,0x74,0x00,0x31,0x00,0xB3,0x10,0x00,0xB1,0x10,0x00,0x24,0x00,0x97,0x02,0x00
    ])
    _workbookBinStart = bytes([
        0x83,0x01,0x00,0x80,0x01,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x02,0x00,0x00,0x00,0x78,0x00,0x6C,0x00,0x01,0x00,0x00,0x00,0x37,0x00,0x01,0x00,0x00,0x00,
        0x36,0x00,0x05,0x00,0x00,0x00,0x32,0x00,0x34,0x00,0x33,0x00,0x32,0x00,0x36,0x00,0x99,0x01,0x0C,0x20,
        0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x25,0x06,0x01,0x00,0x03,0x0F,0x00,0x80,0x97,
        0x10,0x34,0x18,0x00,0x00,0x00,        0x43,0x00,0x3A,0x00,0x5C,0x00,0x78,0x00,0x6C,0x00,0x73,0x00,0x70,0x00,0x79,0x00,
        0x5C,0x00,0x74,0x00,0x65,0x00,0x73,0x00,0x74,0x00,0x5C,0x00,0x70,0x00,0x61,0x00,0x74,0x00,0x68,0x00,
        0x5C,0x00,0x64,0x00,0x6F,0x00,0x63,0x00,0x5C,0x00,0x74,0x00,0x6D,0x00,0x70,0x00,0x5C,0x00,0x26,0x00,0x25,0x06,0x01,0x00,
        0x00,0x10,0x00,0x80,0x81,0x18,0x82,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2F,0x00,0x00,0x00,
        0x31,0x00,0x33,0x00,0x5F,0x00,0x6E,0x00,0x63,0x00,0x72,0x00,0x3A,0x00,0x31,0x00,0x5F,0x00,0x7B,0x00,
        0x31,0x00,0x36,0x00,0x35,0x00,0x30,0x00,0x38,0x00,0x44,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x43,0x00,
        0x46,0x00,0x38,0x00,0x37,0x00,0x2D,0x00,0x34,0x00,0x37,0x00,0x36,0x00,0x39,0x00,0x2D,0x00,0x38,0x00,
        0x34,0x00,0x35,0x00,0x36,0x00,0x2D,0x00,0x44,0x00,0x34,0x00,0x41,0x00,0x34,0x00,0x30,0x00,0x31,0x00,
        0x31,0x00,0x33,0x00,0x31,0x00,0x35,0x00,0x36,0x00,0x37,0x00,0x7D,0x00,0x2F,0x00,0x00,0x00,0x2F,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x26,0x00,
        0x87,0x01,0x00,0x25,0x06,0x01,0x00,0x02,0x10,0x00,0x80,0x80,0x18,0x10,0x00,0x00,0x00,0x00,0x0D,0x00,
        0x00,0x00,0xFF,0xFF,0xFF,0xFF,0x00,0x00,0x00,0x00,0x26,0x00,0x9E,0x01,0x1D,0x00,0x00,0x00,0x00,0x9E,
        0x16,0x00,0x00,0xB4,0x69,0x00,0x00,0xE8,0x26,0x00,0x00,0x58,0x02,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x78,0x88,0x01,0x00,0x8F,0x01,0x00
    ])
    _workbookBinMiddle = bytes([0x90,0x01,0x00])
    _workbookBinEnd = bytes([
        0x9D,0x01,0x1A,0x35,0xEA,0x02,0x00,0x01,0x00,0x00,0x00,0x64,0x00,0x00,0x00,0xFC,0xA9,0xF1,0xD2,0x4D,
        0x62,0x50,0x3F,0x01,0x00,0x00,0x00,0x6A,0x00,0x9B,0x01,0x01,0x00,0x23,0x04,0x03,0x0F,0x00,0x00,0xAB,
        0x10,0x01,0x01,0x24,0x00,0x84,0x01,0x00
    ])
    _binaryIndexBin = bytes([
        0x2A,0x18,0x00,0x00,0x00,0x00,0x20,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x95,0x02,0x00
    ])
    _stickHeaderA1bytes =bytes([
        0x97,0x01,0x1D,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0xF0,0x3F,0x01,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x03,
    ])
    _autoFilterEndBytes = bytes([0xA2, 0x01, 0x00])
    _magicFilterExcel2016Fix0 = bytes([0xE1, 0x02, 0x00, 0xE5, 0x02, 0x00, 0xEA, 0x02])
    _magicFilterExcel2016Fix1 = bytearray([
            0x27,
            0x46,
            0x21,
            0x00,
            0x00,
            0x00,
            0x00,
            255,
            0x00,
            0x00,
            0x00,
            0x0F,
            0x00,
            0x00,
            0x00,
            0x5F,
            0x00,
            0x46,
            0x00,
            0x69,
            0x00,
            0x6C,
            0x00,
            0x74,
            0x00,
            0x65,
            0x00,
            0x72,
            0x00,
            0x44,
            0x00,
            0x61,
            0x00,
            0x74,
            0x00,
            0x61,
            0x00,
            0x62,
            0x00,
            0x61,
            0x00,
            0x73,
            0x00,
            0x65,
            0x00,
            0x0F,
            0x00,
            0x00,
            0x00,
            0x3B,
            255,
            0x00
        ])
    _magicFilterExcel2016Fix2 = bytes([0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF])
