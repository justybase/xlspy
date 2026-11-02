"""Cell styles of an XLSB package (``xl/styles.bin``) for the editable model.

``XlsbStyleTable`` decodes the binary style sheet into the same component keys
``XlsxStyleTable`` keeps (``split_style``), in file order, so a cell's style id
is the file's XF index. Styles added through the model are appended: on save
the new fonts, fills, borders, number formats and cell XFs are spliced into
their blocks of the original part, which otherwise stays byte for byte.
"""

from __future__ import annotations
from collections.abc import Sequence

from .exceptions import UnsupportedFeatureError
import struct

from typing import TYPE_CHECKING, Any

from ._xlsb_metadata import read_wide
from ._xlsb_style_records import (
    _BORDER_BYTES,
    _FONT_SCHEME,
    _HORIZONTAL_BYTES,
    _PATTERN_BYTES,
    _UNDERLINE_BYTES,
    _VERTICAL_BYTES,
    border_record,
    fill_record,
    font_record,
    numfmt_record,
    xf_record,
)
from .biff_utils import build_record, iter_records
from .styles import (
    Alignment,
    Border,
    Color,
    Font,
    GradientFill,
    NamedStyle,
    PatternFill,
    Protection,
    Side,
    Stop,
    Style,
    _alignment_key_from,
    _border_key,
    _fill_key_from,
    _font_is_empty,
    _font_key_from,
    _protection_key_from,
    fill_key_for,
)
from .workbook._styles import XlsxStyleTable, _Xf

if TYPE_CHECKING:
    from .styles import DifferentialStyle

_FONT, _FMT, _FILL, _BORDER, _XF, _STYLE = 0x2B, 0x2C, 0x2D, 0x2E, 0x2F, 0x30
_DXF = 0x1FB
_BEGIN_DXFS, _END_DXFS = 0x1F9, 0x1FA
_END_STYLES = 0x28C  # BrtEndStyles: the DXFs block follows the cell styles.
_BEGIN_CELL_STYLES, _END_CELL_STYLES = 0x26B, 0x26C  # BrtBeginStyle / BrtEndStyle: the BrtStyle records
_AFTER_DXFS = (0x1FC, 0x1D9, 0x23, 0x117)  # table styles, colors, FRT, end
_BEGIN_FILLS, _END_FILLS = 0x25B, 0x25C
_BEGIN_FONTS, _END_FONTS = 0x263, 0x264
_BEGIN_BORDERS, _END_BORDERS = 0x265, 0x266
_BEGIN_FMTS, _END_FMTS = 0x267, 0x268
_BEGIN_CELL_XFS, _END_CELL_XFS = 0x269, 0x26A
_BEGIN_STYLE_XFS, _END_STYLE_XFS = 0x272, 0x273
_AC_BEGIN = 0x25  # BrtACBegin: the fonts block ends with KnownFonts inside one.

_BIFF_UNDERLINE_NAMES = {value: name for name, value in _UNDERLINE_BYTES.items() if name not in (None, "none")}
_BASELINES = {1: "superscript", 2: "subscript"}
_SCHEMES = {value: name for name, value in _FONT_SCHEME.items() if name is not None}
_BIFF_PATTERN_NAMES = {value: name for name, value in _PATTERN_BYTES.items() if name is not None}
_BIFF_BORDER_NAMES = {value: name for name, value in _BORDER_BYTES.items() if name not in (None, "none")}
_BIFF_HORIZONTAL_NAMES = {value: name for name, value in _HORIZONTAL_BYTES.items() if name not in (None, "general")}
_BIFF_VERTICAL_NAMES = {value: name for name, value in _VERTICAL_BYTES.items() if name not in (None, "bottom")}
# System foreground/background colors: what Excel stores for "no color".
_SYSTEM_COLORS = (64, 65)
_GRADIENT = 0x28


def _read_brt_color(data: bytes, offset: int, *, default_system: bool = False) -> Color | None:
    """A BrtColor ([MS-XLSB] 2.4.334); all zero bytes mean no color."""
    raw = data[offset:offset + 8]
    if len(raw) < 8 or not any(raw):
        return None
    kind = raw[0] >> 1
    if kind == 0:
        return Color(auto=True)
    if kind == 1:
        if default_system and raw[1] in _SYSTEM_COLORS:
            return None
        return Color(indexed=raw[1])
    if kind == 2:
        red, green, blue, alpha = raw[4:8]
        return Color(rgb=f"{alpha:02X}{red:02X}{green:02X}{blue:02X}")
    tint = struct.unpack_from("<h", raw, 2)[0]
    return Color(theme=raw[1], tint=tint / 32767 if tint else None)


def _font(payload: bytes) -> Font:
    height, flags, weight, baseline = struct.unpack_from("<HHHH", payload, 0)
    underline, family, charset = payload[8], payload[9], payload[10]
    color = _read_brt_color(payload, 12)
    scheme = payload[20]
    name, _ = read_wide(payload, 21)

    def flag(bit: int) -> bool | None:
        return True if flags & bit else None

    return Font(
        name=name or None,
        size=height / 20,
        bold=True if weight >= 700 else None,
        italic=flag(0x02),
        strike=flag(0x08),
        underline=_BIFF_UNDERLINE_NAMES.get(underline),
        vertAlign=_BASELINES.get(baseline),
        color=color,
        charset=charset or None,
        family=family or None,
        scheme=_SCHEMES.get(scheme),
        outline=flag(0x10),
        shadow=flag(0x20),
        condense=flag(0x40),
        extend=flag(0x80),
    )


def _fill(payload: bytes) -> PatternFill | GradientFill:
    pattern = struct.unpack_from("<I", payload, 0)[0]
    if pattern == _GRADIENT and len(payload) >= 64:
        kind = struct.unpack_from("<I", payload, 20)[0]
        angle = struct.unpack_from("<d", payload, 24)[0]
        left, right, top, bottom = struct.unpack_from("<dddd", payload, 32)
        count = struct.unpack_from("<I", payload, 64)[0]
        stops = []
        offset = 68
        for _ in range(count):
            if offset + 16 > len(payload):
                break
            color = _read_brt_color(payload, offset)
            position = struct.unpack_from("<d", payload, offset + 8)[0]
            if color is not None:
                stops.append(Stop(color=color, position=position))
            offset += 16
        if not stops:
            return PatternFill(fill_type="none")
        if kind == 0:
            return GradientFill(fill_type="linear", degree=(360 - angle) % 360, stop=tuple(stops))
        return GradientFill(
            fill_type="path",
            left=None if left == 0.0 else left,
            right=None if right == 1.0 else right,
            top=None if top == 0.0 else top,
            bottom=None if bottom == 1.0 else bottom,
            stop=tuple(stops),
        )
    fill_type = _BIFF_PATTERN_NAMES.get(pattern, "none")
    if fill_type == "none":
        return PatternFill(fill_type="none")
    return PatternFill(
        fill_type=fill_type,
        start_color=_read_brt_color(payload, 4, default_system=True),
        end_color=_read_brt_color(payload, 12, default_system=True),
    )


def _border(payload: bytes) -> Border:
    diagonal = payload[0]
    sides = []
    for index in range(5):
        offset = 1 + index * 10
        style = _BIFF_BORDER_NAMES.get(payload[offset])
        color = _read_brt_color(payload, offset + 2, default_system=True) if style else None
        sides.append(Side(style=style, color=color) if style else None)
    top, bottom, left, right, diagonal_side = sides
    return Border(
        left=left, right=right, top=top, bottom=bottom, diagonal=diagonal_side,
        diagonal_up=True if diagonal & 0x02 else None,
        diagonal_down=True if diagonal & 0x01 else None,
    )


def _xf(payload: bytes) -> tuple[int, _Xf]:
    parent, ifmt, font, fill, border = struct.unpack_from("<HHHHH", payload, 0)
    rotation, indent, align, flags = payload[10], payload[11], payload[12], payload[13]
    horizontal = _BIFF_HORIZONTAL_NAMES.get(align & 0x07)
    vertical = _BIFF_VERTICAL_NAMES.get((align >> 3) & 0x07)
    reading_order = (flags >> 2) & 0x03
    alignment = None
    if (horizontal or vertical or align & 0xC0 or flags & 0x01 or rotation or indent
            or reading_order):
        alignment = Alignment(
            horizontal=horizontal,
            vertical=vertical,
            wrap_text=True if align & 0x40 else None,
            shrink_to_fit=True if flags & 0x01 else None,
            indent=indent or None,
            text_rotation=rotation or None,
            justify_last_line=True if align & 0x80 else None,
            reading_order=reading_order or None,
        )
    protection = None
    if not flags & 0x10 or flags & 0x20:
        protection = Protection(
            locked=None if flags & 0x10 else False,
            hidden=True if flags & 0x20 else None,
        )
    return parent, _Xf(
        font_id=font, fill_id=fill, border_id=border, num_fmt_id=ifmt,
        alignment=alignment, protection=protection,
    )


class XlsbStyleTable(XlsxStyleTable):
    """The style sheet of a loaded XLSB package, with append-only additions."""

    # Set by from_bin() and the append paths; declared here for the type checker
    # (annotations only, so no attribute is created at class level).
    _source_bin: bytes
    _saved_counts: tuple[int, int, int, int]
    _cell_style_count: int
    _font_keys: list[Any]
    _fill_keys: list[Any]
    _border_keys: list[Any]
    _dxf_keys: list[Any]
    _new_dxfs: list[Any]
    _new_style_xfs: list[Any]
    _new_cell_styles: list[Any]

    @classmethod
    def from_bin(cls, data: bytes | None) -> XlsbStyleTable:
        table = cls()
        table._source_bin = data or b""
        table._cell_style_count = 0
        if data:
            table._parse_bin(data)
        if not table.xfs:
            table._add_default_components()
        table._saved_counts = (
            len(table.fonts), len(table.fills), len(table.borders), len(table.xfs),
        )
        table._saved_formats = set(table.custom_num_formats)
        table._saved_dxfs = len(table.dxfs)
        table._new_dxfs = []
        table._dxf_keys = []
        table._new_style_xfs = []
        table._new_cell_styles = []
        return table

    def _parse_bin(self, data: bytes) -> None:
        style_xfs: list[tuple[int, _Xf]] = []
        cell_styles: list[tuple[str, int, int | None, bool]] = []
        block = None
        for record in iter_records(data):
            kind = record.record_id
            payload = data[record.data_start:record.data_end]
            if kind in (_BEGIN_STYLE_XFS, _BEGIN_CELL_XFS):
                block = kind
            elif kind in (_END_STYLE_XFS, _END_CELL_XFS):
                block = None
            elif kind == _FMT:
                ifmt = struct.unpack_from("<H", payload, 0)[0]
                code, _ = read_wide(payload, 2)
                self.num_formats[ifmt] = code
                self.custom_num_formats.add(ifmt)
                self._next_numfmt_id = max(self._next_numfmt_id, ifmt + 1)
            elif kind == _FONT:
                font = _font(payload)
                key = _font_key_from(font)
                self.fonts.append(font)
                self._font_keys.append(None if _font_is_empty(key) else key)
                self._font_xml.append("")
            elif kind == _FILL:
                fill = _fill(payload)
                self.fills.append(fill)
                self._fill_keys.append(_fill_key_from(fill))
                self._fill_xml.append("")
            elif kind == _BORDER:
                border = _border(payload)
                self.borders.append(border)
                self._border_keys.append(_border_key(border))
                self._border_xml.append("")
            elif kind == _XF:
                if block == _BEGIN_CELL_XFS:
                    self.xfs.append(_xf(payload)[1])
                    self._xf_xml.append("")
                elif block == _BEGIN_STYLE_XFS:
                    style_xfs.append(_xf(payload))
            elif kind == _DXF:
                from ._xlsb_cf import dxf_from_payload

                try:
                    self.dxfs.append(dxf_from_payload(payload))
                except (struct.error, IndexError, TypeError, ValueError):
                    self.dxfs.append(None)
            elif kind == _STYLE and len(payload) >= 8:
                ixf, grbit = struct.unpack_from("<IH", payload, 0)
                builtin_byte = payload[6]
                name, _ = read_wide(payload, 8)
                cell_styles.append((name, ixf, builtin_byte if grbit & 0x01 else None, bool(grbit & 0x02)))
        self._cell_style_count = len(cell_styles)
        for _parent, xf in style_xfs:
            self.cell_style_xfs.append(self._style_for(xf))
        for name, ixf, builtin, hidden in cell_styles:
            if not name:
                continue
            if name not in self.named_styles:
                self.named_styles.append(name)
            if 0 <= ixf < len(self.cell_style_xfs):
                base = self.cell_style_xfs[ixf]
                self.named_style_defs[name] = NamedStyle(
                    name=name, font=base.font, fill=base.fill, border=base.border,
                    number_format=base.number_format, alignment=base.alignment,
                    protection=base.protection, builtinId=builtin, hidden=hidden,
                )
        if self.fonts:
            self._rebuild_indexes()

    def _style_for(self, xf: _Xf) -> Style:
        def component(values: Sequence[Any], index: int) -> Any:
            return values[index] if 0 <= index < len(values) else None

        number_format = self.num_formats.get(xf.num_fmt_id)
        return Style(
            font=component(self.fonts, xf.font_id),
            fill=component(self.fills, xf.fill_id),
            border=component(self.borders, xf.border_id),
            number_format=None if number_format in (None, "General") else number_format,
            alignment=xf.alignment,
            protection=xf.protection,
        )

    # -- editing ------------------------------------------------------
    def add(self, style: Style) -> int:
        from .writer import XlsbWriter

        XlsbWriter._reject_rich_style(style)
        return super().add(style)

    def add_named_style(self, style: NamedStyle) -> int:
        """Append a named style: a style XF and a BrtStyle record."""
        if not isinstance(style, NamedStyle):
            raise TypeError("style must be a NamedStyle")
        if style.name in self.named_styles:
            raise ValueError(f"a named style {style.name!r} already exists")
        base = Style(
            font=style.font, fill=style.fill, border=style.border,
            number_format=style.number_format, alignment=style.alignment,
            protection=style.protection,
        )
        number_format = base.number_format
        xf = _Xf(
            num_fmt_id=self._intern_numfmt(number_format) if number_format and number_format != "General" else 0,
            font_id=self._intern_font(_font_key_from(base.font)) if base.font else 0,
            fill_id=self._intern_fill(fill_key_for(base.fill)) if base.fill else 0,
            border_id=self._intern_border(_border_key(base.border)) if base.border else 0,
            alignment=base.alignment, protection=base.protection,
        )
        xf_index = len(self.cell_style_xfs)
        record = bytearray(self._xf_record(xf))
        # ixfeParent of a style XF is 0xFFF; build_record put the XF payload last.
        self._new_style_xfs.append(bytes(self._style_xf_record(xf)))
        self.cell_style_xfs.append(base)
        grbit = (0x01 if style.builtinId is not None else 0) | (0x02 if style.hidden else 0)
        name = style.name.encode("utf-16-le")
        self._new_cell_styles.append(build_record(_STYLE, struct.pack(
            "<IHBBI", xf_index, grbit, style.builtinId or 0, 0xFF, len(style.name),
        ) + name))
        self._cell_style_count += 1
        self.named_styles.append(style.name)
        self.named_style_defs[style.name] = style
        del record
        return xf_index

    def _style_xf_record(self, xf: _Xf) -> bytes:
        data = bytearray(self._xf_record(xf))
        # The record is id (1 byte) + length (1 byte) + payload; ixfeParent (0xFFFF for a style XF) leads the payload.
        data[2:4] = bytes((0xFF, 0xFF))
        return bytes(data)

    def to_bytes(self) -> bytes:
        raise UnsupportedFeatureError("an XLSB style sheet is saved with to_bin()")

    def dxf_id(self, dxf: DifferentialStyle) -> int:
        """The index of a conditional format's differential style; a style
        the file does not hold yet is added once."""
        from ._xlsb_cf import dxf_payload

        if dxf is None:
            return -1
        # Styles are compared as written: a loaded style spells out what an
        # edited one leaves implicit (a solid fill's background color).
        wanted = dxf_payload(dxf, lambda _code: 0)
        known = self._dxf_keys
        for index in range(len(known), len(self.dxfs)):
            existing = self.dxfs[index]
            try:
                known.append(None if existing is None else dxf_payload(existing, lambda _code: 0))
            except (KeyError, TypeError, ValueError):
                known.append(None)
        if wanted in known:
            return known.index(wanted)
        self._new_dxfs.append(dxf_payload(dxf, self._intern_numfmt))
        self.dxfs.append(dxf)
        known.append(wanted)
        return len(self.dxfs) - 1

    def mark_saved(self) -> None:
        """The current components are now the file's own."""
        self._source_bin = self.to_bin()
        self._saved_counts = (
            len(self.fonts), len(self.fills), len(self.borders), len(self.xfs),
        )
        self._saved_formats = set(self.custom_num_formats)
        self._saved_dxfs = len(self.dxfs)
        self._new_dxfs = []
        self._new_style_xfs = []
        self._new_cell_styles = []

    def is_modified(self) -> bool:
        counts = (len(self.fonts), len(self.fills), len(self.borders), len(self.xfs))
        return (counts != self._saved_counts or self.custom_num_formats != self._saved_formats
                or bool(self._new_dxfs) or bool(self._new_cell_styles))

    def to_bin(self) -> bytes:
        """The original ``styles.bin`` with the components added since loading."""
        if not self.is_modified():
            return self._source_bin
        saved_fonts, saved_fills, saved_borders, saved_xfs = self._saved_counts
        additions = {
            _END_FMTS: [
                numfmt_record(ifmt, self.num_formats[ifmt])
                for ifmt in sorted(self.custom_num_formats - self._saved_formats)
            ],
            _END_FONTS: [font_record(key) for key in self._font_keys[saved_fonts:]],
            _END_FILLS: [fill_record(key) for key in self._fill_keys[saved_fills:]],
            _END_BORDERS: [border_record(key) for key in self._border_keys[saved_borders:]],
            _END_CELL_XFS: [self._xf_record(xf) for xf in self.xfs[saved_xfs:]],
            _END_DXFS: [build_record(_DXF, payload) for payload in self._new_dxfs],
            _END_STYLE_XFS: self._new_style_xfs,
            _END_CELL_STYLES: self._new_cell_styles,
        }
        counts = {
            _BEGIN_FMTS: (_END_FMTS, len(self.custom_num_formats)),
            _BEGIN_FONTS: (_END_FONTS, len(self.fonts)),
            _BEGIN_FILLS: (_END_FILLS, len(self.fills)),
            _BEGIN_BORDERS: (_END_BORDERS, len(self.borders)),
            _BEGIN_CELL_XFS: (_END_CELL_XFS, len(self.xfs)),
            _BEGIN_DXFS: (_END_DXFS, len(self.dxfs)),
            _BEGIN_STYLE_XFS: (_END_STYLE_XFS, len(self.cell_style_xfs)),
            _BEGIN_CELL_STYLES: (_END_CELL_STYLES, self._cell_style_count),
        }
        data = self._source_bin
        missing = {end for end, records in additions.items() if records}
        if _END_DXFS in missing and not any(
            record.record_id == _BEGIN_DXFS for record in iter_records(data)
        ):
            data = self._with_dxfs_block(data)
        chunks: list[bytes] = []
        in_fonts = False
        for record in iter_records(data):
            kind = record.record_id
            whole = data[record.header_start:record.data_end]
            if kind in counts:
                chunks.append(build_record(kind, struct.pack("<I", counts[kind][1])))
                in_fonts = kind == _BEGIN_FONTS
                continue
            # New fonts go before the BrtACBegin/KnownFonts extension, as Excel writes them.
            if in_fonts and kind in (_AC_BEGIN, _END_FONTS) and _END_FONTS in missing:
                chunks.extend(additions[_END_FONTS])
                missing.discard(_END_FONTS)
            if kind == _END_FONTS:
                in_fonts = False
            elif kind in additions and kind in missing:
                chunks.extend(additions[kind])
                missing.discard(kind)
            chunks.append(whole)
        if missing:
            raise UnsupportedFeatureError(
                "the XLSB style sheet has no block for the new styles "
                f"(record ids {sorted(missing)})"
            )
        return b"".join(chunks)

    @staticmethod
    def _with_dxfs_block(data: bytes) -> bytes:
        """``data`` with an empty DXFs block where Excel writes it."""
        block = build_record(_BEGIN_DXFS, struct.pack("<I", 0)) + build_record(_END_DXFS)
        chunks, inserted, after_styles = [], False, False
        for record in iter_records(data):
            if not inserted and (after_styles or record.record_id == 0x117) and (
                record.record_id in _AFTER_DXFS
            ):
                chunks.append(block)
                inserted = True
            after_styles = after_styles or record.record_id == _END_STYLES
            chunks.append(data[record.header_start:record.data_end])
        if not inserted:
            raise UnsupportedFeatureError("the XLSB style sheet has no place for conditional format styles")
        return b"".join(chunks)

    def _xf_record(self, xf: _Xf) -> bytes:
        alignment = _alignment_key_from(xf.alignment) if xf.alignment is not None else None
        if alignment is not None and all(value is None for value in alignment):
            alignment = None
        protection = _protection_key_from(xf.protection) if xf.protection is not None else None
        if protection is not None and all(value is None for value in protection):
            protection = None
        return xf_record(
            xf.num_fmt_id, font_id=xf.font_id, fill_id=xf.fill_id,
            border_id=xf.border_id, alignment=alignment, protection=protection,
            apply_flags=True,
        )
