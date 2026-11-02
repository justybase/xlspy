"""XlsbWriter style methods: number formats, xf registration, styles.bin records."""

from .exceptions import UnsupportedFeatureError
import struct
from .styles import (
    Style,
    split_style,
)
from .biff_utils import build_record
from ._xlsb_style_records import (
    border_record,
    fill_record,
    font_record,
    xf_record,
)
from ._xlsb_cf import dxfs_bin
from ._xlsb_writer_consts import _BRT_BEGIN_BORDERS, _BRT_BEGIN_CELL_XFS, _BRT_BEGIN_FILLS, _BRT_BEGIN_FMTS, _BRT_BEGIN_FONTS, _BRT_END_BORDERS, _BRT_END_CELL_XFS, _BRT_END_FILLS, _BRT_END_FMTS, _BRT_END_FONTS
from ._xlsb_writer_host import _XlsbWriterHost


class _XlsbStylesMixin(_XlsbWriterHost):
    """Methods of the writer class, mixed into it; see the facade module."""

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
            raise UnsupportedFeatureError(
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
                raise UnsupportedFeatureError(
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
        self._xf_styles[xf_index] = style
        self._style_components[xf_index] = (
            font_id, fill_id, border_id, numfmt_id, alignment, protection,
        )
        return xf_index

    def _xlsb_font_record(self, font_key: tuple) -> bytes:
        return font_record(font_key)

    def _xlsb_fill_record(self, fill_key: tuple) -> bytes:
        return fill_record(fill_key)

    def _xlsb_border_record(self, border_key: tuple) -> bytes:
        return border_record(border_key)

    def _build_style_record(self, ifmt: int = 0, **options) -> bytes:
        return xf_record(ifmt, **options)

    def _build_styles_bin(self) -> bytes:
        if not (self._format_registry or self._style_components or self._dxf_index or self._font_registry):
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
        if self._dxf_index:
            # Template bytes 588..598 are the empty BrtBeginDXFs/BrtEndDXFs pair.
            buf.extend(template[519:588])
            buf.extend(dxfs_bin(list(self._dxf_index)))
            buf.extend(template[598:])
        else:
            buf.extend(template[519:])

        return bytes(buf)
