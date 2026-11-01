"""XlsbWriter sheet-view records: row headers, panes, selections, column widths."""

from typing import TYPE_CHECKING, Any
import struct
import io
from ._widths import file_width
from .biff_utils import build_record
from ._xlsb_writer_consts import _BRT_BEGIN_COL_INFOS, _BRT_COL_INFO, _BRT_END_COL_INFOS, _BRT_PANE, _BRT_ROW_HDR, _BRT_SEL, _PANE_ACTIVE


class _XlsbViewsMixin:
    """Methods of the writer class, mixed into it; see the facade module."""
    if TYPE_CHECKING:
        # Attributes of the writer class that these methods use; it sets them.
        _split_freeze_cell: Any
        _stickHeaderA1bytes: Any

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

        # Explicit widths and hidden columns past the sampled width still need a
        # BrtColInfo; dropping them silently lost the caller's settings.
        columns = set(range(start_col, end_col)) | set(widths) | set(hidden_cols)
        if col_outline is not None:
            # Grouped columns past the data still need a BrtColInfo.
            columns |= set(col_outline.indices())
        for i in sorted(columns):
            explicit = widths.get(i)
            if explicit is not None:
                coldx = int(round(explicit * 256))
            elif i - start_col < len(col_widths_array):
                # Automatic widths count characters; the file adds the cell padding.
                coldx = int(round(file_width(col_widths_array[i - start_col]) * 256))
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
