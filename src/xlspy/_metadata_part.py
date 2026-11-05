"""The ``xl/metadata.xml`` part of a workbook with dynamic-array formulas.

Excel records every dynamic-array anchor as a cell with ``cm="1"`` and writes
this part: one ``XLDAPR`` metadata type, one future-metadata block carrying the
``dynamicArrayProperties`` and one cell-metadata block. The block is the same
for every anchor, so one entry serves the whole workbook.
"""

from __future__ import annotations

METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1">'
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
    'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    'clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
    "</metadataTypes>"
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    "</ext></extLst></bk></futureMetadata>"
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    "</metadata>"
)

# The same part in XLSB, captured from a file Excel saved with a dynamic array.
METADATA_BIN = bytes.fromhex(
    "cc0200ce020401000000cf0218b0c06ad8c0d401000600000058004c004400410050005200d00200"
    "d30214010000000600000058004c0044004100500052003400230402000200802000812002010024"
    "003500d40200d102080100000001000000330c010000000100000000000000d20200cd0200"
)
