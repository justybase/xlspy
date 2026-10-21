"""BIFF12 metadata codecs. Record IDs are the decoded 7-bit BIFF12 IDs.

Layouts follow MS-XLSB and are cross-checked against Excel-produced workbooks.
"""

from __future__ import annotations

import struct
import uuid

from ._metadata import password_hash
from .biff_utils import build_record, iter_records, read_utf16
from .comments import Comment
from .refs import parse_cell_range, column_label

COMMENTS_TYPE = "application/vnd.ms-excel.comments"
BEGIN_COMMENTS, END_COMMENTS = 0x274, 0x275
BEGIN_AUTHORS, END_AUTHORS, AUTHOR = 0x276, 0x277, 0x278
BEGIN_LIST, END_LIST, BEGIN_COMMENT, END_COMMENT, TEXT = range(0x279, 0x27E)
SHEET_PROTECTION, BOOK_PROTECTION = 0x217, 0x216
SHEET_PROTECTION_ISO, BOOK_PROTECTION_ISO = 0x2A6, 0x2A5
LEGACY_DRAWING = 0x227


def wide_string(text):
    encoded = text.encode("utf-16le")
    return struct.pack("<I", len(encoded) // 2) + encoded


def read_wide(data, offset=0, *, nullable=False):
    if offset + 4 > len(data):
        raise ValueError("Truncated BIFF12 string length")
    length = struct.unpack_from("<I", data, offset)[0]
    if nullable and length == 0xFFFFFFFF:
        return None, offset + 4
    return read_utf16(data, offset + 4, length)


def records(data):
    for record in iter_records(data):
        yield record, data[record.data_start:record.data_end]


def replace_records(data, record_ids, replacement=b"", *, before=(0x82,)):
    """Preserve every untargeted binary record, inserting once in schema order."""
    chunks = []
    inserted = False
    for record, payload in records(data):
        if record.record_id in record_ids:
            if not inserted:
                chunks.append(replacement)
                inserted = True
        else:
            if not inserted and record.record_id in before:
                chunks.append(replacement)
                inserted = True
            chunks.append(data[record.header_start:record.data_end])
    if not inserted:
        raise ValueError("BIFF12 insertion point was not found")
    return b"".join(chunks)


def parse_comments(data):
    """Return notes plus their original binary blocks, retaining rich text."""
    authors = []
    notes = {}
    blocks = {}
    current = None
    start = None
    pending_start = None
    for record, payload in records(data):
        kind = record.record_id
        if kind == AUTHOR:
            authors.append(read_wide(payload)[0])
        elif kind == 0x25 and current is None:
            pending_start = record.header_start
        elif kind == BEGIN_COMMENT:
            if len(payload) != 36 or current is not None:
                raise ValueError("Malformed XLSB comment header")
            author, r1, r2, c1, c2 = struct.unpack_from("<iIIII", payload)
            if r1 != r2 or c1 != c2 or not 0 <= author < len(authors):
                raise ValueError("Invalid XLSB comment range or author")
            ref = f"{column_label(c1)}{r1 + 1}"
            parse_cell_range(ref)
            if ref in notes:
                raise ValueError("Duplicate XLSB comment")
            current = (ref, authors[author], "")
            start = pending_start if pending_start is not None else record.header_start
            pending_start = None
        elif kind == TEXT and current is not None:
            if not payload:
                raise ValueError("Truncated XLSB rich comment text")
            text, _ = read_wide(payload, 1)
            current = (current[0], current[1], text)
        elif kind == END_COMMENT:
            if current is None:
                raise ValueError("XLSB comment end without header")
            ref, author, text = current
            notes[ref] = Comment(text, author)
            blocks[ref] = (start, data[start:record.data_end])
            current = None
    if current is not None:
        raise ValueError("Unterminated XLSB comment")
    return notes, authors, blocks


def comments_bin(changes, existing=None):
    if existing is None:
        existing = b"".join(build_record(kind) for kind in
                            (BEGIN_COMMENTS, BEGIN_AUTHORS, END_AUTHORS, BEGIN_LIST, END_LIST, END_COMMENTS))
    notes, authors, blocks = parse_comments(existing)
    additions = []
    for comment in changes.values():
        if comment is not None and comment.author not in authors:
            authors.append(comment.author)
            additions.append(build_record(AUTHOR, wide_string(comment.author)))
    replacement_blocks = {}
    for ref, comment in changes.items():
        if comment is None:
            replacement_blocks[ref] = b""
            continue
        col, row, _, _ = parse_cell_range(ref)
        header = struct.pack("<iIIII", authors.index(comment.author), row, row, col, col) + uuid.uuid4().bytes_le
        # RichStr: no formatting runs and no phonetic information.
        replacement_blocks[ref] = (build_record(BEGIN_COMMENT, header)
                                   + build_record(TEXT, b"\0" + wide_string(comment.text))
                                   + build_record(END_COMMENT))
    output = []
    skip_end = -1
    block_starts = {start: (ref, block) for ref, (start, block) in blocks.items()}
    for record, payload in records(existing):
        if record.header_start < skip_end:
            continue
        if record.header_start in block_starts:
            ref, block = block_starts[record.header_start]
            if ref in replacement_blocks:
                output.append(replacement_blocks.pop(ref))
                skip_end = record.header_start + len(block)
                continue
        if record.record_id == END_AUTHORS:
            output.extend(additions)
        if record.record_id == END_LIST:
            output.extend(replacement_blocks.values())
        output.append(existing[record.header_start:record.data_end])
    return b"".join(output)


def sheet_protection_bin(protection):
    if not protection.enabled:
        return b""
    flags = [1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
             int(protection.select_locked_cells), 0, 0, 0, int(protection.select_unlocked_cells)]
    password = 0 if protection.password is None else int(password_hash(protection.password), 16)
    return build_record(SHEET_PROTECTION, struct.pack("<H16I", password, *flags))


def workbook_protection_bin(protection):
    if not protection.lock_structure:
        return b""
    password = 0 if protection.password is None else int(password_hash(protection.password), 16)
    return build_record(BOOK_PROTECTION, struct.pack("<HHH", password, 0, 1))


def merge_ranges(data):
    return [struct.unpack("<IIII", payload) for record, payload in records(data) if record.record_id == 0xB0]


def check_merged_refs(data, refs):
    for r1, r2, c1, c2 in merge_ranges(data):
        for ref in refs:
            col, row, _, _ = parse_cell_range(ref)
            if c1 <= col <= c2 and r1 <= row <= r2 and (col, row) != (c1, r1):
                raise ValueError("merged cells must use their top-left cell")
