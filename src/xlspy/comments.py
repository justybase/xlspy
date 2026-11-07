"""Classic Excel cell notes, including their VML display shapes."""

from __future__ import annotations

from collections.abc import Callable, Collection, Mapping
from typing import Any
from dataclasses import dataclass
import math
import posixpath
import re
from xml.etree import ElementTree as ET
from ._safe_xml import fromstring as safe_fromstring
from xml.sax.saxutils import escape

from ._metadata import MAIN, PKG_REL, REL, cell_ref, xml_text, append_child, element_pattern, namespace_of
from ._text import check_cell_text, xstring_escape, xstring_unescape
from .refs import RangeIndex, parse_cell_range, column_label

COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
VML_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
X = "urn:schemas-microsoft-com:office:excel"


@dataclass(frozen=True, slots=True)
class Comment:
    """A classic cell note. Width and height are pixels at 96 DPI.

    Examples:
        Notes are keyed by cell address in the ``comments`` sheet option:

        ```python
        from xlspy import Comment, XlsxWriter

        with XlsxWriter("notes.xlsx") as workbook:
            workbook.add_sheet("Data")
            workbook.write_sheet(
                [["Amount"], [1200]],
                comments={"A2": Comment("Confirm this amount", "Jan", 240, 100)},
            )
        ```

        The author is required and is shown in Excel's note header.

    See Also:
        `XlsxUpdater.set_comments` for notes on an existing workbook.
    """

    text: str
    author: str
    width: float = 144
    height: float = 79

    def __post_init__(self):
        xml_text(self.text)
        check_cell_text(self.text)
        xml_text(self.author)
        for name in ("width", "height"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise TypeError(f"{name} must be a positive finite number")
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"{name} must be a positive finite number")


def normalize_comments(comments, *, allow_none=False, merges=()):
    if comments is None:
        return {}
    if not isinstance(comments, Mapping):
        raise TypeError("comments must map cell references to Comment objects")
    result: dict[str, Comment | None] = {}
    merged = None
    for ref, comment in comments.items():
        ref = cell_ref(ref)
        if ref in result:
            raise ValueError(f"duplicate comment reference: {ref}")
        if not isinstance(comment, Comment) and not (allow_none and comment is None):
            raise TypeError("comment values must be Comment objects" + (" or None" if allow_none else ""))
        if comment is not None:
            c, r, _, _ = parse_cell_range(ref)
            if merged is None:
                merged = RangeIndex()
                for merge in merges:
                    merged.add(parse_cell_range(merge))
            for c1, r1, _, _, _ in merged.overlapping((c, r, c, r)):
                if (c, r) != (c1, r1):
                    raise ValueError("comments on merged cells must use the top-left cell")
        result[ref] = comment
    return result


def comments_xml(comments, existing=None):
    """Patch only targeted comments; retain rich text in untouched comments."""
    if existing is None:
        xml = f'<comments xmlns="{MAIN}"><authors/><commentList/></comments>'
    else:
        xml = existing.decode("utf-8") if isinstance(existing, bytes) else existing
    root = safe_fromstring(xml)
    # An author's index is its first place in the list.
    authors: dict[str, int] = {}
    author_nodes = root.findall(f"{{{MAIN}}}authors/{{{MAIN}}}author")
    for index, node in enumerate(author_nodes):
        authors.setdefault(node.text or "", index)
    count = len(author_nodes)
    new_authors = []
    for comment in comments.values():
        if comment is not None and xstring_escape(comment.author) not in authors:
            author = xstring_escape(comment.author)
            authors[author] = count
            count += 1
            new_authors.append(f'<author xmlns="{MAIN}">{escape(author)}</author>')
    if new_authors:
        xml = append_child(xml, "authors", "".join(new_authors))
    seen = set()

    def patch(match):
        # Parse inside the original namespace declarations, including prefixes.
        node = _fragment_node(xml, match[0])
        ref = cell_ref(node.attrib["ref"])
        if ref not in comments:
            return match[0]
        seen.add(ref)
        return _comment_xml(ref, comments[ref], authors)

    xml = element_pattern("comment").sub(patch, xml)
    added = [_comment_xml(ref, comment, authors) for ref, comment in comments.items()
             if ref not in seen and comment is not None]
    if added:
        xml = append_child(xml, "commentList", "".join(added))
    return xml.encode("utf-8")


def _comment_xml(ref, comment, authors):
    if comment is None:
        return ""
    index = authors[xstring_escape(comment.author)]
    return (f'<comment xmlns="{MAIN}" ref="{ref}" authorId="{index}"><text>'
            f'<t xml:space="preserve">{escape(xstring_escape(comment.text))}</t></text></comment>')


def read_comments(data: bytes) -> dict[str, Comment]:
    root = safe_fromstring(data)
    ns = namespace_of(root.tag) or MAIN  # Transitional or Strict
    authors = [xstring_unescape(n.text or "") for n in root.findall(f"{{{ns}}}authors/{{{ns}}}author")]
    result = {}
    for node in root.findall(f"{{{ns}}}commentList/{{{ns}}}comment"):
        text = "".join(n.text or "" for n in node.findall(f"{{{ns}}}text//{{{ns}}}t"))
        result[cell_ref(node.attrib["ref"])] = Comment(xstring_unescape(text), authors[int(node.attrib["authorId"])])
    return result


def _fragment_node(xml, fragment):
    match = re.search(r"<(?![?!])[^>]+>", xml)
    if match is None:
        raise ValueError("XML element was not found")
    root = match[0]
    declarations = " ".join(re.findall(r'xmlns(?::[\w.-]+)?\s*=\s*(?:"[^"]*"|\x27[^\x27]*\x27)', root))
    return safe_fromstring(f"<wrapper {declarations}>{fragment}</wrapper>")[0]


def vml_xml(comments, existing=None):
    xml = (f'<xml xmlns:v="{V}" xmlns:o="{O}" xmlns:x="{X}"></xml>'
           if existing is None else existing.decode("utf-8"))
    root = safe_fromstring(xml)  # Validate before editing any package members.
    shape_ids = {n.get("id") for n in root.iter() if n.get("id")}
    type_id = "_xlspy_t202"
    while type_id in shape_ids:
        type_id += "_"
    seen = set()

    def patch(match):
        shape = _fragment_node(xml, match[0])
        client = shape.find(f"{{{X}}}ClientData")
        if client is None or client.get("ObjectType") != "Note":
            return match[0]
        row, col = client.find(f"{{{X}}}Row"), client.find(f"{{{X}}}Column")
        if row is None or col is None or row.text is None or col.text is None:
            raise ValueError("unsupported VML note without Row/Column; no changes were made")
        ref = f"{column_label(int(col.text))}{int(row.text) + 1}"
        if ref not in comments:
            return match[0]
        if ref in seen:
            raise ValueError("duplicate VML notes; no changes were made")
        seen.add(ref)
        return ""

    xml = element_pattern("shape").sub(patch, xml)
    additions = [(ref, comment) for ref, comment in comments.items() if comment is not None]
    if additions:
        # Self-contained namespace declarations avoid changing existing controls.
        fragments = [
            f'<v:shapetype xmlns:v="{V}" xmlns:o="{O}" id="{type_id}" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
            '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
        ]
        next_id = 1025
        for ref, comment in additions:
            while f"_x0000_s{next_id}" in shape_ids:
                next_id += 1
            col, row, _, _ = parse_cell_range(ref)
            fragments.append(
                f'<v:shape xmlns:v="{V}" xmlns:o="{O}" xmlns:x="{X}" id="_x0000_s{next_id}" type="#{type_id}" '
                f'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:{comment.width * .75:g}pt;height:{comment.height * .75:g}pt;z-index:1;visibility:hidden" '
                'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
                '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>'
                f'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill>'
                f'<x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>'
            )
            next_id += 1
        xml = append_child(xml, "xml", "".join(fragments))
    return xml.encode("utf-8")


def rels_path(part):
    directory, name = posixpath.split(part)
    return f"{directory}/_rels/{name}.rels"


def resolve_part(part, target):
    resolved = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith("/") else target.lstrip("/")
    if resolved.startswith("../") or resolved == "..":
        raise ValueError("relationship target escapes package")
    return resolved


_ANY_VML = object()


def find_comment_parts(get: Callable[[str], bytes], names: Collection[str], sheet_part: str,
                       legacy_rid: object = _ANY_VML) -> tuple[str, ET.Element, dict[str, tuple[str, str]]]:
    """Return the sheet's comments part and its note VML part.

    Editors pass ``legacy_rid`` (the legacyDrawing r:id, or None when the sheet
    has none): other vmlDrawing relationships, such as header/footer pictures
    used by legacyDrawingHF, are not note drawings and are never selected.
    Without it, a sole vmlDrawing relationship is reported.
    """
    path = rels_path(sheet_part)
    root = safe_fromstring(get(path)) if path in names else ET.Element(f"{{{PKG_REL}}}Relationships")
    parts: dict[str, tuple[str, str]] = {}
    drawings = []
    for rel in root:
        kind = rel.get("Type", "").rsplit("/", 1)[-1]
        if kind == "vmlDrawing":
            drawings.append(rel)
            if legacy_rid is _ANY_VML or rel.get("Id") != legacy_rid:
                continue
        elif kind != "comments":
            continue
        if kind in parts or rel.get("TargetMode") == "External":
            raise ValueError("unsupported comments/VML relationships")
        parts[kind] = (resolve_part(sheet_part, rel.attrib["Target"]), rel.attrib["Id"])
    if legacy_rid is _ANY_VML and len(drawings) == 1 and drawings[0].get("TargetMode") != "External":
        parts["vmlDrawing"] = (resolve_part(sheet_part, drawings[0].attrib["Target"]), drawings[0].attrib["Id"])
    return path, root, parts


class NotePartsEdit:
    """The comments and note-VML parts of one sheet, while its notes are edited.

    ``ensure`` gives a missing part a free name and a free relationship id;
    ``drop`` removes a part with its relationship and content-type override.
    ``parts`` and ``relationships`` (from ``find_comment_parts``) are updated
    in place.
    """

    def __init__(self, sheet_path: str, relationships: ET.Element,
                 parts: dict[str, tuple[str, str]], names: Collection[str]) -> None:
        self.sheet_path = sheet_path
        self.relationships = relationships
        self.parts = parts
        self._occupied = set(names)
        self._used_ids = {rel.get("Id") for rel in relationships}

    def ensure(self, kind: str, directory: str, stem: str, extension: str) -> str:
        """The path of the ``kind`` part, created (with a relationship) when missing."""
        if kind in self.parts:
            return self.parts[kind][0]
        index = 1
        while f"{directory}/{stem}{index}.{extension}" in self._occupied:
            index += 1
        part = f"{directory}/{stem}{index}.{extension}"
        self._occupied.add(part)
        index = 1
        while f"rId{index}" in self._used_ids:
            index += 1
        rid = f"rId{index}"
        self._used_ids.add(rid)
        ET.SubElement(self.relationships, f"{{{PKG_REL}}}Relationship", {
            "Id": rid, "Type": REL + "/" + kind,
            "Target": posixpath.relpath(part, posixpath.dirname(self.sheet_path)),
        })
        self.parts[kind] = (part, rid)
        return part

    def drop(self, part: str, kind: str, types: ET.Element, updates: dict[str, Any], removed: list[str]) -> None:
        """Remove ``part`` from the edit: no update, its relationship and override gone."""
        updates.pop(part, None)
        removed.append(part)
        # Only this part's relationship: other vmlDrawing parts may exist.
        for node in list(self.relationships):
            if node.get("Id") == self.parts[kind][1]:
                self.relationships.remove(node)
        for node in list(types):
            if node.get("PartName") == "/" + part:
                types.remove(node)


def comment_part_targets(get, names, sheet_part):
    """All comment and VML parts a worksheet refers to."""
    path = rels_path(sheet_part)
    if path not in names:
        return set()
    return {
        resolve_part(sheet_part, rel.attrib["Target"]) for rel in safe_fromstring(get(path))
        if rel.get("Type", "").rsplit("/", 1)[-1] in {"comments", "vmlDrawing"}
        and rel.get("TargetMode") != "External"
    }
