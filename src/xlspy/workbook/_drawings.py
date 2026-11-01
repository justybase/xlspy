"""Remove loaded charts and images from a package's drawing parts.

The loader records where each chart and image came from (drawing part and
anchor position, see ``chart_reader.read_drawing``). A save removes the
anchors of the objects no longer on the sheet and the parts only they used,
for XLSX and XLSB alike: both keep drawings as DrawingML parts. The drawing
XML is edited as text, so the prefixes and ``mc:Ignorable`` declarations of
the file stay as they are.
"""

from __future__ import annotations

from ..exceptions import UnsupportedFeatureError
import re
from xml.etree import ElementTree as ET
from .._safe_xml import fromstring as safe_fromstring

from .._metadata import CT, PKG_REL
from ._objects import _bytes, _Package, _rels_path, _resolve

_ANCHORS = frozenset(("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"))
_TAG = re.compile(r"<(/?)([\w.-]+:)?([\w.-]+)\b[^>]*?(/?)>|<!--.*?-->|<\?.*?\?>|<!\[CDATA\[.*?\]\]>", re.S)
_REFERENCE = re.compile(r'\s[\w.-]+:(?:id|embed|link)="([^"]+)"')
# Parts a chart brings with it (style and colors), removed with it.
_CHART_COMPANIONS = ("/chartStyle", "/chartColorStyle")


def _drawing_objects(sheet) -> list:
    return [*sheet._charts, *sheet._images]


def _contains(items, item) -> bool:
    return any(existing is item for existing in items)


def removed_drawing_objects(sheet) -> list[tuple[str, int, object]]:
    """The loaded charts and images no longer on the sheet."""
    current = _drawing_objects(sheet)
    return [
        entry for entry in getattr(sheet, "_drawing_sources", ())
        if not _contains(current, entry[2])
    ]


def new_drawing_objects(sheet) -> dict[str, list]:
    """The charts and images not in the file yet."""
    saved = getattr(sheet, "_saved_drawing_objects", None)
    if saved is None:
        saved = [entry[2] for entry in getattr(sheet, "_drawing_sources", ())]
    return {
        name: [item for item in getattr(sheet, name) if not _contains(saved, item)]
        for name in ("charts", "images")
    }


def remember_drawing_objects(sheet) -> None:
    sheet._saved_drawing_objects = _drawing_objects(sheet)


def check_removals(sheet) -> None:
    """Refuse removals that a save cannot make exactly."""
    saved = getattr(sheet, "_saved_drawing_objects", None)
    if saved is None:
        return
    current = _drawing_objects(sheet)
    sources = getattr(sheet, "_drawing_sources", ())
    known = [entry[2] for entry in sources]
    for item in saved:
        # Objects added and saved in this session have no recorded place in the file.
        if not _contains(current, item) and not _contains(known, item):
            raise UnsupportedFeatureError(
                "a chart or image added by an earlier save cannot be removed in the same "
                "session; load the saved workbook again to remove it"
            )
    # Removal drops a whole anchor, so a group must go as a whole.
    removed = {(part, position) for part, position, _item in removed_drawing_objects(sheet)}
    for part, position, item in sources:
        if (part, position) in removed and _contains(current, item):
            raise UnsupportedFeatureError(
                "a chart or image that is grouped with other objects can only be removed "
                "together with the rest of its group"
            )


def _top_level_anchors(xml: str) -> list[tuple[int, int]]:
    """``(start, end)`` of each anchor element directly under the root."""
    spans = []
    depth = 0
    start = None
    for match in _TAG.finditer(xml):
        if match[3] is None:  # comment, declaration or CDATA
            continue
        closing, local, empty = match[1], match[3], match[4]
        if closing:
            depth -= 1
            if depth == 1 and start is not None:
                spans.append((start, match.end()))
                start = None
            continue
        if depth == 1 and local in _ANCHORS:
            if empty:
                spans.append((match.start(), match.end()))
                continue
            start = match.start()
        if not empty:
            depth += 1
    return spans


def _referenced(package: _Package, part: str, skip: set[str]) -> bool:
    """Whether a relationship of any other part targets ``part``."""
    for name in list(package.names):
        if not name.endswith(".rels") or name in skip:
            continue
        data = package.get(name)
        if not data:
            continue
        source = name.replace("_rels/", "")[:-len(".rels")]
        try:
            root = safe_fromstring(data)
        except ET.ParseError:
            return True
        for node in root:
            if node.get("TargetMode") == "External":
                continue
            target = node.get("Target")
            if target and _resolve(source, target) == part:
                return True
    return False


def _drop_part(package: _Package, part: str) -> None:
    package.remove(part)
    if package.get(_rels_path(part)) is not None:
        package.remove(_rels_path(part))
    for node in list(package.types):
        if node.tag == f"{{{CT}}}Override" and node.get("PartName") == "/" + part:
            package.types.remove(node)


def remove_drawing_objects(updater, sheet) -> None:
    """Remove the anchors of the loaded charts and images no longer on the sheet."""
    removed = removed_drawing_objects(sheet)
    if not removed:
        return
    package = _Package(updater)
    by_part: dict[str, set[int]] = {}
    for part, position, _item in removed:
        by_part.setdefault(part, set()).add(position)
    for drawing, positions in by_part.items():
        data = package.get(drawing)
        if data is None:
            continue
        xml = data.decode("utf-8")
        spans = _top_level_anchors(xml)
        dropped = [spans[index] for index in sorted(positions) if index < len(spans)]
        references = {
            rid for start, end in dropped for rid in _REFERENCE.findall(xml[start:end])
        }
        for start, end in reversed(dropped):
            xml = xml[:start] + xml[end:]
        package.set(drawing, xml)
        still_used = set(_REFERENCE.findall(xml))
        rels_path = _rels_path(drawing)
        rels = package.rels(drawing)
        orphans = []
        for node in list(rels):
            if node.get("Id") in references and node.get("Id") not in still_used:
                rels.remove(node)
                target = node.get("Target")
                if node.get("TargetMode") != "External" and target:
                    orphans.append((node.get("Type", ""), _resolve(drawing, target)))
        ET.register_namespace("", PKG_REL)
        package.set(rels_path, _bytes(rels))
        for kind, target in orphans:
            if _referenced(package, target, set()):
                continue
            if kind.endswith("/chart"):
                companions = [
                    _resolve(target, node.get("Target", "")) for node in package.rels(target)
                    if node.get("Type", "").endswith(_CHART_COMPANIONS) and node.get("Target")
                ]
                _drop_part(package, target)
                for companion in companions:
                    if not _referenced(package, companion, set()):
                        _drop_part(package, companion)
            elif package.get(target) is not None:
                _drop_part(package, target)
    package.commit()
    # The remaining anchors moved up.
    kept = []
    for part, position, item in getattr(sheet, "_drawing_sources", ()):
        if any(entry[2] is item for entry in removed):
            continue
        shift = sum(1 for other in by_part.get(part, ()) if other < position)
        kept.append((part, position - shift, item))
    sheet._drawing_sources = kept


def save_chart_edits(updater, sheet) -> None:
    """Write the chart parts whose title, axes or legend were changed."""
    for chart in sheet._charts:
        if getattr(chart, "edited", False) and chart.part:
            updater._package.set(chart.part, chart.xml)
            chart._mark_saved()
