"""Rich text in one cell: runs of text with their own font.

The classes follow ``openpyxl.cell.rich_text``: a ``CellRichText`` is a list
of plain strings and ``TextBlock(InlineFont(...), text)`` parts::

    CellRichText("Total: ", TextBlock(InlineFont(b=True, color="C00000"), "1 234"))

``str(rich)`` is the plain text. XlsxWriter, the workbook model and the
XLSX updater write it as ``<r>`` runs; XlsbWriter as a rich shared string
(``BrtSSTItem`` with ``StrRun`` font runs).
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from ._text import check_cell_text, xstring_escape, xstring_unescape
from .styles import Color, _coerce_color, _color_key

_UNDERLINES = frozenset({"single", "double", "singleAccounting", "doubleAccounting"})
_VERT_ALIGNS = frozenset({"superscript", "subscript", "baseline"})


@dataclass(frozen=True)
class InlineFont:
    """The font of one run, with openpyxl's names: ``b`` bold, ``i`` italic,
    ``u`` underline (``"single"``, ``"double"``, ...), ``strike``, ``sz``
    size in points, ``color`` (``"RRGGBB"`` or a ``Color``), ``rFont`` font
    name, ``vertAlign`` (``"superscript"``/``"subscript"``)."""

    rFont: str | None = None
    b: bool | None = None
    i: bool | None = None
    u: str | None = None
    strike: bool | None = None
    sz: float | None = None
    color: Color | str | None = None
    vertAlign: str | None = None
    family: int | None = None
    charset: int | None = None
    scheme: str | None = None
    outline: bool | None = None
    shadow: bool | None = None
    condense: bool | None = None
    extend: bool | None = None

    def __post_init__(self):
        if self.u is not None and self.u not in _UNDERLINES:
            raise ValueError(f"u must be one of {sorted(_UNDERLINES)}, got {self.u!r}")
        if self.vertAlign is not None and self.vertAlign not in _VERT_ALIGNS:
            raise ValueError(f"vertAlign must be one of {sorted(_VERT_ALIGNS)}")
        if self.sz is not None and (isinstance(self.sz, bool) or not isinstance(self.sz, (int, float))
                                    or not 1 <= self.sz <= 409):
            raise ValueError("sz must be a font size between 1 and 409 points")
        # Normalized, so a font read back from a file compares equal.
        if self.sz is not None:
            object.__setattr__(self, "sz", float(self.sz))
        object.__setattr__(self, "color", _coerce_color(self.color, "color"))

    def _xml(self, prefix: str) -> str:
        parts = []
        for flag in ("b", "i", "strike", "condense", "extend", "outline", "shadow"):
            if getattr(self, flag):
                parts.append(f"<{prefix}{flag}/>")
        if self.u:
            parts.append(f'<{prefix}u/>' if self.u == "single" else f'<{prefix}u val="{self.u}"/>')
        if self.vertAlign:
            parts.append(f'<{prefix}vertAlign val="{self.vertAlign}"/>')
        if self.sz is not None:
            parts.append(f'<{prefix}sz val="{self.sz:g}"/>')
        key = _color_key(self.color)
        if key is not None:
            parts.append(f"<{prefix}color {_color_attrs(key)}/>")
        if self.rFont:
            parts.append(f'<{prefix}rFont val={_quote(self.rFont)}/>')
        if self.family is not None:
            parts.append(f'<{prefix}family val="{int(self.family)}"/>')
        if self.charset is not None:
            parts.append(f'<{prefix}charset val="{int(self.charset)}"/>')
        if self.scheme:
            parts.append(f'<{prefix}scheme val={_quote(self.scheme)}/>')
        return f"<{prefix}rPr>{''.join(parts)}</{prefix}rPr>" if parts else ""


@dataclass(frozen=True)
class TextBlock:
    """A run of ``text`` in ``font``."""

    font: InlineFont
    text: str

    def __post_init__(self):
        if not isinstance(self.font, InlineFont):
            raise TypeError("TextBlock font must be an InlineFont")
        if not isinstance(self.text, str):
            raise TypeError("TextBlock text must be a str")

    def __str__(self):
        return self.text


class CellRichText(list):
    """A cell value made of plain strings and ``TextBlock`` runs."""

    def __init__(self, *args):
        if len(args) == 1 and isinstance(args[0], (list, tuple)):
            args = args[0]
        super().__init__()
        for part in args:
            self.append(part)

    def append(self, part) -> None:
        if not isinstance(part, (str, TextBlock)):
            raise TypeError("CellRichText parts must be str or TextBlock")
        super().append(part)

    def extend(self, parts) -> None:
        for part in parts:
            self.append(part)

    def __str__(self):
        return "".join(str(part) for part in self)

    def __repr__(self):
        return f"CellRichText({list.__repr__(self)})"

    def as_list(self) -> list[str]:
        return [str(part) for part in self]

    def runs_xml(self, prefix: str = "") -> str:
        """The ``<r>`` runs for an ``<si>`` or ``<is>`` element."""
        check_cell_text(str(self))
        runs = []
        for part in self:
            if isinstance(part, TextBlock):
                font, text = part.font._xml(prefix), part.text
            else:
                font, text = "", part
            if not text:
                continue
            space = ' xml:space="preserve"' if text != text.strip() else ""
            runs.append(
                f"<{prefix}r>{font}<{prefix}t{space}>{xstring_escape(escape(text))}</{prefix}t></{prefix}r>"
            )
        return "".join(runs) or f"<{prefix}t/>"


def _quote(value: str) -> str:
    return '"' + escape(value, {'"': "&quot;"}) + '"'


def _color_attrs(key) -> str:
    kind = key[0]
    if kind == "rgb":
        return f'rgb="{key[1]}"'
    if kind == "indexed":
        return f'indexed="{key[1]}"'
    if kind == "auto":
        return 'auto="1"'
    tint = f' tint="{key[2]}"' if key[2] is not None else ""
    return f'theme="{key[1]}"{tint}'


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _font_from_rpr(rpr) -> InlineFont:
    values = {}
    for child in rpr:
        name = _local(child.tag)
        val = child.get("val")
        if name in ("b", "i", "strike", "outline", "shadow", "condense", "extend"):
            values[name] = val is None or val.lower() not in {"0", "false"}
        elif name == "u":
            values["u"] = val or "single"
        elif name == "sz" and val:
            values["sz"] = float(val)
        elif name == "rFont" and val:
            values["rFont"] = val
        elif name == "vertAlign" and val in _VERT_ALIGNS:
            values["vertAlign"] = val
        elif name in ("family", "charset") and val:
            values[name] = int(val)
        elif name == "scheme" and val:
            values["scheme"] = val
        elif name == "color":
            if child.get("rgb"):
                values["color"] = Color(rgb=child.get("rgb"))
            elif child.get("theme"):
                tint = child.get("tint")
                values["color"] = Color(theme=int(child.get("theme")),
                                        tint=float(tint) if tint else None)
            elif child.get("indexed"):
                values["color"] = Color(indexed=int(child.get("indexed")))
    return InlineFont(**values)


def rich_text_from_si(element) -> str | CellRichText:
    """The value of an ``<si>`` or ``<is>`` element: plain text, or a
    ``CellRichText`` when it has formatted runs. Phonetic runs are skipped."""
    runs = [child for child in element if _local(child.tag) == "r"]
    if not runs:
        return "".join(xstring_unescape(child.text or "") for child in element if _local(child.tag) == "t")
    parts = CellRichText()
    for run in runs:
        text = "".join(xstring_unescape(t.text or "") for t in run if _local(t.tag) == "t")
        rpr = next((child for child in run if _local(child.tag) == "rPr"), None)
        parts.append(TextBlock(_font_from_rpr(rpr), text) if rpr is not None and len(rpr) else text)
    return parts


def parse_rich_shared_strings(xml: bytes) -> list[str | CellRichText]:
    root = ET.fromstring(xml)
    return [rich_text_from_si(si) for si in root if _local(si.tag) == "si"]
