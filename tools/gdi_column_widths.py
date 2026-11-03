"""Generate src/xlspy/_font_metrics.py: advance widths of popular fonts as GDI draws them.

Excel's AutoFit measures a cell's text with GDI at the display's resolution.
Column widths are defined at 96 DPI (the digit 0 of Calibri 11 is 7 pixels),
so the estimate uses GDI's pixel advances at 96 DPI: each font's design
advance scaled to the pixel size and rounded, plus the pixels the font's
hinting moves some glyphs at some sizes. The module stores, per font and
style (regular, bold, italic, bold italic), the design advances, and per
pixel size from 8 to 64 the shift most glyphs take and the glyphs that differ
from it, so _fonts.py rebuilds any size's pixel widths exactly. It also keeps
each style's ascent per pixel size (italic text gets room for its slant from
it) and the widths of the fallback fonts Excel draws a character with when
the cell's font lacks it (emoji, symbols, some scripts).

Measuring through Excel only works on a display at 100% scaling: at other
scalings Excel measures text at that size, where hinting rounds differently.

    python tools/gdi_column_widths.py         rewrites src/xlspy/_font_metrics.py
    python tools/gdi_column_widths.py --check  fails if the module is out of date

Needs Windows (GDI) with the fonts below installed.
"""

import base64
import collections
import ctypes
import json
import os
import sys
import unicodedata
import zlib
from ctypes import wintypes

UNITS = 2048  # the design units the module stores (every font here has 2048 per em)
PIXEL_SIZES = range(8, 65)  # font heights in pixels (6 to 48 pt at 96 DPI)
FONTS = [
    "Calibri", "Calibri Light", "Aptos", "Aptos Narrow", "Arial", "Arial Narrow", "Arial Black",
    "Times New Roman", "Cambria", "Consolas", "Courier New", "Lucida Console", "Segoe UI", "Tahoma",
    "Verdana", "Georgia", "Trebuchet MS", "Century Gothic", "Lucida Sans Unicode", "Microsoft Sans Serif",
    "Comic Sans MS",
]
# Office's cloud fonts (Aptos, the default since Office 2023) are not installed
# in Windows: Office keeps them in its font cache and loads them for itself.
CLOUD_FONTS = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "FontCache", "4", "CloudFonts")
_CLOUD_FAMILIES = ["Aptos", "Aptos Narrow"]
STYLES = [(False, False), (True, False), (False, True), (True, True)]  # (bold, italic)
_BLOCKS = [
    (0x20, 0x7F), (0xA0, 0x250),      # Latin
    (0x370, 0x400), (0x400, 0x530),   # Greek, Cyrillic
    (0x590, 0x600),                   # Hebrew
    (0x1E00, 0x1F00),                 # Latin Extended Additional (Vietnamese)
    (0x2000, 0x2070),                 # General Punctuation, with the Unicode spaces
    (0x2070, 0x20D0),                 # super- and subscripts, currency signs
    (0x2100, 0x2200),                 # letterlike symbols, number forms, arrows
    (0x2200, 0x2400),                 # mathematical operators, technical symbols
    (0x2460, 0x2500),                 # enclosed alphanumerics
    (0x2500, 0x2600),                 # box drawing, blocks, geometric shapes
    (0x2600, 0x27C0),                 # miscellaneous symbols, dingbats
    (0x2B00, 0x2C00),                 # miscellaneous symbols and arrows
]
# Scripts only the fallback fonts draw (Devanagari, Thai).
_FALLBACK_BLOCKS = [(0x900, 0x980), (0xE00, 0xE80)]
# Emoji outside the BMP: Excel draws them all with Segoe UI Emoji.
_EMOJI_BLOCKS = [(0x1F000, 0x1F100), (0x1F1E6, 0x1F200), (0x1F300, 0x1F650), (0x1F680, 0x1F700),
                 (0x1F7E0, 0x1F7F0), (0x1F900, 0x1FA00), (0x1FA70, 0x1FB00)]
# BMP characters with Emoji_Presentation=Yes: Excel takes them from Segoe UI
# Emoji, the other symbols from Segoe UI Symbol.
_EMOJI_PRESENTATION = (
    "⌚⌛⏩⏪⏫⏬⏰⏳◽◾☔☕"
    + "".join(chr(c) for c in range(0x2648, 0x2654))
    + "♿⚓⚡⚪⚫⚽⚾⛄⛅⛎⛔⛪⛲⛳⛵⛺"
    "⛽✅✊✋✨❌❎❓❔❕❗➕➖➗➰➿"
    "⬛⬜⭐⭕"
)
# Excel's own font fallback picks among these by character (Cambria Math for
# some arrows, Arial- or Tahoma-wide box drawing, Segoe UI Symbol for most
# symbols); the module keeps the widest, so no fallback character is too narrow.
FALLBACK_FONTS = ["Segoe UI Symbol", "Cambria Math", "Segoe UI", "Tahoma", "Arial", "Nirmala UI", "Leelawadee UI"]
EMOJI_FONT = "Segoe UI Emoji"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "xlspy", "_font_metrics.py")


def _drawn(c: int) -> bool:
    """A character that takes room: not a control, format or combining mark."""
    category = unicodedata.category(chr(c))
    return category not in ("Cc", "Cf", "Cs", "Co", "Cn", "Mn", "Me", "Zl", "Zp")


CHARS = "".join(chr(c) for lo, hi in _BLOCKS for c in range(lo, hi) if _drawn(c))
FALLBACK_CHARS = CHARS + "".join(chr(c) for lo, hi in _FALLBACK_BLOCKS for c in range(lo, hi) if _drawn(c))

gdi = ctypes.windll.gdi32 if sys.platform == "win32" else None


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


class TEXTMETRICW(ctypes.Structure):
    _fields_ = [(name, ctypes.c_long) for name in (
        "tmHeight", "tmAscent", "tmDescent", "tmInternalLeading", "tmExternalLeading", "tmAveCharWidth",
        "tmMaxCharWidth", "tmWeight", "tmOverhang", "tmDigitizedAspectX", "tmDigitizedAspectY")] + [
        ("tmFirstChar", ctypes.c_wchar), ("tmLastChar", ctypes.c_wchar), ("tmDefaultChar", ctypes.c_wchar),
        ("tmBreakChar", ctypes.c_wchar), ("tmItalic", ctypes.c_byte), ("tmUnderlined", ctypes.c_byte),
        ("tmStruckOut", ctypes.c_byte), ("tmPitchAndFamily", ctypes.c_byte), ("tmCharSet", ctypes.c_byte)]


class Gdi:
    """A memory DC with one font selected at a time."""

    def __init__(self):
        gdi.CreateFontW.restype = wintypes.HFONT
        gdi.CreateCompatibleDC.restype = wintypes.HDC
        gdi.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
        gdi.SelectObject.restype = wintypes.HGDIOBJ
        gdi.DeleteObject.argtypes = [wintypes.HGDIOBJ]
        gdi.GetTextExtentPoint32W.argtypes = [wintypes.HDC, ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
        gdi.GetGlyphIndicesW.argtypes = [wintypes.HDC, ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p,
                                         wintypes.DWORD]
        gdi.GetTextFaceW.argtypes = [wintypes.HDC, ctypes.c_int, wintypes.LPWSTR]
        gdi.GetTextMetricsW.argtypes = [wintypes.HDC, ctypes.c_void_p]
        self.dc = gdi.CreateCompatibleDC(None)

    def select(self, name: str, height: int, bold: bool = False, italic: bool = False) -> None:
        # A negative height is the em size in pixels; 5 = CLEARTYPE_QUALITY.
        font = gdi.CreateFontW(-height, 0, 0, 0, 700 if bold else 400, int(italic), 0, 0, 1, 0, 0, 5, 0, name)
        old = gdi.SelectObject(self.dc, font)
        if old:
            gdi.DeleteObject(old)

    def face(self) -> str:
        buf = ctypes.create_unicode_buffer(64)
        gdi.GetTextFaceW(self.dc, 64, buf)
        return buf.value

    def px(self, text: str) -> int:
        units = text.encode("utf-16-le")
        buf = ctypes.create_string_buffer(units)
        size = SIZE()
        gdi.GetTextExtentPoint32W(self.dc, buf, len(units) // 2, ctypes.byref(size))
        return size.cx

    def has(self, char: str) -> bool:
        index = ctypes.c_ushort()
        buf = ctypes.create_unicode_buffer(char)
        gdi.GetGlyphIndicesW(self.dc, buf, 1, ctypes.byref(index), 1)  # GGI_MARK_NONEXISTING_GLYPHS
        return index.value != 0xFFFF

    def ascent(self) -> int:
        tm = TEXTMETRICW()
        gdi.GetTextMetricsW(self.dc, ctypes.byref(tm))
        return tm.tmAscent


def scaled(units: int, pixels: int) -> int:
    """Design advance at a pixel size, rounded half up (as _fonts.py does)."""
    return (units * pixels * 2 + UNITS) // (2 * UNITS)


def _hinting(pixel_widths: dict, design: list) -> dict:
    """Per pixel size: [shift of most glyphs, {further delta: glyphs}] against ``design``."""
    hinting = {}
    for pixels, widths in pixel_widths.items():
        deltas = {c: widths[c] - scaled(d, pixels) for c, d in zip(CHARS, design) if d >= 0}
        shift = collections.Counter(deltas.values()).most_common(1)[0][0]
        groups = collections.defaultdict(str)
        for c, delta in deltas.items():
            if delta != shift:
                groups[delta - shift] += c
        if shift or groups:
            hinting[pixels] = [shift, {str(k): v for k, v in sorted(groups.items())}]
    return hinting


def _style(g: Gdi, name: str, chars: str, bold: bool, italic: bool, regular: list | None) -> dict:
    g.select(name, UNITS, bold, italic)
    design = [g.px(c) if c in chars else -1 for c in CHARS]
    pixel_widths, ascents = {}, []
    sample = "The quick brown fox, Zażółć gęślą jaźń 1234567890 WAVE To"
    for pixels in PIXEL_SIZES:
        g.select(name, pixels, bold, italic)
        ascents.append(g.ascent())
        pixel_widths[pixels] = {c: g.px(c) for c in chars}
        # GDI does not kern: a line is the sum of its characters.
        assert g.px(sample) == sum(g.px(c) for c in sample), (name, bold, italic, pixels)
    hinting = _hinting(pixel_widths, design)
    if regular is not None:
        # A style Windows synthesizes (Calibri Light bold) widens every glyph by
        # a pixel or two: the regular advances plus a shift describe it better.
        from_regular = _hinting(pixel_widths, regular)
        if len(json.dumps(from_regular)) < len(json.dumps(hinting)):
            design, hinting = regular, from_regular
    g.select(name, UNITS, bold, italic)
    return {"design": design, "ascent": ascents, "ascent_units": g.ascent(), "hinting": hinting}


def _pack(data) -> str:
    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode()
    return base64.b85encode(zlib.compress(raw, 9)).decode("ascii")


def _fallback(g: Gdi) -> list:
    """Runs ``[first code point, count, design units]`` of the fallback widths."""
    widths = {}
    for font in [EMOJI_FONT] + FALLBACK_FONTS:
        g.select(font, UNITS)
        if g.face() != font:
            raise SystemExit(f"fallback font {font!r} is not installed")
        for char in FALLBACK_CHARS:
            if (font == EMOJI_FONT) != (char in _EMOJI_PRESENTATION) or not g.has(char):
                continue
            widths[ord(char)] = max(widths.get(ord(char), 0), g.px(char))
    g.select(EMOJI_FONT, UNITS)
    for lo, hi in _EMOJI_BLOCKS:
        for cp in range(lo, hi):
            if unicodedata.category(chr(cp)) != "Cn":
                widths[cp] = g.px(chr(cp))
    runs = []
    for cp in sorted(widths):
        if runs and runs[-1][0] + runs[-1][1] == cp and runs[-1][2] == widths[cp]:
            runs[-1][1] += 1
        else:
            runs.append([cp, 1, widths[cp]])
    return runs


def _load_cloud_fonts() -> None:
    """Make Office's cached cloud fonts available to GDI in this process only."""
    gdi.AddFontResourceExW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_void_p]
    for family in _CLOUD_FAMILIES:
        folder = os.path.join(CLOUD_FONTS, family)
        files = [f for f in os.listdir(folder) if f.lower().endswith((".ttf", ".otf"))] if os.path.isdir(folder) else []
        if not files:
            raise SystemExit(f"{family!r} is not in Office's font cache ({folder}): use it once in Excel")
        for name in files:
            gdi.AddFontResourceExW(os.path.join(folder, name), 0x10, None)  # FR_PRIVATE


def generate() -> str:
    _load_cloud_fonts()
    g = Gdi()
    families = {}
    for name in FONTS:
        g.select(name, UNITS)
        if g.face() != name:
            raise SystemExit(f"font {name!r} is not installed (GDI gave {g.face()!r})")
        chars = "".join(c for c in CHARS if g.has(c))
        styles = {}
        for bold, italic in STYLES:
            regular = styles["00"]["design"] if styles else None
            styles[f"{int(bold)}{int(italic)}"] = _style(g, name, chars, bold, italic, regular)
        families[name.lower()] = (name, _pack(styles))
    g.select("Calibri", 15)
    assert g.px("0") == 7, "Calibri 11 at 96 DPI has a 7-pixel digit"
    fallback = _pack(_fallback(g))

    def wrapped(text: str, indent: str) -> list[str]:
        return [f'{indent}"{text[i:i + 96]}"' for i in range(0, len(text), 96)]

    lines = [
        '"""Font advance widths for the column-width estimate (generated; do not edit).',
        "",
        "Written by tools/gdi_column_widths.py from GDI on Windows. Per font, a",
        "base85 zlib JSON object keyed by style (bold, italic as '00'..'11'): design",
        f"advances in font units ({UNITS} per em, -1 where the font has no glyph) for",
        "CHARS, the ascent per pixel size, and per pixel size the shift GDI's hinting",
        "gives most glyphs with the glyphs that move further. FALLBACK holds runs of",
        "[code point, count, units] for the characters Excel draws from a fallback font.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        f"UNITS = {UNITS}",
        f"PIXEL_SIZES = range({PIXEL_SIZES.start}, {PIXEL_SIZES.stop})",
        f"CHARS = {CHARS!r}",
        "# lower-case family name: (family name, packed styles)",
        "FAMILIES = {",
    ]
    for key, (name, packed) in families.items():
        lines.append(f"    {key!r}: ({name!r}, (")
        lines.extend(wrapped(packed, "        "))
        lines.append("    )),")
    lines.append("}")
    lines.append("FALLBACK = (")
    lines.extend(wrapped(fallback, "    "))
    lines.append(")")
    return "\n".join(lines) + "\n"


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if gdi is None:
        sys.exit("needs Windows (GDI) with the fonts installed")
    text = generate()
    path = os.path.normpath(OUT)
    if "--check" in sys.argv:
        with open(path, encoding="utf-8") as handle:
            if handle.read() != text:
                sys.exit(f"{path} is out of date; run tools/gdi_column_widths.py")
        print("up to date")
        return
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)
    print(f"wrote {path} ({len(text) // 1024} KiB)")


if __name__ == "__main__":
    main()
