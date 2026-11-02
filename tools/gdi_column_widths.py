"""Generate src/xlspy/_calibri_metrics.py: Calibri advance widths as GDI draws them.

Excel's AutoFit measures a cell's text with GDI at the display's resolution.
Column widths are defined at 96 DPI (the digit 0 of Calibri 11 is 7 pixels),
so the estimate uses GDI's pixel advances at 96 DPI: the font's design
advance scaled to the pixel size and rounded, plus the +-1 pixel the font's
hinting moves some glyphs at some sizes. Those exceptions are few (645 over
8-64 pixels, regular and bold), so the module stores the design advances and
the exceptions, and _widths.py rebuilds any size's pixel widths exactly.

Measuring through Excel only works on a display at 100% scaling: at other
scalings Excel measures text at that size, where hinting rounds differently.

    python tools/gdi_column_widths.py         rewrites src/xlspy/_calibri_metrics.py
    python tools/gdi_column_widths.py --check  fails if the module is out of date

Needs Windows (GDI) with the Calibri font.
"""

import ctypes
import os
import sys
from ctypes import wintypes

FONT = "Calibri"
UNITS = 2048  # Calibri's units per em
PIXEL_SIZES = range(8, 65)  # font heights in pixels (6 to 48 pt at 96 DPI)
CHARS = "".join(
    [chr(c) for c in range(0x20, 0x7F)]
    + [chr(c) for c in range(0xA0, 0x250)]
    + [chr(c) for c in range(0x370, 0x400)]
    + [chr(c) for c in range(0x400, 0x530)]
    + list("–—‘’‚“”„†‡•…‰‹›€₴₽™←↑→↓√∞≈≠≤≥")
)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "xlspy", "_calibri_metrics.py")

gdi = ctypes.windll.gdi32 if sys.platform == "win32" else None


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


def measurer(height, bold):
    gdi.CreateFontW.restype = wintypes.HFONT
    gdi.CreateCompatibleDC.restype = wintypes.HDC
    gdi.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi.GetTextExtentPoint32W.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.c_void_p]
    gdi.GetGlyphIndicesW.argtypes = [wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
    dc = gdi.CreateCompatibleDC(None)
    # A negative height is the em size in pixels; 5 = CLEARTYPE_QUALITY.
    font = gdi.CreateFontW(-height, 0, 0, 0, 700 if bold else 400, 0, 0, 0, 1, 0, 0, 5, 0, FONT)
    gdi.SelectObject(dc, font)

    def text_px(text):
        size = SIZE()
        gdi.GetTextExtentPoint32W(dc, text, len(text), ctypes.byref(size))
        return size.cx

    def has_glyph(char):
        index = ctypes.c_ushort()
        gdi.GetGlyphIndicesW(dc, char, 1, ctypes.byref(index), 1)  # GGI_MARK_NONEXISTING_GLYPHS
        return index.value != 0xFFFF

    return text_px, has_glyph


def scaled(units, pixels):
    """Design advance at a pixel size, rounded half up (as _widths.py does)."""
    return (units * pixels * 2 + UNITS) // (2 * UNITS)


def generate():
    design_px, has_glyph = measurer(UNITS, False)
    chars = "".join(c for c in CHARS if has_glyph(c) and c not in "\xad")
    tables, exceptions = {}, {}
    for bold in (False, True):
        design = measurer(UNITS, bold)[0]
        units = [design(c) for c in chars]
        tables[bold] = units
        for pixels in PIXEL_SIZES:
            text_px = measurer(pixels, bold)[0]
            plus, minus = [], []
            for char, advance in zip(chars, units):
                delta = text_px(char) - scaled(advance, pixels)
                if delta == 1:
                    plus.append(char)
                elif delta == -1:
                    minus.append(char)
                elif delta:
                    raise SystemExit(f"{char!r} at {pixels}px is {delta} pixels off; extend the format")
            if plus or minus:
                exceptions[(bold, pixels)] = ("".join(plus), "".join(minus))
            # GDI does not kern: a line is the sum of its characters.
            sample = "The quick brown fox, Zażółć gęślą jaźń 1234567890 WAVE To"
            assert text_px(sample) == sum(text_px(c) for c in sample)
    assert measurer(15, False)[0]("0") == 7, "Calibri 11 at 96 DPI has a 7-pixel digit"

    lines = [
        '"""Calibri advance widths for the column-width estimate (generated; do not edit).',
        "",
        "Written by tools/gdi_column_widths.py from GDI on Windows: design advances in",
        "font units (2048 per em) and, per weight and pixel size, the characters GDI's",
        'hinting draws one pixel wider (+1) or narrower (-1) than the rounded advance.',
        '"""',
        "",
        "from __future__ import annotations",
        "",
        f"UNITS = {UNITS}",
        f"PIXEL_SIZES = range({PIXEL_SIZES.start}, {PIXEL_SIZES.stop})",
        f"CHARS = {chars!r}",
    ]
    for name, bold in (("REGULAR", False), ("BOLD", True)):
        lines.append(f"{name} = (")
        units = tables[bold]
        for i in range(0, len(units), 16):
            lines.append("    " + ", ".join(str(u) for u in units[i:i + 16]) + ",")
        lines.append(")")
    lines.append("# (bold, pixels): (one pixel wider, one pixel narrower)")
    lines.append("HINTING = {")
    for (bold, pixels), (plus, minus) in sorted(exceptions.items()):
        lines.append(f"    ({bold}, {pixels}): ({plus!r}, {minus!r}),")
    lines.append("}")
    return "\n".join(lines) + "\n"


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if gdi is None:
        sys.exit("needs Windows (GDI) with the Calibri font")
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
    print(f"wrote {path}")


if __name__ == "__main__":
    main()
