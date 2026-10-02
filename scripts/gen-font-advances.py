"""Generate src/GanttCreator.Core/Scene/FontAdvanceTable.cs from the real font binaries.

The advance widths are read from each family's TrueType tables -- `head.unitsPerEm`,
`hmtx` for advances, `cmap` for the Unicode-to-glyph map -- so the committed table
is the font's own metric rather than an approximation.

Why this exists: the previous table was a flat 4pt per character, which is wrong in
BOTH directions. Measured against a live Excel textbox at 8pt, 40x'W' is 287.03pt
(the flat table said 160pt, 44% too small) and 40x'i' is 78.03pt (the flat table
said 160pt, 105% too large). The small end is the one that clips: the scene sized
the label box from the underestimate and the host has no auto-size to recover.

Office cloud fonts are NOT in C:\\Windows\\Fonts, NOT in the font registry, and NOT
enumerated by System.Drawing's InstalledFontCollection. They live under
%LOCALAPPDATA%\\Microsoft\\FontCache\\4\\CloudFonts\\<Family>. Searching any of the
other three locations reports the font as absent when it is installed -- which is
how this table was nearly built against the wrong font.

Run:  python scripts/gen-font-advances.py
"""

from __future__ import annotations

import os
import struct
from pathlib import Path

CLOUD_FONTS = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts"))
OUTPUT = Path(__file__).resolve().parent.parent / "src" / "GanttCreator.Core" / "Scene" / "FontAdvanceTable.cs"

# The two faces of each family that labels actually use. The typography tokens are
# 8pt regular (body/header), 8pt bold (delineator), 9pt bold (year) and 11pt bold
# (title), so Regular and Bold cover all of them.
FACES = {
    "Aptos": {"regular": "30153066857.ttf", "bold": "31531277363.ttf"},
    "AptosNarrow": {"regular": "27970306811.ttf", "bold": "24328775610.ttf"},
}

# Non-ASCII code points labels can contain. Zero advance in the output means the font
# genuinely has no glyph, which is a real answer rather than a missing entry.
EXTRA_CODE_POINTS = [
    0x00A0,  # no-break space
    0x2013,  # en dash
    0x2014,  # em dash
    0x2018,  # left single quote
    0x2019,  # right single quote
    0x201C,  # left double quote
    0x201D,  # right double quote
    0x2026,  # ellipsis -- the label truncation marker
    0x00D7,  # multiplication sign
    0x2022,  # bullet
]

ASCII_FIRST = 0x20
ASCII_LAST = 0x7E


def table_directory(data: bytes) -> dict[str, tuple[int, int]]:
    count = struct.unpack(">H", data[4:6])[0]
    tables: dict[str, tuple[int, int]] = {}
    for index in range(count):
        record = 12 + index * 16
        tag = data[record : record + 4].decode("latin-1").strip()
        offset, length = struct.unpack(">II", data[record + 8 : record + 16])
        tables[tag] = (offset, length)
    return tables


def unicode_to_glyph(data: bytes, tables: dict[str, tuple[int, int]]) -> dict[int, int]:
    """cmap format 4, the BMP subtable every one of these fonts carries."""
    base, _ = tables["cmap"]
    num_tables = struct.unpack(">H", data[base + 2 : base + 4])[0]
    subtable = None
    for index in range(num_tables):
        platform, encoding, offset = struct.unpack(">HHI", data[base + 4 + index * 8 : base + 12 + index * 8])
        if (platform, encoding) in ((3, 1), (3, 10), (0, 3), (0, 4)):
            subtable = base + offset
    if subtable is None:
        raise RuntimeError("no BMP cmap subtable")

    seg_x2 = struct.unpack(">H", data[subtable + 6 : subtable + 8])[0]
    segments = seg_x2 // 2
    ends = read_u16(data, subtable + 14, segments)
    start_off = subtable + 16 + seg_x2
    starts = read_u16(data, start_off, segments)
    delta_off = start_off + seg_x2
    deltas = [struct.unpack(">h", data[delta_off + i * 2 : delta_off + 2 + i * 2])[0] for i in range(segments)]
    range_off = delta_off + seg_x2
    ranges = read_u16(data, range_off, segments)

    mapping: dict[int, int] = {}
    for index in range(segments):
        for code in range(starts[index], min(ends[index], 0xFFFF) + 1):
            if ranges[index] == 0:
                glyph = (code + deltas[index]) & 0xFFFF
            else:
                at = range_off + index * 2 + ranges[index] + (code - starts[index]) * 2
                glyph = struct.unpack(">H", data[at : at + 2])[0]
                if glyph:
                    glyph = (glyph + deltas[index]) & 0xFFFF
            if glyph:
                mapping[code] = glyph
    return mapping


def read_u16(data: bytes, offset: int, count: int) -> list[int]:
    return [struct.unpack(">H", data[offset + i * 2 : offset + 2 + i * 2])[0] for i in range(count)]


def load_face(path: Path) -> tuple[int, list[int], dict[int, int]]:
    data = path.read_bytes()
    tables = table_directory(data)
    units_per_em = struct.unpack(">H", data[tables["head"][0] + 18 : tables["head"][0] + 20])[0]
    num_glyphs = struct.unpack(">H", data[tables["maxp"][0] + 4 : tables["maxp"][0] + 6])[0]
    num_h_metrics = struct.unpack(">H", data[tables["hhea"][0] + 34 : tables["hhea"][0] + 36])[0]

    hmtx, _ = tables["hmtx"]
    advances = [0] * num_glyphs
    for index in range(min(num_h_metrics, num_glyphs)):
        advances[index] = struct.unpack(">H", data[hmtx + index * 4 : hmtx + index * 4 + 2])[0]
    # Trailing glyphs share the last advance, which is how the table is defined.
    for index in range(num_h_metrics, num_glyphs):
        advances[index] = advances[num_h_metrics - 1]

    return units_per_em, advances, unicode_to_glyph(data, tables)


def advance_for(units_per_em: int, advances: list[int], mapping: dict[int, int], code: int) -> int:
    glyph = mapping.get(code)
    return advances[glyph] if glyph else 0


def emit_array(name: str, values: list[int], per_line: int = 12) -> str:
    lines = [f"    internal static readonly ushort[] {name} =", "    ["]
    for index in range(0, len(values), per_line):
        chunk = ", ".join(str(value) for value in values[index : index + per_line])
        lines.append(f"        {chunk},")
    lines.append("    ];")
    return "\n".join(lines)


def build() -> str:
    loaded = {}
    for family, faces in FACES.items():
        for weight, filename in faces.items():
            path = CLOUD_FONTS / family.replace("AptosNarrow", "Aptos Narrow") / filename
            if not path.exists():
                raise FileNotFoundError(f"missing {path}")
            loaded[(family, weight)] = load_face(path)

    units = {units for units, _, _ in loaded.values()}
    if len(units) != 1:
        raise RuntimeError(f"families disagree on unitsPerEm: {units}")
    units_per_em = units.pop()

    parts: list[str] = []
    header = """// <auto-generated>
//   Advance widths for the Aptos and Aptos Narrow families, read from the real
//   font binaries. DO NOT EDIT BY HAND -- regenerate with
//   scripts/gen-font-advances.py, which is committed for exactly that reason.
// </auto-generated>
#nullable enable

namespace GanttCreator.Core.Scene;

/// <summary>
/// Glyph advance widths, in font design units, for the families the label
/// typography tokens use.
/// </summary>
/// <remarks>
/// <para>
/// These are the fonts' own <c>hmtx</c> advances, not an approximation. The
/// previous table was a flat 4pt per character, which is wrong in BOTH
/// directions: measured against a live Excel textbox at 8pt, 40 characters of
/// <c>W</c> is 287.03pt where the flat table said 160pt (44 percent too small),
/// and 40 characters of <c>i</c> is 78.03pt where it said 160pt (105 percent
/// too large). The small end is what clips a label: the scene sized the box from
/// the underestimate and the host has no auto-size to recover the difference.
/// </para>
/// <para>
/// Regular and Bold of both families are carried, which covers the typography
/// tokens at 8pt regular, 8pt bold, 9pt bold and 11pt bold.
/// </para>
/// </remarks>
internal static class FontAdvanceTable
{
    /// <summary>The shared design-unit denominator.</summary>
    internal const int UnitsPerEm = __UPEM__;

    /// <summary>The first code point in the ASCII run, U+0020.</summary>
    internal const int AsciiFirst = 0x__FIRST__;

    /// <summary>One past the last code point in the ASCII run, U+007E.</summary>
    internal const int AsciiLast = 0x__LAST__;
"""
    # Token replacement rather than %-formatting: the prose above is full of
    # literal percent signs, and %-formatting would read them as format specifiers.
    header = header.replace("__UPEM__", str(units_per_em))
    header = header.replace("__FIRST__", f"{ASCII_FIRST:02X}")
    header = header.replace("__LAST__", f"{ASCII_LAST:02X}")
    parts.append(header)

    for family, stem in (("Aptos", "Aptos"), ("AptosNarrow", "AptosNarrow")):
        for weight, suffix in (("regular", "Regular"), ("bold", "Bold")):
            units, advances, mapping = loaded[(family, weight)]
            values = [advance_for(units, advances, mapping, code) for code in range(ASCII_FIRST, ASCII_LAST + 1)]
            parts.append("\n" + emit_array(f"{stem}{suffix}Ascii", values))

    parts.append(
        """

    /// <summary>
    /// Code points carried outside ASCII, in the order the <c>*Extra</c> arrays
    /// use: no-break space, en dash, em dash, left and right single quote, left
    /// and right double quote, ellipsis, multiplication sign, bullet.
    /// </summary>
    /// <remarks>A zero means the font has no glyph, which is a real answer.</remarks>
    internal static readonly int[] ExtraCodePoints =
    ["""
    )
    for index in range(0, len(EXTRA_CODE_POINTS), 6):
        chunk = ", ".join(f"0x{code:04X}" for code in EXTRA_CODE_POINTS[index : index + 6])
        parts.append(f"        {chunk},")
    parts.append("    ];")

    for family, stem in (("Aptos", "Aptos"), ("AptosNarrow", "AptosNarrow")):
        for weight, suffix in (("regular", "Regular"), ("bold", "Bold")):
            units, advances, mapping = loaded[(family, weight)]
            values = [advance_for(units, advances, mapping, code) for code in EXTRA_CODE_POINTS]
            parts.append("\n" + emit_array(f"{stem}{suffix}Extra", values))

    parts.append(
        """

    /// <summary>
    /// The advance for one code point, or 0 when the font carries no glyph.
    /// </summary>
    /// <param name="codePoint">The Unicode code point.</param>
    /// <param name="bold">Whether to read the bold face.</param>
    /// <param name="narrow">Whether to read the Narrow family.</param>
    /// <returns>The advance in design units, or 0 when the glyph is absent.</returns>
    internal static int AdvanceFor(int codePoint, bool bold, bool narrow)
    {
        ushort[] ascii = (narrow, bold) switch
        {
            (false, false) => AptosRegularAscii,
            (false, true) => AptosBoldAscii,
            (true, false) => AptosNarrowRegularAscii,
            (true, true) => AptosNarrowBoldAscii,
        };

        if (codePoint >= AsciiFirst && codePoint <= AsciiLast)
        {
            return ascii[codePoint - AsciiFirst];
        }

        ushort[] extra = (narrow, bold) switch
        {
            (false, false) => AptosRegularExtra,
            (false, true) => AptosBoldExtra,
            (true, false) => AptosNarrowRegularExtra,
            (true, true) => AptosNarrowBoldExtra,
        };

        for (int index = 0; index < ExtraCodePoints.Length; index++)
        {
            if (ExtraCodePoints[index] == codePoint)
            {
                return extra[index];
            }
        }

        return 0;
    }
}
"""
    )
    return "".join(parts)


def main() -> None:
    OUTPUT.write_text(build(), encoding="utf-8", newline="\n")
    print(f"wrote {OUTPUT}")


if __name__ == "__main__":
    main()