"""Encode BIFF12 style records independently for format import tests."""

import struct

from xlsb_builder import cell, package, record, replace_part, row_header, wide, worksheet


def color(kind="rgb", value="000000", *, tint=0, alpha=255):
    kinds = {"auto": 0, "indexed": 1, "rgb": 2, "theme": 3}
    rgb = bytes.fromhex(value) if kind == "rgb" else b"\0\0\0"
    index = value if kind in ("indexed", "theme") else 0
    return struct.pack("<BBh", kinds[kind] * 2 + 1, index, tint) + rgb + bytes([alpha])


def font(*, name="Arial", size=11, bold=False, italic=False, underline=0,
         family=2, charset=0, scheme=0, text_color=None):
    return (struct.pack("<HHHHBBBB", int(size * 20), 2 if italic else 0,
                        700 if bold else 400, 0, underline, family, charset, 0)
            + (text_color or color("theme", 1)) + bytes([scheme]) + wide(name))


def fill(*, pattern=0, fg=None, bg=None):
    return (struct.pack("<I", pattern) + (fg or color("indexed", 64))
            + (bg or color("indexed", 65)) + struct.pack("<I5dI", 0, 0, 0, 0, 0, 0, 0))


def border(*, left=0, right=0, top=0, bottom=0, diagonal=0,
           line_color=None, diagonal_up=False, diagonal_down=False):
    flags = int(diagonal_down) | (int(diagonal_up) << 1)
    shade = line_color or color("auto")
    return bytes([flags]) + b"".join(bytes([style, 0]) + shade
                                    for style in (top, bottom, left, right, diagonal))


def xf(*, parent=0, numfmt=0, font_id=0, fill_id=0, border_id=0,
       rotation=0, indent=0, flags=0x1010, apply=0):
    return struct.pack("<5HBBHH", parent, numfmt, font_id, fill_id, border_id,
                       rotation, indent, flags, apply)


def named_style(name="Normal", *, parent=0, builtin=0, hidden=False):
    flags = (builtin is not None) | (int(hidden) << 1)
    return struct.pack("<IHBB", parent, flags, builtin or 0, 0) + wide(name)


def stylesheet(*, fonts=None, fills=None, borders=None, xfs=None,
               style_xfs=None, styles=None, numfmts=(), palette=None, extra=b""):
    groups = (
        (615, 616, 44, [struct.pack("<H", number) + wide(code) for number, code in numfmts]),
        (611, 612, 43, fonts if fonts is not None else [font()]),
        (603, 604, 45, fills if fills is not None else [fill()]),
        (613, 614, 46, borders if borders is not None else [border()]),
        (626, 627, 47, style_xfs if style_xfs is not None else [xf(parent=65535)]),
        (617, 618, 47, xfs if xfs is not None else [xf()]),
        (619, 620, 48, styles if styles is not None else [named_style()]),
    )
    result = record(278)
    for begin, end, item, values in groups:
        result += record(begin, struct.pack("<I", len(values)))
        result += b"".join(record(item, value) for value in values)
        result += record(end)
    if palette is not None:
        result += record(473) + record(565)
        result += b"".join(record(475, bytes.fromhex(rgb) + b"\xff") for rgb in palette)
        result += record(566) + record(474)
    return result + extra + record(279)


def styled_book(styles, *, index=0, value=1, extra_rows=(), columns=()):
    rows = [row_header(0, spans=((0, 0),)), cell(0, "number", value, ixfe=index)]
    rows.extend(extra_rows)
    data = package([("売上", worksheet(rows, dimensions=(0, 1048575, 0, 16383), columns=columns))])
    return replace_part(data, "xl/styles.bin", styles)
