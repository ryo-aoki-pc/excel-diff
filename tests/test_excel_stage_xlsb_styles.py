"""Compare and import BIFF12 styles without changing unrelated records."""

import io
from pathlib import Path
import struct
import sys
import unittest
from unittest import mock
import zipfile


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
import excel_stage_ooxml as OOXML
import excel_stage_xlsb as XLSB
import excel_stage_xlsb_styles as STYLES

from xlsb_builder import cell, package, record, records, replace_part, row_header, worksheet
from test_excel_stage_xlsb import parts, raw_cells, replace_cell
from xlsb_style_builder import border, color, fill, font, named_style, stylesheet, styled_book, xf

FIXTURES = Path(__file__).resolve().parent / "fixtures"
STYLE_PATH = "xl/styles.bin"
SHEET_PATH = "xl/worksheets/sheet1.bin"
COLLECTIONS = {
    "numFmts": (615, 616, 44), "fonts": (611, 612, 43),
    "fills": (603, 604, 45), "borders": (613, 614, 46),
    "cellStyleXfs": (626, 627, 47), "cellXfs": (617, 618, 47),
    "cellStyles": (619, 620, 48),
}


def collection(data, name):
    begin, end, entry = COLLECTIONS[name]
    active = False
    found = []
    for kind, payload, start, stop in records(data):
        if kind == begin:
            active = True
        elif kind == end:
            active = False
        elif active and kind == entry:
            found.append(payload)
    return found


def replace_collection(data, name, values):
    begin, end, entry = COLLECTIONS[name]
    result = []
    active = False
    for kind, payload, start, stop in records(data):
        if kind == begin:
            result.append(record(begin, struct.pack("<I", len(values))))
            result.extend(record(entry, value) for value in values)
            active = True
        elif kind == end:
            result.append(data[start:stop])
            active = False
        elif not active:
            result.append(data[start:stop])
    return b"".join(result)


def add_part(data, name, value):
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(output, "w") as result:
        for info in source.infolist():
            result.writestr(info, source.read(info.filename))
        result.writestr(name, value)
    return output.getvalue()


def set_cell_style(data, coordinate, index):
    row, col = OOXML.coordinate_key(coordinate)
    current = parts(data)
    raw = raw_cells(current[SHEET_PATH])[(row - 1, col - 1)]
    kind, payload, start, stop = next(records(raw))
    flags = struct.unpack_from("<I", payload, 4)[0] & 0xFF000000
    encoded = record(kind, payload[:4] + struct.pack("<I", flags | index) + payload[8:])
    changed = replace_cell(current[SHEET_PATH], row - 1, col - 1, encoded)
    return replace_part(data, SHEET_PATH, changed)


class XlsbFormatFixtureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = (FIXTURES / "sales.xlsb").read_bytes()
        cls.changed = (FIXTURES / "sales-edited.xlsb").read_bytes()

    def books(self):
        return XLSB.Workbook(self.original), XLSB.Workbook(self.changed)

    def test_fixture_fields_and_projection_match_ooxml(self):
        binary = self.books()
        xml = tuple(OOXML.Workbook((FIXTURES / name).read_bytes())
                    for name in ("sales.xlsx", "sales-edited.xlsx"))
        for version in (0, 1):
            with self.subTest(version=version):
                for bin_sheet, xml_sheet in zip(binary[version].sheets, xml[version].sheets):
                    self.assertEqual(bin_sheet.name, xml_sheet.name)
                    coordinates = {key[1] for key in binary[version].cells if key[0] == bin_sheet.id}
                    coordinates |= {key[1] for key in binary[version].cell_styles if key[0] == bin_sheet.id}
                    coordinates |= {key[1] for key in xml[version].cell_styles if key[0] == xml_sheet.id}
                    for coordinate in coordinates:
                        bin_index = binary[version].effective_style((bin_sheet.id, coordinate))
                        xml_index = xml[version].effective_style((xml_sheet.id, coordinate))
                        self.assertEqual(binary[version].stylesheet().fields(bin_index),
                                         xml[version].stylesheet().fields(xml_index), coordinate)
                actual = binary[version].compare_formats(binary[1 - version]).projection()
                expected = xml[version].compare_formats(xml[1 - version]).projection()
                for side in (0, 1, 2):
                    self.assertEqual([actual[side][sheet.id] for sheet in binary[version].sheets],
                                     [expected[side][sheet.id] for sheet in xml[version].sheets])

    def assert_selected_formats(self, base, source, values, formats):
        result_data = base.rebuild(source, values, formats)
        result = XLSB.Workbook(result_data)
        keys = set(base.cells) | set(source.cells) | set(base.cell_styles) | set(source.cell_styles)
        for key in keys:
            self.assertEqual(result.cells.get(key),
                             source.cells.get(key) if key in values else base.cells.get(key), key)
            expected = source if key in formats else base
            self.assertEqual(result.stylesheet().exact(result.effective_style(key)),
                             expected.stylesheet().exact(expected.effective_style(key)), key)
        before, after = parts(base.data), parts(result_data)
        self.assertEqual(after["xl/sharedStrings.bin"], before["xl/sharedStrings.bin"])
        STYLES.check_styles_appended(before[STYLE_PATH], after[STYLE_PATH])
        for sheet in base.sheets:
            old, new = raw_cells(before[sheet.path]), raw_cells(after[sheet.path])
            for position in set(old) | set(new):
                coordinate = XLSB.xlsb._coordinate(*position)
                if (sheet.id, coordinate) not in values | formats:
                    self.assertEqual(new.get(position), old.get(position), coordinate)
        return result_data, result

    def test_fixture_format_selections_preserve_values_and_unselected_styles(self):
        for coordinate in ("B2", "A1", "M3"):
            with self.subTest(coordinate=coordinate):
                base, source = self.books()
                output, result = self.assert_selected_formats(base, source, set(), {("1", coordinate)})
                if coordinate == "M3":
                    self.assertNotIn(("1", coordinate), result.cells)
                    raw = raw_cells(parts(output)[SHEET_PATH])[(2, 12)]
                    self.assertEqual(next(records(raw))[0], 1)
                # A format-only edit does not invalidate calculation caches.
                self.assertEqual(parts(output)["xl/workbook.bin"], parts(base.data)["xl/workbook.bin"])
                self.assertEqual(parts(output).get("xl/calcChain.bin"), parts(base.data).get("xl/calcChain.bin"))

    def test_value_and_format_can_be_selected_independently(self):
        key = ("1", "B2")
        for values, formats in (({key}, set()), (set(), {key}), ({key}, {key})):
            with self.subTest(values=bool(values), formats=bool(formats)):
                base, source = self.books()
                self.assert_selected_formats(base, source, values, formats)

    def test_repeated_import_reuses_existing_xf_without_appending(self):
        base, source = self.books()
        key = ("1", "B2")
        first = base.rebuild(source, set(), {key})
        result = XLSB.Workbook(first)
        self.assertEqual(result.rebuild(source, set(), {key}), first)
        # A second cell importing that same style uses the appended xf.
        source = XLSB.Workbook(set_cell_style(source.data, "C2", source.effective_style(key)))
        output = result.rebuild(source, set(), {("1", "C2")})
        self.assertEqual(parts(output)[STYLE_PATH], parts(first)[STYLE_PATH])
        self.assertEqual(XLSB.Workbook(output).effective_style(("1", "C2")), result.effective_style(key))

    def test_fixture_unknown_style_records_are_preserved_after_import(self):
        base, source = self.books()
        output = base.rebuild(source, set(), {("1", "A1"), ("1", "B2"), ("1", "M3")})
        def unknown(contents):
            known = {278, 279} | {kind for kinds in COLLECTIONS.values() for kind in kinds}
            return [contents[start:stop] for kind, payload, start, stop in records(contents) if kind not in known]
        self.assertTrue(unknown(parts(base.data)[STYLE_PATH]))
        self.assertEqual(unknown(parts(output)[STYLE_PATH]), unknown(parts(base.data)[STYLE_PATH]))

    def test_format_rebuild_preserves_rich_array_and_merged_cell_records(self):
        for coordinate in ("E2", "J2", "J3", "L1"):
            with self.subTest(coordinate=coordinate):
                base, source = self.books()
                self.assertTrue(base.unsupported_reason(("1", coordinate)))
                source = XLSB.Workbook(set_cell_style(base.data, coordinate, source.effective_style(("1", "B2"))))
                output = base.rebuild(source, set(), {("1", coordinate)})
                pos = tuple(value - 1 for value in OOXML.coordinate_key(coordinate))
                old = next(records(raw_cells(parts(base.data)[SHEET_PATH])[pos]))[1]
                new = next(records(raw_cells(parts(output)[SHEET_PATH])[pos]))[1]
                self.assertEqual(new[:4] + new[8:], old[:4] + old[8:])
                with self.assertRaises(OOXML.WorkbookError):
                    base.rebuild(source, {("1", coordinate)})


class BinaryStylesTests(unittest.TestCase):
    def test_fields_match_ooxml_for_components_alignment_and_protection(self):
        data = stylesheet(
            fonts=[font(), font(bold=True, italic=True, underline=2, text_color=color("rgb", "FF0000"))],
            fills=[fill(), fill(pattern=1, fg=color("rgb", "FFFF00"))],
            borders=[border(), border(left=1, bottom=6, line_color=color("rgb", "0000FF"))],
            numfmts=[(164, "0.000")],
            xfs=[xf(), xf(numfmt=164, font_id=1, fill_id=1, border_id=1, rotation=45, indent=2,
                          flags=1 | (1 << 3) | 0x40 | 0x100 | (2 << 10) | 0x2000 | 0x8000)])
        xml = b'''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
          <fonts count="2"><font><name val="Arial"/><sz val="11"/><family val="2"/><color theme="1"/></font>
            <font><b/><i/><u val="double"/><name val="Arial"/><sz val="11"/><family val="2"/><color rgb="FFFF0000"/></font></fonts>
          <fills count="2"><fill><patternFill/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
          <borders count="2"><border/><border><left style="thin"><color rgb="FF0000FF"/></left><bottom style="double"><color rgb="FF0000FF"/></bottom></border></borders>
          <cellStyleXfs count="1"><xf/></cellStyleXfs>
          <cellXfs count="2"><xf/><xf numFmtId="164" fontId="1" fillId="1" borderId="1" quotePrefix="1"><alignment horizontal="left" vertical="center" textRotation="45" indent="2" wrapText="1" shrinkToFit="1" readingOrder="2"/><protection locked="0" hidden="1"/></xf></cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>'''
        expected = OOXML._Styles(xml, "styles.xml")
        actual = STYLES.BinaryStyles(data, STYLE_PATH)
        for index in (0, 1):
            self.assertEqual(actual.fields(index), expected.fields(index))
            self.assertEqual(actual.canonical(index), expected.canonical(index))
            self.assertIsNone(actual.problem(index))
        self.assertEqual(actual.style_names(), expected.style_names())

    def test_gradient_fields_match_ooxml_and_import_round_trips(self):
        gradient = (struct.pack("<I", 40) + color("indexed", 64) + color("indexed", 65)
                    + struct.pack("<I5dI", 0, 45, 0, 0, 0, 0, 2)
                    + color("rgb", "FF0000") + struct.pack("<d", 0)
                    + color("rgb", "0000FF") + struct.pack("<d", 1))
        source_styles = stylesheet(fills=[fill(), gradient], xfs=[xf(), xf(fill_id=1)])
        expected = OOXML._Styles(b'''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1"><font><name val="Arial"/><sz val="11"/><family val="2"/><color theme="1"/></font></fonts>
          <fills count="2"><fill><patternFill/></fill><fill><gradientFill degree="45"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill></fills>
          <borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>
          <cellXfs count="2"><xf/><xf fillId="1"/></cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>''', "styles.xml")
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        self.assertEqual(source.stylesheet().fields(1), expected.fields(1))
        base = XLSB.Workbook(styled_book(stylesheet()))
        result = XLSB.Workbook(base.rebuild(source, set(), {("1", "A1")}))
        self.assertEqual(result.stylesheet().exact(result.effective_style(("1", "A1"))),
                         source.stylesheet().exact(1))

    def test_colors_normalize_alpha_and_keep_theme_tint_indexed_and_auto(self):
        shades = (color("rgb", "FF0000", alpha=0), color("rgb", "FF0000", alpha=255),
                  color("theme", 4, tint=-16384), color("indexed", 2), color("auto"))
        styles = STYLES.BinaryStyles(stylesheet(fonts=[font(text_color=shade) for shade in shades],
                                               xfs=[xf(font_id=index) for index in range(len(shades))]))
        self.assertEqual(styles.fields(0), styles.fields(1))
        self.assertEqual(styles.fields(0)["font.color"], ("rgb", "FF0000"))
        self.assertEqual(styles.fields(2)["font.color"][:2], ("theme", "4"))
        self.assertAlmostEqual(float(styles.fields(2)["font.color"][2][1]), -0.500015259254738, delta=1e-12)
        self.assertEqual(styles.fields(3)["font.color"], ("indexed", "2"))
        # The automatic font color is omitted by Excel's XML writer.
        self.assertNotIn("font.color", styles.fields(4))

    def test_theme_tints_match_excels_positive_and_negative_values(self):
        for encoded, expected in ((-8191, -0.249977111117893),
                                  (-16384, -0.500015259254738),
                                  (13106, 0.39997558519241921),
                                  (16384, 0.500015259254738),
                                  (-32767, -1), (32767, 1)):
            with self.subTest(encoded=encoded):
                styles = STYLES.BinaryStyles(stylesheet(fonts=[font(text_color=color("theme", 4, tint=encoded))]))
                shade = styles.fields(0)["font.color"]
                self.assertEqual(shade[:2], ("theme", "4"))
                self.assertEqual(shade[2][0], "tint")
                self.assertAlmostEqual(float(shade[2][1]), expected, delta=1e-12)

    def test_gradient_defaults_are_omitted_and_nondefaults_are_kept(self):
        cases = (
            ((0, 0, 0, 0, 0, 0), {}),
            ((0, 45, 0, 0, 0, 0), {"degree": "45"}),
            ((1, 0, 0, 0, 0, 0), {"type": "path"}),
            ((1, 0, 0.25, 0.5, 0, 0.75),
             {"type": "path", "left": "0.25", "right": "0.5", "bottom": "0.75"}),
        )
        for values, expected in cases:
            with self.subTest(values=values):
                gradient = (struct.pack("<I", 40) + color("indexed", 64) + color("indexed", 65)
                            + struct.pack("<I5dI", *values, 2)
                            + color("rgb", "FF0000") + struct.pack("<d", 0)
                            + color("rgb", "0000FF") + struct.pack("<d", 1))
                node = STYLES._fill(gradient)[0]
                self.assertEqual(dict(node.attrib), expected)
                self.assertEqual([stop.get("position") for stop in node], ["0", "1"])

    def test_builtin_formats_and_apply_flags_follow_ooxml_semantics(self):
        data = stylesheet(numfmts=[(164, "0.00"), (165, "0.000")],
                          xfs=[xf(numfmt=2), xf(numfmt=164, apply=0x3F), xf(numfmt=14), xf(numfmt=165)])
        styles = STYLES.BinaryStyles(data, STYLE_PATH)
        self.assertEqual(styles.fields(0), styles.fields(1))
        self.assertNotEqual(styles.exact(0), styles.exact(1))
        self.assertEqual(styles.numfmt(2), ("code", "0.00"))
        self.assertEqual(styles.numfmt(14), ("builtin", 14))
        self.assertEqual(styles.numfmt(165), ("code", "0.000"))
        self.assertEqual(styles.fields(2)["numFmt"], ("builtin", 14))
        self.assertEqual(styles.fields(3)["numFmt"], ("code", "0.000"))

    def test_import_appends_components_and_named_style_with_valid_counts(self):
        original_styles = stylesheet()
        target = XLSB.Workbook(styled_book(original_styles))
        source_styles = stylesheet(
            fonts=[font(), font(name="Calibri", bold=True)],
            fills=[fill(), fill(pattern=1, fg=color("rgb", "00FF00"))],
            borders=[border(), border(left=1, line_color=color("rgb", "0000FF"))],
            numfmts=[(164, "0.000")],
            style_xfs=[xf(parent=65535), xf(parent=65535, font_id=1, fill_id=1, border_id=1, numfmt=164)],
            styles=[named_style(), named_style("売上形式", parent=1, builtin=None)],
            xfs=[xf(), xf(parent=1, font_id=1, fill_id=1, border_id=1, numfmt=164, apply=0x1B)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        output = target.rebuild(source, set(), {("1", "A1")})
        result = XLSB.Workbook(output)
        updated = parts(output)[STYLE_PATH]
        STYLES.check_styles_appended(original_styles, updated)
        for name, (begin, end, item) in COLLECTIONS.items():
            values = collection(updated, name)
            before = collection(original_styles, name)
            self.assertGreater(len(values), len(before), name)
            self.assertEqual(values[:len(before)], before, name)
            count = next(struct.unpack("<I", payload)[0]
                         for kind, payload, start, stop in records(updated) if kind == begin)
            self.assertEqual(count, len(values), name)
        self.assertEqual(result.stylesheet().exact(result.effective_style(("1", "A1"))),
                         source.stylesheet().exact(1))
        self.assertIn(("custom", "売上形式"), result.stylesheet().style_names())

    def test_existing_equivalent_xf_is_reused_without_style_changes(self):
        target_styles = stylesheet(fonts=[font(), font(bold=True)], xfs=[xf(), xf(font_id=1)])
        target = XLSB.Workbook(styled_book(target_styles))
        # Source numbering is different while the font and xf are equivalent.
        source_styles = stylesheet(fonts=[font(), font(italic=True), font(bold=True)],
                                   xfs=[xf(), xf(font_id=2)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        output = target.rebuild(source, set(), {("1", "A1")})
        self.assertEqual(parts(output)[STYLE_PATH], target_styles)
        self.assertEqual(XLSB.Workbook(output).effective_style(("1", "A1")), 1)

    def test_blank_style_import_creates_new_rows_and_removes_redundant_records(self):
        styles = stylesheet(fonts=[font(), font(bold=True)], xfs=[xf(), xf(font_id=1)])
        def book(rows, columns=()):
            data = package([("売上", worksheet(rows, dimensions=(0, 5, 0, 0), columns=columns))])
            return XLSB.Workbook(replace_part(data, STYLE_PATH, styles))
        base = book([row_header(0, spans=((0, 0),)), cell(0, "blank", ixfe=1)])
        source = book([row_header(5, spans=((0, 0),)), cell(0, "blank", ixfe=1)])
        output = base.rebuild(source, set(), {("1", "A1"), ("1", "A6")})
        updated = XLSB.Workbook(output)
        encoded = raw_cells(parts(output)[SHEET_PATH])
        self.assertNotIn((0, 0), encoded)
        self.assertEqual(next(records(encoded[(5, 0)]))[0], 1)
        self.assertEqual(updated.cells, {})
        self.assertEqual(updated.effective_style(("1", "A6")), 1)
        self.assertEqual(parts(output)[STYLE_PATH], styles)
        # A style already inherited from the target column needs no blank.
        base = book([], columns=((0, 0, 1),))
        source = book([row_header(5, spans=((0, 0),)), cell(0, "blank", ixfe=1)], columns=((0, 0, 1),))
        self.assertEqual(base.rebuild(source, set(), {("1", "A6")}), base.data)

    def test_number_format_import_does_not_reuse_an_occupied_custom_id(self):
        base_styles = stylesheet(numfmts=[(164, "0.0000"), (205, "0.00000")])
        base = XLSB.Workbook(styled_book(base_styles))
        source = XLSB.Workbook(styled_book(stylesheet(numfmts=[(164, "0.000")],
                                                    xfs=[xf(), xf(numfmt=164)]), index=1))
        output = base.rebuild(source, set(), {("1", "A1")})
        updated = STYLES.BinaryStyles(parts(output)[STYLE_PATH], STYLE_PATH)
        self.assertEqual(updated.numfmt(164), ("code", "0.0000"))
        self.assertEqual(updated.numfmt(205), ("code", "0.00000"))
        result = XLSB.Workbook(output)
        index = result.effective_style(("1", "A1"))
        self.assertEqual(result.stylesheet().fields(index)["numFmt"], ("code", "0.000"))
        imported = [struct.unpack_from("<H", value)[0] for value in collection(parts(output)[STYLE_PATH], "numFmts")]
        self.assertGreater(imported[-1], 205)

    def test_unavailable_style_components_and_invalid_indices_are_blocked(self):
        target = XLSB.Workbook(styled_book(stylesheet()))
        cases = (
            (stylesheet(xfs=[xf(), xf(font_id=99)]), 1, "フォント"),
            (stylesheet(xfs=[xf(), xf(fill_id=99)]), 1, "塗りつぶし"),
            (stylesheet(xfs=[xf(), xf(border_id=99)]), 1, "罫線"),
            (stylesheet(xfs=[xf(), xf(numfmt=164)]), 1, "表示形式"),
            (stylesheet(), 999, "書式"),
        )
        for styles, index, label in cases:
            with self.subTest(label=label):
                source = XLSB.Workbook(styled_book(styles, index=index))
                self.assertTrue(target.compare_formats(source).unsupported_reason(("1", "A1")))
                with self.assertRaisesRegex(OOXML.WorkbookError, "売上!A1 の書式は反映できません"):
                    target.rebuild(source, set(), {("1", "A1")})

    def test_palette_mismatch_only_blocks_low_indexed_colors(self):
        default = list(OOXML._DEFAULT_PALETTE)
        custom = default.copy()
        custom[2] = "123456"
        original = stylesheet()
        target = XLSB.Workbook(styled_book(original))
        for palette, kind, value, blocked in ((default, "indexed", 2, False),
                                               (custom, "indexed", 2, True),
                                               (custom, "theme", 1, False)):
            with self.subTest(palette=palette[2], kind=kind):
                styles = stylesheet(fonts=[font(), font(bold=True, text_color=color(kind, value))],
                                    xfs=[xf(), xf(font_id=1)], palette=palette)
                source = XLSB.Workbook(styled_book(styles, index=1))
                reason = target.compare_formats(source).unsupported_reason(("1", "A1"))
                self.assertEqual(bool(reason), blocked)
                if blocked:
                    self.assertEqual(reason, "ブックのパレットが異なる")
                    with self.assertRaises(OOXML.WorkbookError):
                        target.rebuild(source, set(), {("1", "A1")})
                else:
                    target.rebuild(source, set(), {("1", "A1")})

    def test_named_style_clash_and_style_limit_are_blocked(self):
        target_styles = stylesheet(styles=[named_style(), named_style("Report", parent=0, builtin=None)])
        source_styles = stylesheet(fonts=[font(), font(bold=True)],
                                   style_xfs=[xf(parent=65535), xf(parent=65535, font_id=1)],
                                   styles=[named_style(), named_style("report", parent=1, builtin=None)],
                                   xfs=[xf(), xf(parent=1, font_id=1)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        target = XLSB.Workbook(styled_book(target_styles))
        self.assertEqual(target.compare_formats(source).unsupported_reason(("1", "A1")),
                         "同じ名前で中身の違うスタイル")
        target = XLSB.Workbook(styled_book(stylesheet(xfs=[xf()] * 64000)))
        self.assertEqual(target.compare_formats(source).unsupported_reason(("1", "A1")),
                         "セルの書式の数が 64000 を超える")
        with self.assertRaises(OOXML.WorkbookError):
            target.rebuild(source, set(), {("1", "A1")})

    def test_cell_style_xf_limit_and_builtin_number_format_conflict_are_blocked(self):
        source_styles = stylesheet(fonts=[font(), font(bold=True)],
                                   style_xfs=[xf(parent=65535), xf(parent=65535, font_id=1)],
                                   styles=[named_style(), named_style("Report", parent=1, builtin=None)],
                                   xfs=[xf(), xf(parent=1, font_id=1)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        base = XLSB.Workbook(styled_book(stylesheet(style_xfs=[xf(parent=65535)] * 64000)))
        self.assertEqual(base.compare_formats(source).unsupported_reason(("1", "A1")),
                         "セルのスタイルの数が 64000 を超える")
        with self.assertRaises(OOXML.WorkbookError):
            base.rebuild(source, set(), {("1", "A1")})
        base = XLSB.Workbook(styled_book(stylesheet(numfmts=[(14, "0.000")])))
        source = XLSB.Workbook(styled_book(stylesheet(xfs=[xf(), xf(numfmt=14)]), index=1))
        self.assertEqual(base.compare_formats(source).unsupported_reason(("1", "A1")),
                         "組み込みの表示形式番号が別の定義に使われている")
        with self.assertRaises(OOXML.WorkbookError):
            base.rebuild(source, set(), {("1", "A1")})

    def test_extended_xfs_and_components_are_blocked_without_losing_their_bytes(self):
        original = stylesheet(fonts=[font(), font(bold=True)],
                              fills=[fill(), fill(pattern=1, fg=color("rgb", "FFFF00"))],
                              borders=[border(), border(left=1)],
                              xfs=[xf(), xf(font_id=1, fill_id=1, border_id=1)])
        target = XLSB.Workbook(styled_book(stylesheet()))
        for name in ("cellXfs", "fonts", "fills", "borders"):
            end = COLLECTIONS[name][1]
            extended = b"".join((record(35, struct.pack("<I", 0x1234)) + record(4095, b"unknown extension") + record(36)
                                 if kind == end else b"") + original[start:stop]
                                for kind, payload, start, stop in records(original))
            source = XLSB.Workbook(styled_book(extended, index=1))
            with self.subTest(collection=name):
                self.assertTrue(target.compare_formats(source).unsupported_reason(("1", "A1")))
                with self.assertRaises(OOXML.WorkbookError):
                    target.rebuild(source, set(), {("1", "A1")})
                if name != "cellXfs":
                    # An exact existing xf needs no copied component records.
                    already_present = XLSB.Workbook(styled_book(extended))
                    output = already_present.rebuild(source, set(), {("1", "A1")})
                    self.assertEqual(parts(output)[STYLE_PATH], extended)

    def test_same_named_style_with_different_parent_xf_is_blocked_unless_xf_exists(self):
        fonts = [font(), font(bold=True), font(italic=True)]
        fills = [fill(), fill(pattern=1, fg=color("rgb", "FFFF00"))]
        names = [named_style(), named_style("Report", parent=1, builtin=None)]
        source_styles = stylesheet(fonts=fonts, fills=fills,
                                   style_xfs=[xf(parent=65535), xf(parent=65535, font_id=2)],
                                   styles=names, xfs=[xf(), xf(parent=1, font_id=2, fill_id=1)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        target_styles = stylesheet(fonts=fonts, fills=fills,
                                   style_xfs=[xf(parent=65535), xf(parent=65535, font_id=1)],
                                   styles=names)
        target = XLSB.Workbook(styled_book(target_styles))
        self.assertEqual(target.compare_formats(source).unsupported_reason(("1", "A1")),
                         "同じ名前で中身の違うスタイル")
        with self.assertRaisesRegex(OOXML.WorkbookError, "同じ名前で中身の違うスタイル"):
            target.rebuild(source, set(), {("1", "A1")})
        # Exact existing cell XFs require no new style or parent XF records.
        target_styles = replace_collection(target_styles, "cellXfs",
                                           collection(source_styles, "cellXfs"))
        target = XLSB.Workbook(styled_book(target_styles))
        self.assertIsNone(target.compare_formats(source).unsupported_reason(("1", "A1")))
        output = target.rebuild(source, set(), {("1", "A1")})
        self.assertEqual(parts(output)[STYLE_PATH], target_styles)
        self.assertEqual(XLSB.Workbook(output).effective_style(("1", "A1")), 1)

    def named_style_with_prefix(self, prefix):
        source_styles = stylesheet(fonts=[font(), font(bold=True)],
                                   style_xfs=[xf(parent=65535), xf(parent=65535, font_id=1)],
                                   styles=[named_style(), named_style("Report", parent=1, builtin=None)],
                                   xfs=[xf(), xf(parent=1, font_id=1)])
        chunks = []
        active, count = False, 0
        for kind, payload, start, stop in records(source_styles):
            if kind == 619:
                active = True
            elif kind == 620:
                active = False
            if active and kind == 48:
                if count == 1:
                    chunks.append(prefix)
                count += 1
            chunks.append(source_styles[start:stop])
        return b"".join(chunks)

    def test_new_named_style_with_guid_is_imported_without_its_identifier_records(self):
        prefix = (record(37, bytes.fromhex("010002110080"))
                  + record(3072, bytes(range(16))) + record(38))
        source_styles = self.named_style_with_prefix(prefix)
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        target = XLSB.Workbook(styled_book(stylesheet()))
        key = ("1", "A1")
        self.assertIsNone(target.compare_formats(source).unsupported_reason(key))
        output = target.rebuild(source, set(), {key})
        result = XLSB.Workbook(output)
        self.assertEqual(result.stylesheet().exact(result.effective_style(key)), source.stylesheet().exact(1))
        self.assertIn(("custom", "report"), result.stylesheet().style_names())
        self.assertFalse({37, 3072, 38} & {kind for kind, *_ in records(parts(output)[STYLE_PATH])})
        STYLES.check_styles_appended(parts(target.data)[STYLE_PATH], parts(output)[STYLE_PATH])
        # Reusing an exact target XF preserves its existing identifier bytes.
        target = XLSB.Workbook(styled_book(source_styles))
        self.assertIsNone(target.compare_formats(source).unsupported_reason(key))
        output = target.rebuild(source, set(), {key})
        self.assertEqual(parts(output)[STYLE_PATH], source_styles)
        self.assertEqual(XLSB.Workbook(output).effective_style(key), 1)

    def test_new_named_style_with_unknown_alternate_content_is_blocked(self):
        attachments = (record(3073, bytes(range(16))),
                       record(3072, bytes(15)), record(3072, bytes(17)),
                       record(3072, bytes(16)) + record(4095, b"unknown"))
        for attached in attachments:
            with self.subTest(attachment=attached):
                prefix = record(37, bytes.fromhex("010002110080")) + attached + record(38)
                source_styles = self.named_style_with_prefix(prefix)
                source = XLSB.Workbook(styled_book(source_styles, index=1))
                target = XLSB.Workbook(styled_book(stylesheet()))
                self.assertEqual(target.compare_formats(source).unsupported_reason(("1", "A1")),
                                 "書式の拡張機能")
                with self.assertRaises(OOXML.WorkbookError):
                    target.rebuild(source, set(), {("1", "A1")})
                # Existing exact target XFs need no copied extension records.
                target = XLSB.Workbook(styled_book(source_styles))
                self.assertIsNone(target.compare_formats(source).unsupported_reason(("1", "A1")))
                output = target.rebuild(source, set(), {("1", "A1")})
                self.assertEqual(parts(output)[STYLE_PATH], source_styles)
                self.assertEqual(XLSB.Workbook(output).effective_style(("1", "A1")), 1)

    def test_unknown_pivot_binary_blocks_custom_number_format_allocation(self):
        target = styled_book(stylesheet())
        target = add_part(target, "xl/pivotCache/pivotCacheDefinition1.bin", record(4095, b"unknown"))
        source_styles = stylesheet(numfmts=[(164, "0.000")], xfs=[xf(), xf(numfmt=164)])
        source = XLSB.Workbook(styled_book(source_styles, index=1))
        base = XLSB.Workbook(target)
        self.assertTrue(base.compare_formats(source).unsupported_reason(("1", "A1")))
        with self.assertRaises(OOXML.WorkbookError):
            base.rebuild(source, set(), {("1", "A1")})
        # No format id is allocated for an existing invariant builtin code.
        source = XLSB.Workbook(styled_book(stylesheet(xfs=[xf(), xf(numfmt=2)]), index=1))
        base.rebuild(source, set(), {("1", "A1")})

    def test_unread_differential_formats_block_new_custom_number_format_ids(self):
        source = XLSB.Workbook(styled_book(stylesheet(numfmts=[(164, "0.000")],
                                                    xfs=[xf(), xf(numfmt=164)]), index=1))
        for kind in (507, 1171, 2102):
            with self.subTest(record_type=kind):
                # DXF extensions can also reference ids that are not exposed
                # by the ordinary XF and BrtFmt collections.
                extra = record(35, struct.pack("<I", 0)) + record(kind, b"unknown number-format references") + record(36)
                base = XLSB.Workbook(styled_book(stylesheet(extra=extra)))
                reason = base.compare_formats(source).unsupported_reason(("1", "A1"))
                self.assertEqual(reason, "条件付き書式の表示形式番号を確認できない")
                with self.assertRaises(OOXML.WorkbookError):
                    base.rebuild(source, set(), {("1", "A1")})
                known = XLSB.Workbook(styled_book(stylesheet(xfs=[xf(), xf(numfmt=2)]), index=1))
                base.rebuild(known, set(), {("1", "A1")})

    def test_append_validator_rejects_changed_existing_entries_counts_and_extensions(self):
        original = stylesheet(extra=record(35, b"header") + record(4095, b"keep") + record(36))
        for name in ("fonts", "fills", "borders", "cellXfs", "cellStyleXfs", "cellStyles", "numFmts"):
            entries = collection(original, name)
            if entries:
                damaged = replace_collection(original, name, [entries[0] + b"corrupt"] + entries[1:])
            else:
                damaged = original.replace(record(COLLECTIONS[name][0], struct.pack("<I", 0)),
                                           record(COLLECTIONS[name][0], struct.pack("<I", 1)))
            with self.subTest(collection=name), self.assertRaises(OOXML.WorkbookError):
                STYLES.check_styles_appended(original, damaged)
        with self.assertRaises(OOXML.WorkbookError):
            STYLES.check_styles_appended(original, original.replace(b"keep", b"lose"))

    def test_result_validation_detects_unselected_and_row_style_changes(self):
        styles = stylesheet(fonts=[font(), font(bold=True)], xfs=[xf(), xf(font_id=1)])
        rows = [row_header(0, spans=((0, 1),)), cell(0, "number", 1), cell(1, "number", 2)]
        original = replace_part(package([("売上", worksheet(rows, dimensions=(0, 0, 0, 1)))]), STYLE_PATH, styles)
        base = XLSB.Workbook(original)
        source = XLSB.Workbook(set_cell_style(original, "A1", 1))
        rebuild_sheet = XLSB.Workbook._rebuild_sheet
        for corruption in ("unselected", "row", "phonetic"):
            def broken(book, *args, **kwargs):
                data = rebuild_sheet(book, *args, **kwargs)
                if corruption == "unselected":
                    return replace_cell(data, 0, 1, cell(1, "number", 2, ixfe=1))
                if corruption == "phonetic":
                    return replace_cell(data, 0, 0, cell(0, "number", 1, ixfe=1, phonetic=True))
                result = []
                for kind, payload, start, stop in records(data):
                    if kind == 0:
                        payload = payload[:4] + struct.pack("<I", 1) + payload[8:11] + bytes([payload[11] | 0x40]) + payload[12:]
                        result.append(record(0, payload))
                    else:
                        result.append(data[start:stop])
                return b"".join(result)
            with self.subTest(corruption=corruption), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, "作り直したブックの検証に失敗しました"):
                    base.rebuild(source, set(), {("1", "A1")})

    def test_result_validation_detects_wrong_imported_style(self):
        original = styled_book(stylesheet())
        source = XLSB.Workbook(styled_book(stylesheet(fonts=[font(), font(bold=True)],
                                                    xfs=[xf(), xf(font_id=1)]), index=1))
        base = XLSB.Workbook(original)
        with mock.patch.object(STYLES.BinaryStyleImporter, "index_for", return_value=0):
            with self.assertRaises(OOXML.WorkbookError):
                base.rebuild(source, set(), {("1", "A1")})


if __name__ == "__main__":
    unittest.main()
