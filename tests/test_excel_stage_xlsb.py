"""Verify xlsb value imports preserve binary records and package metadata."""

import hashlib
import io
from pathlib import Path
import struct
import sys
import unittest
from unittest import mock
import zipfile

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
import excel_stage_ooxml as OOXML
import excel_stage_xlsb as XLSB
import excel_xlsb as BINARY

from xlsb_builder import (
    cell, formula_range, package, parsed_formula, ptg_exp, ptg_int,
    record, records, recompute_index, replace_part, rich_string, row_header,
    wide, with_indexes, worksheet,
)

FIXTURES = Path(__file__).resolve().parent / "fixtures"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_PATH = "xl/worksheets/sheet1.bin"
INDEX_PATH = "xl/worksheets/binaryIndex1.bin"
VALIDATION_ERROR = "作り直したブックの検証に失敗しました。書き込みは行っていません"
CELL_TYPES = set(range(1, 12)) | {62}


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def raw_cells(contents):
    result = {}
    row = None
    for kind, payload, start, end in records(contents):
        if kind == 0:
            row = struct.unpack_from("<I", payload)[0]
        elif kind in CELL_TYPES:
            col = struct.unpack_from("<I", payload)[0]
            result[(row, col)] = contents[start:end]
    return result


def replace_cell(contents, wanted_row, wanted_col, replacement):
    result = []
    row = None
    found = False
    for kind, payload, start, end in records(contents):
        if kind == 0:
            row = struct.unpack_from("<I", payload)[0]
        if kind in CELL_TYPES and (row, struct.unpack_from("<I", payload)[0]) == (wanted_row, wanted_col):
            result.append(replacement)
            found = True
        else:
            result.append(contents[start:end])
    if not found:
        raise AssertionError("cell to replace is missing")
    return b"".join(result)


def replace_record(contents, wanted_kind, change):
    result = []
    changed = False
    for kind, payload, start, end in records(contents):
        if kind == wanted_kind and not changed:
            result.append(record(kind, change(payload)))
            changed = True
        else:
            result.append(contents[start:end])
    if not changed:
        raise AssertionError("record to replace is missing")
    return b"".join(result)


def minimal(rows, *, name="売上", dimensions=None, default_height=None,
            columns=(), workbook_records=b"", date1904=False, shared_strings=(),
            before_data=b"", after_data=b"", indexes=True):
    if dimensions is None:
        positions = raw_cells(b"".join(rows))
        if positions:
            dimensions = (min(row for row, col in positions), max(row for row, col in positions),
                          min(col for row, col in positions), max(col for row, col in positions))
        else:
            dimensions = (0, 0, 0, 0)
    sheet = worksheet(rows, dimensions=dimensions, default_height=default_height,
                      columns=columns, before_data=before_data, after_data=after_data)
    data = package([(name, sheet)], cell_formats=(0,) * 10,
                   shared_strings=shared_strings, date1904=date1904,
                   workbook_records=record(157, bytes(26)) + workbook_records)
    return with_indexes(data) if indexes else data


def rows_by_number(contents):
    return {struct.unpack_from("<I", payload)[0]: payload
            for kind, payload, start, end in records(contents) if kind == 0}


class XlsbStagingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base_bytes = (FIXTURES / "sales.xlsb").read_bytes()
        cls.source_bytes = (FIXTURES / "sales-edited.xlsb").read_bytes()

    def books(self):
        return XLSB.Workbook(self.base_bytes), XLSB.Workbook(self.source_bytes)

    def assert_rebuild(self, base, source, selected):
        output = base.rebuild(source, selected)
        result = XLSB.Workbook(output)
        for key in set(base.cells) | set(source.cells) | set(result.cells):
            expected = source.cells.get(key) if key in selected else base.cells.get(key)
            self.assertEqual(result.cells.get(key), expected, key)
            if key in selected:
                self.assertEqual(result.effective_style(key), base.effective_style(key), key)
        original_parts, result_parts = parts(base.data), parts(output)
        for sheet in base.sheets:
            old = raw_cells(original_parts[sheet.path])
            new = raw_cells(result_parts[sheet.path])
            for position in set(old) | set(new):
                coordinate = BINARY._coordinate(*position)
                if (sheet.id, coordinate) not in selected:
                    self.assertEqual(new.get(position), old.get(position), coordinate)
            preserved = lambda contents: [contents[start:end] for kind, payload, start, end in records(contents)
                                           if kind not in CELL_TYPES | {0, 148}]
            self.assertEqual(preserved(original_parts[sheet.path]), preserved(result_parts[sheet.path]))
        return output, result

    def assert_indexes(self, data):
        current = parts(data)
        for path, contents in current.items():
            if path.startswith("xl/worksheets/sheet") and path.endswith(".bin"):
                index = path.replace("/sheet", "/binaryIndex")
                if index in current:
                    self.assertEqual(current[index], recompute_index(contents), path)
                    self.assertEqual(XLSB.recompute_index(contents), current[index])

    def test_projection_and_unsupported_reasons_match_ooxml_oracle(self):
        binary = self.books()
        xml = tuple(OOXML.Workbook((FIXTURES / name).read_bytes())
                    for name in ("sales.xlsx", "sales-edited.xlsx"))
        for index in (0, 1):
            with self.subTest(version=index):
                actual = binary[index].projection(binary[1 - index])
                expected = xml[index].projection(xml[1 - index])
                # OOXML stores one shared follower without formula text, and its
                # time literal has more decimal digits than repr(binary64).
                # Keep these verified format differences explicit; all other
                # projection records must agree exactly.
                self.assertEqual(xml[index].cells[("1", "H3")], OOXML.Cell("number", "0.57291666666666663"))
                self.assertEqual(binary[index].cells[("1", "H3")], OOXML.Cell("number", "0.5729166666666666"))
                expected["1"] = expected["1"].replace(
                    '["H3","number","0.57291666666666663"]', '["H3","number","0.5729166666666666"]')
                if index == 0:
                    self.assertEqual(xml[index].cells[("1", "D4")], OOXML.Cell("formula", "="))
                    self.assertEqual(binary[index].cells[("1", "D4")], OOXML.Cell("formula", "=B4*C4"))
                    expected["1"] = expected["1"].replace('["D4","formula","="]', '["D4","formula","=B4*C4"]')
                self.assertEqual([actual[sheet.id] for sheet in binary[index].sheets],
                                 [expected[sheet.id] for sheet in xml[index].sheets])
                for binary_sheet, xml_sheet in zip(binary[index].sheets, xml[index].sheets):
                    self.assertEqual(binary_sheet.name, xml_sheet.name)
                    coordinates = {key[1] for key in binary[index].cells if key[0] == binary_sheet.id}
                    coordinates |= {key[1] for key in xml[index].cells if key[0] == xml_sheet.id}
                    coordinates |= {"J2", "J3", "L1", "M2", "M3"}
                    for coordinate in coordinates:
                        actual_reason = binary[index].unsupported_reason((binary_sheet.id, coordinate))
                        expected_reason = xml[index].unsupported_reason((xml_sheet.id, coordinate))
                        if index == 0 and binary_sheet.id == "1" and coordinate in {"D3", "D4"}:
                            # BIFF12 stores independent formulas; the xlsx
                            # counterpart stores a shared formula attribute.
                            self.assertEqual(expected_reason, "数式の拡張機能")
                            self.assertIsNone(actual_reason)
                        else:
                            self.assertEqual(actual_reason, expected_reason, (binary_sheet.name, coordinate))

    def test_phonetic_projection_and_unsupported_reasons_match_ooxml_oracle(self):
        names = ("phonetic", "phonetic-edited")
        binary = tuple(XLSB.Workbook((FIXTURES / (name + ".xlsb")).read_bytes()) for name in names)
        xml = tuple(OOXML.Workbook((FIXTURES / (name + ".xlsx")).read_bytes()) for name in names)
        for index in (0, 1):
            with self.subTest(version=names[index]):
                actual = binary[index].projection(binary[1 - index])
                expected = xml[index].projection(xml[1 - index])
                self.assertEqual([sheet.name for sheet in binary[index].sheets], ["ふりがな"])
                self.assertEqual([sheet.name for sheet in xml[index].sheets], ["ふりがな"])
                self.assertEqual([actual[sheet.id] for sheet in binary[index].sheets],
                                 [expected[sheet.id] for sheet in xml[index].sheets])
                coordinates = {key[1] for key in binary[index].cells} | {key[1] for key in xml[index].cells}
                for coordinate in coordinates:
                    key = ("1", coordinate)
                    self.assertEqual(binary[index].unsupported_reason(key), xml[index].unsupported_reason(key), key)
                self.assertEqual(binary[index].unsupported_reason(("1", "A5")), "リッチテキスト")

    def test_rebuild_phonetic_fixture_changes_preserves_shared_strings(self):
        base = XLSB.Workbook((FIXTURES / "phonetic.xlsb").read_bytes())
        source = XLSB.Workbook((FIXTURES / "phonetic-edited.xlsb").read_bytes())
        changed = {key for key in set(base.cells) | set(source.cells)
                   if base.cells.get(key) != source.cells.get(key)}
        self.assertEqual(changed, {("1", "A2"), ("1", "B1"), ("1", "A8")})
        selected = {key for key in changed
                    if not base.unsupported_reason(key) and not source.unsupported_reason(key)}
        self.assertEqual(selected, changed)
        output, result = self.assert_rebuild(base, source, selected)
        self.assert_indexes(output)
        self.assertEqual(parts(output)["xl/sharedStrings.bin"], parts(base.data)["xl/sharedStrings.bin"])

    def test_shared_strings_with_multiple_phonetic_runs_are_plain_values(self):
        inline = rich_string("東京と大阪", phonetic="トウキョウオオサカ",
                             phonetic_runs=((0, 0, 2), (5, 3, 2)))
        data = minimal([row_header(0, spans=((0, 1),)), cell(0, "shared-string", 0),
                        cell(1, "rich-string", inline)],
                       shared_strings=[rich_string("今日は東京へ行く", phonetic="キョウトウキョウイ",
                                                   phonetic_runs=((0, 0, 2), (3, 3, 2), (8, 6, 1)))])
        book = XLSB.Workbook(data)
        self.assertEqual(book.cells[("1", "A1")], OOXML.Cell("string", "今日は東京へ行く"))
        self.assertIsNone(book.unsupported_reason(("1", "A1")))
        self.assertEqual(book.cells[("1", "B1")], OOXML.Cell("string", "東京と大阪"))
        self.assertEqual(book.unsupported_reason(("1", "B1")), "リッチテキスト")

    def test_cells_use_shared_contract_and_decimal_number_normalization(self):
        data = minimal([row_header(0, spans=((0, 5),)), cell(0, "rk", (12 << 2) | 2),
                        cell(1, "number", 0.1), cell(2, "number", -0.0),
                        cell(3, "number", 1e-7), cell(4, "blank"), cell(5, "blank", ixfe=3)])
        book = XLSB.Workbook(data)
        self.assertEqual(book.data, data)
        self.assertTrue(all(isinstance(sheet, OOXML.Sheet) for sheet in book.sheets))
        self.assertEqual([book.cells[("1", coordinate)] for coordinate in ("A1", "B1", "C1", "D1")],
                         [OOXML.Cell("number", value) for value in ("12", "0.1", "0", "1E-7")])
        self.assertEqual(book.cell_styles[("1", "E1")], 0)
        self.assertEqual(book.cell_styles[("1", "F1")], 3)
        self.assertNotIn(("1", "A1"), book.cell_styles)

    def test_rebuild_each_fixture_change_preserves_other_cells_and_indexes(self):
        for coordinate in ("B2", "A6", "AMK3", "A50", "F3", "E5", "D4"):
            with self.subTest(coordinate=coordinate):
                base, source = self.books()
                output, result = self.assert_rebuild(base, source, {("1", coordinate)})
                self.assert_indexes(output)
                current = parts(output)
                self.assertNotIn("xl/calcChain.bin", current)
                self.assertNotIn("xl/_rels/calcChain.bin.rels", current)
                calc = next(payload for kind, payload, start, end in records(current["xl/workbook.bin"]) if kind == 157)
                self.assertTrue(struct.unpack_from("<H", calc, 24)[0] & 1)
                rels = etree.fromstring(current["xl/_rels/workbook.bin.rels"])
                self.assertFalse(any(node.get("Type", "").endswith("/calcChain") for node in rels))
                types = etree.fromstring(current["[Content_Types].xml"])
                self.assertFalse(any(node.get("PartName") == "/xl/calcChain.bin" for node in types))

    def test_rebuild_all_changed_values_preserves_styles_and_unmodified_parts(self):
        base, source = self.books()
        selected = {key for key in set(base.cells) | set(source.cells)
                    if base.cells.get(key) != source.cells.get(key)
                    and not base.unsupported_reason(key) and not source.unsupported_reason(key)}
        self.assertGreater(len(selected), 10)
        output, result = self.assert_rebuild(base, source, selected)
        self.assert_indexes(output)
        original, current = parts(base.data), parts(output)
        changed = {sheet.path for sheet in base.sheets if any(key[0] == sheet.id for key in selected)}
        changed |= {"xl/workbook.bin", "xl/_rels/workbook.bin.rels", "[Content_Types].xml"}
        changed |= {path for path in original if "binaryIndex" in path}
        removed = {"xl/calcChain.bin", "xl/_rels/calcChain.bin.rels"}
        for path in set(original) - changed - removed:
            self.assertEqual(current[path], original[path], path)
        self.assertEqual(list(current), [path for path in original if path not in removed])
        self.assertEqual(current["xl/sharedStrings.bin"], original["xl/sharedStrings.bin"])
        self.assertEqual(current["xl/styles.bin"], original["xl/styles.bin"])

    def test_unchanged_value_selection_returns_exact_original_even_with_format_changes(self):
        base, source = self.books()
        self.assertEqual(base.rebuild(source, {("1", "A1")}), base.data)
        self.assertEqual(base.rebuild(source, set()), base.data)

    def test_shared_string_import_is_inline_and_formula_tokens_are_copied(self):
        base, source = self.books()
        output, result = self.assert_rebuild(base, source, {("1", "E5"), ("1", "D4")})
        target_cells = raw_cells(parts(output)[SHEET_PATH])
        source_cells = raw_cells(parts(source.data)[SHEET_PATH])
        kind, payload, start, end = next(records(target_cells[(4, 4)]))
        self.assertEqual(kind, 6)
        self.assertEqual(payload[8:], wide(source.cells[("1", "E5")].value))
        target_formula = next(records(target_cells[(3, 3)]))[1]
        source_formula = next(records(source_cells[(3, 3)]))[1]
        self.assertEqual(target_formula[8:], source_formula[8:])

    def test_target_phonetic_flag_and_style_are_preserved(self):
        source = minimal([row_header(0, spans=((0, 0),)), cell(0, "string", "source", ixfe=7)])
        target = minimal([row_header(0, spans=((0, 0),)), cell(0, "string", "target", ixfe=3, phonetic=True)])
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "A1")})
        payload = next(records(raw_cells(parts(output)[SHEET_PATH])[(0, 0)]))[1]
        self.assertEqual(struct.unpack_from("<I", payload, 4)[0], 3 | 0x01000000)

    def test_deletion_keeps_nonzero_style_as_blank_and_removes_zero_style_record(self):
        target = minimal([row_header(0, spans=((0, 1),)), cell(0, "number", 1), cell(1, "number", 2, ixfe=3)])
        source = minimal([row_header(0, spans=((0, 1),))])
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "A1"), ("1", "B1")})
        encoded = raw_cells(parts(output)[SHEET_PATH])
        self.assertNotIn((0, 0), encoded)
        self.assertEqual(next(records(encoded[(0, 1)]))[0], 1)
        self.assertEqual(result.cell_styles[("1", "B1")], 3)

    def test_deletion_preserves_explicit_zero_over_inherited_row_or_column_style(self):
        cases = (
            ([row_header(0, spans=((0, 0),)), cell(0, "number", 1)], ((0, 0, 2),)),
            ([row_header(0, ixfe=4, spans=((0, 0),), ghost=True), cell(0, "number", 1)], ()),
        )
        for rows, columns in cases:
            target = minimal(rows, columns=columns)
            source = minimal([rows[0]], columns=columns)
            base, other = XLSB.Workbook(target), XLSB.Workbook(source)
            with self.subTest(columns=columns):
                self.assertEqual(base.effective_style(("1", "A1")), 0)
                output, result = self.assert_rebuild(base, other, {("1", "A1")})
                self.assertEqual(raw_cells(parts(output)[SHEET_PATH])[(0, 0)], cell(0, "blank"))
                self.assertEqual(result.effective_style(("1", "A1")), 0)
                self.assert_indexes(output)

    def test_expansion_coalesces_all_spans_in_the_same_column_block(self):
        spans = ((0, 1), (5, 7), (1024, 1026))
        rows = [row_header(0, spans=spans), cell(0, "number", 1)]
        base = XLSB.Workbook(minimal(rows))
        source = XLSB.Workbook(minimal([row_header(0, spans=((0, 10), (1024, 1026))),
                                       cell(0, "number", 1), cell(10, "number", 2)]))
        output, result = self.assert_rebuild(base, source, {("1", "K1")})
        self.assertEqual(result._binary_sheets["1"].rows[0].spans, ((0, 10), (1024, 1026)))
        self.assert_indexes(output)

    def test_empty_and_blank_only_sheets_have_excel_index_records(self):
        empty = worksheet([])
        expected_empty = record(42, struct.pack("<IIII", 0, 32, 0, 0) + bytes(8)) + record(277)
        self.assertEqual(XLSB.recompute_index(empty), expected_empty)
        self.assertEqual(recompute_index(empty), expected_empty)
        blank = worksheet([row_header(4, spans=((1024, 1024),)), cell(1024, "blank", ixfe=2)])
        encoded = XLSB.recompute_index(blank)
        self.assertEqual(encoded, recompute_index(blank))
        block, rowblock, end = list(records(encoded))
        self.assertEqual(block[0], 42)
        self.assertEqual(struct.unpack_from("<II", block[1]), (4, 36))
        self.assertEqual(rowblock[0], 40)
        self.assertEqual(struct.unpack_from("<I", rowblock[1])[0], 1)
        self.assertEqual(struct.unpack_from("<H", rowblock[1], 12)[0], 2)

    def test_missing_dimension_is_inserted_after_sheet_properties_before_views(self):
        rows = [row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                row_header(7, spans=((2, 2),)), cell(2, "number", 8)]
        for properties in (b"", record(147, bytes(23))):
            views = record(133) + record(134)
            contents = worksheet(rows, before_data=views)
            contents = contents[:len(record(129))] + properties + contents[len(record(129)):]
            data = with_indexes(package([("売上", contents)]))
            base = XLSB.Workbook(data)
            source = XLSB.Workbook(replace_part(data, SHEET_PATH,
                replace_cell(contents, 0, 0, cell(0, "number", 2))))
            with self.subTest(properties=bool(properties)):
                output, result = self.assert_rebuild(base, source, {("1", "A1")})
                types = [kind for kind, _, _, _ in records(parts(output)[SHEET_PATH])]
                self.assertEqual(types[:4] if properties else types[:3],
                                 [129, 147, 148, 133] if properties else [129, 148, 133])
                self.assertEqual(result._binary_sheets["1"].dimensions, (0, 7, 0, 2))
                self.assert_indexes(output)

    def test_missing_calculation_properties_follow_names_before_later_records(self):
        group = record(664) + record(665) + record(666)
        externals = record(353) + record(357) + record(354)
        name = record(39, struct.pack("<IBI", 0, 0, 0xFFFFFFFF) + wide("名前") + parsed_formula(ptg_int(1)))
        tail = record(549) + record(155, b"\0") + record(35, bytes(4)) + record(36)
        for preceding in (b"", group, externals, group + externals + name):
            contents = worksheet([row_header(0, spans=((0, 0),)), cell(0, "number", 1)])
            data = with_indexes(package([("売上", contents)], workbook_records=preceding + tail))
            base = XLSB.Workbook(data)
            source = XLSB.Workbook(replace_part(data, SHEET_PATH,
                replace_cell(contents, 0, 0, cell(0, "number", 2))))
            with self.subTest(preceding=preceding):
                output, result = self.assert_rebuild(base, source, {("1", "A1")})
                workbook = parts(output)["xl/workbook.bin"]
                types = [kind for kind, _, _, _ in records(workbook)]
                self.assertEqual(types[types.index(157) + 1], 549)
                self.assertLess(types.index(144), types.index(157))
                if name in preceding:
                    self.assertLess(types.index(39), types.index(157))

    def test_validation_rejects_wrong_dimension_position_or_excluded_unselected_cell(self):
        rows = [row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                row_header(7, spans=((2, 2),)), cell(2, "number", 8)]
        contents = worksheet(rows, before_data=record(133) + record(134))
        data = with_indexes(package([("売上", contents)]))
        base = XLSB.Workbook(data)
        source = XLSB.Workbook(replace_part(data, SHEET_PATH,
            replace_cell(contents, 0, 0, cell(0, "number", 2))))
        original = XLSB.Workbook._rebuild_sheet
        for corruption in ("position", "bounds"):
            def broken(book, sheet, other, coordinates):
                contents = original(book, sheet, other, coordinates)
                if corruption == "bounds":
                    return replace_record(contents, 148, lambda _: struct.pack("<IIII", 0, 0, 0, 0))
                dimension = next(contents[start:end] for kind, _, start, end in records(contents) if kind == 148)
                contents = contents.replace(dimension, b"", 1)
                begin = next(start for kind, _, start, end in records(contents) if kind == 145)
                return contents[:begin] + dimension + contents[begin:]
            with self.subTest(corruption=corruption), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {("1", "A1")})

    def test_validation_rejects_calculation_properties_after_later_records(self):
        contents = worksheet([row_header(0, spans=((0, 0),)), cell(0, "number", 1)])
        data = package([("売上", contents)], workbook_records=record(155, b"\0"))
        base = XLSB.Workbook(data)
        source = XLSB.Workbook(replace_part(data, SHEET_PATH,
            replace_cell(contents, 0, 0, cell(0, "number", 2))))
        original = XLSB.Workbook._recalculate
        def broken(book, updates, removed):
            original(book, updates, removed)
            contents = updates[book._workbook_path]
            calc = next(contents[start:end] for kind, _, start, end in records(contents) if kind == 157)
            contents = contents.replace(calc, b"", 1)
            end = next(start for kind, _, start, end in records(contents) if kind == 132)
            updates[book._workbook_path] = contents[:end] + calc + contents[end:]
        with mock.patch.object(XLSB.Workbook, "_recalculate", broken):
            with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                base.rebuild(source, {("1", "A1")})

    def test_new_row_uses_default_height_and_inherited_target_column_style(self):
        common = dict(default_height=420, columns=((0, 1, 2),))
        target = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1)], **common)
        source = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                          row_header(5, ixfe=7, spans=((0, 0),), height=900, ghost=True),
                          cell(0, "string", "new", ixfe=9)], **common)
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "A6")})
        row = rows_by_number(parts(output)[SHEET_PATH])[5]
        self.assertEqual(struct.unpack_from("<I", row, 4)[0], 0)
        self.assertEqual(struct.unpack_from("<H", row, 8)[0], 420)
        self.assertFalse(row[11] & 0x40)
        self.assertEqual(struct.unpack_from("<I", row, 13)[0], 1)
        self.assertEqual(struct.unpack_from("<II", row, 17), (0, 0))
        self.assertEqual(result.effective_style(("1", "A6")), 2)
        dimensions = next(payload for kind, payload, start, end in records(parts(output)[SHEET_PATH]) if kind == 148)
        self.assertEqual(struct.unpack("<IIII", dimensions), (0, 5, 0, 0))

    def test_effective_style_uses_explicit_cell_then_ghost_row_then_column(self):
        data = minimal([row_header(0, ixfe=4, spans=((0, 1),), ghost=True),
                        cell(0, "number", 1), cell(1, "blank", ixfe=3),
                        row_header(1, ixfe=5, spans=(), ghost=False)], columns=((0, 2, 2),))
        book = XLSB.Workbook(data)
        expected = {"A1": 0, "B1": 3, "C1": 4, "A2": 2, "D2": 0}
        for coordinate, style in expected.items():
            self.assertEqual(book.effective_style(("1", coordinate)), style, coordinate)

    def test_row_with_sixteen_spans_expands_existing_block_without_adding_a_span(self):
        spans = tuple((block * 1024, block * 1024) for block in range(16))
        target = minimal([row_header(0, spans=spans), cell(0, "number", 1)])
        source = minimal([row_header(0, spans=spans[:-1] + ((15360, 16383),)),
                          cell(0, "number", 1), cell(16383, "number", 2)])
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "XFD1")})
        row = rows_by_number(parts(output)[SHEET_PATH])[0]
        self.assertEqual(struct.unpack_from("<I", row, 13)[0], 16)
        actual_spans = tuple(struct.unpack_from("<II", row, 17 + 8 * index) for index in range(16))
        self.assertEqual(actual_spans, spans[:-1] + ((15360, 16383),))
        self.assert_indexes(output)

    def test_empty_sheet_accepts_new_first_cell_and_indexes_it(self):
        base = XLSB.Workbook(minimal([]))
        source = XLSB.Workbook(minimal([row_header(9, spans=((1024, 1024),)), cell(1024, "string", "first")]))
        output, result = self.assert_rebuild(base, source, {("1", "AMK10")})
        self.assertEqual(result.cells[("1", "AMK10")], OOXML.Cell("string", "first"))
        self.assert_indexes(output)

    def test_many_rows_rebuild_recomputes_all_index_blocks(self):
        data = (FIXTURES / "many-rows.xlsb").read_bytes()
        current = parts(data)
        sheet = replace_cell(current[SHEET_PATH], 39, 0, cell(0, "string", "a longer replacement at the first indexed row"))
        sheet = replace_cell(sheet, 44, 1029, cell(1029, "string", "replacement in the second column block"))
        base = XLSB.Workbook(data)
        source = XLSB.Workbook(replace_part(data, SHEET_PATH, sheet))
        output, result = self.assert_rebuild(base, source, {("1", "A40"), ("1", "AMP45")})
        self.assert_indexes(output)
        self.assertNotEqual(parts(output)[INDEX_PATH], current[INDEX_PATH])
        self.assertGreater(sum(kind == 42 for kind, payload, start, end in records(parts(output)[INDEX_PATH])), 3)

    def test_unknown_index_shape_is_removed_with_relationship_and_content_type(self):
        target = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1)], indexes=False)
        target = with_indexes(target, {1: record(42, b"unknown index") + record(277)})
        source = minimal([row_header(0, spans=((0, 0),)), cell(0, "string", "changed")])
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "A1")})
        current = parts(output)
        self.assertNotIn(INDEX_PATH, current)
        relpath = "xl/worksheets/_rels/sheet1.bin.rels"
        if relpath in current:
            self.assertFalse(any(node.get("Type", "").endswith("/xlBinaryIndex")
                                 for node in etree.fromstring(current[relpath])))
        self.assertFalse(any(node.get("PartName") == "/" + INDEX_PATH
                             for node in etree.fromstring(current["[Content_Types].xml"])))

    def test_array_merged_and_rich_text_selections_are_rejected(self):
        base, source = self.books()
        for coordinate, reason in (("J2", "数式の拡張機能"),
                                   ("J3", "配列・共有・データテーブル数式の範囲"),
                                   ("L1", "結合セルの左上以外"), ("E2", "リッチテキスト")):
            with self.subTest(coordinate=coordinate):
                self.assertEqual(base.unsupported_reason(("1", coordinate)), reason)
                with self.assertRaisesRegex(OOXML.WorkbookError, reason):
                    base.rebuild(source, {("1", coordinate)})

    def test_rstring_and_rich_shared_string_are_unsupported_even_when_value_is_plain(self):
        data = minimal([row_header(0, spans=((0, 1),)),
                        cell(0, "rich-string", rich_string("plain")),
                        cell(1, "shared-string", 0)],
                       shared_strings=(rich_string("rich", runs=((0, 0),)),))
        book = XLSB.Workbook(data)
        for coordinate in ("A1", "B1"):
            self.assertEqual(book.unsupported_reason(("1", coordinate)), "リッチテキスト")
            with self.assertRaises(OOXML.WorkbookError):
                book.rebuild(book, {("1", coordinate)})

    def test_shared_formula_and_data_table_ranges_are_unsupported(self):
        rgce, rgcb = ptg_exp(0, 0)
        shared = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1, rgce=rgce, rgcb=rgcb),
                          formula_range((0, 1, 0, 1), ptg_int(1)),
                          row_header(1, spans=((0, 0),)), cell(0, "number", 1, rgce=rgce, rgcb=rgcb)])
        table = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1, rgce=rgce, rgcb=rgcb),
                         record(428, struct.pack("<IIIIIIII", 0, 1, 0, 1, 2, 2, 3, 3) + b"\0")])
        for data in (shared, table):
            book = XLSB.Workbook(data)
            for coordinate in ("A1", "A2", "B1", "B2"):
                self.assertTrue(book.unsupported_reason(("1", coordinate)), coordinate)
                with self.assertRaises(OOXML.WorkbookError):
                    book.rebuild(book, {("1", coordinate)})

    def test_unparsed_formula_keeps_hash_and_is_never_copyable(self):
        tokens = b"\xff"
        data = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1, rgce=tokens)])
        book = XLSB.Workbook(data)
        value = book.cells[("1", "A1")]
        self.assertEqual(value.kind, "formula")
        self.assertTrue(value.value.startswith("="))
        self.assertIn(hashlib.sha256(tokens).hexdigest()[:8], value.value)
        self.assertTrue(book.unsupported_reason(("1", "A1")))
        with self.assertRaises(OOXML.WorkbookError):
            book.rebuild(book, {("1", "A1")})

    def test_formats_use_the_shared_comparison_and_can_be_selected(self):
        base, source = self.books()
        key = ("1", "B2")
        comparison = base.compare_formats(source)
        self.assertIsInstance(comparison, OOXML.FormatComparison)
        self.assertIn(key, set(comparison.keys()))
        self.assertTrue(comparison.differs(key))
        self.assertIsNone(comparison.unsupported_reason(key))
        result = XLSB.Workbook(base.rebuild(source, set(), {key}))
        self.assertEqual(result.cells, base.cells)
        self.assertEqual(result.stylesheet().exact(result.effective_style(key)),
                         source.stylesheet().exact(source.effective_style(key)))

    def test_compatibility_checks_external_link_parts_relationships_and_targets(self):
        external = (record(353) + record(355, wide("external")) +
                    record(362, struct.pack("<IIii", 1, 0, 0, 0)) + record(354))
        data = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1)],
                       workbook_records=external)

        def linked(*, target="file:///before.xlsx", sheet_name="Data", filename="externalLink1.bin"):
            current = parts(data)
            path = "xl/externalLinks/" + filename
            relpath = "xl/_rels/workbook.bin.rels"
            relations = etree.fromstring(current[relpath])
            etree.SubElement(relations, "{" + REL_NS + "}Relationship", {
                "Id": "external", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
                "Target": "externalLinks/" + filename,
            })
            current[relpath] = etree.tostring(relations)
            # External Link ABNF: BeginSupBook, SupTabs, EndSupBook. The
            # workbook link uses sbt=0, a RelID and a nullable unused string.
            current[path] = (record(360, struct.pack("<H", 0) + wide("rId1") + wide(None)) +
                             record(359, struct.pack("<I", 1) + wide(sheet_name)) + record(588))
            external_rels = etree.Element("{" + REL_NS + "}Relationships", nsmap={None: REL_NS})
            etree.SubElement(external_rels, "{" + REL_NS + "}Relationship", {
                "Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                "Target": target, "TargetMode": "External",
            })
            current["xl/externalLinks/_rels/" + filename + ".rels"] = etree.tostring(external_rels)
            types = etree.fromstring(current["[Content_Types].xml"])
            etree.SubElement(types, "{" + TYPE_NS + "}Override", {
                "PartName": "/" + path, "ContentType": "application/vnd.ms-excel.externalLink",
            })
            current["[Content_Types].xml"] = etree.tostring(types)
            output = io.BytesIO()
            with zipfile.ZipFile(output, "w") as archive:
                for name, contents in current.items():
                    archive.writestr(name, contents)
            return XLSB.Workbook(output.getvalue())

        base = linked()
        base.assert_compatible(linked())
        for change in ({"sheet_name": "Different"}, {"target": "file:///after.xlsx"},
                       {"filename": "externalLink2.bin"}):
            with self.subTest(change=change), self.assertRaisesRegex(XLSB.WorkbookError, "外部参照"):
                base.assert_compatible(linked(**change))

    def test_compatibility_checks_sheet_structure_date_names_and_external_references(self):
        rows = [row_header(0, spans=((0, 0),)), cell(0, "number", 1)]
        defined = lambda token: record(39, struct.pack("<IBI", 0, 0, 0xFFFFFFFF) + wide("名前") + parsed_formula(token))
        external = lambda first: record(353) + record(357) + record(362, struct.pack("<IIii", 1, 0, first, first)) + record(354)
        cases = (
            (minimal(rows), minimal(rows, name="別名"), "シートの名前・順序・内部の参照"),
            (minimal(rows), minimal(rows, date1904=True), "日付基準"),
            (minimal(rows, workbook_records=defined(ptg_int(1))), minimal(rows, workbook_records=defined(ptg_int(2))), "名前定義"),
            (minimal(rows, workbook_records=external(0)), minimal(rows, workbook_records=external(-1)), "外部参照"),
        )
        for target, source, reason in cases:
            base, other = XLSB.Workbook(target), XLSB.Workbook(source)
            with self.subTest(reason=reason), self.assertRaisesRegex(OOXML.WorkbookError, reason + "が違うブックはセル単位で反映できません"):
                base.assert_compatible(other)

    def test_invalid_selection_uses_shared_workbook_error(self):
        base, source = self.books()
        for key, reason in ((('unknown', 'A1'), "セル選択に未知のシートがあります"),
                            (('1', 'A0'), "セル番地")):
            with self.subTest(key=key), self.assertRaises(OOXML.WorkbookError):
                base.rebuild(source, {key})

    def test_source_nonminimal_record_header_is_preserved_when_copying_cell(self):
        target = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1, ixfe=3)])
        source = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 2, ixfe=7)])
        payload = struct.pack("<II", 0, 7) + struct.pack("<d", 2)
        encoded = b"\x85\x00\x90\x00" + payload
        incoming = parts(source)[SHEET_PATH]
        source = replace_part(source, SHEET_PATH, replace_cell(incoming, 0, 0, encoded))
        base, other = XLSB.Workbook(target), XLSB.Workbook(source)
        output, result = self.assert_rebuild(base, other, {("1", "A1")})
        expected = encoded[:8] + struct.pack("<I", 3) + encoded[12:]
        self.assertEqual(raw_cells(parts(output)[SHEET_PATH])[(0, 0)], expected)
        self.assert_indexes(output)

    def test_appended_cell_follows_formula_definition_after_last_existing_cell(self):
        rgce, rgcb = ptg_exp(0, 0)
        definitions = (
            formula_range((0, 0, 0, 0), ptg_int(1), array=True),
            formula_range((0, 0, 0, 0), ptg_int(1)),
            record(428, struct.pack("<IIIIIIII", 0, 0, 0, 0, 2, 2, 3, 3) + b"\0"),
        )
        for definition in definitions:
            original_row = [row_header(0, spans=((0, 0),)),
                            cell(0, "number", 1, rgce=rgce, rgcb=rgcb), definition]
            incoming_row = [row_header(0, spans=((0, 1),)),
                            cell(0, "number", 1, rgce=rgce, rgcb=rgcb), definition,
                            cell(1, "string", "new")]
            base, other = XLSB.Workbook(minimal(original_row)), XLSB.Workbook(minimal(incoming_row))
            with self.subTest(definition=next(records(definition))[0]):
                output, result = self.assert_rebuild(base, other, {("1", "B1")})
                kinds = [kind for kind, payload, start, end in records(parts(output)[SHEET_PATH])]
                anchor = kinds.index(9)
                self.assertEqual(kinds[anchor:anchor + 3], [9, next(records(definition))[0], 6])
                self.assert_indexes(output)

    def test_new_rows_are_inserted_before_following_rows_adornments(self):
        adornment = record(37, struct.pack("<HH", 1, 0)) + record(1024, struct.pack("<H", 99)) + record(38)
        original = [row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                    adornment, row_header(9, spans=((0, 0),)), cell(0, "number", 10)]
        incoming = [row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                    row_header(5, spans=((0, 0),)), cell(0, "string", "new"),
                    adornment, row_header(9, spans=((0, 0),)), cell(0, "number", 10)]
        base, source = XLSB.Workbook(minimal(original)), XLSB.Workbook(minimal(incoming))
        output, result = self.assert_rebuild(base, source, {("1", "A6")})
        encoded = parts(output)[SHEET_PATH]
        self.assertIn(adornment + row_header(9, spans=((0, 0),)), encoded)
        row6 = next(start for kind, payload, start, end in records(encoded)
                    if kind == 0 and struct.unpack_from("<I", payload)[0] == 5)
        self.assertLess(row6, encoded.index(adornment))

    def test_validation_rejects_new_row_defaults_and_moved_row_adornment(self):
        target = XLSB.Workbook(minimal([], default_height=420))
        source = XLSB.Workbook(minimal([row_header(5, spans=((0, 0),)), cell(0, "number", 2)], default_height=420))
        original = XLSB.Workbook._rebuild_sheet
        for corruption in ("height", "style", "ghost"):
            def broken(book, sheet, other, coordinates):
                contents = original(book, sheet, other, coordinates)
                def change(payload):
                    mutable = bytearray(payload)
                    if corruption == "height":
                        struct.pack_into("<H", mutable, 8, 999)
                    elif corruption == "style":
                        struct.pack_into("<I", mutable, 4, 7)
                    else:
                        mutable[11] |= 0x40
                    return bytes(mutable)
                return replace_record(contents, 0, change)
            with self.subTest(corruption=corruption), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    target.rebuild(source, {("1", "A6")})
        adornment = record(1024, struct.pack("<H", 99))
        common_rows = [row_header(0, spans=((0, 0),)), cell(0, "number", 1),
                       adornment, row_header(9, spans=((0, 0),)), cell(0, "number", 10)]
        target_data = minimal(common_rows)
        changed_sheet = replace_cell(parts(target_data)[SHEET_PATH], 0, 0, cell(0, "number", 2))
        target, source = XLSB.Workbook(target_data), XLSB.Workbook(replace_part(target_data, SHEET_PATH, changed_sheet))
        def move_adornment(book, sheet, other, coordinates):
            contents = original(book, sheet, other, coordinates)
            marker = adornment + row_header(9, spans=((0, 0),))
            if marker not in contents:
                raise AssertionError("adornment marker missing")
            return contents.replace(marker, row_header(9, spans=((0, 0),)) + adornment)
        with mock.patch.object(XLSB.Workbook, "_rebuild_sheet", move_adornment):
            with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                target.rebuild(source, {("1", "A1")})

    def test_untouched_chart_sheet_custom_parts_and_zip_metadata_are_preserved(self):
        target_data = minimal([row_header(0, spans=((0, 0),)), cell(0, "number", 1)])
        original = parts(target_data)
        descriptor = record(156, struct.pack("<II", 0, 2) + wide("chart") + wide("グラフ"))
        original["xl/workbook.bin"] = replace_record(original["xl/workbook.bin"], 144, lambda payload: payload)
        original["xl/workbook.bin"] = original["xl/workbook.bin"].replace(record(144), descriptor + record(144), 1)
        rels = original["xl/_rels/workbook.bin.rels"].decode("utf-8")
        original["xl/_rels/workbook.bin.rels"] = rels.replace("</Relationships>",
            '<Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.bin"/></Relationships>').encode("utf-8")
        original["xl/chartsheets/sheet1.bin"] = record(129) + record(3072, b"chart data") + record(130)
        original["custom/unknown.bin"] = b"\x00\xff preserved content"
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.comment = b"package comment"
            for path, contents in original.items():
                info = zipfile.ZipInfo(path, date_time=(2026, 1, 1, 0, 0, 0))
                info.comment = b"part comment"
                info.extra = b"\xfe\xca\x02\x00xy"
                archive.writestr(info, contents)
        target_data = output.getvalue()
        incoming_sheet = replace_cell(original[SHEET_PATH], 0, 0, cell(0, "number", 2))
        source_data = replace_part(target_data, SHEET_PATH, incoming_sheet)
        base, source = XLSB.Workbook(target_data), XLSB.Workbook(source_data)
        self.assertEqual([(sheet.id, sheet.name) for sheet in base.sheets], [("1", "売上")])
        updated, result = self.assert_rebuild(base, source, {("1", "A1")})
        self.assertEqual(parts(updated)["xl/chartsheets/sheet1.bin"], original["xl/chartsheets/sheet1.bin"])
        self.assertEqual(parts(updated)["custom/unknown.bin"], original["custom/unknown.bin"])
        with zipfile.ZipFile(io.BytesIO(target_data)) as before, zipfile.ZipFile(io.BytesIO(updated)) as after:
            self.assertEqual(after.namelist(), before.namelist())
            self.assertEqual(after.comment, before.comment)
            for first, last in zip(before.infolist(), after.infolist()):
                self.assertEqual((last.date_time, last.comment, last.extra, last.external_attr),
                                 (first.date_time, first.comment, first.extra, first.external_attr))

    def test_validation_rejects_deletion_of_untouched_package_parts(self):
        base, source = self.books()
        original = XLSB.Workbook._recalculate
        for path in ("docProps/core.xml", "xl/worksheets/binaryIndex2.bin"):
            self.assertIn(path, parts(base.data))
            def broken(book, updates, removed):
                original(book, updates, removed)
                removed.add(path)
            with self.subTest(path=path), mock.patch.object(XLSB.Workbook, "_recalculate", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {("1", "B2")})

    def test_validation_requires_known_index_to_be_recomputed_instead_of_removed(self):
        base, source = self.books()
        original = XLSB.Workbook._reindex
        def broken(book, sheet, updates, removed):
            original(book, sheet, updates, removed)
            updates.pop(INDEX_PATH, None)
            book._remove_relationships(sheet.path, "/xlBinaryIndex", updates, removed)
        with mock.patch.object(XLSB.Workbook, "_reindex", broken):
            with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                base.rebuild(source, {("1", "B2")})

    def test_validation_requires_source_formula_cache_tokens_and_extra_bytes(self):
        fixture_base, fixture_source = self.books()
        hint = b"\x19\x08\0\0"
        normal = [row_header(0, spans=((0, 0),)), cell(0, "number", 1, rgce=ptg_int(1))]
        hinted = [row_header(0, spans=((0, 0),)), cell(0, "number", 2, rgce=ptg_int(2) + hint)]
        array_tokens = b"\x20" + bytes(14)
        array_extra = struct.pack("<II", 1, 1) + b"\0" + struct.pack("<d", 0.0)
        array = [row_header(0, spans=((0, 0),)), cell(0, "number", 0, rgce=array_tokens, rgcb=array_extra)]
        cases = (
            (fixture_base, fixture_source, ("1", "D4"), "cache"),
            (XLSB.Workbook(minimal(normal)), XLSB.Workbook(minimal(hinted)), ("1", "A1"), "token-hint"),
            (XLSB.Workbook(minimal(normal)), XLSB.Workbook(minimal(array)), ("1", "A1"), "extra-negative-zero"),
        )
        original = XLSB.Workbook._rebuild_sheet
        for base, source, key, corruption in cases:
            def broken(book, sheet, other, coordinates):
                contents = original(book, sheet, other, coordinates)
                row, col = OOXML.coordinate_key(key[1])
                encoded = raw_cells(contents)[(row - 1, col - 1)]
                kind, payload, start, end = next(records(encoded))
                if corruption == "cache":
                    altered = payload[:8] + struct.pack("<d", 12345.0) + payload[16:]
                elif corruption == "token-hint":
                    altered = payload.replace(hint, b"\x19\x08\x01\0", 1)
                else:
                    altered = payload[:-8] + struct.pack("<d", -0.0)
                self.assertNotEqual(payload, altered)
                corrupted = replace_cell(contents, row - 1, col - 1, record(kind, altered))
                # Each mutation keeps the decoded selected formula identical.
                changed = XLSB.Workbook(replace_part(book.data, sheet.path, corrupted))
                self.assertEqual(changed.cells[key], other.cells[key])
                return corrupted
            with self.subTest(corruption=corruption), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {key})

    def test_validation_rejects_relationship_and_content_type_root_attribute_changes(self):
        base, source = self.books()
        original = XLSB.Workbook._recalculate
        for path in ("[Content_Types].xml", "xl/_rels/workbook.bin.rels"):
            def broken(book, updates, removed):
                original(book, updates, removed)
                tree = etree.fromstring(updates.get(path, book._parts[path]))
                tree.set("changed", "unexpected")
                updates[path] = etree.tostring(tree)
            with self.subTest(path=path), mock.patch.object(XLSB.Workbook, "_recalculate", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {("1", "B2")})

    def test_new_cell_is_inserted_before_metadata_of_the_following_cell(self):
        for kind in (49, 50):
            metadata = record(kind, struct.pack("<I", 1))
            target = minimal([row_header(0, spans=((0, 1),)), metadata, cell(1, "number", 2)])
            incoming = minimal([row_header(0, spans=((0, 1),)), cell(0, "number", 1), metadata, cell(1, "number", 2)])
            base, source = XLSB.Workbook(target), XLSB.Workbook(incoming)
            with self.subTest(kind=kind):
                output, result = self.assert_rebuild(base, source, {("1", "A1")})
                encoded = parts(output)[SHEET_PATH]
                self.assertIn(cell(0, "number", 1) + metadata + cell(1, "number", 2), encoded)
                self.assertIsNone(result.unsupported_reason(("1", "A1")))
                self.assertEqual(result.unsupported_reason(("1", "B1")), "セルのメタデータ")
                self.assert_indexes(output)

    def test_cell_with_metadata_cannot_be_copied_or_deleted(self):
        for kind in (49, 50):
            metadata = record(kind, struct.pack("<I", 1))
            tagged = XLSB.Workbook(minimal([row_header(0, spans=((0, 1),)), metadata, cell(1, "number", 2)]))
            plain = XLSB.Workbook(minimal([row_header(0, spans=((0, 1),)), cell(1, "number", 3)]))
            empty = XLSB.Workbook(minimal([row_header(0, spans=((0, 1),))]))
            self.assertEqual(tagged.unsupported_reason(("1", "B1")), "セルのメタデータ")
            for target, source in ((tagged, plain), (plain, tagged), (tagged, empty)):
                with self.subTest(kind=kind, deletion=source is empty):
                    with self.assertRaisesRegex(OOXML.WorkbookError, "セルのメタデータ"):
                        target.rebuild(source, {("1", "B1")})

    def test_validation_rejects_metadata_moved_to_another_cell(self):
        original = XLSB.Workbook._rebuild_sheet
        for kind in (49, 50):
            metadata = record(kind, struct.pack("<I", 1))
            target = minimal([row_header(0, spans=((0, 1),)), metadata, cell(1, "number", 2)])
            incoming = minimal([row_header(0, spans=((0, 1),)), cell(0, "number", 1), metadata, cell(1, "number", 2)])
            base, source = XLSB.Workbook(target), XLSB.Workbook(incoming)
            def broken(book, sheet, other, coordinates):
                contents = original(book, sheet, other, coordinates)
                marker = cell(0, "number", 1) + metadata + cell(1, "number", 2)
                self.assertIn(marker, contents)
                corrupted = contents.replace(marker, metadata + cell(0, "number", 1) + cell(1, "number", 2))
                actual = XLSB.Workbook(replace_part(book.data, sheet.path, corrupted))
                self.assertEqual(actual.cells, other.cells)
                return corrupted
            with self.subTest(kind=kind), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {("1", "A1")})

    def test_validation_detects_unselected_record_row_and_noncell_corruption(self):
        base, source = self.books()
        original = XLSB.Workbook._rebuild_sheet
        for corruption in ("unselected", "row", "noncell", "selected-style", "missing-selected"):
            def broken(book, sheet, other, coordinates):
                contents = original(book, sheet, other, coordinates)
                if corruption == "unselected":
                    encoded = raw_cells(contents)[(0, 0)]
                    kind, payload, start, end = next(records(encoded))
                    altered = payload[:4] + struct.pack("<I", 9) + payload[8:]
                    return replace_cell(contents, 0, 0, record(kind, altered))
                if corruption == "row":
                    return replace_record(contents, 0, lambda payload: payload[:8] + struct.pack("<H", 999) + payload[10:])
                if corruption == "noncell":
                    return replace_record(contents, 485, lambda payload: payload[:-1] + bytes([payload[-1] ^ 1]))
                if corruption == "selected-style":
                    encoded = raw_cells(contents)[(1, 1)]
                    kind, payload, start, end = next(records(encoded))
                    return replace_cell(contents, 1, 1, record(kind, payload[:4] + struct.pack("<I", 9) + payload[8:]))
                return replace_cell(contents, 1, 1, b"")

            with self.subTest(corruption=corruption), mock.patch.object(XLSB.Workbook, "_rebuild_sheet", broken):
                with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                    base.rebuild(source, {("1", "B2")})


if __name__ == "__main__":
    unittest.main()
