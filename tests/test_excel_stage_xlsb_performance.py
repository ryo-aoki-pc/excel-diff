"""Guard bounded xlsb validation and reusable decoding without timing assertions."""

from collections import Counter
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

from xlsb_builder import (
    cell, package, parsed_formula, ptg_int, ptg_ref, record, recompute_index,
    replace_part, row_header, wide, with_indexes, worksheet,
)

SHEET_PATH = "xl/worksheets/sheet1.bin"
SECOND_PATH = "xl/worksheets/sheet2.bin"
INDEX_PATH = "xl/worksheets/binaryIndex1.bin"
VALIDATION_ERROR = "作り直したブックの検証に失敗しました。書き込みは行っていません"


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def sheet(value):
    return worksheet([row_header(0, spans=((0, 0),)), cell(0, "number", value)],
                     dimensions=(0, 0, 0, 0))


def pair(*, indexed=True):
    target = package([("変更", sheet(1)), ("維持", sheet(7))],
                     workbook_records=record(157, bytes(26)))
    if indexed:
        target = with_indexes(target)
    source = replace_part(target, SHEET_PATH, sheet(2))
    return XLSB.Workbook(target), XLSB.Workbook(source)


class XlsbPerformanceRegressionTests(unittest.TestCase):
    def test_repeated_position_independent_formula_decodes_once(self):
        # The integer contains 0x2c in its payload; only token boundaries make
        # it a RefN, so a byte search must not disable otherwise safe reuse.
        tokens = ptg_ref(1, 1) + ptg_int(17964) + b"\x03"
        rows = []
        for row in (1, 3, 5):
            rows.extend((row_header(row, spans=((1, 1),)),
                         cell(1, "number", 0, rgce=tokens)))
        with mock.patch.object(XLSB, "decode_formula", wraps=XLSB.decode_formula) as decode:
            book = XLSB.Workbook(package([("再利用", worksheet(rows))]))
        for coordinate in ("B2", "B4", "B6"):
            self.assertEqual(book.cells[("1", coordinate)], OOXML.Cell("formula", "=B2+17964"))
        self.assertLessEqual(decode.call_count, 1)

    def test_identical_refn_and_arean_resolve_at_each_cell_position(self):
        area = b"\x4d" + struct.pack("<IIHH", 0, 1, 0xC000, 0xC001)
        cases = ((ptg_ref(0, 0, shared=True), ("=B2", "=D4")),
                 (area, ("=B2:C3", "=D4:E5")))
        for tokens, expected in cases:
            with self.subTest(tokens=tokens):
                rows = [row_header(1, spans=((1, 1),)), cell(1, "number", 0, rgce=tokens),
                        row_header(3, spans=((3, 3),)), cell(3, "number", 0, rgce=tokens)]
                book = XLSB.Workbook(package([("相対参照", worksheet(rows))]))
                self.assertEqual([book.cells[("1", coordinate)].value for coordinate in ("B2", "D4")],
                                 list(expected))

    def test_attribute_scan_follows_decoder_precedence_before_relative_reference(self):
        # Space attributes are consumed before Choose hints by the decoder,
        # even when flags overlap. Their operands must not hide a later RefN.
        tokens = ptg_int(1) + b"\x19" + struct.pack("<BH", 0x44, 3) + ptg_ref(0, 0, shared=True) + b"\x03"
        rows = [row_header(1, spans=((1, 1),)), cell(1, "number", 0, rgce=tokens),
                row_header(3, spans=((3, 3),)), cell(3, "number", 0, rgce=tokens)]
        book = XLSB.Workbook(package([("属性", worksheet(rows))]))
        self.assertEqual(book.cells[("1", "B2")], OOXML.Cell("formula", "=1+B2"))
        self.assertEqual(book.cells[("1", "D4")], OOXML.Cell("formula", "=1+D4"))

    def test_array_constant_cache_includes_extra_formula_data(self):
        tokens = b"\x60" + bytes(14)
        extra = lambda value: struct.pack("<II", 1, 1) + b"\0" + struct.pack("<d", value)
        rows = [row_header(0, spans=((0, 1),)),
                cell(0, "number", 0, rgce=tokens, rgcb=extra(1)),
                cell(1, "number", 0, rgce=tokens, rgcb=extra(2))]
        book = XLSB.Workbook(package([("配列定数", worksheet(rows))]))
        self.assertEqual(book.cells[("1", "A1")], OOXML.Cell("formula", "={1}"))
        self.assertEqual(book.cells[("1", "B1")], OOXML.Cell("formula", "={2}"))

    def test_formula_reuse_keeps_sheet_local_name_context(self):
        tokens = b"\x23" + struct.pack("<I", 1)
        formula_sheet = worksheet([row_header(0, spans=((0, 0),)),
                                   cell(0, "number", 0, rgce=tokens)])
        name = record(39, struct.pack("<IBI", 0, 0, 0) + wide("局所") + parsed_formula(ptg_int(3)))
        book = XLSB.Workbook(package([("一枚目", formula_sheet), ("二枚目", formula_sheet)],
                                     workbook_records=name))
        self.assertEqual(book.cells[("1", "A1")], OOXML.Cell("formula", "=局所"))
        self.assertEqual(book.cells[("2", "A1")], OOXML.Cell("formula", "=一枚目!局所"))

    def test_fast_integer_numbers_keep_the_shared_decimal_representation(self):
        values = (-0.0, -123.0, 12.0, 0.1, 1e-7, 1e15, 1e16, 1e20,
                  float(2 ** 53 + 2), float(2 ** 63), 1.234567890123456e30)
        rows = [row_header(0, spans=((0, len(values) + 1),))]
        rows.extend(cell(column, "number", value) for column, value in enumerate(values))
        rows.extend((cell(len(values), "rk", (123 << 2) | 2),
                     cell(len(values) + 1, "rk", ((-123 & 0x3FFFFFFF) << 2) | 2)))
        book = XLSB.Workbook(package([("数値", worksheet(rows))]))
        expected = [OOXML._number(repr(value)) for value in values] + ["123", "-123"]
        self.assertEqual([value.value for value in book.cells.values()], expected)

    def test_validation_does_not_decode_unchanged_sheet_and_keeps_its_bytes(self):
        base, source = pair()
        read_sheet = XLSB._BinaryWorkbook._read_sheet
        decoded = []
        def read_changed_only(book, descriptor, data):
            decoded.append(descriptor.path)
            if descriptor.path == SECOND_PATH:
                self.fail("unchanged worksheet was decoded again during validation")
            return read_sheet(book, descriptor, data)
        with mock.patch.object(XLSB._BinaryWorkbook, "_read_sheet", read_changed_only):
            output = base.rebuild(source, {("1", "A1")})
        self.assertEqual(parts(output)[SECOND_PATH], parts(base.data)[SECOND_PATH])
        self.assertNotIn(SECOND_PATH, decoded)
        self.assertEqual(XLSB.Workbook(output).cells[("1", "A1")], OOXML.Cell("number", "2"))

    def test_selected_raw_value_is_checked_against_the_source_cell_model(self):
        base, source = pair()
        # A bad read/model result must not be accepted just because the writer
        # faithfully copies the bytes belonging to that incorrect model.
        source.cells[("1", "A1")] = OOXML.Cell("number", "3")
        with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
            base.rebuild(source, {("1", "A1")})

    def test_raw_unchanged_sheet_check_rejects_corruption_in_updates(self):
        base, source = pair()
        recalculate = XLSB.Workbook._recalculate
        def corrupt_untouched(book, updates, removed):
            recalculate(book, updates, removed)
            # Also put the corrupt bytes in updates, so ZIP-versus-updates
            # agreement alone cannot establish that this part is untouched.
            updates[SECOND_PATH] = sheet(8)
        with mock.patch.object(XLSB.Workbook, "_recalculate", corrupt_untouched):
            with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                base.rebuild(source, {("1", "A1")})

    def test_original_and_rebuilt_index_are_each_computed_at_most_once(self):
        base, source = pair()
        counts = Counter()
        recompute = XLSB.recompute_index
        def counted(data):
            counts[data] += 1
            return recompute(data)
        with mock.patch.object(XLSB, "recompute_index", counted):
            output = base.rebuild(source, {("1", "A1")})
        current = parts(output)
        self.assertEqual(current[INDEX_PATH], recompute_index(current[SHEET_PATH]))
        self.assertTrue(counts)
        self.assertTrue(all(count <= 1 for count in counts.values()), counts)
        self.assertLessEqual(sum(counts.values()), 2)

    def test_cached_result_index_does_not_hide_a_corrupt_written_index(self):
        base, source = pair()
        reindex = XLSB.Workbook._reindex
        def corrupt_index(book, descriptor, updates, removed):
            reindex(book, descriptor, updates, removed)
            updates[INDEX_PATH] = record(42, bytes(24)) + record(277)
        with mock.patch.object(XLSB.Workbook, "_reindex", corrupt_index):
            with self.assertRaisesRegex(OOXML.WorkbookError, VALIDATION_ERROR):
                base.rebuild(source, {("1", "A1")})


if __name__ == "__main__":
    unittest.main()
