"""Packed BIFF12 sheet views preserve the public read-only cell contract."""

import io
from pathlib import Path
import struct
import sys
import unittest
from unittest import mock
import zipfile

from xlsb_builder import cell, package, record, replace_part, rich_string, row_header, worksheet

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import excel_xlsb as xlsb


class CompactBinaryWorkbookTests(unittest.TestCase):
    def assert_same_sheet(self, original, compact):
        fields = ("type", "start", "payload_start", "end", "payload")
        self.assertEqual([tuple(getattr(value, field) for field in fields) for value in original.records],
                         [tuple(getattr(value, field) for field in fields) for value in compact.records])
        fields = ("row", "col", "kind", "value", "ixfe", "start", "end", "rgce", "rgcb", "rich", "coordinate")
        self.assertEqual([tuple(getattr(value, field) for field in fields) for value in original.cells],
                         [tuple(getattr(value, field) for field in fields) for value in compact.cells])
        self.assertEqual([(value.row, value.ixfe, value.spans) for value in original.rows],
                         [(value.row, value.ixfe, value.spans) for value in compact.rows])
        self.assertEqual(original.dimensions, compact.dimensions)
        for value in compact.cells:
            self.assertEqual(value.read(), (value.kind, value.value, value.rich, value.rgce, value.rgcb))

    def test_fixture_views_preserve_payloads_offsets_values_and_formula_bytes(self):
        for name in ("sales", "sales-edited", "many-rows"):
            with self.subTest(name=name):
                data = (ROOT / "tests" / "fixtures" / (name + ".xlsb")).read_bytes()
                original, compact = xlsb.BinaryWorkbook(data), xlsb.BinaryWorkbook(data, compact=True)
                self.assertEqual([sheet.name for sheet in original.sheets], [sheet.name for sheet in compact.sheets])
                for old, new in zip(original.sheets, compact.sheets):
                    self.assert_same_sheet(old, new)

    def test_values_use_the_original_sheet_without_materializing_cell_dataclasses(self):
        data = package([("値", worksheet([
            row_header(0), cell(0, "rk", 6), cell(1, "rk", ((-3 << 2) | 2) & 0xFFFFFFFF),
            cell(2, "rk", 12345 << 2 | 3), cell(3, "number", -0.0), cell(4, "boolean", True),
            cell(5, "error", 42), cell(6, "string", "日本語"), cell(7, "shared-string", 0),
            cell(8, "rich-string", rich_string("直接", runs=((0, 1),))), cell(9, "blank"),
            row_header(2), cell(2, "string", "cached", rgce=b"\x1e\x01\x00", rgcb=b"extra"),
            cell(0, "number", 7, rgce=b"\x1e\x01\x00"),
        ]))], shared_strings=[rich_string("共有", phonetic="キョウユウ")])
        original = xlsb.BinaryWorkbook(data)
        with mock.patch.object(xlsb, "Cell", side_effect=AssertionError("cell dataclass")):
            compact = xlsb.BinaryWorkbook(data, compact=True)
        self.assert_same_sheet(original.sheets[0], compact.sheets[0])
        self.assertEqual(compact.sheets[0].cells[-1].coordinate, "C3")
        self.assertEqual(compact.sheets[0].cells[0:2][1].value, -3)
        with self.assertRaises(IndexError):
            compact.sheets[0].cells[len(compact.sheets[0].cells)]

    def test_scalar_iteration_preserves_values_without_cell_or_record_views(self):
        data = (ROOT / "tests" / "fixtures" / "sales.xlsb").read_bytes()
        book = xlsb.BinaryWorkbook(data, compact=True)
        expected = [[(value.row, value.col, value.ixfe, value.record.type, value.start,
                      value.kind, value.value, value.rich, value.rgce, value.rgcb, index)
                     for index, value in enumerate(sheet.cells)] for sheet in book.sheets]
        with mock.patch.object(xlsb, "_CellView", side_effect=AssertionError("cell view")), \
                mock.patch.object(xlsb, "_RecordView", side_effect=AssertionError("record view")):
            actual = [list(sheet.cells.iter_values()) for sheet in book.sheets]
        self.assertEqual(actual, expected)

    def test_shared_parts_are_reused_and_untouched_sheets_are_not_parsed(self):
        data = package([("変更", worksheet([row_header(0), cell(0, "number", 1)])),
                        ("保持", worksheet([row_header(0), cell(0, "number", 2)]))])
        data = replace_part(data, "xl/worksheets/sheet2.bin", b"broken")
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        with mock.patch.object(zipfile, "ZipFile", side_effect=AssertionError("parts decompressed twice")):
            book = xlsb.BinaryWorkbook(data, parts=parts, sheet_paths={"xl/worksheets/sheet1.bin"}, compact=True)
        self.assertIs(book.parts, parts)
        self.assertEqual(book.sheets[0].cells[0].value, 1)
        self.assertEqual(book.sheets[1].cells, [])
        self.assertEqual([sheet.name for sheet in book.sheets], ["変更", "保持"])
        self.assertIs(book.sheets[0].records.data, parts["xl/worksheets/sheet1.bin"])

    def test_malformed_compact_records_fail_before_any_cell_is_accessed(self):
        malformed = [b"\x80", b"\x80\x80\x00", b"\x00\x80", b"\x00\x80\x80\x80\x80", b"\x00\x02x",
                     worksheet([cell(0, "number", 1)]), worksheet([row_header(1048576)]),
                     worksheet([row_header(0), cell(16384, "number", 1)]),
                     worksheet([row_header(0), cell(0, "shared-string", 999)]),
                     worksheet([row_header(0), cell(0, "number", 1), cell(0, "number", 2)]),
                     worksheet([row_header(0), cell(0, "number", float("inf"))]),
                     worksheet([row_header(0), cell(0, "rk", 0x7FF00000)]),
                     worksheet([row_header(0), record(6, struct.pack("<III", 0, 0, 3) + b"x")]),
                     worksheet([row_header(0), record(6, struct.pack("<IIIH", 0, 0, 1, 0xD800))]),
                     worksheet([row_header(0), record(9, struct.pack("<IIdHI", 0, 0, 1, 0, 20) + b"x")]),
                     worksheet([row_header(0), record(9, struct.pack("<IIdHII", 0, 0, 1, 0, 0, 20))])]
        for sheet in malformed:
            with self.subTest(sheet=sheet[:40]), self.assertRaises(xlsb.WorkbookError):
                xlsb.BinaryWorkbook(package([("破損", sheet)]), compact=True)


if __name__ == "__main__":
    unittest.main()
