"""BIFF12 reading, fixture parity, malformed input, and Git textconv tests."""

import hashlib
import importlib.util
import io
from pathlib import Path
import shutil
import struct
import sys
import unittest
from unittest import mock
import zipfile

import test_excel_textconv as textconv_tests
from xlsb_builder import cell, formula_range, package, parsed_formula, ptg_exp, ptg_int, ptg_ref, record, records, replace_part, rich_string, row_header, wide, worksheet


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import excel_xlsb as xlsb

_SPEC = importlib.util.spec_from_file_location("xlsb_textconv", ROOT / "scripts" / "excel-textconv.py")
textconv = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(textconv)
FIXTURES = ROOT / "tests" / "fixtures"


class BinaryRecordTests(unittest.TestCase):
    def test_record_headers_and_offsets_at_variable_length_boundaries(self):
        for kind in (0, 127, 128, 255, 16383):
            for size in (0, 1, 127, 128, 16383, 16384):
                with self.subTest(kind=kind, size=size):
                    payload = b"x" * size
                    encoded = record(kind, payload)
                    self.assertEqual(xlsb.record_header(kind, size) + payload, encoded)
                    result, = xlsb.records(encoded)
                    self.assertEqual(list(xlsb._stream_records(io.BytesIO(encoded))), [(kind, payload)])
                    self.assertEqual((result.type, result.start, result.payload_start, result.end, result.payload), (kind, 0, len(encoded) - size, len(encoded), payload))
        joined = record(127, b"one") + record(128, b"two")
        first, second = xlsb.records(joined)
        self.assertEqual(first.end, second.start)
        self.assertEqual(second.end, len(joined))
        for size in (2097151, 2097152, 0x0FFFFFFF):
            self.assertEqual(len(xlsb.record_header(128, size)), 5 if size < 2097152 else 6)

    def test_truncated_and_oversized_record_headers_fail(self):
        for data in (b"\x80", b"\x80\x80\0", b"\0", b"\0\x80", b"\0\x80\x80\x80\x80", b"\0\x02x"):
            with self.subTest(data=data), self.assertRaises(xlsb.WorkbookError):
                list(xlsb.records(data))
            with self.subTest(stream=data), self.assertRaises(xlsb.WorkbookError):
                list(xlsb._stream_records(io.BytesIO(data)))
        for kind, size in ((-1, 0), (16384, 0), (0, -1), (0, 0x10000000)):
            with self.assertRaises(xlsb.WorkbookError):
                xlsb.record_header(kind, size)

    def test_filtered_stream_discards_large_unknown_records_in_bounded_chunks(self):
        requested = []

        class TrackedStream(io.BytesIO):
            def read(self, size=-1):
                requested.append(size)
                return super().read(size)

        data = record(3072, b"x" * 200000) + row_header(0)
        for include in ({0}, xlsb._CELL_TYPES | {0}):
            with self.subTest(include=include):
                requested.clear()
                self.assertEqual(list(xlsb._stream_records(TrackedStream(data), include)), [(0, next(xlsb.records(row_header(0))).payload)])
                self.assertLessEqual(max(requested), 65536)

    def test_cells_retain_original_offsets_kinds_styles_and_tokens(self):
        tokens = ptg_int(12)
        sheet = worksheet([row_header(39, 7, ((0, 2), (1024, 1024))), cell(0, "number", 12, ixfe=0, rgce=tokens), cell(2, "blank", ixfe=0), cell(1024, "string", "遠い列")])
        data = package([("位置", sheet)], workbook_records=record(157, bytes(26)))
        book = xlsb.read_workbook(data)
        self.assertEqual(book.calc_record.type, 157)
        result = book.sheets[0]
        self.assertEqual(result.rows[0].spans, ((0, 2), (1024, 1024)))
        formula, blank, distant = result.cells
        self.assertEqual((formula.row, formula.col, formula.kind, formula.value, formula.ixfe, formula.rgce), (39, 0, "number", 12, 0, tokens))
        self.assertEqual(sheet[formula.start:formula.end], record(9, formula.record.payload))
        self.assertEqual((blank.kind, distant.coordinate), ("blank", "AMK40"))
        self.assertEqual(book.parts[result.path], sheet)


class BinaryWorkbookTests(unittest.TestCase):
    def test_excel_generated_fixtures_match_xlsx_byte_for_byte(self):
        for name in ("sales", "sales-edited", "phonetic", "phonetic-edited"):
            with self.subTest(name=name):
                data = (FIXTURES / (name + ".xlsb")).read_bytes()
                self.assertTrue(xlsb.is_xlsb(data))
                expected = textconv.convert((FIXTURES / (name + ".xlsx")).read_bytes())
                self.assertEqual(xlsb.render_text(data), expected)
                self.assertEqual(textconv.convert(data), expected)
                titles = ["ふりがな"] if name.startswith("phonetic") else ["売上", "備考"]
                self.assertEqual(xlsb.sheet_titles(data), titles)
        book = xlsb.read_workbook((FIXTURES / "sales.xlsb").read_bytes())
        cells = {cell.coordinate: cell for cell in book.sheets[0].cells}
        self.assertTrue(cells["E2"].rich)
        self.assertFalse(cells["E3"].rich)
        self.assertEqual(cells["E3"].value, "売上")
        self.assertEqual([(name.name, name.scope) for name in book.names], [("_xlfn.IFERROR", 0xFFFFFFFF), ("税率", 0xFFFFFFFF)])
        self.assertEqual(len(book.sheets[0].merges), 1)

    def test_phonetic_fixture_contains_multiple_and_implicit_runs(self):
        with zipfile.ZipFile(FIXTURES / "phonetic.xlsb") as archive:
            strings = archive.read("xl/sharedStrings.bin")
        counts = []
        for kind, payload, start, end in records(strings):
            if kind != 19 or not payload[0] & 2:
                continue
            position = 1
            position += 4 + struct.unpack_from("<I", payload, position)[0] * 2
            if payload[0] & 1:
                position += 4 + struct.unpack_from("<I", payload, position)[0] * 4
            position += 4 + struct.unpack_from("<I", payload, position)[0] * 2
            count = struct.unpack_from("<I", payload, position)[0]
            self.assertEqual(len(payload) - position - 4, count * 6 + 4)
            counts.append(count)
        self.assertIn(0, counts)
        self.assertTrue(any(count >= 2 for count in counts))

    def test_many_rows_and_columns_are_in_coordinate_order(self):
        text = xlsb.render_text((FIXTURES / "many-rows.xlsb").read_bytes())
        expected = ['Sheet: "Data"']
        for row in range(40, 201):
            if row % 7 == 0:
                continue
            expected.append(f"A{row}\t{row}")
            if row % 3 == 0:
                expected.append(f'AMP{row}\t"x{row}"')
            if row % 11 == 0:
                expected.append(f'GJH{row}\tformula: "=A{row}*2"')
        self.assertEqual(text, "\n".join(expected) + "\n")

    def test_all_value_records_and_rich_phonetic_strings(self):
        sheet = worksheet([row_header(0), cell(0, "rk", (123 << 2) | 2), cell(1, "rk", ((-3 << 2) | 2) & 0xFFFFFFFF), cell(2, "rk", (12345 << 2) | 3), cell(3, "number", 1.25), cell(4, "string", "日本語\n\t"), cell(5, "boolean", True), cell(6, "error", 42), cell(7, "shared-string", 0), cell(8, "shared-string", 1), cell(9, "rich-string", rich_string("直接", runs=((0, 1),))), cell(10, "blank")])
        data = package([("値", sheet)], shared_strings=[rich_string("書式", runs=((0, 1),)), rich_string("売上", phonetic="ウリアゲ")])
        self.assertEqual(xlsb.render_text(data), 'Sheet: "値"\nA1\t123\nB1\t-3\nC1\t123.45\nD1\t1.25\nE1\t"日本語\\n\\t"\nF1\ttrue\nG1\terror: "#N/A"\nH1\t"書式"\nI1\t"売上"\nJ1\t"直接"\n')
        cells = xlsb.read_workbook(data).sheets[0].cells
        self.assertTrue(cells[7].rich)
        self.assertFalse(cells[8].rich)

    def test_excel_phonetic_runs_match_payloads_and_are_ignored(self):
        shared = rich_string("今日は東京へ行く", phonetic="キョウトウキョウイ",
                             phonetic_runs=((0, 0, 2), (3, 3, 2), (8, 6, 1)))
        inline = rich_string("第1回東京会議の議事録", runs=((3, 2), (5, 0)),
                             phonetic="ダイカイトウキョウカイギギジロク",
                             phonetic_runs=((0, 0, 1), (2, 2, 1), (4, 3, 2), (9, 5, 2), (12, 8, 3)))
        plain_inline = rich_string("東京と大阪", phonetic="トウキョウオオサカ",
                                   phonetic_runs=((0, 0, 2), (5, 3, 2)))
        self.assertEqual(shared, bytes.fromhex(
            "0208000000ca4ee5656f307167ac4e78304c884f3009000000ad30e730a630c830a630ad30e730a630a430"
            "0300000000000000020003000300020008000600010001003500"))
        self.assertEqual(inline, bytes.fromhex(
            "030b0000002c7b3100de567167ac4e1a4f708b6e30708b8b4e3293020000000300020005000000"
            "10000000c030a430ab30a430c830a630ad30e730a630ab30a430ae30ae30b830ed30af30"
            "050000000000000001000200020001000400030002000900050002000c000800030001003500"))
        data = package([("ふりがな", worksheet([
            row_header(0), cell(0, "shared-string", 0), cell(1, "rich-string", inline),
            cell(2, "rich-string", plain_inline),
        ]))], shared_strings=[shared])
        expected = ('Sheet: "ふりがな"\nA1\t"今日は東京へ行く"\n'
                    'B1\t"第1回東京会議の議事録"\nC1\t"東京と大阪"\n')
        self.assertEqual(xlsb.render_text(data), expected)
        self.assertEqual(textconv.convert(data), expected)
        for book in (xlsb.read_workbook(data), xlsb.BinaryWorkbook(data, compact=True)):
            with self.subTest(compact=isinstance(book.sheets[0].cells, xlsb._CellSequence)):
                values = [(item.value, item.rich) for item in book.sheets[0].cells]
                self.assertEqual(values, [("今日は東京へ行く", False), ("第1回東京会議の議事録", True),
                                          ("東京と大阪", False)])

    def test_specification_phonetic_runs_without_trailing_settings_are_read(self):
        for runs in ((), ((0, 0, 2, 1, 0x35), (5, 3, 2, 1, 0x35))):
            with self.subTest(runs=runs):
                payload = b"\x02" + wide("東京と大阪") + wide("トウキョウオオサカ")
                payload += struct.pack("<I", len(runs))
                payload += b"".join(struct.pack("<HHHHI", *run) for run in runs)
                data = package([("ふりがな", worksheet([
                    row_header(0), cell(0, "shared-string", 0), cell(1, "rich-string", payload),
                ]))], shared_strings=[payload])
                expected = 'Sheet: "ふりがな"\nA1\t"東京と大阪"\nB1\t"東京と大阪"\n'
                self.assertEqual(xlsb.render_text(data), expected)
                self.assertEqual(textconv.convert(data), expected)
                for book in (xlsb.read_workbook(data), xlsb.BinaryWorkbook(data, compact=True)):
                    self.assertEqual([(item.value, item.rich) for item in book.sheets[0].cells],
                                     [("東京と大阪", False), ("東京と大阪", False)])

    def test_malformed_phonetic_run_lengths_fail_with_phonetic_message(self):
        payload = rich_string("東京と大阪", phonetic="トウキョウオオサカ",
                              phonetic_runs=((0, 0, 2), (5, 3, 2)))
        count_position = 1 + len(wide("東京と大阪")) + len(wide("トウキョウオオサカ"))
        too_many = payload[:count_position] + struct.pack("<I", 3) + payload[count_position + 4:]
        runs = ((0, 0, 2, 1, 0x35), (5, 3, 2, 1, 0x35))
        specification = b"\x02" + wide("東京と大阪") + wide("トウキョウオオサカ") + struct.pack("<I", len(runs))
        specification += b"".join(struct.pack("<HHHHI", *run) for run in runs)
        for damaged in (too_many, payload[:-4], payload + b"\0", specification + b"\0", specification[:-1]):
            for kind in ("shared-string", "rich-string"):
                with self.subTest(payload=damaged, kind=kind):
                    data = package([("ふりがな", worksheet([
                        row_header(0), cell(0, kind, 0 if kind == "shared-string" else damaged),
                    ]))], shared_strings=[damaged] if kind == "shared-string" else [])
                    for read in (xlsb.render_text, xlsb.read_workbook, textconv.convert,
                                 lambda value: xlsb.BinaryWorkbook(value, compact=True)):
                        with self.assertRaisesRegex(xlsb.WorkbookError, "ふりがな"):
                            read(data)

    def test_dates_times_durations_and_1904_epoch_match_openpyxl(self):
        sheet = worksheet([row_header(0), cell(0, "number", 1, ixfe=1), cell(1, "number", 0.5, ixfe=2), cell(2, "number", 1.5, ixfe=3), cell(3, "number", 1, ixfe=0)])
        data = package([("日時", sheet)], formats={164: "yyyy-mm-dd", 165: "[h]:mm:ss"}, cell_formats=(0, 164, 21, 165), date1904=True)
        self.assertEqual(xlsb.render_text(data), 'Sheet: "日時"\nA1\tdatetime: 1904-01-02T00:00:00\nB1\ttime: 12:00:00\nC1\tduration-seconds: 129600.0\nD1\t1\n')

    def test_shared_and_array_formulas_use_anchor_and_relative_cell(self):
        exp, extra = ptg_exp(0, 1)
        shared = worksheet([row_header(0), cell(1, "number", 2, rgce=exp, rgcb=extra), formula_range((0, 1, 1, 1), ptg_ref(0, -1, shared=True) + ptg_int(2) + b"\x05"), row_header(1), cell(1, "number", 4, rgce=exp, rgcb=extra)])
        data = package([("共有", shared)])
        self.assertEqual(xlsb.render_text(data), 'Sheet: "共有"\nB1\tformula: "=A1*2"\nB2\tformula: "=A2*2"\n')
        array = worksheet([row_header(0), cell(1, "number", 2, rgce=exp, rgcb=extra), formula_range((0, 1, 1, 1), ptg_int(2), array=True), row_header(1), cell(1, "number", 2, rgce=exp, rgcb=extra)])
        self.assertEqual(xlsb.render_text(package([("配列", array)])), 'Sheet: "配列"\nB1\tarray-formula: {"text": "=2", "ref": "B1:B2"}\nB2\t2\n')

    def test_shared_formula_anchor_can_be_inside_the_range(self):
        exp, extra = ptg_exp(0, 2)
        shared = worksheet([row_header(0), cell(2, "number", 2, rgce=exp, rgcb=extra), formula_range((0, 1, 1, 2), ptg_ref(0, -1, shared=True) + ptg_int(2) + b"\x05"), row_header(1), cell(1, "number", 4, rgce=exp, rgcb=extra)])
        self.assertEqual(xlsb.render_text(package([("共有", shared)])), 'Sheet: "共有"\nC1\tformula: "=B1*2"\nB2\tformula: "=A2*2"\n')

    def test_data_table_properties_match_excel_xlsx_dictionary(self):
        cases = [
            (0, '"dt2D": "0", "dtr": "0", "r1": "A1"'),
            (1, '"dt2D": "0", "dtr": "1", "r1": "A1"'),
            (2, '"dt2D": "1", "dtr": "1", "r1": "A1", "r2": "D1"'),
            (3, '"dt2D": "1", "dtr": "1", "r1": "A1", "r2": "D1"'),
            (4, '"dt2D": "0", "dtr": "0", "del1": "1"'),
            (10, '"dt2D": "1", "dtr": "1", "r1": "A1", "del2": "1"'),
            (16, '"dt2D": "0", "dtr": "0", "r1": "A1", "ca": "1"'),
        ]
        for flags, properties in cases:
            with self.subTest(flags=flags):
                table = record(428, struct.pack("<IIIIIIII B", 1, 2, 1, 2, 0, 0, 0, 3, flags))
                sheet = worksheet([row_header(1), cell(1, "number", 10), table, cell(2, "number", 20), row_header(2), cell(1, "number", 30), cell(2, "number", 40)])
                expected = 'Sheet: "テーブル"\nB2\tdata-table-formula: {"t": "dataTable", "ref": "B2:C3", ' + properties + '}\nC2\t20\nB3\t30\nC3\t40\n'
                self.assertEqual(xlsb.render_text(package([("テーブル", sheet)])), expected)

    def test_later_shared_definition_resolves_earlier_references(self):
        exp, extra = ptg_exp(1, 2)
        shared = worksheet([
            row_header(0), cell(1, "number", 2, rgce=exp, rgcb=extra),
            row_header(1), cell(1, "number", 4, rgce=exp, rgcb=extra), cell(2, "number", 6, rgce=exp, rgcb=extra),
            formula_range((0, 1, 1, 2), ptg_ref(0, -1, shared=True) + ptg_int(2) + b"\x05"),
        ])
        self.assertEqual(xlsb.render_text(package([("共有", shared)])), 'Sheet: "共有"\nB1\tformula: "=A1*2"\nB2\tformula: "=A2*2"\nC2\tformula: "=B2*2"\n')

    def test_render_streams_parts_and_keeps_cells_in_coordinate_order(self):
        rows = [row_header(row) + cell(2, "number", row + 1) + cell(0, "number", row) for row in range(1000)]
        data = package([("逐次", worksheet(rows))])
        archive_read = zipfile.ZipFile.read

        def read_metadata(archive, name, *args, **kwargs):
            self.assertFalse(str(name).startswith("xl/worksheets/"))
            return archive_read(archive, name, *args, **kwargs)

        with mock.patch.object(zipfile.ZipFile, "read", read_metadata), mock.patch.object(xlsb.BinaryWorkbook, "_read_sheet", side_effect=AssertionError("whole-sheet model")), mock.patch.object(xlsb, "Cell", side_effect=AssertionError("cell dataclass")):
            actual = xlsb.render_text(data)
        expected = ['Sheet: "逐次"']
        for row in range(1000):
            expected.extend((f"A{row + 1}\t{row}", f"C{row + 1}\t{row + 1}"))
        self.assertEqual(actual, "\n".join(expected) + "\n")

    def test_broken_formula_at_data_table_anchor_is_not_hidden(self):
        table = record(428, struct.pack("<IIIIIIII B", 0, 1, 0, 0, 0, 1, 0, 0, 0))
        broken = record(9, struct.pack("<IIdH", 0, 0, 3, 0) + struct.pack("<I", 100) + b"\x1e")
        data = package([("破損", worksheet([row_header(0), broken, table]))])
        with self.assertRaises(xlsb.WorkbookError):
            xlsb.render_text(data)

    def test_unknown_and_broken_formulas_retain_cached_value_and_hash(self):
        for tokens in (b"\xff", b"\xfe", b"", b"\x03"):
            with self.subTest(tokens=tokens):
                data = package([("未解読", worksheet([row_header(0), cell(0, "number", 7, rgce=tokens)]))])
                digest = hashlib.sha256(tokens).hexdigest()[:8]
                self.assertEqual(xlsb.render_text(data), 'Sheet: "未解読"\nA1\tformula-unparsed: "#' + digest + '" 7\n')
        data = package([("外部", worksheet([row_header(0), cell(0, "error", 23, rgce=b"\x3a" + struct.pack("<HIH", 0, 0, 0))]))], workbook_records=record(355, wide("rIdExternal")) + record(362, struct.pack("<IIii", 1, 0, 0, 0)))
        self.assertIn('formula-unparsed:', xlsb.render_text(data))
        self.assertIn('error: "#REF!"', xlsb.render_text(data))

    def test_names_and_internal_addin_functions_use_only_local_links(self):
        name = record(39, struct.pack("<IBI", 0, 0, 0xFFFFFFFF) + wide("税率") + parsed_formula(ptg_int(1)) + wide(None))
        externals = record(667) + record(361, wide("_xlfn.IFERROR")) + record(362, struct.pack("<IIii", 1, 0, -2, -2))
        udf = b"\x59" + struct.pack("<HI", 0, 1) + ptg_int(1) + ptg_int(2) + b"\x42\x03\xff\x00"
        data = package([("名前", worksheet([row_header(0), cell(0, "number", 1, rgce=b"\x43\x01\0\0\0"), cell(1, "number", 1, rgce=udf)]))], workbook_records=name + externals)
        self.assertEqual(xlsb.render_text(data), 'Sheet: "名前"\nA1\tformula: "=税率"\nB1\tformula: "=IFERROR(1,2)"\n')

    def test_sheet_scoped_name_is_qualified_from_another_sheet(self):
        name = record(39, struct.pack("<IBI", 0, 0, 0) + wide("税率") + parsed_formula(ptg_int(1)) + wide(None))
        formula = worksheet([row_header(0), cell(0, "number", 1, rgce=b"\x43\x01\0\0\0")])
        data = package([("売上 表", formula), ("備考", formula)], workbook_records=name)
        self.assertEqual(xlsb.render_text(data), 'Sheet: "売上 表"\nA1\tformula: "=税率"\nSheet: "備考"\nA1\tformula: "=\'売上 表\'!税率"\n')

    def test_unknown_records_are_kept_and_broken_payloads_fail(self):
        unknown = record(3072, b"unknown data")
        data = package([("不明", worksheet([unknown, row_header(0), cell(0, "string", "値")]))])
        self.assertIn('A1\t"値"', xlsb.render_text(data))
        self.assertTrue(any(rec.type == 3072 for rec in xlsb.read_workbook(data).sheets[0].records))
        for part in (b"\0\x7ftruncated", worksheet([cell(0, "number", 1)]), worksheet([row_header(1048576)]), worksheet([row_header(0), cell(16384, "number", 1)]), worksheet([row_header(0), cell(0, "shared-string", 999)]), worksheet([row_header(0), cell(0, "number", 1), cell(0, "number", 2)]), worksheet([row_header(0), record(9, struct.pack("<II", 0, 0) + bytes(10) + struct.pack("<I", 20) + b"\x1e")])):
            with self.subTest(part=part[:40]), self.assertRaises(xlsb.WorkbookError):
                xlsb.render_text(replace_part(data, "xl/worksheets/sheet1.bin", part))

    def test_kind_detection_uses_relationship_and_main_content_type(self):
        data = package([("空", worksheet([]))])
        self.assertTrue(xlsb.is_xlsb(data))
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            types = archive.read("[Content_Types].xml")
        override = types.replace(b"</Types>", ('<Override PartName="/xl/workbook.bin" ContentType="' + xlsb.MAIN_TYPE + '"/></Types>').encode())
        self.assertTrue(xlsb.is_xlsb(replace_part(data, "[Content_Types].xml", override)))
        self.assertFalse(xlsb.is_xlsb(replace_part(data, "[Content_Types].xml", types.replace(xlsb.MAIN_TYPE.encode(), b"other"))))
        self.assertFalse(xlsb.is_xlsb(b"PKbroken"))
        self.assertFalse(xlsb.is_xlsb((FIXTURES / "sales.xlsx").read_bytes()))

    def test_large_placeholder_lists_binary_sheet_titles_and_size(self):
        data = (FIXTURES / "sales.xlsb").read_bytes()
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            expected = sum(info.file_size for info in archive.infolist() if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".bin") or info.filename == "xl/sharedStrings.bin")
        self.assertEqual(textconv.input_size(data), expected)
        self.assertEqual(textconv.sheet_titles(data), ["売上", "備考"])
        self.assertTrue(textconv.placeholder("a" * 64, data).endswith('Sheet: "売上"\nSheet: "備考"\n'))
        broken_sheet = replace_part(data, "xl/worksheets/sheet1.bin", b"broken")
        self.assertEqual(xlsb.sheet_titles(broken_sheet), ["売上", "備考"])


class BinaryCliTests(textconv_tests.ConverterTestCase):
    def test_extensionless_binary_blob_and_corruption_messages(self):
        path = self.directory / "git-blob"
        path.write_bytes((FIXTURES / "sales.xlsb").read_bytes())
        self.assertEqual(self.convert(path).stdout, textconv.convert((FIXTURES / "sales.xlsx").read_bytes()))
        path.write_bytes(replace_part(path.read_bytes(), "xl/worksheets/sheet1.bin", b"\0\x7fbroken"))
        result = self.convert(path, expected_returncode=1)
        self.assertIn("xlsb", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        path.write_bytes(textconv.OLE_SIGNATURE + bytes(504))
        result = self.convert(path, expected_returncode=1)
        self.assertIn("xlsb", result.stderr)
        self.assertIn("パスワード", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_large_binary_book_returns_a_named_placeholder(self):
        path = self.directory / "大きい.xlsb"
        rows = [row_header(row) + cell(0, "number", row) for row in range(2000)]
        path.write_bytes(package([("大きい表", worksheet(rows))]))
        cache = self.directory / "cache"
        cache.mkdir()
        (cache / "worker.lock").touch()
        env = textconv_tests.os.environ.copy()
        env["EXCEL_TEXTCONV_TIMEOUT"] = "0.001"
        result = self.convert(path, options=("--cache-dir", str(cache)), env=env)
        self.assertTrue(result.stdout.startswith("# Excel を裏で変換中"))
        self.assertTrue(result.stdout.endswith('Sheet: "大きい表"\n'))


@unittest.skipUnless(shutil.which("git"), "Git is required for textconv integration tests")
class BinaryGitTextconvTests(unittest.TestCase):
    git = textconv_tests.GitTextconvTests.git
    commit_book = textconv_tests.GitTextconvTests.commit_book

    def setUp(self):
        textconv_tests.GitTextconvTests.setUp(self)
        converter_dir = self.directory / "変換 ツール"
        modules = converter_dir / "excel-textconv"
        modules.mkdir()
        for name in ("excel_formula.py", "excel_xlsb.py"):
            shutil.copyfile(ROOT / "scripts" / name, modules / name)
        (self.repo / ".gitattributes").write_text("*.xlsb diff=excel\n", encoding="utf-8")
        self.git("add", "--", ".gitattributes")
        self.git("commit", "--quiet", "-m", "xlsb の差分を設定")

    def test_staged_and_unstaged_binary_workbook_blobs(self):
        path = self.repo / "売上 表.xlsb"
        path.write_bytes((FIXTURES / "sales.xlsb").read_bytes())
        self.commit_book(path)
        path.write_bytes((FIXTURES / "sales-edited.xlsb").read_bytes())
        unstaged = self.git("diff", "--textconv", "--", path.name)
        self.assertIn('-B2\t10', unstaged)
        self.assertIn('+B2\t12', unstaged)
        self.assertIn('-D4\tformula: "=B4*C4"', unstaged)
        self.assertIn('+D4\tformula: "=B4*C4*2"', unstaged)
        self.assertNotIn("Binary files", unstaged)
        self.git("add", "--", path.name)
        staged = self.git("diff", "--cached", "--textconv", "--", path.name)
        self.assertEqual(staged, unstaged)
        path.write_bytes((FIXTURES / "sales.xlsb").read_bytes())
        reversed_diff = self.git("diff", "--textconv", "--", path.name)
        self.assertIn('-B2\t12', reversed_diff)
        self.assertIn('+B2\t10', reversed_diff)
        self.assertEqual(self.git("diff", "--cached", "--textconv", "--", path.name), staged)


if __name__ == "__main__":
    unittest.main()
