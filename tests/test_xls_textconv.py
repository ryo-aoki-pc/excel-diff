"""Check BIFF8 formulas and Excel's equivalent xlsx fixture."""
import hashlib
import importlib.util
import io
from pathlib import Path
import struct
import sys
from types import SimpleNamespace
import unittest
import xlrd

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
sys.path.insert(0, str(SCRIPTS))
import excel_xls
spec = importlib.util.spec_from_file_location("xls_textconv", SCRIPTS / "excel-textconv.py")
converter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(converter)


def record(kind, data=b""):
    return struct.pack("<HH", kind, len(data)) + data


def formula_cell(row, col, rgce, rgcb=b""):
    data = struct.pack("<HHHdHIH", row, col, 0, 42, 0, 0, len(rgce)) + rgce + rgcb
    return record(excel_xls.FORMULA, data)


def fake_book(sheet_records, globals_records=b""):
    globals_stream = globals_records + record(excel_xls.EOF)
    sheet = SimpleNamespace(cell=lambda row, col: SimpleNamespace(ctype=xlrd.XL_CELL_NUMBER, value=42))
    return SimpleNamespace(biff_version=80, nsheets=1, base=0, datemode=0,
        mem=globals_stream + sheet_records + record(excel_xls.EOF),
        _sh_abs_posn=[len(globals_stream)], sheet_by_index=lambda index: sheet)


class XlsTextconvTests(unittest.TestCase):
    def test_sales_formulas_match_xlsx(self):
        expected = io.StringIO()
        converter.render_ooxml(io.BytesIO((FIXTURES / "sales.xlsx").read_bytes()), expected)
        actual = excel_xls.render_text((FIXTURES / "sales.xls").read_bytes())
        def formula_lines(text):
            return [line for line in text.splitlines() if "\tformula:" in line or "\tarray-formula:" in line]
        self.assertEqual(formula_lines(actual), formula_lines(expected.getvalue()))
        expected_lines = [line for line in expected.getvalue().splitlines()
                          if not line.startswith("AMK1\t")]
        self.assertEqual(actual, "\n".join(expected_lines) + "\n")
        self.assertIn('D2\tformula: "=B2*C2"\n', actual)
        self.assertIn('J2\tarray-formula: {"text": "=B2:B3*C2:C3", "ref": "J2:J3"}\n', actual)
        self.assertIn('J3\t400\n', actual)
        self.assertNotIn("formula-unparsed:", actual)
        self.assertNotIn("AMK1\t", actual)

    def test_public_converter_uses_xls_formulas(self):
        data = (FIXTURES / "sales.xls").read_bytes()
        self.assertEqual(converter.convert(data), excel_xls.render_text(data))

    def test_shared_formula_resolves_each_cell(self):
        anchor = b"\x01" + struct.pack("<HH", 1, 1)
        # Row-relative zero and column-relative +1 refer to C2/C3.
        expression = b"\x4c" + struct.pack("<HH", 0, 0xc001) + b"\x1e\x02\x00\x05"
        shared = struct.pack("<HHBBBBH", 1, 2, 1, 1, 0, 2, len(expression)) + expression
        book = fake_book(formula_cell(1, 1, anchor) + record(excel_xls.SHRFMLA, shared) + formula_cell(2, 1, anchor))
        self.assertEqual(excel_xls.formula_cells(book), [{(1, 1): 'formula: "=C2*2"', (2, 1): 'formula: "=C3*2"'}])

    def test_shared_anchor_need_not_be_range_top_left(self):
        anchor = b"\x01" + struct.pack("<HH", 2, 1)
        expression = b"\x4c" + struct.pack("<HH", 0, 0xc001) + b"\x1e\x02\x00\x05"
        # B2 is in the shared range but does not use the shared formula.
        shared = struct.pack("<HHBBBBH", 1, 3, 1, 1, 0, 2, len(expression)) + expression
        book = fake_book(formula_cell(2, 1, anchor) + record(excel_xls.SHRFMLA, shared)
                         + formula_cell(3, 1, anchor))
        self.assertEqual(excel_xls.formula_cells(book), [{
            (2, 1): 'formula: "=C3*2"', (3, 1): 'formula: "=C4*2"'}])

    def test_continued_array_strings_keep_their_original_bytes(self):
        expression = b"\x60" + bytes(7)
        # SerStr uses XLUnicodeString, rather than SST's rich string form.
        # Its continuation is raw bytes; do not consume a 0/1 as a new flag.
        extra = b"\x01\x00\x00\x02\x01\x00\x00A\x02\x01\x00\x01\x01\x01"
        payload = formula_cell(0, 0, expression, extra)[4:]
        for split in range(len(payload) - len(extra), len(payload)):
            with self.subTest(split=split):
                book = fake_book(record(excel_xls.FORMULA, payload[:split])
                                 + record(excel_xls.CONTINUE, payload[split:]))
                self.assertEqual(excel_xls.formula_cells(book), [{
                    (0, 0): 'formula: "={\\"A\\",\\"ā\\"}"'}])

    def test_unparsed_formula_hashes_tokens_and_extra_data(self):
        for extra in (b"", b"changed-array-data"):
            with self.subTest(extra=extra):
                expression = b"\x18"
                digest = hashlib.sha256(expression + extra).hexdigest()[:8]
                result = excel_xls.formula_cells(fake_book(formula_cell(0, 0, expression, extra)))
                self.assertEqual(result, [{(0, 0): 'formula-unparsed: "#' + digest + '" 42'}])

    def test_unknown_shared_formula_hashes_definition(self):
        anchor = b"\x01" + struct.pack("<HH", 1, 1)
        expression = b"\x18"
        shared = struct.pack("<HHBBBBH", 1, 2, 1, 1, 0, 2, 1) + expression
        book = fake_book(formula_cell(1, 1, anchor) + record(excel_xls.SHRFMLA, shared) + formula_cell(2, 1, anchor))
        text = 'formula-unparsed: "#' + hashlib.sha256(expression).hexdigest()[:8] + '" 42'
        self.assertEqual(excel_xls.formula_cells(book), [{(1, 1): text, (2, 1): text}])

    def test_data_table_matches_openpyxl_attributes(self):
        anchor = b"\x02" + struct.pack("<HH", 1, 1)
        cases = (
            (0x00, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "0", "dtr": "0", "r1": "A1"}'),
            (0x04, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "0", "dtr": "1", "r1": "A1"}'),
            (0x08, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "1", "dtr": "1", "r1": "A1", "r2": "D1"}'),
            (0x0d, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "1", "dtr": "1", "r1": "A1", "r2": "D1", "ca": "1"}'),
            (0x39, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "1", "dtr": "1", "del1": "1", "del2": "1", "ca": "1"}'),
            (0x34, '{"t": "dataTable", "ref": "B2:C3", "dt2D": "0", "dtr": "1", "del1": "1"}'),
        )
        for flags, attributes in cases:
            with self.subTest(flags=flags):
                table = struct.pack("<HHBB5H", 1, 2, 1, 2, flags, 0, 0, 0, 3)
                book = fake_book(formula_cell(1, 1, anchor)
                                 + record(excel_xls.TABLE, table) + formula_cell(2, 1, anchor))
                self.assertEqual(excel_xls.formula_cells(book), [{
                    (1, 1): "data-table-formula: " + attributes}])

    def test_continue_preserves_formula_tokens(self):
        payload = formula_cell(0, 0, b"\x1e\x2a\x00")[4:]
        book = fake_book(record(excel_xls.FORMULA, payload[:-1]) + record(excel_xls.CONTINUE, payload[-1:]))
        self.assertEqual(excel_xls.formula_cells(book), [{(0, 0): 'formula: "=42"'}])

    def test_continue_name_can_change_unicode_compression(self):
        header = struct.pack("<HBBHHH4B", 0, 0, 2, 0, 0, 0, 0, 0, 0, 0)
        names = record(excel_xls.NAME, header + b"\x00A")
        names += record(excel_xls.CONTINUE, b"\x01" + "名".encode("utf-16le"))
        self.assertEqual(excel_xls.formula_context(fake_book(b"", names))["names"], {1: "A名"})

    def test_continue_unicode_surrogate_pair_is_not_split(self):
        text = "A😀B"
        data = text.encode("utf-16le")
        reader = excel_xls.Reader((b"\x01" + data[:4], b"\x01" + data[4:]))
        self.assertEqual(reader.unicode(4), text)

    def test_local_names_are_qualified_only_on_other_worksheets(self):
        def sheet(name, kind):
            encoded = name.encode("utf-16le")
            return record(excel_xls.BOUNDSHEET,
                          struct.pack("<IBBB", 0, 0, kind, len(name)) + b"\x01" + encoded)
        # itab counts all BOUNDSHEET records, including the chart xlrd ignores.
        globals_records = sheet("図表", 2) + sheet("売上 表", 0) + sheet("備考", 0)
        header = struct.pack("<HBBHHH4B", 0, 0, 2, 0, 0, 2, 0, 0, 0, 0)
        globals_records += record(excel_xls.NAME, header + b"\x01" + "税率".encode("utf-16le"))
        book = fake_book(b"", globals_records)
        self.assertEqual(excel_xls.formula_context(book)["names"], {1: "税率"})
        self.assertEqual(excel_xls.formula_context(book, 0)["names"], {1: "税率"})
        context = excel_xls.formula_context(book, 1)
        self.assertEqual(context["names"], {1: "'売上 表'!税率"})
        self.assertEqual(excel_xls.decode_formula(b"\x43\x01\x00\x00\x00", **context),
                         "'売上 表'!税率")

    def test_external_names_are_not_resolved(self):
        globals_records = record(excel_xls.SUPBOOK, b"\x01\x00\x04\x00file")
        # Do not parse a DDE/external record as an add-in name.
        globals_records += record(excel_xls.EXTERNNAME, b"\x00")
        globals_records += record(excel_xls.EXTERNSHEET, struct.pack("<4H", 1, 0, 0, 0))
        context = excel_xls.formula_context(fake_book(b"", globals_records))
        self.assertEqual(context["sheets"], {})
        self.assertEqual(context["external_names"], {})

    def test_old_biff_does_not_parse_formulas(self):
        self.assertEqual(excel_xls.formula_cells(SimpleNamespace(biff_version=70, nsheets=2)), [{}, {}])

    def test_broken_record_lengths_fail(self):
        for data in (b"\x06", struct.pack("<HH", 6, 10) + b"short"):
            with self.subTest(data=data), self.assertRaisesRegex(ValueError, "xls"):
                list(excel_xls.records(data, 0))


if __name__ == "__main__":
    unittest.main()
