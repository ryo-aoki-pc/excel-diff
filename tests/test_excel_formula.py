"""Tests for lossless display of ordinary BIFF8 and BIFF12 formulas."""

from pathlib import Path
import struct
import sys
import unittest

import xlrd.formula


ROOT = Path(__file__).resolve().parents[1]
if ROOT.name == ".stage1-draft":
    ROOT = ROOT.parent
sys.path.insert(0, str(ROOT / "scripts"))
import excel_formula as formula


def integer(value):
    return b"\x1e" + struct.pack("<H", value)


def string(value, version=12, compressed=False):
    if version == 12:
        return b"\x17" + struct.pack("<H", len(value)) + value.encode("utf-16le")
    return b"\x17" + bytes((len(value), not compressed)) + value.encode("latin1" if compressed else "utf-16le")


def ref(row, col, version=12, token=0x24):
    return bytes((token,)) + struct.pack("<IH" if version == 12 else "<HH", row, col)


def area(first_row, last_row, first_col, last_col, version=12, token=0x25):
    return bytes((token,)) + struct.pack("<IIHH" if version == 12 else "<HHHH", first_row, last_row, first_col, last_col)


def attr(flags, value=0):
    return b"\x19" + struct.pack("<BH", flags, value)


def func(index, count=None):
    return (b"\x21" if count is None else b"\x22" + bytes((count,))) + struct.pack("<H", index)


class FormulaTests(unittest.TestCase):
    def assert_both(self, tokens, expected, **kwargs):
        for version in (8, 12):
            with self.subTest(version=version):
                data = tokens(version) if callable(tokens) else tokens
                self.assertEqual(formula.decode_formula(data, version=version, **kwargs), expected)

    def test_constants_and_string_escaping(self):
        self.assert_both(integer(65535), "65535")
        self.assert_both(b"\x1f" + struct.pack("<d", 1.25), "1.25")
        self.assert_both(lambda v: string('a"b\n日本語', v), '"a""b\n日本語"')
        self.assertEqual(formula.decode_formula(string("abc", 8, True), version=8), '"abc"')
        self.assert_both(b"\x1d\x01", "TRUE")
        self.assert_both(b"\x1d\x00", "FALSE")
        self.assert_both(b"\x1c\x2a", "#N/A")
        self.assert_both(b"\x1c\x07", "#DIV/0!")

    def test_numeric_spelling(self):
        cases = (
            (1e20, "100000000000000000000"),
            (1e21, "1E+21"),
            # Excel already reduces an entered 123456789012345678 to this stored value.
            (123456789012345000, "123456789012345000"),
            (123456789012345678, "123456789012346000"),
            (1.23456789012345e21, "1.23456789012345E+21"),
            (1e-5, "0.00001"),
            (1e-10, "0.0000000001"),
            (1e-20, "1E-20"),
            (1.5e-19, "1.5E-19"),
            (1.5e-18, "0.0000000000000000015"),
            (1.23456789012345e-6, "1.23456789012345E-06"),
            (1.23456789012345e-5, "0.0000123456789012345"),
            (0.30000000000000004, "0.3"),
            (-1e19, "-10000000000000000000"),
            (-1e20, "-1E+20"),
            (-0.0, "0"),
        )
        for version in (8, 12):
            token = b"\x60" + b"\0" * (14 if version == 12 else 7)
            dimensions = struct.pack("<II", 1, 1) if version == 12 else struct.pack("<BH", 0, 0)
            for value, expected in cases:
                with self.subTest(version=version, value=value):
                    payload = struct.pack("<d", value)
                    self.assertEqual(formula.decode_formula(b"\x1f" + payload, version=version), expected)
                    extra = dimensions + bytes((0 if version == 12 else 1,)) + payload
                    self.assertEqual(formula.decode_formula(token, extra, version=version), "{" + expected + "}")

    def test_operators_and_precedence(self):
        a, b, c = integer(1), integer(2), integer(3)
        self.assert_both(a + b + c + b"\x05\x03", "1+2*3")
        self.assert_both(a + b + b"\x03" + c + b"\x05", "(1+2)*3")
        self.assert_both(a + b + c + b"\x04\x04", "1-(2-3)")
        self.assert_both(a + b + c + b"\x07\x07", "1^(2^3)")
        self.assert_both(a + b + b"\x07" + c + b"\x07", "1^2^3")
        for token, text in ((3, "+"), (4, "-"), (5, "*"), (6, "/"), (7, "^"),
                            (8, "&"), (9, "<"), (10, "<="), (11, "="),
                            (12, ">="), (13, ">"), (14, "<>"), (15, " "),
                            (16, ","), (17, ":")):
            self.assert_both(a + b + bytes((token,)), "1" + text + "2")

    def test_unary_percent_and_explicit_parentheses(self):
        self.assert_both(integer(1) + b"\x12", "+1")
        self.assert_both(integer(1) + b"\x13\x14", "-1%")
        self.assert_both(integer(1) + integer(2) + b"\x03\x13", "-(1+2)")
        self.assert_both(integer(1) + b"\x15", "(1)")
        self.assert_both(integer(1) + integer(2) + b"\x03\x15" + integer(3) + b"\x05", "(1+2)*3")

    def test_relative_and_absolute_references(self):
        for version in (8, 12):
            with self.subTest(version=version):
                self.assertEqual(formula.decode_formula(ref(1, 0xC001, version), version=version), "B2")
                self.assertEqual(formula.decode_formula(ref(1, 1, version), version=version), "$B$2")
                self.assertEqual(formula.decode_formula(ref(1, 0x4001, version), version=version), "B$2")
                self.assertEqual(formula.decode_formula(ref(1, 0x8001, version), version=version), "$B2")
                self.assertEqual(formula.decode_formula(area(1, 3, 0xC001, 0xC002, version), version=version), "B2:C4")

    def test_shared_formula_signed_offsets_and_wrapping(self):
        for version in (8, 12):
            max_rows, max_cols = (1048576, 16384) if version == 12 else (65536, 256)
            with self.subTest(version=version):
                tokens = ref(max_rows - 1, 0xC000 | (max_cols - 2), version, 0x2C)
                self.assertEqual(formula.decode_formula(tokens, version=version, row=3, col=4), "C3")
                self.assertEqual(formula.decode_formula(tokens, version=version), formula._column(max_cols - 2) + str(max_rows))
                tokens = area(max_rows - 1, 0, 0xC000 | (max_cols - 1), 0xC001, version, 0x2D)
                self.assertEqual(formula.decode_formula(tokens, version=version, row=2, col=2), "B2:D3")

    def test_whole_rows_and_columns(self):
        for version in (8, 12):
            max_row, max_col = (1048575, 16383) if version == 12 else (65535, 255)
            with self.subTest(version=version):
                self.assertEqual(formula.decode_formula(area(0, max_row, 0, 1, version), version=version), "$A:$B")
                self.assertEqual(formula.decode_formula(area(0, max_row, 0x4000, 0x4000, version), version=version), "A:A")
                self.assertEqual(formula.decode_formula(area(0, 1, 0, max_col, version), version=version), "$1:$2")
                self.assertEqual(formula.decode_formula(area(0, max_row, 0, max_col, version), version=version), "$1:$" + str(max_row + 1))
                self.assertEqual(formula.decode_formula(area(0, max_row, 0x4000, 0x4000 | max_col, version), version=version), "$1:$" + str(max_row + 1))
                self.assertEqual(formula.decode_formula(area(0, max_row, 0xC000, 0xC000 | max_col, version), version=version), "1:" + str(max_row + 1))
                self.assertEqual(formula.decode_formula(area(0, max_row, 0, 0x8000 | max_col, version), version=version), "$1:" + str(max_row + 1))

    def test_three_dimensional_reference_and_sheet_quoting(self):
        for version in (8, 12):
            for name, expected in (
                    ("備考", "備考"), ("Sheet1", "Sheet1"), ("A1", "'A1'"),
                    ("RC", "'RC'"), ("R1C", "'R1C'"), ("RC1", "'RC1'"),
                    ("C", "'C'"), ("R", "'R'"), ("R1", "'R1'"), ("C1", "'C1'"),
                    ("R2C3", "'R2C3'"), ("r2c3", "'r2c3'"),
                    ("TRUE", "'TRUE'"), ("false", "'false'"), ("XFD1", "'XFD1'"),
                    ("A1048576", "'A1048576'"), ("A1048577", "A1048577"),
                    ("1abc", "'1abc'"), ("名前 空白", "'名前 空白'"),
                    ("シート'名", "'シート''名'"), ("A1B", "A1B"),
                    ("a.b", "a.b"), ("XFE1", "XFE1"), ("Main", "Main"),
                    ("_名前", "_名前"), (".名前", ".名前"),
                    ("名前・中点", "'名前・中点'"), ("名前２", "名前２"),
                    ("１abc", "'１abc'"),
                    ("a b", "'a b'"), ("O'Brien", "'O''Brien'"), ("123", "'123'")):
                with self.subTest(version=version, name=name):
                    tokens = b"\x3a\x00\x00" + ref(1, 0xC001, version)[1:]
                    self.assertEqual(formula.decode_formula(tokens, version=version, sheets=[name]), expected + "!B2")
            tokens = b"\x3b\x00\x00" + area(0, 2, 0xC001, 0xC001, version)[1:]
            self.assertEqual(formula.decode_formula(tokens, version=version, sheets=[("先頭", "末尾")]), "先頭:末尾!B1:B3")
            self.assertEqual(formula.decode_formula(tokens, version=version, sheets=[("a b", "末尾")]), "'a b:末尾'!B1:B3")

    def test_names_and_addin_functions(self):
        name = b"\x23" + struct.pack("<I", 1)
        self.assert_both(name, "税率", names={1: "税率"})
        self.assert_both(name + integer(1) + integer(0) + b"\x06" + integer(2) + func(255, 3),
                         "IFERROR(1/0,2)", names={1: "_xlfn.IFERROR"})
        external = b"\x39" + struct.pack("<HI", 3, 1)
        self.assert_both(external + integer(1) + func(255, 2), "_xlfn.XLOOKUP(1)",
                         external_names={(3, 1): "_xlfn.XLOOKUP"})
        self.assert_both(external + integer(1) + integer(2) + func(255, 3), "IFERROR(1,2)",
                         external_names={(3, 1): "_xlfn.IFERROR"})

    def test_fixed_variable_and_empty_arguments(self):
        self.assert_both(integer(1) + integer(2) + func(27), "ROUND(1,2)")
        self.assert_both(integer(1) + integer(2) + integer(3) + func(4, 3), "SUM(1,2,3)")
        self.assert_both(func(10), "NA()")
        self.assert_both(integer(1) + b"\x16" + integer(3) + func(1, 3), "IF(1,,3)")
        self.assert_both(integer(1) + integer(2) + b"\x10" + func(4, 1), "SUM((1,2))")

    def test_array_constants(self):
        for version in (8, 12):
            token = b"\x60" + b"\x00" * (14 if version == 12 else 7)
            dimensions = struct.pack("<II", 2, 2) if version == 12 else struct.pack("<BH", 1, 1)
            data = dimensions + b"".join(bytes((0 if version == 12 else 1,)) + struct.pack("<d", n) for n in (1, 2, 3, 4))
            self.assertEqual(formula.decode_formula(token + func(7, 1), data, version=version), "MAX({1,2;3,4})")
            if version == 12:
                data = struct.pack("<II", 1, 3) + b"\x01\x03\x00" + 'a"b'.encode("utf-16le") + b"\x02\x01\x04\x2a\x00\x00\x00"
            else:
                data = struct.pack("<BH", 2, 0) + b"\x02\x03\x00\x01" + 'a"b'.encode("utf-16le") + b"\x04\x01" + b"\0" * 7 + b"\x10\x2a" + b"\0" * 7
            self.assertEqual(formula.decode_formula(token, data, version=version), '{"a""b",TRUE,#N/A}')

    def test_attributes_and_whitespace(self):
        self.assert_both(integer(1) + attr(0x10), "SUM(1)")
        for flags in (0, 1, 2, 8, 0x20, 0x21, 0x80):
            self.assert_both(integer(1) + attr(flags), "1")
        self.assert_both(integer(1) + attr(4, 1) + b"\x00\x00\x00\x00", "1")
        self.assert_both(attr(0x40, 2 << 8) + integer(1), "  1")
        self.assert_both(integer(1) + integer(2) + attr(0x40, 1 << 8) + b"\x03", "1 +2")
        self.assert_both(integer(1) + attr(0x40, (2 << 8) | 2) + attr(0x40, (1 << 8) | 4) + b"\x15", "  (1 )")
        self.assert_both(attr(0x41, (1 << 8) | 1) + integer(1), "\r1")

    def test_function_table_matches_xlrd(self):
        for number, value in xlrd.formula.func_defs.items():
            with self.subTest(number=number):
                self.assertIn(number, formula.FTAB)
                self.assertEqual(formula.FTAB[number][0], value[0])
        self.assertEqual(formula.FTAB[480], ("IFERROR", 2, 2))
        self.assertEqual(formula.FTAB[484][0], "AVERAGEIFS")
        self.assertGreater(len(formula.FTAB), 470)

    def test_token_classes(self):
        for version in (8, 12):
            for token in (0x24, 0x44, 0x64):
                self.assertEqual(formula.decode_formula(ref(0, 0, version, token), version=version), "$A$1")

    def test_memory_tokens(self):
        self.assert_both(b"\x29\x03\x00" + integer(1), "1")
        for version in (8, 12):
            extra = struct.pack("<I" if version == 12 else "<H", 0)
            self.assertEqual(formula.decode_formula(b"\x26" + b"\0" * 6 + integer(1), extra, version=version), "1")

    def test_reference_errors(self):
        for version in (8, 12):
            for token, count in ((0x2A, 1), (0x2B, 2)):
                self.assertEqual(formula.decode_formula(bytes((token,)) + b"\0" * ((6 if version == 12 else 4) * count), version=version), "#REF!")

    def test_malformed_and_unsupported_tokens(self):
        for tokens in (b"", b"\x03", integer(1) + integer(2), b"\x18\x19", b"\xff",
                       b"\x1e", b"\x1d\x02", b"\x1c\xff", b"\x01" + b"\0" * 4,
                       integer(1) + func(27), func(4), func(999), b"\x17\xff\x00",
                       integer(1) + attr(0x40, 7), b"\x60" + b"\0" * 14,
                       b"\x1f" + struct.pack("<d", float("nan"))):
            with self.subTest(tokens=tokens):
                with self.assertRaises(formula.FormulaError):
                    formula.decode_formula(tokens)
        with self.assertRaises(formula.FormulaError):
            formula.decode_formula(integer(1), b"extra")
        with self.assertRaises(formula.FormulaError):
            formula.decode_formula(b"\x3a\0\0" + ref(0, 0)[1:], sheets={0: None})
        with self.assertRaises(formula.FormulaError):
            formula.decode_formula(b"\x39" + struct.pack("<HI", 0, 1), external_names={(0, 1): None})
        with self.assertRaises(formula.FormulaError):
            formula.decode_formula(integer(1), version=5)
        with self.assertRaises(formula.FormulaError):
            formula.decode_formula(ref(0, 256, 8), version=8)


if __name__ == "__main__":
    unittest.main()
