"""Compare Excel-made XLSB formatting with its OOXML counterpart."""

import json
from pathlib import Path
import re
import sys
import unittest


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
import excel_stage_ooxml as OOXML
import excel_stage_xlsb as XLSB


FIXTURES = Path(__file__).resolve().parent / "fixtures"
TINT = re.compile(r"(?<=濃淡)[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")


def books(module, extension):
    return tuple(module.Workbook((FIXTURES / (name + extension)).read_bytes())
                 for name in ("formats-base", "formats"))


def format_rows(base, source):
    comparison = base.compare_formats(source)
    left, right, coordinates = comparison.projection()
    rows = {}
    for sheet in base.sheets:
        baseline = [json.loads(line) for line in left[sheet.id].splitlines() if line]
        changed = [json.loads(line) for line in right[sheet.id].splitlines() if line]
        if len(baseline) != len(changed) or len(baseline) != len(coordinates[sheet.id]):
            raise AssertionError("書式の射影の行数が一致しません")
        for old, new, coordinate in zip(baseline, changed, coordinates[sheet.id]):
            if old[:2] != [coordinate, "format"] or new[:2] != [coordinate, "format"]:
                raise AssertionError("書式の射影の番地が一致しません")
            rows[sheet.name, coordinate] = (
                old, new, comparison.unsupported_reason((sheet.id, coordinate)))
    return rows


class XlsbFormatsOracleTests(unittest.TestCase):
    def assert_description_equal(self, actual, expected):
        self.assertEqual(actual[:2], expected[:2])
        self.assertEqual(TINT.sub("<tint>", actual[2]), TINT.sub("<tint>", expected[2]))
        actual_tints = [float(match[0]) for match in TINT.finditer(actual[2])]
        expected_tints = [float(match[0]) for match in TINT.finditer(expected[2])]
        self.assertEqual(len(actual_tints), len(expected_tints))
        for actual_tint, expected_tint in zip(actual_tints, expected_tints):
            self.assertLess(abs(actual_tint - expected_tint), 1e-9)

    def test_format_projection_and_reasons_match_xlsx(self):
        expected = format_rows(*books(OOXML, ".xlsx"))
        actual = format_rows(*books(XLSB, ".xlsb"))
        self.assertEqual(len(expected), 40)
        self.assertEqual(actual.keys(), expected.keys())
        for key, (old, new, reason) in actual.items():
            with self.subTest(sheet=key[0], coordinate=key[1]):
                expected_old, expected_new, expected_reason = expected[key]
                self.assert_description_equal(old, expected_old)
                self.assert_description_equal(new, expected_new)
                self.assertEqual(reason, expected_reason)

    def test_all_formats_import_with_exact_source_styles_and_preserved_values(self):
        base, source = books(XLSB, ".xlsb")
        comparison = base.compare_formats(source)
        selected = set(comparison.keys())
        self.assertEqual(len(selected), 40)
        result = XLSB.Workbook(base.rebuild(source, set(), selected))
        self.assertEqual(dict(result.cells), dict(base.cells))
        keys = set(base.cells) | set(source.cells) | set(base.cell_styles) | set(source.cell_styles)
        result_styles = result.stylesheet()
        for key in keys:
            with self.subTest(key=key):
                expected = source if key in selected else base
                self.assertEqual(result_styles.exact(result.effective_style(key)),
                                 expected.stylesheet().exact(expected.effective_style(key)))
        sheet = next(sheet for sheet in result.sheets if sheet.name == "Fmt")
        style = result_styles.fields(result.effective_style((sheet.id, "B11")))["style"]
        self.assertEqual(style, ("custom", "試験のスタイル"))
        self.assertEqual(result_styles.style_names()[style], "試験のスタイル")


if __name__ == "__main__":
    unittest.main()
