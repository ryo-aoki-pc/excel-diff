#!/usr/bin/env python3
"""Compare what Microsoft Excel sees with what the binary readers and rebuilds produce.

These tests run only on Windows with Excel installed and not already running;
tests/excel_open.ps1 opens each workbook read-only in a private hidden instance.
"""

import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import unittest

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"
HELPER = ROOT / "tests" / "excel_open.ps1"
sys.path.insert(0, str(ROOT / "scripts"))

import excel_formula  # noqa: E402
import excel_xlsb  # noqa: E402

POWERSHELL = shutil.which("pwsh") or shutil.which("powershell")


def excel_running():
    result = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"], capture_output=True)
    return b"EXCEL.EXE" in result.stdout.upper()


def excel_view(paths, format_cells=()):
    """Return Excel's cells for each path, or skip when Excel cannot be used."""
    if os.name != "nt" or POWERSHELL is None:
        raise unittest.SkipTest("Excel の確かめは Windows だけで行います")
    # The instance of a previous run exits a few seconds after its PowerShell ends.
    deadline = time.monotonic() + 30
    while excel_running() and time.monotonic() < deadline:
        time.sleep(1)
    with tempfile.TemporaryDirectory() as directory:
        listing = Path(directory) / "paths.txt"
        listing.write_text("\n".join(str(path) for path in paths), encoding="utf-8")
        command = [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HELPER),
                   "-ListFile", str(listing), "-FormatCells", ",".join(format_cells)]
        result = subprocess.run(command, capture_output=True, timeout=600)
    output = result.stdout.decode("utf-8", errors="replace").strip()
    if result.returncode == 3:
        raise unittest.SkipTest(output or "Excel を使えません")
    if result.returncode != 0:
        raise AssertionError("excel_open.ps1 failed: " + output + result.stderr.decode("utf-8", errors="replace"))
    entries = json.loads(output)
    return {Path(entry["path"]).name: entry for entry in entries}


def reader_cells(data):
    """Cells as Excel reports them: Value2 plus the English A1 formula."""
    book = excel_xlsb.read_workbook(data)
    result = {}
    for sheet in book.sheets:
        context = book.formula_context(sheet)
        cells = {}
        arrays = {}
        for anchor, array in sheet.arrays.items():
            formula = "=" + excel_formula.decode_formula(
                array.rgce, array.rgcb, version=12, row=anchor[0], col=anchor[1], **context)
            first_row, last_row, first_col, last_col = array.bounds
            for row in range(first_row, last_row + 1):
                for col in range(first_col, last_col + 1):
                    arrays[(row, col)] = formula
        for cell in sheet.cells:
            if cell.kind == "blank":
                continue
            value = {"error": cell.value} if cell.kind == "error" else cell.value
            item = {"value": value}
            if cell.record.type in (8, 9, 10, 11):
                if (cell.row, cell.col) in arrays:
                    item["formula"] = arrays[(cell.row, cell.col)]
                else:
                    item["formula"] = "=" + excel_formula.decode_formula(
                        cell.rgce, cell.rgcb, version=12, row=cell.row, col=cell.col, **context)
            cells[cell.coordinate] = item
        result[sheet.name] = cells
    return result


def excel_cells(entry):
    """Excel's cells without the format details, for comparing with reader_cells."""
    sheets = {}
    for name, cells in entry["sheets"].items():
        sheets[name] = {address: {key: value for key, value in item.items() if key != "format"}
                        for address, item in cells.items() if item.get("value") is not None or "formula" in item}
    return sheets


def expected_cells(target, source, selected, names):
    """Cells of target with the selected coordinates taken from source."""
    expected = {name: dict(cells) for name, cells in target.items()}
    for sheet, coordinate in selected:
        name = names[sheet]
        expected[name].pop(coordinate, None)
        if coordinate in source[name]:
            expected[name][coordinate] = source[name][coordinate]
    return expected


def constants_and_formulas(cells):
    """Excel recalculates formulas on open, so compare formula text, not cached values."""
    return {name: {address: ({"formula": item["formula"]} if "formula" in item else item)
                   for address, item in sheet.items()}
            for name, sheet in cells.items()}


class ExcelOpenTests(unittest.TestCase):
    maxDiff = None

    def test_reader_values_and_formulas_match_excel(self):
        names = ("sales.xlsb", "sales-edited.xlsb", "many-rows.xlsb")
        view = excel_view([FIXTURES / name for name in names])
        for name in names:
            with self.subTest(name=name):
                entry = view[name]
                self.assertTrue(entry["ok"], entry.get("error"))
                self.assertEqual(reader_cells((FIXTURES / name).read_bytes()), excel_cells(entry))

    def test_rebuilt_xlsb_opens_in_excel_with_only_selected_cells_changed(self):
        import excel_stage_xlsb

        target_data = (FIXTURES / "sales.xlsb").read_bytes()
        source_data = (FIXTURES / "sales-edited.xlsb").read_bytes()
        target, source = excel_stage_xlsb.Workbook(target_data), excel_stage_xlsb.Workbook(source_data)
        names = {sheet.id: sheet.name for sheet in target.sheets}
        ids = {sheet.name: sheet.id for sheet in target.sheets}
        changed = sorted(key for key in set(target.cells) | set(source.cells)
                         if target.cells.get(key) != source.cells.get(key)
                         and not target.unsupported_reason(key) and not source.unsupported_reason(key))
        some = {(ids["売上"], coordinate) for coordinate in ("B2", "D4", "A6", "AMK3", "A50", "E5", "F3")}
        some.add((ids["備考"], "B2"))
        self.assertTrue(some <= set(changed))
        comparison = target.compare_formats(source)
        formats = sorted(key for key in comparison.keys() if not comparison.unsupported_reason(key))
        format_cells = ("売上!B2", "売上!A1", "売上!M3", "売上!M2")
        self.assertTrue({(ids["売上"], "B2"), (ids["売上"], "A1"), (ids["売上"], "M3")} <= set(formats))
        cases = {"all-values.xlsb": (set(changed), set()), "some-values.xlsb": (some, set()),
                 "all-formats.xlsb": (set(), set(formats)), "both.xlsb": (set(changed), set(formats))}
        with tempfile.TemporaryDirectory(dir=os.environ.get("EXCEL_OPEN_TEMP")) as directory:
            paths = []
            for name, (selected, styles) in cases.items():
                path = Path(directory) / name
                path.write_bytes(target.rebuild(source, selected, styles))
                paths.append(path)
            view = excel_view(paths + [FIXTURES / "sales.xlsb", FIXTURES / "sales-edited.xlsb"], format_cells)
        before, after = reader_cells(target_data), reader_cells(source_data)
        target_view, source_view = view["sales.xlsb"]["sheets"], view["sales-edited.xlsb"]["sheets"]
        for name, (selected, styles) in cases.items():
            with self.subTest(case=name):
                entry = view[name]
                self.assertTrue(entry["ok"], entry.get("error"))
                self.assertNotIn("修復", entry.get("caption", ""))
                self.assertNotIn("Repaired", entry.get("caption", ""))
                expected = expected_cells(before, after, selected, names)
                self.assertEqual(constants_and_formulas(excel_cells(entry)), constants_and_formulas(expected))
                for item in format_cells:
                    sheet, coordinate = item.split("!")
                    origin = source_view if (ids[sheet], coordinate) in styles else target_view
                    self.assertEqual(entry["sheets"][sheet][coordinate]["format"],
                                     origin[sheet][coordinate]["format"], item)

    def test_imported_formats_look_the_same_in_excel(self):
        import excel_stage_xlsb

        target_data = (FIXTURES / "formats-base.xlsb").read_bytes()
        source_data = (FIXTURES / "formats.xlsb").read_bytes()
        target, source = excel_stage_xlsb.Workbook(target_data), excel_stage_xlsb.Workbook(source_data)
        names = {sheet.id: sheet.name for sheet in target.sheets}
        comparison = target.compare_formats(source)
        keys = sorted(comparison.keys())
        unsupported = {key: comparison.unsupported_reason(key) for key in keys if comparison.unsupported_reason(key)}
        self.assertEqual(unsupported, {})
        self.assertGreaterEqual(len(keys), 35)
        cells = [names[sheet] + "!" + coordinate for sheet, coordinate in keys]
        with tempfile.TemporaryDirectory(dir=os.environ.get("EXCEL_OPEN_TEMP")) as directory:
            path = Path(directory) / "imported.xlsb"
            path.write_bytes(target.rebuild(source, (), keys))
            view = excel_view([path, FIXTURES / "formats.xlsb"], cells)
        entry = view["imported.xlsb"]
        self.assertTrue(entry["ok"], entry.get("error"))
        self.assertNotIn("修復", entry.get("caption", ""))
        for item in cells:
            sheet, coordinate = item.split("!")
            with self.subTest(cell=item):
                got = dict(entry["sheets"][sheet][coordinate]["format"])
                want = dict(view["formats.xlsb"]["sheets"][sheet][coordinate]["format"])
                if want["fontName"] is None:
                    # Excel reports no single font name when the string itself holds
                    # runs in several fonts; such text runs are values, not cell formats.
                    del got["fontName"], want["fontName"]
                self.assertEqual(got, want)


if __name__ == "__main__":
    unittest.main()
