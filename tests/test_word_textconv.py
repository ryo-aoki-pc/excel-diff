"""Exercise Word textconv dispatch, installed layouts, cache, and Git."""

import hashlib
import importlib.util
import io
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock
import zipfile

from test_excel_textconv import isolated_env, save_book
from word_fixtures import make_docx


ROOT = Path(__file__).resolve().parents[1]
CONVERTER = ROOT / "scripts" / "excel-textconv.py"
WORD_MODULE = ROOT / "scripts" / "word_ooxml.py"


def body(text="合成の本文", bold=False):
    props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return "<w:p><w:r>" + props + "<w:t>" + text + "</w:t></w:r></w:p>"


def rewrite_zip(data, updates):
    """Replace selected package parts without changing the other members."""
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                target.writestr(item, updates.get(item.filename, source.read(item.filename)))
    return output.getvalue()


class WordConverterTestCase(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(prefix="word-textconv-", dir=ROOT)
        self.addCleanup(directory.cleanup)
        self.directory = Path(directory.name)
        self.converter = CONVERTER
        self.env = os.environ.copy()
        self.env["PYTHONUTF8"] = "1"
        self.env["PYTHONDONTWRITEBYTECODE"] = "1"

    def convert(self, path, expected_returncode=0, options=(), env=None):
        result = subprocess.run(
            [sys.executable, "-B", str(self.converter), *options, str(path)],
            capture_output=True, encoding="utf-8", env=env or self.env, timeout=120,
        )
        self.assertEqual(result.returncode, expected_returncode, result.stderr)
        return result

    def document(self, name="合成 文書.docx", **options):
        path = self.directory / name
        path.write_bytes(make_docx(body(), **options))
        return path

    def copy_converter(self, module=None, fallback=None):
        directory = self.directory / "変換 ツール"
        directory.mkdir()
        self.converter = directory / "excel-textconv-build.py"
        shutil.copyfile(CONVERTER, self.converter)
        if fallback is not None:
            (directory / "word_ooxml.py").write_bytes(fallback.encode("utf-8"))
        if module is not None:
            build = self.converter.with_suffix("")
            build.mkdir()
            (build / "word_ooxml.py").write_bytes(module.encode("utf-8"))


class WordConverterTests(WordConverterTestCase):
    def test_repository_fallback_extensionless_blob_and_four_content_types(self):
        expected = self.convert(self.document()).stdout
        self.assertTrue(expected.startswith("# 本文\n"))
        self.assertIn('"合成の本文"', expected)
        for main_type in ("docx", "docm", "dotx", "dotm"):
            with self.subTest(main_type=main_type):
                path = self.document("文書." + main_type, main_type=main_type)
                self.assertEqual(self.convert(path).stdout, expected)
        blob = self.document("一時 blob")
        self.assertEqual(self.convert(blob).stdout, expected)
        self.assertEqual(self.convert(blob).stdout, self.convert(blob).stdout)

    def test_main_content_type_override_is_matched_case_insensitively(self):
        path = self.document()
        data = path.read_bytes()
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            types = archive.read("[Content_Types].xml")
        types = types.replace(b'/word/document.xml', b'/WORD/DOCUMENT.XML')
        path.write_bytes(rewrite_zip(data, {"[Content_Types].xml": types}))
        self.assertIn('"合成の本文"', self.convert(path).stdout)

    def test_main_relationship_target_is_resolved_from_the_package_root(self):
        path = self.document(main="word/alternate body.xml")
        data = path.read_bytes()
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            rels = archive.read("_rels/.rels")
        rels = rels.replace(b"word/alternate body.xml", b"/word/alternate%20body.xml")
        path.write_bytes(rewrite_zip(data, {"_rels/.rels": rels}))
        self.assertIn('"合成の本文"', self.convert(path).stdout)

    def test_content_wins_over_the_extension_without_loading_word_for_excel(self):
        self.copy_converter()
        path = self.directory / "Excel の内容.docx"
        save_book(path)
        self.assertIn('Sheet: "売上"', self.convert(path).stdout)
        path = self.document("Word の内容.xlsx")
        result = self.convert(path, expected_returncode=1)
        self.assertIn("Word の差分変換に失敗しました:", result.stderr)
        self.assertIn("word_ooxml.py がありません", result.stderr)

    def test_installed_sibling_module_wins_over_script_directory(self):
        self.copy_converter(
            module='def render_text(data):\n    return "# sibling build\\n"\n',
            fallback='raise RuntimeError("script directory module was imported")\n',
        )
        self.assertEqual(self.convert(self.document()).stdout, "# sibling build\n")

    def test_placeholder_and_convert_insert_the_word_module_path_only_once(self):
        name = "word_path_test_converter"
        spec = importlib.util.spec_from_file_location(name, CONVERTER)
        converter = importlib.util.module_from_spec(spec)
        sys.modules[name] = converter
        self.addCleanup(sys.modules.pop, name, None)
        spec.loader.exec_module(converter)
        module_directory = str(WORD_MODULE.parent.resolve())
        original_path = list(sys.path)
        self.addCleanup(setattr, sys, "path", original_path)
        sys.path[:] = [entry for entry in sys.path if entry != module_directory]
        with mock.patch.dict(sys.modules, {"word_ooxml": mock.Mock(
                placeholder_lines=mock.Mock(return_value=["# 本文"]),
                render_text=mock.Mock(return_value="# 本文\n"))}):
            data = make_docx(body())
            self.assertIn("# 本文", converter.placeholder("synthetic-key", data))
            self.assertEqual(converter.convert(data), "# 本文\n")
            converter.word_module()
        self.assertEqual(sys.path.count(module_directory), 1)
        self.assertEqual(sys.path[0], module_directory)

    def test_missing_word_module_fails_without_traceback_and_excel_still_works(self):
        self.copy_converter()
        result = self.convert(self.document(), expected_returncode=1)
        self.assertEqual(result.stdout, "")
        self.assertIn("Word の差分変換に失敗しました:", result.stderr)
        self.assertIn("セットアップを再実行してください。", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        path = self.directory / "合成 ブック.xlsx"
        save_book(path)
        self.assertIn('A1\t"元の値"', self.convert(path).stdout)

    def test_plain_corrupt_ole_and_missing_word_files_fail_clearly(self):
        cases = (
            ("テキスト.docx", b"plain text", "対応する Word ファイルではありません (docx/docm/dotx/dotm)"),
            ("壊れた.docx", b"PK\x03\x04truncated", None),
            ("暗号化.docm", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"broken", "暗号化された Word 文書と .doc 形式は表示できません"),
        )
        for name, data, text in cases:
            with self.subTest(name=name):
                path = self.directory / name
                path.write_bytes(data)
                result = self.convert(path, expected_returncode=1)
                self.assertEqual(result.stdout, "")
                self.assertIn("Word の差分変換に失敗しました:", result.stderr)
                self.assertNotIn("Traceback", result.stderr)
                if text is not None:
                    self.assertIn(text, result.stderr)
        result = self.convert(self.directory / "存在しない.docx", expected_returncode=1)
        self.assertIn("Word の差分変換に失敗しました:", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_word_load_failure_for_extensionless_blob_uses_word_label(self):
        path = self.document("blob")
        path.write_bytes(rewrite_zip(path.read_bytes(), {"word/document.xml": b"broken XML"}))
        result = self.convert(path, expected_returncode=1)
        self.assertIn("Word の差分変換に失敗しました:", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_help_and_missing_argument_include_word(self):
        result = subprocess.run([sys.executable, "-B", str(CONVERTER), "--help"],
                                capture_output=True, encoding="utf-8", env=self.env, timeout=120)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Excel/Word file or Git temporary blob", result.stdout)
        result = subprocess.run([sys.executable, "-B", str(CONVERTER)],
                                capture_output=True, encoding="utf-8", env=self.env, timeout=120)
        self.assertEqual(result.returncode, 2)
        self.assertIn("Excel か Word のファイルを指定してください", result.stderr)

    def test_sniffer_never_raises_and_input_size_includes_word_xml(self):
        name = "word_test_converter"
        spec = importlib.util.spec_from_file_location(name, CONVERTER)
        converter = importlib.util.module_from_spec(spec)
        sys.modules[name] = converter
        self.addCleanup(sys.modules.pop, name, None)
        spec.loader.exec_module(converter)
        for data in (b"", b"PKbroken", b"plain", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            self.assertEqual(converter.package_kind(data), "excel")
        data = make_docx(body(), parts={"word/media/image1.png": b"png"})
        self.assertEqual(converter.package_kind(data), "word")
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            expected = sum(info.file_size for info in archive.infolist()
                           if info.filename.startswith("word/") and info.filename.endswith(".xml"))
        self.assertEqual(converter.input_size(data), expected)
        cache = self.directory / "cache"
        cache.mkdir()
        (cache / (hashlib.sha256(data).hexdigest() + ".txt")).write_bytes(b"cached\n")
        with mock.patch.object(converter, "package_kind", side_effect=AssertionError("cache hit sniffed")):
            self.assertEqual(converter.cached_convert(data, cache), ("cached\n", True))


class WordCacheTests(WordConverterTestCase):
    def setUp(self):
        super().setUp()
        self.cache = self.directory / "変換 キャッシュ"
        self.options = ("--cache-dir", str(self.cache))
        self.wait_env = dict(self.env, EXCEL_TEXTCONV_TIMEOUT="0")
        self.slow_env = dict(self.env, EXCEL_TEXTCONV_TIMEOUT="0.000001")

    def key(self, path):
        return hashlib.sha256(path.read_bytes()).hexdigest()

    def hold_worker_lock(self):
        self.cache.mkdir(parents=True, exist_ok=True)
        (self.cache / "worker.lock").touch()

    def run_worker(self):
        (self.cache / "worker.lock").unlink(missing_ok=True)
        result = subprocess.run([sys.executable, "-B", str(self.converter), "--worker", *self.options],
                                capture_output=True, encoding="utf-8", env=self.env, timeout=120)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_cache_reuse_needs_no_word_module(self):
        path = self.document()
        expected = self.convert(path, options=self.options, env=self.wait_env).stdout
        self.copy_converter()
        self.assertEqual(self.convert(path, options=self.options).stdout, expected)
        cached = self.cache / (self.key(path) + ".txt")
        cached.write_bytes(b"cached result\n")
        self.assertEqual(self.convert(path, options=self.options).stdout, "cached result\n")

    def test_queued_placeholder_ends_with_headings_and_worker_converts(self):
        path = self.document(headers={"default": body("合成ヘッダー")})
        expected = self.convert(path).stdout
        headings = [line for line in expected.splitlines() if line.startswith("# ")]
        self.hold_worker_lock()
        first = self.convert(path, options=self.options, env=self.slow_env).stdout
        self.assertTrue(first.startswith("# Word を裏で変換中 (" + self.key(path)[:12] + ")\n"))
        self.assertIn("# 完了後に選び直すか R で表示\n", first)
        placeholder_headings = first.splitlines()[2:]
        self.assertTrue(placeholder_headings, first)
        self.assertEqual(placeholder_headings, headings[:len(placeholder_headings)])
        self.assertTrue(first.endswith("\n"), first)
        self.assertTrue((self.cache / (self.key(path) + ".in")).exists())
        self.assertEqual(self.convert(path, options=self.options, env=self.slow_env).stdout, first)
        self.run_worker()
        self.assertFalse((self.cache / (self.key(path) + ".in")).exists())
        self.assertEqual(self.convert(path, options=self.options).stdout, expected)

    def test_slow_conversion_finishes_in_a_background_worker(self):
        self.copy_converter(module=(
            'import time\n'
            'def render_text(data):\n    time.sleep(0.5)\n    return "# 本文\\n合成の出力\\n"\n'
            'def placeholder_lines(data):\n    return ["# 本文"]\n'
        ))
        path = self.document()
        started = time.monotonic()
        result = self.convert(path, options=self.options, env=self.slow_env)
        self.assertLess(time.monotonic() - started, 15)
        self.assertTrue(result.stdout.startswith("# Word を裏で変換中"))
        cached = self.cache / (self.key(path) + ".txt")
        deadline = time.monotonic() + 30
        while not cached.exists():
            self.assertLess(time.monotonic(), deadline, "background Word conversion did not finish")
            time.sleep(0.05)
        self.assertEqual(self.convert(path, options=self.options).stdout, "# 本文\n合成の出力\n")
        while (self.cache / "worker.lock").exists():
            self.assertLess(time.monotonic(), deadline, "background Word worker did not stop")
            time.sleep(0.05)

    def test_worker_failure_is_reported_as_word_without_traceback(self):
        self.copy_converter()
        path = self.document("一時 blob")
        self.hold_worker_lock()
        first = self.convert(path, options=self.options, env=self.slow_env).stdout
        self.assertTrue(first.startswith("# Word を裏で変換中"))
        self.run_worker()
        result = self.convert(path, options=self.options, expected_returncode=1)
        self.assertEqual(result.stdout, "")
        self.assertIn("Word の差分変換に失敗しました:", result.stderr)
        self.assertIn("word_ooxml.py がありません", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_timeout_zero_waits_even_with_a_queued_request(self):
        path = self.document()
        self.hold_worker_lock()
        first = self.convert(path, options=self.options, env=self.slow_env).stdout
        self.assertTrue(first.startswith("# Word を裏で変換中"))
        self.assertEqual(self.convert(path, options=self.options, env=self.wait_env).stdout,
                         self.convert(path).stdout)


@unittest.skipUnless(shutil.which("git"), "Git is required for Word textconv integration tests")
class WordGitTextconvTests(WordConverterTestCase):
    def setUp(self):
        super().setUp()
        self.repo = self.directory / "日本語 repository"
        self.repo.mkdir()
        home = self.directory / "home"
        home.mkdir()
        self.env = isolated_env(home)
        self.env["PYTHONUTF8"] = "1"
        self.env["PYTHONDONTWRITEBYTECODE"] = "1"
        self.copy_converter()
        build = self.converter.with_suffix("")
        build.mkdir()
        shutil.copyfile(WORD_MODULE, build / "word_ooxml.py")
        self.git("init", "--quiet")
        self.git("config", "user.name", "Word test")
        self.git("config", "user.email", "word-test@example.invalid")
        self.git("config", "core.quotePath", "false")
        self.git("config", "core.autocrlf", "false")
        self.git("config", "diff.word.textconv", shlex.join([sys.executable, "-B", str(self.converter)]))
        self.git("config", "diff.word.binary", "true")
        (self.repo / ".gitattributes").write_bytes(b"*.docx diff=word\n")
        self.git("add", "--", ".gitattributes")
        self.git("commit", "--quiet", "-m", "Word textconv の設定")

    def git(self, *args):
        result = subprocess.run(["git", "--no-pager", *args], cwd=self.repo, env=self.env,
                                capture_output=True, encoding="utf-8", timeout=120)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout

    def commit_document(self, path):
        self.git("add", "--", path.name)
        self.git("commit", "--quiet", "-m", "合成文書を追加")

    def test_staged_and_unstaged_changes_compare_correct_word_blobs(self):
        path = self.repo / "合成 文書.docx"
        path.write_bytes(make_docx(body("元の本文")))
        self.commit_document(path)
        path.write_bytes(make_docx(body("ステージ済みの本文")))
        self.git("add", "--", path.name)
        path.write_bytes(make_docx(body("作業中の本文")))
        staged = self.git("diff", "--cached", "--textconv", "--", path.name)
        self.assertIn('-"元の本文"', staged)
        self.assertIn('+"ステージ済みの本文"', staged)
        self.assertNotIn("作業中の本文", staged)
        unstaged = self.git("diff", "--textconv", "--", path.name)
        self.assertIn('-"ステージ済みの本文"', unstaged)
        self.assertIn('+"作業中の本文"', unstaged)
        self.assertNotIn("Binary files", staged + unstaged)

    def test_formatting_only_change_is_visible_in_word_diff(self):
        path = self.repo / "書式 文書.docx"
        path.write_bytes(make_docx(body("書式だけ変更")))
        self.commit_document(path)
        path.write_bytes(make_docx(body("書式だけ変更", bold=True)))
        diff = self.git("diff", "--textconv", "--", path.name)
        self.assertIn('-"書式だけ変更"', diff)
        self.assertIn('+"書式だけ変更" 太字', diff)
        self.assertRegex(diff, r"#[0-9a-f]{8}")
        self.assertNotIn("Binary files", diff)

    def test_cached_word_driver_reuses_both_blobs(self):
        path = self.repo / "キャッシュ 文書.docx"
        path.write_bytes(make_docx(body("元の本文")))
        self.commit_document(path)
        path.write_bytes(make_docx(body("変更した本文")))
        expected = self.git("diff", "--textconv", "--", path.name)
        cache = self.directory / "Git キャッシュ"
        self.env["EXCEL_TEXTCONV_TIMEOUT"] = "0"
        command = self.git("config", "diff.word.textconv").strip()
        self.git("config", "diff.word.textconv", command + " --cache-dir " + shlex.quote(str(cache)))
        self.assertEqual(self.git("diff", "--textconv", "--", path.name), expected)
        self.assertEqual(self.git("diff", "--textconv", "--", path.name), expected)
        self.assertEqual(len(list(cache.glob("*.txt"))), 2)


if __name__ == "__main__":
    unittest.main()
