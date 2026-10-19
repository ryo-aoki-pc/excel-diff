"""Exercise two-document imports without opening WinMerge or lazygit."""

import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import sys
import tempfile
import unittest
from unittest import mock

from word_fixtures import hyperlink, make_docx, p, run, table, unpack


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "excel-stage.py"
HEADER = "header:1:default"
SELECTED = {("body", "1"), ("body", "3 R2C2"), ("body", "4"), (HEADER, "1")}


def document(updated=False, main_type="docx", **options):
    body = p("項目 A: 新しい値" if updated else "項目 A: 元の値")
    body += p("項目 B: 25" if updated else "項目 B: 10")
    body += table([["項目", "値"], ["電圧", "24" if updated else "12"]])
    body += p(runs=run("書式のみ", "<w:b/>" if updated else "")) + p("保持する末尾")
    return make_docx(body, main_type=main_type,
                     headers={"default": p("合成ヘッダー: 新しい値" if updated else "合成ヘッダー: 元の値")},
                     footers={"default": p("合成フッター")}, **options)


def comment_paragraph(text):
    return p(runs=run(text) + '<w:r><w:commentReference w:id="1"/></w:r>')


def reference_document(updated=False):
    name, ident = ("_TocSyntheticNew", "8") if updated else ("_TocSyntheticOld", "1")
    heading = p(runs='<w:bookmarkStart w:id="' + ident + '" w:name="' + name + '"/>'
                + run("合成の見出し") + '<w:bookmarkEnd w:id="' + ident + '"/>')
    return make_docx(p(runs=hyperlink("更新後の参照" if updated else "更新前の参照", anchor=name))
                     + heading + p("末尾"))


@unittest.skipUnless(shutil.which("git"), "Git is required for merge picker tests")
class WordMergeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scripts_path = str(SCRIPT.parent)
        sys.path.insert(0, cls.scripts_path)
        spec = importlib.util.spec_from_file_location("word_merge_integration", SCRIPT)
        cls.stage = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = cls.stage
        spec.loader.exec_module(cls.stage)
        import word_ooxml
        import word_stage_ooxml
        cls.word, cls.core = word_ooxml, word_stage_ooxml

    @classmethod
    def tearDownClass(cls):
        sys.path.remove(cls.scripts_path)
        sys.modules.pop("word_merge_integration", None)

    def setUp(self):
        self.directory = Path(tempfile.mkdtemp(prefix="word-merge-test-"))
        self.addCleanup(self.stage.cleanup_session, self.directory, retries=0, delay=0)
        self.target = self.directory / "取り込み先 と空白.docx"
        self.source = self.directory / "取り込み元 と空白.docx"
        self.target.write_bytes(document())
        self.source.write_bytes(document(True))
        self.env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
        self.env.update(HOME=str(self.directory), XDG_CONFIG_HOME=str(self.directory / ".config"),
                        GIT_CONFIG_GLOBAL=os.devnull, GIT_CONFIG_NOSYSTEM="1", GIT_TERMINAL_PROMPT="0")
        environment = mock.patch.dict(os.environ, self.env, clear=True)
        environment.start()
        self.addCleanup(environment.stop)
        self.number = 0

    def git(self, *args, cwd=None, input=None):
        result = subprocess.run(["git", "--no-pager", *args], cwd=cwd or self.directory,
                                env=self.env, input=input, capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors="replace"))
        return result.stdout

    def prepare(self, tool="lazygit"):
        self.number += 1
        directory = self.directory / ("session-" + str(self.number))
        return directory, self.stage.prepare_merge_session(self.target, self.source, directory, tool=tool)

    def select(self, directory, state, selected):
        picker = directory / "picker"
        for part in state["parts"]:
            filename = part["filename"]
            baseline = self.git("show", "HEAD:" + filename, cwd=picker).decode("utf-8").splitlines()
            changed = (picker / filename).read_text(encoding="utf-8").splitlines()
            lines = [new if (part["id"], json.loads(old)[0].split("/", 1)[0]) in selected else old
                     for old, new in zip(baseline, changed)]
            (picker / filename).write_bytes(("\n".join(lines) + ("\n" if lines else "")).encode("utf-8"))
            self.git("add", "--", filename, cwd=picker)

    def win_paths(self, directory, state):
        return tuple(directory / state["winmerge"][key] for key in ("left", "right", "baseline"))

    def copied_lines(self, directory, state, selected=None):
        left, right, _ = self.win_paths(directory, state)
        baseline = left.read_text(encoding="utf-8-sig").splitlines()
        changed = right.read_text(encoding="utf-8-sig").splitlines()
        story_id, story_index = None, -1
        for index, line in enumerate(baseline):
            if line.startswith("# 部分 "):
                story_index += 1
                story_id = state["parts"][story_index]["id"]
            elif line.startswith("[") and (selected is None or
                    (story_id, json.loads(line)[0].split("/", 1)[0]) in selected):
                baseline[index] = changed[index]
        return baseline

    def save_win(self, directory, state, selected=None, encoding="utf-8-sig", separator="\n", final="\n"):
        left, _, _ = self.win_paths(directory, state)
        left.write_bytes((separator.join(self.copied_lines(directory, state, selected)) + final).encode(encoding))

    def run_win(self, directory, state, edit=None):
        left, _, _ = self.win_paths(directory, state)

        def launch(command, *args, **kwargs):
            if edit is not None:
                data = edit(left.read_bytes())
                if data is not None:
                    left.write_bytes(data)
            return subprocess.CompletedProcess(command, 0)

        with mock.patch.object(self.stage, "find_winmerge", return_value="WinMergeU.exe"), mock.patch.object(
            self.stage.subprocess, "run", side_effect=launch
        ) as launched, mock.patch.object(self.stage.sys, "stdin", io.StringIO()):
            selected = self.stage.run_winmerge(directory)
        return selected, launched

    def assert_unchanged_failure(self, operation, pattern=None):
        before = self.target.read_bytes() if self.target.is_file() else None
        with self.assertRaisesRegex((ValueError, OSError), pattern or "."):
            operation()
        if before is not None:
            self.assertEqual(self.target.read_bytes(), before)

    def main(self, arguments, picker=None):
        output, errors = io.StringIO(), io.StringIO()
        with mock.patch.object(self.stage, "choose_tool", return_value="lazygit"), mock.patch.object(
            self.stage, "check_tools"
        ) as tools, mock.patch.object(self.stage, "run_picker", side_effect=picker) as lazy, mock.patch.object(
            self.stage, "run_winmerge"
        ) as win, mock.patch.object(self.stage.sys, "stdout", output), mock.patch.object(
            self.stage.sys, "stderr", errors
        ), mock.patch.object(self.stage.sys, "stdin", io.StringIO()):
            status = self.stage.main(["--merge", *arguments])
        return status, output.getvalue(), errors.getvalue(), tools, lazy, win

    def test_prepare_preserves_inputs_and_builds_word_picker_projections(self):
        snapshots = [(path.read_bytes(), path.stat().st_mtime_ns) for path in (self.target, self.source)]
        directory, state = self.prepare()
        self.assertEqual(state["format"], "word")
        self.assertEqual(state["target"], "file")
        self.assertEqual(state["path"], str(self.target.resolve()))
        self.assertEqual(state["source"], str(self.source.resolve()))
        self.assertFalse(state["applied"])
        self.assertEqual(state["differences"], 5)
        self.assertEqual(state["unsupported"], [])
        self.assertNotIn("sheets", state)
        self.assertEqual((directory / "base.docx").read_bytes(), snapshots[0][0])
        self.assertEqual((directory / "work.docx").read_bytes(), snapshots[1][0])
        comparison = self.core.compare(snapshots[0][0], snapshots[1][0])
        for number, (part, story) in enumerate(zip(state["parts"], comparison.stories), 1):
            self.assertEqual((part["id"], part["title"]), (story.id, story.title))
            self.assertTrue(part["filename"].startswith("{:03d}-".format(number)))
            self.assertTrue(part["filename"].endswith(".word.txt"))
            baseline = (directory / "baseline" / part["filename"]).read_bytes()
            self.assertEqual(hashlib.sha256(baseline).hexdigest(), part["baseline_sha256"])
            self.assertEqual(baseline.decode("utf-8"), story.baseline)
            self.assertEqual(self.git("show", "HEAD:" + part["filename"], cwd=directory / "picker").decode(), story.baseline)
            self.assertEqual((directory / "picker" / part["filename"]).read_text(encoding="utf-8"), story.changed)
        self.assertEqual(self.git("log", "-1", "--format=%s", cwd=directory / "picker").decode().strip(), "取り込み先の Word")
        self.assertEqual(self.git("config", "user.name", cwd=directory / "picker").decode().strip(), "Word staging")
        self.assertEqual(self.git("config", "user.email", cwd=directory / "picker").decode().strip(), "word-stage@example.invalid")
        self.assertEqual([(path.read_bytes(), path.stat().st_mtime_ns) for path in (self.target, self.source)], snapshots)

    def test_formatting_only_change_is_visible_and_selectable(self):
        self.target.write_bytes(make_docx(p("書式だけ") + p("末尾")))
        self.source.write_bytes(make_docx(p(runs=run("書式だけ", '<w:b/><w:color w:val="FF0000"/>')) + p("末尾")))
        directory, state = self.prepare()
        self.assertEqual(state["differences"], 1)
        self.select(directory, state, {("body", "1")})
        self.assertEqual(self.stage.apply_merge_session(directory), 1)
        records = [json.loads(line) for line in self.core.compare(self.target.read_bytes(), self.source.read_bytes()).stories[0].baseline.splitlines()]
        self.assertIn("太字", records[0][4])
        self.assertEqual(self.core.compare(self.target.read_bytes(), self.source.read_bytes()).differences, 0)

    def test_same_file_normalized_path_and_hardlink_are_rejected(self):
        (self.directory / "sub").mkdir()
        for source in (self.target, self.directory / "sub" / ".." / self.target.name):
            self.assert_unchanged_failure(lambda: self.stage.prepare_merge_session(
                self.target, source, self.directory / "bad", tool="lazygit"), "同じファイル")
        alias = self.directory / "hardlink.docx"
        try:
            os.link(self.target, alias)
        except OSError as error:
            self.skipTest(str(error))
        self.assert_unchanged_failure(lambda: self.stage.prepare_merge_session(
            self.target, alias, self.directory / "bad", tool="lazygit"), "ハードリンク")

    @unittest.skipUnless(os.name == "posix", "POSIX symlinks do not require Windows privileges")
    def test_symlink_on_either_side_is_rejected(self):
        alias = self.directory / "symlink.docx"
        alias.symlink_to(self.target)
        for target, source in ((alias, self.source), (self.target, alias)):
            self.assert_unchanged_failure(lambda: self.stage.prepare_merge_session(
                target, source, self.directory / "bad", tool="lazygit"))

    def test_mixed_excel_word_and_unsupported_extensions_report_the_right_family(self):
        for side in ("target", "source"):
            for suffix in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlt", ".xlsb", ".txt", ".doc"):
                with self.subTest(side=side, suffix=suffix):
                    path = self.directory / (side + suffix)
                    path.write_bytes(getattr(self, side).read_bytes())
                    target, source = (path, self.source) if side == "target" else (self.target, path)
                    pattern = "Excel と Word は混ぜて取り込めません" if suffix not in (".txt", ".doc") else "docx.*docm.*dotx.*dotm"
                    self.assert_unchanged_failure(lambda: self.stage.prepare_merge_session(
                        target, source, self.directory / "bad", tool="lazygit"), pattern)

    def test_all_four_word_suffixes_are_accepted_including_mixed_templates(self):
        for suffix in ("docx", "docm", "dotx", "dotm"):
            with self.subTest(suffix=suffix):
                self.source = self.directory / ("取り込み元 と空白." + suffix)
                self.source.write_bytes(document(True, main_type=suffix))
                _, state = self.prepare()
                self.assertEqual(state["format"], "word")
                self.assertEqual(state["differences"], 5)

    def test_target_validation_precedes_reading_word_documents(self):
        for target in (self.directory / "missing.docx", self.directory / "folder.docx"):
            if target.name.startswith("folder"):
                target.mkdir()
            with mock.patch.object(self.core, "compare") as comparison:
                status, _, errors, _, lazy, win = self.main([str(target), str(self.source)])
            self.assertEqual(status, 1)
            self.assertIn("Word の取り込みに失敗しました:", errors)
            comparison.assert_not_called()
            lazy.assert_not_called()
            win.assert_not_called()
        with mock.patch.object(self.stage.os, "access", return_value=False):
            self.assert_unchanged_failure(lambda: self.prepare(), "読み取り専用")

    def test_unsupported_changes_on_either_side_are_hidden_for_both_tools(self):
        for side in ("target", "source"):
            for tool in ("lazygit", "winmerge"):
                with self.subTest(side=side, tool=tool):
                    self.target.write_bytes(make_docx(p("比較する前") + p("末尾")))
                    self.source.write_bytes(make_docx(p("比較する後") + p("末尾")))
                    getattr(self, side).write_bytes(make_docx(comment_paragraph("比較する注釈") + p("末尾"),
                        comments='<w:comment w:id="1">' + p("合成の注記") + '</w:comment>'))
                    directory, state = self.prepare(tool)
                    self.assertTrue(state["unsupported"])
                    self.assertIn("コメント", state["unsupported"][0]["reason"])
                    comparison = self.core.compare(self.target.read_bytes(), self.source.read_bytes())
                    blocked = {item.slot for item in comparison.unsupported if item.story_id == "body"}
                    raw = comparison.stories[0].raw_changed.splitlines()
                    if tool == "lazygit":
                        part = state["parts"][0]
                        before = self.git("show", "HEAD:" + part["filename"], cwd=directory / "picker").decode().splitlines()
                        after = (directory / "picker" / part["filename"]).read_text(encoding="utf-8").splitlines()
                    else:
                        left, right, _ = self.win_paths(directory, state)
                        before = [line for line in left.read_text(encoding="utf-8-sig").splitlines() if line.startswith("[")]
                        after = [line for line in right.read_text(encoding="utf-8-sig").splitlines() if line.startswith("[")]
                    blocked_lines = [index for index, line in enumerate(before)
                                     if json.loads(line)[0].split("/", 1)[0] in blocked]
                    self.assertTrue(blocked_lines)
                    self.assertTrue(any(before[index] != raw[index] for index in blocked_lines))
                    for index in blocked_lines:
                        self.assertEqual(before[index], after[index])

    def test_only_unsupported_changes_are_summarized_without_opening_picker(self):
        self.target.write_bytes(make_docx(p("比較する前") + p("末尾")))
        self.source.write_bytes(make_docx(comment_paragraph("比較する後") + p("末尾"),
            comments='<w:comment w:id="1">' + p("合成の注記") + '</w:comment>'))
        status, output, errors, _, lazy, win = self.main([str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors)
        self.assertIn("取り込めない違い（1 か所）: 本文 #1（コメント）", output)
        self.assertIn("取り込める差分がありません（取り込めない違い 1 か所）。取り込み先は変えていません。", output)
        lazy.assert_not_called()
        win.assert_not_called()

    def test_unsupported_summary_precedes_picker_and_limits_twenty_examples(self):
        self.target.write_bytes(make_docx("".join(p("比較する前 " + str(i)) for i in range(23)) + p("末尾")))
        self.source.write_bytes(make_docx("".join(comment_paragraph("比較する後 " + str(i)) for i in range(23))
                                         + p("末尾"), comments='<w:comment w:id="1">' + p("注記") + '</w:comment>'))
        status, output, errors, _, lazy, _ = self.main([str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors)
        self.assertIn("取り込めない違い（23 か所）:", output)
        self.assertIn("本文 #20（コメント）", output)
        self.assertNotIn("本文 #21", output)
        self.assertIn("、ほか 3 か所", output)
        lazy.assert_not_called()
        self.target.write_bytes(make_docx(p("比較する前") + p("選べる前") + p("末尾")))
        self.source.write_bytes(make_docx(comment_paragraph("比較する後") + p("選べる後") + p("末尾"),
            comments='<w:comment w:id="1">' + p("注記") + '</w:comment>'))
        output = io.StringIO()
        def picker(directory, merge=False):
            self.assertTrue(merge)
            self.assertIn("取り込めない違い（1 か所）:", output.getvalue())
            return 0
        with mock.patch.object(self.stage, "choose_tool", return_value="lazygit"), mock.patch.object(
            self.stage, "check_tools"
        ), mock.patch.object(self.stage, "run_picker", side_effect=picker), mock.patch.object(self.stage.sys, "stdout", output):
            self.assertEqual(self.stage.main(["--merge", str(self.target), str(self.source)]), 0)

    def test_selected_slots_preserve_unselected_values_macros_parts_and_source(self):
        self.target.write_bytes(document(main_type="docm", parts={"custom/opaque.bin": b"unchanged opaque bytes"}))
        source = self.source.read_bytes(), self.source.stat().st_mtime_ns
        base = self.target.read_bytes()
        directory, state = self.prepare()
        self.select(directory, state, SELECTED)
        self.assertEqual(self.stage.apply_merge_session(directory), len(SELECTED))
        expected = self.core.compare(base, source[0]).rebuild(SELECTED)
        self.assertEqual(self.word.render_text(self.target.read_bytes()), self.word.render_text(expected))
        self.assertIn("項目 B: 10", self.word.render_text(self.target.read_bytes()))
        self.assertIn("合成ヘッダー: 新しい値", self.word.render_text(self.target.read_bytes()))
        for name in ("word/vbaProject.bin", "custom/opaque.bin", "word/footer1.xml"):
            self.assertEqual(unpack(self.target.read_bytes())[name], unpack(base)[name])
        self.assertEqual((self.source.read_bytes(), self.source.stat().st_mtime_ns), source)
        saved = json.loads((directory / "session.json").read_text(encoding="utf-8"))
        self.assertTrue(saved["applied"])
        self.assertEqual(saved["count"], len(SELECTED))

    def test_no_difference_or_selection_preserves_target_bytes_and_timestamp(self):
        before = self.target.read_bytes(), self.target.stat().st_mtime_ns
        directory, _ = self.prepare()
        self.assertEqual(self.stage.apply_merge_session(directory), 0)
        self.assertEqual((self.target.read_bytes(), self.target.stat().st_mtime_ns), before)
        self.assertFalse(json.loads((directory / "session.json").read_text(encoding="utf-8"))["applied"])
        self.source.write_bytes(self.target.read_bytes())
        status, output, errors, _, lazy, win = self.main([str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors)
        self.assertIn("差分がありません。取り込み先は変えていません。", output)
        lazy.assert_not_called()
        win.assert_not_called()

    def test_second_apply_and_wrong_session_targets_are_rejected(self):
        directory, state = self.prepare()
        self.select(directory, state, {("body", "1")})
        self.assert_unchanged_failure(lambda: self.stage.apply_session(directory), "取り込みの選択画面")
        self.assertEqual(self.stage.apply_merge_session(directory), 1)
        self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory), "反映済み")
        directory, state = self.prepare()
        state["target"] = "index"
        self.stage.write_json(directory / "session.json", state)
        self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory), "部分ステージの選択画面")

    def test_lazygit_malformed_records_and_modified_picker_file_set_are_rejected(self):
        for change in ("manual", "removed", "added", "file-removed", "file-added"):
            with self.subTest(change=change):
                directory, state = self.prepare()
                part = state["parts"][0]
                picker, filename = directory / "picker", part["filename"]
                path = picker / filename
                lines = path.read_text(encoding="utf-8").splitlines()
                if change == "manual":
                    record = json.loads(lines[0])
                    record[3] = "手で作った値"
                    lines[0] = json.dumps(record, ensure_ascii=False)
                elif change == "removed":
                    del lines[0]
                elif change == "added":
                    lines.insert(0, lines[0])
                elif change == "file-removed":
                    path.unlink()
                else:
                    (picker / "extra.word.txt").write_bytes(b"extra")
                if change not in ("file-removed", "file-added"):
                    path.write_bytes(("\n".join(lines) + "\n").encode("utf-8"))
                self.git("add", "--all", cwd=picker)
                self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory),
                    "選択画面のファイル" if change.startswith("file-") else "本文")

    def test_changed_or_deleted_target_leaves_no_temporary_file(self):
        for deleted in (False, True):
            with self.subTest(deleted=deleted):
                self.target.write_bytes(document())
                directory, state = self.prepare()
                self.select(directory, state, {("body", "1")})
                if deleted:
                    self.target.unlink()
                else:
                    self.target.write_bytes(make_docx(p("別の操作による変更")))
                files = set(self.directory.iterdir())
                with self.assertRaises((ValueError, OSError)):
                    self.stage.apply_merge_session(directory)
                self.assertEqual(set(self.directory.iterdir()), files)
                self.assertFalse(json.loads((directory / "session.json").read_text(encoding="utf-8"))["applied"])

    def test_source_is_the_opening_snapshot(self):
        directory, state = self.prepare()
        self.select(directory, state, {("body", "1")})
        self.source.write_bytes(make_docx(p("あとで変更した取り込み元")))
        self.assertEqual(self.stage.apply_merge_session(directory), 1)
        self.assertIn("項目 A: 新しい値", self.word.render_text(self.target.read_bytes()))
        self.assertNotIn("あとで変更", self.word.render_text(self.target.read_bytes()))

    def test_snapshot_and_baseline_tampering_are_rejected(self):
        for change in ("base.docx", "work.docx", "old", "baseline", "forged-baseline"):
            with self.subTest(change=change):
                directory, state = self.prepare()
                self.select(directory, state, {("body", "1")})
                if change == "old":
                    del state["base_sha256"]
                elif change in ("baseline", "forged-baseline"):
                    part = state["parts"][0]
                    path = directory / "baseline" / part["filename"]
                    data = path.read_bytes().replace("元の値".encode(), "偽の値".encode())
                    path.write_bytes(data)
                    if change == "forged-baseline":
                        part["baseline_sha256"] = hashlib.sha256(data).hexdigest()
                else:
                    (directory / change).write_bytes(make_docx(p("変更した控え")))
                self.stage.write_json(directory / "session.json", state)
                self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory), "変更|射影|形式")

    def test_merge_picker_configuration_uses_word_renderers_and_texts(self):
        directory, _ = self.prepare()
        config = json.loads(self.stage.picker_config(directory, merge=True).read_text(encoding="utf-8"))
        command = config["customCommands"][0]
        self.assertIn("--apply-merge-session", command["command"])
        self.assertEqual(command["description"], "選択した変更を取り込み先へ書き込む (反映後 q で戻る)")
        self.assertEqual(command["outputTitle"], "Word の取り込み")
        self.assertEqual(config["git"]["diffRenderers"], [
            {"command": "delta --no-gitconfig --dark --paging=never --syntax-theme=none --tabs=4 --side-by-side --wrap-max-lines=unlimited", "name": "delta side-by-side wrapped"},
            {"command": "delta --no-gitconfig --dark --paging=never --syntax-theme=none --tabs=4 --side-by-side", "name": "delta side-by-side"},
            {"type": "rawGit", "name": "default"}])

    def test_winmerge_text_has_bom_lf_aligned_headers_and_word_paths(self):
        directory, state = self.prepare("winmerge")
        left, right, baseline = self.win_paths(directory, state)
        self.assertEqual(left.relative_to(directory).as_posix(), "winmerge/取り込み先.word.txt")
        self.assertEqual(right.relative_to(directory).as_posix(), "winmerge/取り込み元.word.txt")
        self.assertEqual(baseline.relative_to(directory).as_posix(), "baseline/取り込み先.word.txt")
        self.assertEqual(left.read_bytes(), baseline.read_bytes())
        texts = [path.read_bytes() for path in (left, right)]
        for text in texts:
            self.assertTrue(text.startswith(b"\xef\xbb\xbf"))
            self.assertNotIn(b"\r", text)
            self.assertTrue(text.endswith(b"\n"))
        rows = [text.decode("utf-8-sig").splitlines() for text in texts]
        self.assertEqual(len(rows[0]), len(rows[1]))
        headers = [line for line in rows[0] if line.startswith("#")]
        self.assertEqual(headers, [line for line in rows[1] if line.startswith("#")])
        self.assertEqual(headers[:2], ["# 右（取り込み元）から左（取り込み先）へ、取り込む行をコピーする",
                                     "# 上書き保存（Ctrl+S）してから WinMerge を閉じると、取り込み先の文書に書き込む"])
        self.assertIn('# 部分 1: "本文"', headers)
        self.assertIn('# 部分 2: "ヘッダー（セクション 1・既定）"', headers)

    def test_winmerge_some_and_all_record_copies(self):
        for selected in (SELECTED, None):
            with self.subTest(selected=selected):
                self.target.write_bytes(document())
                directory, state = self.prepare("winmerge")
                self.save_win(directory, state, selected)
                chosen, _ = self.run_win(directory, state)
                self.assertEqual(len(chosen), len(SELECTED) if selected else 5)
                self.assertEqual(self.stage.apply_merge_session(directory), len(chosen))
                if selected is None:
                    self.assertEqual(self.core.compare(self.target.read_bytes(), self.source.read_bytes()).differences, 0)
                else:
                    self.assertIn("項目 B: 10", self.word.render_text(self.target.read_bytes()))

    def test_winmerge_close_without_save_preserves_target_and_warns_on_quick_exit(self):
        before = self.target.read_bytes(), self.target.stat().st_mtime_ns
        directory, state = self.prepare("winmerge")
        errors = io.StringIO()
        with mock.patch.object(self.stage.sys, "stderr", errors), mock.patch.object(self.stage.time, "monotonic", side_effect=[10, 11]):
            selected, _ = self.run_win(directory, state)
        self.assertEqual(selected, set())
        self.assertIn("WinMerge がすぐ終了しました", errors.getvalue())
        self.assertEqual(self.stage.apply_merge_session(directory), 0)
        self.assertEqual((self.target.read_bytes(), self.target.stat().st_mtime_ns), before)

    def test_winmerge_line_endings_optional_bom_and_final_newline_are_accepted(self):
        for encoding, separator, final in (("utf-8-sig", "\r\n", "\r\n"), ("utf-8", "\n", "\n"),
                                            ("utf-8", "\n", ""), ("utf-8-sig", "\n", "\n\n")):
            with self.subTest(encoding=encoding, separator=repr(separator), final=repr(final)):
                self.target.write_bytes(document())
                directory, state = self.prepare("winmerge")
                self.save_win(directory, state, {("body", "1")}, encoding, separator, final)
                self.assertEqual(self.stage.apply_merge_session(directory), 1)
        self.target.write_bytes(document())
        directory, state = self.prepare("winmerge")
        lines = self.copied_lines(directory, state, {("body", "1")})
        left, _, _ = self.win_paths(directory, state)
        left.write_bytes("".join(line + ("\n" if index % 2 else "\r\n") for index, line in enumerate(lines)).encode("utf-8-sig"))
        self.assertEqual(self.stage.apply_merge_session(directory), 1)

    def test_winmerge_non_utf8_saves_are_rejected(self):
        for encoding in ("utf-16", "cp932"):
            with self.subTest(encoding=encoding):
                directory, state = self.prepare("winmerge")
                self.save_win(directory, state, encoding=encoding)
                self.assert_unchanged_failure(lambda: self.run_win(directory, state), "UTF-8")

    def test_winmerge_modified_headers_counts_and_invalid_records_have_locations(self):
        for change in ("header", "removed", "added", "reordered", "manual", "other-story"):
            with self.subTest(change=change):
                directory, state = self.prepare("winmerge")
                lines = self.copied_lines(directory, state)
                records = [index for index, line in enumerate(lines) if line.startswith("[")]
                first, second, last = records[0], records[1], records[-1]
                if change == "header":
                    lines[0] += "変更"
                elif change == "removed":
                    del lines[first]
                elif change == "added":
                    lines.insert(first, lines[first])
                elif change == "reordered":
                    lines[first], lines[second] = lines[second], lines[first]
                elif change == "other-story":
                    lines[first] = lines[last]
                else:
                    record = json.loads(lines[first])
                    record[3] = "手で打った値"
                    lines[first] = json.dumps(record, ensure_ascii=False)
                left, _, _ = self.win_paths(directory, state)
                left.write_bytes(("\n".join(lines) + "\n").encode("utf-8-sig"))
                self.assert_unchanged_failure(lambda: self.run_win(directory, state), "見出し" if change == "header" else "本文 #1")

    def test_winmerge_baseline_hash_and_recomputed_projection_are_checked(self):
        for forged in (False, True):
            with self.subTest(forged=forged):
                directory, state = self.prepare("winmerge")
                _, _, baseline = self.win_paths(directory, state)
                data = baseline.read_bytes().replace("元の値".encode(), "偽の値".encode())
                baseline.write_bytes(data)
                if forged:
                    state["winmerge"]["baseline_sha256"] = hashlib.sha256(data).hexdigest()
                    self.stage.write_json(directory / "session.json", state)
                self.assert_unchanged_failure(lambda: self.run_win(directory, state), "射影|控え|変更")

    def test_winmerge_arguments_keep_source_readonly_and_avoid_persistent_settings(self):
        directory, state = self.prepare("winmerge")
        _, launched = self.run_win(directory, state)
        left, right, _ = self.win_paths(directory, state)
        self.assertEqual(launched.call_args.args[0], ["WinMergeU.exe", "/s-", "/e", "/u", "/fl", "/wr", "/t", "text",
                         "/dl", "取り込み先（書き換える）: " + self.target.name,
                         "/dr", "取り込み元: " + self.source.name, str(left), str(right)])
        self.assertFalse(any(argument.lower().startswith(("/cfg", "/ignore", "/cp")) for argument in launched.call_args.args[0]))

    def test_core_selection_error_is_a_retryable_stage_error(self):
        directory, state = self.prepare("winmerge")
        left, _, _ = self.win_paths(directory, state)
        valid = ("\n".join(self.copied_lines(directory, state, {("body", "1")})) + "\n").encode("utf-8-sig")
        saved = []
        def launch(command, *args, **kwargs):
            saved.append(left.read_bytes())
            left.write_bytes(valid)
            return subprocess.CompletedProcess(command, 0)
        original_choose = self.core.Comparison.choose
        calls = [0]
        def choose(comparison, *args, **kwargs):
            calls[0] += 1
            if calls[0] == 1:
                raise self.word.DocumentError("選択の試験用エラー")
            return original_choose(comparison, *args, **kwargs)
        errors = io.StringIO()
        with mock.patch.object(self.stage, "find_winmerge", return_value="WinMergeU.exe"), mock.patch.object(
            self.stage.subprocess, "run", side_effect=launch
        ), mock.patch.object(self.core.Comparison, "choose", choose), mock.patch.object(
            self.stage.sys.stdin, "isatty", return_value=True
        ), mock.patch("builtins.input", return_value="y") as prompt, mock.patch.object(self.stage.sys, "stderr", errors):
            self.assertEqual(len(self.stage.run_winmerge(directory)), 1)
        self.assertEqual(len(saved), 2)
        self.assertEqual(saved[1], valid)
        self.assertIn("選択の試験用エラー", errors.getvalue())
        prompt.assert_called_once_with("WinMerge を開き直して直しますか [y/N]: ")

    def test_main_winmerge_applies_word_without_checking_lazygit_and_cleans_up(self):
        sessions = []
        original_run = subprocess.run
        def launch(command, *args, **kwargs):
            if command[0] != "WinMergeU.exe":
                return original_run(command, *args, **kwargs)
            directory = Path(command[-2]).parent.parent
            sessions.append(directory)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            self.save_win(directory, state, {("body", "1")})
            return subprocess.CompletedProcess(command, 0)
        output, errors = io.StringIO(), io.StringIO()
        with mock.patch.object(self.stage, "find_winmerge", return_value="WinMergeU.exe"), mock.patch.object(
            self.stage.subprocess, "run", side_effect=launch
        ), mock.patch.object(self.stage, "check_tools") as tools, mock.patch.object(
            self.stage, "run_picker"
        ) as lazy, mock.patch.object(self.stage.sys, "stdout", output), mock.patch.object(
            self.stage.sys, "stderr", errors
        ), mock.patch.object(self.stage.sys, "stdin", io.StringIO()):
            status = self.stage.main(["--merge", "--tool", "winmerge", "--", str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors.getvalue())
        self.assertIn("書き換えるファイル: " + str(self.target.resolve()), output.getvalue())
        self.assertIn("Word を読み込んでいます。大きい文書では画面が開くまで時間がかかります。", output.getvalue())
        self.assertIn("1 か所の変更を取り込み先へ書き込みました: " + str(self.target.resolve()), output.getvalue())
        self.assertIn("取り込む行を右から左へコピー", output.getvalue())
        self.assertNotIn("セル", output.getvalue())
        tools.assert_not_called()
        lazy.assert_not_called()
        self.assertTrue(sessions)
        self.assertTrue(all(not directory.exists() for directory in sessions))

    def test_main_lazygit_cancel_apply_and_failure_messages(self):
        before = self.target.read_bytes()
        status, output, errors, tools, _, _ = self.main([str(self.target), str(self.source)], picker=lambda directory, merge=False: 0)
        self.assertEqual(status, 0, errors)
        self.assertIn("取り込む変更が選ばれていません。取り込み先は変えていません。", output)
        self.assertIn("Enter → Space", output)
        tools.assert_called_once()
        self.assertEqual(self.target.read_bytes(), before)
        sessions = []
        def picker(directory, merge=False):
            sessions.append(directory)
            self.assertTrue(merge)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            self.select(directory, state, {("body", "1")})
            self.assertEqual(self.stage.apply_merge_session(directory), 1)
            return 0
        status, output, errors, _, _, _ = self.main([str(self.target), str(self.source)], picker=picker)
        self.assertEqual(status, 0, errors)
        self.assertIn("1 か所の変更を取り込み先へ書き込みました:", output)
        def fail(directory, merge=False):
            sessions.append(directory)
            raise self.stage.StageError("選択の試験用エラー")
        before = self.target.read_bytes()
        status, _, errors, _, _, _ = self.main([str(self.target), str(self.source)], picker=fail)
        self.assertEqual(status, 1)
        self.assertIn("Word の取り込みに失敗しました: 選択の試験用エラー", errors)
        self.assertEqual(self.target.read_bytes(), before)
        self.assertTrue(all(not directory.exists() for directory in sessions))

    def test_hidden_apply_merge_cli_uses_word_texts(self):
        directory, state = self.prepare()
        output = io.StringIO()
        with mock.patch.object(self.stage.sys, "stdout", output):
            self.assertEqual(self.stage.main(["--apply-merge-session", str(directory)]), 0)
        self.assertIn("変更が選択されていません。Enter・Space でハンクを選び、X で反映してください。", output.getvalue())
        self.select(directory, state, {("body", "1")})
        output = io.StringIO()
        with mock.patch.object(self.stage.sys, "stdout", output):
            self.assertEqual(self.stage.main(["--apply-merge-session", str(directory)]), 0)
        self.assertIn("1 か所の変更を取り込み先へ書き込みました。q で元の画面へ戻ってください。", output.getvalue())
        errors = io.StringIO()
        with mock.patch.object(self.stage.sys, "stderr", errors):
            self.assertEqual(self.stage.main(["--apply-merge-session", str(directory)]), 1)
        self.assertIn("Word の取り込みに失敗しました:", errors.getvalue())

    def test_equal_reference_swap_count_is_printed_only_after_successful_write(self):
        self.target.write_bytes(reference_document())
        self.source.write_bytes(reference_document(True))
        directory, state = self.prepare()
        self.select(directory, state, {("body", "1")})
        output = io.StringIO()
        with mock.patch.object(self.stage, "replace_file", side_effect=PermissionError("busy")), mock.patch.object(self.stage.sys, "stdout", output):
            self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory), "busy")
        self.assertNotIn("目次などの参照先", output.getvalue())
        with mock.patch.object(self.stage.sys, "stdout", output):
            self.assertEqual(self.stage.apply_merge_session(directory), 1)
        self.assertIn("目次などの参照先に合わせて、表示が同じ 1 か所を取り込み元のものにしました。", output.getvalue())
        self.assertIn(b"_TocSyntheticNew", unpack(self.target.read_bytes())["word/document.xml"])

    def test_unpaired_header_only_is_unsupported_even_without_changed_slots(self):
        self.target.write_bytes(make_docx(p("本文は同じ")))
        self.source.write_bytes(make_docx(p("本文は同じ"), headers={"default": p("追加したヘッダー")}))
        directory, state = self.prepare()
        self.assertEqual(state["selectable"], 0)
        self.assertTrue(any("ヘッダー・フッターの追加・削除" in item["reason"] for item in state["unsupported"]))
        status, output, errors, _, lazy, win = self.main([str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors)
        self.assertIn("取り込める差分がありません", output)
        self.assertNotIn("差分がありません。取り込み先は変えていません。", output)
        lazy.assert_not_called()
        win.assert_not_called()

    def test_unsliced_table_requires_all_changed_lines_for_both_tools(self):
        self.target.write_bytes(make_docx(p("先頭") + table([["項目", "値"], ["電圧", "12"]]) + p("末尾")))
        self.source.write_bytes(make_docx(p("先頭") + table([["項目", "値"], ["電圧", "24"], ["電流", "2"]]) + p("末尾")))
        for tool in ("lazygit", "winmerge"):
            with self.subTest(tool=tool):
                directory, state = self.prepare(tool)
                if tool == "lazygit":
                    part = state["parts"][0]
                    path = directory / "picker" / part["filename"]
                    lines = self.git("show", "HEAD:" + part["filename"], cwd=directory / "picker").decode().splitlines()
                    changed = path.read_text(encoding="utf-8").splitlines()
                else:
                    path, right, _ = self.win_paths(directory, state)
                    lines = path.read_text(encoding="utf-8-sig").splitlines()
                    changed = right.read_text(encoding="utf-8-sig").splitlines()
                differences = [index for index, (old, new) in enumerate(zip(lines, changed)) if old != new]
                self.assertGreater(len(differences), 1)
                lines[differences[0]] = changed[differences[0]]
                path.write_bytes(("\n".join(lines) + "\n").encode("utf-8"))
                if tool == "lazygit":
                    self.git("add", "--", part["filename"], cwd=directory / "picker")
                self.assert_unchanged_failure(lambda: self.stage.apply_merge_session(directory), "まとまりの変更行をすべて")
        directory, state = self.prepare()
        self.select(directory, state, {("body", "2")})
        self.assertEqual(self.stage.apply_merge_session(directory), 1)
        self.assertEqual(self.core.compare(self.target.read_bytes(), self.source.read_bytes()).differences, 0)

    def test_inserted_and_deleted_paragraphs_are_whole_place_selections(self):
        self.target.write_bytes(make_docx(p("先頭") + p("削除対象の段落") + p("共通の保持する段落") + p("末尾")))
        self.source.write_bytes(make_docx(p("先頭") + p("共通の保持する段落") + p("追加された新規文章") + p("末尾")))
        for tool in ("lazygit", "winmerge"):
            with self.subTest(tool=tool):
                self.target.write_bytes(make_docx(p("先頭") + p("削除対象の段落") + p("共通の保持する段落") + p("末尾")))
                directory, state = self.prepare(tool)
                comparison = self.core.compare(self.target.read_bytes(), self.source.read_bytes())
                selected = comparison.choose("body", comparison.stories[0].changed.splitlines(), projection="changed")
                if tool == "lazygit":
                    self.select(directory, state, selected)
                else:
                    self.save_win(directory, state)
                self.assertEqual(self.stage.apply_merge_session(directory), len(selected))
                self.assertEqual(self.core.compare(self.target.read_bytes(), self.source.read_bytes()).differences, 0)

    def test_unicode_record_separators_stay_escaped_and_can_be_imported(self):
        value = "項目 A: 前\u2028中\x85後\u2029終"
        for tool in ("lazygit", "winmerge"):
            with self.subTest(tool=tool):
                self.target.write_bytes(make_docx(p("項目 A: 元の値") + p("末尾")))
                self.source.write_bytes(make_docx(p(value) + p("末尾")))
                directory, state = self.prepare(tool)
                if tool == "lazygit":
                    texts = [(directory / "picker" / part["filename"]).read_bytes() for part in state["parts"]]
                    self.select(directory, state, {("body", "1")})
                else:
                    left, right, _ = self.win_paths(directory, state)
                    texts = [left.read_bytes(), right.read_bytes()]
                    self.save_win(directory, state, {("body", "1")})
                for text in texts:
                    for separator in ("\u2028", "\x85", "\u2029"):
                        self.assertNotIn(separator.encode("utf-8"), text)
                self.assertEqual(self.stage.apply_merge_session(directory), 1)
                self.assertEqual(self.core.compare(self.target.read_bytes(), self.source.read_bytes()).differences, 0)

    def test_winmerge_permission_error_retries_the_same_word_selection(self):
        original_run, original_replace = subprocess.run, self.stage.replace_file
        replacements = []
        def launch(command, *args, **kwargs):
            if command[0] != "WinMergeU.exe":
                return original_run(command, *args, **kwargs)
            directory = Path(command[-2]).parent.parent
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            self.save_win(directory, state, {("body", "1")})
            return subprocess.CompletedProcess(command, 0)
        def replace(*args):
            replacements.append(args)
            if len(replacements) == 1:
                raise PermissionError("busy")
            return original_replace(*args)
        output, errors = io.StringIO(), io.StringIO()
        with mock.patch.object(self.stage, "find_winmerge", return_value="WinMergeU.exe"), mock.patch.object(
            self.stage.subprocess, "run", side_effect=launch
        ), mock.patch.object(self.stage, "replace_file", side_effect=replace), mock.patch.object(
            self.stage.sys.stdin, "isatty", return_value=True
        ), mock.patch("builtins.input", return_value="") as prompt, mock.patch.object(
            self.stage.sys, "stdout", output
        ), mock.patch.object(self.stage.sys, "stderr", errors):
            status = self.stage.main(["--merge", "--tool", "winmerge", str(self.target), str(self.source)])
        self.assertEqual(status, 0, errors.getvalue())
        self.assertEqual(len(replacements), 2)
        self.assertEqual(replacements[0], replacements[1])
        prompt.assert_called_once_with("取り込み先を閉じてから Enter で再試行します（q で中止）: ")
        self.assertIn("1 か所の変更を取り込み先へ書き込みました:", output.getvalue())

    def test_replace_file_uses_word_temporary_prefix_and_permission_labels(self):
        digest = hashlib.sha256(self.target.read_bytes()).hexdigest()
        before = set(self.directory.iterdir())
        original_mkstemp = self.stage.tempfile.mkstemp
        with mock.patch.object(self.stage.tempfile, "mkstemp", wraps=original_mkstemp) as create:
            self.stage.replace_file(self.target, b"replacement", digest)
        self.assertEqual(create.call_args.kwargs["prefix"], ".word-merge-")
        self.assertEqual(Path(create.call_args.kwargs["dir"]), self.target.parent)
        self.assertEqual(set(self.directory.iterdir()), before)
        digest = hashlib.sha256(self.target.read_bytes()).hexdigest()
        with mock.patch.object(self.stage.os, "replace", side_effect=PermissionError("busy")), mock.patch.object(self.stage.time, "sleep"):
            with self.assertRaisesRegex(PermissionError, "取り込み先を Word などで開いているなら閉じてから再試行してください:"):
                self.stage.replace_file(self.target, b"replacement", digest)
        self.assertEqual(set(self.directory.iterdir()), before)

    @unittest.skipUnless(os.name == "nt", "Windows sharing violations")
    def test_open_word_target_is_rejected_before_document_comparison(self):
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                      wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        kernel.CreateFileW.restype = wintypes.HANDLE
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = kernel.CreateFileW(str(self.target), 0x80000000, 1, None, 3, 0x80, None)
        self.assertNotEqual(handle, wintypes.HANDLE(-1).value)
        try:
            with mock.patch.object(self.core, "compare") as comparison:
                self.assert_unchanged_failure(lambda: self.prepare(), "Word|開")
            comparison.assert_not_called()
        finally:
            kernel.CloseHandle(handle)

    @unittest.skipUnless(os.name == "posix", "POSIX mode bits")
    def test_replace_file_preserves_posix_mode(self):
        self.target.chmod(0o640)
        self.stage.replace_file(self.target, b"replacement", hashlib.sha256(self.target.read_bytes()).hexdigest())
        self.assertEqual(stat.S_IMODE(self.target.stat().st_mode), 0o640)


if __name__ == "__main__":
    unittest.main()
