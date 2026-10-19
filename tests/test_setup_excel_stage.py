"""Verify installer isolation, updates, and Git alias argument forwarding.

The small installed fixture exercises the installer's entry point/module layout;
the real workbook and selector behavior is covered by test_excel_stage.py.
"""

import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import venv


ROOT = Path(__file__).resolve().parents[1]
RUN = subprocess.run
ENV_BUILDER = venv.EnvBuilder
HELPER_SOURCE = (
    "import json, sys\n"
    "from excel_stage_ooxml import MARKER\n"
    "sys.stdout.reconfigure(encoding='utf-8')\n"
    "print(json.dumps({'arguments': sys.argv[1:], 'module': MARKER}, ensure_ascii=False))\n"
).encode("utf-8")
MODULE_SOURCE = b"MARKER = 'installed module'\n"
WORD_MODULE_SOURCE = b"MARKER = 'installed Word module'\n"
WORD_STAGE_MODULE_SOURCE = b"MARKER = 'installed Word stage module'\n"
# Keep a literal copy of the released Excel-only block for upgrade coverage.
OLD_ATTRIBUTE_BLOCK = (
    b"# BEGIN lazygit Excel textconv (setup-excel-diff.py)\n"
    b"*.[xX][lL][sS][xX] diff=excel\n"
    b"*.[xX][lL][sS][mM] diff=excel\n"
    b"*.[xX][lL][tT][xX] diff=excel\n"
    b"*.[xX][lL][tT][mM] diff=excel\n"
    b"*.[xX][lL][sS] diff=excel\n"
    b"*.[xX][lL][tT] diff=excel\n"
    b"# END lazygit Excel textconv (setup-excel-diff.py)\n"
)
VERSIONS = "3.1.5\n2.0.2\n6.1.3"


class SetupExcelStageTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="excel-stage-setup-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.checkout = self.directory / "checkout"
        scripts = self.checkout / "scripts"
        scripts.mkdir(parents=True)
        for name in ("setup-excel-diff.py", "excel-textconv.py"):
            shutil.copyfile(ROOT / "scripts" / name, scripts / name)
        shutil.copyfile(ROOT / "requirements-excel-diff.txt",
                        self.checkout / "requirements-excel-diff.txt")
        (scripts / "excel-stage.py").write_bytes(HELPER_SOURCE)
        (scripts / "excel_stage_ooxml.py").write_bytes(MODULE_SOURCE)
        (scripts / "word_ooxml.py").write_bytes(WORD_MODULE_SOURCE)
        (scripts / "word_stage_ooxml.py").write_bytes(WORD_STAGE_MODULE_SOURCE)
        spec = importlib.util.spec_from_file_location("excel_stage_setup", scripts / "setup-excel-diff.py")
        self.setup = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.setup)
        self.install = self.directory / "インストール 空白 O'Brien $dollar"
        self.config = self.directory / "global.gitconfig"
        self.attributes = self.directory / "config" / "git" / "attributes"
        self.attributes.parent.mkdir(parents=True)
        self.original_attributes = b"# Existing attributes\r\n*.csv diff=csv"
        self.attributes.write_bytes(self.original_attributes)
        self.env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
        self.env.update(
            GIT_CONFIG_GLOBAL=str(self.config),
            GIT_CONFIG_NOSYSTEM="1",
            GIT_CONFIG_SYSTEM=os.devnull,
            GIT_TERMINAL_PROMPT="0",
            XDG_CONFIG_HOME=str(self.directory / "config"),
            XDG_DATA_HOME=str(self.directory / "data"),
        )
        self.pip_calls = []

    def run_git(self, *args):
        return RUN(["git", *args], cwd=self.checkout, env=self.env,
                   capture_output=True, encoding="utf-8", check=True)

    def run_setup(self, *, check=False, uninstall=False, explicit_directory=True, versions=VERSIONS, env=None):
        arguments = ["setup-excel-diff.py"]
        if explicit_directory:
            arguments += ["--install-dir", str(self.install)]
        if check:
            arguments.append("--check")
        if uninstall:
            arguments.append("--uninstall")

        def run(command, *args, **kwargs):
            if command[1:4] == ["-m", "pip", "install"]:
                self.pip_calls.append(command)
                return subprocess.CompletedProcess(command, 0)
            return RUN(command, *args, **kwargs)

        def dependency_versions(_python):
            if isinstance(versions, Exception):
                raise versions
            return versions

        output = io.BytesIO()
        errors = io.BytesIO()
        stdout = io.TextIOWrapper(output, encoding="utf-8")
        stderr = io.TextIOWrapper(errors, encoding="utf-8")
        # Create a real lightweight venv without downloading dependencies. Its
        # Python runs the installed fixture through Git's actual shell alias.
        with mock.patch.dict(os.environ, env or self.env, clear=True), \
                mock.patch.object(sys, "argv", arguments), \
                mock.patch.object(sys, "stdout", stdout), \
                mock.patch.object(sys, "stderr", stderr), \
                mock.patch.object(self.setup.venv, "EnvBuilder",
                                  side_effect=lambda **_kwargs: ENV_BUILDER(with_pip=False)), \
                mock.patch.object(self.setup.subprocess, "run", side_effect=run), \
                mock.patch.object(self.setup, "dependency_versions", side_effect=dependency_versions):
            result = self.setup.main()
        stdout.flush()
        stderr.flush()
        return result, output.getvalue().decode("utf-8"), errors.getvalue().decode("utf-8")

    def installed_helper(self):
        requirements = (self.checkout / "requirements-excel-diff.txt").read_bytes()
        modules = tuple((self.checkout / "scripts" / name).read_bytes()
                        for name in self.setup.STAGE_MODULES)
        return self.setup.stage_path(self.install, HELPER_SOURCE, modules, requirements, VERSIONS)

    def installed_converter(self):
        source = (self.checkout / "scripts" / "excel-textconv.py").read_bytes()
        requirements = (self.checkout / "requirements-excel-diff.txt").read_bytes()
        modules = tuple((self.checkout / "scripts" / name).read_bytes()
                        for name in self.setup.CONVERTER_MODULES)
        return self.setup.converter_path(self.install, source, requirements, VERSIONS, modules=modules)

    def snapshot(self):
        paths = [self.config, self.attributes]
        if self.install.exists():
            paths += [path for path in self.install.rglob("*") if path.is_file()]
        return {str(path): (path.stat().st_mtime_ns, hashlib.sha256(path.read_bytes()).hexdigest())
                for path in paths if path.exists()}

    def test_install_and_repeat_preserve_attributes_and_alias_arguments(self):
        for _ in range(2):
            result, _, error = self.run_setup()
            self.assertEqual(result, 0, error)
        helper = self.installed_helper()
        self.assertEqual(helper.read_bytes(), HELPER_SOURCE)
        self.assertEqual(helper.with_name("excel_stage_ooxml.py").read_bytes(), MODULE_SOURCE)
        self.assertEqual(self.attributes.read_bytes(),
                         self.original_attributes + b"\n" + self.setup.ATTRIBUTE_BLOCK)
        converter = self.installed_converter()
        library = self.setup.converter_library(converter)
        self.assertEqual((library / "word_ooxml.py").read_bytes(), WORD_MODULE_SOURCE)
        self.assertEqual(helper.with_name("word_ooxml.py").read_bytes(), WORD_MODULE_SOURCE)
        self.assertEqual(helper.with_name("word_stage_ooxml.py").read_bytes(), WORD_STAGE_MODULE_SOURCE)
        self.assertEqual(len(list(self.install.glob("excel-stage-*"))), 1)
        for family, suffix, source_suffix in (("excel", "xlsx", "xlsm"), ("word", "docx", "docm")):
            with self.subTest(family=family):
                alias = self.run_git("config", "--global", "--get-all", "alias." + family + "-stage")
                merge_alias = self.run_git("config", "--global", "--get-all", "alias." + family + "-merge")
                self.assertEqual(len(alias.stdout.splitlines()), 1)
                self.assertEqual(len(merge_alias.stdout.splitlines()), 1)
                filename = "-売上 表 O'Brien $dollar." + suffix
                result = self.run_git(family + "-stage", "--", filename)
                self.assertEqual(json.loads(result.stdout),
                                 {"arguments": ["--", filename], "module": "installed module"})
                source = "元の 表 O'Brien $dollar." + source_suffix
                merged = self.run_git(family + "-merge", "--", filename, source)
                self.assertEqual(json.loads(merged.stdout),
                                 {"arguments": ["--merge", "--", filename, source],
                                  "module": "installed module"})
                self.assertEqual(self.run_git("config", "--global", "--get",
                                              "diff." + family + ".binary").stdout.strip(), "true")
                self.assertEqual(self.run_git("config", "--global", "--get",
                                              "diff." + family + ".cachetextconv").stdout.strip(), "false")
        for ending in ("stage", "merge"):
            self.assertEqual(self.run_git("config", "--global", "--get", "alias.word-" + ending).stdout,
                             self.run_git("config", "--global", "--get", "alias.excel-" + ending).stdout)
        self.assertEqual(self.run_git("config", "--global", "--get", "diff.word.textconv").stdout,
                         self.run_git("config", "--global", "--get", "diff.excel.textconv").stdout)

    def test_check_is_read_only_and_discovers_registered_directory(self):
        self.assertEqual(self.run_setup()[0], 0)
        before = self.snapshot()
        changed_env = dict(self.env, XDG_DATA_HOME=str(self.directory / "different-data"))
        result, _, error = self.run_setup(check=True, explicit_directory=False, env=changed_env)
        self.assertEqual(result, 0, error)
        self.assertEqual(self.snapshot(), before)
        self.assertEqual(len(self.pip_calls), 1)
        self.assertFalse((self.directory / "different-data").exists())

    def test_check_rejects_changed_helper_without_repairing_it(self):
        self.assertEqual(self.run_setup()[0], 0)
        helper = self.installed_helper()
        helper.write_bytes(b"# changed entry point\n")
        before = self.snapshot()
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("部分ステージ", error)
        self.assertEqual(self.snapshot(), before)

    def test_check_rejects_missing_module_without_repairing_it(self):
        self.assertEqual(self.run_setup()[0], 0)
        self.installed_helper().with_name("excel_stage_ooxml.py").unlink()
        before = self.snapshot()
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("部分ステージ", error)
        self.assertEqual(self.snapshot(), before)

    def test_module_update_installs_a_new_complete_build(self):
        self.assertEqual(self.run_setup()[0], 0)
        previous = self.installed_helper()
        (self.checkout / "scripts" / "excel_stage_ooxml.py").write_bytes(b"MARKER = 'updated module'\n")
        result, _, error = self.run_setup()
        self.assertEqual(result, 0, error)
        current = self.installed_helper()
        self.assertNotEqual(previous.parent, current.parent)
        self.assertEqual(previous.with_name("excel_stage_ooxml.py").read_bytes(), MODULE_SOURCE)
        self.assertEqual(current.read_bytes(), HELPER_SOURCE)
        self.assertEqual(json.loads(self.run_git("excel-stage", "--", "file.xlsx").stdout)["module"],
                         "updated module")

    def test_check_rejects_changed_alias_without_modifying_config(self):
        self.assertEqual(self.run_setup()[0], 0)
        self.run_git("config", "--global", "alias.excel-stage", "!printf changed")
        before = self.snapshot()
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("alias.excel-stage", error)
        self.assertEqual(self.snapshot(), before)

    def test_check_without_install_or_xml_dependencies_makes_no_changes(self):
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("仮想環境", error)
        self.assertFalse(self.install.exists())
        self.assertFalse(self.config.exists())
        self.assertEqual(self.run_setup()[0], 0)
        before = self.snapshot()
        result, _, error = self.run_setup(check=True, versions=RuntimeError("lxml is unavailable"))
        self.assertEqual(result, 1)
        self.assertIn("lxml", error)
        self.assertEqual(self.snapshot(), before)

    def test_check_rejects_changed_or_missing_merge_alias_without_modifying_config(self):
        for value in ("!printf changed", None):
            with self.subTest(value=value):
                self.assertEqual(self.run_setup()[0], 0)
                if value is None:
                    self.run_git("config", "--global", "--unset-all", "alias.excel-merge")
                else:
                    self.run_git("config", "--global", "alias.excel-merge", value)
                before = self.snapshot()
                result, _, error = self.run_setup(check=True)
                self.assertEqual(result, 1)
                self.assertIn("alias.excel-merge", error)
                self.assertEqual(self.snapshot(), before)

    def test_default_install_directory_is_named_excel_diff(self):
        with mock.patch.dict(os.environ, self.env, clear=True):
            for platform in ("win32", "linux", "darwin"):
                with self.subTest(platform=platform), mock.patch.object(self.setup.sys, "platform", platform):
                    self.assertEqual(self.setup.default_install_directory().name, "excel-diff")

    def test_uninstall_restores_attributes_and_settings_and_can_repeat(self):
        original = b"\xef\xbb\xbf# Existing attributes\r\n*.csv diff=csv\n*.txt text\r\n"
        self.attributes.write_bytes(original)
        self.run_git("config", "--global", "core.attributesFile", self.attributes.as_posix())
        self.run_git("config", "--global", "alias.keep", "status")
        self.assertEqual(self.run_setup()[0], 0)
        suffix = b"# User rules after the installed block\r\n*.custom diff=custom\r\n"
        with self.attributes.open("ab") as output:
            output.write(suffix)
        pip_count = len(self.pip_calls)
        changed_env = dict(self.env, XDG_DATA_HOME=str(self.directory / "elsewhere"))
        result, output, error = self.run_setup(
            uninstall=True, explicit_directory=False, env=changed_env,
            versions=RuntimeError("dependencies must not be checked"),
        )
        self.assertEqual(result, 0, error)
        self.assertIn("削除した導入先", output)
        for key in self.setup.UNINSTALL_KEYS:
            query = RUN(["git", "config", "--global", "--get-all", key],
                        cwd=self.checkout, env=self.env, capture_output=True)
            self.assertEqual(query.returncode, 1, key)
            self.assertIn(key, output)
        self.assertEqual(self.attributes.read_bytes(), original + suffix)
        self.assertFalse(self.install.exists())
        self.assertEqual(self.run_git("config", "--global", "--get", "core.attributesFile").stdout.strip(),
                         self.attributes.as_posix())
        self.assertEqual(self.run_git("config", "--global", "--get", "alias.keep").stdout.strip(), "status")
        self.assertEqual(len(self.pip_calls), pip_count)
        before = self.snapshot()
        result, output, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertIn("消すものがありません", output)
        self.assertEqual(self.snapshot(), before)

    def test_uninstall_leaves_attributes_unchanged_without_markers(self):
        self.assertEqual(self.run_setup()[0], 0)
        original = b"\xef\xbb\xbf# unrelated\r\n*.csv diff=csv"
        self.attributes.write_bytes(original)
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertEqual(self.attributes.read_bytes(), original)
        self.assertFalse(self.install.exists())

    def test_uninstall_preserves_bom_when_managed_block_is_first(self):
        self.attributes.write_bytes(b"")
        self.assertEqual(self.run_setup()[0], 0)
        self.attributes.write_bytes(b"\xef\xbb\xbf" + self.attributes.read_bytes() + b"*.csv diff=csv\r\n")
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertEqual(self.attributes.read_bytes(), b"\xef\xbb\xbf*.csv diff=csv\r\n")

    def test_uninstall_rejects_broken_markers_before_any_change(self):
        self.assertEqual(self.run_setup()[0], 0)
        begin = self.setup.ATTRIBUTE_BLOCK.splitlines(keepends=True)[0]
        end = self.setup.ATTRIBUTE_BLOCK.splitlines(keepends=True)[-1]
        cases = {
            "begin_only": begin + b"*.csv diff=csv\r\n",
            "end_only": end,
            "reversed": end + begin,
            "multiple_blocks": self.setup.ATTRIBUTE_BLOCK * 2,
            "multiple_begins": begin + self.setup.ATTRIBUTE_BLOCK,
            "multiple_ends": self.setup.ATTRIBUTE_BLOCK + end,
        }
        for name, content in cases.items():
            with self.subTest(name=name):
                self.attributes.write_bytes(b"\xef\xbb\xbf" + content)
                before = self.snapshot()
                result, _, error = self.run_setup(uninstall=True)
                self.assertEqual(result, 1)
                self.assertIn("目印", error)
                self.assertEqual(self.snapshot(), before)

    def test_uninstall_refuses_unknown_directory_without_changing_anything(self):
        self.assertEqual(self.run_setup()[0], 0)
        (self.install / "unrelated.txt").write_bytes(b"keep this file")
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("作った形", error)
        self.assertEqual(self.snapshot(), before)

    def test_uninstall_refuses_directory_without_installed_scripts(self):
        self.assertEqual(self.run_setup()[0], 0)
        for entry in self.install.glob("excel-textconv-*.py"):
            entry.unlink()
        for entry in self.install.glob("excel-stage-*"):
            shutil.rmtree(entry)
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("作った形", error)
        self.assertEqual(self.snapshot(), before)

    def test_uninstall_refuses_directory_inside_any_repository(self):
        self.assertEqual(self.run_setup()[0], 0)
        RUN(["git", "init", "-q", str(self.directory)], env=self.env, check=True)
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("リポジトリ", error)
        self.assertEqual(self.snapshot(), before)

    def test_uninstall_refuses_linked_worktree_and_separate_git_directory(self):
        self.assertEqual(self.run_setup()[0], 0)
        repository = self.directory / "repository"
        RUN(["git", "init", "-q", str(repository)], env=self.env, check=True)
        RUN(["git", "-C", str(repository), "-c", "user.name=Test", "-c", "user.email=test@example.invalid",
             "commit", "--allow-empty", "-q", "-m", "fixture"], env=self.env, check=True)
        linked = self.directory / "linked-worktree"
        RUN(["git", "-C", str(repository), "worktree", "add", "-q", "-b", "linked-test", str(linked)],
            env=self.env, check=True)
        separate = self.directory / "separate-worktree"
        metadata = self.directory / "separate-metadata"
        RUN(["git", "init", "-q", "--separate-git-dir", str(metadata), str(separate)],
            env=self.env, check=True)
        for location in (linked, separate):
            with self.subTest(location=location.name):
                previous = self.install
                self.install = previous.rename(location / "installed")
                self.assertTrue((location / ".git").is_file())
                before = self.snapshot()
                result, _, error = self.run_setup(uninstall=True)
                self.assertEqual(result, 1)
                self.assertIn("リポジトリ", error)
                self.assertEqual(self.snapshot(), before)
                self.install = self.install.rename(previous)

    def test_uninstall_refuses_worktree_defined_only_by_git_environment(self):
        self.assertEqual(self.run_setup()[0], 0)
        metadata = self.directory / "custom-metadata"
        RUN(["git", "init", "-q", "--bare", str(metadata)], env=self.env, check=True)
        location = self.directory / "custom-worktree"
        location.mkdir()
        self.install = self.install.rename(location / "installed")
        self.assertFalse((location / ".git").exists())
        RUN(["git", "--git-dir", str(metadata), "config", "core.bare", "false"], env=self.env, check=True)
        RUN(["git", "--git-dir", str(metadata), "config", "core.worktree", str(location)],
            env=self.env, check=True)
        for explicit_worktree in (False, True):
            with self.subTest(explicit_worktree=explicit_worktree):
                environment = dict(self.env, GIT_DIR=str(metadata))
                if explicit_worktree:
                    environment["GIT_WORK_TREE"] = str(location)
                before = self.snapshot()
                result, _, error = self.run_setup(uninstall=True, env=environment)
                self.assertEqual(result, 1)
                self.assertIn("リポジトリ", error)
                self.assertEqual(self.snapshot(), before)

    def test_uninstall_refuses_missing_venv_without_changing_anything(self):
        self.assertEqual(self.run_setup()[0], 0)
        shutil.rmtree(self.install / "venv")
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("作った形", error)
        self.assertEqual(self.snapshot(), before)

    @unittest.skipUnless(os.name == "nt", "Windows junctions only")
    def test_uninstall_refuses_junction_before_resolving_its_target(self):
        self.assertEqual(self.run_setup()[0], 0)
        target = self.install.rename(self.directory / "real-install")
        created = RUN(["cmd", "/c", "mklink", "/J", str(self.install), str(target)],
                      env=self.env, capture_output=True)
        if created.returncode != 0:
            self.skipTest("Directory junction creation is unavailable")
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("リンク", error)
        self.assertEqual(self.snapshot(), before)
        self.assertTrue(target.is_dir())

    def test_uninstall_refuses_included_settings_before_any_change(self):
        self.assertEqual(self.run_setup()[0], 0)
        included = self.directory / "included.gitconfig"
        included.write_bytes(b'[alias]\n\texcel-merge = "!keep included command"\n')
        self.run_git("config", "--global", "include.path", included.as_posix())
        before = self.snapshot()
        included_before = (included.stat().st_mtime_ns, included.read_bytes())
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("include", error)
        self.assertEqual(self.snapshot(), before)
        self.assertEqual((included.stat().st_mtime_ns, included.read_bytes()), included_before)

    def test_uninstall_stops_before_any_change_on_preflight_permission_error(self):
        self.assertEqual(self.run_setup()[0], 0)
        original_exists = Path.exists
        original_read = Path.read_bytes
        original_lstat = Path.lstat
        for location in (self.attributes, self.install):
            with self.subTest(location=location.name):
                before = self.snapshot()

                def exists(path):
                    return False if path == location else original_exists(path)

                def read_bytes(path):
                    if path == location:
                        raise PermissionError("preflight access denied")
                    return original_read(path)

                def lstat(path):
                    if path == location:
                        raise PermissionError("preflight access denied")
                    return original_lstat(path)

                with mock.patch.object(Path, "exists", exists), \
                        mock.patch.object(Path, "read_bytes", read_bytes), \
                        mock.patch.object(Path, "lstat", lstat):
                    if location == self.install:
                        with self.assertRaises(PermissionError):
                            self.setup.validate_uninstall_directory(self.install, self.checkout)
                    result, _, error = self.run_setup(uninstall=True)
                self.assertEqual(result, 1)
                self.assertIn("preflight access denied", error)
                self.assertEqual(self.snapshot(), before)

    def test_upgrade_excel_only_attribute_block_preserves_other_bytes(self):
        bom = b"\xef\xbb\xbf"
        crlf_prefix = bom + b"# Existing attributes\r\n*.csv diff=csv\r\n"
        suffix = b"# User rules\r\n*.docx diff=custom"
        cases = {
            "at_end": (b"*.csv diff=csv\n" + OLD_ATTRIBUTE_BLOCK, b"*.csv diff=csv\n"),
            "first_with_bom": (bom + OLD_ATTRIBUTE_BLOCK + b"*.csv diff=csv\r\n",
                               bom + b"*.csv diff=csv\r\n"),
            "after_crlf_content": (crlf_prefix + OLD_ATTRIBUTE_BLOCK.replace(b"\n", b"\r\n"),
                                   crlf_prefix),
            "with_user_rules_after": (b"*.csv diff=csv\n" + OLD_ATTRIBUTE_BLOCK + suffix,
                                      b"*.csv diff=csv\n" + suffix + b"\n"),
        }
        for name, (existing, remaining) in cases.items():
            with self.subTest(name=name):
                self.attributes.write_bytes(existing)
                result, _, error = self.run_setup()
                self.assertEqual(result, 0, error)
                self.assertEqual(self.attributes.read_bytes(), remaining + self.setup.ATTRIBUTE_BLOCK)
                before = (self.attributes.stat().st_mtime_ns, self.attributes.read_bytes())
                result, _, error = self.run_setup()
                self.assertEqual(result, 0, error)
                self.assertEqual((self.attributes.stat().st_mtime_ns, self.attributes.read_bytes()), before)

    def test_install_moves_current_attribute_block_after_user_rules(self):
        suffix = b"*.docx diff=custom\r\n*.txt text"
        self.attributes.write_bytes(b"\xef\xbb\xbf" + self.setup.ATTRIBUTE_BLOCK + suffix)
        result, _, error = self.run_setup()
        self.assertEqual(result, 0, error)
        self.assertEqual(self.attributes.read_bytes(),
                         b"\xef\xbb\xbf" + suffix + b"\n" + self.setup.ATTRIBUTE_BLOCK)

    def test_install_rejects_broken_markers_before_any_change(self):
        begin = OLD_ATTRIBUTE_BLOCK.splitlines(keepends=True)[0]
        end = OLD_ATTRIBUTE_BLOCK.splitlines(keepends=True)[-1]
        cases = {
            "begin_only": begin + b"*.csv diff=csv\r\n",
            "end_only": end,
            "reversed": end + begin,
            "multiple_blocks": OLD_ATTRIBUTE_BLOCK * 2,
            "multiple_begins": begin + OLD_ATTRIBUTE_BLOCK,
            "multiple_ends": OLD_ATTRIBUTE_BLOCK + end,
        }
        for name, content in cases.items():
            with self.subTest(name=name):
                self.attributes.write_bytes(b"\xef\xbb\xbf" + content)
                before = self.snapshot()
                result, _, error = self.run_setup()
                self.assertEqual(result, 1)
                self.assertIn("目印", error)
                self.assertEqual(self.snapshot(), before)
                self.assertFalse(self.install.exists())
                self.assertFalse(self.config.exists())
                self.assertEqual(self.pip_calls, [])

    def test_check_rejects_old_attribute_block_without_upgrading_it(self):
        self.assertEqual(self.run_setup()[0], 0)
        self.attributes.write_bytes(b"*.csv diff=csv\r\n" + OLD_ATTRIBUTE_BLOCK)
        before = self.snapshot()
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("attributes", error)
        self.assertEqual(self.snapshot(), before)

    def test_check_rejects_duplicate_markers_even_when_new_block_is_at_end(self):
        self.assertEqual(self.run_setup()[0], 0)
        self.attributes.write_bytes(OLD_ATTRIBUTE_BLOCK + self.setup.ATTRIBUTE_BLOCK)
        before = self.snapshot()
        result, _, error = self.run_setup(check=True)
        self.assertEqual(result, 1)
        self.assertIn("目印", error)
        self.assertEqual(self.snapshot(), before)

    def test_check_rejects_missing_or_changed_word_modules_without_repairing_them(self):
        self.assertEqual(self.run_setup()[0], 0)
        helper = self.installed_helper()
        converter = self.installed_converter()
        modules = (
            (self.setup.converter_library(converter) / "word_ooxml.py", WORD_MODULE_SOURCE, "変換"),
            (helper.with_name("word_ooxml.py"), WORD_MODULE_SOURCE, "部分ステージ"),
            (helper.with_name("word_stage_ooxml.py"), WORD_STAGE_MODULE_SOURCE, "部分ステージ"),
        )
        for module, original, message in modules:
            for changed in (None, b"# modified module\n"):
                with self.subTest(module=str(module), changed=changed):
                    if changed is None:
                        module.unlink()
                    else:
                        module.write_bytes(changed)
                    before = self.snapshot()
                    result, _, error = self.run_setup(check=True)
                    self.assertEqual(result, 1)
                    self.assertIn(message, error)
                    self.assertEqual(self.snapshot(), before)
                    module.write_bytes(original)

    def test_word_module_updates_install_complete_immutable_builds(self):
        self.assertEqual(self.run_setup()[0], 0)
        for name, replacement in (
            ("word_ooxml.py", b"MARKER = 'updated Word module'\n"),
            ("word_stage_ooxml.py", b"MARKER = 'updated Word stage module'\n"),
        ):
            with self.subTest(module=name):
                previous_helper = self.installed_helper()
                previous_converter = self.installed_converter()
                previous_library = self.setup.converter_library(previous_converter)
                old_module = previous_helper.with_name(name).read_bytes()
                (self.checkout / "scripts" / name).write_bytes(replacement)
                result, _, error = self.run_setup()
                self.assertEqual(result, 0, error)
                current_helper = self.installed_helper()
                current_converter = self.installed_converter()
                self.assertNotEqual(current_helper.parent, previous_helper.parent)
                self.assertEqual(previous_helper.with_name(name).read_bytes(), old_module)
                self.assertEqual(current_helper.with_name(name).read_bytes(), replacement)
                self.assertEqual(current_helper.read_bytes(), HELPER_SOURCE)
                for module_name in self.setup.STAGE_MODULES:
                    self.assertEqual(current_helper.with_name(module_name).read_bytes(),
                                     (self.checkout / "scripts" / module_name).read_bytes())
                if name == "word_ooxml.py":
                    self.assertNotEqual(current_converter, previous_converter)
                    self.assertEqual((previous_library / name).read_bytes(), WORD_MODULE_SOURCE)
                    self.assertEqual((self.setup.converter_library(current_converter) / name).read_bytes(),
                                     replacement)
                else:
                    self.assertEqual(current_converter, previous_converter)

    def test_install_writes_imported_modules_before_entry_points(self):
        writes = []
        original_write = Path.write_bytes

        def write_bytes(path, content):
            if path.name in ("word_ooxml.py", "word_stage_ooxml.py", "excel_stage_ooxml.py",
                              "excel-stage.py") or path.name.startswith("excel-textconv-"):
                writes.append(path)
            return original_write(path, content)

        with mock.patch.object(Path, "write_bytes", write_bytes):
            result, _, error = self.run_setup()
        self.assertEqual(result, 0, error)
        converter = self.installed_converter()
        helper = self.installed_helper()
        self.assertEqual(writes, [
            self.setup.converter_library(converter) / "word_ooxml.py",
            converter,
            helper.with_name("excel_stage_ooxml.py"),
            helper.with_name("word_ooxml.py"),
            helper.with_name("word_stage_ooxml.py"),
            helper,
        ])

    def test_check_rejects_changed_or_missing_word_settings_without_modifying_config(self):
        self.assertEqual(self.run_setup()[0], 0)
        keys = ("diff.word.textconv", "diff.word.binary", "diff.word.cachetextconv",
                "alias.word-stage", "alias.word-merge")
        for key in keys:
            original = self.run_git("config", "--global", "--get", key).stdout.rstrip("\n")
            for value in ("!printf changed", None):
                with self.subTest(key=key, value=value):
                    if value is None:
                        self.run_git("config", "--global", "--unset-all", key)
                    else:
                        self.run_git("config", "--global", "--replace-all", key, value)
                    before = self.snapshot()
                    result, _, error = self.run_setup(check=True)
                    self.assertEqual(result, 1)
                    self.assertIn(key, error)
                    self.assertEqual(self.snapshot(), before)
                    self.run_git("config", "--global", "--replace-all", key, original)

    def test_check_and_uninstall_discover_directory_through_word_textconv_alone(self):
        self.assertEqual(self.run_setup()[0], 0)
        self.run_git("config", "--global", "--unset-all", "diff.excel.textconv")
        before = self.snapshot()
        changed_env = dict(self.env, XDG_DATA_HOME=str(self.directory / "different-data"),
                           LOCALAPPDATA=str(self.directory / "different-local"))
        result, _, error = self.run_setup(check=True, explicit_directory=False, env=changed_env)
        self.assertEqual(result, 1)
        # Discovery succeeds; --check still requires the missing Excel key.
        self.assertIn("diff.excel.textconv", error)
        self.assertEqual(self.snapshot(), before)
        result, _, error = self.run_setup(uninstall=True, explicit_directory=False, env=changed_env)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        self.assertFalse((self.directory / "different-data").exists())
        self.assertFalse((self.directory / "different-local").exists())

    def test_install_refuses_foreign_word_commands_before_any_change(self):
        for key in ("diff.word.textconv", "alias.word-stage", "alias.word-merge"):
            for value in ("!printf foreign", "'unterminated"):
                with self.subTest(key=key, value=value):
                    self.run_git("config", "--global", "--replace-all", key, value)
                    before = self.snapshot()
                    result, _, error = self.run_setup()
                    self.assertEqual(result, 1)
                    self.assertIn(key, error)
                    self.assertIn("別の設定", error)
                    self.assertEqual(self.snapshot(), before)
                    self.assertFalse(self.install.exists())
                    self.assertEqual(self.pip_calls, [])
                    self.run_git("config", "--global", "--unset-all", key)

    def test_install_refuses_mixed_owned_and_foreign_word_commands(self):
        self.assertEqual(self.run_setup()[0], 0)
        for key in ("diff.word.textconv", "alias.word-stage", "alias.word-merge"):
            with self.subTest(key=key):
                self.run_git("config", "--global", "--add", key, "!printf foreign")
                before = self.snapshot()
                pip_count = len(self.pip_calls)
                result, _, error = self.run_setup()
                self.assertEqual(result, 1)
                self.assertIn(key, error)
                self.assertEqual(self.snapshot(), before)
                self.assertEqual(len(self.pip_calls), pip_count)
                self.run_git("config", "--global", "--unset-all", key, "foreign")

    def test_uninstall_accepts_new_layout_with_module_bytecode(self):
        self.assertEqual(self.run_setup()[0], 0)
        for directory in (self.installed_helper().parent,
                          self.setup.converter_library(self.installed_converter())):
            bytecode = directory / "__pycache__"
            bytecode.mkdir()
            (bytecode / "word_ooxml.cpython-314.pyc").write_bytes(b"fixture bytecode")
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())

    def test_uninstall_accepts_hand_built_excel_only_layout(self):
        python = self.install / "venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        python.parent.mkdir(parents=True)
        source = (self.checkout / "scripts" / "excel-textconv.py").read_bytes()
        requirements = (self.checkout / "requirements-excel-diff.txt").read_bytes()
        converter_hash = hashlib.sha256(
            b"\0".join((source, requirements, VERSIONS.encode("utf-8")))
        ).hexdigest()[:16]
        converter = self.install / ("excel-textconv-" + converter_hash + ".py")
        self.assertEqual(self.setup.converter_path(self.install, source, requirements, VERSIONS), converter)
        old_hash = hashlib.sha256(
            b"\0".join((HELPER_SOURCE, MODULE_SOURCE, requirements, VERSIONS.encode("utf-8")))
        ).hexdigest()[:16]
        helper = self.install / ("excel-stage-" + old_hash) / "excel-stage.py"
        helper.parent.mkdir()
        converter.write_bytes(source)
        helper.write_bytes(HELPER_SOURCE)
        helper.with_name("excel_stage_ooxml.py").write_bytes(MODULE_SOURCE)
        self.assertFalse(self.setup.converter_library(converter).exists())
        self.attributes.write_bytes(b"*.csv diff=csv\r\n" + OLD_ATTRIBUTE_BLOCK)
        expected = {
            "diff.excel.textconv": self.setup.textconv_command(python, converter),
            "diff.excel.binary": "true",
            "diff.excel.cachetextconv": "false",
            "alias.excel-stage": self.setup.stage_alias(python, helper),
            "alias.excel-merge": self.setup.merge_alias(python, helper),
        }
        for key, value in expected.items():
            self.run_git("config", "--global", key, value)
        result, _, error = self.run_setup(uninstall=True, explicit_directory=False)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        self.assertEqual(self.attributes.read_bytes(), b"*.csv diff=csv\r\n")
        self.assertEqual(self.pip_calls, [])
        for key in expected:
            query = RUN(["git", "config", "--global", "--get-all", key],
                        cwd=self.checkout, env=self.env, capture_output=True)
            self.assertEqual(query.returncode, 1, key)

    def test_uninstall_refuses_converter_library_without_required_module(self):
        self.assertEqual(self.run_setup()[0], 0)
        library = self.setup.converter_library(self.installed_converter())
        (library / "word_ooxml.py").unlink()
        before = self.snapshot()
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 1)
        self.assertIn("作った形", error)
        self.assertEqual(self.snapshot(), before)
        self.assertTrue(library.is_dir())

    def test_uninstall_preserves_foreign_word_commands_and_their_driver_settings(self):
        self.assertEqual(self.run_setup()[0], 0)
        foreign = {
            "diff.word.textconv": "!printf foreign",
            "diff.word.binary": "false",
            "diff.word.cachetextconv": "true",
            "alias.word-stage": "!printf foreign-stage",
            "alias.word-merge": "'unterminated",
        }
        for key, value in foreign.items():
            self.run_git("config", "--global", "--replace-all", key, value)
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        for key, value in foreign.items():
            self.assertEqual(self.run_git("config", "--global", "--get-all", key).stdout.rstrip("\n"),
                             value)

    def test_uninstall_preserves_foreign_word_aliases_with_owned_word_driver(self):
        self.assertEqual(self.run_setup()[0], 0)
        for key in ("alias.word-stage", "alias.word-merge"):
            self.run_git("config", "--global", "--replace-all", key, "!printf foreign")
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        for key in ("alias.word-stage", "alias.word-merge"):
            self.assertEqual(self.run_git("config", "--global", "--get-all", key).stdout.strip(),
                             "!printf foreign")
        for key in ("diff.word.textconv", "diff.word.binary", "diff.word.cachetextconv"):
            query = RUN(["git", "config", "--global", "--get-all", key],
                        cwd=self.checkout, env=self.env, capture_output=True)
            self.assertEqual(query.returncode, 1, key)

    def test_uninstall_preserves_foreign_included_word_settings(self):
        self.assertEqual(self.run_setup()[0], 0)
        keys = ("diff.word.textconv", "diff.word.binary", "diff.word.cachetextconv",
                "alias.word-stage", "alias.word-merge")
        for key in keys:
            self.run_git("config", "--global", "--unset-all", key)
        included = self.directory / "included-word.gitconfig"
        included.write_bytes(
            b'[diff "word"]\n\ttextconv = "!printf foreign"\n\tbinary = false\n\tcachetextconv = true\n'
            b'[alias]\n\tword-stage = "!printf foreign-stage"\n\tword-merge = "!printf foreign-merge"\n'
        )
        self.run_git("config", "--global", "include.path", included.as_posix())
        before = (included.stat().st_mtime_ns, included.read_bytes())
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        self.assertEqual((included.stat().st_mtime_ns, included.read_bytes()), before)
        for key in keys:
            self.assertTrue(self.run_git("config", "--global", "--includes", "--get-all", key).stdout)

    def test_uninstall_preserves_mixed_owned_and_foreign_word_commands(self):
        self.assertEqual(self.run_setup()[0], 0)
        keys = ("diff.word.textconv", "alias.word-stage", "alias.word-merge")
        for key in keys:
            self.run_git("config", "--global", "--add", key, "!printf foreign")
        preserved_keys = keys + ("diff.word.binary", "diff.word.cachetextconv")
        originals = {key: self.run_git("config", "--global", "--get-all", key).stdout
                     for key in preserved_keys}
        result, _, error = self.run_setup(uninstall=True)
        self.assertEqual(result, 0, error)
        self.assertFalse(self.install.exists())
        for key, original in originals.items():
            self.assertEqual(self.run_git("config", "--global", "--get-all", key).stdout, original)


if __name__ == "__main__":
    unittest.main()
