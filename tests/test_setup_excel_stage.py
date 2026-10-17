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
    "print(json.dumps({'arguments': sys.argv[1:], 'module': MARKER}, ensure_ascii=False))\n"
).encode("utf-8")
MODULE_SOURCE = b"MARKER = 'installed module'\n"
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
        module = (self.checkout / "scripts" / "excel_stage_ooxml.py").read_bytes()
        return self.setup.stage_path(self.install, HELPER_SOURCE, module, requirements, VERSIONS)

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
        alias = self.run_git("config", "--global", "--get-all", "alias.excel-stage").stdout.splitlines()
        self.assertEqual(len(alias), 1)
        merge_alias = self.run_git("config", "--global", "--get-all", "alias.excel-merge").stdout.splitlines()
        self.assertEqual(len(merge_alias), 1)
        self.assertEqual(len(list(self.install.glob("excel-stage-*"))), 1)
        filename = "-売上 表 O'Brien $dollar.xlsx"
        result = self.run_git("excel-stage", "--", filename)
        self.assertEqual(json.loads(result.stdout),
                         {"arguments": ["--", filename], "module": "installed module"})
        source = "元の 表 O'Brien $dollar.xlsm"
        merged = self.run_git("excel-merge", "--", filename, source)
        self.assertEqual(json.loads(merged.stdout),
                         {"arguments": ["--merge", "--", filename, source],
                          "module": "installed module"})
        self.assertEqual(self.run_git("config", "--global", "--get", "diff.excel.binary").stdout.strip(), "true")
        self.assertEqual(self.run_git("config", "--global", "--get", "diff.excel.cachetextconv").stdout.strip(), "false")

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


if __name__ == "__main__":
    unittest.main()
