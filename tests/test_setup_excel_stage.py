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

    def run_setup(self, *, check=False, explicit_directory=True, versions=VERSIONS, env=None):
        arguments = ["setup-excel-diff.py"]
        if explicit_directory:
            arguments += ["--install-dir", str(self.install)]
        if check:
            arguments.append("--check")

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
        self.assertEqual(len(list(self.install.glob("excel-stage-*"))), 1)
        filename = "-売上 表 O'Brien $dollar.xlsx"
        result = self.run_git("excel-stage", "--", filename)
        self.assertEqual(json.loads(result.stdout),
                         {"arguments": ["--", filename], "module": "installed module"})
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


if __name__ == "__main__":
    unittest.main()
