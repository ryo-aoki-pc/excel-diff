#!/usr/bin/env python3
"""Excel の Git textconv をユーザー全体に設定する (Python 3.9 以降)。"""

import argparse
import hashlib
import os
import shlex
import shutil
import subprocess
import sys
import venv
from pathlib import Path


ATTRIBUTE_BLOCK = (
    "# BEGIN lazygit Excel textconv (setup-excel-diff.py)\n"
    "*.[xX][lL][sS][xX] diff=excel\n"
    "*.[xX][lL][sS][mM] diff=excel\n"
    "*.[xX][lL][tT][xX] diff=excel\n"
    "*.[xX][lL][tT][mM] diff=excel\n"
    "*.[xX][lL][sS] diff=excel\n"
    "*.[xX][lL][tT] diff=excel\n"
    "# END lazygit Excel textconv (setup-excel-diff.py)\n"
).encode("ascii")


def xdg_directory(variable, fallback):
    value = os.environ.get(variable)
    return Path(value) if value and Path(value).is_absolute() else fallback


def default_install_directory():
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData/Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library/Application Support"
    else:
        base = xdg_directory("XDG_DATA_HOME", Path.home() / ".local/share")
    return base / "lazygit-excel-diff"


def git_values(key, path=False, scope="global"):
    command = ["git", "config", "--" + scope, "--includes", "--null"]
    if path:
        command.append("--path")
    result = subprocess.run(
        command + ["--get-all", key], capture_output=True, text=True, encoding="utf-8"
    )
    if result.returncode == 1:
        return []
    if result.returncode != 0:
        raise RuntimeError("Git 設定を読み取れません: " + result.stderr.strip())
    return result.stdout.split("\0")[:-1]


def attributes_file():
    values = git_values("core.attributesFile", path=True)
    if not values:
        # Respect an existing system attributes path without letting a local
        # repository override a user-wide installation.
        values = git_values("core.attributesFile", path=True, scope="system")
    if len(values) > 1:
        raise RuntimeError(
            "core.attributesFile が複数あります。グローバル設定を 1 つに整理してください。"
        )
    if values:
        path = Path(values[0])
        if not values[0] or not path.is_absolute():
            raise RuntimeError(
                "core.attributesFile は空でない絶対パスにしてください "
                "(既存の設定は変更していません)。"
            )
        return path
    return xdg_directory("XDG_CONFIG_HOME", Path.home() / ".config") / "git/attributes"


def dependency_versions(python):
    result = subprocess.run(
        [str(python), "-c", "import openpyxl, xlrd; print(openpyxl.__version__); print(xlrd.__version__)"],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError("Excel 変換用の依存関係を確認できません: " + result.stderr.strip())
    return result.stdout.strip()


def converter_path(install_directory, source, requirements, versions):
    # Git invalidates its textconv cache when the command changes. Include the
    # converter and library versions so reinstalls cannot reuse stale text.
    digest = hashlib.sha256(source + b"\0" + requirements + b"\0" + versions.encode("utf-8"))
    return install_directory / ("excel-textconv-" + digest.hexdigest()[:16] + ".py")


def cache_directory(converter):
    # One cache per converter build, so an update never reuses stale text.
    return converter.parent / "cache" / converter.stem


def textconv_command(python, converter):
    # Git for Windows also executes textconv with sh; use POSIX quoting and
    # forward slashes, keeping the venv Python path (not its symlink target).
    parts = [python, converter, "--cache-dir", cache_directory(converter)]
    return " ".join(shlex.quote(part if isinstance(part, str) else part.as_posix()) for part in parts)


def registered_converter(command):
    for part in shlex.split(command):
        name = Path(part).name
        if name.startswith("excel-textconv-") and name.endswith(".py"):
            return Path(part)
    return None


def remove_stale_caches(converter):
    cache_root = converter.parent / "cache"
    if not cache_root.is_dir():
        return
    for directory in cache_root.iterdir():
        if directory.name != converter.stem:
            # A worker of the old build may still hold a file open on Windows.
            shutil.rmtree(directory, ignore_errors=True)


def verify(converter, source, attributes, expected):
    if not converter.is_file() or converter.read_bytes() != source:
        raise RuntimeError("変換スクリプトが未設定か更新されています。セットアップを再実行してください。")
    for key, value in expected.items():
        values = git_values(key)
        if not values or values[-1] != value:
            raise RuntimeError(key + " が設定と一致しません。セットアップを再実行してください。")
    if not attributes.is_file() or not attributes.read_bytes().rstrip(b"\r\n").endswith(
        ATTRIBUTE_BLOCK.rstrip(b"\n")
    ):
        raise RuntimeError("Excel 用の attributes が未設定か変更されています。セットアップを再実行してください。")


def main():
    if sys.version_info < (3, 9):
        print("Python 3.9 以降で実行してください。", file=sys.stderr)
        return 1
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--install-dir", type=Path,
                        help="仮想環境と変換スクリプトの保存先 (同期対象リポジトリの外)")
    parser.add_argument("--check", action="store_true", help="設定を変更せず、インストール状態を確認する")
    args = parser.parse_args()
    try:
        root = Path(__file__).resolve().parent.parent
        selected_directory = args.install_dir
        if args.check and selected_directory is None:
            commands = git_values("diff.excel.textconv")
            registered = registered_converter(commands[-1]) if commands else None
            if registered is not None:
                selected_directory = registered.parent
        directory = (selected_directory or default_install_directory()).expanduser().resolve()
        if directory == root or root in directory.parents:
            raise RuntimeError("--install-dir は同期対象リポジトリの外を指定してください。")
        source_path = root / "scripts/excel-textconv.py"
        source = source_path.read_bytes()
        requirements_path = root / "requirements-excel-diff.txt"
        requirements = requirements_path.read_bytes()
        attributes = attributes_file()
        # Read before installing, so unreadable attributes do not cause a partial setup.
        existing_attributes = attributes.read_bytes() if attributes.exists() else b""
        python = directory / "venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        if args.check:
            if not python.is_file():
                raise RuntimeError("Excel 変換用の仮想環境がありません。セットアップを実行してください。")
        else:
            print("Excel 変換用の仮想環境と依存関係を準備しています。", flush=True)
            venv.EnvBuilder(with_pip=True).create(directory / "venv")
            subprocess.run(
                [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(requirements_path)],
                check=True,
            )
        converter = converter_path(directory, source, requirements, dependency_versions(python))
        expected = {
            "diff.excel.textconv": textconv_command(python, converter),
            "diff.excel.binary": "true",
            # The converter keeps its own cache. Git's cache would store the
            # temporary "converting" text of a large workbook for good.
            "diff.excel.cachetextconv": "false",
        }
        if not git_values("core.attributesFile", path=True):
            # Pin the selected path so the rules also apply when another shell
            # or editor starts Git with a different XDG_CONFIG_HOME.
            expected["core.attributesFile"] = attributes.as_posix()
        if not args.check:
            converter.write_bytes(source)
            remove_stale_caches(converter)
            if not existing_attributes.rstrip(b"\r\n").endswith(ATTRIBUTE_BLOCK.rstrip(b"\n")):
                attributes.parent.mkdir(parents=True, exist_ok=True)
                # Append bytes to preserve unrelated rules, encodings and line endings.
                with attributes.open("ab") as output:
                    if existing_attributes and not existing_attributes.endswith(b"\n"):
                        output.write(b"\n")
                    output.write(ATTRIBUTE_BLOCK)
            for key, value in expected.items():
                subprocess.run(["git", "config", "--global", "--replace-all", key, value], check=True)
        verify(converter, source, attributes, expected)
        print("Excel 差分の設定を確認しました。" if args.check else "Excel 差分をグローバルに設定しました。")
        print("変換スクリプト: " + str(converter))
        print("attributes: " + str(attributes))
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.CalledProcessError, KeyboardInterrupt) as error:
        print("Excel 差分のセットアップに失敗しました: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
