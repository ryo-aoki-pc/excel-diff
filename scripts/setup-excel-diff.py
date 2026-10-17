#!/usr/bin/env python3
"""Excel の差分表示と部分ステージをユーザー全体に設定する (Python 3.9 以降)。"""

import argparse
import hashlib
import os
import shlex
import shutil
import stat
import subprocess
import sys
import venv
from pathlib import Path


# Keep these markers unchanged: a reinstall must recognize blocks that
# earlier setups already appended, instead of adding a second one.
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
    return base / "excel-diff"


def git_values(key, path=False, scope="global", includes=True):
    command = ["git", "config", "--" + scope,
               "--includes" if includes else "--no-includes", "--null"]
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
        [str(python), "-B", "-c",
         "import openpyxl, xlrd, lxml.etree; "
         "print(openpyxl.__version__); print(xlrd.__version__); "
         "print(lxml.__version__)"],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError("Excel 用の依存関係を確認できません: " + result.stderr.strip())
    return result.stdout.strip()


def converter_path(install_directory, source, requirements, versions):
    # Git invalidates its textconv cache when the command changes. Include the
    # converter and library versions so reinstalls cannot reuse stale text.
    digest = hashlib.sha256(source + b"\0" + requirements + b"\0" + versions.encode("utf-8"))
    return install_directory / ("excel-textconv-" + digest.hexdigest()[:16] + ".py")


def stage_path(install_directory, source, module, requirements, versions):
    # Keep the entry point and its imported module in one immutable build so a
    # running selector cannot accidentally import a different installed version.
    digest = hashlib.sha256(
        b"\0".join((source, module, requirements, versions.encode("utf-8")))
    )
    return install_directory / ("excel-stage-" + digest.hexdigest()[:16]) / "excel-stage.py"


def cache_directory(converter):
    # One cache per converter build, so an update never reuses stale text.
    return converter.parent / "cache" / converter.stem


def textconv_command(python, converter):
    # Git for Windows also executes textconv with sh; use POSIX quoting and
    # forward slashes, keeping the venv Python path (not its symlink target).
    parts = [python, converter, "--cache-dir", cache_directory(converter)]
    return " ".join(shlex.quote(part if isinstance(part, str) else part.as_posix()) for part in parts)


def helper_alias(python, helper, *options):
    # Git shell aliases append the user's arguments, including the -- path
    # separator. POSIX quoting works with Git for Windows' bundled shell too.
    return "!" + " ".join(
        shlex.quote(value.as_posix() if isinstance(value, Path) else value)
        for value in (python, helper, *options)
    )


def stage_alias(python, helper):
    return helper_alias(python, helper)


def merge_alias(python, helper):
    return helper_alias(python, helper, "--merge")


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


def verify(converter, source, attributes, expected, helper, helper_source, module_source):
    if not converter.is_file() or converter.read_bytes() != source:
        raise RuntimeError("変換スクリプトが未設定か更新されています。セットアップを再実行してください。")
    module = helper.with_name("excel_stage_ooxml.py")
    if (not helper.is_file() or helper.read_bytes() != helper_source
            or not module.is_file() or module.read_bytes() != module_source):
        raise RuntimeError("Excel 部分ステージのスクリプトが未設定か更新されています。セットアップを再実行してください。")
    for key, value in expected.items():
        values = git_values(key)
        if not values or values[-1] != value:
            raise RuntimeError(key + " が設定と一致しません。セットアップを再実行してください。")
    if not attributes.is_file() or not attributes.read_bytes().rstrip(b"\r\n").endswith(
        ATTRIBUTE_BLOCK.rstrip(b"\n")
    ):
        raise RuntimeError("Excel 用の attributes が未設定か変更されています。セットアップを再実行してください。")


UNINSTALL_KEYS = (
    "diff.excel.textconv",
    "diff.excel.binary",
    "diff.excel.cachetextconv",
    "alias.excel-stage",
    "alias.excel-merge",
)


def without_attribute_block(content):
    # Work on byte ranges so unrelated encodings, BOMs and line endings survive.
    begin = b"# BEGIN lazygit Excel textconv (setup-excel-diff.py)"
    end = b"# END lazygit Excel textconv (setup-excel-diff.py)"
    starts = []
    ends = []
    offset = 0
    bom = b"\xef\xbb\xbf"
    for line in content.splitlines(keepends=True):
        marker = line.rstrip(b"\r\n")
        marker_offset = offset
        if offset == 0 and marker.startswith(bom):
            marker = marker[len(bom):]
            marker_offset += len(bom)
        if marker == begin:
            starts.append(marker_offset)
        elif marker == end:
            ends.append(offset + len(line))
        offset += len(line)
    if not starts and not ends:
        return content
    if len(starts) != 1 or len(ends) != 1 or starts[0] >= ends[0]:
        raise RuntimeError(
            "attributes の Excel 用区画の目印が壊れています。設定は変更していません。"
        )
    return content[:starts[0]] + content[ends[0]:]


def linked_directory(path):
    try:
        info = path.lstat()
    except FileNotFoundError:
        return False
    # Windows junctions are not symlinks, but resolve() would follow them before
    # rmtree() sees the requested path. Reject all reparse-point directories.
    return path.is_symlink() or bool(
        getattr(info, "st_file_attributes", 0)
        & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
    )


def validate_uninstall_directory(directory, root):
    try:
        directory.lstat()
    except FileNotFoundError:
        return False
    if directory == root or directory in root.parents or root in directory.parents:
        raise RuntimeError("リポジトリのフォルダーは削除できません: " + str(directory))
    # Check ancestors directly as Git may refuse repositories owned by another
    # account. A worktree's .git file also counts as a repository marker.
    for parent in (directory, *directory.parents):
        if (parent / ".git").exists() or (parent / ".git").is_symlink() or (
            (parent / "HEAD").is_file() and (parent / "objects").is_dir()
            and (parent / "refs").is_dir()
        ):
            raise RuntimeError("リポジトリの中のフォルダーは削除できません: " + str(directory))
    if not directory.is_dir() or linked_directory(directory):
        raise RuntimeError("導入先の形を確認できないため削除しません: " + str(directory))
    # GIT_DIR/GIT_WORK_TREE can define a worktree without a .git marker.
    repository = subprocess.run(
        ["git", "-C", str(directory), "rev-parse", "--is-inside-work-tree"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if repository.returncode == 0 and repository.stdout.strip() == "true":
        raise RuntimeError("リポジトリの中のフォルダーは削除できません: " + str(directory))
    if repository.returncode not in (0, 128) or (
        repository.returncode != 0 and (os.environ.get("GIT_DIR") or os.environ.get("GIT_WORK_TREE"))
    ):
        raise RuntimeError("導入先がリポジトリの外か確認できないため削除しません: " + str(directory))
    converter_found = False
    helper_found = False
    for entry in directory.iterdir():
        if linked_directory(entry):
            raise RuntimeError("導入先にリンクがあるため削除しません: " + str(entry))
        if entry.name in ("venv", "cache") and entry.is_dir():
            continue
        if (entry.name.startswith("excel-textconv-") and entry.name.endswith(".py")
                and entry.is_file()):
            converter_found = True
            continue
        if (entry.name.startswith("excel-stage-") and entry.is_dir()
                and (entry / "excel-stage.py").is_file()):
            helper_found = True
            continue
        raise RuntimeError("導入スクリプトが作った形ではないため削除しません: " + str(directory))
    if not (directory / "venv").is_dir() or not (converter_found or helper_found):
        raise RuntimeError("導入スクリプトが作った形ではないため削除しません: " + str(directory))
    return True


def uninstall(directory, root):
    # Complete every validation before changing Git, attributes or installation.
    attributes = attributes_file()
    try:
        existing = attributes.read_bytes()
    except FileNotFoundError:
        existing = b""
    remaining = without_attribute_block(existing)
    remove_directory = validate_uninstall_directory(directory, root)
    configured = {key: git_values(key) for key in UNINSTALL_KEYS}
    for key, values in configured.items():
        if values != git_values(key, includes=False):
            raise RuntimeError(
                "Git の include 先に " + key + " があります。include 先を整理してから再実行してください。"
                "設定は変更していません。"
            )
    removed = False
    for key, values in configured.items():
        if not values:
            continue
        result = subprocess.run(
            ["git", "config", "--global", "--unset-all", key],
            capture_output=True, text=True, encoding="utf-8",
        )
        if result.returncode not in (0, 5):
            raise RuntimeError("Git 設定を削除できません: " + result.stderr.strip())
        if result.returncode == 0:
            print("削除した Git 設定: " + key)
            removed = True
    if remaining != existing:
        attributes.write_bytes(remaining)
        print("削除した attributes の Excel 用区画: " + str(attributes))
        removed = True
    if remove_directory:
        shutil.rmtree(directory)
        print("削除した導入先: " + str(directory))
        removed = True
    if not removed:
        print("消すものがありません")
    return 0


def main():
    if sys.version_info < (3, 9):
        print("Python 3.9 以降で実行してください。", file=sys.stderr)
        return 1
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--install-dir", type=Path,
                        help="仮想環境と Excel 用スクリプトの保存先 (既定: ユーザーのデータ領域の excel-diff。リポジトリの外)")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="設定を変更せず、インストール状態を確認する")
    mode.add_argument("--uninstall", action="store_true", help="Excel 用の Git 設定・attributes の区画・導入先を削除する")
    args = parser.parse_args()
    try:
        root = Path(__file__).resolve().parent.parent
        selected_directory = args.install_dir
        if (args.check or args.uninstall) and selected_directory is None:
            commands = git_values("diff.excel.textconv")
            registered = registered_converter(commands[-1]) if commands else None
            if registered is not None:
                selected_directory = registered.parent
        requested_directory = (selected_directory or default_install_directory()).expanduser()
        if args.uninstall and linked_directory(requested_directory):
            raise RuntimeError("導入先のリンクは削除しません: " + str(requested_directory))
        directory = requested_directory.resolve()
        if directory == root or root in directory.parents:
            raise RuntimeError("--install-dir はこのリポジトリの外を指定してください。")
        if args.uninstall:
            return uninstall(directory, root)
        source_path = root / "scripts/excel-textconv.py"
        source = source_path.read_bytes()
        helper_source = (root / "scripts/excel-stage.py").read_bytes()
        module_source = (root / "scripts/excel_stage_ooxml.py").read_bytes()
        requirements_path = root / "requirements-excel-diff.txt"
        requirements = requirements_path.read_bytes()
        attributes = attributes_file()
        # Read before installing, so unreadable attributes do not cause a partial setup.
        existing_attributes = attributes.read_bytes() if attributes.exists() else b""
        python = directory / "venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        if args.check:
            if not python.is_file():
                raise RuntimeError("Excel 用の仮想環境がありません。セットアップを実行してください。")
        else:
            print("Excel 用の仮想環境と依存関係を準備しています。", flush=True)
            venv.EnvBuilder(with_pip=True).create(directory / "venv")
            subprocess.run(
                [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(requirements_path)],
                check=True,
            )
        versions = dependency_versions(python)
        converter = converter_path(directory, source, requirements, versions)
        helper = stage_path(directory, helper_source, module_source, requirements, versions)
        expected = {
            "diff.excel.textconv": textconv_command(python, converter),
            "alias.excel-stage": stage_alias(python, helper),
            "alias.excel-merge": merge_alias(python, helper),
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
            helper.parent.mkdir(parents=True, exist_ok=True)
            helper.write_bytes(helper_source)
            helper.with_name("excel_stage_ooxml.py").write_bytes(module_source)
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
        verify(converter, source, attributes, expected, helper, helper_source, module_source)
        print("Excel 差分と部分ステージの設定を確認しました。" if args.check
              else "Excel 差分と部分ステージをグローバルに設定しました。")
        print("変換スクリプト: " + str(converter))
        print("部分ステージ: " + str(helper))
        print("取り込み: git excel-merge")
        print("attributes: " + str(attributes))
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.CalledProcessError, KeyboardInterrupt) as error:
        print("Excel 差分のセットアップに失敗しました: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
