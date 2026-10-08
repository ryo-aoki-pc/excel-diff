#!/usr/bin/env python3
"""Excel のセル変更を delta の左右比較で選び、Excel のまま部分ステージする。"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile

from excel_stage_ooxml import Workbook, WorkbookError, parse_record


class StageError(ValueError):
    """An unsupported selection or a changed source; never apply blindly."""


SUPPORTED = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def isolated_env():
    env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
    env.update(GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull,
               GIT_LITERAL_PATHSPECS="1", GIT_TERMINAL_PROMPT="0")
    return env


def source_env(index=None):
    # The explicit repository and index also work when called from the picker.
    env = {key: value for key, value in os.environ.items() if not key.startswith("GIT_")}
    env["GIT_LITERAL_PATHSPECS"] = "1"
    if index is not None:
        env["GIT_INDEX_FILE"] = str(index)
    return env


def git(repo, *args, env=None, data=None, check=True):
    result = subprocess.run(["git", "-C", str(repo), *args], input=data,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
    if check and result.returncode:
        message = result.stderr.decode("utf-8", errors="replace").strip()
        raise StageError("Git の操作に失敗しました: " + message)
    return result


def entries(repo, env, path=None):
    args = ["ls-files", "--stage", "-z"]
    if path is not None:
        args += ["--", path]
    records = []
    for row in git(repo, *args, env=env).stdout.split(b"\0"):
        if row:
            metadata, name = row.split(b"\t", 1)
            mode, oid, stage = metadata.decode("ascii").split()
            records.append((os.fsdecode(name), mode, oid, stage))
    return records


def workbook_entry(repo, path, env):
    rows = entries(repo, env, path)
    if len(rows) != 1 or rows[0][0] != path or rows[0][3] != "0":
        raise StageError("追跡済みで競合のない Excel を選んでください。新規ファイルは全体をステージします。")
    if rows[0][1] not in ("100644", "100755"):
        raise StageError("通常の Excel ファイルだけを部分ステージできます。")
    return rows[0][1], rows[0][2]


def write_json(path, value):
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def prepare_session(repo, relative, directory, progress=None):
    """Capture the workbook pair and create the isolated text selection repo."""
    repo, directory = Path(repo).resolve(), Path(directory).resolve()
    relative = str(relative)
    if Path(relative).suffix.lower() not in SUPPORTED:
        raise StageError("部分ステージの対象は .xlsx / .xlsm / .xltx / .xltm です。それ以外のファイルは全体をステージしてください。")
    work_path = repo / relative
    try:
        normalized = work_path.resolve().relative_to(repo).as_posix()
    except ValueError:
        raise StageError("リポジトリ内の Excel を指定してください。")
    if work_path.is_symlink() or not work_path.is_file():
        raise StageError("削除されたファイルやリンクは部分ステージできません。")
    relative = normalized
    # Respect an alternate index at entry; the private picker never inherits it.
    index = Path(git(repo, "rev-parse", "--path-format=absolute", "--git-path", "index").stdout.decode().strip())
    env = source_env(index)
    mode, oid = workbook_entry(repo, relative, env)
    head = git(repo, "ls-tree", "-z", "HEAD", "--", relative, env=env, check=False)
    if head.returncode or not head.stdout:
        raise StageError("新規ファイルは Excel 全体をステージしてください。既存ファイルのセル変更が対象です。")
    base_data = git(repo, "cat-file", "blob", oid, env=env).stdout
    work_data = work_path.read_bytes()
    if progress is not None:
        # Reading a large workbook takes seconds with nothing else on screen.
        progress("Excel を読み込んでいます。大きいブックでは選択画面が開くまで時間がかかります。")
    base, work = Workbook(base_data), Workbook(work_data)
    base.assert_compatible(work)
    baseline, changed = base.projection(work), work.projection(base)
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "base.xlsx").write_bytes(base_data)
    (directory / "work.xlsx").write_bytes(work_data)
    picker = directory / "picker"
    picker.mkdir()
    picker_env = isolated_env()
    git(picker, "init", "--quiet", "--template=", env=picker_env)
    options = {"user.name": "Excel staging", "user.email": "excel-stage@example.invalid",
               "core.autocrlf": "false", "core.quotePath": "false", "core.attributesFile": os.devnull,
               "core.hooksPath": str(directory / "no-hooks"), "core.fsmonitor": "false",
               "commit.gpgSign": "false", "tag.gpgSign": "false"}
    for key, value in options.items():
        git(picker, "config", key, value, env=picker_env)
    sheets = []
    for number, sheet in enumerate(base.sheets, 1):
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", sheet.name).strip(" .") or "Sheet"
        filename = "{:03d}-{}.cells.txt".format(number, safe_name)
        sheets.append({"id": sheet.id, "name": sheet.name, "filename": filename})
        with (picker / filename).open("w", encoding="utf-8", newline="\n") as output:
            output.write(baseline[sheet.id])
    git(picker, "add", "--all", env=picker_env)
    git(picker, "commit", "--quiet", "--no-verify", "-m", "現在の Excel インデックス", env=picker_env)
    for sheet in sheets:
        with (picker / sheet["filename"]).open("w", encoding="utf-8", newline="\n") as output:
            output.write(changed[sheet["id"]])
    state = {"repo": str(repo), "path": relative, "index": str(index), "mode": mode,
             "oid": oid, "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "sheets": sheets, "applied": False}
    write_json(directory / "session.json", state)
    return state


def selected_cells(directory, state, base, work):
    picker, env = directory / "picker", isolated_env()
    rows = entries(picker, env)
    names = {sheet["filename"] for sheet in state["sheets"]}
    if {row[0] for row in rows} != names or any(row[1] != "100644" or row[3] != "0" for row in rows):
        raise StageError("選択画面のファイルが削除・追加・競合しています。セルのハンクだけを選んでください。")
    selected = set()
    baseline = base.projection(work)
    for sheet in state["sheets"]:
        sheet_id = sheet["id"]
        expected = [parse_record(line)[0] for line in baseline[sheet_id].splitlines()]
        text = git(picker, "show", ":" + sheet["filename"], env=env).stdout.decode("utf-8")
        records = [parse_record(line) for line in text.splitlines()]
        if [coord for coord, _ in records] != expected:
            raise StageError(sheet["name"] + " のセル置換が不完全です。削除行と追加行を一緒に選んでください。")
        for coordinate, cell in records:
            key = sheet_id, coordinate
            before, after = base.cells.get(key), work.cells.get(key)
            if cell != before and cell != after:
                raise StageError(sheet["name"] + "!" + coordinate + " の選択が不完全です。セル変更全体を選んでください。")
            if cell != before:
                selected.add(key)
    return selected


def assert_work_unchanged(repo, state):
    path = repo / state["path"]
    if path.is_symlink() or not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != state["work_sha256"]:
        raise StageError("作業ファイルが選択開始後に変更されています。画面を閉じて選び直してください。")


def apply_session(directory):
    """Apply once, preserving concurrent index edits to other files."""
    directory = Path(directory).resolve()
    state_path = directory / "session.json"
    state = json.loads(state_path.read_text(encoding="utf-8"))
    if state["applied"]:
        raise StageError("この選択は反映済みです。q で戻り、E で新しく開いてください。")
    base = Workbook((directory / "base.xlsx").read_bytes())
    work = Workbook((directory / "work.xlsx").read_bytes())
    selected = selected_cells(directory, state, base, work)
    if not selected:
        return 0
    rebuilt = base.rebuild(work, selected)
    repo, index = Path(state["repo"]), Path(state["index"])
    assert_work_unchanged(repo, state)
    lock = Path(str(index) + ".lock")
    try:
        descriptor = os.open(lock, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    except FileExistsError:
        raise StageError("Git のインデックスは別の操作が使用中です。完了後に X を押してください。")
    os.close(descriptor)
    try:
        env = source_env(index)
        mode, oid = workbook_entry(repo, state["path"], env)
        if (mode, oid) != (state["mode"], state["oid"]):
            raise StageError("対象 Excel のインデックスが変更されています。画面を閉じて選び直してください。")
        shutil.copyfile(index, lock)
        os.chmod(lock, stat.S_IMODE(index.stat().st_mode))
        new_oid = git(repo, "hash-object", "-w", "--no-filters", "--stdin", env=env, data=rebuilt).stdout.decode("ascii").strip()
        git(repo, "update-index", "--add", "--cacheinfo", mode, new_oid, state["path"], env=source_env(lock))
        assert_work_unchanged(repo, state)
        os.replace(lock, index)
        state["applied"] = True
        write_json(state_path, state)
        return len(selected)
    finally:
        if lock.exists():
            lock.unlink()


def check_tools():
    for name, minimum, pattern in (
        ("lazygit", (0, 66, 0), r"version=(\d+)\.(\d+)\.(\d+)"),
        ("delta", (0, 20, 1), r"delta\s+(\d+)\.(\d+)\.(\d+)"),
    ):
        if shutil.which(name) is None:
            raise StageError(name + " がありません。先に導入してください。")
        result = subprocess.run([name, "--version"], capture_output=True, text=True, encoding="utf-8")
        match = re.search(pattern, result.stdout)
        if result.returncode or not match or tuple(map(int, match.groups())) < minimum:
            raise StageError(name + " は " + ".".join(map(str, minimum)) + " 以降が必要です。")


def picker_config(directory):
    # JSON is also valid YAML; no YAML dependency or per-machine paths in the
    # synced outer config. Paths here exist only for this private session.
    helper = Path(__file__).resolve()
    command_parts = [sys.executable, str(helper), "--apply-session", str(directory)]
    if os.name == "nt":
        command = subprocess.list2cmdline(command_parts)
    else:
        import shlex
        command = shlex.join(command_parts)
    config = {
        "gui": {"useHunkModeInDiffView": True, "language": "ja", "mouseEvents": False,
                "showRandomTip": False, "showFileTree": False},
        "git": {"autoFetch": False, "diffRenderers": [
            {"command": "delta --no-gitconfig --dark --paging=never --syntax-theme=none --tabs=4 --side-by-side",
             "name": "delta side-by-side"}, {"type": "rawGit", "name": "default"}]},
        "update": {"method": "never"}, "disableStartupPopups": True,
        "customCommands": [{"key": "X", "context": "global", "command": command,
                            "description": "選択セルを Excel へステージ (反映後 q で戻る)",
                            "output": "popup", "outputTitle": "Excel の部分ステージ"}],
    }
    path = directory / "config.json"
    write_json(path, config)
    return path


def run_picker(directory):
    config = picker_config(directory)
    state_dir = directory / "lazygit-state"
    state_dir.mkdir()
    return subprocess.run(["lazygit", "--use-config-dir", str(state_dir),
                           "--use-config-file", str(config)], cwd=directory / "picker",
                          env=isolated_env()).returncode


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", nargs="?", help="既存の Excel ファイル")
    parser.add_argument("--apply-session", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    try:
        if args.apply_session is not None:
            count = apply_session(args.apply_session)
            print(str(count) + " セルを Excel のままステージしました。q で元の画面へ戻ってください。" if count else
                  "セル変更が選択されていません。Enter・Space でハンクを選び、X で反映してください。")
            return 0
        if not args.file:
            parser.error("Excel ファイルを指定してください")
        check_tools()
        repo = Path(git(Path.cwd(), "rev-parse", "--show-toplevel").stdout.decode().strip())
        prefix = os.environ.get("GIT_PREFIX", "")
        path = str(Path(prefix) / args.file) if prefix else args.file
        with tempfile.TemporaryDirectory(prefix="lazygit-excel-stage-") as temporary:
            directory = Path(temporary)
            prepare_session(repo, path, directory, progress=lambda message: print(message, flush=True))
            result = run_picker(directory)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            print("Excel の部分ステージを反映しました。" if state["applied"] else "選択を取り消しました。Excel のインデックスは変更していません。")
            return result
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print("Excel の部分ステージに失敗しました: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
