#!/usr/bin/env python3
"""Excel のセル変更を delta の左右比較で選び、Excel のまま部分ステージする。"""

import argparse
import hashlib
import io
from itertools import zip_longest
import json
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time

from excel_stage_ooxml import Workbook, WorkbookError, coordinate_key, parse_record


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
    base, work, baseline, changed = _load_books(base_data, work_data)
    _save_snapshots(directory, base_data, work_data)
    sheets = _build_picker(directory, base, baseline, changed, "現在の Excel インデックス")
    state = {"repo": str(repo), "path": relative, "index": str(index), "mode": mode,
             "oid": oid, "base_sha256": hashlib.sha256(base_data).hexdigest(),
             "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "sheets": sheets, "target": "index", "applied": False}
    write_json(directory / "session.json", state)
    return state


def _load_books(base_data, work_data):
    base, work = Workbook(base_data), Workbook(work_data)
    base.assert_compatible(work)
    return base, work, base.projection(work), work.projection(base)


def _save_snapshots(directory, base_data, work_data):
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "base.xlsx").write_bytes(base_data)
    (directory / "work.xlsx").write_bytes(work_data)


def _build_picker(directory, base, baseline, changed, message):
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
    baseline_directory = directory / "baseline"
    baseline_directory.mkdir(exist_ok=True)
    sheets = []
    for number, sheet in enumerate(base.sheets, 1):
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", sheet.name).strip(" .") or "Sheet"
        filename = "{:03d}-{}.cells.txt".format(number, safe_name)
        baseline_data = baseline[sheet.id].encode("utf-8")
        sheets.append({"id": sheet.id, "name": sheet.name, "filename": filename,
                       "baseline_sha256": hashlib.sha256(baseline_data).hexdigest()})
        # Keep the trusted original projection outside the editable picker repo.
        (baseline_directory / filename).write_bytes(baseline_data)
        (picker / filename).write_bytes(baseline_data)
    git(picker, "add", "--all", env=picker_env)
    git(picker, "commit", "--quiet", "--no-verify", "-m", message, env=picker_env)
    for sheet in sheets:
        with (picker / sheet["filename"]).open("w", encoding="utf-8", newline="\n") as output:
            output.write(changed[sheet["id"]])
    return sheets


def _record_lines(data):
    """Yield splitlines-compatible records without a list of every cell line."""
    with io.TextIOWrapper(io.BytesIO(data), encoding="utf-8", newline="") as source:
        pending = ""
        while True:
            chunk = source.read(65536)
            if not chunk:
                yield from pending.splitlines()
                return
            lines = (pending + chunk).splitlines(keepends=True)
            # Retain the last line even when terminated: a CR at the chunk
            # boundary may be the first half of CRLF in the next chunk.
            pending = lines.pop() if lines else ""
            for line in lines:
                yield line[:-2] if line.endswith("\r\n") else line[:-1]


def _selected_key(before_line, selected_line, sheet_id, base, work, incomplete, invalid):
    expected, _ = parse_record(before_line)
    coordinate, cell = parse_record(selected_line)
    if coordinate != expected:
        raise StageError(incomplete)
    key = sheet_id, coordinate
    before, after = base.cells.get(key), work.cells.get(key)
    if cell != before and cell != after:
        raise StageError(invalid(coordinate) if callable(invalid) else invalid)
    return key if cell != before else None


def selected_cells(directory, state, base, work):
    picker, env = directory / "picker", isolated_env()
    rows = entries(picker, env)
    names = {sheet["filename"] for sheet in state["sheets"]}
    if {row[0] for row in rows} != names or any(row[1] != "100644" or row[3] != "0" for row in rows):
        raise StageError("選択画面のファイルが削除・追加・競合しています。セルのハンクだけを選んでください。")
    base.assert_compatible(work)
    selected = set()
    missing = object()
    for sheet in state["sheets"]:
        sheet_id = sheet["id"]
        if "baseline_sha256" not in sheet:
            raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
        baseline = (directory / "baseline" / sheet["filename"]).read_bytes()
        if hashlib.sha256(baseline).hexdigest() != sheet["baseline_sha256"]:
            raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
        # splitlines historically treats these unescaped JSON string
        # characters as record boundaries. Preserve its rejection of such
        # projections instead of accepting split fragments without parsing.
        validate_all = any(separator in baseline for separator in (b"\xc2\x85", b"\xe2\x80\xa8", b"\xe2\x80\xa9"))
        text = git(picker, "show", ":" + sheet["filename"], env=env).stdout
        for before_line, selected_line in zip_longest(_record_lines(baseline), _record_lines(text), fillvalue=missing):
            if before_line is missing or selected_line is missing:
                raise StageError(sheet["name"] + " のセル置換が不完全です。削除行と追加行を一緒に選んでください。")
            # The captured projection already describes the validated base.
            # Most lines are identical, including when a large book has only
            # a handful of edits; only replacements need JSON validation.
            if before_line == selected_line and not validate_all:
                continue
            key = _selected_key(before_line, selected_line, sheet_id, base, work,
                                sheet["name"] + " のセル置換が不完全です。削除行と追加行を一緒に選んでください。",
                                lambda coordinate: sheet["name"] + "!" + coordinate + " の選択が不完全です。セル変更全体を選んでください。")
            if key is not None:
                selected.add(key)
    return selected


def assert_work_unchanged(repo, state):
    path = repo / state["path"]
    if path.is_symlink() or not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != state["work_sha256"]:
        raise StageError("作業ファイルが選択開始後に変更されています。画面を閉じて選び直してください。")


def _session_books(directory, state):
    if "base_sha256" not in state:
        raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    base_data = (directory / "base.xlsx").read_bytes()
    work_data = (directory / "work.xlsx").read_bytes()
    if (hashlib.sha256(base_data).hexdigest() != state["base_sha256"] or
            hashlib.sha256(work_data).hexdigest() != state["work_sha256"]):
        raise StageError("選択開始時の Excel が変更されています。画面を閉じて選び直してください。")
    base, work = Workbook(base_data), Workbook(work_data)
    return base, work


def apply_session(directory):
    """Apply once, preserving concurrent index edits to other files."""
    directory = Path(directory).resolve()
    state_path = directory / "session.json"
    state = json.loads(state_path.read_text(encoding="utf-8"))
    if state.get("target", "index") != "index":
        raise StageError("取り込みの選択画面は部分ステージに使えません。")
    if state["applied"]:
        raise StageError("この選択は反映済みです。q で戻り、E で新しく開いてください。")
    base, work = _session_books(directory, state)
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


def picker_config(directory, merge=False):
    # JSON is also valid YAML; no YAML dependency or per-machine paths in the
    # synced outer config. Paths here exist only for this private session.
    helper = Path(__file__).resolve()
    command_parts = [sys.executable, str(helper), "--apply-merge-session" if merge else "--apply-session", str(directory)]
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
                            "description": ("選択セルを取り込み先へ書き込む (反映後 q で戻る)" if merge else
                                            "選択セルを Excel へステージ (反映後 q で戻る)"),
                            "output": "popup", "outputTitle": "Excel の取り込み" if merge else "Excel の部分ステージ"}],
    }
    path = directory / "config.json"
    write_json(path, config)
    return path


def run_picker(directory, merge=False):
    config = picker_config(directory, merge=merge)
    state_dir = directory / "lazygit-state"
    state_dir.mkdir()
    return subprocess.run(["lazygit", "--use-config-dir", str(state_dir),
                           "--use-config-file", str(config)], cwd=directory / "picker",
                          env=isolated_env()).returncode


def cleanup_session(directory, retries=20, delay=0.5):
    """Remove a private session without changing the staging result."""
    def remove_readonly(function, path, exc_info):
        error = exc_info[1]
        if isinstance(error, FileNotFoundError):
            return
        if not isinstance(error, PermissionError):
            raise error
        try:
            attributes = os.lstat(path)
            # Do not change permissions through a symlink or Windows junction.
            if stat.S_ISLNK(attributes.st_mode) or (
                getattr(attributes, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
            ):
                raise error
            os.chmod(path, attributes.st_mode | stat.S_IWRITE)
            # POSIX callbacks such as os.open require more than a path. Let
            # the next rmtree attempt retry them after fixing permissions.
            if function not in (os.unlink, os.remove, os.rmdir):
                raise error
            function(path)
        except FileNotFoundError:
            pass

    # Windows lazygit can leave git diff/delta running after it exits. Wait
    # briefly for their handles to close; cleanup must not undo a success.
    reason = ""
    for attempt in range(retries + 1):
        try:
            shutil.rmtree(directory, onerror=remove_readonly)
        except FileNotFoundError:
            return True
        except PermissionError:
            if attempt < retries:
                time.sleep(delay)
                continue
            reason = "（ほかのプロセスが使用中）"
        except OSError:
            reason = ""
        else:
            return True
        break
    print("一時フォルダーを消せませんでした" + reason + "。後で消してください: " + str(directory), file=sys.stderr)
    return False


def resolve_user_path(value):
    path = Path(value)
    if not path.is_absolute():
        path = Path.cwd() / os.environ.get("GIT_PREFIX", "") / path
    # abspath normalizes x/../ without hiding a final symlink from the checks.
    return Path(os.path.abspath(path))


def check_target_writable(path):
    path = Path(path)
    try:
        attributes = path.lstat()
    except FileNotFoundError:
        raise StageError("取り込み先がありません: " + str(path)) from None
    if stat.S_ISLNK(attributes.st_mode) or not stat.S_ISREG(attributes.st_mode):
        raise StageError("取り込み先はリンクではない通常のファイルを指定してください: " + str(path))
    if attributes.st_nlink > 1:
        raise StageError("ハードリンクの取り込み先には書き込めません: " + str(path))
    if not os.access(path, os.W_OK) or not attributes.st_mode & (stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH):
        raise StageError("取り込み先に書き込めません（読み取り専用）: " + str(path))
    if os.name == "nt":
        try:
            descriptor = os.open(path, os.O_RDWR)
            os.close(descriptor)
        except PermissionError as error:
            raise PermissionError("取り込み先を Excel などで開いているなら閉じてからやり直してください: " + str(path)) from error
    elif (path.parent / (".~lock." + path.name + "#")).exists():
        raise StageError("LibreOffice のロックファイルがあります。取り込み先を閉じてください: " + str(path))
    return attributes


def _same_file(target, source):
    left, right = target.stat(), source.stat()
    if left.st_ino and right.st_ino:
        return os.path.samefile(target, source)
    return os.path.normcase(os.path.realpath(target)) == os.path.normcase(os.path.realpath(source))


def _merge_projections(base, work, baseline=None, changed=None):
    if baseline is None:
        baseline, changed = base.projection(work), work.projection(base)
    differences, unsupported = 0, []
    blocked = set()
    names = {sheet.id: sheet.name for sheet in base.sheets}
    for key in base.cells.keys() | work.cells.keys():
        if base.cells.get(key) == work.cells.get(key):
            continue
        differences += 1
        reasons = list(dict.fromkeys(reason for reason in
                                    (base.unsupported_reason(key), work.unsupported_reason(key)) if reason))
        if reasons:
            blocked.add(key)
            unsupported.append({"sheet": names[key[0]], "coordinate": key[1], "reason": "・".join(reasons)})
    # Escape the additional Unicode line separators only in merge projections.
    for sheet in base.sheets:
        identifier = sheet.id
        before = baseline[identifier].split("\n")
        after = changed[identifier].split("\n")
        if blocked:
            for number, line in enumerate(before):
                # Only differing records can be blocked; skip parsing the rest.
                if line != after[number] and (identifier, json.loads(line)[0]) in blocked:
                    after[number] = line
        for projection, lines in ((baseline, before), (changed, after)):
            projection[identifier] = "\n".join(lines).replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    unsupported.sort(key=lambda item: (list(names.values()).index(item["sheet"]), coordinate_key(item["coordinate"])))
    return baseline, changed, differences, unsupported


_WINMERGE_HEADERS = (
    "# 右（取り込み元）から左（取り込み先）へ、取り込むセルの行をコピーする",
    "# 上書き保存（Ctrl+S）してから WinMerge を閉じると、取り込み先のブックに書き込む",
)


def _winmerge_text(base, projection):
    pieces = [*_WINMERGE_HEADERS]
    for number, sheet in enumerate(base.sheets, 1):
        title = json.dumps(sheet.name, ensure_ascii=False)
        title = title.replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
        pieces.append("# シート " + str(number) + ": " + title)
        # Empty worksheets contribute a heading, without an extra blank record.
        pieces.extend(projection[sheet.id].rstrip("\n").split("\n") if projection[sheet.id] else [])
    return ("\ufeff" + "\n".join(pieces) + "\n").encode("utf-8")


def prepare_merge_session(target, source, directory, tool="lazygit"):
    target, source, directory = resolve_user_path(target), resolve_user_path(source), Path(directory).resolve()
    for path in (target, source):
        if path.suffix.lower() not in SUPPORTED:
            raise StageError("取り込みの対象は .xlsx / .xlsm / .xltx / .xltm です: " + str(path))
    check_target_writable(target)
    if not source.is_file():
        raise StageError("取り込み元は通常のファイルを指定してください: " + str(source))
    if _same_file(target, source):
        raise StageError("取り込み先と取り込み元が同じファイルです。")
    base_data, work_data = target.read_bytes(), source.read_bytes()
    base, work, baseline, changed = _load_books(base_data, work_data)
    baseline, changed, differences, unsupported = _merge_projections(base, work, baseline, changed)
    _save_snapshots(directory, base_data, work_data)
    state = {"target": "file", "tool": tool, "path": str(target), "source": str(source),
             "base_sha256": hashlib.sha256(base_data).hexdigest(),
             "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "differences": differences, "unsupported": unsupported, "applied": False, "sheets": []}
    if tool == "lazygit":
        state["sheets"] = _build_picker(directory, base, baseline, changed, "取り込み先の Excel")
    elif tool == "winmerge":
        (directory / "winmerge").mkdir()
        (directory / "baseline").mkdir()
        left, right = "winmerge/取り込み先.cells.txt", "winmerge/取り込み元.cells.txt"
        backup = "baseline/取り込み先.cells.txt"
        original = _winmerge_text(base, baseline)
        (directory / left).write_bytes(original)
        (directory / right).write_bytes(_winmerge_text(base, changed))
        (directory / backup).write_bytes(original)
        state["sheets"] = [{"id": sheet.id, "name": sheet.name} for sheet in base.sheets]
        state["winmerge"] = {"left": left, "right": right, "baseline": backup,
                             "baseline_sha256": hashlib.sha256(original).hexdigest()}
    else:
        raise StageError("選択画面の道具が不正です: " + tool)
    write_json(directory / "session.json", state)
    return state


def _assert_target_unchanged(path, expected_sha256):
    check_target_writable(path)
    if hashlib.sha256(path.read_bytes()).hexdigest() != expected_sha256:
        raise StageError("取り込み先が選択開始後に変更されています。画面を閉じて選び直してください: " + str(path))


def replace_file(path, data, expected_sha256):
    """Replace in the same directory, checking the captured target before every attempt."""
    path = Path(path)
    attributes = check_target_writable(path)
    descriptor, temporary_name = tempfile.mkstemp(prefix=".excel-merge-", suffix=path.suffix, dir=path.parent)
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as output:
            output.write(data)
            output.flush()
            os.fsync(output.fileno())
        os.chmod(temporary, stat.S_IMODE(attributes.st_mode))
        if os.name == "nt":
            # Keep the downloaded-file mark so Excel still opens it in
            # Protected View. FAT and some network shares have no streams.
            try:
                zone = Path(str(path) + ":Zone.Identifier").read_bytes()
            except OSError:
                zone = None
            if zone is not None:
                Path(str(temporary) + ":Zone.Identifier").write_bytes(zone)
        for attempt in range(6):
            _assert_target_unchanged(path, expected_sha256)
            try:
                os.replace(temporary, path)
                return
            except PermissionError as error:
                if attempt == 5:
                    raise PermissionError("取り込み先を Excel などで開いているなら閉じてから再試行してください: " + str(path)) from error
                time.sleep(0.2)
    finally:
        temporary.unlink(missing_ok=True)


def _merge_state(directory):
    state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
    if state.get("target") != "file":
        raise StageError("部分ステージの選択画面は取り込みに使えません。")
    if state["applied"]:
        raise StageError("この選択は反映済みです。画面を閉じて新しく選び直してください。")
    return state


def _write_merge(directory, state, base, work, selected):
    if not selected:
        return 0
    rebuilt = base.rebuild(work, selected)
    replace_file(Path(state["path"]), rebuilt, state["base_sha256"])
    state["applied"], state["count"] = True, len(selected)
    write_json(directory / "session.json", state)
    return len(selected)


def apply_merge_session(directory):
    directory = Path(directory).resolve()
    state = _merge_state(directory)
    base, work = _session_books(directory, state)
    selected = (winmerge_selected_cells(directory, state, base, work) if state["tool"] == "winmerge" else
                selected_cells(directory, state, base, work))
    return _write_merge(directory, state, base, work, selected)


def _winmerge_lines(data):
    # Validate encoding strictly; compare records as bytes, without normalizing JSON.
    data.decode("utf-8-sig")
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    lines = data.replace(b"\r\n", b"\n").split(b"\n")
    while lines and lines[-1] == b"":
        lines.pop()
    return lines


def winmerge_selected_cells(directory, state, base, work):
    base.assert_compatible(work)
    details = state["winmerge"]
    original = (directory / details["baseline"]).read_bytes()
    if hashlib.sha256(original).hexdigest() != details["baseline_sha256"]:
        raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
    baseline, changed, _, _ = _merge_projections(base, work)
    if original != _winmerge_text(base, baseline):
        raise StageError("選択開始時の射影が Excel と一致しません。画面を閉じて選び直してください。")
    try:
        saved = _winmerge_lines((directory / details["left"]).read_bytes())
    except UnicodeError as error:
        raise StageError("WinMerge の保存文字コードは UTF-8 にしてください。") from error
    before, after = _winmerge_lines(original), _winmerge_lines(_winmerge_text(base, changed))
    selected, sheet_index, last_coordinate = set(), -1, "A1"
    missing = object()
    for old, new, source in zip_longest(before, saved, after, fillvalue=missing):
        if old is missing or new is missing:
            if isinstance(old, bytes) and old.startswith(b"["):
                last_coordinate = parse_record(old.decode("utf-8"))[0]
            raise StageError("セルの行数が変わっています: " + last_coordinate + "。セルの行全体をコピーしてください。")
        if old.startswith(b"#"):
            if old != new:
                raise StageError("WinMerge の見出しが変更されています。見出しの行を元に戻してください。")
            if old.startswith("# シート ".encode("utf-8")):
                sheet_index += 1
            continue
        coordinate = parse_record(old.decode("utf-8"))[0]
        last_coordinate = coordinate
        sheet = base.sheets[sheet_index]
        location = sheet.name + "!" + coordinate
        if new == old:
            continue
        if new != source:
            try:
                saved_coordinate = parse_record(new.decode("utf-8"))[0]
            except WorkbookError:
                saved_coordinate = "不正なセル行"
            raise StageError(location + " の選択が不正です（保存行: " + saved_coordinate +
                             "）。取り込み先か取り込み元のセルの行全体をコピーしてください。")
        key = _selected_key(old.decode("utf-8"), new.decode("utf-8"), sheet.id, base, work,
                            location + " のセル置換が不完全です。", location + " の選択が不完全です。")
        if key is not None:
            selected.add(key)
    return selected


def find_winmerge():
    if os.name == "nt":
        try:
            import winreg
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, r"Software\Thingamahoochie\WinMerge") as key:
                        candidate, _ = winreg.QueryValueEx(key, "Executable")
                    if isinstance(candidate, str) and os.path.isfile(candidate):
                        return candidate
                except OSError:
                    continue
        except ImportError:
            pass
        for variable, suffix in (("LOCALAPPDATA", ("Programs", "WinMerge", "WinMergeU.exe")),
                                 ("ProgramFiles", ("WinMerge", "WinMergeU.exe"))):
            parent = os.environ.get(variable)
            if parent:
                candidate = os.path.join(parent, *suffix)
                if os.path.isfile(candidate):
                    return candidate
    return shutil.which("WinMergeU.exe")


def choose_tool(name="auto"):
    if name == "winmerge":
        if not find_winmerge():
            raise StageError("WinMerge がありません。WinMerge を導入するか --tool lazygit を指定してください。")
        return name
    if name == "auto":
        return "winmerge" if os.name == "nt" and find_winmerge() else "lazygit"
    if name != "lazygit":
        raise StageError("選択画面の道具が不正です: " + name)
    return name


def _launch_winmerge(directory, state):
    executable = find_winmerge()
    if not executable:
        raise StageError("WinMerge がありません。WinMerge を導入するか --tool lazygit を指定してください。")
    details = state["winmerge"]
    left, right = directory / details["left"], directory / details["right"]
    command = [executable, "/s-", "/e", "/u", "/fl", "/wr", "/t", "text",
               "/dl", "取り込み先（書き換える）: " + Path(state["path"]).name,
               "/dr", "取り込み元: " + Path(state["source"]).name, str(left), str(right)]
    left_before = left.read_bytes()
    started = time.monotonic()
    # The exit status says nothing about the choice; the saved left text does.
    subprocess.run(command)
    if time.monotonic() - started <= 2 and left.read_bytes() == left_before:
        print("WinMerge がすぐ終了しました。保存せずに閉じた場合、取り込み先は変えていません。", file=sys.stderr)


def _winmerge_select(directory, state):
    """Open WinMerge until the saved left text is a valid choice."""
    books = None
    while True:
        _launch_winmerge(directory, state)
        try:
            # Read the snapshots once, after the window closes, so a large
            # workbook does not stay in memory while WinMerge is open.
            if books is None:
                books = _session_books(directory, state)
            return winmerge_selected_cells(directory, state, *books), books
        except (StageError, WorkbookError, UnicodeError) as error:
            if not sys.stdin.isatty():
                raise
            print(str(error), file=sys.stderr)
            if input("WinMerge を開き直して直しますか [y/N]: ").strip().lower() != "y":
                raise


def run_winmerge(directory):
    directory = Path(directory).resolve()
    return _winmerge_select(directory, _merge_state(directory))[0]


def _merge_with_winmerge(directory):
    directory = Path(directory).resolve()
    state = _merge_state(directory)
    selected, (base, work) = _winmerge_select(directory, state)
    while True:
        try:
            return _write_merge(directory, state, base, work, selected)
        except PermissionError as error:
            if not sys.stdin.isatty():
                raise
            print(str(error), file=sys.stderr)
            if input("取り込み先を閉じてから Enter で再試行します（q で中止）: ").strip().lower() == "q":
                raise


def _apply_merge_cli(directory):
    try:
        count = apply_merge_session(directory)
        print(str(count) + " セルを取り込み先へ書き込みました。q で元の画面へ戻ってください。" if count else
              "セル変更が選択されていません。Enter・Space でハンクを選び、X で反映してください。")
        return 0
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print("Excel の取り込みに失敗しました: " + str(error), file=sys.stderr)
        return 1


def merge_main(argv=None):
    parser = argparse.ArgumentParser(description="2 つの Excel のセル変更を取り込む")
    parser.add_argument("--tool", choices=("auto", "winmerge", "lazygit"), default="auto")
    parser.add_argument("target")
    parser.add_argument("source")
    args = parser.parse_args(argv)
    directory = None
    try:
        target, source = resolve_user_path(args.target), resolve_user_path(args.source)
        tool = choose_tool(args.tool)
        if tool == "lazygit":
            check_tools()
        print("書き換えるファイル: " + str(target), flush=True)
        print("取り込み元: " + str(source), flush=True)
        print("Excel を読み込んでいます。大きいブックでは画面が開くまで時間がかかります。", flush=True)
        directory = Path(tempfile.mkdtemp(prefix="excel-merge-"))
        state = prepare_merge_session(target, source, directory, tool=tool)
        unsupported = state["unsupported"]
        if unsupported:
            examples = ", ".join(item["sheet"] + "!" + item["coordinate"] + "（" + item["reason"] + "）" for item in unsupported[:20])
            if len(unsupported) > 20:
                examples += "、ほか " + str(len(unsupported) - 20) + " セル"
            print("取り込めない違い（" + str(len(unsupported)) + " セル）: " + examples, flush=True)
        if state["differences"] == len(unsupported):
            print("取り込める差分がありません（取り込めない違い " + str(len(unsupported)) + " セル）。取り込み先は変えていません。" if unsupported else
                  "差分がありません。取り込み先は変えていません。")
            return 0
        if tool == "winmerge":
            print("WinMerge で、取り込むセルの行を右から左へコピーし、上書き保存（Ctrl+S）してから WinMerge の窓を閉じてください。", flush=True)
            count = _merge_with_winmerge(directory)
            result = 0
        else:
            print("専用画面で、取り込むハンクを Enter → Space で選び、X で書き込んでください。q で閉じると取り消します。", flush=True)
            result = run_picker(directory, merge=True)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            count = state.get("count", 0) if state["applied"] else 0
        print(str(count) + " セルを取り込み先へ書き込みました: " + str(target) if count else
              "取り込むセルが選ばれていません。取り込み先は変えていません。")
        return result
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print("Excel の取り込みに失敗しました: " + str(error), file=sys.stderr)
        return 1
    finally:
        if directory is not None:
            cleanup_session(directory)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments[:1] == ["--merge"]:
        return merge_main(arguments[1:])
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", nargs="?", help="既存の Excel ファイル")
    parser.add_argument("--apply-session", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--apply-merge-session", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(arguments)
    if args.apply_merge_session is not None:
        return _apply_merge_cli(args.apply_merge_session)
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
        directory = Path(tempfile.mkdtemp(prefix="lazygit-excel-stage-"))
        try:
            prepare_session(repo, path, directory, progress=lambda message: print(message, flush=True))
            result = run_picker(directory)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            print("Excel の部分ステージを反映しました。" if state["applied"] else "選択を取り消しました。Excel のインデックスは変更していません。")
            return result
        finally:
            cleanup_session(directory)
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print("Excel の部分ステージに失敗しました: " + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
