#!/usr/bin/env python3
"""Excel のセル変更を delta の左右比較で選び、Excel のまま部分ステージする。"""

import argparse
from contextlib import contextmanager
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
from typing import NamedTuple

from excel_stage_ooxml import (Cell, FORMAT_KIND, Workbook, WorkbookError, cell_record,
                              coordinate_key, parse_record)


class StageError(ValueError):
    """An unsupported selection or a changed source; never apply blindly."""


SUPPORTED = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb"}
WORD_SUPPORTED = {".docx", ".docm", ".dotx", ".dotm"}
EXCEL_FAMILY = SUPPORTED | {".xls", ".xlt", ".xlsb"}


_TEXTS = {
    "excel": {
        "reading_stage": "Excel を読み込んでいます。大きいブックでは選択画面が開くまで時間がかかります。",
        "reading_merge": "Excel を読み込んでいます。大きいブックでは画面が開くまで時間がかかります。",
        "outside": "リポジトリ内の Excel を指定してください。",
        "untracked": "追跡済みで競合のない Excel を選んでください。新規ファイルは全体をステージします。",
        "irregular": "通常の Excel ファイルだけを部分ステージできます。",
        "new_file": "新規ファイルは Excel 全体をステージしてください。既存ファイルのセル変更が対象です。",
        "index_commit": "現在の Excel インデックス",
        "merge_commit": "取り込み先の Excel",
        "snapshots_changed": "選択開始時の Excel が変更されています。画面を閉じて選び直してください。",
        "index_changed": "対象 Excel のインデックスが変更されています。画面を閉じて選び直してください。",
        "work_changed": "作業ファイルが選択開始後に変更されています。画面を閉じて選び直してください。",
        "open_elsewhere": "取り込み先を Excel などで開いているなら閉じてからやり直してください: ",
        "open_retry": "取り込み先を Excel などで開いているなら閉じてから再試行してください: ",
        "merge_suffix": "取り込みの対象は .xlsx / .xlsm / .xltx / .xltm / .xlam / .xlsb です: ",
        "staged": "{}を Excel のままステージしました。q で元の画面へ戻ってください。",
        "nothing_selected": "セル変更が選択されていません。Enter・Space でハンクを選び、X で反映してください。",
        "stage_done": "Excel の部分ステージを反映しました。",
        "stage_cancelled": "選択を取り消しました。Excel のインデックスは変更していません。",
        "stage_failed": "Excel の部分ステージに失敗しました: ",
        "merged_picker": "{}を取り込み先へ書き込みました。q で元の画面へ戻ってください。",
        "merged": "{}を取り込み先へ書き込みました: ",
        "merge_nothing": "取り込むセルが選ばれていません。取り込み先は変えていません。",
        "merge_failed": "Excel の取り込みに失敗しました: ",
        "no_difference": "差分がありません。取り込み先は変えていません。",
        "only_unsupported": "取り込める差分がありません（取り込めない違い {} セル）。取り込み先は変えていません。",
        "unsupported_merge": "取り込めない違い（{} セル）: ",
        "more": "、ほか {} セル",
        "winmerge_hint": "WinMerge で、取り込むセルの行を右から左へコピーし、上書き保存（Ctrl+S）してから WinMerge の窓を閉じてください。",
        "picker_hint": "専用画面で、取り込むハンクを Enter → Space で選び、X で書き込んでください。q で閉じると取り消します。",
        "picker_stage": "選択セルを Excel へステージ (反映後 q で戻る)",
        "picker_merge": "選択セルを取り込み先へ書き込む (反映後 q で戻る)",
        "title_stage": "Excel の部分ステージ",
        "title_merge": "Excel の取り込み",
        "stage_prefix": "lazygit-excel-stage-",
        "merge_prefix": "excel-merge-",
        "replace_prefix": ".excel-merge-",
        "identity_name": "Excel staging",
        "identity_email": "excel-stage@example.invalid",
    },
    "word": {
        "reading_stage": "Word を読み込んでいます。大きい文書では選択画面が開くまで時間がかかります。",
        "reading_merge": "Word を読み込んでいます。大きい文書では画面が開くまで時間がかかります。",
        "outside": "リポジトリ内の Word を指定してください。",
        "untracked": "追跡済みで競合のない Word を選んでください。新規ファイルは全体をステージします。",
        "irregular": "通常の Word ファイルだけを部分ステージできます。",
        "new_file": "新規ファイルは Word 全体をステージしてください。既存ファイルの変更が対象です。",
        "index_commit": "現在の Word インデックス",
        "merge_commit": "取り込み先の Word",
        "snapshots_changed": "選択開始時の Word が変更されています。画面を閉じて選び直してください。",
        "index_changed": "対象 Word のインデックスが変更されています。画面を閉じて選び直してください。",
        "work_changed": "作業ファイルが選択開始後に変更されています。画面を閉じて選び直してください。",
        "open_elsewhere": "取り込み先を Word などで開いているなら閉じてからやり直してください: ",
        "open_retry": "取り込み先を Word などで開いているなら閉じてから再試行してください: ",
        "merge_suffix": "取り込みの対象は .docx / .docm / .dotx / .dotm です: ",
        "staged": "{} か所の変更を Word のままステージしました。q で元の画面へ戻ってください。",
        "nothing_selected": "変更が選択されていません。Enter・Space でハンクを選び、X で反映してください。",
        "stage_done": "Word の部分ステージを反映しました。",
        "stage_cancelled": "選択を取り消しました。Word のインデックスは変更していません。",
        "stage_failed": "Word の部分ステージに失敗しました: ",
        "merged_picker": "{} か所の変更を取り込み先へ書き込みました。q で元の画面へ戻ってください。",
        "merged": "{} か所の変更を取り込み先へ書き込みました: ",
        "merge_nothing": "取り込む変更が選ばれていません。取り込み先は変えていません。",
        "merge_failed": "Word の取り込みに失敗しました: ",
        "no_difference": "差分がありません。取り込み先は変えていません。",
        "only_unsupported": "取り込める差分がありません（取り込めない違い {} か所）。取り込み先は変えていません。",
        "unsupported_merge": "取り込めない違い（{} か所）: ",
        "unsupported_stage": "部分ステージできない違い（{} か所）: ",
        "more": "、ほか {} か所",
        "winmerge_hint": "WinMerge で、取り込む行を右から左へコピーし、上書き保存（Ctrl+S）してから WinMerge の窓を閉じてください。",
        "picker_hint": "専用画面で、取り込むハンクを Enter → Space で選び、X で書き込んでください。q で閉じると取り消します。",
        "picker_stage": "選択した変更を Word へステージ (反映後 q で戻る)",
        "picker_merge": "選択した変更を取り込み先へ書き込む (反映後 q で戻る)",
        "title_stage": "Word の部分ステージ",
        "title_merge": "Word の取り込み",
        "stage_prefix": "lazygit-word-stage-",
        "merge_prefix": "word-merge-",
        "replace_prefix": ".word-merge-",
        "identity_name": "Word staging",
        "identity_email": "word-stage@example.invalid",
    },
}


def _kind(path):
    return "word" if Path(path).suffix.lower() in WORD_SUPPORTED else "excel"


def _session_kind(directory):
    try:
        state = json.loads((Path(directory) / "session.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return "excel"
    return "word" if isinstance(state, dict) and state.get("format") == "word" else "excel"


def _word_module():
    try:
        import word_stage_ooxml
    except ImportError as error:
        raise StageError("Word 用のモジュールを読み込めません。setup-excel-diff.py を再実行してください: " + str(error)) from error
    return word_stage_ooxml


@contextmanager
def _word_errors():
    module = _word_module()
    try:
        yield module
    except module.word.DocumentError as error:
        raise StageError(str(error)) from error


def _word_unsupported(comparison):
    return [item._asdict() for item in comparison.unsupported]


def _build_word_picker(directory, comparison, message, staging=False):
    units = [{"id": story.id, "title": story.title} for story in comparison.stories]
    baseline = {story.id: story.baseline for story in comparison.stories}
    changed = {story.id: story.raw_changed if staging else story.changed for story in comparison.stories}
    texts = _TEXTS["word"]
    return _build_picker_files(directory, units, ((".word.txt", "", baseline, changed),), message,
                               (texts["identity_name"], texts["identity_email"]))


def _report_word_swaps(comparison):
    if comparison.last_swaps:
        print("目次などの参照先に合わせて、表示が同じ {} か所を取り込み元のものにしました。".format(comparison.last_swaps))


class Selection(NamedTuple):
    values: set
    formats: set


class Counts(NamedTuple):
    values: int
    formats: int


class Projections(NamedTuple):
    baseline: dict
    changed: dict
    format_baseline: dict
    format_changed: dict
    format_coordinates: dict = None


class MergeProjections(NamedTuple):
    baseline: dict
    changed: dict
    format_baseline: dict
    format_changed: dict
    differences: int
    format_differences: int
    selectable: int
    unsupported: list


def _count_text(values, formats):
    if not formats:
        return str(values) + " セル"
    if not values:
        return "書式 " + str(formats) + " セル"
    return "値 " + str(values) + " セル・書式 " + str(formats) + " セル"


def _counted(count):
    # Excel reports Counts; Word reports how many places changed.
    return any(count) if isinstance(count, Counts) else bool(count)


def _count_label(count):
    return _count_text(*count) if isinstance(count, Counts) else count


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
    texts = _TEXTS[_kind(path)]
    rows = entries(repo, env, path)
    if len(rows) != 1 or rows[0][0] != path or rows[0][3] != "0":
        raise StageError(texts["untracked"])
    if rows[0][1] not in ("100644", "100755"):
        raise StageError(texts["irregular"])
    return rows[0][1], rows[0][2]


def write_json(path, value):
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def prepare_session(repo, relative, directory, progress=None):
    """Capture the workbook pair and create the isolated text selection repo."""
    repo, directory = Path(repo).resolve(), Path(directory).resolve()
    relative = str(relative)
    kind = _kind(relative)
    texts = _TEXTS[kind]
    if kind == "excel" and Path(relative).suffix.lower() not in SUPPORTED:
        raise StageError("部分ステージの対象は Excel（.xlsx / .xlsm / .xltx / .xltm / .xlam / .xlsb）と Word（.docx / .docm / .dotx / .dotm）です。それ以外のファイルは全体をステージしてください。")
    work_path = repo / relative
    try:
        normalized = work_path.resolve().relative_to(repo).as_posix()
    except ValueError:
        raise StageError(texts["outside"])
    if work_path.is_symlink() or not work_path.is_file():
        raise StageError("削除されたファイルやリンクは部分ステージできません。")
    relative = normalized
    # Respect an alternate index at entry; the private picker never inherits it.
    index = Path(git(repo, "rev-parse", "--path-format=absolute", "--git-path", "index").stdout.decode().strip())
    env = source_env(index)
    mode, oid = workbook_entry(repo, relative, env)
    head = git(repo, "ls-tree", "-z", "HEAD", "--", relative, env=env, check=False)
    if head.returncode or not head.stdout:
        raise StageError(texts["new_file"])
    base_data = git(repo, "cat-file", "blob", oid, env=env).stdout
    work_data = work_path.read_bytes()
    if progress is not None:
        # Reading a large workbook takes seconds with nothing else on screen.
        progress(texts["reading_stage"])
    if kind == "word":
        with _word_errors() as module:
            comparison = module.compare(base_data, work_data)
        _save_snapshots(directory, base_data, work_data, ".docx")
        parts = _build_word_picker(directory, comparison, texts["index_commit"], staging=True)
        state = {"format": "word", "repo": str(repo), "path": relative, "index": str(index), "mode": mode,
                 "oid": oid, "base_sha256": hashlib.sha256(base_data).hexdigest(),
                 "work_sha256": hashlib.sha256(work_data).hexdigest(), "parts": parts,
                 "differences": comparison.differences, "unsupported": _word_unsupported(comparison),
                 "selectable": comparison.selectable, "target": "index", "applied": False}
        write_json(directory / "session.json", state)
        return state
    base, work, projections = _load_books(base_data, work_data)
    _save_snapshots(directory, base_data, work_data, "." + base.workbook_format)
    sheets = _build_picker(directory, base, projections, texts["index_commit"])
    state = {"workbook": base.workbook_format, "repo": str(repo), "path": relative,
             "index": str(index), "mode": mode,
             "oid": oid, "base_sha256": hashlib.sha256(base_data).hexdigest(),
             "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "sheets": sheets, "target": "index", "applied": False}
    write_json(directory / "session.json", state)
    return state


def workbook_from_data(data):
    """Choose the workbook implementation from package content, not its suffix."""
    try:
        from excel_xlsb import is_xlsb
    except ModuleNotFoundError as error:
        # A copied OOXML-only helper can still run without the binary modules.
        if error.name != "excel_xlsb":
            raise
        binary = False
    else:
        binary = is_xlsb(data)

    if binary:
        from excel_stage_xlsb import Workbook as BinaryWorkbook

        book, kind = BinaryWorkbook(data), "xlsb"
    else:
        book, kind = Workbook(data), "xlsx"
    book.workbook_format = kind
    return book


def _load_books(base_data, work_data):
    base, work = workbook_from_data(base_data), workbook_from_data(work_data)
    if base.workbook_format != work.workbook_format:
        raise StageError("形式の違うブック（xlsb と xlsx など）は取り込めません。同じ形式のブックを 2 つ指定してください。")
    base.assert_compatible(work)
    format_baseline, format_changed, coordinates = base.compare_formats(work).projection()
    return base, work, Projections(base.projection(work), work.projection(base),
                                   format_baseline, format_changed, coordinates)


def _save_snapshots(directory, base_data, work_data, suffix=".xlsx"):
    directory.mkdir(parents=True, exist_ok=True)
    (directory / ("base" + suffix)).write_bytes(base_data)
    (directory / ("work" + suffix)).write_bytes(work_data)


def _build_picker(directory, base, projections, message):
    units = [{"id": sheet.id, "name": sheet.name} for sheet in base.sheets]
    texts = _TEXTS["excel"]
    layers = ((".cells.txt", "", projections.baseline, projections.changed),
              (".formats.txt", "formats_", projections.format_baseline, projections.format_changed))
    return _build_picker_files(directory, units, layers, message,
                               (texts["identity_name"], texts["identity_email"]))


def _build_picker_files(directory, units, layers, message, identity):
    """Write each unit's files; a layer is (suffix, state key prefix, baseline, changed)."""
    picker = directory / "picker"
    picker.mkdir()
    picker_env = isolated_env()
    git(picker, "init", "--quiet", "--template=", env=picker_env)
    options = {"user.name": identity[0], "user.email": identity[1],
               "core.autocrlf": "false", "core.quotePath": "false", "core.attributesFile": os.devnull,
               "core.hooksPath": str(directory / "no-hooks"), "core.fsmonitor": "false",
               "commit.gpgSign": "false", "tag.gpgSign": "false"}
    for key, value in options.items():
        git(picker, "config", key, value, env=picker_env)
    baseline_directory = directory / "baseline"
    baseline_directory.mkdir(exist_ok=True)
    sheets = []
    for number, unit in enumerate(units, 1):
        title_key = "title" if "title" in unit else "name"
        title = unit[title_key]
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", title).strip(" .") or ("Part" if title_key == "title" else "Sheet")
        stem = "{:03d}-{}".format(number, safe_name)
        entry = {"id": unit["id"], title_key: title}
        for suffix, prefix, baseline, _ in layers:
            filename = stem + suffix
            baseline_data = baseline[unit["id"]].encode("utf-8")
            entry[prefix + "filename"] = filename
            entry[prefix + "baseline_sha256"] = hashlib.sha256(baseline_data).hexdigest()
            # Keep the trusted original projection outside the editable picker repo.
            (baseline_directory / filename).write_bytes(baseline_data)
            (picker / filename).write_bytes(baseline_data)
        sheets.append(entry)
    git(picker, "add", "--all", env=picker_env)
    git(picker, "commit", "--quiet", "--no-verify", "-m", message, env=picker_env)
    for sheet in sheets:
        for _, prefix, _, changed in layers:
            with (picker / sheet[prefix + "filename"]).open("w", encoding="utf-8", newline="\n") as output:
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


def _changed_lines(directory, picker, env, filename, digest, incomplete):
    baseline = (directory / "baseline" / filename).read_bytes()
    if hashlib.sha256(baseline).hexdigest() != digest:
        raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
    # Preserve splitlines' historical rejection of unescaped value separators.
    validate_all = any(separator in baseline for separator in (b"\xc2\x85", b"\xe2\x80\xa8", b"\xe2\x80\xa9"))
    text = git(picker, "show", ":" + filename, env=env).stdout
    missing, previous = object(), ""
    for before, selected in zip_longest(_record_lines(baseline), _record_lines(text), fillvalue=missing):
        if before is missing or selected is missing:
            # A separating blank line names no cell; use the record before it.
            line = (selected if before is missing else before) or previous
            coordinate = parse_record(line)[0]
            raise StageError(incomplete(coordinate) if callable(incomplete) else incomplete)
        if before:
            previous = before
        # Blank lines only separate format records; equal ones carry no choice.
        if before != selected or (validate_all and before):
            yield before, selected


def _selected_format_key(before_line, selected_line, sheet_id, comparison, incomplete, invalid):
    if not before_line or not selected_line:
        # Staging only one side of a record shifts the records against the
        # separating blank lines; name the record that lost its pair.
        expected = parse_record(before_line or selected_line)[0]
        raise StageError(incomplete(expected) if callable(incomplete) else incomplete)
    expected, before = parse_record(before_line)
    message = incomplete(expected) if callable(incomplete) else incomplete
    coordinate, selected = parse_record(selected_line)
    if (coordinate != expected or before is None or selected is None or
            before.kind != FORMAT_KIND or selected.kind != FORMAT_KIND):
        raise StageError(message)
    key = sheet_id, coordinate
    texts = comparison.texts(key)
    if texts is None or before.value != texts[0] or selected.value not in texts:
        raise StageError(invalid(coordinate) if callable(invalid) else invalid)
    return key if selected.value != texts[0] else None


def selected_cells(directory, state, base, work):
    picker, env = directory / "picker", isolated_env()
    for sheet in state["sheets"]:
        if "baseline_sha256" not in sheet or "formats_baseline_sha256" not in sheet:
            raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    rows = entries(picker, env)
    names = {sheet[key] for sheet in state["sheets"] for key in ("filename", "formats_filename")}
    if {row[0] for row in rows} != names or any(row[1] != "100644" or row[3] != "0" for row in rows):
        raise StageError("選択画面のファイルが削除・追加・競合しています。セルのハンクだけを選んでください。")
    base.assert_compatible(work)
    selected, formats, comparison = set(), set(), None
    for sheet in state["sheets"]:
        incomplete = sheet["name"] + " のセル置換が不完全です。削除行と追加行を一緒に選んでください。"
        invalid = lambda coordinate: sheet["name"] + "!" + coordinate + " の選択が不完全です。セル変更全体を選んでください。"
        for before, after in _changed_lines(directory, picker, env, sheet["filename"], sheet["baseline_sha256"], incomplete):
            key = _selected_key(before, after, sheet["id"], base, work, incomplete, invalid)
            if key is not None:
                selected.add(key)
        format_incomplete = lambda coordinate: (sheet["name"] + "!" + coordinate +
            " の書式の選択が不完全です。書式の行の削除側と追加側を一緒に選んでください。")
        for before, after in _changed_lines(directory, picker, env, sheet["formats_filename"],
                                             sheet["formats_baseline_sha256"], format_incomplete):
            if comparison is None:
                comparison = base.compare_formats(work)
            key = _selected_format_key(before, after, sheet["id"], comparison, format_incomplete, format_incomplete)
            if key is not None:
                formats.add(key)
    return Selection(selected, formats)


def _session_word_comparison(directory, state):
    if "base_sha256" not in state or "work_sha256" not in state:
        raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    base_data = (directory / "base.docx").read_bytes()
    work_data = (directory / "work.docx").read_bytes()
    if (hashlib.sha256(base_data).hexdigest() != state["base_sha256"] or
            hashlib.sha256(work_data).hexdigest() != state["work_sha256"]):
        raise StageError(_TEXTS["word"]["snapshots_changed"])
    with _word_errors() as module:
        return module.compare(base_data, work_data)


def selected_word(directory, state, comparison):
    picker, env = directory / "picker", isolated_env()
    rows = entries(picker, env)
    names = {part["filename"] for part in state["parts"]}
    if {row[0] for row in rows} != names or any(row[1] != "100644" or row[3] != "0" for row in rows):
        raise StageError("選択画面のファイルが削除・追加・競合しています。変更のハンクだけを選んでください。")
    if [part["id"] for part in state["parts"]] != [story.id for story in comparison.stories]:
        raise StageError("選択開始時の射影が Word 文書と一致しません。画面を閉じて選び直してください。")
    selected = set()
    for part, story in zip(state["parts"], comparison.stories):
        if "baseline_sha256" not in part:
            raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
        baseline = (directory / "baseline" / part["filename"]).read_bytes()
        if hashlib.sha256(baseline).hexdigest() != part["baseline_sha256"]:
            raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
        if baseline != story.baseline.encode("utf-8"):
            raise StageError("選択開始時の射影が Word 文書と一致しません。画面を閉じて選び直してください。")
        staged = git(picker, "show", ":" + part["filename"], env=env).stdout.decode("utf-8").splitlines()
        with _word_errors():
            selected.update(comparison.choose(story.id, staged,
                                              projection="changed" if state.get("target") == "file" else "raw_changed"))
    return selected


def assert_work_unchanged(repo, state):
    texts = _TEXTS[state.get("format", "excel")]
    path = repo / state["path"]
    if path.is_symlink() or not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != state["work_sha256"]:
        raise StageError(texts["work_changed"])


def _session_books(directory, state):
    texts = _TEXTS[state.get("format", "excel")]
    if "base_sha256" not in state:
        raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    kind = state.get("workbook", "xlsx")
    if kind not in ("xlsx", "xlsb"):
        raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    base_data = (directory / ("base." + kind)).read_bytes()
    work_data = (directory / ("work." + kind)).read_bytes()
    if (hashlib.sha256(base_data).hexdigest() != state["base_sha256"] or
            hashlib.sha256(work_data).hexdigest() != state["work_sha256"]):
        raise StageError(texts["snapshots_changed"])
    base, work = workbook_from_data(base_data), workbook_from_data(work_data)
    if base.workbook_format != kind or work.workbook_format != kind:
        raise StageError(texts["snapshots_changed"])
    return base, work


def apply_session(directory):
    """Apply once, preserving concurrent index edits to other files."""
    directory = Path(directory).resolve()
    state_path = directory / "session.json"
    state = json.loads(state_path.read_text(encoding="utf-8"))
    texts = _TEXTS[state.get("format", "excel")]
    if state.get("target", "index") != "index":
        raise StageError("取り込みの選択画面は部分ステージに使えません。")
    if state["applied"]:
        raise StageError("この選択は反映済みです。q で戻り、E で新しく開いてください。")
    word_session = state.get("format") == "word"
    if word_session:
        comparison = _session_word_comparison(directory, state)
        selected = selected_word(directory, state, comparison)
        if not selected:
            return 0
        with _word_errors():
            rebuilt = comparison.rebuild(selected)
        count = len(selected)
    else:
        base, work = _session_books(directory, state)
        selected = selected_cells(directory, state, base, work)
        if not any(selected):
            return Counts(0, 0)
        rebuilt = base.rebuild(work, selected.values, selected.formats)
        count = Counts(len(selected.values), len(selected.formats))
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
            raise StageError(texts["index_changed"])
        shutil.copyfile(index, lock)
        os.chmod(lock, stat.S_IMODE(index.stat().st_mode))
        new_oid = git(repo, "hash-object", "-w", "--no-filters", "--stdin", env=env, data=rebuilt).stdout.decode("ascii").strip()
        git(repo, "update-index", "--add", "--cacheinfo", mode, new_oid, state["path"], env=source_env(lock))
        assert_work_unchanged(repo, state)
        os.replace(lock, index)
        state["applied"] = True
        write_json(state_path, state)
        if word_session:
            _report_word_swaps(comparison)
        return count
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
    texts = _TEXTS[_session_kind(directory)]
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
                            "description": (texts["picker_merge"] if merge else
                                            texts["picker_stage"]),
                            "output": "popup", "outputTitle": texts["title_merge"] if merge else texts["title_stage"]}],
    }
    if _session_kind(directory) == "word":
        config["git"]["diffRenderers"].insert(0, {
            "command": "delta --no-gitconfig --dark --paging=never --syntax-theme=none --tabs=4 --side-by-side --wrap-max-lines=unlimited",
            "name": "delta side-by-side wrapped"})
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
    texts = _TEXTS[_kind(path)]
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
            raise PermissionError(texts["open_elsewhere"] + str(path)) from error
    elif (path.parent / (".~lock." + path.name + "#")).exists():
        raise StageError("LibreOffice のロックファイルがあります。取り込み先を閉じてください: " + str(path))
    return attributes


def _same_file(target, source):
    left, right = target.stat(), source.stat()
    if left.st_ino and right.st_ino:
        return os.path.samefile(target, source)
    return os.path.normcase(os.path.realpath(target)) == os.path.normcase(os.path.realpath(source))


def _merge_projections(base, work, projections=None):
    if projections is None:
        format_baseline, format_changed, coordinates = base.compare_formats(work).projection()
        projections = Projections(base.projection(work), work.projection(base), format_baseline, format_changed,
                                  coordinates)
    baseline, changed, format_baseline, format_changed = (dict(part) for part in projections[:4])
    differences, format_differences, selectable = 0, 0, 0
    blocked, format_blocked, reasons = set(), set(), {}
    names = {sheet.id: sheet.name for sheet in base.sheets}
    order = {sheet.id: number for number, sheet in enumerate(base.sheets)}
    for key in base.cells.keys() | work.cells.keys():
        if base.cells.get(key) == work.cells.get(key):
            continue
        differences += 1
        problems = list(dict.fromkeys(reason for reason in
                                     (base.unsupported_reason(key), work.unsupported_reason(key)) if reason))
        if problems:
            blocked.add(key)
            reasons[key] = problems
        else:
            selectable += 1
    comparison = base.compare_formats(work)
    # The projection already found the differing formats; do not compare again.
    format_keys = ((identifier, coordinate) for identifier, coordinates in projections.format_coordinates.items()
                   for coordinate in coordinates) if projections.format_coordinates is not None else comparison.keys()
    for key in format_keys:
        format_differences += 1
        reason = comparison.unsupported_reason(key)
        if reason:
            format_blocked.add(key)
            reasons.setdefault(key, []).append(reason)
        else:
            selectable += 1
    # Escape the additional Unicode line separators only in merge projections.
    for before_text, after_text, hidden in ((baseline, changed, blocked),
                                            (format_baseline, format_changed, format_blocked)):
        for sheet in base.sheets:
            identifier = sheet.id
            before, after = before_text[identifier].split("\n"), after_text[identifier].split("\n")
            if hidden:
                for number, line in enumerate(before):
                    if line != after[number] and (identifier, json.loads(line)[0]) in hidden:
                        after[number] = line
            for projection, lines in ((before_text, before), (after_text, after)):
                projection[identifier] = "\n".join(lines).replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    unsupported = [{"sheet": names[key[0]], "coordinate": key[1], "reason": "・".join(dict.fromkeys(reasons[key]))}
                   for key in sorted(reasons, key=lambda key: (order[key[0]], coordinate_key(key[1])))]
    return MergeProjections(baseline, changed, format_baseline, format_changed, differences,
                            format_differences, selectable, unsupported)


_WINMERGE_HEADERS = (
    "# 右（取り込み元）から左（取り込み先）へ、取り込むセルの行をコピーする",
    "# 上書き保存（Ctrl+S）してから WinMerge を閉じると、取り込み先のブックに書き込む",
)


def _winmerge_text(base, projection, formats=None):
    pieces = [*_WINMERGE_HEADERS]
    for number, sheet in enumerate(base.sheets, 1):
        title = json.dumps(sheet.name, ensure_ascii=False)
        title = title.replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
        pieces.append("# シート " + str(number) + ": " + title)
        # Empty worksheets contribute a heading, without an extra blank record.
        pieces.extend(projection[sheet.id].rstrip("\n").split("\n") if projection[sheet.id] else [])
        if formats and formats[sheet.id]:
            pieces.append("# 書式 シート " + str(number) + ": " + title)
            pieces.extend(formats[sheet.id].rstrip("\n").split("\n"))
    return ("\ufeff" + "\n".join(pieces) + "\n").encode("utf-8")


def _merge_kind(target, source):
    word_paths = [path for path in (target, source) if _kind(path) == "word"]
    if len(word_paths) == 2:
        return "word"
    if word_paths:
        other = source if _kind(target) == "word" else target
        if Path(other).suffix.lower() in EXCEL_FAMILY:
            raise StageError("Excel と Word は混ぜて取り込めません。同じ種類のファイルを 2 つ指定してください。")
        raise StageError(_TEXTS["word"]["merge_suffix"] + str(other))
    return "excel"


_WORD_WINMERGE_HEADERS = (
    "# 右（取り込み元）から左（取り込み先）へ、取り込む行をコピーする",
    "# 上書き保存（Ctrl+S）してから WinMerge を閉じると、取り込み先の文書に書き込む",
)


def _word_winmerge_text(comparison, projection="baseline"):
    pieces = [*_WORD_WINMERGE_HEADERS]
    for number, story in enumerate(comparison.stories, 1):
        title = json.dumps(story.title, ensure_ascii=False)
        title = title.replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
        pieces.append("# 部分 " + str(number) + ": " + title)
        text = getattr(story, projection)
        pieces.extend(text.rstrip("\n").split("\n") if text else [])
    return ("\ufeff" + "\n".join(pieces) + "\n").encode("utf-8")


def _prepare_word_merge_session(target, source, directory, tool="lazygit"):
    check_target_writable(target)
    if not source.is_file():
        raise StageError("取り込み元は通常のファイルを指定してください: " + str(source))
    if _same_file(target, source):
        raise StageError("取り込み先と取り込み元が同じファイルです。")
    base_data, work_data = target.read_bytes(), source.read_bytes()
    with _word_errors() as module:
        comparison = module.compare(base_data, work_data)
    _save_snapshots(directory, base_data, work_data, ".docx")
    state = {"format": "word", "target": "file", "tool": tool, "path": str(target), "source": str(source),
             "base_sha256": hashlib.sha256(base_data).hexdigest(),
             "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "differences": comparison.differences, "unsupported": _word_unsupported(comparison),
             "selectable": comparison.selectable, "applied": False, "parts": []}
    if tool == "lazygit":
        state["parts"] = _build_word_picker(directory, comparison, _TEXTS["word"]["merge_commit"])
    elif tool == "winmerge":
        (directory / "winmerge").mkdir()
        (directory / "baseline").mkdir()
        left, right = "winmerge/取り込み先.word.txt", "winmerge/取り込み元.word.txt"
        backup = "baseline/取り込み先.word.txt"
        original = _word_winmerge_text(comparison)
        (directory / left).write_bytes(original)
        (directory / right).write_bytes(_word_winmerge_text(comparison, "changed"))
        (directory / backup).write_bytes(original)
        state["parts"] = [{"id": story.id, "title": story.title} for story in comparison.stories]
        state["winmerge"] = {"left": left, "right": right, "baseline": backup,
                             "baseline_sha256": hashlib.sha256(original).hexdigest()}
    else:
        raise StageError("選択画面の道具が不正です: " + tool)
    write_json(directory / "session.json", state)
    return state


def prepare_merge_session(target, source, directory, tool="lazygit"):
    target, source, directory = resolve_user_path(target), resolve_user_path(source), Path(directory).resolve()
    if _merge_kind(target, source) == "word":
        return _prepare_word_merge_session(target, source, directory, tool=tool)
    texts = _TEXTS["excel"]
    for path in (target, source):
        if path.suffix.lower() not in SUPPORTED:
            raise StageError(texts["merge_suffix"] + str(path))
    check_target_writable(target)
    if not source.is_file():
        raise StageError("取り込み元は通常のファイルを指定してください: " + str(source))
    if _same_file(target, source):
        raise StageError("取り込み先と取り込み元が同じファイルです。")
    base_data, work_data = target.read_bytes(), source.read_bytes()
    base, work, projections = _load_books(base_data, work_data)
    projections = _merge_projections(base, work, projections)
    _save_snapshots(directory, base_data, work_data, "." + base.workbook_format)
    state = {"workbook": base.workbook_format, "target": "file", "tool": tool,
             "path": str(target), "source": str(source),
             "base_sha256": hashlib.sha256(base_data).hexdigest(),
             "work_sha256": hashlib.sha256(work_data).hexdigest(),
             "differences": projections.differences, "format_differences": projections.format_differences,
             "selectable": projections.selectable, "unsupported": projections.unsupported, "applied": False, "sheets": []}
    if tool == "lazygit":
        state["sheets"] = _build_picker(directory, base, Projections(*projections[:4]), texts["merge_commit"])
    elif tool == "winmerge":
        (directory / "winmerge").mkdir()
        (directory / "baseline").mkdir()
        left, right = "winmerge/取り込み先.cells.txt", "winmerge/取り込み元.cells.txt"
        backup = "baseline/取り込み先.cells.txt"
        original = _winmerge_text(base, projections.baseline, projections.format_baseline)
        (directory / left).write_bytes(original)
        (directory / right).write_bytes(_winmerge_text(base, projections.changed, projections.format_changed))
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
    texts = _TEXTS[_kind(path)]
    attributes = check_target_writable(path)
    descriptor, temporary_name = tempfile.mkstemp(prefix=texts["replace_prefix"], suffix=path.suffix, dir=path.parent)
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
                    raise PermissionError(texts["open_retry"] + str(path)) from error
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
    if not any(selected):
        return Counts(0, 0)
    rebuilt = base.rebuild(work, selected.values, selected.formats)
    replace_file(Path(state["path"]), rebuilt, state["base_sha256"])
    counts = Counts(len(selected.values), len(selected.formats))
    state["applied"], state["count"], state["format_count"] = True, counts.values, counts.formats
    write_json(directory / "session.json", state)
    return counts


def _write_word_merge(directory, state, comparison, selected):
    if not selected:
        return 0
    with _word_errors():
        rebuilt = comparison.rebuild(selected)
    replace_file(Path(state["path"]), rebuilt, state["base_sha256"])
    state["applied"], state["count"] = True, len(selected)
    write_json(directory / "session.json", state)
    _report_word_swaps(comparison)
    return len(selected)


def apply_merge_session(directory):
    directory = Path(directory).resolve()
    state = _merge_state(directory)
    if state.get("format") == "word":
        comparison = _session_word_comparison(directory, state)
        selected = (winmerge_selected_word(directory, state, comparison) if state["tool"] == "winmerge" else
                    selected_word(directory, state, comparison))
        return _write_word_merge(directory, state, comparison, selected)
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
    if "format_differences" not in state:
        raise StageError("選択画面の形式が更新されています。画面を閉じて選び直してください。")
    details = state["winmerge"]
    original = (directory / details["baseline"]).read_bytes()
    if hashlib.sha256(original).hexdigest() != details["baseline_sha256"]:
        raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
    try:
        saved = _winmerge_lines((directory / details["left"]).read_bytes())
    except UnicodeError as error:
        raise StageError("WinMerge の保存文字コードは UTF-8 にしてください。") from error
    before = _winmerge_lines(original)
    selected, formats, comparison = set(), set(), None
    sheet_index, format_section, last_coordinate = -1, False, "A1"
    missing = object()
    for old, new in zip_longest(before, saved, fillvalue=missing):
        if old is missing or new is missing:
            if isinstance(old, bytes) and old.startswith(b"["):
                last_coordinate = parse_record(old.decode("utf-8"))[0]
            raise StageError("セルの行数が変わっています: " + last_coordinate + "。セルの行全体をコピーしてください。")
        if not old:
            # Blank lines separate format records; like headings, they stay unchanged.
            if new:
                raise StageError("書式の行の間の空行が変更されています: " + last_coordinate + "。空行を元に戻してください。")
            continue
        if old.startswith(b"#"):
            if old != new:
                raise StageError("WinMerge の見出しが変更されています。見出しの行を元に戻してください。")
            if old.startswith("# シート ".encode("utf-8")):
                sheet_index += 1
                format_section = False
            elif old.startswith("# 書式 シート ".encode("utf-8")):
                format_section = True
            continue
        coordinate, cell = parse_record(old.decode("utf-8"))
        last_coordinate = coordinate
        sheet = base.sheets[sheet_index]
        location = sheet.name + "!" + coordinate
        if (cell is not None and cell.kind == FORMAT_KIND) != format_section:
            raise StageError(location + " のセル行の種類が不正です。元の節へ戻してください。")
        if new == old:
            continue
        try:
            saved_coordinate = parse_record(new.decode("utf-8"))[0]
        except WorkbookError:
            saved_coordinate = "不正なセル行"
        invalid = (location + " の選択が不正です（保存行: " + saved_coordinate +
                   "）。取り込み先か取り込み元のセルの行全体をコピーしてください。")
        if format_section:
            if comparison is None:
                comparison = base.compare_formats(work)
            texts = comparison.texts((sheet.id, coordinate))
            if texts is None:
                raise StageError(invalid)
            expected = cell_record(coordinate, Cell(FORMAT_KIND, texts[1]))
        else:
            expected = cell_record(coordinate, work.cells.get((sheet.id, coordinate)))
        # Check only copied records against the snapshots; do not rebuild the
        # complete format projection while applying a large selection.
        expected = expected.replace("\x85", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
        if new != expected.encode("utf-8"):
            raise StageError(invalid)
        if format_section:
            incomplete = location + " の書式の選択が不完全です。書式の行の削除側と追加側を一緒に選んでください。"
            key = _selected_format_key(old.decode("utf-8"), new.decode("utf-8"), sheet.id,
                                       comparison, incomplete, invalid)
            destination = formats
        else:
            key = _selected_key(old.decode("utf-8"), new.decode("utf-8"), sheet.id, base, work,
                                location + " のセル置換が不完全です。", location + " の選択が不完全です。")
            destination = selected
        if key is not None:
            destination.add(key)
    return Selection(selected, formats)


def winmerge_selected_word(directory, state, comparison):
    details = state["winmerge"]
    original = (directory / details["baseline"]).read_bytes()
    if hashlib.sha256(original).hexdigest() != details["baseline_sha256"]:
        raise StageError("選択開始時の射影が変更されています。画面を閉じて選び直してください。")
    if original != _word_winmerge_text(comparison):
        raise StageError("選択開始時の射影が Word 文書と一致しません。画面を閉じて選び直してください。")
    try:
        saved = _winmerge_lines((directory / details["left"]).read_bytes())
    except UnicodeError as error:
        raise StageError("WinMerge の保存文字コードは UTF-8 にしてください。") from error
    before = _winmerge_lines(original)
    after = _winmerge_lines(_word_winmerge_text(comparison, "changed"))
    locations = [None, None]
    for story in comparison.stories:
        locations.append(None)
        with _word_errors() as module:
            locations.extend(story.title + " #" + module.parse_record(line).id.split("/", 1)[0]
                             for line in story.baseline.splitlines())
    if len(saved) != len(before):
        mismatch = next((i for i, (old, new) in enumerate(zip(before, saved)) if old != new),
                        min(len(before), len(saved)))
        nearby = locations[mismatch:] + list(reversed(locations[:mismatch]))
        location = next((value for value in nearby if value is not None), "本文 #1")
        raise StageError("行の数が変わっています: " + location + "。行全体をコピーしてください。")
    chosen = {story.id: [] for story in comparison.stories}
    story_index = -1
    for old, new, source, location in zip(before, saved, after, locations):
        if old.startswith(b"#"):
            if old != new:
                raise StageError("WinMerge の見出しが変更されています。見出しの行を元に戻してください。")
            if old.startswith("# 部分 ".encode("utf-8")):
                story_index += 1
            continue
        if new != old and new != source:
            raise StageError(location + " の選択が不正です。取り込み先か取り込み元の行全体をコピーしてください。")
        chosen[comparison.stories[story_index].id].append(new.decode("utf-8"))
    selected = set()
    with _word_errors():
        for story in comparison.stories:
            selected.update(comparison.choose(story.id, chosen[story.id], projection="changed"))
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
                books = (_session_word_comparison(directory, state) if state.get("format") == "word" else
                         _session_books(directory, state))
            if state.get("format") == "word":
                return winmerge_selected_word(directory, state, books), books
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
    selected, books = _winmerge_select(directory, state)
    while True:
        try:
            if state.get("format") == "word":
                return _write_word_merge(directory, state, books, selected)
            return _write_merge(directory, state, *books, selected)
        except PermissionError as error:
            if not sys.stdin.isatty():
                raise
            print(str(error), file=sys.stderr)
            if input("取り込み先を閉じてから Enter で再試行します（q で中止）: ").strip().lower() == "q":
                raise


def _apply_merge_cli(directory):
    texts = _TEXTS[_session_kind(directory)]
    try:
        count = apply_merge_session(directory)
        print(texts["merged_picker"].format(_count_label(count)) if _counted(count) else
              texts["nothing_selected"])
        return 0
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print(texts["merge_failed"] + str(error), file=sys.stderr)
        return 1


def _unsupported_summary(kind, items):
    if kind == "word":
        examples = ", ".join(item["story_title"] + (" #" + item["slot"] if item["slot"] else "") +
                             "（" + item["reason"] + "）" for item in items[:20])
    else:
        examples = ", ".join(item["sheet"] + "!" + item["coordinate"] + "（" + item["reason"] + "）" for item in items[:20])
    if len(items) > 20:
        examples += _TEXTS[kind]["more"].format(len(items) - 20)
    return examples


def merge_main(argv=None):
    parser = argparse.ArgumentParser(description="2 つの Excel ブックのセル、または 2 つの Word 文書の変更を取り込む")
    parser.add_argument("--tool", choices=("auto", "winmerge", "lazygit"), default="auto")
    parser.add_argument("target")
    parser.add_argument("source")
    args = parser.parse_args(argv)
    kind = _kind(args.target)
    texts = _TEXTS[kind]
    directory = None
    try:
        target, source = resolve_user_path(args.target), resolve_user_path(args.source)
        tool = choose_tool(args.tool)
        if tool == "lazygit":
            check_tools()
        print("書き換えるファイル: " + str(target), flush=True)
        print("取り込み元: " + str(source), flush=True)
        print(texts["reading_merge"], flush=True)
        directory = Path(tempfile.mkdtemp(prefix=texts["merge_prefix"]))
        state = prepare_merge_session(target, source, directory, tool=tool)
        unsupported = state["unsupported"]
        if unsupported:
            print(texts["unsupported_merge"].format(len(unsupported)) + _unsupported_summary(kind, unsupported), flush=True)
        if not state["selectable"]:
            print(texts["only_unsupported"].format(len(unsupported)) if unsupported else
                  texts["no_difference"])
            return 0
        if tool == "winmerge":
            print(texts["winmerge_hint"], flush=True)
            count = _merge_with_winmerge(directory)
            result = 0
        else:
            print(texts["picker_hint"], flush=True)
            result = run_picker(directory, merge=True)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            if not state["applied"]:
                count = 0
            elif kind == "word":
                count = state.get("count", 0)
            else:
                count = Counts(state.get("count", 0), state.get("format_count", 0))
        print(texts["merged"].format(_count_label(count)) + str(target) if _counted(count) else
              texts["merge_nothing"])
        return result
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print(texts["merge_failed"] + str(error), file=sys.stderr)
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
    parser.add_argument("file", nargs="?", help="既存の Excel または Word のファイル")
    parser.add_argument("--apply-session", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--apply-merge-session", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args(arguments)
    texts = _TEXTS[_session_kind(args.apply_session) if args.apply_session is not None else _kind(args.file or "")]
    if args.apply_merge_session is not None:
        return _apply_merge_cli(args.apply_merge_session)
    try:
        if args.apply_session is not None:
            count = apply_session(args.apply_session)
            print(texts["staged"].format(_count_label(count)) if _counted(count) else
                  texts["nothing_selected"])
            return 0
        if not args.file:
            parser.error("Excel か Word のファイルを指定してください")
        check_tools()
        repo = Path(git(Path.cwd(), "rev-parse", "--show-toplevel").stdout.decode().strip())
        prefix = os.environ.get("GIT_PREFIX", "")
        path = str(Path(prefix) / args.file) if prefix else args.file
        directory = Path(tempfile.mkdtemp(prefix=texts["stage_prefix"]))
        try:
            prepare_session(repo, path, directory, progress=lambda message: print(message, flush=True))
            if _kind(path) == "word":
                state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
                unsupported = state["unsupported"]
                if unsupported:
                    print(texts["unsupported_stage"].format(len(unsupported)) +
                          _unsupported_summary("word", unsupported), flush=True)
            result = run_picker(directory)
            state = json.loads((directory / "session.json").read_text(encoding="utf-8"))
            print(texts["stage_done"] if state["applied"] else texts["stage_cancelled"])
            return result
        finally:
            cleanup_session(directory)
    except (StageError, WorkbookError, OSError, UnicodeError, ValueError, KeyboardInterrupt) as error:
        print(texts["stage_failed"] + str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
