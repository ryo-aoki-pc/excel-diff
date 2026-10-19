# AGENTS.md

このリポジトリで作業するコーディングエージェント（Claude Code・Codex・Grok Build・Antigravity CLI）への指示。Claude Code は CLAUDE.md の `@AGENTS.md` で、Codex・Grok Build・Antigravity CLI はこのファイルを直接読む。

利用者への回答・質問・報告は、常に日本語で書く（コードのコメントなどの言語は、このファイルのほかの決まりに従う）。

## 導入手順の正本

導入・更新・使い方の手順は、setup-notes の [almalinux-setup.md](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff) と [windows-setup.md](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff) の「excel-diff」とその子の項にある。このリポジトリには置かない。ロールバックと macOS の導入は、このリポジトリに置く。

## このリポジトリは何か

Excel のブックと Word の文書を、セル・段落の単位のテキストにして Git の差分に表示し、選んだ変更だけを部分ステージする道具。任意の 2 つのブック・文書を比べ、WinMerge または lazygit で選んだ変更を一方へ取り込むこともできる。導入の表は [README.md](README.md)、仕様と検査・macOS の導入は [docs/reference/readme.md](docs/reference/readme.md)、注意点は [docs/extra/readme.md](docs/extra/readme.md)、検証記録は [非公開側の readme.md](https://github.com/ryo-aoki-pc/setup-notes-verification/blob/main/repos/excel-diff/docs/verification/readme.md)。

lazygit の設定は [ryo-aoki-pc/lazygit](https://github.com/ryo-aoki-pc/lazygit) にあり、`E` のキーが、このリポジトリの導入スクリプトで登録する `git excel-stage` を呼ぶ。

ドキュメント・コミットメッセージ・Pull Request の説明は日本語で書く。

### ブランチ構成

`main` が既定のブランチ。作業はここから始める。

## 構成

- `scripts/excel-textconv.py` — Git の textconv で、Excel のセルと Word の段落・表などを差分用のテキストにする
- `scripts/excel-stage.py` — Excel・Word の変更を専用画面で部分ステージし、任意の 2 ファイル間で選んだ変更を取り込む
- `scripts/excel_stage_ooxml.py` — ブックのセルを比較し、選んだセルだけを作り直す
- `scripts/word_ooxml.py` — Word の段落・表・書式などを比較用のテキストにする
- `scripts/word_stage_ooxml.py` — Word の変更を比較し、選んだ段落・セル・表などだけを作り直す
- `scripts/setup-excel-diff.py` — Excel・Word の差分表示・部分ステージ・取り込みを、その端末の Git 全体に設定する
- `requirements-excel-diff.txt` — リポジトリの外の仮想環境に入れる依存関係
- `tests/test_excel_textconv.py` — textconv と Git の差分表示の自動テスト
- `tests/test_excel_stage.py` — Git のインデックスを使った部分ステージの自動テスト
- `tests/test_excel_stage_ooxml.py` — セルの比較とブックの再構成の自動テスト
- `tests/test_excel_merge.py` — 任意の 2 つのブックの取り込みの自動テスト
- `tests/word_fixtures.py` — Word の自動テストに使う文書を作る
- `tests/test_word_ooxml.py` — Word の読み込み・表示・比較用データの自動テスト
- `tests/test_word_stage_ooxml.py` — Word の変更の比較と文書の再構成の自動テスト
- `tests/test_word_textconv.py` — Word の textconv と Git の差分表示の自動テスト
- `tests/test_word_stage.py` — Git のインデックスを使った Word の部分ステージの自動テスト
- `tests/test_word_merge.py` — 任意の 2 つの Word 文書の取り込みの自動テスト
- `tests/test_setup_excel_stage.py` — 導入・更新・Git の alias の自動テスト
- `README.md` — 道具の概要・導入の表・表示例・文書の一覧
- `AGENTS.md` — コーディングエージェントの作業規則
- `CLAUDE.md` — Claude Code から AGENTS.md を読むための指定
- `.gitignore` — Python の実行時キャッシュとローカルの設定・同期マーカーを除く
- `docs/reference/readme.md` — Excel・Word の差分表示・取り込みと Word の部分ステージの仕様、自動検証・macOS の導入
- `docs/extra/readme.md` — Excel・Word の注意点とロールバック

## テスト

```sh
python3 -B -m unittest discover -s tests -v
```

- リポジトリの直下で動かす。Windows は `python`
- openpyxl・xlrd・lxml などの依存は、`scripts/setup-excel-diff.py` が作る仮想環境の Python で動かすと揃う（[docs/reference/readme.md](docs/reference/readme.md) の「自動検証」）

## 文書の置き場所

- 導入・利用・更新の手順は正本に置き、README の「導入」の表から正本の項へ直接リンクする（手順書の先頭やカテゴリの見出しへはリンクしない）。執筆規則は setup-notes の `docs/dev/writing-rules.md`
- `docs/extra/` はロールバックと注意点、`docs/reference/` は仕様・自動検証・macOS の導入。検証記録（実施日・環境・結果・未確認事項）は非公開の `ryo-aoki-pc/setup-notes-verification` の `repos/excel-diff/docs/verification/` に置く。テストのコマンドは上の「テスト」
- 検証記録を書く・直すときは、`setup-notes-verification` を clone し、そちらへ Pull Request を作る
- 文書を移す・分ける・まとめるときは、移した経緯と、旧ファイル名・旧見出しとの対応を文書に書かない（経緯は git の履歴と Pull Request の本文に残す）。リンクを今の場所へ直す。中身が新しい場所への案内だけになったファイルは消す
- 検証していないことを「動く」と書かない
- 残す検証記録の実施日・対象版・環境・出力・成功と失敗・未実施の範囲・記録したときの手順番号は書き換えない。再検証せずに検証範囲を広げない。記録の手順番号が今の手順と違い、本文を読むのに要るときだけ「記録の手順 | 今の手順」の表で結ぶ

## 共同作業の規則

このリポジトリでは、Claude Code・Codex・Grok Build・Antigravity CLI が同じ規則で作業する。分担は人が決める。

- 起動された worktree（作業ディレクトリ）の中だけでファイルを変える。ほかの worktree のファイルは変えない
- 今のブランチにだけコミットする。`main` にはコミットも push もしない
- 頼まれた範囲のファイルだけを変える。範囲の外を変えるときは、変える前に理由を書いて確かめる
- 終わったら、テストとリンターを通してから、目的ごとにコミットする。通らなければコミットせずに、結果を報告する
- コミットしたら、今のブランチを push し、`main` への Pull Request を作る（既にあれば足す）。今のブランチに `main` を取り込むのは、頼まれたときと、Pull Request が競合したときだけ
- `main` への取り込み（マージ）は、エージェントが行ってよい。テストとリンターが通り、Pull Request が競合していなければ、`gh pr merge --squash` で Pull Request を取り込む（手元の `main` にはコミットも push もしない）。利用者が取り込まないよう頼んだときは取り込まず、そのことを報告する。取り込んだブランチには、その後コミットも push もせず、次の作業は新しいブランチで始める
  - リモートのブランチ: 取り込んだ Pull Request のブランチは、GitHub が自動で消す（リポジトリの Automatically delete head branches）。取り込む前に `gh api 'repos/{owner}/{repo}' --jq .delete_branch_on_merge` で確かめ、`false` なら `gh api -X PATCH 'repos/{owner}/{repo}' -F delete_branch_on_merge=true` で有効にする（`{owner}/{repo}` は gh が今のリポジトリで埋めるので、そのまま打つ）。取り込んだ後もリモートに残った自分のブランチは、`git push origin --delete <ブランチ>` で消す
  - 手元のブランチ: 自分が作ったブランチは、Pull Request が取り込まれたこと（`gh pr view <番号> --json state` が `MERGED`）と、どの worktree でも使っていないこと（`git worktree list`）を確かめてから、`git branch -D <ブランチ>` で消す（squash で取り込むので `git branch -d` は通らない）。今いるブランチは、新しいブランチへ移ってから消す。ほかの担当のブランチは消さない
  - Claude Code の worktree とブランチ: セッションの worktree は、Pull Request が閉じるとアプリがセッションをアーカイブして消す（Claude Code の設定の Auto-archive sessions when their pull request closes を有効にしておく。消えることはアプリの設定の説明による。コミットしていない変更がある worktree は残る）。残った `claude/` のブランチは、ほかの担当のブランチを消さない決まりの例外とする。Claude Code は、ほかのセッションが作ったものも、Pull Request が取り込まれたこと（`gh pr list --head <ブランチ> --state merged` に出る）と、どの worktree でも使っていないこと（`git worktree list`）を確かめてから、手元は `git branch -D <ブランチ>`、リモートに残っていれば `git push origin --delete <ブランチ>` で消す
- 秘密情報（`.env`・鍵・トークン・パスワード）を読まない・書かない・出力しない
- レビューを頼まれたら、ファイルを変えずに、指摘を「重大度・場所（ファイル:行）・理由・直し方」で挙げる
- ほかの担当の変更は、`git diff main...agent/codex` のように git で読む（ほかの worktree へ移らない）

## 既定の分担

利用者が Claude Code に、分担を指定せずにファイルを変える作業を頼んだときは、Claude Code が取りまとめて次の 4 つで分担する。「共同作業の規則」の「分担は人が決める」で、人が分担を決めていないときの既定にする。利用者がほかのエージェントに直接頼んだときは、この節は使わず「共同作業の規則」に従う。

- Claude Code だけで行うもの: 利用者の質問への答え、ファイルを変えない調べもの、1 つのファイルの数行で振る舞いを変えない直し（誤字・リンク・値の直しなど）
- Claude Code: 計画・実機での確かめ・ほかの 3 つへの依頼・指摘の確かめ・統合・コミットと Pull Request・取り込み
- Codex: 実装と下書き。Claude Code が用意した別の clone か worktree の中だけで書く。この分担で頼まれたときは、「共同作業の規則」のコミット・push・Pull Request・取り込みを行わず、変更を残して報告する（Claude Code が確かめ、自分の worktree に移してからコミットする）
- Grok Build: 差分の読み取り専用のレビュー。ファイルを変えない（Windows 11 では、Claude Code は `--always-approve` を付けずに頼む）
- Antigravity CLI: 差分の読み取り専用のレビューと、その作業のための調べもの。ファイルを変えない（Claude Code は、入力を写した使い捨てのフォルダーで 1 回だけ実行し、`--dangerously-skip-permissions`・`--mode` を付けない）
- Claude Code から頼まれたエージェントは、上の自分の役割の範囲で作業する
- レビューの指摘は、実物で確かめてから採る。採らなかった指摘は、理由を付けて報告する
- 使えないエージェント（入っていない・ログインが切れた）は飛ばし、そのことを報告する
- 頼み方は、[setup-notes](https://github.com/ryo-aoki-pc/setup-notes) の `docs/ai-agents.md` の「コーディングエージェントの共同作業: Windows 11 で既定で分担する」と、`docs/reference/coding-agents.md` の「既定の分担」
