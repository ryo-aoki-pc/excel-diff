# AGENTS.md

このリポジトリで作業するコーディングエージェント（Claude Code・Codex・Grok Build・Antigravity CLI）への指示。Claude Code は CLAUDE.md の `@AGENTS.md` で、Codex・Grok Build・Antigravity CLI はこのファイルを直接読む。

## 導入手順の正本

導入・更新・使い方の手順は、setup-notes の [almalinux-setup.md](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff) と [windows-setup.md](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff) の「excel-diff」とその子の項にある。このリポジトリには置かない。macOS の導入は、このリポジトリに置く。

## このリポジトリは何か

Excel のブックと Word の文書を、セル・段落の単位のテキストにして Git の差分に表示し、選んだ変更だけを部分ステージする道具。任意の 2 つのブック・文書を比べ、WinMerge または lazygit で選んだ変更を一方へ取り込むこともできる。導入の表は [README.md](README.md)、仕様と検査・macOS の導入は [docs/reference/readme.md](docs/reference/readme.md)、注意点は [docs/extra/readme.md](docs/extra/readme.md)、検証記録は `repos/excel-diff/docs/verification/readme.md`。

lazygit の設定は [ryo-aoki-pc/lazygit](https://github.com/ryo-aoki-pc/lazygit) にあり、`E` のキーが、このリポジトリの導入スクリプトで登録する `git excel-stage` を呼ぶ。

ドキュメント・コミットメッセージ・Pull Request の説明は日本語で書く。

### ブランチ構成

`main` が既定のブランチ。作業はここから始める。

## 構成

- `scripts/excel-textconv.py` — Git の textconv で、Excel のセルと Word の段落・表などを差分用のテキストにする
- `scripts/excel-stage.py` — Excel・Word の変更を専用画面で部分ステージし、任意の 2 ファイル間で選んだ変更を取り込む
- `scripts/excel_stage_ooxml.py` — xlsx 系のブックのセルの値・数式と書式を比較し、選んだ内容だけを作り直す
- `scripts/excel_formula.py` — BIFF8・BIFF12 の数式のトークン列を式の文字列に戻す
- `scripts/excel_xls.py` — xls・xlt の BIFF8 の数式を差分表示用に読む
- `scripts/excel_xlsb.py` — xlsb の記録・セル・数式などを読み、差分表示用のテキストにする
- `scripts/excel_stage_xlsb.py` — xlsb のセルの値・数式と書式を比較し、選んだ内容だけを作り直す
- `scripts/excel_stage_xlsb_styles.py` — xlsb の styles.bin を比較し、必要な書式の定義を追記する
- `scripts/word_ooxml.py` — Word の段落・表・書式などを比較用のテキストにする
- `scripts/word_stage_ooxml.py` — Word の変更を比較し、選んだ段落・セル・表などだけを作り直す
- `scripts/setup-excel-diff.py` — Excel・Word の差分表示・部分ステージ・取り込みを、その端末の Git 全体に設定する
- `requirements-excel-diff.txt` — リポジトリの外の仮想環境に入れる依存関係
- `tests/test_excel_textconv.py` — textconv と Git の差分表示の自動テスト
- `tests/test_excel_stage.py` — Git のインデックスを使った値・数式と書式の部分ステージの自動テスト
- `tests/test_excel_stage_ooxml.py` — セルの値・数式と書式の比較、ブックの再構成の自動テスト
- `tests/test_excel_merge.py` — 任意の 2 つのブックの値・数式と書式の取り込みの自動テスト
- `tests/test_excel_formula.py` — BIFF8・BIFF12 の数式の解読の自動テスト
- `tests/test_xls_textconv.py`・`tests/test_xlsb_textconv.py` — xls・xlsb の数式の表示と xlsx との突き合わせの自動テスト
- `tests/test_excel_stage_xlsb.py` — xlsb のセルの値・数式の比較と作り直し、部分ステージと取り込みの自動テスト
- `tests/test_excel_stage_xlsb_styles.py` — xlsb の書式の比較・追記・取り込みの自動テスト
- `tests/test_excel_stage_xlsb_performance.py`・`tests/test_xlsb_record_views.py` — 読み込み・検証の処理と、位置から参照する記録の自動テスト
- `tests/test_xlsb_formats_oracle.py` — Excel が作った xlsx・xlsb の書式の説明の突き合わせ
- `tests/xlsb_builder.py`・`tests/xlsb_style_builder.py` — BIFF12 の記録・数式・書式と試験用のパッケージを組み立てる
- `tests/fixtures/` — Excel で作った試験用のブックと、作成用の `make_fixtures.ps1`
- `tests/excel_open.ps1`・`tests/test_excel_open.py` — Windows の Excel で読み取り結果と作り直した xlsb を確かめる
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
- `docs/extra/readme.md` — Excel・Word の注意点

## テスト

```sh
python3 -B -m unittest discover -s tests -v
```

- リポジトリの直下で動かす。Windows は `python`
- openpyxl・xlrd・lxml などの依存は、`scripts/setup-excel-diff.py` が作る仮想環境の Python で動かすと揃う（[docs/reference/readme.md](docs/reference/readme.md) の「自動検証」）

## 文書の置き場所

- 導入・利用・更新の手順は正本に置き、README の「導入」の表から正本の項へ直接リンクする（手順書の先頭やカテゴリの見出しへはリンクしない）。執筆規則は setup-notes の `docs/dev/writing-rules.md`
- `docs/extra/` は注意点。ロールバック（取り消し）の手順は書かない。`docs/reference/` は仕様・自動検証・macOS の導入。検証記録（実施日・環境・結果・未確認事項）は Forgejo の非公開リポジトリ `setup-notes-verification` の `repos/excel-diff/docs/verification/` に置く。テストのコマンドは上の「テスト」
- 検証記録を書く・直すときは、Forgejo の非公開リポジトリ `setup-notes-verification` を clone し、そちらへ Pull Request を作る。このリポジトリから検証記録へは URL で指さず、「`<検証記録のパス>` の「<見出し>」」と文字で書く（見出しを指さないときは、同じ名前の文書の記録なら「検証記録」とだけ書き、ほかの記録はパスを添える）
- 文書を移す・分ける・まとめるときは、移した経緯と、旧ファイル名・旧見出しとの対応を文書に書かない（経緯は git の履歴と Pull Request の本文に残す）。リンクを今の場所へ直す。中身が新しい場所への案内だけになったファイルは消す
- 検証していないことを「動く」と書かない
- 残す検証記録の実施日・対象版・環境・出力・成功と失敗・未実施の範囲・記録したときの手順番号は書き換えない。再検証せずに検証範囲を広げない。記録の手順番号が今の手順と違い、本文を読むのに要るときだけ「記録の手順 | 今の手順」の表で結ぶ
