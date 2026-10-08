# Excel の差分表示と部分ステージ

[README](../README.md#excel-の差分表示) / [検証記録](verification/readme.md)

Excel はそのままではバイナリ差分になります。[変換スクリプト](../scripts/excel-textconv.py)を Git の `diff.excel.textconv` に登録すると、Git が変更前後の Excel をテキスト化して差分を生成し、lazygit の既存の delta が左右比較で表示します。delta の `--no-gitconfig` は delta 自身の表示設定だけに作用するため、この Git の変換設定は有効です。

既存の `.xlsx` / `.xlsm` / `.xltx` / `.xltm` のセル値・通常の数式は、ファイル一覧の `E` で開く[専用画面](#excel-の部分ステージ)から部分ステージできます。Git には元の Excel 形式で保存します。

## 導入

Python 3.9 以降と Git が必要です。部分ステージには lazygit 0.66.0 以降と delta 0.20.1 以降が必要です。lazygit と delta は [README](../README.md#delta-の設定と表示)の手順で導入してください。

このリポジトリのフォルダーで実行します。

Linux / macOS:

```sh
python3 --version
python3 scripts/setup-excel-diff.py
python3 scripts/setup-excel-diff.py --check
```

Windows(PowerShell):

```powershell
python --version
python scripts/setup-excel-diff.py
python scripts/setup-excel-diff.py --check
```

Windows で Python が無い場合は `scoop install python` で導入できます。Linux では Python の `venv` と `pip` が使える必要があります。仮想環境作成時に `ensurepip` のエラーが出る場合は、その OS の Python 用パッケージを導入してから再実行します(Ubuntu なら `sudo apt install python3-venv`、AlmaLinux なら `sudo dnf install python3-pip`)。

導入スクリプトは次を設定します。

- リポジトリの外に専用の仮想環境を作り、[openpyxl / xlrd / lxml](../requirements-excel-diff.txt)、変換スクリプト、部分ステージ用のスクリプトを置く。変換結果のキャッシュも同じ場所の `cache` フォルダーに作る
- グローバル Git 設定に `diff.excel.textconv`、`diff.excel.binary=true`、`diff.excel.cachetextconv=false` を登録する([大きいブックの表示](#大きいブックの表示)のため、Git の変換キャッシュは使わない)
- グローバル Git 設定に `alias.excel-stage` を登録し、`git excel-stage -- <ファイル>` で部分ステージ画面を開けるようにする。同梱の `config.yml` の `E` キーもこのコマンドを使う
- 既存の `core.attributesFile`、未指定なら Git の既定のユーザー属性ファイルに Excel 用の `diff=excel` を追記する。既存の内容は保持し、同じ行を重複追加しない
- グローバル `core.attributesFile` が未指定なら、選んだ属性ファイルの絶対パスを登録する。エディタとシェルで `XDG_CONFIG_HOME` が違う場合も同じ設定を使う。システム設定に既存の属性ファイルがある場合はそのパスを使い、書き込み権限が無ければエラーで止まる

仮想環境の置き場所は Linux では `${XDG_DATA_HOME:-~/.local/share}/lazygit-excel-diff`、macOS では `~/Library/Application Support/lazygit-excel-diff`、Windows では `%LOCALAPPDATA%\lazygit-excel-diff` です。`--install-dir <パス>` で変更できます。Git 設定は各端末の絶対パスを使うため、Syncthing で clone を同期している場合も各端末で導入します。変換・部分ステージ処理を更新したら導入コマンドを再実行してください。

`--check` は登録済みの変換コマンドから設置先を調べるので、導入後に `XDG_DATA_HOME` が変わっても現在の設定を確認できます。依存ライブラリ、設置済みの変換・部分ステージスクリプト、Git の alias と属性設定を読み取りだけで確認します。

## 表示する内容

| 形式 | 表示内容 |
| --- | --- |
| `.xlsx` / `.xlsm` / `.xltx` / `.xltm` | 全シートのセル値、セル番地、数式。数式は式そのものを表示し、再計算しない |
| `.xls` / `.xlt` | 全シートのセル値、セル番地。数式セルはファイルに保存された計算結果を表示する |

拡張子の大文字・小文字はどちらも対象です。Git が過去の版を拡張子の無い一時ファイルに取り出す場合も、内容から形式を判定して変換します。

例えば、売上シートの B2 を `1000` から `1200` に変えると、Git の差分は次のようになります。delta では削除側と追加側が左右に並びます。

```diff
 Sheet: "売上"
 A1	"金額"
-B2	1000
+B2	1200
 C2	formula: "=B2*1.1"
```

空のセルは省略します。文字列中の改行やタブは `\n` / `\t` として表示し、1 セルを 1 行で比較します。書式・色・図・マクロのコードは比較しません。`.xlsb` とパスワードで暗号化されたブックは対象外です。

## 大きいブックの表示

変換に時間がかかっても、lazygit の操作は止まりません。1 回の変換が 2 秒以内に終わらない場合、変換スクリプトはすぐに次のような仮の内容を返し、残りの変換を裏で続けます。シートを展開した大きさから 2 秒では終わらないと分かるブック(展開後 16 MB 超。検証に使った表では約 30 万セル)は、待たずにすぐ返します。

```diff
+# Excel を裏で変換中 (1a2b3c4d5e6f)
+# 完了後に選び直すか R で表示
 Sheet: "売上"
```

lazygit はファイル一覧を 10 秒ごとに更新する(`refresher.refreshInterval` の既定値)ため、変換が終われば選んだままでもセルの差分に切り替わります。すぐ見たい場合は、ファイルを選び直すか `R` を押します。比較する 2 つの版のうち片方だけが変換済みの場合は、一時的にその版の全セルを削除(または追加)したように表示されます。先頭の「Excel を裏で変換中」の行で見分けてください。

Windows の lazygit 0.66.0 は、別のファイルへ移っても実行中の差分コマンドを止められません。待ち時間を区切らないと、離れたブックの変換が裏で重なって CPU を使い続けるため、次のようにしています。

- 裏の変換は 1 つずつ、低い優先度で実行する。最後に表示しようとしたブックから変換し、1 時間以上表示されなかった依頼は捨てる
- 変換結果はブックの内容(SHA-256)ごとにキャッシュし、同じ内容は 2 回目から変換しない。キャッシュは合計 512 MB を超えた分と、30 日使わなかった分から消える。導入スクリプトを再実行すると、古い版の変換スクリプトのキャッシュも消す
- Git の変換キャッシュ(`diff.excel.cachetextconv`)は仮の内容まで保存してしまうため使わない

待つ時間は環境変数 `EXCEL_TEXTCONV_TIMEOUT`(秒)で変えられます。`0` にすると変換が終わるまで待ちます。コマンドラインの `git diff` で必ず全セルを表示する場合は次のように実行します。

```sh
EXCEL_TEXTCONV_TIMEOUT=0 git diff -- "売上表.xlsx"
```

```powershell
$env:EXCEL_TEXTCONV_TIMEOUT = "0"; git diff -- "売上表.xlsx"
```

PowerShell の設定は、そのウィンドウで後から実行するコマンドにも効きます。元に戻すときは `Remove-Item Env:EXCEL_TEXTCONV_TIMEOUT` を実行します。

## 確認と操作

Excel のある Git リポジトリで、実際のファイル名を指定します。

```sh
git check-attr diff -- "売上表.xlsx"
git diff -- "売上表.xlsx"
git diff --cached -- "売上表.xlsx"
```

最初のコマンドが `売上表.xlsx: diff: excel` を返し、差分に `Sheet:` とセル番地が出れば設定できています。lazygit を開き直してファイルを選ぶと、同じ内容を delta で表示します。`|` で内蔵表示に切り替えた場合も Excel のテキスト差分を表示します。

プロジェクトの `.gitattributes` や `.git/info/attributes` はユーザー属性ファイルより優先します。`git check-attr` が `unset` や別のドライバーを返す場合は、そのリポジトリで Excel に指定している属性を確認してください。そのリポジトリだけに適用するには、`.git/info/attributes` に例えば `*.[xX][lL][sS][xX] diff=excel` を追加できます。

通常の Excel 差分は閲覧用です。lazygit の**ファイル一覧**で `Space` を押すと Excel 全体をステージします。通常の差分に `Enter` で入り、行・ハンクをステージしたり破棄したりしないでください。既存の Excel にはテキストのパッチを適用できず、新規ファイルでは変換後のテキストを Excel の代わりにステージしてしまう場合があります。`diff.excel.binary=true` はバイナリ扱いを維持しますが、lazygit のこれらの操作を禁止する設定ではありません。部分ステージは、次の `E` キーの専用画面から行います。

## Excel の部分ステージ

1. lazygit のファイル一覧で、変更した既存の `.xlsx` / `.xlsm` / `.xltx` / `.xltm` を選び、`E` を押します。
2. 専用画面で `Enter` を押して delta の左右比較へ入り、`Space` で必要なハンクを選択します。左右の矢印でハンクを移動できます。選択した内容は専用画面のステージ済み差分で確認できます。
3. 選択を確認して `X` を押すと、選択したセルの変更を元のリポジトリへステージします。適用結果を確認したら `q` で元の画面へ戻ります。`X` はこの専用画面だけの操作です。適用前に `q` で終了すれば、元のリポジトリは変更せず選択を取り消します。

端末から直接開くこともできます。

```sh
git excel-stage -- "売上表.xlsx"
```

対象は HEAD と index の両方に既存の Excel があるファイルです。すでにステージ済みの変更がある場合は、その版を基準に、残りのセル変更を追加できます。新規ファイルは一度全体をコミットしてから部分ステージの対象になります。作業中の Excel は書き換えません。ファイル全体をステージする場合は、元の lazygit のファイル一覧で `Space` を使ってください。

専用画面ではセルの値と通常の数式を比較し、index 側の書式・画像・マクロ・テンプレート形式を保持します。数式の計算結果キャッシュや、書式だけの変更は選択対象にしません。通常の閲覧用 textconv は表示形式に応じて日付・時刻を表示するため、専用画面とは値の表示が異なる場合があります。

セル以外の書式・画像・マクロの変更は選択対象にせず、元の index の内容を保持します。行列の挿入・削除・移動を操作ごと反映することもできません。こうした変更を含むブックは全体のステージを使ってください。

日本語版 Excel が文字列に付けるふりがな(読み)は、文字列の一部として扱わず比較しません。ふりがな付きのセルも部分ステージできますが、選んだセルには新しい値の文字だけを書き込み、ふりがなは引き継ぎません。未選択のセルのふりがなはそのまま残ります。

次の場合は、元の index を変更せずエラーで止まります。

- 新規ファイル、削除・名前変更、競合中のファイル、`.xls` / `.xlt` / `.xlsb`、暗号化・電子署名付きのブック
- シートの追加・削除・名前変更、日付基準・名前定義・外部参照の変更
- 配列数式・共有数式・データテーブルなど、セル同士の依存関係を安全に反映できない変更
- 文字ごとに書式を変えたリッチテキストなど対応できないセル内容や、結合セルの左上以外のセルの変更
- 値や数式の置換で、削除行と追加行の片側だけを選択した場合。対応する両方の行を選択してください

画面を開いた後に対象 Excel の index や作業ファイルが変わった場合も、古い選択を適用せず止まります。閉じて `E` から開き直してください。他ファイルのステージ変更は保持します。ブックを読み取って差分を作る処理は完了まで待つため、大きいブックでは専用画面が開くまで時間がかかります。その間は端末に「Excel を読み込んでいます」と表示します。Windows 11 での例では、50 万セルのブックで専用画面が開くまで約 16 秒、`X` に約 22 秒かかり、最大で 2.5 GB ほどのメモリを使いました([検証記録](verification/readme.md#windows-11-での-excel-の部分ステージの検証2026-10-09))。

表示確認と部分ステージのために使うスクリプトは、Excel のマクロを実行せず、外部リンクの更新や数式の再計算も行いません。部分ステージしたブックには、次に Excel で開いたときに一度だけ再計算する指定(`fullCalcOnLoad`)を付けます。常に完全再計算する設定(`forceFullCalc`)は付けません。

専用画面の一時リポジトリには `core.quotePath=false` を設定します。lazygit 0.66.0 と delta 0.20.1 では、日本語のシート名から作ったファイル名が Git の引用形式になると、ハンクを選んでも対象ファイルを照合できません。この設定で表示と選択を同じファイル名に揃えます。設定の適用先は一時リポジトリです。

## 自動検証

自動検証は、このリポジトリのフォルダーで導入した仮想環境の Python を使って実行できます(既定の置き場所の例)。

Linux:

```sh
"${XDG_DATA_HOME:-$HOME/.local/share}/lazygit-excel-diff/venv/bin/python" -B -m unittest discover -s tests -v
```

Windows(PowerShell):

```powershell
& "$env:LOCALAPPDATA\lazygit-excel-diff\venv\Scripts\python.exe" -B -m unittest discover -s tests -v
```
