# Excel の差分表示

[README](../README.md#excel-の差分表示) / [検証記録](verification/readme.md)

Excel はそのままではバイナリ差分になります。[変換スクリプト](../scripts/excel-textconv.py)を Git の `diff.excel.textconv` に登録すると、Git が変更前後の Excel をテキスト化して差分を生成し、lazygit の既存の delta が左右比較で表示します。delta の `--no-gitconfig` は delta 自身の表示設定だけに作用するため、この Git の変換設定は有効です。

## 導入

Python 3.9 以降と Git が必要です。変換キャッシュの保存には、Git のコミット時と同じ `user.name` / `user.email` の設定を使います。lazygit と delta は [README](../README.md#delta-の設定と表示)の手順で導入してください。

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

- リポジトリの外に専用の仮想環境を作り、[openpyxl / xlrd](../requirements-excel-diff.txt)と変換スクリプトを置く。変換結果のキャッシュも同じ場所の `cache` フォルダーに作る
- グローバル Git 設定に `diff.excel.textconv`、`diff.excel.binary=true`、`diff.excel.cachetextconv=false` を登録する([大きいブックの表示](#大きいブックの表示)のため、Git の変換キャッシュは使わない)
- 既存の `core.attributesFile`、未指定なら Git の既定のユーザー属性ファイルに Excel 用の `diff=excel` を追記する。既存の内容は保持し、同じ行を重複追加しない
- グローバル `core.attributesFile` が未指定なら、選んだ属性ファイルの絶対パスを登録する。エディタとシェルで `XDG_CONFIG_HOME` が違う場合も同じ設定を使う。システム設定に既存の属性ファイルがある場合はそのパスを使い、書き込み権限が無ければエラーで止まる

仮想環境の置き場所は Linux では `${XDG_DATA_HOME:-~/.local/share}/lazygit-excel-diff`、macOS では `~/Library/Application Support/lazygit-excel-diff`、Windows では `%LOCALAPPDATA%\lazygit-excel-diff` です。`--install-dir <パス>` で変更できます。Git 設定は各端末の絶対パスを使うため、Syncthing で clone を同期している場合も各端末で導入します。変換処理を更新したら導入コマンドを再実行してください。

`--check` は登録済みの変換コマンドから設置先を調べるので、導入後に `XDG_DATA_HOME` が変わっても現在の設定を確認できます。

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

この差分は閲覧用です。lazygit の**ファイル一覧**で `Space` を押して Excel 全体をステージします。`Enter` で差分に入り、行・ハンクをステージしたり破棄したりしないでください。既存の Excel にはテキストのパッチを適用できず、新規ファイルでは変換後のテキストを Excel の代わりにステージしてしまう場合があります。`diff.excel.binary=true` はバイナリ扱いを維持しますが、lazygit のこれらの操作を禁止する設定ではありません。

自動検証は、このリポジトリのフォルダーで導入した仮想環境の Python を使って実行できます(既定の置き場所の例)。

Linux:

```sh
"${XDG_DATA_HOME:-$HOME/.local/share}/lazygit-excel-diff/venv/bin/python" -B -m unittest discover -s tests -v
```

Windows(PowerShell):

```powershell
& "$env:LOCALAPPDATA\lazygit-excel-diff\venv\Scripts\python.exe" -B -m unittest discover -s tests -v
```
