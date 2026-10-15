# Excel の差分表示の仕様と検査

導入と使い方の手順は setup-notes の[lazygit の Excel 差分表示を導入する](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-の-excel-差分表示を導入する)（Windows 11 は[同名の項](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-の-excel-差分表示を導入する)）と、その後ろの項にある。この文書には、表示する内容・自動検証・macOS の導入を置く。対応条件と制限は[注意点](extra/excel-diff.md)。

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

## macOS の導入

リポジトリのフォルダーで `python3 scripts/setup-excel-diff.py`、確認は `python3 scripts/setup-excel-diff.py --check` を行う。
