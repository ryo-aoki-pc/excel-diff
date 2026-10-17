# excel-diff

Excel のブックをセル単位のテキストにして Git の差分に表示し、lazygit の `E` で選んだセル変更だけを部分ステージする道具です。任意の 2 つのブックを比べ、WinMerge または lazygit で選んだセルを一方へ取り込むこともできます。

lazygit の設定は [ryo-aoki-pc/lazygit](https://github.com/ryo-aoki-pc/lazygit) にあります。その設定の `E` のキーが、このリポジトリの導入スクリプトで登録する `git excel-stage` を呼びます。

## 導入

AlmaLinux 10 と Windows 11 の自分のユーザーを対象にしています。導入・更新と使い方は、setup-notes の次の項を参照してください。macOS は[仕様と検査の「macOS の導入」](docs/reference/readme.md#macos-の導入)にあります。

| 用途 | AlmaLinux 10 | Windows 11 |
| --- | --- | --- |
| 導入・更新 | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff-を導入更新する) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff-を導入更新する) |
| 差分を確かめる | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-excel-の差分を確かめる) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で-excel-の差分を確かめる) |
| 部分ステージ | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-excel-のセル変更を部分ステージする) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で-excel-のセル変更を部分ステージする) |
| 2 つのブックの取り込み | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-2-つの-excel-のセルの変更を取り込む) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#winmerge-で-2-つの-excel-のセルの変更を取り込む) |
| 大きいブックの変換を待つ | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で大きな-excel-の変換を待つ) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で大きな-excel-の変換を待つ) |
| 戻す | [手順](docs/extra/readme.md#almalinux-10-のロールバック) | [手順](docs/extra/readme.md#windows-11-のロールバック) |

## 2 つのブックのセルを取り込む

```sh
git excel-merge -- "取り込み先.xlsx" "取り込み元.xlsx"
```

書き換えるのは最初の引数（取り込み先）です。取り込み元は変更しません。Git のリポジトリの外でも使えます。値と通常の数式を取り込み、書式・画像・マクロは取り込み先のものを保ちます。上書き前のバックアップは作りません。

既定の `--tool auto` は、Windows で WinMerge が見つかれば WinMerge、それ以外では lazygit を選びます。`--tool winmerge` または `--tool lazygit` で指定することもできます。

- WinMerge は左が取り込み先、右が取り込み元です。取り込むセルの行を右から左へコピーし、上書き保存（`Ctrl+S`）してから窓を閉じます。
- lazygit は専用画面で `Enter` → `Space` でハンクを選び、`X` で取り込み先へ書き込みます。書き込む前に `q` で閉じると取り消します。

**WinMerge** — 左が取り込み先（見出しに「取り込み先（書き換える）」）、右が取り込み元（読み取り専用）。色の付いた行がセルの違いで、右から左へコピーした行のセルを取り込みます（Windows 11 の WinMerge 2.16.58.2 で、試験用のブックを開いた画面）。

![WinMerge の取り込みの画面。左の見出しが「取り込み先（書き換える）: 見積.xlsx」、右が「取り込み元: 見積-更新.xlsx」で、B2・C3・A4〜D5・D6 などの行が違いとして色付きで並ぶ](docs/images/winmerge-merge.png)

**lazygit** — 左が取り込み先、右が取り込み元のセル。`Enter` → `Space` でステージしたハンクのセルを、`X` で取り込み先へ書き込みます（同じブックを Windows 11 の lazygit 0.66.0・delta 0.20.1 で開いた画面）。

![lazygit の取り込みの専用画面。ファイル一覧に「001-見積.cells.txt」「002-備考.cells.txt」、右の差分に B2 の 10 と 12、C3 の 80 と 90 などが左右に並ぶ](docs/images/lazygit-merge.png)

操作の手順は「導入」の表から、[取り込みの仕様](docs/reference/readme.md#2-つのブックの取り込み)と[注意点](docs/extra/readme.md#取り込みの注意点)は各文書を参照してください。

## 表示の例

売上シートの `B2` を `1000` から `1200` に変えると、Git の差分は次のようになります。delta では削除側と追加側が左右に並びます。

```diff
 Sheet: "売上"
 A1	"金額"
-B2	1000
+B2	1200
 C2	formula: "=B2*1.1"
```

lazygit では、Git のリポジトリの中の変更したブックを選ぶと、セルの差分が delta の左右比較で出ます（Windows 11 の lazygit 0.66.0・delta 0.20.1、試験用のブック）。

![lazygit の Excel の差分の表示。「見積.xlsx」を選ぶと、B2 の 10 と 12、C3 の 80 と 90、D6 の数式の違いなどが左右に並ぶ](docs/images/lazygit-diff.png)

## 文書

- [仕様と検査](docs/reference/readme.md) — 表示する内容・取り込みの仕様・自動検証・macOS の導入
- [注意点とロールバック](docs/extra/readme.md) — 対応条件と制限・取り込みの注意点・戻す手順
- [検証記録](https://github.com/ryo-aoki-pc/setup-notes-verification/blob/main/repos/excel-diff/docs/verification/readme.md)（非公開） — 実施日・環境・結果・未確認の範囲
- [AGENTS.md](AGENTS.md) — テストのコマンドと、変えるときの決まり
