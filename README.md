# excel-diff

Excel のブックをセル単位のテキストにして Git の差分に表示し、lazygit の `E` で選んだセル変更だけを部分ステージする道具です。

lazygit の設定は [ryo-aoki-pc/lazygit](https://github.com/ryo-aoki-pc/lazygit) にあります。その設定の `E` のキーが、このリポジトリの導入スクリプトで登録する `git excel-stage` を呼びます。

## 導入

AlmaLinux 10 と Windows 11 の自分のユーザーを対象にしています。導入・更新と使い方は、setup-notes の次の項を参照してください。macOS は[仕様と検査の「macOS の導入」](docs/reference/readme.md#macos-の導入)にあります。

| 用途 | AlmaLinux 10 | Windows 11 |
| --- | --- | --- |
| 導入・更新 | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff-を導入更新する) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff-を導入更新する) |
| 差分を確かめる | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-excel-の差分を確かめる) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で-excel-の差分を確かめる) |
| 部分ステージ | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-excel-のセル変更を部分ステージする) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で-excel-のセル変更を部分ステージする) |
| 大きいブックの変換を待つ | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で大きな-excel-の変換を待つ) | [手順](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で大きな-excel-の変換を待つ) |

## 表示の例

売上シートの `B2` を `1000` から `1200` に変えると、Git の差分は次のようになります。delta では削除側と追加側が左右に並びます。

```diff
 Sheet: "売上"
 A1	"金額"
-B2	1000
+B2	1200
 C2	formula: "=B2*1.1"
```

## 文書

- [仕様と検査](docs/reference/readme.md) — 表示する内容・自動検証・macOS の導入
- [注意点](docs/extra/readme.md) — 対応条件と制限
- [検証記録](docs/verification/readme.md) — 実施日・環境・結果・未確認の範囲
- [AGENTS.md](AGENTS.md) — テストのコマンドと、変えるときの決まり
