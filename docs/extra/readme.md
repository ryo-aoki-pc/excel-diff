# excel-diff の注意点とロールバック

[手順書（AlmaLinux 10）](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff-を導入更新する)・[手順書（Windows 11）](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff-を導入更新する)・[検証記録](https://github.com/ryo-aoki-pc/setup-notes-verification/blob/main/repos/excel-diff/docs/verification/readme.md)・[表示する内容と自動検証](../reference/readme.md)

## 注意点

- プロジェクトの `.gitattributes` や `.git/info/attributes` は、ユーザーの属性ファイルより優先する

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

待つ時間は環境変数 `EXCEL_TEXTCONV_TIMEOUT`(秒)で変えられます。`0` にすると変換が終わるまで待ちます。設定する操作は [AlmaLinux 10](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で大きな-excel-の変換を待つ) / [Windows 11](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で大きな-excel-の変換を待つ) を参照してください。

PowerShell の設定は、そのウィンドウで後から実行するコマンドにも効きます。元に戻すときは `Remove-Item Env:EXCEL_TEXTCONV_TIMEOUT` を実行します。

部分ステージの操作と端末から直接開く方法は [AlmaLinux 10](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#lazygit-で-excel-のセル変更を部分ステージする) / [Windows 11](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#lazygit-で-excel-のセル変更を部分ステージする) にあります。
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

画面を開いた後に対象 Excel の index や作業ファイルが変わった場合も、古い選択を適用せず止まります。閉じて `E` から開き直してください。他ファイルのステージ変更は保持します。ブックを読み取って差分を作る処理は完了まで待つため、大きいブックでは専用画面が開くまで時間がかかります。その間は端末に「Excel を読み込んでいます」と表示します。Windows 11 での例では、50 万セルのブックで `E` を押してから専用画面に差分が出るまで約 13 秒、`X` に 13〜18 秒ほどかかり、メモリは最大で 0.9 GB ほど使いました([検証記録](https://github.com/ryo-aoki-pc/setup-notes-verification/blob/main/repos/excel-diff/docs/verification/readme.md#windows-11-での大きいブックの部分ステージの高速化2026-10-17))。

表示確認と部分ステージのために使うスクリプトは、Excel のマクロを実行せず、外部リンクの更新や数式の再計算も行いません。部分ステージしたブックには、次に Excel で開いたときに一度だけ再計算する指定(`fullCalcOnLoad`)を付けます。常に完全再計算する設定(`forceFullCalc`)は付けません。

専用画面の一時リポジトリには `core.quotePath=false` を設定します。lazygit 0.66.0 と delta 0.20.1 では、日本語のシート名から作ったファイル名が Git の引用形式になると、ハンクを選んでも対象ファイルを照合できません。この設定で表示と選択を同じファイル名に揃えます。設定の適用先は一時リポジトリです。

## 取り込みの注意点

`git excel-merge` は、最初の引数のブックを上書きします。`.bak` は作りません。必要な控えは実行前に用意してください。書き込めば更新日時も変わります。取り込み元は準備時の写しを使い、元のファイルは変更しません。

- 画面を開いた後に取り込み先が変わったり削除されたりした場合は、上書きせず止まります。画面を閉じて最初から選び直してください。
- Excel などで取り込み先を開いている場合は閉じてください。Windows は読み書き用に開けないファイル、Linux は LibreOffice のロックファイルがあるファイルで止まります。シンボリックリンク、複数のハードリンクを持つ取り込み先、読み取り専用の取り込み先も対象外です。
- 同期や共同編集のすべてを検出するものではありません。OneDrive などの共同編集中かどうかは検出しません。
- 値と通常の数式だけを取り込みます。書式・画像・マクロ・テンプレート形式は取り込み先のものを保ちます。セル以外の変更を取り込む用途には使えません。
- リッチテキスト、共有・配列数式などの範囲、結合セルの左上以外など、セル単位で反映できない違いは選べません。取り込み先・取り込み元の両方を調べ、端末に理由とセル番地を表示します。
- 行・列の挿入、削除、移動そのものは反映しません。セル番地で比べるため、行・列をずらしたブックは多数の値の変更として見えます。
- 選んだ文字列のふりがなは引き継ぎません。数式はその場で再計算せず、次に Excel で開くときの再計算を指定します。

WinMerge では、取り込むセルの行を右から左へ行ごとにコピーしてください。手で値を打つ、見出しを変える、行を増減・並べ替える操作は照合で止まります。コピーした後は上書き保存（`Ctrl+S`）し、WinMerge の窓を閉じると取り込み先へ書き込みます。保存せずに閉じれば、画面上でコピーしていても取り込み先は変わりません。

WinMerge の比較設定によっては、違いが画面に出ないことがあります。比較の無視設定やフィルターを確かめてください。WinMerge のバックアップ設定によっては、一時フォルダー以外にもセルの値を含むテキストの控えが残ることがあります。扱う値に応じて保存先を確かめてください。

取り込みの準備と書き込みは、ブック全体の読み込み・検証が終わるまで待ちます。大きいブックでは時間とメモリが必要です。Windows 11 での例では、50 万セルのブックで、画面を開くまでの準備に約 6 秒、書き込みに約 11 秒かかり、WinMerge は約 1 秒で窓を開いて 208 MB ほど使いました([検証記録](https://github.com/ryo-aoki-pc/setup-notes-verification/blob/main/repos/excel-diff/docs/verification/readme.md#2-つのブックの取り込み2026-10-18))。

ファイルは同じフォルダーの一時ファイルから置き換えます。元のモードと、Windows の `Zone.Identifier` は引き継ぎます。Windows の明示的な ACL、`Zone.Identifier` 以外の代替データストリームは引き継ぎません。

## Windows 11 のロールバック

excel-diff の Git 設定と導入先を外し、最後に clone を削除します。Git・Python・WinMerge・lazygit は削除しません。

- 前提: [excel-diff を導入・更新する](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md#excel-diff-を導入更新する)で、自分のユーザーに導入している
- 実行場所: Windows 11 の自分のユーザー、管理者ではない Windows PowerShell 5.1
- 貼った後は、反転表示の「確認」から後ろの出力を箇条書きで確かめる

> [!CAUTION]
> この節の手順 3 は `%USERPROFILE%\excel-diff` を削除します。未コミット・未 push の変更は取り戻せません。手順 1 の確認と手順 2 の削除が成功してから進んでください。

1. clone で、未コミット・未 push の変更が無いことを確認する。

   ```powershell
   & {
       Set-Location -LiteralPath "$env:USERPROFILE\excel-diff" -ErrorAction Stop
       git fetch --all --prune
       if ($LASTEXITCODE -ne 0) { throw "中断: リモートを確認できませんでした" }
       $excelDiffChanges = @(git status --porcelain=v1 --untracked-files=all)
       if ($LASTEXITCODE -ne 0) { throw "中断: 作業ツリーを確認できませんでした" }
       $excelDiffUnpushed = @(git log --oneline HEAD --branches --not --remotes)
       if ($LASTEXITCODE -ne 0) { throw "中断: 未 push のコミットを確認できませんでした" }
       "`n$([char]27)[7m 確認 $([char]27)[0m"
       $excelDiffChanges
       $excelDiffUnpushed
       if ($excelDiffChanges.Count -gt 0 -or $excelDiffUnpushed.Count -gt 0) {
           throw "中断: 未コミット・未 push の変更を保管してからやり直してください"
       }
       "未コミット・未 push の変更はありません"
   }
   ```

   - 「未コミット・未 push の変更はありません」と出ること
   - 変更やエラーが出た場合は、手順 2・3 へ進まず、必要なファイルを保管し、コミット・push の扱いを決めてからやり直す
   - **次の手順は、この確認が成功してから貼る**

1. clone で、導入スクリプトの設定と導入先を削除する。

   ```powershell
   & {
       Set-Location -LiteralPath "$env:USERPROFILE\excel-diff" -ErrorAction Stop
       "`n$([char]27)[7m 確認 $([char]27)[0m"
       python scripts/setup-excel-diff.py --uninstall
       if ($LASTEXITCODE -ne 0) { throw "中断: clone を削除せず、表示された理由を確認してください" }
   }
   ```

   - 削除した Git 設定・attributes の区画・導入先、または「消すものがありません」と出て、エラーが無いこと
   - `core.attributesFile` と、attributes の管理区画以外の行は残る
   - 導入先を自動で見つけられない場合は、導入時と同じ `--install-dir` を指定してこの手順をやり直す
   - 管理区画の目印や導入先の確認で止まった場合は、clone を消さず、表示された理由を確認する
   - **次の手順は、この削除が成功してから貼る**

1. ホームへ移り、clone を削除する（取り戻せない）。

   ```powershell
   & {
       $excelDiffClone = Get-Item -LiteralPath "$env:USERPROFILE\excel-diff" -Force -ErrorAction Stop
       if (-not $excelDiffClone.PSIsContainer -or ($excelDiffClone.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
           throw "中断: 通常の clone フォルダーではありません"
       }
       if ($excelDiffClone.FullName -ne [IO.Path]::GetFullPath("$env:USERPROFILE\excel-diff")) {
           throw "中断: 削除先のパスが一致しません"
       }
       Set-Location -LiteralPath $env:USERPROFILE -ErrorAction Stop
       Remove-Item -LiteralPath $excelDiffClone.FullName -Recurse -Force -ErrorAction Stop
       "`n$([char]27)[7m 確認 $([char]27)[0m"
       Test-Path -LiteralPath "$env:USERPROFILE\excel-diff"
   }
   ```

   - `False` と出ること
   - エラーが出た場合は、削除先を広げず、表示されたパスと使用中のファイルを確認する

## AlmaLinux 10 のロールバック

excel-diff の Git 設定と導入先を外し、最後に clone を削除します。Git・Python・lazygit は削除しません。

- 前提: [excel-diff を導入・更新する](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#excel-diff-を導入更新する)で、自分のユーザーに導入している
- 実行場所: AlmaLinux 10 の自分のユーザー、bash
- 貼った後は、反転表示の「確認」から後ろの出力を箇条書きで確かめる

> [!CAUTION]
> この節の手順 3 は `~/excel-diff` を削除します。未コミット・未 push の変更は取り戻せません。手順 1 の確認と手順 2 の削除が成功してから進んでください。

1. clone で、未コミット・未 push の変更が無いことを確認する。

   ```bash
   (
     cd "$HOME/excel-diff" || exit
     git fetch --all --prune || exit
     excel_diff_changes=$(git status --porcelain=v1 --untracked-files=all) || exit
     excel_diff_unpushed=$(git log --oneline HEAD --branches --not --remotes) || exit
     printf '\n\033[7m 確認 \033[0m\n'
     if [ -n "$excel_diff_changes" ] || [ -n "$excel_diff_unpushed" ]; then
       printf '%s\n' "$excel_diff_changes" "$excel_diff_unpushed"
       printf '%s\n' '中断: 未コミット・未 push の変更を保管してからやり直してください' >&2
       exit 1
     fi
     printf '%s\n' '未コミット・未 push の変更はありません'
   )
   ```

   - 「未コミット・未 push の変更はありません」と出ること
   - 変更やエラーが出た場合は、手順 2・3 へ進まず、必要なファイルを保管し、コミット・push の扱いを決めてからやり直す
   - **次の手順は、この確認が成功してから貼る**

1. clone で、導入スクリプトの設定と導入先を削除する。

   ```bash
   (
     cd "$HOME/excel-diff" || exit
     printf '\n\033[7m 確認 \033[0m\n'
     python3 scripts/setup-excel-diff.py --uninstall
   )
   ```

   - 削除した Git 設定・attributes の区画・導入先、または「消すものがありません」と出て、エラーが無いこと
   - `core.attributesFile` と、attributes の管理区画以外の行は残る
   - 導入先を自動で見つけられない場合は、導入時と同じ `--install-dir` を指定してこの手順をやり直す
   - 管理区画の目印や導入先の確認で止まった場合は、clone を消さず、表示された理由を確認する
   - **次の手順は、この削除が成功してから貼る**

1. ホームへ移り、clone を削除する（取り戻せない）。

   ```bash
   (
     if [ -L "$HOME/excel-diff" ] || [ ! -d "$HOME/excel-diff/.git" ]; then
       printf '%s\n' '中断: 通常の clone フォルダーではありません' >&2
       exit 1
     fi
     cd "$HOME" || exit
     rm -rf -- "$HOME/excel-diff" || exit
     printf '\n\033[7m 確認 \033[0m\n'
     if [ ! -e "$HOME/excel-diff" ]; then
       printf '%s\n' 'clone を削除しました'
     else
       printf '%s\n' '中断: clone が残っています' >&2
       exit 1
     fi
   )
   ```

   - 「clone を削除しました」と出ること
   - エラーが出た場合は、削除先を広げず、表示されたパスと使用中のファイルを確認する
