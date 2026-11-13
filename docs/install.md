# excel-diff の導入と更新（AlmaLinux 10・Windows 11）

- Excel のブックと Word の文書の差分表示・部分ステージ・2 つのファイルの取り込みに使う道具（[excel-diff](../README.md)）
- 使い方は [excel-diff の使い方](usage.md) にある
- どの項も単独で通せる。項のリードの前提を先に満たす
- bash のブロックはブラケットペーストの効く端末に、PowerShell のブロックは PowerShell 7 に貼り、貼った後に Enter を押す。貼り方と、反転表示の「確認」から後ろの出力の読み方は、setup-notes の [AlmaLinux 10 のセットアップ](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md) と [Windows 11 のセットアップ](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/windows-setup.md) の冒頭に従う
- macOS の導入は [仕様と検査の「macOS の導入」](reference/readme.md#macos-の導入) にある

<details id="この文書の補足と注意点">
  <summary>この文書の補足と注意点</summary>

**参照**

- 検証記録（setup-notes-verification の `docs/verification/almalinux-setup.md`）
- 検証記録（setup-notes-verification の `docs/verification/windows-setup.md`）

</details>

## AlmaLinux 10 で使う

- GNOME にログインした自分のユーザーで行う。前提の setup-notes の項は、各項のリードのリンクから開く

### excel-diff を導入・更新する

- Excel のブックと Word の文書の差分をテキストで出す textconv と、`git excel-stage`・`git excel-merge`・`git word-stage`・`git word-merge` を Git 全体に入れ、導入状態を確かめる
- 前提: Python 3.9 以上と git が入っている（git は[Homebrew](https://github.com/ryo-aoki-pc/setup-notes/blob/main/docs/almalinux-setup.md#homebrew)の手順 1 で入る）
  - 部分ステージと lazygit での取り込みには、lazygit 0.66.0 以上・delta 0.20.1 以上も要る
- 実行場所: GNOME にログインした自分のユーザーの端末（bash）。`sudo` は、ensurepip で止まったときの `python3-pip` の導入だけで使う
- 更新時は `git -C ~/excel-diff status` で既存の変更を確かめてから進む

1. Python の版を確かめる。

   ```bash
   python3 --version
   ```

   - `Python 3.9` 以上の版が出ればよい
   - 3.9 より古い版やエラーが出たら、その Python を更新してから貼り直す

1. リポジトリを clone し、導入して確かめる。

   ```bash
   # excel-diff のソースを取得する
   git clone https://github.com/ryo-aoki-pc/excel-diff.git ~/excel-diff
   # コマンドを導入して設定を検査する
   (
     cd ~/excel-diff || exit
     printf '\n\033[7m 確認 \033[0m\n'
     python3 scripts/setup-excel-diff.py &&
       python3 scripts/setup-excel-diff.py --check
   )
   ```

   - 導入時に「Excel・Word の差分と部分ステージをグローバルに設定しました。」、確認時に「Excel・Word の差分と部分ステージの設定を確認しました。」が出ればよい
   - その後ろに「変換スクリプト:」「部分ステージ: … (git excel-stage / git word-stage)」「取り込み: git excel-merge / git word-merge」「attributes:」の行が出ればよい
   - 仮想環境の作成が ensurepip で止まる場合は `sudo dnf install python3-pip` を行い、この項の手順 2 の `(`〜`)` 部分を貼り直す
   - 同期したリポジトリでも端末ごとに一度実施し、スクリプトを更新したら再実行する
   - 更新時は `git -C ~/excel-diff pull --ff-only` を実行し、この項の手順 2 の `(`〜`)` 部分（導入・確認）を貼り直す。`git clone` は再実行しない

<details id="excel-diff-を導入更新する-の補足と注意点">
  <summary>excel-diff を導入・更新する の補足と注意点</summary>

**[excel-diff を導入・更新する](#excel-diff-を導入更新する)の手順 2**

- Git 全体の textconv・alias・属性を設定し、依存はリポジトリ外の仮想環境へ入れる

- `--check` は読み取りだけの確認

- 終わると元の作業ディレクトリへ戻る

**[excel-diff を導入・更新する](#excel-diff-を導入更新する)の手順 2: スクリプトを更新した後に導入し直す理由**

- Excel の変換スクリプトの写しは、中身などから決まる名前で Git の設定に書かれる。`scripts/excel-textconv.py` を変える更新を取り込むと、前の設定では確認が失敗する
- [excel-diff を導入・更新する](#excel-diff-を導入更新する)の手順 2 の導入・確認を通し直すと、新しい変換スクリプトの名前で設定される

</details>

---

## Windows 11 で使う

- この PC のデスクトップで、自分のユーザーが行う。前提の setup-notes の項は、各項のリードのリンクから開く

### Windows 11 で excel-diff を導入・更新する

- Excel のブックと Word の文書の差分をテキストで出す textconv と、`git excel-stage`・`git excel-merge`・`git word-stage`・`git word-merge` を Git 全体に入れ、導入状態を確かめる
- 前提: [Git for Windows](https://gitforwindows.org/)と[scoop](https://scoop.sh/)が入っている。部分ステージと lazygit での取り込みには lazygit 0.66.0 以上・delta 0.20.1 以上が要る
- 実行場所: この PC のデスクトップの管理者ではない PowerShell 7。この項の手順 1 で開く
- Python 3.9 以上が要る。この項の手順 2 で、無ければ scoop で入れる
- WinMerge での取り込みには WinMerge が要る。[WinMerge で 2 つの Excel のセルの変更を取り込む](usage.md#winmerge-で-2-つの-excel-のセルの変更を取り込む)の手順 2 か、[WinMerge で 2 つの Word の変更を取り込む](usage.md#winmerge-で-2-つの-word-の変更を取り込む)の手順 3 で、無ければ入れる
- 更新時は `git -C "$env:USERPROFILE\excel-diff" status` で既存の変更を確かめてから進む

1. 管理者ではない PowerShell 7 を開く。

   - スタートメニューで「PowerShell」を探し、クリックして開く（「Windows PowerShell」ではない。「管理者として実行」にはしない）
   - 「PowerShell」が無ければ、[PowerShell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell-on-windows)を入れる

1. Python が無ければ scoop で入れ、版を確かめる。

   ```powershell
   # Python を導入する
   if (-not (Get-Command python -ErrorAction SilentlyContinue)) { scoop install python }
   "`n$([char]27)[7m 確認 $([char]27)[0m"
   python --version
   ```

   - `Python 3.9` 以上の版が出ればよい
   - 3.9 より古い版やエラーが出たら、その Python を更新してから貼り直す

1. リポジトリを clone し、導入して確かめる。

   ```powershell
   # excel-diff のリポジトリを取得する
   git clone https://github.com/ryo-aoki-pc/excel-diff.git "$env:USERPROFILE\excel-diff"
   & {
     # excel-diff の導入と確認を実行する
     Push-Location -LiteralPath "$env:USERPROFILE\excel-diff" -ErrorAction Stop
     try {
       "`n$([char]27)[7m 確認 $([char]27)[0m"
       python scripts/setup-excel-diff.py
       python scripts/setup-excel-diff.py --check
     } finally {
       Pop-Location
     }
   }
   ```

   - 導入時に `Excel・Word の差分と部分ステージをグローバルに設定しました。`、確認時に `Excel・Word の差分と部分ステージの設定を確認しました。` が出ればよい
   - その後ろに `変換スクリプト:`・`部分ステージ: … (git excel-stage / git word-stage)`・`取り込み: git excel-merge / git word-merge`・`attributes:` の行が出ればよい
   - 同期したリポジトリでも端末ごとに一度実施し、スクリプトを更新したら再実行する（[Windows 11 で excel-diff を導入・更新する の補足と注意点](#windows-11-で-excel-diff-を導入更新する-の補足と注意点)の「Windows 11 で excel-diff を導入・更新するの手順 3: スクリプトを更新した後に導入し直す理由」）
   - 更新時は `git -C "$env:USERPROFILE\excel-diff" pull --ff-only` を実行し、この項の手順 3 の `& { ... }` 部分（導入・確認）を貼り直す。`git clone` は再実行しない

<details id="windows-11-で-excel-diff-を導入更新する-の補足と注意点">
  <summary>Windows 11 で excel-diff を導入・更新する の補足と注意点</summary>

**[Windows 11 で excel-diff を導入・更新する](#windows-11-で-excel-diff-を導入更新する)の手順 3**

- Git 全体の textconv・alias・属性を設定し、依存はリポジトリ外の仮想環境へ入れる

- `--check` は読み取りだけの確認

- 終わると元の作業ディレクトリへ戻る

**[Windows 11 で excel-diff を導入・更新する](#windows-11-で-excel-diff-を導入更新する)の手順 3: スクリプトを更新した後に導入し直す理由**

- Excel の変換スクリプトを更新した後に導入し直す理由は、[excel-diff を導入・更新する の補足と注意点](#excel-diff-を導入更新する-の補足と注意点)の「excel-diff を導入・更新するの手順 2: スクリプトを更新した後に導入し直す理由」を参照する

</details>

---
