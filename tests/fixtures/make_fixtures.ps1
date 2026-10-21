<#
試験用の xlsb・xlsx・xls・xlam を Excel で作る。Windows の Excel（COM）が要る。

    pwsh -File tests/fixtures/make_fixtures.ps1 -OutDir C:\tmp\excel-diff-fixtures

- 自分で起動した Excel だけを使い、最後に閉じる。EXCEL.EXE が既に動いていれば止める
- 作成者・最終更新者は「試験 太郎」、日時は 2026-01-01 にする。Excel が書く利用者名
  （Application.UserName）・Windows のユーザー名・ホームのパスが残っていないことを確かめる
- 保存先のフォルダー名はブックの中（BrtAbsPath15）に残るので、利用者名を含まないパスにする
#>
param(
    [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3

$author = '試験 太郎'
$OutDir = [IO.Path]::GetFullPath($OutDir)
if ($OutDir -match [regex]::Escape($env:USERNAME) -or $OutDir.StartsWith($env:USERPROFILE, [StringComparison]::OrdinalIgnoreCase)) {
    throw "中断: 保存先のパスに利用者名が入っています。C:\tmp の下などを指定してください: $OutDir"
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw '中断: Excel が動いています。閉じてからやり直してください。'
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

function Set-Property($object, [string]$name, $value) {
    [void][System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::SetProperty, $null, $object, @($value))
}

function Get-Item2($object, $index) {
    [System.__ComObject].InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $object, @($index))
}

function Set-Author($book) {
    $properties = $book.BuiltinDocumentProperties
    foreach ($name in 'Author', 'Last Author') {
        Set-Property (Get-Item2 $properties $name) 'Value' $author
    }
}

function Save-As($book, [string]$name, [int]$format) {
    Set-Author $book
    $path = Join-Path $OutDir $name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
    $book.SaveAs($path, $format)
    $path
}

function Find-BuiltinStyle($book, [string[]]$names) {
    foreach ($style in $book.Styles) {
        if ($style.BuiltIn -and $names -contains $style.NameLocal) { return $style.Name }
    }
    throw "中断: 組み込みのスタイルが見つかりません: $($names -join ', ')"
}

function Build-Sales($book) {
    while ($book.Worksheets.Count -lt 2) {
        [void]$book.Worksheets.Add([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    }
    $sales = $book.Worksheets.Item(1); $sales.Name = '売上'
    $notes = $book.Worksheets.Item(2); $notes.Name = '備考'
    $notes.Range('A1').Value2 = 'メモ'
    $notes.Range('B1').Value2 = 0.1
    $notes.Range('B2').Value2 = 2
    $notes.Range('B3').Value2 = 3
    [void]$book.Names.Add('税率', '=備考!$B$1')

    $sales.Range('A1:E1').Value2 = [object[]]@('品目', '数量', '単価', '金額', 'メモ')
    $sales.Range('A2').Value2 = 'りんご'; $sales.Range('B2').Value2 = 10; $sales.Range('C2').Value2 = 120
    $sales.Range('A3').Value2 = 'みかん'; $sales.Range('B3').Value2 = 5; $sales.Range('C3').Value2 = 80
    $sales.Range('A4').Value2 = 'ぶどう'; $sales.Range('B4').Value2 = 2; $sales.Range('C4').Value2 = 450
    $sales.Range('D2').Formula = '=B2*C2'
    [void]$sales.Range('D2:D4').FillDown()
    $sales.Range('A5').Value2 = '合計'
    $sales.Range('B5').Formula = '=SUM(B2:B4)'
    $sales.Range('D5').Formula = '=SUM(D2:D4)'

    $sales.Range('E2').Value2 = '重要な点'
    $sales.Range('E2').Characters(1, 2).Font.Bold = $true
    $sales.Range('E3').Value2 = '売上'
    $sales.Range('E3').SetPhonetic()
    $sales.Range('E4').Value2 = "改行`nあり`tタブ"

    $sales.Range('F2').Value2 = $true
    $sales.Range('F3').Value2 = $false
    $sales.Range('G2').Formula = '=NA()'
    Set-Property $sales.Range('G3') 'Value2' ([Runtime.InteropServices.ErrorWrapper]::new(-2146826281))  # #DIV/0!

    $sales.Range('H2').Formula = '=DATE(2026,10,21)'
    $sales.Range('H2').Value2 = $sales.Range('H2').Value2
    $sales.Range('H2').NumberFormat = 'yyyy/m/d'
    $sales.Range('H3').Value2 = 0.5729166666666666
    $sales.Range('H3').NumberFormat = 'h:mm'
    $sales.Range('H4').Value2 = 46316.75
    $sales.Range('H4').NumberFormat = 'yyyy/m/d h:mm'
    $sales.Range('H5').Value2 = 0.125
    $sales.Range('H5').NumberFormat = '0.0%'

    $formulas = [ordered]@{
        'I2'  = '=IF(D2>1000,"高い","安い")'
        'I3'  = '=ROUND(C3*1.08,0)'
        'I4'  = '=VLOOKUP("みかん",A2:C4,3,FALSE)'
        'I5'  = '=TEXT(H2,"yyyy年m月d日")'
        'I6'  = '="a"&"b"&1'
        'I7'  = '=-B2+B3^2/4'
        'I8'  = '=B2>=B3'
        'I9'  = '=備考!B2*2'
        'I10' = '=税率*D2'
        'I11' = '=SUM(備考!B1:B3)'
        'I12' = '=COUNTIF(B2:B4,">3")'
        'I13' = '=IFERROR(1/0,"ゼロ")'
        'I14' = '=AVERAGE(B2:B4)%'
        'I15' = '=$B$2+B$3+$C4'
        'I16' = '=SUM(B2:B4,C2:C4,1)'
        'I17' = '=TRUE'
        'I18' = '=#N/A'
        'I19' = '=MAX({1,2;3,4})'
        'I20' = '=CONCATENATE(A2,"・",A3)'
        'I21' = '=(B2+B3)*(C2-C3)'
    }
    foreach ($key in $formulas.Keys) { $sales.Range($key).Formula = $formulas[$key] }
    $sales.Range('J2:J3').FormulaArray = '=B2:B3*C2:C3'

    $sales.Range('K1').Value2 = '結合'
    $sales.Range('K1:L1').Merge()
    $sales.Range('AMK1').Value2 = '列のブロック 2'
    $sales.Range('A40').Value2 = '40 行目'

    $sales.Range('A1:E1').Font.Bold = $true
    $sales.Range('A1:E1').Interior.Color = 0x00FFFF
    $sales.Range('C2:C4').NumberFormat = '#,##0"円"'
    $sales.Range('D5').Borders.Item(8).LineStyle = 1   # xlEdgeTop
    $sales.Range('D5').Borders.Item(9).LineStyle = -4119  # xlEdgeBottom, xlDouble
    $sales.Range('B2').Font.Color = 0x0000FF
    $sales.Range('A4').Style = Find-BuiltinStyle $book @('良い', 'Good')
    $sales.Range('M2').Interior.Color = 0xCCFFCC
}

function Edit-Sales($book) {
    $sales = $book.Worksheets.Item('売上')
    $notes = $book.Worksheets.Item('備考')
    $sales.Range('B2').Value2 = 12
    $sales.Range('C3').Value2 = 90
    $sales.Range('D4').Formula = '=B4*C4*2'
    $sales.Range('I2').Formula = '=IF(D2>1000,"とても高い","安い")'
    $sales.Range('E4').Value2 = "改行`nなし`tタブ"
    $sales.Range('A6').Value2 = 'もも'
    $sales.Range('AMK3').Value2 = 7
    $sales.Range('F3').ClearContents()
    $sales.Range('B2').Font.Bold = $true
    $sales.Range('A1').Interior.Color = 0xFFCC99
    $sales.Range('H2').Value2 = $sales.Range('H2').Value2 + 1
    $sales.Range('A50').Value2 = '末尾'
    $sales.Range('E5').Value2 = '新しい文字列'
    $sales.Range('M3').Interior.Color = 0x99CCFF
    $notes.Range('B2').Value2 = 20
}

function Build-ManyRows($book) {
    $sheet = $book.Worksheets.Item(1); $sheet.Name = 'Data'
    foreach ($r in 40..200) {
        if ($r % 7 -eq 0) { continue }
        $sheet.Cells.Item($r, 1).Value2 = [double]$r
        if ($r % 3 -eq 0) { $sheet.Cells.Item($r, 1030).Value2 = "x$r" }
        if ($r % 5 -eq 0) { $sheet.Cells.Item($r, 3).Interior.Color = 255 }
        if ($r % 11 -eq 0) { $sheet.Cells.Item($r, 5000).Formula = "=A$r*2" }
    }
}

function Build-FormatsBase($book) {
    $sheet = $book.Worksheets.Item(1); $sheet.Name = 'Fmt'
    foreach ($r in 1..30) {
        $sheet.Cells.Item($r, 1).Value2 = "行$r"
        $sheet.Cells.Item($r, 2).Value2 = [double]$r
    }
}

function Edit-Formats($book) {
    # One kind of format per cell, so each description can be compared with the xlsx.
    $s = $book.Worksheets.Item('Fmt')
    $s.Range('A1').Font.Italic = $true
    $s.Range('A2').Font.Underline = 2        # single
    $s.Range('A3').Font.Underline = -4119    # double
    $s.Range('A4').Font.Strikethrough = $true
    $s.Range('A5').Font.Superscript = $true
    $s.Range('A6').Font.Subscript = $true
    $s.Range('A7').Font.Size = 16
    $s.Range('A8').Font.Name = 'Arial'
    $s.Range('A9').Font.ThemeColor = 5; $s.Range('A9').Font.TintAndShade = 0.4
    $s.Range('A10').Interior.ThemeColor = 6; $s.Range('A10').Interior.TintAndShade = -0.25
    $s.Range('A11').Interior.Pattern = 17; $s.Range('A11').Interior.PatternColor = 255
    $s.Range('A12').Interior.Pattern = 4000  # linear gradient
    $gradient = $s.Range('A12').Interior.Gradient; $gradient.Degree = 90
    $gradient.ColorStops.Clear(); [void]$gradient.ColorStops.Add(0); [void]$gradient.ColorStops.Add(1)
    $gradient.ColorStops.Item(1).Color = 0xFFFFFF; $gradient.ColorStops.Item(2).Color = 0x0000FF
    $s.Range('A13').Borders.Item(5).LineStyle = 1   # diagonal down
    $s.Range('A14').Borders.Item(6).LineStyle = 1   # diagonal up
    $s.Range('A15').Borders.Item(7).Weight = -4138; $s.Range('A15').Borders.Item(7).LineStyle = 1
    $s.Range('A16').Borders.Item(9).LineStyle = -4115; $s.Range('A16').Borders.Item(9).Color = 0x00FF00
    $s.Range('A17').HorizontalAlignment = -4108; $s.Range('A17').VerticalAlignment = -4160
    $s.Range('A18').Orientation = 45
    $s.Range('A19').IndentLevel = 2
    $s.Range('A20').WrapText = $true
    $s.Range('A21').ShrinkToFit = $true
    $s.Range('A22').Locked = $false
    $s.Range('A23').FormulaHidden = $true
    $s.Range('B1').NumberFormat = '0.00%'
    $s.Range('B2').NumberFormat = 'yyyy"年"m"月"d"日"'
    $s.Range('B3').NumberFormat = '#,##0;-#,##0'
    $s.Range('B4').NumberFormat = '@'
    $s.Range('B5').NumberFormat = 'm/d/yyyy'
    $s.Range('B6').NumberFormat = '0.000E+00'
    $s.Range('B7').Style = Find-BuiltinStyle $book @('悪い', 'Bad')
    $s.Range('B8').Style = Find-BuiltinStyle $book @('見出し 1', 'Heading 1')
    $s.Range('B9').Style = Find-BuiltinStyle $book @('メモ', 'Note')
    $s.Range('B10').Style = Find-BuiltinStyle $book @('桁区切り', 'Comma')
    $custom = $book.Styles.Add('試験のスタイル'); $custom.Font.Bold = $true; $custom.Interior.Color = 0xCCCCFF
    $s.Range('B11').Style = '試験のスタイル'
    $s.Range('C1:C3').Interior.Color = 0x99FF99   # cells with a format and no value
    $s.Rows.Item(25).Font.Bold = $true             # row format
    $s.Columns.Item(4).Interior.Color = 0xFFCC00   # column format
}

function Clear-Package([string]$path) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('docProps/core.xml')
        $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
        $xml = $reader.ReadToEnd(); $reader.Close()
        $xml = $xml -replace '<dc:creator>[^<]*</dc:creator>', "<dc:creator>$author</dc:creator>"
        $xml = $xml -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', "<cp:lastModifiedBy>$author</cp:lastModifiedBy>"
        $xml = $xml -replace '(<dcterms:(?:created|modified)[^>]*>)[^<]*', '${1}2026-01-01T00:00:00Z'
        $entry.Delete()
        $entry = $zip.CreateEntry('docProps/core.xml')
        $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding $false))
        $writer.Write($xml); $writer.Close()
    } finally {
        $zip.Dispose()
    }
}

function Clear-Legacy([string]$path, [string[]]$names) {
    # Excel writes the Office user name as the last author of an .xls. Overwrite each
    # occurrence with the same number of bytes so the compound file keeps its layout.
    $bytes = [IO.File]::ReadAllBytes($path)
    $hex = [BitConverter]::ToString($bytes)
    foreach ($name in $names) {
        foreach ($encoding in [Text.Encoding]::GetEncoding(932), [Text.Encoding]::Unicode) {
            $needle = $encoding.GetBytes($name)
            $filler = $encoding.GetBytes(('x' * $name.Length))
            if ($filler.Length -ne $needle.Length) { continue }
            $pattern = [BitConverter]::ToString($needle)
            $index = $hex.IndexOf($pattern)
            while ($index -ge 0) {
                $offset = [int]($index / 3)
                [Array]::Copy($filler, 0, $bytes, $offset, $filler.Length)
                $hex = [BitConverter]::ToString($bytes)
                $index = $hex.IndexOf($pattern)
            }
        }
    }
    [IO.File]::WriteAllBytes($path, $bytes)
}

function Assert-Clean([string]$path, [string[]]$names) {
    $blobs = [Collections.Generic.List[byte[]]]::new()
    $blobs.Add([IO.File]::ReadAllBytes($path))
    try {
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($path)
        try {
            foreach ($entry in $zip.Entries) {
                $stream = New-Object IO.MemoryStream
                $source = $entry.Open(); $source.CopyTo($stream); $source.Close()
                $blobs.Add($stream.ToArray())
            }
        } finally { $zip.Dispose() }
    } catch [IO.InvalidDataException] { }
    foreach ($blob in $blobs) {
        $texts = @([Text.Encoding]::UTF8.GetString($blob), [Text.Encoding]::Unicode.GetString($blob),
            [Text.Encoding]::Unicode.GetString($blob, 1, [Math]::Max(0, $blob.Length - 1)),
            [Text.Encoding]::GetEncoding(932).GetString($blob))
        foreach ($name in $names) {
            foreach ($text in $texts) {
                if ($text.IndexOf($name, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    throw "中断: $path に利用者を示す文字列が残っています"
                }
            }
        }
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$secrets = @($excel.UserName, $env:USERNAME, $env:USERPROFILE) | Where-Object { $_ } | Select-Object -Unique
$written = @()
try {
    $book = $excel.Workbooks.Add()
    Build-Sales $book
    $written += Save-As $book 'sales.xlsb' 50
    $written += Save-As $book 'sales.xlsx' 51
    $written += Save-As $book 'sales.xlam' 55
    # Last: saving as .xls switches the open workbook to the old row and column limits.
    $written += Save-As $book 'sales.xls' 56
    $book.Close($false)

    $book = $excel.Workbooks.Open((Join-Path $OutDir 'sales.xlsb'))
    Edit-Sales $book
    $written += Save-As $book 'sales-edited.xlsb' 50
    $written += Save-As $book 'sales-edited.xlsx' 51
    $written += Save-As $book 'sales-edited.xlam' 55
    $book.Close($false)

    $book = $excel.Workbooks.Add()
    Build-ManyRows $book
    $written += Save-As $book 'many-rows.xlsb' 50
    $book.Close($false)

    $book = $excel.Workbooks.Add()
    Build-FormatsBase $book
    $written += Save-As $book 'formats-base.xlsb' 50
    $written += Save-As $book 'formats-base.xlsx' 51
    Edit-Formats $book
    $written += Save-As $book 'formats.xlsb' 50
    $written += Save-As $book 'formats.xlsx' 51
    $book.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

foreach ($path in $written) {
    if ($path.EndsWith('.xls')) { Clear-Legacy $path $secrets } else { Clear-Package $path }
    Assert-Clean $path $secrets
}
Get-ChildItem -LiteralPath $OutDir | Select-Object Name, Length
