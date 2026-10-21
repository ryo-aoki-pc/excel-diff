<#
Open workbooks in a private, hidden Excel instance and print what Excel sees as JSON.
Used by tests/test_excel_open.py. Exit code 3 means Excel is unavailable or already running.

    pwsh -NoProfile -File tests/excel_open.ps1 -ListFile <paths.txt> [-FormatCells '売上!M2,売上!M3']

For every non-empty cell the JSON has the value (Value2) and, for formulas, Excel's
English A1 formula. Font bold, fill colour and number format are reported for the
non-empty cells and for the cells named in -FormatCells.
#>
param(
    [Parameter(Mandatory = $true)][string]$ListFile,
    [string]$FormatCells = ''
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Output 'Excel is already running'
    exit 3
}
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Output ('Excel is not available: ' + $_.Exception.Message)
    exit 3
}
$errorCodes = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'
                 -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!' }

function Convert-Value($value) {
    # COM returns cell errors (VT_ERROR) as negative integers.
    if ($value -is [int] -and $errorCodes.ContainsKey($value)) { return [ordered]@{ error = $errorCodes[$value] } }
    return $value
}

function Get-Format($cell) {
    $font = $cell.Font; $interior = $cell.Interior
    $format = [ordered]@{
        bold = [bool]$font.Bold; italic = [bool]$font.Italic; underline = $font.Underline
        strikethrough = [bool]$font.Strikethrough; superscript = [bool]$font.Superscript; subscript = [bool]$font.Subscript
        fontName = $font.Name; size = $font.Size; fontColor = $font.Color
        fill = $interior.Color; pattern = $interior.Pattern; patternColor = $interior.PatternColor
        numberFormat = $cell.NumberFormat; horizontal = $cell.HorizontalAlignment; vertical = $cell.VerticalAlignment
        orientation = $cell.Orientation; indent = $cell.IndentLevel; wrap = [bool]$cell.WrapText; shrink = [bool]$cell.ShrinkToFit
        locked = [bool]$cell.Locked; formulaHidden = [bool]$cell.FormulaHidden; style = $cell.Style.NameLocal
    }
    if ($interior.Pattern -eq 4000) { $format.gradientDegree = $interior.Gradient.Degree }
    foreach ($edge in 5, 6, 7, 8, 9, 10) {
        $border = $cell.Borders.Item($edge)
        $format['border' + $edge] = '{0}/{1}/{2}' -f $border.LineStyle, $border.Weight, $border.Color
    }
    $format
}

$extra = @{}
foreach ($item in ($FormatCells -split ',')) {
    if ($item -match '^(.+)!([A-Z]+[0-9]+)$') {
        if (-not $extra.ContainsKey($Matches[1])) { $extra[$Matches[1]] = @() }
        $extra[$Matches[1]] += $Matches[2]
    }
}

$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: never run macros
$results = @()
try {
    foreach ($path in [IO.File]::ReadAllLines($ListFile, [Text.Encoding]::UTF8)) {
        if (-not $path) { continue }
        $entry = [ordered]@{ path = $path; ok = $false }
        try {
            $book = [System.__ComObject].InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null,
                $excel.Workbooks, @($path, 0, $true), $null, $null, @('Filename', 'UpdateLinks', 'ReadOnly'))
            try {
                $entry.caption = $excel.Caption
                $sheets = [ordered]@{}
                foreach ($sheet in $book.Worksheets) {
                    $cells = [ordered]@{}
                    $used = $sheet.UsedRange
                    $top = $used.Row; $left = $used.Column
                    $values = $used.Value2; $formulas = $used.Formula
                    if ($used.Cells.Count -eq 1) {
                        $values = [object[,]]::new(1, 1); $values[0, 0] = $used.Value2
                        $formulas = [object[,]]::new(1, 1); $formulas[0, 0] = $used.Formula
                    }
                    $rows = $values.GetLength(0); $columns = $values.GetLength(1)
                    $base0 = $values.GetLowerBound(0); $base1 = $values.GetLowerBound(1)
                    for ($r = 0; $r -lt $rows; $r++) {
                        for ($c = 0; $c -lt $columns; $c++) {
                            $value = $values[($r + $base0), ($c + $base1)]
                            $formula = [string]$formulas[($r + $base0), ($c + $base1)]
                            if ($null -eq $value -and -not $formula) { continue }
                            $cell = $sheet.Cells.Item($top + $r, $left + $c)
                            $item = [ordered]@{ value = (Convert-Value $value) }
                            if ($formula.StartsWith('=')) { $item.formula = $formula }
                            $item.format = Get-Format $cell
                            $cells[$cell.Address($false, $false)] = $item
                        }
                    }
                    foreach ($address in @($extra[$sheet.Name])) {
                        if (-not $address -or $cells.Contains($address)) { continue }
                        $cells[$address] = [ordered]@{ value = $null; format = (Get-Format $sheet.Range($address)) }
                    }
                    $sheets[$sheet.Name] = $cells
                }
                $entry.sheets = $sheets
                $entry.ok = $true
            } finally {
                $book.Close($false)
            }
        } catch {
            $entry.error = $_.Exception.Message
        }
        $results += [pscustomobject]$entry
    }
} finally {
    $excel.Quit()
    # Excel exits only after every COM reference is released.
    $book = $null; $sheet = $null; $used = $null; $cell = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}
ConvertTo-Json -InputObject @($results) -Depth 8 -Compress
