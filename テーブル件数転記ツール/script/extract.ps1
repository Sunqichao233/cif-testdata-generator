<#
    統計情報ファイル（exp_Statics_yyyymmddhhmm.txt）から
    「テーブル統計情報」の部分だけを切り出して CSV にする。

    出力先 : csv\<元のファイル名>.csv
    この CSV を Excel のマクロ n_live_tup取込 で読み込む。
#>
param(
    [string] $Path
)

$ErrorActionPreference = 'Stop'

# ---- 設定 ----
$SectionTitle = 'テーブル統計情報'   # 切り出すセクション名
$Delim        = '|'                  # 統計情報ファイルの区切り文字
$KeyColumns   = @('schemaname', 'relname', 'n_live_tup')   # 見出し行の判定に使う列名
$OutCharset   = 'shift_jis'          # CSV の文字コード（Excel がそのまま開けるように）

$ToolRoot = Split-Path -Parent $PSScriptRoot
$OutDir   = Join-Path $ToolRoot 'csv'

function Get-Enc {
    param([string]$Name)
    switch ($Name.ToLower()) {
        'shift_jis' { [System.Text.Encoding]::GetEncoding(932) }
        'utf-8'     { New-Object System.Text.UTF8Encoding($false) }
        default     { throw "未対応の文字コードです: $Name" }
    }
}

# 1行を CSV の1フィールドにする。カンマ・引用符・改行があれば "" で囲む
function To-CsvField {
    param([string]$s)
    if ($s -match '[",\r\n]') { return '"' + ($s -replace '"', '""') + '"' }
    return $s
}

# 罫線行（----+----）かどうか
function Test-RuleLine {
    param([string]$s)
    $t = $s.Trim()
    if ($t -eq '') { return $false }
    return ($t -notmatch '[^\-\+\| ]')
}

# ---- 入力ファイルの決定 ----
if (-not $Path) {
    $cand = @(Get-ChildItem -LiteralPath $ToolRoot -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -like '*.txt' } |
              Sort-Object LastWriteTime -Descending)

    if ($cand.Count -eq 1) {
        $Path = $cand[0].FullName
    }
    elseif ($cand.Count -gt 1) {
        Write-Host ''
        Write-Host ' 入力ファイルの候補が複数あります。どれを使うか決められません:' -ForegroundColor Yellow
        $cand | ForEach-Object { Write-Host ('   ' + $_.Name) }
        Write-Host ''
        Write-Host ' 対象のファイルを bat に直接ドラッグ＆ドロップして実行してください。'
        Write-Host ''
        exit 1
    }
    else {
        Write-Host ''
        Write-Host ' 統計情報ファイルが見つかりません。' -ForegroundColor Yellow
        Write-Host ''
        Write-Host ' 対処:'
        Write-Host '   ・exp_Statics_*.txt を bat にドラッグ＆ドロップする'
        Write-Host '   ・または .txt をこのフォルダに置いて実行する'
        Write-Host ''
        exit 1
    }
}
if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host ''
    Write-Host " ファイルが存在しません: $Path" -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

# ---- 読み込み（Shift_JIS と UTF-8 を試し、セクション名が読めた方を使う）----
$bytes = [System.IO.File]::ReadAllBytes($Path)
$text  = $null
$usedCharset = ''
foreach ($cs in @('shift_jis', 'utf-8')) {
    $t = (Get-Enc $cs).GetString($bytes)
    if ($t.IndexOf($SectionTitle) -ge 0) { $text = $t; $usedCharset = $cs; break }
}
if ($null -eq $text) {
    Write-Host ''
    Write-Host " 「$SectionTitle」のセクションが見つかりません。" -ForegroundColor Yellow
    Write-Host ' ファイルを確認してください。'
    Write-Host ''
    exit 1
}

$lines = ($text -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' テーブル統計情報の切り出し' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  入力       : $Path"
Write-Host "  文字コード : 入力 $usedCharset / 出力 $OutCharset"

# ---- 切り出し ----
$out       = New-Object 'System.Collections.Generic.List[string]'
$inSection = $false
$headerOk  = $false
$started   = $false
$dataRows  = 0
$headLine  = ''

foreach ($ln in $lines) {

    if (-not $inSection) {
        # セクションの開始を探す
        if ($ln.IndexOf($SectionTitle) -ge 0) { $inSection = $true }
        continue
    }

    if (-not $headerOk) {
        # 見出し行を探す。3つの列名がそろって初めて見出しとみなす。
        # 「テーブルサイズ」のように schemaname と relname だけある行を拾わないため。
        if ($ln.Contains($Delim)) {
            $names = @($ln -split [regex]::Escape($Delim) | ForEach-Object { $_.Trim().ToLower() })
            $miss  = @($KeyColumns | Where-Object { $names -notcontains $_ })
            if ($miss.Count -eq 0) {
                $headerOk = $true
                $headLine = $ln
                [void]$out.Add((@($ln -split [regex]::Escape($Delim) |
                                  ForEach-Object { To-CsvField $_.Trim() }) -join ','))
            }
        }
        continue
    }

    # データ行
    if ($ln.Trim() -eq '')        { if ($started) { break } else { continue } }
    if (Test-RuleLine $ln)        { continue }
    if (-not $ln.Contains($Delim)) { break }     # (16 rows) や次のセクション

    $started = $true
    $dataRows++
    [void]$out.Add((@($ln -split [regex]::Escape($Delim) |
                      ForEach-Object { To-CsvField $_.Trim() }) -join ','))
}

if (-not $headerOk) {
    Write-Host ''
    Write-Host ' 見出し行が見つかりません。' -ForegroundColor Yellow
    Write-Host ("  必要な列名 : " + ($KeyColumns -join ' / '))
    Write-Host ''
    exit 1
}
if ($dataRows -eq 0) {
    Write-Host ''
    Write-Host ' データ行が0件です。ファイルを確認してください。' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

# ---- 出力 ----
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$base    = [System.IO.Path]::GetFileNameWithoutExtension($Path)
$outPath = Join-Path $OutDir ($base + '.csv')

$oenc = Get-Enc $OutCharset
[System.IO.File]::WriteAllBytes($outPath, $oenc.GetBytes((($out -join "`r`n") + "`r`n")))

# 見出しの列位置を出す。マクロ側の既定（B/C/O列）と合っているかの確認用
$names = @($headLine -split [regex]::Escape($Delim) | ForEach-Object { $_.Trim().ToLower() })
function Get-ColLetter {
    param([int]$n)   # 1 始まり
    $s = ''
    while ($n -gt 0) {
        $n--
        $s = [char](65 + ($n % 26)) + $s
        $n = [int][math]::Floor($n / 26)
    }
    return $s
}

Write-Host ''
Write-Host '--- 結果 ---' -ForegroundColor Cyan
Write-Host ("  データ行数 : {0} 行" -f $dataRows)
Write-Host ("  列数       : {0} 列" -f $names.Count)
foreach ($k in $KeyColumns) {
    $i = [array]::IndexOf($names, $k)
    if ($i -ge 0) {
        Write-Host ("  {0,-12}: {1} 列目（{2}列）" -f $k, ($i + 1), (Get-ColLetter ($i + 1)))
    } else {
        Write-Host ("  {0,-12}: 見つかりません" -f $k) -ForegroundColor Yellow
    }
}
Write-Host ''
Write-Host ("  出力 : {0}" -f $outPath) -ForegroundColor Green
Write-Host ''
Write-Host ' この CSV を Excel のマクロ n_live_tup取込 で読み込んでください。' -ForegroundColor Green
Write-Host ''
exit 0
