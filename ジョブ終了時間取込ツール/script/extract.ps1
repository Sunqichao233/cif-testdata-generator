<#
    ev_app_*.log（イベントログのテキスト出力）から
    ジョブの開始・終了時刻を切り出して CSV にする。

    出力先 : csv\jobtime_<yyyymmdd>.csv
    この CSV を Excel のマクロで @ジョブ終了時間確認_filter.xlsx に取り込む。
#>
param(
    [string] $Path,                       # ログ1本、またはフォルダ
    [int]    $UtcOffsetHours = 0,         # Date が UTC の場合は 9 を指定する
    [int]    $DayCutoffHour  = 7           # 業務日の切れ目。7 なら 07:00〜翌06:59 が同じ業務日
                                           # 橋本さん「5月1日というのは7時から始まるところの5月1日」
)

$ErrorActionPreference = 'Stop'

# ---- 設定 ----
$Source     = 'OCULUSBT'     # 対象のイベントソース。これ以外は無視する
$IdStart    = 'COM100001'    # 「〜 を開始します」
$IdEnd      = 'COM100003'    # 「〜 が正常終了しました」
$GroupSep   = '－'           # ジョブ名の区切り（全角ハイフン）。グループ－サブタスク
$OutCharset = 'shift_jis'

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

function To-CsvField {
    param([string]$s)
    if ($s -match '[",\r\n]') { return '"' + ($s -replace '"', '""') + '"' }
    return $s
}

# ---- 入力ファイルの決定 ----
if (-not $Path) { $Path = $ToolRoot }

if (Test-Path -LiteralPath $Path -PathType Container) {
    $files = @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like 'ev_app*.log' } |
               Sort-Object Name)
} elseif (Test-Path -LiteralPath $Path) {
    $files = @(Get-Item -LiteralPath $Path)
} else {
    Write-Host ''
    Write-Host " ファイルが見つかりません: $Path" -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

if ($files.Count -eq 0) {
    Write-Host ''
    Write-Host ' ev_app*.log が見つかりません。' -ForegroundColor Yellow
    Write-Host ''
    Write-Host ' 対処:'
    Write-Host '   ・ログを bat にドラッグ＆ドロップする'
    Write-Host '   ・または ev_app*.log をこのフォルダに置いて実行する'
    Write-Host ''
    exit 1
}

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' ジョブ開始・終了時刻の切り出し' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host ("  対象ファイル : {0} 本" -f $files.Count)
if ($UtcOffsetHours -ne 0) {
    Write-Host ("  時差補正     : +{0} 時間" -f $UtcOffsetHours) -ForegroundColor Yellow
}

# ---- 解析 ----
# key = コンピュータ名 + ジョブグループ名
# 開始は COM100001 の最も早いもの、終了は COM100003 の最も遅いものを採る。
# 1つのグループに複数のサブタスクがあるため。
$jobs    = @{}
$nEvent  = 0
$nHit    = 0
$skipSrc = 0

foreach ($f in $files) {

    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $text  = $null
    foreach ($cs in @('shift_jis', 'utf-8')) {
        $t = (Get-Enc $cs).GetString($bytes)
        if ($t.IndexOf('Log Name:') -ge 0) { $text = $t; break }
    }
    if ($null -eq $text) {
        Write-Host ("  読めません: {0}" -f $f.Name) -ForegroundColor Yellow
        continue
    }

    $lines = ($text -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"

    $curSrc = ''; $curDate = ''; $curComp = ''; $inDesc = $false

    foreach ($ln in $lines) {

        if ($ln -match '^Event\[\d+\]') {
            $curSrc = ''; $curDate = ''; $curComp = ''; $inDesc = $false
            $nEvent++
            continue
        }
        if ($ln -match '^\s*Source:\s*(.+?)\s*$')   { $curSrc  = $Matches[1]; continue }
        if ($ln -match '^\s*Date:\s*(.+?)\s*$')     { $curDate = $Matches[1]; continue }
        if ($ln -match '^\s*Computer:\s*(.+?)\s*$') { $curComp = $Matches[1]; continue }
        if ($ln -match '^\s*Description:')          { $inDesc  = $true; continue }

        if (-not $inDesc) { continue }
        if ($ln.Trim() -eq '') { $inDesc = $false; continue }

        # ---- ここが Description の本文 ----
        $inDesc = $false
        if ($curSrc -ne $Source) { $skipSrc++; continue }

        $isStart = $ln.Contains($IdStart)
        $isEnd   = $ln.Contains($IdEnd)
        if (-not ($isStart -or $isEnd)) { continue }

        # "OCULUS-I COM100001 業後処理－ファイルバックアップ を開始します"
        $m = [regex]::Match($ln, '\sCOM\d+\s+(.+?)\s+(?:を開始します|が正常終了しました)')
        if (-not $m.Success) { continue }
        $jobName = $m.Groups[1].Value.Trim()
        $group   = $jobName
        $i = $jobName.IndexOf($GroupSep)
        if ($i -gt 0) { $group = $jobName.Substring(0, $i).Trim() }

        # 日時
        # 末尾の Z を外して、書かれている時刻をそのまま使う。
        # 本当に UTC だった場合は -UtcOffsetHours 9 を付けて補正する。
        $d  = $curDate -replace 'Z$', ''
        $dt = [datetime]::MinValue
        $ok = $false
        try {
            $dt = [datetime]::Parse($d, [System.Globalization.CultureInfo]::InvariantCulture)
            $ok = $true
        } catch { }
        if (-not $ok) { continue }
        if ($UtcOffsetHours -ne 0) { $dt = $dt.AddHours($UtcOffsetHours) }

        # 業務日 : 夜間バッチは日をまたぐため、切れ目をずらして同じ日にまとめる
        #          既定 7 なら 07:00〜翌06:59 が同じ業務日になる（夕会で確認）
        $bizDate = $dt.AddHours(-1 * $DayCutoffHour).Date

        $key = $curComp + '|' + $bizDate.ToString('yyyyMMdd') + '|' + $group
        if (-not $jobs.ContainsKey($key)) {
            $jobs[$key] = [ordered]@{
                Computer = $curComp; BizDate = $bizDate; Group = $group
                Start = $null; End = $null; Subs = @{}
            }
        }
        $o = $jobs[$key]
        $o.Subs[$jobName] = $true
        if ($isStart) {
            if ($null -eq $o.Start -or $dt -lt $o.Start) { $o.Start = $dt }
        } else {
            if ($null -eq $o.End   -or $dt -gt $o.End)   { $o.End   = $dt }
        }
        $nHit++
    }
}

if ($jobs.Count -eq 0) {
    Write-Host ''
    Write-Host ' ジョブの開始・終了が1件も取れませんでした。' -ForegroundColor Yellow
    Write-Host ("  ソース   : {0}" -f $Source)
    Write-Host ("  開始     : {0} / 終了 : {1}" -f $IdStart, $IdEnd)
    Write-Host ' ログの内容を確認してください。'
    Write-Host ''
    exit 1
}

# ---- 出力 ----
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

$rows = @($jobs.Values | Sort-Object { $_.BizDate }, { $_.Computer }, { $_.Start })

$head = @('業務日','コンピュータ','ジョブグループ','開始日付','開始時刻','終了日付','終了時刻',
          '処理時間','サブタスク数','業務日の曜日','警告') -join ','
$out = New-Object 'System.Collections.Generic.List[string]'
[void]$out.Add($head)

$wd = @('日曜','月曜','火曜','水曜','木曜','金曜','土曜')
foreach ($r in $rows) {
    $warn = ''
    if ($null -eq $r.Start) { $warn = '開始なし' }
    elseif ($null -eq $r.End) { $warn = '終了なし' }

    $dur = ''
    if ($r.Start -and $r.End) {
        $ts = $r.End - $r.Start
        $dur = '{0}:{1:00}:{2:00}' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
    }
    $f = @(
        $r.BizDate.ToString('yyyy/MM/dd'), $r.Computer, $r.Group,
        $(if ($r.Start) { $r.Start.ToString('yyyy/MM/dd') } else { '' }),
        $(if ($r.Start) { $r.Start.ToString('HH:mm:ss') } else { '' }),
        $(if ($r.End)   { $r.End.ToString('yyyy/MM/dd') }   else { '' }),
        $(if ($r.End)   { $r.End.ToString('HH:mm:ss') }   else { '' }),
        $dur, $r.Subs.Count,
        $wd[[int]$r.BizDate.DayOfWeek],
        $warn
    ) | ForEach-Object { To-CsvField ([string]$_) }
    [void]$out.Add(($f -join ','))
}

$stamp = $rows[0].BizDate.ToString('yyyyMMdd')
if ($rows[-1].BizDate -ne $rows[0].BizDate) {
    $stamp = $stamp + '-' + $rows[-1].BizDate.ToString('yyyyMMdd')
}
$outPath = Join-Path $OutDir ("jobtime_{0}.csv" -f $stamp)
[System.IO.File]::WriteAllBytes($outPath,
    (Get-Enc $OutCharset).GetBytes((($out -join "`r`n") + "`r`n")))

# ---- 結果 ----
Write-Host ''
Write-Host '--- 結果 ---' -ForegroundColor Cyan
Write-Host ("  読んだイベント : {0} 件（うち対象 {1} 件 / 他ソース {2} 件）" -f $nEvent, $nHit, $skipSrc)
Write-Host ("  ジョブグループ : {0} 件" -f $rows.Count)
Write-Host ''
Write-Host ('  {0,-11} {1,-4} {2,-12} {3,-34} {4,-9} {5,-9} {6}' -f `
            '業務日','曜日','コンピュータ','ジョブグループ','開始','終了','処理時間')
foreach ($r in $rows) {
    $dur = ''
    if ($r.Start -and $r.End) {
        $ts = $r.End - $r.Start
        $dur = '{0}:{1:00}:{2:00}' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
    }
    Write-Host ('  {0,-11} {1,-4} {2,-12} {3,-34} {4,-9} {5,-9} {6}' -f `
        $r.BizDate.ToString('yyyy/MM/dd'), $wd[[int]$r.BizDate.DayOfWeek],
        $r.Computer, $r.Group,
        $(if ($r.Start) { $r.Start.ToString('HH:mm:ss') } else { '-' }),
        $(if ($r.End)   { $r.End.ToString('HH:mm:ss') }   else { '-' }),
        $dur)
}
Write-Host ''
Write-Host ("  出力 : {0}" -f $outPath) -ForegroundColor Green
Write-Host ''
exit 0
