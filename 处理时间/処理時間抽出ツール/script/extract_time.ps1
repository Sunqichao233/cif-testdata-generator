<#
    処理時間抽出ツール

    イベントビューアーのログ（ev_app_YYYYMMDDHHII.log）を複数まとめて読み込み、
    タスクごとの開始・終了時刻と処理時間を一覧にする。

    入力 : ログファイル または フォルダ（複数可）。省略時はツール直下の「入力」フォルダ。
           フォルダの場合は中の ev_app_*.log をすべて読む（サブフォルダも含む）。
    出力 : 出力\ 以下（毎回同じフォルダ）
           同じ名前のファイルがある場合は、置き換えではなく更新する。
           読み込んだログは 出力\_蓄積\ に保存しておき、次回は今回分と合わせて集計し直す。
           00_メッセージ一覧.csv          … ログに出てきたメッセージの種類（タスク定義の調整用）
           01_処理時間一覧.csv            … 業務日 × タスクごとの開始・終了・処理時間
           02_処理時間表_<区分>.csv       … タスク × 業務日 の処理時間表（M／MTWTFS／S）
           03_最遅最速_<基準タスク>.csv   … 基準タスクが最も遅い日／早い日の、各タスクの処理時間
           切り出し\<区分>-<No>-1 <タスク名>_開始.txt ／ -2 …_終了.txt
                                          … 前任者の切り出しファイルと同じ形式
#>
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $Paths
)

$ErrorActionPreference = 'Stop'

# ======================= 設定 =======================
$DayStart       = '07:00'        # 業務日の開始時刻。この時刻より前のログは前日分として扱う
$Source         = 'OCULUSBT'     # 対象のイベントソース
$BaseTask       = '業後処理'     # 最遅／最速を判定する基準タスク（タスク定義.csv のタスク名）
$CompareBy      = '終了時刻'     # 最遅／最速の判定基準：'終了時刻' または '処理時間'
$GroupBy        = '月'           # 最遅／最速を判定する単位：'月'（月ごと）／'月×区分'／'区分'（全期間を区分ごと）
$SaturdayClass  = 'MTWTFS'       # 土曜日の区分（M=第1月曜／MTWTFS=平日／S=日曜）
$FilePattern    = 'ev_app_*.log' # フォルダ指定時に読み込むファイル名
$OutCharset     = 'utf-8-bom'    # 出力の文字コード：'utf-8-bom'（Excel でそのまま開ける・文字化けしない）または 932（Shift_JIS）
# ====================================================

$ToolRoot = Split-Path -Parent $PSScriptRoot
$DefFile  = Join-Path $ToolRoot 'タスク定義.csv'
$OutEnc   = if ($OutCharset -eq 'utf-8-bom') { New-Object System.Text.UTF8Encoding($true) } else { [System.Text.Encoding]::GetEncoding($OutCharset) }
$DayStartSpan = [TimeSpan]::Parse($DayStart)

function Write-Info([string]$msg) { Write-Host ('  ' + $msg) }

# ---------- 文字コード判定（BOM → UTF-16(BOM無し) → UTF-8 → Shift_JIS の順） ----------
function Get-FileEncoding([string]$file) {
    $fs = [System.IO.File]::OpenRead($file)
    try {
        $buf = New-Object byte[] 65536
        $n = $fs.Read($buf, 0, $buf.Length)
    } finally { $fs.Close() }
    if ($n -ge 2 -and $buf[0] -eq 0xFF -and $buf[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
    if ($n -ge 2 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
    if ($n -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8 }
    $zeroOdd = 0
    for ($i = 1; $i -lt [Math]::Min($n, 2000); $i += 2) { if ($buf[$i] -eq 0) { $zeroOdd++ } }
    if ($zeroOdd -gt 300) { return [System.Text.Encoding]::Unicode }
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$strict.GetString($buf, 0, [Math]::Max(0, $n - 4))
        return [System.Text.Encoding]::UTF8
    } catch {
        return [System.Text.Encoding]::GetEncoding(932)
    }
}

# ---------- CSV ----------
function To-CsvField($v) {
    $s = [string]$v
    if ($s -match '[",\r\n]') { return '"' + ($s -replace '"', '""') + '"' }
    return $s
}
function Write-Csv([string]$path, [string[]]$header, $rows) {
    # 1行だけのときに配列が展開されてしまうのを防ぐ
    if ($null -eq $rows) { $rows = @() }
    elseif ($rows -is [System.Collections.IList] -and $rows.Count -gt 0 -and
            -not ($rows[0] -is [System.Collections.IList])) { $rows = ,$rows }
    $sw = New-Object System.IO.StreamWriter($path, $false, $OutEnc)
    try {
        $sw.WriteLine((($header | ForEach-Object { To-CsvField $_ }) -join ','))
        foreach ($r in $rows) { $sw.WriteLine((($r | ForEach-Object { To-CsvField $_ }) -join ',')) }
    } finally { $sw.Close() }
}
function Fmt-Span($ts) {
    if ($null -eq $ts) { return '' }
    $sign = ''
    if ($ts.Ticks -lt 0) { $sign = '-'; $ts = $ts.Negate() }
    return ('{0}{1:00}:{2:00}:{3:00}' -f $sign, [Math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds)
}
function Fmt-Dt($dt) { if ($null -eq $dt) { return '' } return $dt.ToString('yyyy/MM/dd HH:mm:ss') }

# ---------- タスク定義 ----------
if (-not (Test-Path $DefFile)) { throw "タスク定義.csv が見つかりません: $DefFile" }
$defLines = [System.IO.File]::ReadAllLines($DefFile, (Get-FileEncoding $DefFile))
$Tasks = @()
foreach ($line in ($defLines | Select-Object -Skip 1)) {
    if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) { continue }
    $c = $line -split ','
    if ($c.Count -lt 4) { continue }
    $Tasks += [pscustomobject]@{
        No       = $c[0].Trim()
        Name     = $c[1].Trim()
        StartRx  = New-Object System.Text.RegularExpressions.Regex($c[2].Trim())
        EndRx    = New-Object System.Text.RegularExpressions.Regex($c[3].Trim())
    }
}
if ($Tasks.Count -eq 0) { throw 'タスク定義.csv にタスクがありません。' }
Write-Info ('タスク定義      : {0} 件' -f $Tasks.Count)

$sw0 = [Diagnostics.Stopwatch]::StartNew()
function Write-Step([string]$msg) { Write-Host ('  [{0:mm\:ss}] {1}' -f $sw0.Elapsed, $msg) -ForegroundColor Cyan }

# ---------- 入力ファイル ----------
if (-not $Paths -or $Paths.Count -eq 0) { $Paths = @(Join-Path $ToolRoot '入力') }
$files = @()
foreach ($p in $Paths) {
    if (-not $p) { continue }
    $p = $p.Trim('"')
    if (Test-Path $p -PathType Container) {
        $files += Get-ChildItem -Path $p -Filter $FilePattern -Recurse -File | ForEach-Object { $_.FullName }
    } elseif (Test-Path $p -PathType Leaf) {
        $files += (Resolve-Path $p).Path
    } else {
        Write-Warning "見つかりません: $p"
    }
}
$files = $files | Sort-Object -Unique
if (-not $files -or @($files).Count -eq 0) { throw 'ログファイルがありません。ファイルかフォルダをドラッグ＆ドロップしてください。' }
Write-Info ('ログファイル    : {0} 件' -f @($files).Count)

# ---------- 読み込み（Event[n] ブロック単位） ----------
# メッセージ例：OCULUS-I COM900100 長期保管用ファイルバックアップ処理を開始します。
$msgRx  = New-Object System.Text.RegularExpressions.Regex('^\s*OCULUS-([A-Z])\s*([A-Z]+\d+)\s*(.*)$')
$seen   = New-Object 'System.Collections.Generic.HashSet[string]'
$events = New-Object System.Collections.Generic.List[object]
$total = 0; $dup = 0

$ranges = New-Object System.Collections.Generic.List[object]

# ---------- 出力フォルダ・蓄積データ ----------
$OutDir   = Join-Path $ToolRoot '出力'
$CutDir   = Join-Path $OutDir '切り出し'
$StoreDir = Join-Path $OutDir '_蓄積'
$StoreEv  = Join-Path $StoreDir 'イベント.tsv'
$StoreRg  = Join-Path $StoreDir 'ログ範囲.tsv'
$fmtT     = 'yyyy-MM-dd HH:mm:ss.fffffff'
$Inv      = [Globalization.CultureInfo]::InvariantCulture
New-Item -ItemType Directory -Path $StoreDir -Force | Out-Null
function Get-EvKey($t, $lv, $id, $tx) { return ('{0}|{1}|{2}|{3}' -f $t.Ticks, $lv, $id, $tx) }
$stored = 0
if (Test-Path $StoreEv) {
    foreach ($l in [System.IO.File]::ReadAllLines($StoreEv, [System.Text.Encoding]::UTF8)) {
        $c = $l -split "`t", 5
        if ($c.Count -lt 5 -or $c[0] -eq 'Time') { continue }
        $t = [DateTime]::ParseExact($c[0], $fmtT, $Inv)
        if ($seen.Add((Get-EvKey $t $c[1] $c[2] $c[4]))) {
            $events.Add([pscustomobject]@{ Time = $t; Level = $c[1]; Id = $c[2]; Text = $c[4]; File = $c[3] })
            $stored++
        }
    }
}
if (Test-Path $StoreRg) {
    foreach ($l in [System.IO.File]::ReadAllLines($StoreRg, [System.Text.Encoding]::UTF8)) {
        $c = $l -split "`t"
        if ($c.Count -lt 2 -or $c[0] -eq 'From') { continue }
        $ranges.Add(@([DateTime]::ParseExact($c[0], $fmtT, $Inv), [DateTime]::ParseExact($c[1], $fmtT, $Inv)))
    }
}
Write-Info ('前回までの蓄積  : {0} 件' -f $stored)

Write-Step 'ログを読み込んでいます（1ファイル数秒〜十数秒かかります）…'
foreach ($f in $files) {
    $enc = Get-FileEncoding $f
    $sr = New-Object System.IO.StreamReader($f, $enc, $true)
    $cnt = 0
    $fFirst = $null; $fLast = $null   # このファイルに含まれるログの時刻範囲（全ソース）
    try {
        $src = $null; $date = $null; $inDesc = $false; $desc = $null
        $flush = {
            if ($src -eq $Source -and $date -and $desc) {
                $m = $msgRx.Match($desc)
                if ($m.Success) {
                    $script:total++
                    $dt = [DateTime]::Parse($date.TrimEnd('Z'), $Inv)   # Date の末尾の Z は無視し、ログに書かれた時刻をそのまま使う
                    $lv = $m.Groups[1].Value; $id = $m.Groups[2].Value; $tx = ($m.Groups[3].Value.Trim() -replace "`t", ' ')
                    if ($seen.Add((Get-EvKey $dt $lv $id $tx))) {
                        $events.Add([pscustomobject]@{ Time = $dt; Level = $lv; Id = $id; Text = $tx; File = [IO.Path]::GetFileName($f) })
                        $script:cnt++
                    } else { $script:dup++ }
                }
            }
        }
        while ($null -ne ($line = $sr.ReadLine())) {
            if ($line.StartsWith('Event[')) {
                . $flush
                $src = $null; $date = $null; $inDesc = $false; $desc = $null
                continue
            }
            if ($inDesc) {
                if ($line.Trim().Length -gt 0) {
                    if ($null -eq $desc) { $desc = $line.Trim() } else { $desc += ' ' + $line.Trim() }
                }
                continue
            }
            $t = $line.TrimStart()
            if     ($t.StartsWith('Source:'))      { $src  = $t.Substring(7).Trim() }
            elseif ($t.StartsWith('Date:'))        {
                $date = $t.Substring(5).Trim()
                if (-not $fFirst) { $fFirst = $date }
                $fLast = $date
            }
            elseif ($t.StartsWith('Description:')) {
                $inDesc = $true
                $rest = $t.Substring(12).Trim()
                if ($rest) { $desc = $rest }
            }
        }
        . $flush
    } finally { $sr.Close() }
    if ($fFirst) {
        $a = [DateTime]::Parse($fFirst.TrimEnd('Z'), [Globalization.CultureInfo]::InvariantCulture)
        $b = [DateTime]::Parse($fLast.TrimEnd('Z'),  [Globalization.CultureInfo]::InvariantCulture)
        $ranges.Add(@([DateTime[]]@($a, $b) | Sort-Object))
    }
    Write-Info ('  {0}  ({1}) 新規 {2} 件' -f [IO.Path]::GetFileName($f), $enc.WebName, $cnt)
}
Write-Info ('{0} イベント : 今回 {1} 件（うち重複・蓄積済み {2} 件を除外）→ 蓄積と合わせて {3} 件' -f $Source, $total, $dup, $events.Count)
if ($events.Count -eq 0) { throw "$Source のイベントが見つかりませんでした。" }

$events = @($events | Sort-Object Time)

# 蓄積データを保存（次回の実行で、そのとき読み込むログと合わせて集計する）
$utf8 = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($StoreEv, [string[]](@("Time`tLevel`tId`tFile`tText") +
    @($events | ForEach-Object { ($_.Time.ToString($fmtT), $_.Level, $_.Id, $_.File, $_.Text) -join "`t" })), $utf8)
[System.IO.File]::WriteAllLines($StoreRg, [string[]](@("From`tTo") +
    @($ranges | ForEach-Object { ($_[0].ToString($fmtT), $_[1].ToString($fmtT)) -join "`t" })), $utf8)

# ---------- 業務日・区分 ----------
function Get-BizDay([DateTime]$t) { return ($t - $DayStartSpan).Date }
function Get-Class([DateTime]$d) {
    switch ($d.DayOfWeek) {
        'Sunday'   { return 'S' }
        'Saturday' { return $SaturdayClass }
        'Monday'   { if ($d.Day -le 7) { return 'M' } else { return 'MTWTFS' } }
        default    { return 'MTWTFS' }
    }
}
$WeekJa = @{ Sunday='日'; Monday='月'; Tuesday='火'; Wednesday='水'; Thursday='木'; Friday='金'; Saturday='土' }

Write-Step '集計しています…'
# ---------- 00 メッセージ一覧 ----------
$msgStat = @{}
foreach ($e in $events) {
    $norm = $e.Text -replace '\d+', '#'
    $k = $e.Level + '|' + $e.Id + '|' + $norm
    if (-not $msgStat.ContainsKey($k)) {
        $hit = ($Tasks | Where-Object { $_.StartRx.IsMatch($e.Text) -or $_.EndRx.IsMatch($e.Text) } | Select-Object -First 1)
        $msgStat[$k] = [pscustomobject]@{ Level=$e.Level; Id=$e.Id; Text=$e.Text; Count=0; First=$e.Time; Last=$e.Time
                                         Task = $(if ($hit) { $hit.Name } else { '' }) }
    }
    $s = $msgStat[$k]; $s.Count++; $s.Last = $e.Time
}
Write-Csv (Join-Path $OutDir '00_メッセージ一覧.csv') `
    @('レベル','メッセージID','メッセージ（例）','件数','初回','最終','該当タスク') `
    ($msgStat.Values | Sort-Object Id, Text | ForEach-Object {
        ,@($_.Level, $_.Id, $_.Text, $_.Count, (Fmt-Dt $_.First), (Fmt-Dt $_.Last), $_.Task) })

# ---------- 業務日 × タスク ----------
$days = $events | ForEach-Object { Get-BizDay $_.Time } | Sort-Object -Unique
# 業務日（開始時刻〜翌日の開始時刻）がログの範囲で完全にカバーされているか
# 各ファイルの時刻範囲をつなげ、途切れなく覆われている日だけを「完全」とする
$merged = New-Object System.Collections.Generic.List[object]
foreach ($rg in ($ranges | Sort-Object { $_[0] })) {
    if ($merged.Count -and $rg[0] -le $merged[$merged.Count - 1][1]) {
        if ($rg[1] -gt $merged[$merged.Count - 1][1]) { $merged[$merged.Count - 1][1] = $rg[1] }
    } else { $merged.Add(@($rg[0], $rg[1])) }
}
$fullDay = @{}
foreach ($d in $days) {
    $from = $d + $DayStartSpan; $to = $from.AddDays(1)
    $fullDay[$d] = [bool](@($merged | Where-Object { $_[0] -le $from -and $_[1] -ge $to }).Count)
}

$byDay = @{}
foreach ($e in $events) {
    $d = Get-BizDay $e.Time
    if (-not $byDay.ContainsKey($d)) { $byDay[$d] = New-Object System.Collections.Generic.List[object] }
    $byDay[$d].Add($e)
}

$result = New-Object System.Collections.Generic.List[object]
foreach ($d in $days) {
    $cls = Get-Class $d
    foreach ($tk in $Tasks) {
        $st = @($byDay[$d] | Where-Object { $tk.StartRx.IsMatch($_.Text) })
        $en = @($byDay[$d] | Where-Object { $tk.EndRx.IsMatch($_.Text) })
        $errs = @($byDay[$d] | Where-Object { $_.Level -eq 'E' -and ($tk.StartRx.IsMatch($_.Text) -or $tk.EndRx.IsMatch($_.Text)) })
        if ($st.Count -eq 0 -and $en.Count -eq 0) { continue }
        $s = if ($st.Count) { $st[0] } else { $null }
        $e = if ($en.Count) { $en[$en.Count - 1] } else { $null }
        # 開始・終了パターンが同じでログが1件しかない場合、同じログを開始と終了の両方に使わない
        if ($s -and $e -and [object]::ReferenceEquals($s, $e)) { $e = $null }
        $note = @()
        if (-not $fullDay[$d]) { $note += 'ログ不完全（業務日の一部しかログがない）' }
        if (-not $s) { $note += '開始ログなし' }
        if (-not $e) { $note += '終了ログなし' }
        $dur = $null
        if ($s -and $e) {
            # ログの表示と合うよう、秒未満を切り捨ててから差を取る
            $dur = [TimeSpan]::FromSeconds([Math]::Floor(($e.Time.Ticks / 1e7)) - [Math]::Floor(($s.Time.Ticks / 1e7)))
            if ($dur.Ticks -lt 0) { $note += '終了が開始より前' }
        }
        if ($errs.Count) { $note += ('エラーログ {0} 件' -f $errs.Count) }
        $result.Add([pscustomobject]@{
            Day = $d; Class = $cls; No = $tk.No; Name = $tk.Name; Full = $fullDay[$d]
            Start = $(if ($s) { $s.Time } else { $null }); End = $(if ($e) { $e.Time } else { $null })
            StartEv = $s; EndEv = $e; Dur = $dur; NS = $st.Count; NE = $en.Count; Note = ($note -join '／')
        })
    }
}

Write-Step '01〜03 の CSV を出力しています…'
# ---------- 01 処理時間一覧 ----------
Write-Csv (Join-Path $OutDir '01_処理時間一覧.csv') `
    @('業務日','曜日','区分','No','タスク名','開始','終了','処理時間','処理時間(秒)','開始ログ件数','終了ログ件数','備考') `
    ($result | ForEach-Object {
        ,@($_.Day.ToString('yyyy/MM/dd'), $WeekJa[[string]$_.Day.DayOfWeek], $_.Class, $_.No, $_.Name,
           (Fmt-Dt $_.Start), (Fmt-Dt $_.End), (Fmt-Span $_.Dur),
           $(if ($_.Dur) { [int]$_.Dur.TotalSeconds } else { '' }), $_.NS, $_.NE, $_.Note) })

# ---------- 02 処理時間表（区分別） ----------
foreach ($cls in @('M', 'MTWTFS', 'S')) {
    $r = @($result | Where-Object { $_.Class -eq $cls })
    if ($r.Count -eq 0) { continue }
    $cd = @($r | ForEach-Object { $_.Day } | Sort-Object -Unique)
    $hdr = @('No', 'タスク名') + ($cd | ForEach-Object { $_.ToString('yyyy/MM/dd') + '(' + $WeekJa[[string]$_.DayOfWeek] + ')' })
    $rows = foreach ($tk in $Tasks) {
        $cells = foreach ($d in $cd) {
            $x = $r | Where-Object { $_.Day -eq $d -and $_.No -eq $tk.No } | Select-Object -First 1
            if (-not $x) { '' } elseif ($x.Dur) { Fmt-Span $x.Dur } else { $x.Note }
        }
        if (@($cells | Where-Object { $_ }).Count -eq 0) { continue }
        ,(@($tk.No, $tk.Name) + @($cells))
    }
    Write-Csv (Join-Path $OutDir ('02_処理時間表_' + $cls + '.csv')) $hdr $rows
}

# ---------- 03 最遅／最速 ----------
$baseRows = @($result | Where-Object { $_.Name -eq $BaseTask -and $_.End -and $_.Start -and $_.Full })
if ($baseRows.Count -eq 0) {
    Write-Warning "基準タスク「$BaseTask」の開始・終了がそろった日がないため、03 は出力しません。"
} else {
    $key = if ($CompareBy -eq '処理時間') { { $_.Dur.Ticks } } else { { ($_.End - $_.Day).Ticks } }
    # 判定する単位（グループ）を決める
    $grpOf = switch ($GroupBy) {
        '月×区分' { { $_.Day.ToString('yyyy/MM') + ' ' + $_.Class } }
        '区分'    { { $_.Class } }
        default   { { $_.Day.ToString('yyyy/MM') } }          # '月'
    }
    $fmtDay = { param($d) '{0}({1}・{2})' -f $d.ToString('yyyy/MM/dd'), $WeekJa[[string]$d.DayOfWeek], (Get-Class $d) }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($g in ($baseRows | Group-Object -Property $grpOf | Sort-Object Name)) {
        $b = @($g.Group | Sort-Object $key)
        $fast = $b[0]; $slow = $b[$b.Count - 1]
        foreach ($tk in $Tasks) {
            $xs = $result | Where-Object { $_.Day -eq $slow.Day -and $_.No -eq $tk.No } | Select-Object -First 1
            $xf = $result | Where-Object { $_.Day -eq $fast.Day -and $_.No -eq $tk.No } | Select-Object -First 1
            if (-not $xs -and -not $xf) { continue }
            $diff = if ($xs -and $xf -and $xs.Dur -and $xf.Dur) { Fmt-Span ($xs.Dur - $xf.Dur) } else { '' }
            $out.Add(@($g.Name, $b.Count, $tk.No, $tk.Name,
                (& $fmtDay $slow.Day), (Fmt-Dt $xs.Start), (Fmt-Dt $xs.End), (Fmt-Span $xs.Dur),
                (& $fmtDay $fast.Day), (Fmt-Dt $xf.Start), (Fmt-Dt $xf.End), (Fmt-Span $xf.Dur),
                $diff))
        }
    }
    Write-Csv (Join-Path $OutDir ('03_最遅最速_' + $BaseTask + '.csv')) `
        @($GroupBy,'対象日数','No','タスク名','最遅日','開始','終了','処理時間','最速日','開始','終了','処理時間','差（最遅－最速）') $out
    Write-Info ('最遅／最速の判定 : {0}ごとに、基準タスク「{1}」の{2}' -f $GroupBy, $BaseTask, $CompareBy)
}

Write-Step '切り出しファイルを出力しています…'
New-Item -ItemType Directory -Path $CutDir -Force | Out-Null
# ---------- 切り出し（前任者と同じ形式） ----------
foreach ($cls in @('M', 'MTWTFS', 'S')) {
    foreach ($tk in $Tasks) {
        $r = @($result | Where-Object { $_.Class -eq $cls -and $_.No -eq $tk.No })
        if ($r.Count -eq 0) { continue }
        foreach ($kind in @(@{ n = 1; l = '開始'; p = 'StartEv' }, @{ n = 2; l = '終了'; p = 'EndEv' })) {
            $newLine = @{}   # 業務日 → 行
            foreach ($x in $r) {
                $ev = $x.($kind.p)
                if ($ev) { $newLine[$x.Day] = '{0} OCULUS-{1}{2}{3}' -f $ev.Time.ToString('yyyy-MM-dd HH:mm:ss'), $ev.Level, $ev.Id, $ev.Text }
            }
            if ($newLine.Count -eq 0) { continue }
            $fn = '{0}-{1}-{2} {3}_{4}.txt' -f $cls, $tk.No, $kind.n, $tk.Name, $kind.l
            $fp = Join-Path $CutDir $fn
            # 既にファイルがある場合は置き換えずに更新する：
            # 今回の集計に無い業務日の行はそのまま残し、同じ業務日の行は今回の内容にする
            $byDayLine = @{}
            if (Test-Path $fp) {
                foreach ($ol in [System.IO.File]::ReadAllLines($fp, (Get-FileEncoding $fp))) {
                    $ot = [DateTime]::MinValue
                    if ($ol.Length -ge 19 -and [DateTime]::TryParseExact($ol.Substring(0, 19), 'yyyy-MM-dd HH:mm:ss', $Inv, 'None', [ref]$ot)) {
                        $byDayLine[(Get-BizDay $ot)] = $ol
                    }
                }
            }
            foreach ($d in $newLine.Keys) { $byDayLine[$d] = $newLine[$d] }
            $all = @($byDayLine.Keys | Sort-Object | ForEach-Object { $byDayLine[$_] })
            [System.IO.File]::WriteAllLines($fp, [string[]]$all, $OutEnc)
        }
    }
}

Write-Info ('業務日           : {0} 〜 {1}（{2} 日）' -f $days[0].ToString('yyyy/MM/dd'), $days[-1].ToString('yyyy/MM/dd'), @($days).Count)
Write-Host ''
Write-Info ('出力先 : ' + $OutDir)
Write-Host ''
Write-Host ('  ★ 完了しました（{0:mm\:ss}）。すべてのファイルを出力しました。' -f $sw0.Elapsed) -ForegroundColor Green
