<#
    libInfo.txt 変換処理（手順書 ④ (1)～(9)）
    テキスト入力 → テキスト出力。Excel は不要。

    -Mode check    調査のみ。ファイルは作成しない
    -Mode convert  変換を実行する
#>
param(
    [string] $Path,
    [ValidateSet('check','convert')][string] $Mode = 'check',
    [string] $InCharset  = 'shift_jis',   # 入力の文字コード。手順書の指定は SJIS
    [string] $OutCharset = 'shift_jis',   # 出力の文字コード ※未確認。UTF-8 にする場合はここ
    [switch] $KeepOld,                    # 付けると "old" に当たった行を削除せず残す（目視確認用）

    # 行末スペースの扱い（手順書(8)）
    #   manual = 1個だけ削除。手作業の結果を再現する。
    #            手順書では「置換の繰返し」が未チェックのため、
    #            サクラエディタの すべて置換 は1個しか消さない。
    #            正解ファイルと1バイトも違わない状態にしたい場合はこちら。
    #   all    = 全部削除（RTrim）。データはきれいになるが、
    #            手作業の結果とは全行で差分が出る。
    #   none   = 何もしない。手順書(8)を適用しない
    [ValidateSet('manual','all','none')][string] $TailSpace = 'manual'
)

$ErrorActionPreference = 'Stop'

# このスクリプトは script\ に置かれている。入力の探索と出力はツールのルート基準にする
$ToolRoot = Split-Path -Parent $PSScriptRoot
$OutDir   = Join-Path $ToolRoot 'output'
$Stamp    = Get-Date -Format 'yyyyMMddHHmm'     # 指摘6・7 の YYYYMMDDHHII（12桁）

# ---- 削除パターン（手順書 (2)～(7)(9)）----
# すべて「部分一致・大文字小文字を区別しない」の文字列比較。正規表現は使わない。
# (6) の bk. / org. の「.」を文字そのものとして扱う必要があるため。
# 正規表現にすると「.」が任意1文字になり、対象外のはずの ORGXX まで削除してしまう。
$Patterns = @(
    '\log\',
    '\save\',
    '\common\tools\JudgedCIF\data\output',
    'コピー',
    'bk.',
    '_bk',
    'bkup',
    'bak',
    'org.',
    'old',
    'origin.txt'
)
$OldIndex = $Patterns.IndexOf('old')

# ---- 例外リスト：削除パターンに当たっても削除しない行 ----
# 手順書の備考「←PKG標準資材の"〜ORGXX"は対象外」に相当する。
# org 付きでも正規の資材で、削除してはいけないものがある。
# ここにキーワード（部分一致・大文字小文字区別なし）を書くと、その行は保護される。
$KeepPatterns = @(
    # 指摘4「保護不要」により、登録はゼロ。
    # 仕組みだけ残してあるので、必要になったらここにキーワードを書けば保護できる。
    # 削除してはいけない TextNormalize_ORG1.xml のような ORG 付きファイルは、
    # org. が「org」＋ピリオドの一致であるため、そもそも削除対象にならない。
)

function Get-Enc {
    param([string]$Name)
    switch ($Name.ToLower()) {
        'shift_jis' { [System.Text.Encoding]::GetEncoding(932) }
        'sjis'      { [System.Text.Encoding]::GetEncoding(932) }
        'utf-8'     { New-Object System.Text.UTF8Encoding($false) }
        'utf-8-bom' { New-Object System.Text.UTF8Encoding($true) }
        default     { throw "未対応の文字コードです: $Name" }
    }
}

# 1行の変換：(1) 半角スペース3つ→タブ、(8) 行末スペースの処理。順序は手順書どおり
function Convert-Line {
    param([string]$s)
    $s = $s -replace '   ', "`t"                    # (1) 半角スペース3つ → タブ
    if ($TailSpace -eq 'all') {
        $s = $s.TrimEnd(' ')                        # (8) 全部削除（スペースのみ。タブは消さない）
    } elseif ($TailSpace -eq 'manual') {
        # (8) 1個だけ削除 … サクラエディタの すべて置換 の実際の挙動を再現
        if ($s.EndsWith(' ')) { $s = $s.Substring(0, $s.Length - 1) }
    }
    # none … 何もしない
    return $s
}

# ---- 入力ファイルの決定 ----
# 引数が無ければ、同じフォルダの libInfo*.txt を探す。
# 実ファイル名は libInfo_横東_開発AP#1.txt のように後ろが付くため、
# 'libInfo.txt' 固定では見つからない。
if (-not $Path) {
    $cand = @(Get-ChildItem -LiteralPath $ToolRoot -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -like 'libInfo*.txt' -and
                             $_.Name -notlike '*_converted.txt' -and
                             $_.Name -notlike '*_old_candidates.txt' -and
                             $_.Name -notlike '*_変換後.txt' } |
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
        Write-Host ' 入力ファイルが見つかりません。' -ForegroundColor Yellow
        Write-Host ''
        Write-Host ' 対処:'
        Write-Host '   ・libInfo のファイルを「ここにlibInfo.txtをドラッグ＆ドロップしてください.bat」に乗せる'
        Write-Host '   ・または libInfo*.txt をツールのフォルダに置いて実行する'
        Write-Host ''
        Write-Host ' このフォルダにある .txt ファイル:'
        $txt = @(Get-ChildItem -LiteralPath $ToolRoot -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -like '*.txt' })
        if ($txt.Count -eq 0) { Write-Host '   (なし)' }
        else { $txt | ForEach-Object { Write-Host ('   ' + $_.Name) } }
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

$enc   = Get-Enc $InCharset
$all   = $enc.GetString([System.IO.File]::ReadAllBytes($Path))
$lines = ($all -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' libInfo.txt 変換処理' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  入力         : $Path"
Write-Host "  文字コード   : 入力 $InCharset / 出力 $OutCharset"
Write-Host "  モード       : $Mode"
$tsLabel = switch ($TailSpace) {
    'all'    { 'all（全部削除）' }
    'none'   { 'none（何もしない）' }
    default  { 'manual（1個だけ削除）' }
}
Write-Host ("  行末スペース : {0}" -f $tsLabel)

# ---- 1行ずつ処理 ----
$hit      = New-Object 'int[]' $Patterns.Count
$out      = New-Object 'System.Collections.Generic.List[string]'
$delBuf   = New-Object 'System.Collections.Generic.List[string]'   # 指摘4(b)・7: 削除した行を全部ここに
$colCount = @{}
$total = 0; $blank = 0; $deleted = 0; $kept = 0
$sp4 = 0; $tailSpaceLines = 0; $maxTail = 0; $folderCnt = 0; $tailTab = 0
$tailDist = @{}     # 行末スペースが何個の行が何行あるか
$firstLine = if ($lines.Count -gt 0) { $lines[0] } else { '' }

for ($i = 1; $i -lt $lines.Count; $i++) {     # 1 から開始 = 1行目を読み飛ばす
    $s = $lines[$i]
    if ($s -eq '') { $blank++; continue }
    $total++

    # --- 調査用のカウント ---
    if ($s.Contains('    '))                                   { $sp4++ }
    $n = $s.Length - $s.TrimEnd(' ').Length
    if ($n -gt 0) { $tailSpaceLines++; if ($n -gt $maxTail) { $maxTail = $n } }
    if ($tailDist.ContainsKey($n)) { $tailDist[$n]++ } else { $tailDist[$n] = 1 }
    if ($s.IndexOf('folder', [StringComparison]::OrdinalIgnoreCase) -ge 0) { $folderCnt++ }

    # --- 削除判定 ---
    # まず例外を見る。例外リストにある行は、何に当たっても残す
    $protected = $false
    foreach ($kp in $KeepPatterns) {
        if ($s.IndexOf($kp, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $protected = $true; break }
    }

    $killed = $false; $onlyOld = $false
    $matched = New-Object 'System.Collections.Generic.List[string]'
    for ($j = 0; $j -lt $Patterns.Count; $j++) {
        if ($s.IndexOf($Patterns[$j], [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $hit[$j]++
            [void]$matched.Add($Patterns[$j])
            if ($j -eq $OldIndex) { $onlyOld = $true } else { $killed = $true }
        }
    }
    # "old" は部分一致なので folder / holder などにも当たる。
    # 既定は手順書どおり削除。誤爆が心配な場合は -KeepOld を付けると削除しない。
    if ($onlyOld -and (-not $KeepOld)) { $killed = $true }

    if ($protected) { $killed = $false; $kept++ }
    if ($killed) {
        $deleted++
        # 削除した行はすべて変更ログへ。行頭にどのパターンで消えたかを付ける
        [void]$delBuf.Add('[' + ($matched -join '][') + "]`t" + $s)
        continue
    }

    $conv = Convert-Line $s
    [void]$out.Add($conv)
    $c = ($conv -split "`t").Count
    if ($colCount.ContainsKey($c)) { $colCount[$c]++ } else { $colCount[$c] = 1 }
    if ($conv.EndsWith("`t")) { $tailTab++ }
}

# ---- 調査結果 ----
# 指摘3: 変換実行後は調査結果を再表示しない。check のときだけ出す
if ($Mode -eq 'check') {
Write-Host ''
Write-Host '--- 調査結果 ---' -ForegroundColor Cyan
Write-Host ("  1行目(削除対象) : {0}" -f $firstLine)
Write-Host ("  全行数(1行目除く): {0}   空行: {1}" -f $total, $blank)
Write-Host ''
function Show-Chk {
    param([string]$Label, [int]$Val, [string]$Warn)
    if ($Val -eq 0) { Write-Host ("  {0,-34}: {1}" -f $Label, $Val) -ForegroundColor Green }
    else            { Write-Host ("  {0,-34}: {1}   ★{2}" -f $Label, $Val, $Warn) -ForegroundColor Yellow }
}
Show-Chk '半角スペース4個以上を含む行'  $sp4       '6個あるとタブ2個になり、列がずれる'
Show-Chk '行末にスペースがある行'        $tailSpaceLines "最大 $maxTail 個"
# 何個の行が何行あるかを出す。正解ファイルと合わない時の切り分けに使う
Write-Host '    行末スペースの個数の内訳:'
foreach ($k in ($tailDist.Keys | Sort-Object)) {
    Write-Host ("      {0} 個 : {1} 行" -f $k, $tailDist[$k])
}
Show-Chk '変換後に行末がタブになる行'    $tailTab   '表に貼ると空列が1つ増える'
Show-Chk "'folder' を含む行"             $folderCnt "old の部分一致で巻き添え削除される"

Write-Host ''
Write-Host '  変換後のタブ区切り列数:'
foreach ($k in ($colCount.Keys | Sort-Object)) {
    Write-Host ("    {0} 列 : {1} 行" -f $k, $colCount[$k])
}

Write-Host ''
Write-Host '  削除パターン別の該当行数:'
for ($j = 0; $j -lt $Patterns.Count; $j++) {
    $note = if ($j -eq $OldIndex) { if ($KeepOld) { '   ※削除せず別ファイルへ' } else { '   ※削除。記録も残す' } } else { '' }
    Write-Host ("    {0,-38} {1}{2}" -f $Patterns[$j], $hit[$j], $note)
}
Write-Host '    (1行が複数パターンに当たることがあるため、合計は削除数と一致しません)'

    Write-Host ''
    Write-Host " 調査のみ完了。ファイルは作成していません。" -ForegroundColor Green
    Write-Host " 上の ★ が付いた項目を確認してから変換してください。"
    exit 0
}

# ---- 出力（指摘6・7・8）----
# 出力先は常にツールの output\ 配下。ファイル名にタイムスタンプを付けて上書きを防ぐ
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$base    = [System.IO.Path]::GetFileNameWithoutExtension($Path)
$outPath = Join-Path $OutDir ("{0}_変換後_{1}.txt"   -f $base, $Stamp)
$logPath = Join-Path $OutDir ("{0}_変更ログ_{1}.txt" -f $base, $Stamp)

$oenc = Get-Enc $OutCharset
[System.IO.File]::WriteAllBytes($outPath, $oenc.GetBytes(($out -join "`r`n") + "`r`n"))

# 変更ログ：削除した行を全部、どのパターンで消えたかを付けて出力
$logHead = @(
    "変更ログ  $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')",
    "入力 : $Path",
    "読込 $total 行 / 削除 $($delBuf.Count) 行 / 出力 $($out.Count) 行",
    "",
    "行頭の [ ] は、その行を削除した削除パターン。",
    "タブ変換・行末スペース削除は全行に掛かる処理のため記録していません。",
    "------------------------------------------------------------"
)
[System.IO.File]::WriteAllBytes($logPath, $oenc.GetBytes((($logHead + $delBuf) -join "`r`n") + "`r`n"))

Write-Host ''
Write-Host '--- 変換完了 ---' -ForegroundColor Cyan
Write-Host ("  読込 {0} 行 / 削除 {1} 行 / 出力 {2} 行" -f $total, $deleted, $out.Count)
if ($kept -gt 0) {
    Write-Host ("  例外で保護 {0} 行" -f $kept) -ForegroundColor Cyan
}
Write-Host ("  変換後   : {0}" -f $outPath) -ForegroundColor Green
Write-Host ("  変更ログ : {0}" -f $logPath) -ForegroundColor Green
Write-Host ''
Write-Host ' 結果 : OK' -ForegroundColor Green
exit 0
