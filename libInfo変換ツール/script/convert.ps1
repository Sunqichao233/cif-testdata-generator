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
    #
    #   ★既定は manual（手順書(8)どおり）
    #     実データの行末スペースは1個で、手作業の正解ファイルにも1個残っている。
    #     そのため manual にすると、正解ファイルとは全行で行末スペース1個分の差分が出る。
    #     正解ファイル作成時に手順書(8)が実施されていないためと思われる。
    #     正解ファイルに合わせたい場合は none にする。
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
$OrgIndex = $Patterns.IndexOf('org.')   # 指摘2: org. だけは該当行を画面に出す

# ---- 例外リスト：削除パターンに当たっても削除しない行 ----
# 手順書の備考「←PKG標準資材の"〜ORGXX"は対象外」に相当する。
# org 付きでも正規の資材で、削除してはいけないものがある。
# ここにキーワード（部分一致・大文字小文字区別なし）を書くと、その行は保護される。
$KeepPatterns = @(
    # 削除パターンに当たっても削除しない行のキーワード（部分一致・大小文字区別なし）
    #
    # レビュー指摘4「keepParamにある AfterRcv_org.bat は削除OK」の確認が取れたため、
    # 登録はゼロにしている。下記2行は org. に当たり、削除される。
    #     \common\batches\CIF_AfterRcv_org.bat
    #     \common\batches\USERFILE_AfterRcv_org.bat
    # 手作業の正解ファイルには残っているため、突き合わせでは2行の差分として出る。
    # これは正解ファイル側の消し忘れであり、正しい差分。
    #
    # 削除してはいけない TextNormalize_ORG1.xml のような ORG 付きファイルは、
    # org. が「org」＋ピリオドの一致であるため、そもそも削除対象にならない。
    # 仕組みは残してあるので、保護が必要になったらここにキーワードを書く。
)

# ---- 画面の桁そろえ ----
# PowerShell の -f の {0,-34} は「文字数」で詰めるため、全角文字が混ざると桁がそろわない。
# コマンドプロンプトでは全角1文字＝半角2文字分の幅になるので、
# 表示幅を数えて半角スペースを足す。
function Get-DispWidth {
    param([string]$s)
    $w = 0
    foreach ($ch in $s.ToCharArray()) {
        $c = [int]$ch
        if ($c -ge 0x1100 -and (
              $c -le 0x115F -or
             ($c -ge 0x2E80 -and $c -le 0xA4CF -and $c -ne 0x303F) -or
             ($c -ge 0xAC00 -and $c -le 0xD7A3) -or
             ($c -ge 0xF900 -and $c -le 0xFAFF) -or
             ($c -ge 0xFE30 -and $c -le 0xFE6F) -or
             ($c -ge 0xFF00 -and $c -le 0xFF60) -or
             ($c -ge 0xFFE0 -and $c -le 0xFFE6))) { $w += 2 } else { $w += 1 }
    }
    return $w
}
function Format-Pad {
    param([string]$Text, [int]$Width)
    $n = $Width - (Get-DispWidth $Text)
    if ($n -lt 0) { $n = 0 }
    return ($Text + (' ' * $n))
}

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
$hitKept  = New-Object 'int[]' $Patterns.Count   # 当たったが例外で保護した行数
$out      = New-Object 'System.Collections.Generic.List[string]'
$delBuf   = New-Object 'System.Collections.Generic.List[string]'   # 指摘4(b)・7: 削除した行を全部ここに
# 指摘2: 「org.」に当たった行は本文を画面に出す。
#        ORG を含むが削除してはいけない行（PKG標準資材の ～ORGXX など）を
#        巻き込んでいないか、変換前に目視で確認するため。
$orgBuf   = New-Object 'System.Collections.Generic.List[string]'
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
            if ($j -eq $OrgIndex) {
                # 指摘2: org. の該当行は本文を控えておく
                $mark = if ($protected) { '[保護・残す] ' } else { '[削除] ' }
                [void]$orgBuf.Add($mark + $s)
            }
            if ($j -eq $OldIndex) { $onlyOld = $true } else { $killed = $true }
        }
    }
    # "old" は部分一致なので folder / holder などにも当たる。
    # 既定は手順書どおり削除。誤爆が心配な場合は -KeepOld を付けると削除しない。
    if ($onlyOld -and (-not $KeepOld)) { $killed = $true }

    if ($protected) {
        $killed = $false
        if ($matched.Count -gt 0) {
            $kept++
            # どのパターンに当たったのに残したのかを記録する
            foreach ($m in $matched) { $hitKept[$Patterns.IndexOf($m)]++ }
        }
    }
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
Write-Host ("  {0}: {1}" -f (Format-Pad '1行目(削除対象)'  18), $firstLine)
Write-Host ("  {0}: {1}   空行: {2}" -f (Format-Pad '全行数(1行目除く)' 18), $total, $blank)
Write-Host ''
function Show-Chk {
    param([string]$Label, [int]$Val, [string]$Warn)
    $L = Format-Pad $Label 34
    if ($Val -eq 0) { Write-Host ("  {0}: {1}" -f $L, $Val) -ForegroundColor Green }
    else            { Write-Host ("  {0}: {1}   ★{2}" -f $L, $Val, $Warn) -ForegroundColor Yellow }
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
Write-Host '  削除パターン別の該当行数（当たった行数。削除数とは別）:'
for ($j = 0; $j -lt $Patterns.Count; $j++) {
    $note = ''
    if ($hitKept[$j] -gt 0) {
        # 当たったのに残した行がある場合は必ず書く。削除済みと誤解されないように
        $note = "   ※うち {0} 行は例外リストで保護（削除しない）" -f $hitKept[$j]
    } elseif ($j -eq $OldIndex -and $KeepOld) {
        $note = '   ※削除せず別ファイルへ'
    }
    Write-Host ("    {0} {1,4}{2}" -f (Format-Pad $Patterns[$j] 38), $hit[$j], $note)

    # 指摘2: org. は該当行の本文も出す。
    # ORG を含むが削除してはいけない行を巻き込んでいないか、ここで目視確認する。
    if ($j -eq $OrgIndex -and $orgBuf.Count -gt 0) {
        foreach ($line in $orgBuf) {
            Write-Host ('      ' + $line) -ForegroundColor Yellow
        }
    }
}
Write-Host '    (1行が複数パターンに当たることがあるため、合計は削除数と一致しません)'
Write-Host ''
Write-Host ("  実際に削除する行数 : {0} 行" -f $deleted) -ForegroundColor Cyan
if ($kept -gt 0) {
    Write-Host ("  例外で保護する行数 : {0} 行（上の ※ の行）" -f $kept) -ForegroundColor Cyan
}
Write-Host ("  出力される行数     : {0} 行" -f $out.Count) -ForegroundColor Cyan

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
