<#
    libInfo.txt 变换处理（手順書 ④ (1)~(9)）
    输入 txt -> 输出 txt。不需要 Excel。

    -Mode check    只调查，不生成文件
    -Mode convert  执行变换
#>
param(
    [string] $Path,
    [ValidateSet('check','convert')][string] $Mode = 'check',
    [string] $InCharset  = 'shift_jis',   # 输入编码。手順書要求 SJIS
    [string] $OutCharset = 'shift_jis',   # 输出编码。※未确认，要 UTF-8 就改这里
    [switch] $KeepOld                     # 加上这个开关，"old" 命中的行先不删（留着人工确认）
)

$ErrorActionPreference = 'Stop'

# ---- 11 个删除模式（手順書 (2)~(7)(9)）----
# 全部是「部分一致 + 不区分大小写」的字面匹配，不用正则。
# (6) 的 bk. / org. 里的「.」必须当成字面的点：
# 用正则的话「.」变成通配符，会把备注写明「対象外」的 ORGXX 一起删掉。
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

function Get-Enc {
    param([string]$Name)
    switch ($Name.ToLower()) {
        'shift_jis' { [System.Text.Encoding]::GetEncoding(932) }
        'sjis'      { [System.Text.Encoding]::GetEncoding(932) }
        'utf-8'     { New-Object System.Text.UTF8Encoding($false) }
        'utf-8-bom' { New-Object System.Text.UTF8Encoding($true) }
        default     { throw "不支持的编码: $Name" }
    }
}

# 一行的变换：(1) 3个空格->Tab、(8) 删行末空格
# 顺序照手順書。TrimEnd(' ') 只删空格不删 Tab，和 VBA 的 RTrim 一致。
function Convert-Line { param([string]$s) ($s -replace '   ', "`t").TrimEnd(' ') }

# ---- 找输入文件 ----
if (-not $Path) {
    $here = Join-Path $PSScriptRoot 'libInfo.txt'
    if (Test-Path -LiteralPath $here) { $Path = $here }
    else { throw "没找到输入文件。把 libInfo.txt 拖到 run.bat 上，或放到脚本同一个文件夹里。" }
}
if (-not (Test-Path -LiteralPath $Path)) { throw "文件不存在: $Path" }

$enc   = Get-Enc $InCharset
$all   = $enc.GetString([System.IO.File]::ReadAllBytes($Path))
$lines = ($all -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' libInfo.txt 变换处理' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  输入   : $Path"
Write-Host "  编码   : 输入 $InCharset / 输出 $OutCharset"
Write-Host "  模式   : $Mode"

# ---- 逐行处理 ----
$hit      = New-Object 'int[]' $Patterns.Count
$out      = New-Object 'System.Collections.Generic.List[string]'
$oldBuf   = New-Object 'System.Collections.Generic.List[string]'
$colCount = @{}
$total = 0; $blank = 0; $deleted = 0
$sp4 = 0; $tailSpace = 0; $maxTail = 0; $folderCnt = 0; $tailTab = 0
$firstLine = if ($lines.Count -gt 0) { $lines[0] } else { '' }

for ($i = 1; $i -lt $lines.Count; $i++) {     # 从 1 开始 = 跳过第 1 行
    $s = $lines[$i]
    if ($s -eq '') { $blank++; continue }
    $total++

    # --- 调查用的计数 ---
    if ($s.Contains('    '))                                   { $sp4++ }
    $n = $s.Length - $s.TrimEnd(' ').Length
    if ($n -gt 0) { $tailSpace++; if ($n -gt $maxTail) { $maxTail = $n } }
    if ($s.IndexOf('folder', [StringComparison]::OrdinalIgnoreCase) -ge 0) { $folderCnt++ }

    # --- 删除判定 ---
    $killed = $false; $onlyOld = $false
    for ($j = 0; $j -lt $Patterns.Count; $j++) {
        if ($s.IndexOf($Patterns[$j], [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $hit[$j]++
            if ($j -eq $OldIndex) { $onlyOld = $true } else { $killed = $true }
        }
    }
    # "old" 是部分一致，folder / holder 之类也会被命中。
    # 默认照手順書删掉，但同时写到 _old_candidates.txt 留档，事后能查。
    # 担心误伤时加 -KeepOld，这些行就不删，先人工过目。
    if ($onlyOld) {
        [void]$oldBuf.Add($s)
        if (-not $KeepOld) { $killed = $true }
    }

    if ($killed) { $deleted++; continue }

    $conv = Convert-Line $s
    [void]$out.Add($conv)
    $c = ($conv -split "`t").Count
    if ($colCount.ContainsKey($c)) { $colCount[$c]++ } else { $colCount[$c] = 1 }
    if ($conv.EndsWith("`t")) { $tailTab++ }
}

# ---- 调查报告 ----
Write-Host ''
Write-Host '--- 调查结果 ---' -ForegroundColor Cyan
Write-Host ("  第1行(将被删除) : {0}" -f $firstLine)
Write-Host ("  总行数(不含第1行): {0}   空行: {1}" -f $total, $blank)
Write-Host ''
function Show-Chk {
    param([string]$Label, [int]$Val, [string]$Warn)
    if ($Val -eq 0) { Write-Host ("  {0,-30}: {1}" -f $Label, $Val) -ForegroundColor Green }
    else            { Write-Host ("  {0,-30}: {1}   ★{2}" -f $Label, $Val, $Warn) -ForegroundColor Yellow }
}
Show-Chk '含4个以上连续空格的行'  $sp4       '6个空格会变成2个Tab，列会错位'
Show-Chk '行末有空格的行'          $tailSpace "最多 $maxTail 个"
Show-Chk '变换后行末是Tab的行'     $tailTab   '贴进表格会多一个空列'
Show-Chk "含 folder 的行"          $folderCnt 'old 会把这些一起删掉'

Write-Host ''
Write-Host '  变换后的列数分布:'
foreach ($k in ($colCount.Keys | Sort-Object)) {
    Write-Host ("    {0} 列 : {1} 行" -f $k, $colCount[$k])
}

Write-Host ''
Write-Host '  各删除模式命中行数:'
for ($j = 0; $j -lt $Patterns.Count; $j++) {
    $note = if ($j -eq $OldIndex) { if ($KeepOld) { '   ※不删，另存' } else { '   ※删除，并留档' } } else { '' }
    Write-Host ("    {0,-38} {1}{2}" -f $Patterns[$j], $hit[$j], $note)
}
Write-Host '    (一行可能命中多个模式，合计不等于删除数)'

if ($Mode -eq 'check') {
    Write-Host ''
    Write-Host " 调查完成。没有生成任何文件。" -ForegroundColor Green
    Write-Host " 上面 ★ 的项目确认好之后，再执行变换。"
    exit 0
}

# ---- 输出 ----
$dir  = Split-Path -Parent $Path
$base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
$outPath = Join-Path $dir "$base`_converted.txt"
$oldPath = Join-Path $dir "$base`_old_candidates.txt"

$oenc = Get-Enc $OutCharset
[System.IO.File]::WriteAllBytes($outPath, $oenc.GetBytes(($out -join "`r`n") + "`r`n"))
if ($oldBuf.Count -gt 0) {
    [System.IO.File]::WriteAllBytes($oldPath, $oenc.GetBytes(($oldBuf -join "`r`n") + "`r`n"))
}

Write-Host ''
Write-Host '--- 变换完成 ---' -ForegroundColor Cyan
Write-Host ("  读入 {0} 行 / 删除 {1} 行 / 输出 {2} 行" -f $total, $deleted, $out.Count)
Write-Host ("  输出 : {0}" -f $outPath) -ForegroundColor Green
if ($oldBuf.Count -gt 0) {
    Write-Host ''
    if ($KeepOld) {
        Write-Host ("  ★ 只被 old 命中的 {0} 行【没有删除】，写到了：" -f $oldBuf.Count) -ForegroundColor Yellow
        Write-Host ("    {0}" -f $oldPath) -ForegroundColor Yellow
        Write-Host '    确认这些确实该删之后，去掉 -KeepOld 再跑一次。'
    } else {
        Write-Host ("  只被 old 命中的 {0} 行已删除。留档在：" -f $oldBuf.Count)
        Write-Host ("    {0}" -f $oldPath)
        Write-Host '    担心误伤（folder / holder 之类）时，加 -KeepOld 重跑可以先不删。'
    }
}
Write-Host ''
Write-Host ' 结果 : OK' -ForegroundColor Green
exit 0
