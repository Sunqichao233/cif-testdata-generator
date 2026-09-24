<#
.SYNOPSIS
    CIF 测试数据生成（Tab 分隔 27 列）

.DESCRIPTION
    按 layout.csv 的定义生成 Tab 分隔的测试数据，UTF-8 输出。

    每一列的内容由 layout.csv 的「種別」决定：
      序号      递增序号，「値」写位数（补零）
      固定      固定值，直接写在「値」里
      姓名漢字  自动生成的汉字姓名（-NameMode kanji 时填）
      姓名カナ  自动生成的片假名姓名（-NameMode kana 时填）
      空        留空

    要改某一列的内容 → 只改 layout.csv，不用动脚本。

.EXAMPLE
    .\make_cif.ps1                        生成 1 条（和「正确的数据」对照用）
    .\make_cif.ps1 -Count 40000           生成 4 万条
    .\make_cif.ps1 -Count 100 -NameMode kana   姓名用片假名
#>
param(
    [string] $Layout,     # 列定义 CSV（默认：脚本同目录 layout.csv）
    [string] $Names,      # 姓名素材 CSV（默认：同目录 names.csv）
    [string] $OutDir,     # 输出目录（默认：同目录 output）
    [string] $YMD,        # 文件名里的日期 YYYYMMDD（默认：当天）
    [string] $FileName,   # 输出文件名（默认：CIF_<YMD>.txt）

    # 生成条数，最大 40000
    [int] $Count = 1,

    # 递增序号的起始值
    [int] $StartNo = 1,

    # 输出编码。要求是 UTF-8
    [ValidateSet('utf-8','utf-8-bom')]
    [string] $Encoding = 'utf-8',

    # 姓名用汉字还是片假名。二选一，不混用
    [ValidateSet('kanji','kana')]
    [string] $NameMode = 'kanji'
)

$ErrorActionPreference = 'Stop'
$MAX_COUNT = 40000          # 条数上限
$ZEN_SPACE = [char]0x3000   # 全角空格。姓和名之间用这个

# =============================================================
# 通用函数
# =============================================================
function Write-Head { param([string]$m) Write-Host ''; Write-Host $m -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-NG   { param([string]$m) Write-Host "  [NG]   $m" -ForegroundColor Red }

function Get-Gcd {
    # 最大公约数。用来找一个和组合数互质的步长（姓名分散用）
    param([long]$a, [long]$b)
    while ($b -ne 0) { $t = $b; $b = $a % $b; $a = $t }
    return $a
}

function Read-CsvAuto {
    # 自动判别 CSV 编码后读取（UTF-8 BOM / UTF-8 / Shift_JIS）
    # 注意：Import-Csv -Encoding Default 会跟着系统区域走（中文 Windows 是 936），所以不用它
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "找不到文件: $Path" }
    $bytes = [System.IO.File]::ReadAllBytes($Path)

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        $detected = 'UTF-8 (BOM)'
    } else {
        try {
            # 遇到非法字节会抛异常，能通过说明是 UTF-8
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            $text = $strict.GetString($bytes)
            $detected = 'UTF-8'
        } catch {
            $text = [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
            $detected = 'Shift_JIS'
        }
    }
    Write-Host ("  编码       : {0}  [{1}]" -f $detected, (Split-Path $Path -Leaf))
    return ($text | ConvertFrom-Csv)
}

# =============================================================
# 1. 参数整理
# =============================================================
$here = $PSScriptRoot
if (-not $Layout)   { $Layout   = Join-Path $here 'layout.csv' }
if (-not $Names)    { $Names    = Join-Path $here 'names.csv' }
if (-not $OutDir)   { $OutDir   = Join-Path $here 'output' }
if (-not $YMD)      { $YMD      = (Get-Date).ToString('yyyyMMdd') }
if (-not $FileName) { $FileName = "CIF_$YMD.txt" }

if ($YMD -notmatch '^\d{8}$')  { throw "YMD 要写成 8 位数字(YYYYMMDD): '$YMD'" }
if ($Count -lt 1)              { throw "条数至少是 1: $Count" }
if ($Count -gt $MAX_COUNT)     { throw "条数超过上限: $Count > $MAX_COUNT" }

if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$outPath = Join-Path $OutDir $FileName

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' CIF 测试数据生成（Tab 分隔）' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  列定义     : $Layout"
Write-Host "  输出       : $outPath"
Write-Host "  条数       : $Count"
Write-Host "  起始序号   : $StartNo"
Write-Host "  输出编码   : $Encoding"
Write-Host "  姓名       : $(if ($NameMode -eq 'kanji') { '漢字（例：佐藤　一郎）' } else { 'フリガナ（例：サトウ　イチロウ）' })"

# =============================================================
# 2. 读列定义
# =============================================================
Write-Head '--- 列定义 ---'

$layoutRows = @(Read-CsvAuto $Layout)
if ($layoutRows.Count -eq 0) { throw "layout.csv 是空的" }

# 按「列」排序，并确认列号是 1,2,3... 连续的
$cols = @()
foreach ($r in $layoutRows) {
    $noRaw = "$($r.'列')".Trim()
    if ($noRaw -eq '') { continue }
    if ($noRaw -notmatch '^\d+$') { throw "「列」必须是数字: '$noRaw'" }
    $cols += [pscustomobject]@{
        No   = [int]$noRaw
        Name = "$($r.'項目名')".Trim()
        Kind = "$($r.'種別')".Trim()
        Val  = "$($r.'値')".Trim()
    }
}
$cols = @($cols | Sort-Object No)

for ($i = 0; $i -lt $cols.Count; $i++) {
    if ($cols[$i].No -ne ($i + 1)) {
        throw "列号不连续。第 $($i+1) 个应该是 $($i+1)，实际是 $($cols[$i].No)。"
    }
}
$colCount = $cols.Count
Write-Host ("  列数       : {0} 列（Tab 数 = {1}）" -f $colCount, ($colCount - 1))

# 检查每种「種別」都认识
$known = @('序号','固定','日付','姓名漢字','姓名カナ','空')
foreach ($c in $cols) {
    if ($known -notcontains $c.Kind) {
        throw "第 $($c.No) 列的「種別」不认识: '$($c.Kind)'`n　 只能是: $($known -join ' / ')"
    }
}

# 找出姓名列。二选一：只填一个，另一个留空
$idxKanji = -1; $idxKana = -1
for ($i = 0; $i -lt $colCount; $i++) {
    if ($cols[$i].Kind -eq '姓名漢字') { $idxKanji = $i }
    if ($cols[$i].Kind -eq '姓名カナ') { $idxKana  = $i }
}
$idxName = if ($NameMode -eq 'kanji') { $idxKanji } else { $idxKana }
if ($idxName -lt 0) {
    throw "layout.csv 里没有「$(if($NameMode -eq 'kanji'){'姓名漢字'}else{'姓名カナ'})」这一列。"
}

# 运行时实际有值的列数（姓名是二选一，所以要减掉没用到的那个）
$valued = @($cols | Where-Object { $_.Kind -ne '空' }).Count
if ($idxKanji -ge 0 -and $idxKana -ge 0) { $valued-- }
Write-Host ("  有值的列   : {0} 列（姓名是二选一，另一个留空）" -f $valued)

# =============================================================
# 3. 生成姓名（不重复、不含数字）
# =============================================================
Write-Head '--- 姓名生成 ---'

$nameRows = @(Read-CsvAuto $Names)
$sei = @($nameRows | Where-Object { "$($_.'種別')".Trim() -eq '姓' })
$mei = @($nameRows | Where-Object { "$($_.'種別')".Trim() -eq '名' })
if ($sei.Count -eq 0 -or $mei.Count -eq 0) { throw "names.csv 里没有「姓」或「名」的行。" }

$combo = $sei.Count * $mei.Count
Write-Host ("  素材       : 姓 {0} × 名 {1} = {2} 种组合" -f $sei.Count, $mei.Count, $combo)

if ($Count -gt $combo) {
    throw "组合数不够：要 $Count 条，但只有 $combo 种组合。`n" +
          "　 姓名里不能加数字，所以不能靠后缀凑数。`n" +
          "　 请往 names.csv 里加「姓」或「名」的行。"
}

# 用 List，不要用 $arr += （数组 += 每次整个复制，4 万条会慢 10 倍以上）
$nameList = New-Object 'System.Collections.Generic.List[string]'
$used = New-Object 'System.Collections.Generic.HashSet[string]'

# 姓名分散用的步长。
# 如果老老实实按 0,1,2,... 取组合，会变成「佐藤　太郎 / 鈴木　太郎 / 高橋　太郎…」
# 同一个「名」连着出现 205 次，看起来很假。
# 所以每次跳 stride 格去取。stride 和组合数互质，就能保证不重复地走遍全部组合，
# 又不会扎堆。用黄金比例算出来的步长分散得最均匀。
$stride = [long][Math]::Floor($combo * 0.6180339887)
if ($stride -lt 2) { $stride = 1 }
while ((Get-Gcd $stride $combo) -ne 1) { $stride++ }
Write-Host ("  分散步长   : {0}（和组合数 {1} 互质，保证不重复且不扎堆）" -f $stride, $combo)

$i = 0
$skipped = 0
while ($nameList.Count -lt $Count) {
    if ($i -ge $combo) {
        throw "组合用完了还凑不够 $Count 条（重名跳过了 $skipped 次）。请往 names.csv 加行。"
    }
    $idx = [int](([long]$i * $stride) % $combo)
    $s = $sei[$idx % $sei.Count]
    $m = $mei[[Math]::Floor($idx / $sei.Count)]
    $i++

    # 姓和名之间是全角空格
    $nm = if ($NameMode -eq 'kanji') {
        "$($s.'漢字')$ZEN_SPACE$($m.'漢字')"
    } else {
        "$($s.'カナ')$ZEN_SPACE$($m.'カナ')"
    }

    # 读音相同的姓（渡辺/渡部 都是ワタナベ）在片假名模式下会撞名，跳过即可
    if (-not $used.Add($nm)) { $skipped++; continue }
    [void]$nameList.Add($nm)
}
Write-OK ("生成 {0} 条，全部不重复（跳过重名 {1} 次）" -f $nameList.Count, $skipped)
Write-Host ("  示例       : {0}" -f $nameList[0])

# =============================================================
# 4. 生成文件
# =============================================================
Write-Head '--- 生成 ---'

# 预先算好每一列的固定内容，热循环里只换序号和姓名
$fixed = New-Object string[] $colCount
$seqCols  = New-Object 'System.Collections.Generic.List[int]'
$seqLen   = New-Object 'System.Collections.Generic.List[int]'
$dateCols = New-Object 'System.Collections.Generic.List[int]'
$dateList = New-Object 'System.Collections.Generic.List[string[]]'

for ($c = 0; $c -lt $colCount; $c++) {
    switch ($cols[$c].Kind) {
        '固定' { $fixed[$c] = $cols[$c].Val }
        '序号' {
            $fixed[$c] = ''
            [void]$seqCols.Add($c)
            $w = 6
            if ($cols[$c].Val -match '^\d+$') { $w = [int]$cols[$c].Val }
            [void]$seqLen.Add($w)
        }
        '日付' {
            # 「値」写年份（如 2000），生成该年份里的某一天 YYYYMMDD
            $fixed[$c] = ''
            $y = 2000
            if ($cols[$c].Val -match '^\d{4}$') { $y = [int]$cols[$c].Val }
            # 把这一年的每一天先算好存起来，循环里直接取，比每次算日期快得多
            $days = if ([datetime]::IsLeapYear($y)) { 366 } else { 365 }
            $arr = New-Object string[] $days
            $d0 = New-Object datetime($y, 1, 1)
            for ($d = 0; $d -lt $days; $d++) { $arr[$d] = $d0.AddDays($d).ToString('yyyyMMdd') }
            [void]$dateCols.Add($c)
            [void]$dateList.Add($arr)
        }
        default { $fixed[$c] = '' }   # 空 / 姓名列都先留空
    }
}

# Tab 和换行不能出现在字段里，否则列数会错乱
foreach ($c in 0..($colCount-1)) {
    if ($fixed[$c] -match "[`t`r`n]") {
        throw "第 $($c+1) 列的固定值里有 Tab 或换行，会把列数搞乱: '$($fixed[$c])'"
    }
}

$sb = New-Object System.Text.StringBuilder
$rec = New-Object string[] $colCount
$firstLine = $null

for ($n = 0; $n -lt $Count; $n++) {
    [Array]::Copy($fixed, $rec, $colCount)

    # 递增序号（第 1 列和第 26 列用同一个号）
    $no = $StartNo + $n
    for ($k = 0; $k -lt $seqCols.Count; $k++) {
        $rec[$seqCols[$k]] = ([string]$no).PadLeft($seqLen[$k], '0')
    }

    # 日期。在指定年份里按天循环，让生日分散在全年
    for ($k = 0; $k -lt $dateCols.Count; $k++) {
        $arr = $dateList[$k]
        $rec[$dateCols[$k]] = $arr[$n % $arr.Length]
    }

    # 姓名。二选一，只填一个
    $rec[$idxName] = $nameList[$n]

    $line = [string]::Join("`t", $rec)
    if ($n -eq 0) { $firstLine = $line }
    [void]$sb.Append($line).Append("`r`n")
}

$enc = if ($Encoding -eq 'utf-8-bom') {
    New-Object System.Text.UTF8Encoding($true)
} else {
    New-Object System.Text.UTF8Encoding($false)
}
[System.IO.File]::WriteAllBytes($outPath, $enc.GetBytes($sb.ToString()))

$size = (Get-Item -LiteralPath $outPath).Length
Write-OK "生成完成 : $FileName（$Count 条 / $size 字节）"

# =============================================================
# 5. 把第 1 条按列切回来确认
# =============================================================
Write-Head '--- 第 1 条的各列内容 ---'
$parts = $firstLine -split "`t", -1
Write-Host ('  {0,3} {1,-18} {2,-10} {3}' -f '列', '項目名', '種別', '内容')
Write-Host ('  ' + ('-' * 66))
for ($c = 0; $c -lt $parts.Count; $c++) {
    $nm = if ($cols[$c].Name) { $cols[$c].Name } else { '-' }
    $v  = if ($parts[$c] -eq '') { '(空)' } else { $parts[$c] }
    Write-Host ('  {0,3} {1,-18} {2,-10} {3}' -f ($c+1), $nm, $cols[$c].Kind, $v)
}
Write-Host ''
Write-Host ("  切出来 {0} 列 " -f $parts.Count) -NoNewline
if ($parts.Count -eq $colCount) { Write-Host '-> OK' -ForegroundColor Green }
else { Write-Host "-> NG（定义是 $colCount 列）" -ForegroundColor Red }

Write-Host ''
Write-Host ' 结果 : OK' -ForegroundColor Green
exit 0
