<#
    把脚本的输出 和 手工做的正解 对比，找出差异在哪。
    对应「步骤」文件的 ⑦ 結果検証。

    画面上只显示摘要和少量例子；
    ★全部差异会写进报告文件 compare_report.txt，用编辑器打开慢慢看。

    用法（不带参数会自动找同一文件夹里的 *_converted.txt 和 *_変換後.txt）:
        compare.ps1
        compare.ps1 -Mine 脚本输出.txt -Ref 手工正解.txt
        compare.ps1 -Samples 30           画面上多显示几条
#>
param(
    [string] $Mine,                     # 脚本生成的
    [string] $Ref,                      # 手工做的正解
    [string] $Charset = 'shift_jis',
    [int]    $Samples = 8,              # 画面上每类显示几条（报告文件里始终是全部）
    [string] $Report                    # 报告文件的输出路径
)

$ErrorActionPreference = 'Stop'
$enc = if ($Charset -match 'utf') { New-Object System.Text.UTF8Encoding($false) }
       else { [System.Text.Encoding]::GetEncoding(932) }

# ---- 自动找文件 ----
function Find-One {
    param([string]$Pattern, [string]$What)
    # -Filter 对日文文件名有时不可靠，所以先全取再用 -like 过滤
    # 同名多个时取最新的一个
    $f = @(Get-ChildItem -LiteralPath $PSScriptRoot -File |
           Where-Object { $_.Name -like $Pattern } |
           Sort-Object LastWriteTime -Descending)
    if ($f.Count -eq 0) {
        throw "没找到$What（文件名要像 $Pattern）。请用参数直接指定路径。"
    }
    if ($f.Count -gt 1) {
        Write-Host ("  [注意] 有 {0} 个候选，取最新的: {1}" -f $f.Count, $f[0].Name) -ForegroundColor Yellow
    }
    return $f[0].FullName
}
if (-not $Mine)   { $Mine   = Find-One '*_converted.txt' '脚本的输出' }
if (-not $Ref)    { $Ref    = Find-One '*_変換後.txt'    '手工的正解' }
if (-not $Report) { $Report = Join-Path $PSScriptRoot 'compare_report.txt' }

function Read-Lines {
    param([string]$p)
    $t = $enc.GetString([System.IO.File]::ReadAllBytes($p))
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return @($t -split "`n" | Where-Object { $_ -ne '' })
}

$A = Read-Lines $Mine      # 脚本
$B = Read-Lines $Ref       # 正解

# 报告文件的内容。画面上只印摘要，这里存全部
$rep = New-Object 'System.Collections.Generic.List[string]'
function Rep { param([string]$s = '') [void]$rep.Add($s) }

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' 输出对比（脚本 vs 手工正解）' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  脚本 : $Mine"
Write-Host ("         {0} 行" -f $A.Count)
Write-Host "  正解 : $Ref"
Write-Host ("         {0} 行" -f $B.Count)
Write-Host ("  差   : {0} 行" -f ($A.Count - $B.Count))

Rep '=================================================================='
Rep ' 输出对比报告（脚本 vs 手工正解）'
Rep ('  作成: ' + (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'))
Rep '=================================================================='
Rep ("  脚本 : {0}   {1} 行" -f $Mine, $A.Count)
Rep ("  正解 : {0}   {1} 行" -f $Ref,  $B.Count)
Rep ("  差   : {0} 行" -f ($A.Count - $B.Count))
Rep

# ---- 1) 逐字节比较 ----
$setB = New-Object 'System.Collections.Generic.HashSet[string]' (,[string[]]$B)
$setA = New-Object 'System.Collections.Generic.HashSet[string]' (,[string[]]$A)
$onlyA = @($A | Where-Object { -not $setB.Contains($_) })
$onlyB = @($B | Where-Object { -not $setA.Contains($_) })

Write-Host ''
Write-Host '--- 1) 逐字节比较 ---' -ForegroundColor Cyan
Write-Host ("  两边都有       : {0} 行" -f ($A.Count - $onlyA.Count))
Write-Host ("  只在脚本侧有   : {0} 行" -f $onlyA.Count)
Write-Host ("  只在正解侧有   : {0} 行" -f $onlyB.Count)

# ---- 2) 忽略空白后再比较 ----
function Norm { param([string]$s) ($s -replace "[`t ]+", ' ').Trim() }
$normB = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $B) { [void]$normB.Add((Norm $l)) }
$normA = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $A) { [void]$normA.Add((Norm $l)) }

$onlyA_real = @($onlyA | Where-Object { -not $normB.Contains((Norm $_)) })
$onlyB_real = @($onlyB | Where-Object { -not $normA.Contains((Norm $_)) })
$wsLines    = @($onlyA | Where-Object { $normB.Contains((Norm $_)) })

Write-Host ''
Write-Host '--- 2) 忽略空白差异后 ---' -ForegroundColor Cyan
Write-Host ("  只是空白不同         : {0} 行" -f $wsLines.Count) -ForegroundColor Yellow
Write-Host ("  内容真的多出来(脚本) : {0} 行" -f $onlyA_real.Count) -ForegroundColor $(if($onlyA_real.Count){'Red'}else{'Green'})
Write-Host ("  内容真的缺少(脚本)   : {0} 行" -f $onlyB_real.Count) -ForegroundColor $(if($onlyB_real.Count){'Red'}else{'Green'})

Rep '--- 差异的内訳 ---'
Rep ("  两边完全一致         : {0} 行" -f ($A.Count - $onlyA.Count))
Rep ("  只是空白不同         : {0} 行" -f $wsLines.Count)
Rep ("  内容真的多出来(脚本) : {0} 行" -f $onlyA_real.Count)
Rep ("  内容真的缺少(脚本)   : {0} 行" -f $onlyB_real.Count)
Rep

# ---- 3) 多出来的行是被哪个模式命中的 ----
$pats = @('\log\','\save\','\common\tools\JudgedCIF\data\output','コピー',
          'bk.','_bk','bkup','bak','org.','old','origin.txt')
function Show-ByPattern {
    param([object[]]$Lines, [string]$Title)
    if ($Lines.Count -eq 0) { return }
    Write-Host ''
    Write-Host ("--- {0} ---" -f $Title) -ForegroundColor Cyan
    Rep ("--- {0} ---" -f $Title)
    foreach ($p in $pats) {
        $n = @($Lines | Where-Object { $_.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count
        if ($n -gt 0) { Write-Host ("    {0,-38} {1} 行" -f $p, $n); Rep ("    {0,-38} {1} 行" -f $p, $n) }
    }
    $none = @($Lines | Where-Object {
        $hit = $false
        foreach ($p in $pats) { if ($_.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break } }
        -not $hit })
    if ($none.Count -gt 0) {
        Write-Host ("    哪个模式都没命中                     {0} 行  ★要调查" -f $none.Count) -ForegroundColor Red
        Rep ("    哪个模式都没命中                     {0} 行  ★要调查" -f $none.Count)
    }
    Rep
}
Show-ByPattern $onlyA_real '脚本多出来的行，是被哪个关键词命中的'
Show-ByPattern $onlyB_real '脚本缺少的行，是被哪个关键词删掉的'

# ---- 4) 全部差异写进报告；画面只显示前 N 条 ----
function Dump {
    param([string]$Title, [object[]]$Lines, [string]$Color)
    Rep ('================================================================')
    Rep (' {0}   {1} 行' -f $Title, $Lines.Count)
    Rep ('================================================================')
    if ($Lines.Count -eq 0) { Rep '  (なし)'; Rep; return }
    $i = 0
    foreach ($l in $Lines) { $i++; Rep ('{0,6}: {1}' -f $i, ($l -replace "`t", '[>]')) }
    Rep

    Write-Host ''
    Write-Host ("--- {0}（画面只显示前 {1} 条 / 全部在报告文件里）---" -f `
                $Title, [Math]::Min($Samples, $Lines.Count)) -ForegroundColor $Color
    $Lines | Select-Object -First $Samples | ForEach-Object {
        Write-Host ('    ' + ($_ -replace "`t", '[>]'))
    }
}
Dump '脚本缺少的行（内容差异）' $onlyB_real 'Red'
Dump '脚本多出来的行（内容差异）' $onlyA_real 'Red'

# ---- 5) 空白差异：两边并排，空白显形，全部写进报告 ----
if ($wsLines.Count -gt 0) {
    $refByNorm = @{}
    foreach ($l in $B) {
        $k = Norm $l
        if (-not $refByNorm.ContainsKey($k)) { $refByNorm[$k] = $l }
    }
    function Vis { param([string]$s) ($s -replace "`t", '[>]') -replace ' ', '.' }

    Rep '================================================================'
    Rep (' 只有空白不同的行   {0} 行    （Tab=[>]  半角空格=. ）' -f $wsLines.Count)
    Rep '================================================================'

    # 差异形态的统计。1987 行如果都是同一个原因，看这里就够了
    $kinds = @{}
    $i = 0
    foreach ($mine in $wsLines) {
        $i++
        $ref = $refByNorm[(Norm $mine)]
        $n = [Math]::Min($mine.Length, $ref.Length)
        $d = -1
        for ($k = 0; $k -lt $n; $k++) { if ($mine[$k] -ne $ref[$k]) { $d = $k; break } }
        if ($d -lt 0 -and $mine.Length -ne $ref.Length) { $d = $n }

        $mc = if ($d -ge 0 -and $d -lt $mine.Length) { Vis ([string]$mine[$d]) } else { '(行尾)' }
        $rc = if ($d -ge 0 -and $d -lt $ref.Length)  { Vis ([string]$ref[$d])  } else { '(行尾)' }
        $kind = "脚本={0}  正解={1}" -f $mc, $rc
        if ($kinds.ContainsKey($kind)) { $kinds[$kind]++ } else { $kinds[$kind] = 1 }

        Rep ('{0,6}: 脚本 : {1}' -f $i, (Vis $mine))
        Rep ('        正解 : {0}' -f (Vis $ref))
        Rep ('               第 {0} 个字符开始不同：{1}' -f ($d + 1), $kind)
    }
    Rep

    Write-Host ''
    Write-Host '--- 空白差异的形态（全部行的分类统计）---' -ForegroundColor Yellow
    Rep '--- 空白差异的形态（分类统计）---'
    foreach ($k in ($kinds.Keys | Sort-Object { -$kinds[$_] })) {
        Write-Host ("    {0,-34} {1} 行" -f $k, $kinds[$k])
        Rep ("    {0,-34} {1} 行" -f $k, $kinds[$k])
    }
    if ($kinds.Count -eq 1) {
        Write-Host '    -> 形态只有 1 种，说明是同一个原因造成的' -ForegroundColor Green
        Rep '    -> 形态只有 1 种，说明是同一个原因造成的'
    }
    Rep
}

# ---- 报告文件输出 ----
[System.IO.File]::WriteAllBytes($Report,
    (New-Object System.Text.UTF8Encoding($true)).GetBytes(($rep -join "`r`n") + "`r`n"))

# ---- 结论 ----
Write-Host ''
if ($onlyA_real.Count -eq 0 -and $onlyB_real.Count -eq 0) {
    if ($wsLines.Count -eq 0) {
        Write-Host ' 结论 : 完全一致' -ForegroundColor Green
    } else {
        Write-Host ' 结论 : 内容一致，只有空白（Tab/空格）的差异' -ForegroundColor Yellow
        Write-Host '        -> 看上面的形态统计。行末空格的话，convert.ps1 的 -TailSpace 可以切换。'
    }
} else {
    Write-Host ' 结论 : 有内容差异，要逐条确认' -ForegroundColor Red
    Write-Host '        多出来的行 -> 脚本没删、正解删了'
    Write-Host '        缺少的行   -> ★脚本删多了，看上面是哪个模式干的'
}
Write-Host ''
Write-Host ("  ★全部差异写在这里（{0} 行）: {1}" -f $rep.Count, $Report) -ForegroundColor Cyan
Write-Host ''
