<#
    把脚本的输出 和 手工做的正解 对比，找出差异在哪。
    对应「步骤」文件的 ⑦ 結果検証。

    用法（不带参数会自动找同一文件夹里的 *_converted.txt 和 *_変換後.txt）:
        compare.ps1
        compare.ps1 -Mine 脚本输出.txt -Ref 手工正解.txt
#>
param(
    [string] $Mine,                     # 脚本生成的
    [string] $Ref,                      # 手工做的正解
    [string] $Charset = 'shift_jis',
    [int]    $Samples = 8               # 每类差异显示几条例子
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
if (-not $Mine) { $Mine = Find-One '*_converted.txt' '脚本的输出' }
if (-not $Ref)  { $Ref  = Find-One '*_変換後.txt'    '手工的正解' }

function Read-Lines {
    param([string]$p)
    $t = $enc.GetString([System.IO.File]::ReadAllBytes($p))
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return @($t -split "`n" | Where-Object { $_ -ne '' })
}

$A = Read-Lines $Mine      # 脚本
$B = Read-Lines $Ref       # 正解

Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' 输出对比（脚本 vs 手工正解）' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Host "  脚本 : $Mine"
Write-Host ("         {0} 行" -f $A.Count)
Write-Host "  正解 : $Ref"
Write-Host ("         {0} 行" -f $B.Count)
Write-Host ("  差   : {0} 行" -f ($A.Count - $B.Count))

# ---- 1) 完全一致的比较 ----
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
# 目的：判断差异是「内容不同」还是「只差空白（Tab/空格）」
function Norm { param([string]$s) ($s -replace "[`t ]+", ' ').Trim() }
$normB = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $B) { [void]$normB.Add((Norm $l)) }
$normA = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $A) { [void]$normA.Add((Norm $l)) }

$onlyA_real = @($onlyA | Where-Object { -not $normB.Contains((Norm $_)) })
$onlyB_real = @($onlyB | Where-Object { -not $normA.Contains((Norm $_)) })
$wsOnlyA = $onlyA.Count - $onlyA_real.Count
$wsOnlyB = $onlyB.Count - $onlyB_real.Count

Write-Host ''
Write-Host '--- 2) 忽略空白差异后 ---' -ForegroundColor Cyan
Write-Host ("  只是空白不同(脚本侧) : {0} 行" -f $wsOnlyA) -ForegroundColor Yellow
Write-Host ("  只是空白不同(正解侧) : {0} 行" -f $wsOnlyB) -ForegroundColor Yellow
Write-Host ("  内容真的多出来(脚本) : {0} 行" -f $onlyA_real.Count) -ForegroundColor $(if($onlyA_real.Count){'Red'}else{'Green'})
Write-Host ("  内容真的缺少(脚本)   : {0} 行" -f $onlyB_real.Count) -ForegroundColor $(if($onlyB_real.Count){'Red'}else{'Green'})

# ---- 3) 多出来的行是被哪个模式命中的 ----
if ($onlyA_real.Count -gt 0) {
    Write-Host ''
    Write-Host '--- 3) 脚本多出来的行，是被哪个关键词命中的 ---' -ForegroundColor Cyan
    Write-Host '    （正解把它们删了，脚本没删 -> 多半是 old 安全模式的影响）'
    $pats = @('\log\','\save\','\common\tools\JudgedCIF\data\output','コピー',
              'bk.','_bk','bkup','bak','org.','old','origin.txt')
    foreach ($p in $pats) {
        $n = @($onlyA_real | Where-Object {
                $_.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count
        if ($n -gt 0) { Write-Host ("    {0,-38} {1} 行" -f $p, $n) }
    }
    $none = @($onlyA_real | Where-Object {
        $hit = $false
        foreach ($p in $pats) {
            if ($_.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break }
        }
        -not $hit })
    if ($none.Count -gt 0) {
        Write-Host ("    どの模式にも当たらない              {0} 行  ★要调查" -f $none.Count) -ForegroundColor Red
    }
}

# ---- 4) 例子 ----
function Show-Sample {
    param([string]$Title, [object[]]$Lines, [string]$Color)
    if ($Lines.Count -eq 0) { return }
    Write-Host ''
    Write-Host ("--- {0}（前 {1} 条）---" -f $Title, [Math]::Min($Samples, $Lines.Count)) -ForegroundColor $Color
    $Lines | Select-Object -First $Samples | ForEach-Object {
        Write-Host ('    ' + ($_ -replace "`t", '→'))
    }
}
Show-Sample '脚本多出来的行（内容差异）' $onlyA_real 'Red'
Show-Sample '脚本缺少的行（内容差异）'   $onlyB_real 'Red'
Show-Sample '只有空白不同的行（脚本侧）' @($onlyA | Where-Object { $normB.Contains((Norm $_)) }) 'Yellow'

# ---- 结论 ----
Write-Host ''
if ($onlyA_real.Count -eq 0 -and $onlyB_real.Count -eq 0) {
    if ($wsOnlyA -eq 0 -and $wsOnlyB -eq 0) {
        Write-Host ' 结论 : 完全一致' -ForegroundColor Green
    } else {
        Write-Host ' 结论 : 内容一致，只有空白（Tab/空格）的差异' -ForegroundColor Yellow
        Write-Host '        -> 看上面的例子，判断是不是 (8) 行末空白的处理方式不同。'
        Write-Host '           要贴合手工结果就把 convert.ps1 的 TrimEnd 调整一下。'
    }
} else {
    Write-Host ' 结论 : 有内容差异，要逐条确认' -ForegroundColor Red
    Write-Host '        多出来的行 -> 多半是 old 安全模式（正解删了、脚本留着）'
    Write-Host '        缺少的行   -> ★脚本删多了，要查是哪个模式误伤的'
}
Write-Host ''
