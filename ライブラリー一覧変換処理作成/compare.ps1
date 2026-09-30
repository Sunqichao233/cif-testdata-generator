<#
    把脚本的输出 和 手工做的正解 对比。对应「步骤」⑦ 結果検証。

    报告做成「一屏能截完」的长度：
      ・内容差异（真正的问题）  -> 全部列出，通常只有几行
      ・空白差异（Tab/空格）    -> 只做形态分类 + 各举 2 个例子
                                   形态相同的几千行没必要一条条看
    要全部明细时加 -Full

    用法（不带参数会自动找同一文件夹里的 *_converted.txt 和 *_変換後.txt）:
        compare.ps1
        compare.ps1 -Full                 输出全部明细
        compare.ps1 -Mine A.txt -Ref B.txt
#>
param(
    [string] $Mine,                     # 脚本生成的
    [string] $Ref,                      # 手工做的正解
    [string] $Charset = 'shift_jis',
    [switch] $Full,                     # 连空白差异也全部列出（会很长）
    [string] $Report                    # 报告文件的输出路径
)

$ErrorActionPreference = 'Stop'
$enc = if ($Charset -match 'utf') { New-Object System.Text.UTF8Encoding($false) }
       else { [System.Text.Encoding]::GetEncoding(932) }

# 画面和报告文件输出同样的内容，所以截图 = 报告
$rep = New-Object 'System.Collections.Generic.List[string]'
function Out-Line {
    param([string]$s = '', [string]$Color = '')
    [void]$rep.Add($s)
    if ($Color) { Write-Host $s -ForegroundColor $Color } else { Write-Host $s }
}

# ---- 自动找文件 ----
function Find-One {
    param([string]$Pattern, [string]$What)
    # -Filter 对日文文件名有时不可靠，先全取再用 -like 过滤。同名多个取最新
    $f = @(Get-ChildItem -LiteralPath $PSScriptRoot -File |
           Where-Object { $_.Name -like $Pattern } |
           Sort-Object LastWriteTime -Descending)
    if ($f.Count -eq 0) { throw "没找到$What（文件名要像 $Pattern）。请用参数直接指定路径。" }
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

function Vis  { param([string]$s) ($s -replace "`t", '[>]') -replace ' ', '.' }
function Norm { param([string]$s) ($s -replace "[`t ]+", ' ').Trim() }

# ---- 分类 ----
$setB = New-Object 'System.Collections.Generic.HashSet[string]' (,[string[]]$B)
$setA = New-Object 'System.Collections.Generic.HashSet[string]' (,[string[]]$A)
$onlyA = @($A | Where-Object { -not $setB.Contains($_) })
$onlyB = @($B | Where-Object { -not $setA.Contains($_) })

$normB = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $B) { [void]$normB.Add((Norm $l)) }
$normA = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($l in $A) { [void]$normA.Add((Norm $l)) }

$onlyA_real = @($onlyA | Where-Object { -not $normB.Contains((Norm $_)) })
$onlyB_real = @($onlyB | Where-Object { -not $normA.Contains((Norm $_)) })
$wsLines    = @($onlyA | Where-Object { $normB.Contains((Norm $_)) })
$same       = $A.Count - $onlyA.Count

# ================= 摘要（这一块就是要截图的部分）=================
Out-Line ''
Out-Line '=============================================================='
Out-Line (' 输出对比（脚本 vs 手工正解）        ' + (Get-Date -Format 'yyyy/MM/dd HH:mm'))
Out-Line '=============================================================='
Out-Line ("  脚本 : {0}   {1} 行" -f (Split-Path $Mine -Leaf), $A.Count)
Out-Line ("  正解 : {0}   {1} 行" -f (Split-Path $Ref  -Leaf), $B.Count)
Out-Line ''
Out-Line ("  両方一致             : {0,6} 行" -f $same) $(if($same -gt 0){'Green'}else{''})
Out-Line ("  只是空白不同         : {0,6} 行" -f $wsLines.Count)    $(if($wsLines.Count){'Yellow'}else{''})
Out-Line ("  内容真的多出来(脚本) : {0,6} 行 {1}" -f $onlyA_real.Count, $(if($onlyA_real.Count){'★'}else{''})) $(if($onlyA_real.Count){'Red'}else{'Green'})
Out-Line ("  内容真的缺少(脚本)   : {0,6} 行 {1}" -f $onlyB_real.Count, $(if($onlyB_real.Count){'★'}else{''})) $(if($onlyB_real.Count){'Red'}else{'Green'})

# ---- 内容差异：这是真正的问题，全部列出 ----
$pats = @('\log\','\save\','\common\tools\JudgedCIF\data\output','コピー',
          'bk.','_bk','bkup','bak','org.','old','origin.txt')
function Which-Pattern {
    param([string]$line)
    foreach ($p in $pats) {
        if ($line.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $p }
    }
    return '(哪个模式都没命中)'
}
function Dump-Content {
    param([string]$Title, [object[]]$Lines)
    if ($Lines.Count -eq 0) { return }
    Out-Line ''
    Out-Line ("--- {0}：{1} 行 ---" -f $Title, $Lines.Count) 'Red'
    # 先按「是哪个模式造成的」归类，一眼看出原因
    $g = $Lines | Group-Object { Which-Pattern $_ } | Sort-Object Count -Descending
    foreach ($grp in $g) {
        Out-Line ("  [{0}]  {1} 行" -f $grp.Name, $grp.Count) 'Red'
        $i = 0
        foreach ($l in $grp.Group) {
            $i++
            if ((-not $Full) -and $i -gt 10) {
                Out-Line ("      … 其余 {0} 行（-Full 可全部列出）" -f ($grp.Count - 10))
                break
            }
            Out-Line ('      ' + ($l -replace "`t", '[>]'))
        }
    }
}
Dump-Content '脚本缺少的行（正解有、脚本删掉了）' $onlyB_real
Dump-Content '脚本多出来的行（正解没有、脚本留着）' $onlyA_real

# ---- 空白差异：只做形态分类，不逐行列 ----
if ($wsLines.Count -gt 0) {
    $refByNorm = @{}
    foreach ($l in $B) {
        $k = Norm $l
        if (-not $refByNorm.ContainsKey($k)) { $refByNorm[$k] = $l }
    }
    # 形态 = 「第一个不同的字符是什么」。同一原因造成的差异会归成同一类
    $kinds = @{}
    $examples = @{}
    foreach ($mine in $wsLines) {
        $ref = $refByNorm[(Norm $mine)]
        $n = [Math]::Min($mine.Length, $ref.Length)
        $d = -1
        for ($k = 0; $k -lt $n; $k++) { if ($mine[$k] -ne $ref[$k]) { $d = $k; break } }
        if ($d -lt 0 -and $mine.Length -ne $ref.Length) { $d = $n }
        $mc = if ($d -ge 0 -and $d -lt $mine.Length) { Vis ([string]$mine[$d]) } else { '(行尾)' }
        $rc = if ($d -ge 0 -and $d -lt $ref.Length)  { Vis ([string]$ref[$d])  } else { '(行尾)' }
        $kind = "脚本={0}  正解={1}" -f $mc, $rc
        if ($kinds.ContainsKey($kind)) {
            $kinds[$kind]++
            if ($Full) { $examples[$kind] += ,@($mine, $ref) }
        } else {
            $kinds[$kind] = 1
            $examples[$kind] = @(,@($mine, $ref))
        }
    }

    Out-Line ''
    Out-Line ("--- 空白差异：{0} 行，形态 {1} 种 ---" -f $wsLines.Count, $kinds.Count) 'Yellow'
    foreach ($k in ($kinds.Keys | Sort-Object { -$kinds[$_] })) {
        Out-Line ("  [{0}]   {1} 行" -f $k, $kinds[$k]) 'Yellow'
        $ex = $examples[$k]
        $lim = if ($Full) { $ex.Count } else { 2 }
        for ($i = 0; $i -lt [Math]::Min($lim, $ex.Count); $i++) {
            Out-Line ('      脚本 : ' + (Vis $ex[$i][0]))
            Out-Line ('      正解 : ' + (Vis $ex[$i][1]))
        }
        if ((-not $Full) -and $ex.Count -lt $kinds[$k]) {
            Out-Line ("      （同形态共 {0} 行，-Full 可全部列出）" -f $kinds[$k])
        }
    }
    if ($kinds.Count -eq 1) {
        Out-Line '  -> 形态只有 1 种 = 同一个原因，不用逐行看' 'Green'
    }
}

# ---- 结论 ----
Out-Line ''
if ($onlyA_real.Count -eq 0 -and $onlyB_real.Count -eq 0) {
    if ($wsLines.Count -eq 0) {
        Out-Line ' 结论 : 完全一致' 'Green'
    } else {
        Out-Line ' 结论 : 内容一致，只有空白差异' 'Yellow'
        Out-Line '        行末空格的话 convert.ps1 的 -TailSpace 可切换'
    }
} else {
    Out-Line (' 结论 : 内容差异 {0} 行，要逐条确认' -f ($onlyA_real.Count + $onlyB_real.Count)) 'Red'
}
Out-Line '=============================================================='
Out-Line ("  报告 : {0}" -f $Report)
if (-not $Full) { Out-Line '  （加 -Full 可输出全部明细）' }
Out-Line ''

[System.IO.File]::WriteAllBytes($Report,
    (New-Object System.Text.UTF8Encoding($true)).GetBytes(($rep -join "`r`n") + "`r`n"))
