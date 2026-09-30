<#
    スクリプトの出力 と 手作業の正解 を突き合わせる。「手順」⑦ 結果検証 に対応。

    レポートは「1画面で収まる」長さにしてある：
      ・内容の差分（本当の問題）      → 全件表示。通常は数行
      ・空白の差分（タブ/スペース）  → パターン分類 + 各2件の例のみ
                                        同じ形態の数千行を1行ずつ見る意味はないため
    全明細が必要なときは -Full

    使い方（引数なしなら同じフォルダの *_converted.txt と *_変換後.txt を自動で探す）:
        compare.ps1
        compare.ps1 -Full                 全明細を出力
        compare.ps1 -Mine A.txt -Ref B.txt
#>
param(
    [string] $Mine,                     # スクリプトが作ったもの
    [string] $Ref,                      # 手作業の正解
    [string] $Charset = 'shift_jis',
    [switch] $Full,                     # 空白の差分も全件出力する（かなり長くなる）
    [string] $Report                    # レポートファイルの出力先
)

$ErrorActionPreference = 'Stop'
$enc = if ($Charset -match 'utf') { New-Object System.Text.UTF8Encoding($false) }
       else { [System.Text.Encoding]::GetEncoding(932) }

# 画面とレポートファイルに同じ内容を出すので、画面のスクリーンショット＝レポート
$rep = New-Object 'System.Collections.Generic.List[string]'
function Out-Line {
    param([string]$s = '', [string]$Color = '')
    [void]$rep.Add($s)
    if ($Color) { Write-Host $s -ForegroundColor $Color } else { Write-Host $s }
}

# ---- ファイルの自動検出 ----
function Find-One {
    param([string]$Pattern, [string]$What)
    # -Filter は日本語のファイル名で不安定なことがあるため、全件取得してから -like で絞る。
    # 同名候補が複数ある場合は最新のものを使う
    $f = @(Get-ChildItem -LiteralPath $PSScriptRoot -File |
           Where-Object { $_.Name -like $Pattern } |
           Sort-Object LastWriteTime -Descending)
    if ($f.Count -eq 0) { throw "$What が見つかりません（ファイル名が $Pattern の形）。引数でパスを指定してください。" }
    return $f[0].FullName
}
if (-not $Mine)   { $Mine   = Find-One '*_converted.txt' 'スクリプトの出力' }
if (-not $Ref)    { $Ref    = Find-One '*_変換後.txt'    '手作業の正解' }
if (-not $Report) { $Report = Join-Path $PSScriptRoot 'compare_report.txt' }

function Read-Lines {
    param([string]$p)
    $t = $enc.GetString([System.IO.File]::ReadAllBytes($p))
    $t = $t -replace "`r`n", "`n" -replace "`r", "`n"
    return @($t -split "`n" | Where-Object { $_ -ne '' })
}
$A = Read-Lines $Mine      # スクリプト
$B = Read-Lines $Ref       # 正解

function Vis  { param([string]$s) ($s -replace "`t", '[>]') -replace ' ', '.' }
function Norm { param([string]$s) ($s -replace "[`t ]+", ' ').Trim() }

# ---- 分類 ----
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

# ================= サマリ（ここをスクリーンショットすればよい）=================
Out-Line ''
Out-Line '=============================================================='
Out-Line (' 出力比較（スクリプト vs 手作業の正解）   ' + (Get-Date -Format 'yyyy/MM/dd HH:mm'))
Out-Line '=============================================================='
Out-Line ("  スクリプト : {0}   {1} 行" -f (Split-Path $Mine -Leaf), $A.Count)
Out-Line ("  正解       : {0}   {1} 行" -f (Split-Path $Ref  -Leaf), $B.Count)
Out-Line ''
Out-Line ("  完全に一致               : {0,6} 行" -f $same) $(if($same -gt 0){'Green'}else{''})
Out-Line ("  空白のみ相違             : {0,6} 行" -f $wsLines.Count)    $(if($wsLines.Count){'Yellow'}else{''})
Out-Line ("  内容が余分(スクリプト)   : {0,6} 行 {1}" -f $onlyA_real.Count, $(if($onlyA_real.Count){'★'}else{''})) $(if($onlyA_real.Count){'Red'}else{'Green'})
Out-Line ("  内容が不足(スクリプト)   : {0,6} 行 {1}" -f $onlyB_real.Count, $(if($onlyB_real.Count){'★'}else{''})) $(if($onlyB_real.Count){'Red'}else{'Green'})

# ---- 内容の差分：これが本当の問題。全件出す ----
$pats = @('\log\','\save\','\common\tools\JudgedCIF\data\output','コピー',
          'bk.','_bk','bkup','bak','org.','old','origin.txt')
function Which-Pattern {
    param([string]$line)
    foreach ($p in $pats) {
        if ($line.IndexOf($p, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $p }
    }
    return '(どのパターンにも当たらない)'
}
function Dump-Content {
    param([string]$Title, [object[]]$Lines)
    if ($Lines.Count -eq 0) { return }
    Out-Line ''
    Out-Line ("--- {0}：{1} 行 ---" -f $Title, $Lines.Count) 'Red'
    # 「どのパターンが原因か」でまとめると、理由が一目で分かる
    $g = $Lines | Group-Object { Which-Pattern $_ } | Sort-Object Count -Descending
    foreach ($grp in $g) {
        Out-Line ("  [{0}]  {1} 行" -f $grp.Name, $grp.Count) 'Red'
        $i = 0
        foreach ($l in $grp.Group) {
            $i++
            if ((-not $Full) -and $i -gt 10) {
                Out-Line ("      … 残り {0} 行（-Full で全件表示）" -f ($grp.Count - 10))
                break
            }
            Out-Line ('      ' + ($l -replace "`t", '[>]'))
        }
    }
}
Dump-Content 'スクリプトに無い行（正解にはある＝削りすぎ）' $onlyB_real
Dump-Content 'スクリプトに余分な行（正解には無い＝削り漏れ）' $onlyA_real

# ---- 空白の差分：形態分類のみ。1行ずつは出さない ----
if ($wsLines.Count -gt 0) {
    $refByNorm = @{}
    foreach ($l in $B) {
        $k = Norm $l
        if (-not $refByNorm.ContainsKey($k)) { $refByNorm[$k] = $l }
    }
    # 形態 = 「最初に食い違う文字が何か」。同じ原因の差分は同じ分類にまとまる
    $kinds = @{}
    $examples = @{}
    foreach ($mine in $wsLines) {
        $ref = $refByNorm[(Norm $mine)]
        $n = [Math]::Min($mine.Length, $ref.Length)
        $d = -1
        for ($k = 0; $k -lt $n; $k++) { if ($mine[$k] -ne $ref[$k]) { $d = $k; break } }
        if ($d -lt 0 -and $mine.Length -ne $ref.Length) { $d = $n }
        $mc = if ($d -ge 0 -and $d -lt $mine.Length) { Vis ([string]$mine[$d]) } else { '(行末)' }
        $rc = if ($d -ge 0 -and $d -lt $ref.Length)  { Vis ([string]$ref[$d])  } else { '(行末)' }
        $kind = "スクリプト={0}  正解={1}" -f $mc, $rc
        if ($kinds.ContainsKey($kind)) {
            $kinds[$kind]++
            if ($Full) { $examples[$kind] += ,@($mine, $ref) }
        } else {
            $kinds[$kind] = 1
            $examples[$kind] = @(,@($mine, $ref))
        }
    }

    Out-Line ''
    Out-Line ("--- 空白の差分：{0} 行、形態 {1} 種類 ---" -f $wsLines.Count, $kinds.Count) 'Yellow'
    foreach ($k in ($kinds.Keys | Sort-Object { -$kinds[$_] })) {
        Out-Line ("  [{0}]   {1} 行" -f $k, $kinds[$k]) 'Yellow'
        $ex = $examples[$k]
        $lim = if ($Full) { $ex.Count } else { 2 }
        for ($i = 0; $i -lt [Math]::Min($lim, $ex.Count); $i++) {
            Out-Line ('      スクリプト : ' + (Vis $ex[$i][0]))
            Out-Line ('      正解       : ' + (Vis $ex[$i][1]))
        }
        if ((-not $Full) -and $ex.Count -lt $kinds[$k]) {
            Out-Line ("      （同じ形態が計 {0} 行。-Full で全件表示）" -f $kinds[$k])
        }
    }
    if ($kinds.Count -eq 1) {
        Out-Line '  -> 形態が1種類だけ = 原因は1つ。1行ずつ見る必要はない' 'Green'
    }
}

# ---- 結論 ----
Out-Line ''
if ($onlyA_real.Count -eq 0 -and $onlyB_real.Count -eq 0) {
    if ($wsLines.Count -eq 0) {
        Out-Line ' 結論 : 完全に一致' 'Green'
    } else {
        Out-Line ' 結論 : 内容は一致。空白のみ相違' 'Yellow'
        Out-Line '        行末スペースが原因なら convert.ps1 の -TailSpace で切り替えられる'
    }
} else {
    Out-Line (' 結論 : 内容の差分 {0} 行。1件ずつ確認が必要' -f ($onlyA_real.Count + $onlyB_real.Count)) 'Red'
}
Out-Line '=============================================================='
Out-Line ("  レポート : {0}" -f $Report)
if (-not $Full) { Out-Line '  （-Full を付けると全明細を出力）' }
Out-Line ''

[System.IO.File]::WriteAllBytes($Report,
    (New-Object System.Text.UTF8Encoding($true)).GetBytes(($rep -join "`r`n") + "`r`n"))
