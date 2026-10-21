<#
    テレワーク連絡メールを1クリックで作る。

      ・件名の日付は起動時に今日の日付が入る（画面で変更可）
      ・作業時間は画面で入力できる（既定は config.json の値）
      ・「送信」で Outlook から送信、「Outlook で開く」で下書きを開く

    Outlook（デスクトップ版）の COM を使うため、
    クラシック Outlook が入っている端末で動く。
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Root = Split-Path -Parent $PSScriptRoot
$CfgPath = Join-Path $Root 'config.json'

# ---- 設定の読み込み ----
if (Test-Path -LiteralPath $CfgPath) {
    $cfg = Get-Content -LiteralPath $CfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
    $cfg = [pscustomobject]@{
        To        = '[DL] ASD ASIDC; [DL] ASD ASPDC'
        Subject   = '【テレワーク連絡】{date}：IKI 孫キ超'
        Greeting  = 'お疲れ様です。IKI の孫です。'
        Message   = '作業を開始します。'
        Project   = 'L&F'
        StartTime = '09:00'
        EndTime   = '18:00'
    }
}

function Save-Config {
    param($c)
    try {
        $c | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $CfgPath -Encoding UTF8
    } catch { }
}

function HtmlEsc {
    param([string]$s)
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Build-Mail {
    param($dateText, $st, $et, $proj)
    $subject = $cfg.Subject.Replace('{date}', $dateText)
    $body = @"
<div style="font-family:'游ゴシック',Meiryo,sans-serif;font-size:10.5pt">
各位<br><br>
$(HtmlEsc $cfg.Greeting)<br>
$(HtmlEsc $cfg.Message)<br><br><br>
作業時間:$st-$et<br><br>
・$(HtmlEsc $proj)<br>
</div>
"@
    return @{ Subject = $subject; Html = $body }
}

function Show-Plain {
    param($m, $st, $et, $proj)
    return "件名：$($m.Subject)`r`n`r`n各位`r`n`r`n$($cfg.Greeting)`r`n$($cfg.Message)`r`n`r`n`r`n作業時間:$st-$et`r`n`r`n・$proj"
}

function Send-Or-Show {
    param($m, [bool]$doSend)
    try {
        $ol = New-Object -ComObject Outlook.Application
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Outlook（デスクトップ版）に接続できませんでした。`r`n" +
            "クラシック Outlook を起動してから、もう一度お試しください。`r`n`r`n" +
            "本文はクリップボードにコピーしました。", 'Outlook に接続できません',
            'OK', 'Warning') | Out-Null
        return $false
    }
    $mail = $ol.CreateItem(0)          # olMailItem
    $mail.To = $cfg.To
    $mail.Subject = $m.Subject
    $mail.BodyFormat = 2               # olFormatHTML
    # Inspector を取ると署名が入る。その前に本文を差し込む
    $null = $mail.GetInspector
    $mail.HTMLBody = $m.Html + $mail.HTMLBody
    try { $mail.Recipients.ResolveAll() | Out-Null } catch { }

    if ($doSend) {
        $mail.Send()
    } else {
        $mail.Display()
    }
    return $true
}

# ---- 画面 ----
$f = New-Object System.Windows.Forms.Form
$f.Text = 'テレワーク連絡メール'
$f.Size = New-Object System.Drawing.Size(460, 330)
$f.StartPosition = 'CenterScreen'
$f.FormBorderStyle = 'FixedDialog'
$f.MaximizeBox = $false
$f.Font = New-Object System.Drawing.Font('Meiryo UI', 9)

function NewLabel {
    param($text, $x, $y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.AutoSize = $true
    $f.Controls.Add($l); return $l
}

NewLabel '日付' 20 24 | Out-Null
$dt = New-Object System.Windows.Forms.DateTimePicker
$dt.Location = New-Object System.Drawing.Point(110, 20)
$dt.Size = New-Object System.Drawing.Size(150, 24)
$dt.Format = 'Custom'
$dt.CustomFormat = 'yyyy/MM/dd'
$f.Controls.Add($dt)

NewLabel '作業時間' 20 64 | Out-Null
$tbS = New-Object System.Windows.Forms.TextBox
$tbS.Location = New-Object System.Drawing.Point(110, 60)
$tbS.Size = New-Object System.Drawing.Size(60, 24)
$tbS.Text = $cfg.StartTime
$f.Controls.Add($tbS)
NewLabel '〜' 178 64 | Out-Null
$tbE = New-Object System.Windows.Forms.TextBox
$tbE.Location = New-Object System.Drawing.Point(200, 60)
$tbE.Size = New-Object System.Drawing.Size(60, 24)
$tbE.Text = $cfg.EndTime
$f.Controls.Add($tbE)

NewLabel '案件' 20 104 | Out-Null
$tbP = New-Object System.Windows.Forms.TextBox
$tbP.Location = New-Object System.Drawing.Point(110, 100)
$tbP.Size = New-Object System.Drawing.Size(150, 24)
$tbP.Text = $cfg.Project
$f.Controls.Add($tbP)

$lblTo = NewLabel ("宛先：" + $cfg.To) 20 144
$lblTo.MaximumSize = New-Object System.Drawing.Size(400, 0)
$lblTo.ForeColor = [System.Drawing.Color]::DimGray

$lblPrev = NewLabel '' 20 174
$lblPrev.MaximumSize = New-Object System.Drawing.Size(400, 0)
$lblPrev.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 140)

function Refresh-Preview {
    $lblPrev.Text = "件名：" + $cfg.Subject.Replace('{date}', $dt.Value.ToString('yyyy/MM/dd')) +
                    "`r`n作業時間:" + $tbS.Text + "-" + $tbE.Text
}
$dt.Add_ValueChanged({ Refresh-Preview })
$tbS.Add_TextChanged({ Refresh-Preview })
$tbE.Add_TextChanged({ Refresh-Preview })
Refresh-Preview

$btnSend = New-Object System.Windows.Forms.Button
$btnSend.Text = '送信'
$btnSend.Location = New-Object System.Drawing.Point(110, 232)
$btnSend.Size = New-Object System.Drawing.Size(110, 34)
$f.Controls.Add($btnSend)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'Outlook で開く'
$btnOpen.Location = New-Object System.Drawing.Point(234, 232)
$btnOpen.Size = New-Object System.Drawing.Size(130, 34)
$f.Controls.Add($btnOpen)

function Current {
    return Build-Mail $dt.Value.ToString('yyyy/MM/dd') $tbS.Text $tbE.Text $tbP.Text
}

function Remember {
    $cfg.StartTime = $tbS.Text
    $cfg.EndTime = $tbE.Text
    $cfg.Project = $tbP.Text
    Save-Config $cfg
}

$btnSend.Add_Click({
    $m = Current
    $plain = Show-Plain $m $tbS.Text $tbE.Text $tbP.Text
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "$plain`r`n`r`n宛先：$($cfg.To)`r`n`r`nこの内容で送信しますか？",
        '送信の確認', 'OKCancel', 'Question')
    if ($ans -ne 'OK') { return }
    try { Set-Clipboard -Value $plain } catch { }
    if (Send-Or-Show $m $true) {
        Remember
        [System.Windows.Forms.MessageBox]::Show('送信しました。', '完了', 'OK', 'Information') | Out-Null
        $f.Close()
    }
})

$btnOpen.Add_Click({
    $m = Current
    try { Set-Clipboard -Value (Show-Plain $m $tbS.Text $tbE.Text $tbP.Text) } catch { }
    if (Send-Or-Show $m $false) {
        Remember
        $f.Close()
    }
})

$f.Topmost = $true
[void]$f.ShowDialog()
