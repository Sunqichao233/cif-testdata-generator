<#
    画面スリープ・離席状態になるのを防ぐ。

      ・SetThreadExecutionState で「画面を消さない」ようにOSへ要求する
      ・あわせて一定間隔で F15 キーを送り、無操作タイマーをリセットする
        （F15 はどのアプリでも何も起きないキー。元の KeepAlive.vbs と同じ考え方）

    画面上で「動いているか」「次の送信まで何秒か」が見え、
    ボタンで停止・再開できる。
#>
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    Add-Type -Namespace Native -Name Api -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);

[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, System.UIntPtr dwExtraInfo);
'@

    $ES_CONTINUOUS       = [uint32]'0x80000000'
    $ES_DISPLAY_REQUIRED = [uint32]'0x00000002'
    $ES_SYSTEM_REQUIRED  = [uint32]'0x00000001'
    $VK_F15              = [byte]0x7E
    $KEYEVENTF_KEYUP     = [uint32]0x0002

    function Set-Awake {
        param([bool]$on)
        if ($on) {
            [void][Native.Api]::SetThreadExecutionState(
                $ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED -bor $ES_SYSTEM_REQUIRED)
        } else {
            [void][Native.Api]::SetThreadExecutionState($ES_CONTINUOUS)
        }
    }

    function Send-Dummy {
        [Native.Api]::keybd_event($VK_F15, 0, 0, [System.UIntPtr]::Zero)
        [Native.Api]::keybd_event($VK_F15, 0, $KEYEVENTF_KEYUP, [System.UIntPtr]::Zero)
    }

    # ---- 状態 ----
    $script:running  = $false
    $script:left     = 0
    $script:count    = 0
    $script:started  = $null
    $script:lastSent = $null
    $script:interval = 240

    # ---- 画面 ----
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '画面スリープ防止'
    $f.Size = New-Object System.Drawing.Size(330, 300)
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedSingle'
    $f.MaximizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Meiryo UI', 9)

    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Location = New-Object System.Drawing.Point(20, 16)
    $lblState.Size = New-Object System.Drawing.Size(275, 36)
    $lblState.TextAlign = 'MiddleCenter'
    $lblState.Font = New-Object System.Drawing.Font('Meiryo UI', 14, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($lblState)

    function AddRow {
        param([string]$cap, [int]$y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $cap
        $l.Location = New-Object System.Drawing.Point(26, $y)
        $l.Size = New-Object System.Drawing.Size(120, 20)
        $f.Controls.Add($l)
        $v = New-Object System.Windows.Forms.Label
        $v.Location = New-Object System.Drawing.Point(150, $y)
        $v.Size = New-Object System.Drawing.Size(140, 20)
        $v.Font = New-Object System.Drawing.Font('Consolas', 10)
        $f.Controls.Add($v)
        return $v
    }

    $vNext  = AddRow '次の送信まで' 64
    $vElaps = AddRow '経過時間'     88
    $vCount = AddRow '送信回数'     112
    $vLast  = AddRow '最終送信'     136

    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text = '間隔'
    $lblInt.Location = New-Object System.Drawing.Point(26, 166)
    $lblInt.Size = New-Object System.Drawing.Size(40, 20)
    $f.Controls.Add($lblInt)

    $numInt = New-Object System.Windows.Forms.NumericUpDown
    $numInt.Location = New-Object System.Drawing.Point(90, 163)
    $numInt.Size = New-Object System.Drawing.Size(70, 22)
    $numInt.Minimum = 30
    $numInt.Maximum = 3600
    $numInt.Increment = 30
    $numInt.Value = 240
    $f.Controls.Add($numInt)

    $lblSec = New-Object System.Windows.Forms.Label
    $lblSec.Text = '秒'
    $lblSec.Location = New-Object System.Drawing.Point(166, 166)
    $lblSec.Size = New-Object System.Drawing.Size(30, 20)
    $f.Controls.Add($lblSec)

    $chkTop = New-Object System.Windows.Forms.CheckBox
    $chkTop.Text = '手前に表示'
    $chkTop.Location = New-Object System.Drawing.Point(200, 164)
    $chkTop.Size = New-Object System.Drawing.Size(100, 22)
    $chkTop.Checked = $true
    $f.Controls.Add($chkTop)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Location = New-Object System.Drawing.Point(26, 200)
    $btnRun.Size = New-Object System.Drawing.Size(130, 38)
    $f.Controls.Add($btnRun)

    $btnQuit = New-Object System.Windows.Forms.Button
    $btnQuit.Text = '終了'
    $btnQuit.Location = New-Object System.Drawing.Point(170, 200)
    $btnQuit.Size = New-Object System.Drawing.Size(120, 38)
    $f.Controls.Add($btnQuit)

    function Fmt { param([int]$s); return '{0:00}:{1:00}' -f [int]($s / 60), ($s % 60) }

    function Paint {
        if ($script:running) {
            $lblState.Text = '● 動作中'
            $lblState.ForeColor = [System.Drawing.Color]::FromArgb(0, 140, 60)
            $btnRun.Text = '停止'
            $vNext.Text = Fmt $script:left
        } else {
            $lblState.Text = '■ 停止中'
            $lblState.ForeColor = [System.Drawing.Color]::FromArgb(170, 50, 50)
            $btnRun.Text = '開始'
            $vNext.Text = '-'
        }
        if ($script:started) {
            $e = [int]((Get-Date) - $script:started).TotalSeconds
            $vElaps.Text = '{0:00}:{1:00}:{2:00}' -f [int]($e / 3600), [int](($e % 3600) / 60), ($e % 60)
        } else {
            $vElaps.Text = '-'
        }
        $vCount.Text = "$($script:count) 回"
        $vLast.Text = if ($script:lastSent) { $script:lastSent.ToString('HH:mm:ss') } else { '-' }
    }

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        if (-not $script:running) { return }
        $script:left--
        if ($script:left -le 0) {
            Send-Dummy
            Set-Awake $true          # 念のため毎回要求し直す
            $script:count++
            $script:lastSent = Get-Date
            $script:left = $script:interval
        }
        Paint
    })

    function Start-KeepAlive {
        $script:interval = [int]$numInt.Value
        $script:left = $script:interval
        $script:running = $true
        if (-not $script:started) { $script:started = Get-Date }
        $numInt.Enabled = $false
        Set-Awake $true
        $timer.Start()
        Paint
    }

    function Stop-KeepAlive {
        $script:running = $false
        $timer.Stop()
        $numInt.Enabled = $true
        Set-Awake $false
        Paint
    }

    $btnRun.Add_Click({
        if ($script:running) { Stop-KeepAlive } else { Start-KeepAlive }
    })
    $btnQuit.Add_Click({ $f.Close() })
    $chkTop.Add_CheckedChanged({ $f.TopMost = $chkTop.Checked })
    $f.Add_FormClosing({
        $timer.Stop()
        Set-Awake $false           # 終了時は必ず元に戻す
    })

    $f.TopMost = $true
    $f.Add_Shown({ Start-KeepAlive })   # 起動したらすぐ動き出す
    [void]$f.ShowDialog()

} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "起動できませんでした。`r`n`r`n$($_.Exception.Message)",
        '画面スリープ防止', 'OK', 'Error') | Out-Null
}
