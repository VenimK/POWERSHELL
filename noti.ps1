Add-Type -AssemblyName System.Windows.Forms
$global:balmsg = New-Object System.Windows.Forms.NotifyIcon
$path = (Get-Process -id $pid).Path
$balmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
$balmsg.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
$balmsg.BalloonTipText = "Music Lover Nieuwe Pc Pack Installed"
$balmsg.BalloonTipTitle = "Attention $Env:USERNAME"
$balmsg.Visible = $true
$balmsg.ShowBalloonTip(20000)