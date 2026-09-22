$wshell = New-Object -ComObject WScript.Shell
$wt = Get-Process -Name WindowsTerminal -ErrorAction SilentlyContinue | Select-Object -First 1
if ($wt) {
    $res = $wshell.AppActivate($wt.Id)
    Write-Output "AppActivate by PID ($($wt.Id)): $res"
} else {
    Write-Output "WindowsTerminal not found"
}
