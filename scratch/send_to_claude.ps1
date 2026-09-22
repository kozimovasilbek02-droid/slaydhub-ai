$wshell = New-Object -ComObject WScript.Shell
$targets = @("Loyihani", "Claude", "Windows Terminal", "cmd", "PowerShell")
$activated = $false
foreach ($t in $targets) {
    if ($wshell.AppActivate($t)) {
        $activated = $true
        Write-Output "Successfully activated window: $t"
        Start-Sleep -Milliseconds 400
        $wshell.SendKeys("^v")
        Start-Sleep -Milliseconds 300
        $wshell.SendKeys("{ENTER}")
        Write-Output "Pasted and pressed Enter!"
        break
    }
}
if (-not $activated) {
    Write-Output "Could not automatically activate window. Clipboard is ready for manual paste."
}
