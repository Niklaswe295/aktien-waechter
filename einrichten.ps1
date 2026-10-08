# Richtet den Aktien-Waechter als Windows-Aufgabe ein (alle 15 Minuten, unsichtbar).
# Entfernen: Rechtsklick auf entfernen.ps1 -> "Mit PowerShell ausfuehren"
$Dir = $PSScriptRoot

$vbs = "CreateObject(""WScript.Shell"").Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """"$Dir\waechter.ps1"""""", 0, False"
[IO.File]::WriteAllText("$Dir\start-unsichtbar.vbs", $vbs)

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument """$Dir\start-unsichtbar.vbs"""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 15)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName 'Aktien-Waechter' -Action $action -Trigger $trigger -Settings $settings -Force `
    -Description 'Kursalarme, Nachrichten und Schnaeppchen-Scanner fuer Einzelaktien' | Out-Null

Write-Host 'Fertig! Der Aktien-Waechter laeuft jetzt alle 15 Minuten im Hintergrund.' -ForegroundColor Green
Read-Host 'Enter zum Schliessen'
