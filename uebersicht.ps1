# Erzeugt den Bericht neu und oeffnet ihn im Browser. Mit -Analyse zusaetzlich die Muster-Analyse.
param([switch]$Analyse)
& "$PSScriptRoot\bericht.ps1" -Oeffnen
if ($Analyse) { & "$PSScriptRoot\waechter.ps1" -Scan; & "$PSScriptRoot\analyse.ps1"; Read-Host 'Enter zum Schliessen' }
