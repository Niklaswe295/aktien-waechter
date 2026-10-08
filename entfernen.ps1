# Stoppt den Aktien-Waechter (entfernt die Windows-Aufgabe). Gesammelte Daten bleiben erhalten.
Unregister-ScheduledTask -TaskName 'Aktien-Waechter' -Confirm:$false
Write-Host 'Aktien-Waechter gestoppt.' -ForegroundColor Yellow
Read-Host 'Enter zum Schliessen'
