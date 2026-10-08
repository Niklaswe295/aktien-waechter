# Guru-Depots: liest die Depots bekannter Investoren von dataroma.com (Quelle sind die
# Pflichtmeldungen "13F" an die US-Boersenaufsicht SEC - erscheinen bis 45 Tage nach Quartalsende).
# Speichert je Investor gurus\<code>.json und schickt einen Push, sobald ein neues Quartal da ist.
# Investoren stehen in config.json unter "gurus" (Name -> Dataroma-Kuerzel, z. B. BRK = Berkshire Hathaway).
param([switch]$KeinPush)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Dir = $PSScriptRoot
$Config = Get-Content (Join-Path $Dir 'config.json') -Raw | ConvertFrom-Json
# In der Cloud (GitHub Actions) kommt das ntfy-Thema aus einem Geheimnis statt aus config.json
if ($env:NTFY_TOPIC) { $Config.ntfyTopic = $env:NTFY_TOPIC }
elseif (Test-Path (Join-Path $Dir 'ntfy-thema.txt')) { $Config.ntfyTopic = (Get-Content (Join-Path $Dir 'ntfy-thema.txt') -Raw).Trim() }   # Laptop, nicht im Repository
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$inv = [Globalization.CultureInfo]::InvariantCulture
$GuruDir = Join-Path $Dir 'gurus'
if (-not (Test-Path $GuruDir)) { New-Item -ItemType Directory $GuruDir | Out-Null }

function Zahl($t) { $t = ($t -replace '[\$,%\s]', ''); if ($t -match '^-?[\d.]+$') { [double]::Parse($t, $inv) } else { $null } }

foreach ($g in $Config.gurus.PSObject.Properties) {
    $name = $g.Name; $code = $g.Value
    $h = (Invoke-WebRequest "https://www.dataroma.com/m/holdings.php?m=$code" -UserAgent $UA -UseBasicParsing -TimeoutSec 30).Content
    $periode = if ($h -match 'Period:\s*(?:<[^>]+>\s*)*([^<]+?)\s*<') { $Matches[1].Trim() } else { '?' }
    $stichtag = if ($h -match 'Portfolio date:\s*(?:<[^>]+>)*\s*([^<]+?)\s*<') { $Matches[1].Trim() } else { '' }
    $positionen = foreach ($row in [regex]::Matches($h, '(?s)<tr>\s*<td class="hist">.*?</tr>')) {
        $td = @([regex]::Matches($row.Value, '(?s)<td[^>]*>(.*?)</td>') | ForEach-Object { ($_.Groups[1].Value -replace '<span>', ' |' -replace '<[^>]+>', '').Trim() })
        if ($td.Count -lt 12) { continue }
        $sym, $firma = $td[1] -split '\s*\|\s*-\s*', 2
        $akt = [System.Net.WebUtility]::HtmlDecode($td[3])
        $art = switch -regex ($akt) { '^Buy' { 'Neu gekauft' } '^Add' { 'Aufgestockt' } '^Reduce' { 'Reduziert' } '^Sell' { 'Verkauft' } default { '' } }
        [pscustomobject]@{
            symbol = $sym.Trim(); firma = [System.Net.WebUtility]::HtmlDecode($firma); anteil = (Zahl $td[2])
            aktivitaet = $art; aenderung = $(if ($akt -match '([\d.]+)%') { [double]::Parse($Matches[1], $inv) } else { $null })
            stueck = (Zahl $td[4]); kursQuartal = (Zahl $td[5]); wert = (Zahl $td[6]); kursAktuell = (Zahl $td[8]); seitQuartal = (Zahl $td[9])
        }
    }
    if (-not $positionen) { throw "Dataroma: keine Positionen fuer $code gefunden (Seite geaendert?)" }

    $datei = Join-Path $GuruDir "$code.json"
    $alt = if (Test-Path $datei) { Get-Content $datei -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
    [pscustomobject]@{ name = $name; code = $code; periode = $periode; stichtag = $stichtag; abgerufen = (Get-Date -Format 'yyyy-MM-dd')
                       positionen = @($positionen) } | ConvertTo-Json -Depth 5 | Set-Content $datei -Encoding UTF8

    # Neues Quartal -> Push mit den Aenderungen
    if (-not $KeinPush -and -not $Config.nurTagesbericht -and $alt -and $alt.periode -ne $periode) {
        $bew = @($positionen | Where-Object aktivitaet | Sort-Object anteil -Descending | Select-Object -First 8 |
            ForEach-Object { "$($_.aktivitaet): $($_.firma) ($($_.symbol))$(if ($_.aenderung) { " $($_.aenderung) %" })" })
        $msg = "Neue Depotmeldung fuer $periode`n" + $(if ($bew) { $bew -join "`n" } else { 'Keine Kaeufe oder Verkaeufe.' })
        $body = @{ topic = $Config.ntfyTopic; title = "${name}: neues Quartal"; message = $msg; priority = 4; tags = @('briefcase')
                   click = "https://www.dataroma.com/m/holdings.php?m=$code" }
        Invoke-RestMethod -Method Post -Uri 'https://ntfy.sh/' -Body ([Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress))) -ContentType 'application/json; charset=utf-8' | Out-Null
    }
    Write-Output "$name ($periode): $(@($positionen).Count) Positionen"
    Start-Sleep -Seconds 2   # hoeflich bleiben, falls mehrere Investoren eingetragen sind
}
