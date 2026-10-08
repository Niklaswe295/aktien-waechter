# Erzeugt bericht.html: fuer jedes Scanner-Signal (Papierdepot) und jede Watchlist-Aktie
# Erwartung vs. tatsaechlicher Verlauf, Analysten-Kursziel und Schlagzeilen mit Quelle.
#   -Oeffnen : Bericht danach im Browser oeffnen
param([switch]$Oeffnen)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Dir = $PSScriptRoot
$Config = Get-Content (Join-Path $Dir 'config.json') -Raw | ConvertFrom-Json
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$inv = [Globalization.CultureInfo]::InvariantCulture
$Horizont = 20   # Handelstage, fuer die die Erwartung gilt

function Get-Historie($sym, $range) {
    $r = Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($sym))?range=$range&interval=1d" -UserAgent $UA
    $res = $r.chart.result[0]; $c = $res.indicators.quote[0].close; $t = $res.timestamp
    $tage = New-Object Collections.Generic.List[string]; $kurse = New-Object Collections.Generic.List[double]
    for ($i = 0; $i -lt $t.Count; $i++) {
        if ($null -ne $c[$i] -and $c[$i] -gt 0) { $tage.Add([DateTimeOffset]::FromUnixTimeSeconds([long]$t[$i]).UtcDateTime.ToString('yyyy-MM-dd')); $kurse.Add([double]$c[$i]) }
    }
    [pscustomobject]@{ Tage = $tage.ToArray(); Kurse = $kurse.ToArray(); Name = $res.meta.longName; Waehrung = $res.meta.currency }
}

# --- Erwartungskorridor: Wie entwickelten sich Aktien der Scanner-Liste in den letzten 5 Jahren,
#     wenn sie >= 20 % unter ihrem 52-Wochen-Hoch lagen (wie fast alle Scanner-Signale)?
#     Ergebnis wird 7 Tage zwischengespeichert (korridor.json).
$KorridorFile = Join-Path $Dir 'korridor.json'
if ((Test-Path $KorridorFile) -and (Get-Item $KorridorFile).LastWriteTime -gt (Get-Date).AddDays(-7)) {
    $Korridor = Get-Content $KorridorFile -Raw | ConvertFrom-Json
} else {
    Write-Host 'Berechne Erwartungskorridor aus 5 Jahren Kurshistorie (dauert etwas) ...'
    $pfade = New-Object Collections.Generic.List[double[]]
    foreach ($s in $Config.scannerListe) {
        try { $k = (Get-Historie $s '5y').Kurse } catch { continue }
        $letztes = -999
        for ($i = 252; $i -lt $k.Count - $Horizont; $i++) {
            if ($i - $letztes -lt $Horizont) { continue }
            $mx = 0; for ($j = $i - 251; $j -le $i; $j++) { if ($k[$j] -gt $mx) { $mx = $k[$j] } }
            if (1 - $k[$i] / $mx -lt 0.20) { continue }
            $p = New-Object double[] ($Horizont + 1)
            for ($d = 0; $d -le $Horizont; $d++) { $p[$d] = ($k[$i + $d] / $k[$i] - 1) * 100 }
            $pfade.Add($p); $letztes = $i
        }
    }
    $q = { param($arr, $f) $a = [double[]]$arr; [Array]::Sort($a); $a[[int][math]::Floor(($a.Count - 1) * $f)] }
    $Korridor = [pscustomobject]@{ Faelle = $pfade.Count; p25 = @(); p50 = @(); p75 = @(); Mittel = @() }
    for ($d = 0; $d -le $Horizont; $d++) {
        $werte = foreach ($p in $pfade) { $p[$d] }
        $Korridor.p25 += [math]::Round((& $q $werte 0.25), 2)
        $Korridor.p50 += [math]::Round((& $q $werte 0.50), 2)
        $Korridor.p75 += [math]::Round((& $q $werte 0.75), 2)
        $Korridor.Mittel += [math]::Round(($werte | Measure-Object -Average).Average, 2)
    }
    $Korridor | ConvertTo-Json -Compress | Set-Content $KorridorFile -Encoding UTF8
}

# --- Analysten-Konsens (Yahoo Finance) ---
$sess = New-Object Microsoft.PowerShell.Commands.WebRequestSession
try { Invoke-WebRequest -Uri 'https://fc.yahoo.com' -UserAgent $UA -WebSession $sess -UseBasicParsing | Out-Null } catch {}
$crumb = try { (Invoke-WebRequest -Uri 'https://query2.finance.yahoo.com/v1/test/getcrumb' -UserAgent $UA -WebSession $sess -UseBasicParsing).Content } catch { '' }
function Get-Analysten($sym) {
    try {
        $r = Invoke-RestMethod -Uri "https://query2.finance.yahoo.com/v10/finance/quoteSummary/$([uri]::EscapeDataString($sym))?modules=financialData&crumb=$([uri]::EscapeDataString($crumb))" -UserAgent $UA -WebSession $sess
        $f = $r.quoteSummary.result[0].financialData
        if (-not $f.targetMeanPrice.raw) { return $null }
        [pscustomobject]@{ ziel = $f.targetMeanPrice.raw; hoch = $f.targetHighPrice.raw; tief = $f.targetLowPrice.raw
                           anzahl = $f.numberOfAnalystOpinions.raw; empfehlung = $f.recommendationKey }
    } catch { $null }
}

# --- Schlagzeilen je Aktie ---
$News = @()
$NewsFile = Join-Path $Dir 'news.csv'
if (Test-Path $NewsFile) { $News = @(Import-Csv $NewsFile -Delimiter ';' | Where-Object { $_.Symbole }) }
function Get-Schlagzeilen($sym) {
    @($News | Where-Object { ($_.Symbole -split ',') -contains $sym } | Sort-Object Zeit -Descending | Select-Object -First 8 |
        ForEach-Object { [pscustomobject]@{ zeit = $_.Zeit; quelle = $_.Quelle; signal = $_.Signal; titel = $_.Titel; link = $_.Link } })
}

$Hist = @{}
function Hist($sym) { if (-not $Hist.ContainsKey($sym)) { $Hist[$sym] = Get-Historie $sym '1y' }; $Hist[$sym] }
$Dax = Hist '^GDAXI'
function DaxAm($datum) { $i = [Array]::FindIndex($Dax.Tage, [Predicate[string]] { param($t) $t -ge $datum }); if ($i -lt 0) { $i = $Dax.Tage.Count - 1 }; $Dax.Kurse[$i] }

# --- 1) Papierdepot: Scanner-Signale ---
$Signale = @()
$SigFile = Join-Path $Dir 'signale.csv'
if (Test-Path $SigFile) {
    foreach ($x in (Import-Csv $SigFile -Delimiter ';')) {
        try { $H = Hist $x.Symbol } catch { continue }
        $start = [double]::Parse($x.Kurs, $inv)
        $i0 = [Array]::FindIndex($H.Tage, [Predicate[string]] { param($t) $t -ge $x.Datum })
        if ($i0 -lt 0) { $i0 = $H.Tage.Count - 1 }
        $von = [math]::Max(0, $i0 - $Horizont)
        $punkte = for ($i = $von; $i -lt $H.Tage.Count; $i++) {
            [pscustomobject]@{ d = $i - $i0; datum = $H.Tage[$i]; aktie = [math]::Round(($H.Kurse[$i] / $start - 1) * 100, 2)
                               dax = [math]::Round(((DaxAm $H.Tage[$i]) / (DaxAm $x.Datum) - 1) * 100, 2); kurs = $H.Kurse[$i] }
        }
        $Signale += [pscustomobject]@{
            symbol = $x.Symbol; name = $x.Name; datum = $x.Datum; kurs = $start; waehrung = $x.Waehrung; punkte = [int]$x.Punkte
            gruende = @($x.Gruende -split ', '); verlauf = @($punkte); analysten = (Get-Analysten $x.Symbol); news = (Get-Schlagzeilen $x.Symbol)
        }
    }
}

# --- 2) Watchlist ---
$Watch = foreach ($w in $Config.watchlist) {
    try { $H = Hist $w.symbol } catch { continue }
    $von = [math]::Max(0, $H.Tage.Count - 90)
    [pscustomobject]@{
        symbol = $w.symbol; name = $w.name; waehrung = $H.Waehrung; kaufenUnter = $w.kaufenUnter; verkaufenUeber = $w.verkaufenUeber
        tage = @($H.Tage[$von..($H.Tage.Count - 1)]); kurse = @($H.Kurse[$von..($H.Kurse.Count - 1)] | ForEach-Object { [math]::Round($_, 2) })
        analysten = (Get-Analysten $w.symbol); news = (Get-Schlagzeilen $w.symbol)
    }
}

# --- 3) Guru-Depots (von gurus.ps1 gespeichert) ---
$GuruNews = @{ 'BRK' = 'BRK-B' }   # unter welchem Kuerzel die Schlagzeilen in news.csv stehen
$Gurus = foreach ($f in (Get-ChildItem (Join-Path $Dir 'gurus') -Filter *.json -ErrorAction SilentlyContinue)) {
    $g = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $g | Add-Member news (Get-Schlagzeilen $GuruNews[$g.code]) -Force
    $g
}
$Bekannt = @($Config.watchlist | ForEach-Object { $_.symbol }) + @($Config.scannerListe)

$Daten = [pscustomobject]@{
    erstellt = (Get-Date -Format 'dd.MM.yyyy HH:mm'); horizont = $Horizont; korridor = $Korridor
    signale = @($Signale); watchlist = @($Watch); gurus = @($Gurus); bekannt = $Bekannt
}
$json = $Daten | ConvertTo-Json -Depth 8 -Compress
$html = (Get-Content (Join-Path $Dir 'bericht-vorlage.html') -Raw -Encoding UTF8).Replace('/*__DATEN__*/null', $json)
$ziel = Join-Path $Dir 'bericht.html'
[IO.File]::WriteAllText($ziel, $html, (New-Object Text.UTF8Encoding $false))
$docs = Join-Path $Dir 'docs'
if (Test-Path $docs) { [IO.File]::WriteAllText((Join-Path $docs 'index.html'), $html, (New-Object Text.UTF8Encoding $false)) }
Write-Host "Bericht erstellt: $ziel"
if ($Oeffnen) { Start-Process $ziel }
