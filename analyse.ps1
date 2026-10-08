# Muster-Analyse fuer den Aktien-Waechter (keine Anlageberatung - nur Statistik ueber die Vergangenheit).
# Teil 1: Testet typische Kursmuster auf 5 Jahren Kurshistorie der Scanner-Liste
#         ("Was waere passiert, wenn man immer bei Muster X gekauft haette?")
# Teil 2: Wertet deine eigenen gesammelten Scanner-Signale aus (signale.csv)
# Teil 3: Wie reagierten Kurse am Folgetag auf positive/negative Schlagzeilen (news.csv)?
# Ergebnis steht auch in analyse-ergebnis.txt
param([string]$Zeitraum = '5y')

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Dir = $PSScriptRoot
$Config = Get-Content (Join-Path $Dir 'config.json') -Raw | ConvertFrom-Json
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$DE = [Globalization.CultureInfo]::GetCultureInfo('de-DE')
$Out = New-Object Collections.Generic.List[string]
function Say($t = '') { Write-Host $t; $Out.Add($t) }
function Pct($x) { if ($null -eq $x) { '   -   ' } else { ('{0,6}' -f ($x * 100).ToString('+0.0;-0.0', $DE)) + ' %' } }

# --- Kurshistorie laden (Tagesschlusskurse) ---
function Get-Historie($sym, $range) {
    $r = Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/$([uri]::EscapeDataString($sym))?range=$range&interval=1d" -UserAgent $UA
    $res = $r.chart.result[0]
    $c = $res.indicators.quote[0].close
    $t = $res.timestamp
    $tage = @(); $kurse = @()
    for ($i = 0; $i -lt $t.Count; $i++) {
        if ($null -ne $c[$i] -and $c[$i] -gt 0) { $tage += [DateTimeOffset]::FromUnixTimeSeconds([long]$t[$i]).UtcDateTime.Date; $kurse += [double]$c[$i] }
    }
    [pscustomobject]@{ Symbol = $sym; Tage = $tage; Kurse = $kurse }
}

Say "Lade Kurshistorie ($Zeitraum) fuer $($Config.scannerListe.Count) Aktien + DAX ..."
$Hist = @{}
foreach ($s in @($Config.scannerListe) + '^GDAXI') {
    try { $Hist[$s] = Get-Historie $s $Zeitraum } catch { Say "  $s nicht ladbar: $($_.Exception.Message)" }
}

# --- Teil 1: Muster-Backtest ---
$Haltedauern = 5, 20, 60          # Handelstage (~1 Woche, ~1 Monat, ~3 Monate)
$Pause = 20                        # gleiches Muster pro Aktie hoechstens alle 20 Tage zaehlen
$Muster = [ordered]@{
    'Zufallskauf (Vergleichswert)'               = { param($k, $i, $c) $true }
    'Nach Tagesverlust ab -5 %'                    = { param($k, $i, $c) $k[$i] / $k[$i - 1] - 1 -le -0.05 }
    'Nach Tagesgewinn ab +5 %'                     = { param($k, $i, $c) $k[$i] / $k[$i - 1] - 1 -ge 0.05 }
    'Neues 52-Wochen-Hoch (Momentum)'              = { param($k, $i, $c) $k[$i] -ge $c.Hoch -and $k[$i - 1] -lt $c.HochVortag }
    'Ruecksetzer >=20 % + dreht ueber 50-Tage-Linie' = { param($k, $i, $c) $c.Abstand -ge 0.20 -and $k[$i] -gt $c.Linie50 -and $k[$i - 1] -le $c.Linie50Vortag }
    'Fallendes Messer: >=30 % unter Hoch, unter 50-Tage-Linie' = { param($k, $i, $c) $c.Abstand -ge 0.30 -and $k[$i] -lt $c.Linie50 * 0.95 }
    '3 Verlusttage in Folge'                       = { param($k, $i, $c) $k[$i] -lt $k[$i - 1] -and $k[$i - 1] -lt $k[$i - 2] -and $k[$i - 2] -lt $k[$i - 3] }
    'Ueber 200-Tage-Linie gestiegen'               = { param($k, $i, $c) $k[$i] -gt $c.Linie200 -and $k[$i - 1] -le $c.Linie200Vortag }
}
$Ergebnis = @{}; foreach ($m in $Muster.Keys) { $Ergebnis[$m] = @{}; foreach ($h in $Haltedauern) { $Ergebnis[$m][$h] = New-Object Collections.Generic.List[double] } }

foreach ($s in $Config.scannerListe) {
    $H = $Hist[$s]; if (-not $H) { continue }
    $k = $H.Kurse; $n = $k.Count
    if ($n -lt 300) { continue }
    $letztes = @{}
    # gleitende Werte vorab berechnen
    $sum50 = 0; $sum200 = 0
    $l50 = New-Object double[] $n; $l200 = New-Object double[] $n; $hoch = New-Object double[] $n
    for ($i = 0; $i -lt $n; $i++) {
        $sum50 += $k[$i]; if ($i -ge 50) { $sum50 -= $k[$i - 50] }
        $sum200 += $k[$i]; if ($i -ge 200) { $sum200 -= $k[$i - 200] }
        $l50[$i] = $sum50 / [math]::Min($i + 1, 50); $l200[$i] = $sum200 / [math]::Min($i + 1, 200)
        $von = [math]::Max(0, $i - 251); $mx = 0; for ($j = $von; $j -le $i; $j++) { if ($k[$j] -gt $mx) { $mx = $k[$j] } }; $hoch[$i] = $mx
    }
    for ($i = 252; $i -lt $n - 5; $i++) {
        $ctx = @{ Hoch = $hoch[$i]; HochVortag = $hoch[$i - 1]; Abstand = 1 - $k[$i] / $hoch[$i]
                  Linie50 = $l50[$i]; Linie50Vortag = $l50[$i - 1]; Linie200 = $l200[$i]; Linie200Vortag = $l200[$i - 1] }
        foreach ($m in $Muster.Keys) {
            if ($letztes[$m] -and $i - $letztes[$m] -lt $Pause -and $m -notlike 'Zufall*') { continue }
            if (-not (& $Muster[$m] $k $i $ctx)) { continue }
            $letztes[$m] = $i
            foreach ($h in $Haltedauern) { if ($i + $h -lt $n) { $Ergebnis[$m][$h].Add($k[$i + $h] / $k[$i] - 1) } }
        }
    }
}

function Median($l) { if ($l.Count -eq 0) { return $null }; $a = $l.ToArray(); [Array]::Sort($a); $a[[int][math]::Floor($a.Count / 2)] }
function Mittel($l) { if ($l.Count -eq 0) { return $null }; ($l | Measure-Object -Average).Average }
function Trefferquote($l) { if ($l.Count -eq 0) { return $null }; @($l | Where-Object { $_ -gt 0 }).Count / $l.Count }
function Schlimmster($l) { if ($l.Count -eq 0) { return $null }; ($l | Measure-Object -Minimum).Minimum }

Say ''
Say "=== TEIL 1: Kursmuster der letzten $Zeitraum (Kauf bei Muster, Verkauf nach X Handelstagen) ==="
Say 'Durchschnitt = mittlere Rendite, Treffer = Anteil Faelle mit Gewinn, Schlimmster = groesster Verlust.'
Say 'Ein Muster ist nur interessant, wenn es spuerbar BESSER als der Zufallskauf ist.'
foreach ($h in $Haltedauern) {
    Say ''
    Say "--- Haltedauer $h Handelstage ---"
    Say ('{0,-58} {1,6} {2,10} {3,10} {4,9} {5,11}' -f 'Muster', 'Faelle', 'Durchschn.', 'Median', 'Treffer', 'Schlimmster')
    foreach ($m in $Muster.Keys) {
        $l = $Ergebnis[$m][$h]
        Say ('{0,-58} {1,6} {2,10} {3,10} {4,9} {5,11}' -f $m, $l.Count, (Pct (Mittel $l)), (Pct (Median $l)),
            $(if ($l.Count) { ((Trefferquote $l) * 100).ToString('0', $DE) + ' %' } else { '-' }), (Pct (Schlimmster $l)))
    }
}

# Saisonalitaet DAX: Rendite je Kalendermonat
$dax = $Hist['^GDAXI']
if ($dax) {
    Say ''
    Say '=== DAX: durchschnittliche Rendite je Kalendermonat ==='
    $gruppen = for ($i = 0; $i -lt $dax.Kurse.Count; $i++) { [pscustomobject]@{ M = $dax.Tage[$i].ToString('yyyy-MM'); K = $dax.Kurse[$i] } }
    $monate = $gruppen | Group-Object M | ForEach-Object { [pscustomobject]@{ Monat = [int]$_.Name.Substring(5); R = $_.Group[-1].K / $_.Group[0].K - 1 } }
    foreach ($g in ($monate | Group-Object Monat | Sort-Object { [int]$_.Name })) {
        $avg = ($g.Group | Measure-Object R -Average).Average
        $pos = @($g.Group | Where-Object { $_.R -gt 0 }).Count
        Say ('{0,-10} {1}   ({2} von {3} Jahren im Plus)' -f $DE.DateTimeFormat.GetMonthName([int]$g.Name), (Pct $avg), $pos, $g.Count)
    }
}

# --- Teil 2: eigene Scanner-Signale (Papierdepot) ---
$SigFile = Join-Path $Dir 'signale.csv'
Say ''
Say '=== TEIL 2: Deine Scanner-Signale (Papierdepot) ==='
if (Test-Path $SigFile) {
    $sig = Import-Csv $SigFile -Delimiter ';'
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $daxNow = if ($dax) { $dax.Kurse[-1] } else { $null }
    $zeilen = foreach ($x in $sig) {
        $H = $Hist[$x.Symbol]; if (-not $H) { continue }
        $start = [double]::Parse($x.Kurs, $inv); $jetzt = $H.Kurse[-1]
        $d = [datetime]::ParseExact($x.Datum, 'yyyy-MM-dd', $null)
        $daxStart = $null; if ($dax) { for ($i = 0; $i -lt $dax.Tage.Count; $i++) { if ($dax.Tage[$i] -ge $d) { $daxStart = $dax.Kurse[$i]; break } } }
        $r = $jetzt / $start - 1; $rd = if ($daxStart) { $daxNow / $daxStart - 1 } else { $null }
        [pscustomobject]@{ Datum = $x.Datum; Aktie = $x.Name; Punkte = $x.Punkte; Rendite = $r; DAX = $rd }
    }
    foreach ($z in $zeilen) { Say ('{0}  {1,-40} {2,2} P.  Aktie {3}   DAX {4}' -f $z.Datum, $z.Aktie, $z.Punkte, (Pct $z.Rendite), (Pct $z.DAX)) }
    if ($zeilen) {
        $besser = @($zeilen | Where-Object { $null -ne $_.DAX -and $_.Rendite -gt $_.DAX }).Count
        Say "=> $besser von $(@($zeilen).Count) Signalen liegen vor dem DAX. Durchschnitt: $(Pct (($zeilen | Measure-Object Rendite -Average).Average))"
    }
} else { Say 'Noch keine Signale gesammelt - der Waechter legt signale.csv beim ersten Schnaeppchen-Push an.' }

# --- Teil 3: Reaktion auf Schlagzeilen ---
$NewsFile = Join-Path $Dir 'news.csv'
Say ''
Say '=== TEIL 3: Kursreaktion am naechsten Handelstag nach Schlagzeilen ==='
if (Test-Path $NewsFile) {
    $news = Import-Csv $NewsFile -Delimiter ';' | Where-Object { $_.Symbole -and $_.Signal -match 'positiv|negativ' }
    $react = @{ positiv = New-Object Collections.Generic.List[double]; negativ = New-Object Collections.Generic.List[double] }
    foreach ($x in $news) {
        $d = [datetime]::ParseExact($x.Zeit.Substring(0, 10), 'yyyy-MM-dd', $null)
        foreach ($s in ($x.Symbole -split ',')) {
            $H = $Hist[$s]; if (-not $H) { continue }
            $i = [Array]::FindIndex([datetime[]]$H.Tage, [Predicate[datetime]] { param($t) $t -gt $d })
            if ($i -ge 1) { foreach ($sg in 'positiv', 'negativ') { if ($x.Signal -match $sg) { $react[$sg].Add($H.Kurse[$i] / $H.Kurse[$i - 1] - 1) } } }
        }
    }
    foreach ($sg in 'positiv', 'negativ') {
        $l = $react[$sg]
        Say ('{0,-8} Schlagzeilen: {1,4} Faelle, Folgetag im Schnitt {2}, Treffer {3}' -f $sg, $l.Count, (Pct (Mittel $l)),
            $(if ($l.Count) { ((Trefferquote $l) * 100).ToString('0', $DE) + ' %' } else { '-' }))
    }
    Say '(Aussagekraeftig erst nach einigen Wochen Sammeln - unter ~30 Faellen ist das reiner Zufall.)'
} else { Say 'Noch keine Nachrichten gesammelt.' }

Say ''
Say 'Hinweise: Vergangene Muster garantieren nichts. Die Scanner-Liste enthaelt die HEUTIGEN DAX-Werte'
Say '(Aktien, die abgestiegen oder pleite sind, fehlen) - das schoent die Ergebnisse etwas. Gebuehren/Steuern nicht eingerechnet.'
$Out | Set-Content (Join-Path $Dir 'analyse-ergebnis.txt') -Encoding UTF8
