# Aktien-Waechter: prueft die Watchlist auf Kursalarme (Zielkurse, starke Tagesbewegung)
# und sucht einmal am Tag nach "Schnaeppchen" in der Scanner-Liste. Meldet per ntfy-Push.
# Einstellungen in config.json. Daten: Yahoo Finance (ca. 15 Min. verzoegert).
# Keine Anlageberatung - der Scanner bewertet nur mechanisch nach festen Regeln.
#   -Test   : schickt eine Testnachricht
#   -Scan   : zeigt die Scanner-Bewertung aller Aktien in der Konsole (ohne Push)
#   -Jetzt  : fuehrt den Scanner sofort aus (mit Push), egal welche Uhrzeit
#   -Ueberblick : schickt den Abend-Ueberblick (mit Bericht) sofort aufs Handy
param([switch]$Test, [switch]$Scan, [switch]$Jetzt, [switch]$Ueberblick)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Dir = $PSScriptRoot
$Config = Get-Content (Join-Path $Dir 'config.json') -Raw | ConvertFrom-Json
# In der Cloud (GitHub Actions) kommt das ntfy-Thema aus einem Geheimnis statt aus config.json
if ($env:NTFY_TOPIC) { $Config.ntfyTopic = $env:NTFY_TOPIC }
elseif (Test-Path (Join-Path $Dir 'ntfy-thema.txt')) { $Config.ntfyTopic = (Get-Content (Join-Path $Dir 'ntfy-thema.txt') -Raw).Trim() }   # Laptop, nicht im Repository
$StateFile = Join-Path $Dir 'zustand.txt'
$LogFile = Join-Path $Dir 'alarme.csv'
$ErrFile = Join-Path $Dir 'fehler.log'
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$DE = [Globalization.CultureInfo]::GetCultureInfo('de-DE')

# --- Zustand (was wurde wann schon gemeldet) ---
$State = @{}
if (Test-Path $StateFile) {
    Get-Content $StateFile | ForEach-Object { $k, $v = $_ -split "`t", 2; if ($k) { $State[$k] = $v } }
}
function Save-State { $State.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)`t$($_.Value)" } | Set-Content $StateFile -Encoding UTF8 }

function Fmt($n, $d = 2) { if ($null -eq $n) { '-' } else { ([double]$n).ToString("N$d", $DE) } }

# --- Kurse holen (Yahoo Finance, alle Symbole in einer Anfrage) ---
function Get-Quotes([string[]]$symbols) {
    $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    try { Invoke-WebRequest -Uri 'https://fc.yahoo.com' -UserAgent $UA -WebSession $s -UseBasicParsing | Out-Null } catch {}
    $crumb = (Invoke-WebRequest -Uri 'https://query2.finance.yahoo.com/v1/test/getcrumb' -UserAgent $UA -WebSession $s -UseBasicParsing).Content
    $result = @{}
    for ($i = 0; $i -lt $symbols.Count; $i += 40) {
        $chunk = ($symbols[$i..([math]::Min($i + 39, $symbols.Count - 1))] | ForEach-Object { [uri]::EscapeDataString($_) }) -join ','
        $q = Invoke-RestMethod -Uri "https://query2.finance.yahoo.com/v7/finance/quote?symbols=$chunk&crumb=$([uri]::EscapeDataString($crumb))" -UserAgent $UA -WebSession $s
        foreach ($r in $q.quoteResponse.result) { $result[$r.symbol] = $r }
    }
    $result
}

function Get-Name($q, $fallback) {
    if ($fallback) { return $fallback }
    if ($q.longName) { return $q.longName }
    ($q.shortName -replace '\s+\S?$', '').Trim()
}

function Send-Push($title, $message, $tags, $prio = 4, $symbol, [switch]$Immer) {
    # Im Beobachtungsmodus (nurTagesbericht) nur protokollieren - erscheint im Abendbericht
    if ($Immer -or -not $Config.nurTagesbericht) {
    $body = @{ topic = $Config.ntfyTopic; title = $title; message = $message; priority = $prio; tags = @($tags) }
    if ($symbol) { $body.click = "https://finance.yahoo.com/quote/$symbol" }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress))
    Invoke-RestMethod -Method Post -Uri 'https://ntfy.sh/' -Body $bytes -ContentType 'application/json; charset=utf-8' | Out-Null
    }
    if (-not (Test-Path $LogFile)) { 'Zeit;Titel;Nachricht' | Set-Content $LogFile -Encoding UTF8 }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm');$($title -replace ';', ',');$(($message -replace "`n", ' | ') -replace ';', ',')" | Add-Content $LogFile -Encoding UTF8
}

# --- Schnaeppchen-Regeln: Punkte fuer "guenstig bewertet, aber solide" ---
# Idee: deutlich unter dem 52-Wochen-Hoch, niedriges erwartetes KGV, Dividende,
# Analysten positiv, und der Kurs hat schon begonnen zu drehen (ueber 50-Tage-Linie).
function Get-Bewertung($q) {
    $p = 0; $gruende = @()
    $abstand = if ($q.fiftyTwoWeekHigh) { (1 - $q.regularMarketPrice / $q.fiftyTwoWeekHigh) * 100 } else { 0 }
    if ($abstand -ge 35) { $p += 3; $gruende += "$(Fmt $abstand 0) % unter 52W-Hoch" }
    elseif ($abstand -ge 20) { $p += 2; $gruende += "$(Fmt $abstand 0) % unter 52W-Hoch" }
    elseif ($abstand -lt 5) { $p -= 1 }

    $kgv = $q.forwardPE
    if ($null -eq $kgv -or $kgv -le 0) { $p -= 2; $gruende += 'kein/negatives KGV' }
    elseif ($kgv -lt 4) { }                                  # unplausibel (z. B. Holding) -> ignorieren
    elseif ($kgv -lt 12) { $p += 2; $gruende += "KGV $(Fmt $kgv 1)" }
    elseif ($kgv -lt 16) { $p += 1; $gruende += "KGV $(Fmt $kgv 1)" }
    elseif ($kgv -gt 35) { $p -= 2; $gruende += "teuer: KGV $(Fmt $kgv 0)" }

    $div = $q.dividendYield
    if ($div -ge 5 -and $div -lt 12) { $p += 2; $gruende += "Dividende $(Fmt $div 1) %" }
    elseif ($div -ge 3 -and $div -lt 12) { $p += 1; $gruende += "Dividende $(Fmt $div 1) %" }

    if ($q.averageAnalystRating -match '^([\d.]+)') {
        $r = [double]::Parse($Matches[1], [Globalization.CultureInfo]::InvariantCulture)
        if ($r -le 1.8) { $p += 2; $gruende += "Analysten: $($q.averageAnalystRating)" }
        elseif ($r -le 2.3) { $p += 1; $gruende += "Analysten: $($q.averageAnalystRating)" }
        elseif ($r -ge 3) { $p -= 1; $gruende += "Analysten: $($q.averageAnalystRating)" }
    }

    if ($q.fiftyDayAverage -and $abstand -ge 15) {
        if ($q.regularMarketPrice -gt $q.fiftyDayAverage) { $p += 1; $gruende += 'dreht nach oben (ueber 50-Tage-Linie)' }
        elseif ($q.fiftyDayAverageChangePercent -lt -0.10) { $p -= 1; $gruende += 'faellt noch stark' }
    }
    [pscustomobject]@{ Punkte = $p; Abstand = $abstand; Gruende = $gruende }
}

# --- Nachrichten: Schlagzeilen aus den Feeds in config.json den Aktien zuordnen ---
# Text wird klein geschrieben, Umlaute -> ae/oe/ue/ss (wie beim Kleinanzeigen-Finder)
function Normalize($t) { $t.ToLower().Replace([string][char]0xE4, 'ae').Replace([string][char]0xF6, 'oe').Replace([string][char]0xFC, 'ue').Replace([string][char]0xDF, 'ss') }
$NewsSignale = [ordered]@{
    positiv = 'kursziel\w* (erhoeht|angehoben|hoch)|kurszielerhoehung|kurszielanhebung|hochgestuft|hochstufung|kaufempfehlung|kaufen\b|\bbuy\b|outperform|overweight|uebertrifft|uebertroffen|prognose (erhoeht|angehoben)|rekordgewinn|rekordumsatz|aktienrueckkauf|uebernahmeangebot|uebernahmepraemie|upgrade|raises? (price )?target|target raised|beats? (estimates|expectations)|tops estimates|record (profit|revenue|sales)|buyback|takeover (bid|offer)|soars|surges|jumps'
    negativ = 'kursziel\w* (gesenkt|runter|gekappt)|kurszielsenkung|abgestuft|herabgestuft|abstufung|verkaufen\b|\bsell\b|underperform|underweight|gewinnwarnung|prognose (gesenkt|kassiert|gekappt)|verfehlt|bricht ein|einbruch|absturz|ermittlung|razzia|stellenabbau|insolvenz|downgrade|cuts? (price )?target|target cut|miss(es)? (estimates|expectations)|profit warning|cuts? (outlook|forecast|guidance)|plunges|slumps|tumbles|sinks|probe|lawsuit|layoffs|job cuts|bankruptcy'
    neuling = 'boersengang|\bipo\b|neuemission|erstnotiz|boersendebuet|boersenneuling|zeichnungsfrist|preisspanne|initial public offering|goes public|trading debut|market debut'
}
function Update-News {
    $NewsFile = Join-Path $Dir 'news.csv'
    $SeenFile = Join-Path $Dir 'news-gesehen.txt'
    $seen = @{}; if (Test-Path $SeenFile) { Get-Content $SeenFile | ForEach-Object { $seen[$_] = 1 } }
    $ersterLauf = -not (Test-Path $SeenFile)   # beim allerersten Lauf nur sammeln, nicht pushen
    if (-not (Test-Path $NewsFile)) { 'Zeit;Quelle;Symbole;Signal;Titel;Link' | Set-Content $NewsFile -Encoding UTF8 }
    $watch = @($Config.watchlist | ForEach-Object { $_.symbol })
    $neu = @()
    foreach ($f in $Config.feeds.PSObject.Properties) {
        try {
            $r = Invoke-WebRequest -Uri $f.Value -UserAgent $UA -UseBasicParsing -TimeoutSec 20
            $ms = New-Object IO.MemoryStream; $r.RawContentStream.CopyTo($ms)
            $x = [xml][Text.Encoding]::UTF8.GetString($ms.ToArray()).TrimStart([char]0xFEFF)
        } catch { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Feed $($f.Name): $($_.Exception.Message)" | Add-Content $ErrFile -Encoding UTF8; continue }
        foreach ($it in @($x.rss.channel.item)) {
            if (-not $it) { continue }
            $titel = ([Net.WebUtility]::HtmlDecode($it.SelectSingleNode('title').InnerText) -replace '\s+', ' ').Trim()
            $link = $it.SelectSingleNode('link').InnerText.Trim()
            $key = if ($link) { $link } else { $titel }
            if (-not $titel -or $seen[$key]) { continue }
            $seen[$key] = 1; $neu += $key
            $desc = $it.SelectSingleNode('description'); $desc = if ($desc) { $desc.InnerText -replace '<[^>]+>', ' ' } else { '' }
            $t = Normalize "$titel $desc"
            $syms = @($Config.namen.PSObject.Properties | Where-Object { $t -match $_.Value } | ForEach-Object { $_.Name })
            $signal = @($NewsSignale.Keys | Where-Object { $t -match $NewsSignale[$_] })
            if (-not $syms -and $signal -notcontains 'neuling') { continue }   # nur Relevantes speichern
            $sig = $signal -join '+'
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm');$($f.Name);$($syms -join ',');$sig;$($titel -replace ';', ',');$link" | Add-Content $NewsFile -Encoding UTF8
            # Push: Watchlist-Aktie mit klarem Signal, oder ein Boersenneuling
            $wTreffer = @($syms | Where-Object { $watch -contains $_ })
            if ($ersterLauf -or $Config.nurTagesbericht) { }   # still sammeln - steht im Abendbericht
            elseif ($wTreffer -and ($signal -contains 'positiv' -or $signal -contains 'negativ')) {
                $tag = if ($signal -contains 'negativ') { 'warning' } else { 'newspaper' }
                $body = @{ topic = $Config.ntfyTopic; title = "$($f.Name): $($wTreffer -join ', ')"; message = $titel; priority = 3; tags = @($tag); click = $link }
                Invoke-RestMethod -Method Post -Uri 'https://ntfy.sh/' -Body ([Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress))) -ContentType 'application/json; charset=utf-8' | Out-Null
            } elseif ($signal -contains 'neuling' -and $State["ipo-$(Get-Date -Format 'yyyy-MM-dd')"] -lt 3) {
                $body = @{ topic = $Config.ntfyTopic; title = "Boersenneuling ($($f.Name))"; message = $titel; priority = 2; tags = @('baby'); click = $link }
                Invoke-RestMethod -Method Post -Uri 'https://ntfy.sh/' -Body ([Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress))) -ContentType 'application/json; charset=utf-8' | Out-Null
                $State["ipo-$(Get-Date -Format 'yyyy-MM-dd')"] = [int]$State["ipo-$(Get-Date -Format 'yyyy-MM-dd')"] + 1
            }
        }
    }
    $neu | Add-Content $SeenFile -Encoding UTF8
    # Gesehen-Liste klein halten
    $all = @(Get-Content $SeenFile); if ($all.Count -gt 5000) { $all[-3000..-1] | Set-Content $SeenFile -Encoding UTF8 }
}

# --- Abend-Ueberblick: kurze Zusammenfassung + kompletter Bericht als Datei-Anhang ---
# ntfy.sh bewahrt Anhaenge ein paar Stunden auf - fuer den Abend reicht das.
function Send-Tagesbericht($Quotes) {
    $heute = Get-Date -Format 'yyyy-MM-dd'
    $vz = { param($x) "$(if ($x -ge 0) { '+' })$(Fmt $x 1) %" }
    $zeilen = @()
    $dax = $Quotes['^GDAXI']
    if ($dax) { $zeilen += "DAX $(Fmt $dax.regularMarketPrice 0) ($(& $vz $dax.regularMarketChangePercent))" }

    $zeilen += ''
    $zeilen += 'Watchlist:'
    foreach ($w in $Config.watchlist) {
        $q = $Quotes[$w.symbol]; if (-not $q) { continue }
        $hinweis = if ($w.kaufenUnter -and $q.regularMarketPrice -le $w.kaufenUnter) { ' <- Kaufziel' } elseif ($w.verkaufenUeber -and $q.regularMarketPrice -ge $w.verkaufenUeber) { ' <- Verkaufsziel' } else { '' }
        $zeilen += "$(Get-Name $q $w.name): $(Fmt $q.regularMarketPrice) $($q.currency) ($(& $vz $q.regularMarketChangePercent))$hinweis"
    }

    $SigFile = Join-Path $Dir 'signale.csv'
    if ((Test-Path $SigFile) -and $dax) {
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $sig = @(Import-Csv $SigFile -Delimiter ';')
        $r = foreach ($x in $sig) { $q = $Quotes[$x.Symbol]; if ($q) { ($q.regularMarketPrice / [double]::Parse($x.Kurs, $inv) - 1) * 100 } }
        if ($r) {
            $zeilen += ''
            $zeilen += "Papierdepot: $(@($r).Count) Signale, im Schnitt $(& $vz (($r | Measure-Object -Average).Average)) seit Signal"
        }
        $neu = @($sig | Where-Object Datum -eq $heute | ForEach-Object Name)
        if ($neu) { $zeilen += "Neue Signale heute: $($neu -join ', ')" }
    }

    $NewsFile = Join-Path $Dir 'news.csv'
    if (Test-Path $NewsFile) {
        $n = @(Import-Csv $NewsFile -Delimiter ';' | Where-Object { $_.Zeit -like "$heute*" })
        if ($n) {
            $zeilen += ''
            $zeilen += "Schlagzeilen heute: $($n.Count) ($(@($n | Where-Object Signal -match 'positiv').Count) positiv, $(@($n | Where-Object Signal -match 'negativ').Count) negativ)"
            $watch = @($Config.watchlist | ForEach-Object { $_.symbol })
            $n | Where-Object { $_.Signal -match 'positiv|negativ' -and (@($_.Symbole -split ',') | Where-Object { $watch -contains $_ }) } |
                Select-Object -Last 3 | ForEach-Object { $zeilen += "- $($_.Titel) ($($_.Quelle))" }
        }
    }
    # Kursalarme des Tages (im Beobachtungsmodus nur protokolliert, nicht gepusht)
    if (Test-Path $LogFile) {
        $h = @(Import-Csv $LogFile -Delimiter ';' | Where-Object { $_.Zeit -like "$heute*" -and $_.Titel -notlike 'Schnaeppchen*' -and $_.Titel -notlike 'TEST*' })
        if ($h) { $zeilen += ''; $zeilen += 'Hinweise heute:'; $h | Select-Object -Last 5 | ForEach-Object { $zeilen += "- $($_.Titel)" } }
    }
    # Neue Depotmeldung eines Gurus?
    foreach ($gf in (Get-ChildItem (Join-Path $Dir 'gurus') -Filter *.json -ErrorAction SilentlyContinue)) {
        $g = Get-Content $gf.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($State["guru-$($g.code)"] -and $State["guru-$($g.code)"] -ne $g.periode) {
            $zeilen += ''; $zeilen += "$($g.name): neues Depot fuer $($g.periode) - Details im Bericht"
        }
        $State["guru-$($g.code)"] = $g.periode
    }
    $zeilen += ''
    $zeilen += 'Tippen oeffnet den kompletten Bericht mit Diagrammen (ca. 3 Std. abrufbar).'
    if ($Config.berichtUrl) {
        # Cloud: Bericht liegt dauerhaft auf GitHub Pages (Zeitstempel gegen alten Zwischenspeicher)
        $zeilen[-1] = 'Tippen oeffnet den kompletten Bericht mit Diagrammen.'
        $up = @{ attachment = @{ url = "$($Config.berichtUrl)?d=$(Get-Date -Format 'yyyyMMddHHmm')" } }
    } else {
    # 1) Bericht auf einem unabonnierten Neben-Thema hochladen -> Download-Link
    $up = Invoke-RestMethod -Method Put -Uri "https://ntfy.sh/$($Config.ntfyTopic)-dateien?filename=Aktien-Bericht-$(Get-Date -Format 'dd-MM').html" `
        -InFile (Join-Path $Dir 'bericht.html') -ContentType 'text/html'
    }
    # 2) Ein einziger Push: Text + Tippen/Knopf oeffnet den Bericht im Browser (dort laufen die Diagramme)
    $body = @{ topic = $Config.ntfyTopic; title = "Aktien-Ueberblick $(Get-Date -Format 'dd.MM.')"; message = ($zeilen -join "`n")
               tags = @('bar_chart'); priority = 3; click = $up.attachment.url
               actions = @(@{ action = 'view'; label = 'Bericht oeffnen'; url = $up.attachment.url }) }
    Invoke-RestMethod -Method Post -Uri 'https://ntfy.sh/' -Body ([Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress -Depth 4))) -ContentType 'application/json; charset=utf-8' | Out-Null
}

function Get-Infozeile($q) {
    $info = @()
    if ($q.earningsTimestamp) {
        $e = [DateTimeOffset]::FromUnixTimeSeconds([long]$q.earningsTimestamp).LocalDateTime
        $tage = ($e.Date - (Get-Date).Date).Days
        if ($tage -ge 0 -and $tage -le 14) { $info += "Quartalszahlen am $($e.ToString('dd.MM.'))" }
    }
    $info -join ' - '
}

try {
    if ($Test) {
        Send-Push 'TEST - Aktien-Waechter' 'Wenn du das liest, kommt der Abendbericht an.' 'chart_with_upwards_trend' 3 -Immer
        Write-Output 'Testnachricht gesendet.'
        return
    }

    if ($Ueberblick) {
        $syms = @($Config.watchlist | ForEach-Object { $_.symbol }) + @($Config.scannerListe) + '^GDAXI'
        & (Join-Path $Dir 'bericht.ps1') | Out-Null
        Send-Tagesbericht (Get-Quotes ($syms | Select-Object -Unique))
        Save-State
        Write-Output 'Ueberblick gesendet.'
        return
    }

    $now = Get-Date
    $heute = $now.ToString('yyyy-MM-dd')
    $werktag = $now.DayOfWeek -notin 'Saturday', 'Sunday'
    $scanZeit = [datetime]::ParseExact($Config.scanner.uhrzeit, 'HH:mm', $null)
    $scanFaellig = $Jetzt -or $Scan -or ($werktag -and $now.TimeOfDay -ge $scanZeit.TimeOfDay -and $State['scan-zuletzt'] -ne $heute)
    $alarmZeit = $werktag -and $now.Hour -ge 8 -and $now.Hour -lt 23

    # Nachrichten laufen jeden Tag (6-23 Uhr), auch am Wochenende
    if (-not $Scan -and $now.Hour -ge 6) { Update-News; Save-State }

    $ueberblickFaellig = $Config.ueberblickUhrzeit -and $werktag -and $now.ToString('HH:mm') -ge $Config.ueberblickUhrzeit -and $State['ueberblick-zuletzt'] -ne $heute
    if (-not ($alarmZeit -or $scanFaellig -or $ueberblickFaellig)) { return }

    $symbole = @($Config.watchlist | ForEach-Object { $_.symbol })
    if ($scanFaellig) { $symbole += $Config.scannerListe; $symbole += '^GDAXI' }
    $Quotes = Get-Quotes ($symbole | Select-Object -Unique)

    # --- 1) Kursalarme fuer die Watchlist ---
    if ($alarmZeit -and -not $Scan) {
        foreach ($w in $Config.watchlist) {
            $q = $Quotes[$w.symbol]
            if (-not $q -or -not $q.regularMarketPrice) { continue }
            $name = Get-Name $q $w.name
            $kurs = $q.regularMarketPrice; $cur = $q.currency
            $chg = $q.regularMarketChangePercent
            # Handelstag des Kurses (damit der Vortageswert morgens nicht nochmal meldet)
            $tag = [DateTimeOffset]::FromUnixTimeSeconds([long]$q.regularMarketTime).LocalDateTime.ToString('yyyy-MM-dd')
            $zeile = "Kurs $(Fmt $kurs) $cur ($(if ($chg -ge 0) { '+' })$(Fmt $chg 1) % heute)"
            $info = Get-Infozeile $q

            if ($w.kaufenUnter -and $kurs -le $w.kaufenUnter -and $State["$($w.symbol)-unter"] -ne $tag) {
                Send-Push "$name unter $(Fmt $w.kaufenUnter) $cur" "$zeile`nDein Kaufziel ist erreicht.$(if ($info) { "`n$info" })" 'dart' 5 $w.symbol
                $State["$($w.symbol)-unter"] = $tag
            }
            if ($w.verkaufenUeber -and $kurs -ge $w.verkaufenUeber -and $State["$($w.symbol)-ueber"] -ne $tag) {
                Send-Push "$name ueber $(Fmt $w.verkaufenUeber) $cur" "$zeile`nDein Verkaufsziel ist erreicht.$(if ($info) { "`n$info" })" 'moneybag' 5 $w.symbol
                $State["$($w.symbol)-ueber"] = $tag
            }
            $grenze = if ($w.tagesAlarmProzent) { $w.tagesAlarmProzent } else { $Config.tagesAlarmProzent }
            if ([math]::Abs($chg) -ge $grenze) {
                $richtung = if ($chg -gt 0) { 'plus' } else { 'minus' }
                if ($State["$($w.symbol)-$richtung"] -ne $tag) {
                    $titel = if ($chg -gt 0) { "$name steigt +$(Fmt $chg 1) %" } else { "$name faellt $(Fmt $chg 1) %" }
                    $tags = if ($chg -gt 0) { 'chart_with_upwards_trend' } else { 'chart_with_downwards_trend' }
                    Send-Push $titel "$zeile$(if ($info) { "`n$info" })" $tags 4 $w.symbol
                    $State["$($w.symbol)-$richtung"] = $tag
                }
            }
        }
    }

    # --- 2) Schnaeppchen-Scanner (einmal taeglich) ---
    if ($scanFaellig) {
        $liste = foreach ($sym in $Config.scannerListe) {
            $q = $Quotes[$sym]
            if (-not $q -or -not $q.regularMarketPrice) { continue }
            $b = Get-Bewertung $q
            [pscustomobject]@{ Symbol = $sym; Name = (Get-Name $q $null); Kurs = $q.regularMarketPrice; Waehrung = $q.currency
                               Punkte = $b.Punkte; Gruende = $b.Gruende; Info = (Get-Infozeile $q) }
        }
        $liste = $liste | Sort-Object Punkte -Descending

        if ($Scan) {
            $liste | Select-Object Punkte, Symbol, Name, @{ n = 'Kurs'; e = { "$(Fmt $_.Kurs) $($_.Waehrung)" } }, @{ n = 'Gruende'; e = { $_.Gruende -join ', ' } } |
                Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Output
            Write-Output "Push ab $($Config.scanner.mindestPunkte) Punkten."
            return
        }

        # Tagebuch: taeglicher Schnappschuss aller Werte (fuer die spaetere Muster-Analyse)
        $VerlaufFile = Join-Path $Dir 'verlauf.csv'
        if (-not (Test-Path $VerlaufFile)) { 'Datum;Symbol;Kurs;Waehrung;AenderungProzent;UnterHochProzent;KGV;Dividende;Analysten;UeberLinie50;Punkte' | Set-Content $VerlaufFile -Encoding UTF8 }
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $zeilen = foreach ($sym in @($Config.scannerListe) + '^GDAXI') {
            $q = $Quotes[$sym]; if (-not $q) { continue }
            $a = $liste | Where-Object Symbol -eq $sym
            $unter = if ($q.fiftyTwoWeekHigh) { (1 - $q.regularMarketPrice / $q.fiftyTwoWeekHigh) * 100 } else { '' }
            @(($heute, $sym, $q.regularMarketPrice, $q.currency, $q.regularMarketChangePercent, $unter, $q.forwardPE, $q.dividendYield,
               (($q.averageAnalystRating -split ' ')[0]), [int]($q.regularMarketPrice -gt $q.fiftyDayAverage), $(if ($a) { $a.Punkte })) |
                ForEach-Object { if ($_ -is [double] -or $_ -is [decimal]) { ([double]$_).ToString('0.####', $inv) } else { "$_" } }) -join ';'
        }
        $zeilen | Add-Content $VerlaufFile -Encoding UTF8

        $grenze = (Get-Date).AddDays(-[int]$Config.scanner.wiederholenNachTagen).ToString('yyyy-MM-dd')
        foreach ($a in ($liste | Where-Object { $_.Punkte -ge $Config.scanner.mindestPunkte })) {
            # Gleiche Aktie nur alle X Tage melden - ausser die Punktzahl ist gestiegen
            $alt = $State["$($a.Symbol)-scan"]
            if ($alt) {
                $altTag, $altPunkte = $alt -split '\|'
                if ($altTag -gt $grenze -and $a.Punkte -le [int]$altPunkte) { continue }
            }
            $msg = "Kurs $(Fmt $a.Kurs) $($a.Waehrung) - $($a.Punkte) Punkte`n$($a.Gruende -join "`n")$(if ($a.Info) { "`n$($a.Info)" })"
            Send-Push "Schnaeppchen? $($a.Name)" $msg 'mag' 3 $a.Symbol
            $State["$($a.Symbol)-scan"] = "$heute|$($a.Punkte)"
            # Papierdepot: Signal mit Einstiegskurs merken - analyse.ps1 prueft spaeter, ob es sich gelohnt haette
            $SigFile = Join-Path $Dir 'signale.csv'
            if (-not (Test-Path $SigFile)) { 'Datum;Symbol;Name;Kurs;Waehrung;Punkte;Gruende' | Set-Content $SigFile -Encoding UTF8 }
            "$heute;$($a.Symbol);$($a.Name);$(([double]$a.Kurs).ToString('0.####', [Globalization.CultureInfo]::InvariantCulture));$($a.Waehrung);$($a.Punkte);$($a.Gruende -join ', ')" | Add-Content $SigFile -Encoding UTF8
        }
        $State['scan-zuletzt'] = $heute
        # Guru-Depots (Berkshire Hathaway usw.) - neues Quartal -> Push
        try { & (Join-Path $Dir 'gurus.ps1') | Out-Null }
        catch { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Gurus: $($_.Exception.Message)" | Add-Content $ErrFile -Encoding UTF8 }
        # Bericht (bericht.html) neu erzeugen - ein Fehler dort soll den Waechter nicht stoppen
        try { & (Join-Path $Dir 'bericht.ps1') | Out-Null }
        catch { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Bericht: $($_.Exception.Message)" | Add-Content $ErrFile -Encoding UTF8 }
    }

    # --- 3) Abend-Ueberblick aufs Handy (Kurztext + bericht.html als Anhang), einmal taeglich ---
    $ueZeit = if ($Config.ueberblickUhrzeit) { [datetime]::ParseExact($Config.ueberblickUhrzeit, 'HH:mm', $null) } else { $null }
    if (-not $Scan -and $ueZeit -and $werktag -and $now.TimeOfDay -ge $ueZeit.TimeOfDay -and
        $State['scan-zuletzt'] -eq $heute -and $State['ueberblick-zuletzt'] -ne $heute) {
        try {
            $syms = @($Config.watchlist | ForEach-Object { $_.symbol }) + @($Config.scannerListe) + '^GDAXI'
            & (Join-Path $Dir 'bericht.ps1') | Out-Null
            Send-Tagesbericht (Get-Quotes ($syms | Select-Object -Unique))
            $State['ueberblick-zuletzt'] = $heute
        } catch { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Ueberblick: $($_.Exception.Message)" | Add-Content $ErrFile -Encoding UTF8 }
    }

    Save-State
}
catch {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $($_.Exception.Message)" | Add-Content $ErrFile -Encoding UTF8
    throw
}
