# ==========================================================
#  Whatnot Stickerprinter
#  Print bij elke verkoop in je Whatnot-show automatisch
#  een sticker (koper + bestelnummer) op je DYMO.
# ==========================================================

Add-Type -AssemblyName System.Windows.Forms
try {
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$AppVersion = '2.2'
$ApiBase    = 'https://api.whatnot.com/seller-api/rest/v2'
$DataDir    = Join-Path $env:APPDATA 'WhatnotStickerprinter'
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }
$ConfigPath = Join-Path $DataDir 'config.json'
$SeenPath   = Join-Path $DataDir 'geprint.json'
$LogPath    = Join-Path $DataDir 'log.txt'
$BuyersPath = Join-Path $DataDir 'kopers.json'
$Euro       = [string][char]0x20AC

function Write-Log($msg) {
    try { Add-Content -Path $LogPath -Value ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg) -Encoding UTF8 } catch {}
}

# ---------- Hulpfunctie: JSON-lijst veilig inlezen ----------
function Read-JsonList($path) {
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path $path)) { return ,$out }
    $raw = Get-Content $path -Raw -Encoding UTF8
    if (-not $raw -or -not $raw.Trim()) { return ,$out }
    $parsed = $raw | ConvertFrom-Json
    $stack = New-Object System.Collections.Stack
    $stack.Push($parsed)
    while ($stack.Count -gt 0) {
        $x = $stack.Pop()
        if ($null -eq $x) { continue }
        if ($x -is [System.Array]) { foreach ($y in $x) { $stack.Push($y) } }
        else { [void]$out.Add($x) }
    }
    return ,$out
}

# ---------- Instellingen ----------
function Load-Config {
    $c = [ordered]@{ TokenEnc = ''; Printer = ''; IntervalSec = 5; OnlyShows = $true; LogoPath = ''; InstallTime = ''; PaperName = ''; OffXmm = 0; OffYmm = 0; ScalePct = 100; PerBuyer = $true; BreakWords = 'break, spot'; BreakRound = 1; UpdateUrl = ''; RequirePaid = $true }
    if (Test-Path $ConfigPath) {
        try {
            $j = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @('TokenEnc','Printer','IntervalSec','OnlyShows','LogoPath','InstallTime','PaperName','OffXmm','OffYmm','ScalePct','PerBuyer','BreakWords','BreakRound','UpdateUrl','RequirePaid')) {
                if ($null -ne $j.$k) { $c[$k] = $j.$k }
            }
        } catch { Write-Log "Instellingen onleesbaar: $_" }
    }
    return $c
}
function Save-Config {
    try { ($script:Config | ConvertTo-Json) | Set-Content -Path $ConfigPath -Encoding UTF8 } catch { Write-Log "Opslaan mislukt: $_" }
}
function Get-Token {
    if (-not $script:Config.TokenEnc) { return $null }
    try {
        $ss = $script:Config.TokenEnc | ConvertTo-SecureString
        $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
    } catch { return $null }
}
function Set-Token($plain) {
    # Versleuteld opgeslagen met je Windows-account (alleen leesbaar op deze pc, door jou)
    $ss = ConvertTo-SecureString $plain -AsPlainText -Force
    $script:Config.TokenEnc = ConvertFrom-SecureString $ss
    Save-Config
}

# ---------- Al geprinte orders onthouden ----------
$script:Seen = New-Object 'System.Collections.Generic.HashSet[string]'
$script:Pending = @{}   # order-id -> moment eerst gezien (nog niet betaald)
$script:FirstRun = -not (Test-Path $SeenPath)
if (-not $script:FirstRun) {
    try {
        foreach ($id in (Read-JsonList $SeenPath)) {
            # Oude, verkeerd opgeslagen regels bevatten meerdere ID's met spaties ertussen
            foreach ($part in ([string]$id -split '\s+')) { if ($part) { [void]$script:Seen.Add($part) } }
        }
    } catch { Write-Log "Geprint-bestand onleesbaar: $_" }
}
function Save-Seen {
    try { ConvertTo-Json -InputObject @($script:Seen) -Compress | Set-Content $SeenPath -Encoding UTF8 } catch {}
}

# ---------- Kopers per show (voor extra stickers) ----------
$script:Buyers = @{}
$script:CurrentShow = ''
if (Test-Path $BuyersPath) {
    try {
        $latest = ''
        foreach ($b in (Read-JsonList $BuyersPath)) {
            if (-not $b.Key -or $b.Key -is [System.Array] -or $b.Aankopen -is [System.Array]) { continue }
            $script:Buyers[[string]$b.Key] = $b
            if ([string]$b.LastTime -gt $latest) { $latest = [string]$b.LastTime; $script:CurrentShow = [string]$b.ShowId }
        }
    } catch { Write-Log "Kopers-bestand onleesbaar: $_" }
}
function Save-Buyers {
    try {
        $cut = (Get-Date).AddDays(-14).ToString('o')
        $keep = @($script:Buyers.Values | Where-Object { [string]$_.LastTime -ge $cut })
        ConvertTo-Json -InputObject $keep -Depth 3 | Set-Content $BuyersPath -Encoding UTF8
    } catch { Write-Log "Kopers opslaan mislukt: $_" }
}
function Update-Buyer($o, $data, [bool]$printed) {
    if (-not ($o.show_id -and $o.buyer -and $o.buyer.username)) { return }
    $key = ("B|{0}|{1}" -f $o.show_id, $o.buyer.username).ToLower()
    $b = $script:Buyers[$key]
    if (-not $b) {
        $b = [pscustomobject]@{ Key = $key; ShowId = [string]$o.show_id; Buyer = $data.Buyer; Aankopen = 0; Stickers = 0; EersteOrder = $data.OrderId; LastTime = '' }
        $script:Buyers[$key] = $b
    }
    $b.Aankopen = [int]$b.Aankopen + 1
    if ($printed) { $b.Stickers = [int]$b.Stickers + 1 }
    $b.LastTime = (Get-Date).ToString('o')
    $script:CurrentShow = $b.ShowId
}

# ---------- Whatnot API ----------
function Invoke-Whatnot($path) {
    $token = Get-Token
    if (-not $token) { throw 'GEEN_SLEUTEL' }
    $resp = Invoke-WebRequest -Uri ($ApiBase + $path) -Headers @{ Authorization = "Bearer $token" } -UseBasicParsing -TimeoutSec 15
    $text = [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    return ($text | ConvertFrom-Json)
}
function Get-HttpCode($err) {
    try { return [int]$err.Exception.Response.StatusCode } catch { return 0 }
}
function Get-UtcTime($v) {
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [datetime]) { return $v.ToUniversalTime() }
    try { return [DateTime]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
    catch { return $null }
}
function Get-InstallTime { return (Get-UtcTime $script:Config.InstallTime) }

function Format-Money($m) {
    if (-not $m) { return '' }
    $amt = 0.0
    [void][double]::TryParse([string]$m.amount, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$amt)
    $sym = switch ([string]$m.currency) { 'EUR' { $Euro } 'USD' { '$' } 'GBP' { [string][char]0x00A3 } default { "$($m.currency)" } }
    return $sym + ' ' + $amt.ToString('N2', [Globalization.CultureInfo]'nl-NL')
}
function Test-IsBreak($titles) {
    $words = @(([string]$script:Config.BreakWords).Split(',') | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ })
    if ($words.Count -eq 0) { return $false }
    $all = (@($titles) -join ' ').ToLower()
    foreach ($w in $words) { if ($all.Contains($w)) { return $true } }
    return $false
}

function Get-BreakId($items) {
    # Welke break is dit? Eerst via het product-ID, anders via de titel zonder spotnummer
    foreach ($it in @($items)) {
        $pi = $it.purchased_item
        foreach ($v in @($pi.product_id, $pi.product.id, $it.product_id)) {
            if ($v) { return 'P' + [string]$v }
        }
    }
    foreach ($it in @($items)) {
        $t = ([string]$it.purchased_item.title).ToLower()
        if (-not $t) { continue }
        $t = [regex]::Replace($t, '(spot|plek|slot)\s*#?\s*\d+', '')
        $t = ($t -split '\s[-\u2013|:]\s')[0]
        $t = [regex]::Replace($t, '\s+', ' ').Trim()
        if ($t) { return 'T' + $t }
    }
    return 'onbekend'
}
$script:BreakLogCount = 0

function New-LabelData($order) {
    $items = @()
    try { $items = @((Invoke-Whatnot "/orders/$($order.id)/items?limit=50&reverse=false").data) }
    catch { Write-Log "Producten ophalen mislukt voor order $($order.id): $_" }
    $titles = @()
    foreach ($it in $items) {
        $t = [string]$it.purchased_item.title
        if (-not $t) { continue }
        if ($it.quantity -gt 1) { $t = "$($it.quantity)x $t" }
        $titles += $t
    }
    if ($titles.Count -eq 0)     { $product = '(product onbekend)' }
    elseif ($titles.Count -eq 1) { $product = $titles[0] }
    else                         { $product = "$($titles[0])  (+$($titles.Count - 1) meer)" }

    if ($script:BreakLogCount -lt 3 -and $items.Count -gt 0 -and (Test-IsBreak $titles)) {
        $script:BreakLogCount++
        Write-Log ("Break-order {0} producten: {1}" -f $order.id, (ConvertTo-Json -InputObject @($items) -Depth 5 -Compress))
    }
    if ($order.buyer -and $order.buyer.username) { $buyer = '@' + $order.buyer.username }
    else { $buyer = '(koper onbekend)'; $script:BuyerMissing = $true }

    [pscustomobject]@{
        OrderId = [string]$order.id
        Buyer   = $buyer
        Product = $product
        Price   = (Format-Money $order.amounts.subtotal)
        Time    = (Get-Date)
        IsBreak = (Test-IsBreak $titles)
        BreakId = (Get-BreakId $items)
    }
}

# ---------- Printen ----------
$script:LogoImg = $null
function Load-Logo {
    if ($script:LogoImg) { $script:LogoImg.Dispose(); $script:LogoImg = $null }
    $lp = [string]$script:Config.LogoPath
    if (-not $lp -or -not (Test-Path $lp)) { return }
    try {
        # Kopie in het geheugen, zodat het bestand zelf niet op slot gaat
        $ms = New-Object System.IO.MemoryStream(,[System.IO.File]::ReadAllBytes($lp))
        $tmp = [System.Drawing.Image]::FromStream($ms)
        $script:LogoImg = New-Object System.Drawing.Bitmap($tmp)
        $tmp.Dispose(); $ms.Dispose()
    } catch { Write-Log "Logo laden mislukt: $_"; $script:LogoImg = $null }
}

function Get-FitFont($g, $text, $maxW, $startSize, $style) {
    $size = [double]$startSize
    while ($true) {
        $f = New-Object System.Drawing.Font('Segoe UI', [single]$size, $style)
        if ($g.MeasureString($text, $f).Width -le $maxW -or $size -le 7) { return $f }
        $f.Dispose(); $size -= 1
    }
}

function Draw-Label($ev, $d) {
    $g = $ev.Graphics
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $pb = $ev.PageBounds
    $hx = $ev.PageSettings.HardMarginX; $hy = $ev.PageSettings.HardMarginY
    try { $vc = $g.VisibleClipBounds } catch { $vc = $null }
    $script:LastDiag = ("Printer: {0} | Papier: {1} ({2:N0} x {3:N0} mm) | Liggend: {4} | Pagina: {5}x{6} | Harde marge: {7}/{8} | Zichtbaar: {9}" -f
        $ev.PageSettings.PrinterSettings.PrinterName, $ev.PageSettings.PaperSize.PaperName,
        ($ev.PageSettings.PaperSize.Width * 0.254), ($ev.PageSettings.PaperSize.Height * 0.254),
        $ev.PageSettings.Landscape, $pb.Width, $pb.Height, $hx, $hy, $vc)
    Write-Log $script:LastDiag

    # Kalibratie: verschuiven (mm) en schalen (%), plus correctie voor de harde printermarge
    $mm = 100 / 25.4
    $g.TranslateTransform([single](-$hx + [double]$script:Config.OffXmm * $mm), [single](-$hy + [double]$script:Config.OffYmm * $mm))
    $sc = [Math]::Max(0.3, [double]$script:Config.ScalePct / 100)
    $g.ScaleTransform([single]$sc, [single]$sc)

    # Altijd tekenen op het echte 11352-formaat (54 x 25 mm), los van wat de driver doorgeeft
    $b = New-Object System.Drawing.Rectangle(0, 0, 213, 98)
    $padX = 10; $padY = 5
    $x = $b.X + $padX; $y = $b.Y + $padY
    $w = $b.Width - 2 * $padX
    $bottom = $b.Y + $b.Height - $padY
    $black = [System.Drawing.Brushes]::Black

    # Logo links, vierkant vak over de volle hoogte
    if ($script:LogoImg) {
        $box = $b.Height - 2 * $padY
        $img = $script:LogoImg
        $scale = [Math]::Min($box / $img.Width, $box / $img.Height)
        $lw = $img.Width * $scale; $lh = $img.Height * $scale
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($img, [single]$x, [single]($b.Y + $padY + ($box - $lh) / 2), [single]$lw, [single]$lh)
        $x += $lw + 6
        $w -= $lw + 6
    }

    # Koper bovenaan: zo groot mogelijk, maar altijd op 1 regel
    $fUser = Get-FitFont $g $d.Buyer $w 16 ([System.Drawing.FontStyle]::Bold)
    $g.DrawString($d.Buyer, $fUser, $black, [single]$x, [single]$y)

    # Bestelnummer onderaan
    $nr = '#' + $d.OrderId
    $fNr = Get-FitFont $g $nr $w 14 ([System.Drawing.FontStyle]::Bold)
    $nrH = $g.MeasureString($nr, $fNr).Height
    $g.DrawString($nr, $fNr, $black, [single]$x, [single]($bottom - $nrH))
    $fSmall = New-Object System.Drawing.Font('Segoe UI', [single]7)
    $smallH = $g.MeasureString('Bestelnummer', $fSmall).Height
    $g.DrawString('Bestelnummer', $fSmall, $black, [single]($x + 2), [single]($bottom - $nrH - $smallH + 4))

    $fUser.Dispose(); $fNr.Dispose(); $fSmall.Dispose()
    $ev.HasMorePages = $false
}

function Find-PaperSize($pset) {
    $sizes = @($pset.PaperSizes)
    if ($script:Config.PaperName) {
        foreach ($p in $sizes) { if ($p.PaperName -eq $script:Config.PaperName) { return $p } }
    }
    foreach ($p in $sizes) { if ($p.PaperName -match '11352') { return $p } }
    # Op maat zoeken: 25 x 54 mm (98 x 213 honderdste inch), beide kanten op
    foreach ($p in $sizes) {
        $a = [Math]::Min($p.Width, $p.Height); $c = [Math]::Max($p.Width, $p.Height)
        if ([Math]::Abs($a - 98) -le 8 -and [Math]::Abs($c - 213) -le 12) { return $p }
    }
    return $null
}

function Print-Label($data) {
    $doc = New-Object System.Drawing.Printing.PrintDocument
    try {
        $doc.DocumentName = "Whatnot sticker $($data.OrderId)"
        if ($script:Config.Printer) { $doc.PrinterSettings.PrinterName = $script:Config.Printer }
        if (-not $doc.PrinterSettings.IsValid) { throw "Printer '$($script:Config.Printer)' niet gevonden" }
        $ps = Find-PaperSize $doc.PrinterSettings
        if (-not $ps) { $ps = New-Object System.Drawing.Printing.PaperSize('11352', 98, 213); Write-Log 'Geen 11352-papierformaat gevonden bij deze printer, eigen formaat gebruikt' }
        $doc.DefaultPageSettings.PaperSize = $ps
        $doc.DefaultPageSettings.Landscape = ($ps.Width -lt $ps.Height)
        $doc.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(0, 0, 0, 0)
        $doc.OriginAtMargins = $false
        $doc.PrintController = New-Object System.Drawing.Printing.StandardPrintController
        $script:CurrentLabel = $data
        $doc.add_PrintPage({ param($s, $ev) Draw-Label $ev $script:CurrentLabel })
        $doc.Print()
    } finally { $doc.Dispose() }
}

# ---------- Automatisch updaten via GitHub ----------
$script:RestartAfterClose = $false
function Get-RawBase($url) {
    # Accepteert https://github.com/naam/repo of https://raw.githubusercontent.com/naam/repo/main
    $u = ([string]$url).Trim().TrimEnd('/')
    if (-not $u) { return '' }
    if ($u -match '^https://github\.com/([^/]+)/([^/]+?)(\.git)?(/.*)?$') { return "https://raw.githubusercontent.com/$($Matches[1])/$($Matches[2])/main" }
    return $u
}
function Get-WebBytes($url) {
    $sep = '?'; if ($url.Contains('?')) { $sep = '&' }
    $r = Invoke-WebRequest -Uri ($url + $sep + 't=' + [DateTime]::UtcNow.Ticks) -UseBasicParsing -TimeoutSec 10 -Headers @{ 'Cache-Control' = 'no-cache' }
    return $r.RawContentStream.ToArray()
}
function Check-Update([bool]$manual) {
    $base = Get-RawBase $script:Config.UpdateUrl
    if (-not $base) {
        if ($manual) { [System.Windows.Forms.MessageBox]::Show('Er is nog geen update-adres ingesteld (Instellingen).', 'Updates') | Out-Null }
        return
    }
    try {
        $remote = ([Text.Encoding]::UTF8.GetString((Get-WebBytes "$base/version.txt"))).Trim([char]0xFEFF, ' ', "`r", "`n", "`t")
        $rv = [version]$remote; $lv = [version]$AppVersion
    } catch {
        Write-Log "Update-check mislukt: $_"
        if ($manual) { [System.Windows.Forms.MessageBox]::Show("Kon niet controleren op updates:`n$($_.Exception.Message)", 'Updates') | Out-Null }
        return
    }
    if ($rv -le $lv) {
        if ($manual) { [System.Windows.Forms.MessageBox]::Show("Je hebt de nieuwste versie (v$AppVersion).", 'Updates') | Out-Null }
        return
    }
    $ans = [System.Windows.Forms.MessageBox]::Show("Er is een nieuwe versie beschikbaar: v$remote (je hebt nu v$AppVersion).`n`nNu bijwerken? Het programma start daarna vanzelf opnieuw.", 'Update beschikbaar', 'YesNo', 'Question')
    if ($ans -ne 'Yes') { return }
    try {
        $bytes = Get-WebBytes "$base/WhatnotStickerprinter.ps1"
        $text = [Text.Encoding]::UTF8.GetString($bytes)
        # Veiligheidscontroles: is dit echt de stickerprinter, met het juiste versienummer, zonder fouten?
        if (-not $text.Contains('Whatnot Stickerprinter')) { throw 'Het gedownloade bestand is niet de stickerprinter.' }
        if ($text -notmatch ("\`$AppVersion = '" + [regex]::Escape($remote) + "'")) { throw "Versienummer in het bestand klopt niet met version.txt ($remote)." }
        $errs = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$errs)
        if ($errs -and $errs.Count -gt 0) { throw "Het nieuwe bestand bevat fouten: $($errs[0].Message)" }

        $me = Join-Path $PSScriptRoot 'WhatnotStickerprinter.ps1'
        Copy-Item $me (Join-Path $PSScriptRoot 'WhatnotStickerprinter.vorige-versie.ps1') -Force
        [System.IO.File]::WriteAllBytes($me, $bytes)
        foreach ($extra in @('LEES MIJ EERST - Uitleg.txt')) {
            try { [System.IO.File]::WriteAllBytes((Join-Path $PSScriptRoot $extra), (Get-WebBytes ("$base/" + [uri]::EscapeDataString($extra)))) } catch { Write-Log "Extra bestand $extra niet bijgewerkt: $_" }
        }
        Write-Log "Bijgewerkt van v$AppVersion naar v$remote"
        $script:RestartAfterClose = $true
        $form.Close()
    } catch {
        Write-Log "Update mislukt: $_"
        [System.Windows.Forms.MessageBox]::Show("Bijwerken is niet gelukt, de huidige versie blijft gewoon werken.`n`n$($_.Exception.Message)", 'Update') | Out-Null
    }
}

# ---------- Venster ----------
$script:Config = Load-Config
if (-not $script:Config.UpdateUrl) { $script:Config.UpdateUrl = 'https://github.com/Carthawk/Stickerprinter' }
if (-not $script:Config.InstallTime -and -not $script:FirstRun) { $script:Config.InstallTime = [DateTime]::UtcNow.ToString('o'); Save-Config }
$script:Running = $false
$script:Busy = $false
$script:BuyerMissing = $false

$form = New-Object System.Windows.Forms.Form
$form.Text = "Whatnot Stickerprinter  v$AppVersion"
$form.Size = New-Object System.Drawing.Size(910, 560)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$form.MinimumSize = New-Object System.Drawing.Size(640, 400)

$status = New-Object System.Windows.Forms.Label
$status.Location = New-Object System.Drawing.Point(12, 10)
$status.Size = New-Object System.Drawing.Size(870, 28)
$status.Anchor = 'Top,Left,Right'
$status.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($status)

function Set-Status($text, $kind) {
    $status.Text = $text
    $status.ForeColor = switch ($kind) { 'ok' { [System.Drawing.Color]::ForestGreen } 'fout' { [System.Drawing.Color]::Firebrick } default { [System.Drawing.Color]::DimGray } }
}

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(12, 46); $btnStart.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnStart)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Testlabel'
$btnTest.Location = New-Object System.Drawing.Point(130, 46); $btnTest.Size = New-Object System.Drawing.Size(95, 30)
$form.Controls.Add($btnTest)

$lblPrinter = New-Object System.Windows.Forms.Label
$lblPrinter.Text = 'Printer:'
$btnCal = New-Object System.Windows.Forms.Button
$btnCal.Text = 'Instellingen'
$btnCal.Location = New-Object System.Drawing.Point(233, 46); $btnCal.Size = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($btnCal)

$lblPrinter.Location = New-Object System.Drawing.Point(374, 52); $lblPrinter.AutoSize = $true
$form.Controls.Add($lblPrinter)

$cmbPrinter = New-Object System.Windows.Forms.ComboBox
$cmbPrinter.DropDownStyle = 'DropDownList'
$cmbPrinter.Location = New-Object System.Drawing.Point(428, 49); $cmbPrinter.Size = New-Object System.Drawing.Size(230, 28)
foreach ($p in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) { [void]$cmbPrinter.Items.Add($p) }
$form.Controls.Add($cmbPrinter)

$chkShows = New-Object System.Windows.Forms.CheckBox
$chkShows.Text = 'Alleen verkopen uit live shows'
$chkShows.Location = New-Object System.Drawing.Point(670, 42); $chkShows.AutoSize = $true
$chkShows.Checked = [bool]$script:Config.OnlyShows
$form.Controls.Add($chkShows)

$chkBuyer = New-Object System.Windows.Forms.CheckBox
$chkBuyer.Text = 'Breaks: 1 sticker per koper per break'
$chkBuyer.Location = New-Object System.Drawing.Point(670, 64); $chkBuyer.AutoSize = $true
$chkBuyer.Checked = [bool]$script:Config.PerBuyer
$form.Controls.Add($chkBuyer)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 92)
$tabs.Size = New-Object System.Drawing.Size(870, 374)
$tabs.Anchor = 'Top,Bottom,Left,Right'
$tabSales = New-Object System.Windows.Forms.TabPage; $tabSales.Text = 'Verkopen'
$tabBuyers = New-Object System.Windows.Forms.TabPage; $tabBuyers.Text = 'Kopers deze show'
[void]$tabs.TabPages.Add($tabSales); [void]$tabs.TabPages.Add($tabBuyers)
$form.Controls.Add($tabs)

$list = New-Object System.Windows.Forms.ListView
$list.View = 'Details'; $list.FullRowSelect = $true; $list.GridLines = $true
$list.Dock = 'Fill'
[void]$list.Columns.Add('Tijd', 70)
[void]$list.Columns.Add('Koper', 160)
[void]$list.Columns.Add('Product', 360)
[void]$list.Columns.Add('Bestelnr.', 110)
[void]$list.Columns.Add('Status', 100)
$tabSales.Controls.Add($list)

$listB = New-Object System.Windows.Forms.ListView
$listB.View = 'Details'; $listB.FullRowSelect = $true; $listB.GridLines = $true; $listB.MultiSelect = $false
$listB.Dock = 'Fill'
$listB.Font = New-Object System.Drawing.Font('Segoe UI', 11)
[void]$listB.Columns.Add('Koper', 260)
[void]$listB.Columns.Add('Aankopen', 100)
[void]$listB.Columns.Add('Stickers geprint', 140)
[void]$listB.Columns.Add('Bestelnr. (eerste)', 160)
[void]$listB.Columns.Add('Laatste aankoop', 130)
$pnlB = New-Object System.Windows.Forms.Panel
$pnlB.Dock = 'Top'; $pnlB.Height = 44
$btnExtra = New-Object System.Windows.Forms.Button
$btnExtra.Text = 'Extra sticker printen'
$btnExtra.Location = New-Object System.Drawing.Point(6, 7); $btnExtra.Size = New-Object System.Drawing.Size(170, 30)
$pnlB.Controls.Add($btnExtra)
$lblB = New-Object System.Windows.Forms.Label
$lblB.Text = 'Meeste aankopen bovenaan. Klik een koper aan en druk op de knop, of dubbelklik op de koper.'
$lblB.Location = New-Object System.Drawing.Point(186, 13); $lblB.AutoSize = $true
$pnlB.Controls.Add($lblB)
$tabBuyers.Controls.Add($listB)
$tabBuyers.Controls.Add($pnlB)

function Refresh-Buyers {
    try { Refresh-BuyersInner } catch { Write-Log "Kopers-lijst bijwerken mislukt: $_" }
}
function Refresh-BuyersInner {
    $selKey = $null
    if ($listB.SelectedItems.Count -gt 0) { $selKey = $listB.SelectedItems[0].Tag.Key }
    $listB.BeginUpdate()
    $listB.Items.Clear()
    $rows = @($script:Buyers.Values | Where-Object { $_.ShowId -eq $script:CurrentShow } |
        Sort-Object -Property @{ Expression = { [int]$_.Aankopen }; Descending = $true }, @{ Expression = { [string]$_.LastTime }; Descending = $true })
    foreach ($b in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem([string]$b.Buyer)
        [void]$it.SubItems.Add([string]$b.Aankopen)
        [void]$it.SubItems.Add([string]$b.Stickers)
        [void]$it.SubItems.Add([string]$b.EersteOrder)
        $t = ''; try { $t = ([DateTime]::Parse([string]$b.LastTime)).ToString('HH:mm:ss') } catch {}
        [void]$it.SubItems.Add($t)
        $it.Tag = $b
        if ([int]$b.Aankopen -gt [int]$b.Stickers) { $it.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold) }
        [void]$listB.Items.Add($it)
        if ($selKey -and $b.Key -eq $selKey) { $it.Selected = $true }
    }
    $listB.EndUpdate()
    $tabBuyers.Text = "Kopers deze show ($($rows.Count))"
}

function Print-Extra {
    if ($listB.SelectedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Klik eerst een koper aan.', 'Extra sticker') | Out-Null; return }
    $b = $listB.SelectedItems[0].Tag
    $d = [pscustomobject]@{ OrderId = [string]$b.EersteOrder; Buyer = [string]$b.Buyer; Product = '(extra sticker)'; Price = ''; Time = (Get-Date) }
    try {
        Print-Label $d
        $b.Stickers = [int]$b.Stickers + 1
        Save-Buyers; Refresh-Buyers
        Add-Row $d 'Extra sticker'
    } catch { [System.Windows.Forms.MessageBox]::Show("Printen mislukt: $($_.Exception.Message)", 'Fout') | Out-Null }
}
$btnExtra.add_Click({ Print-Extra })
$listB.add_DoubleClick({ Print-Extra })

$btnReprint = New-Object System.Windows.Forms.Button
$btnReprint.Text = 'Herprint geselecteerde'
$btnReprint.Location = New-Object System.Drawing.Point(12, 475); $btnReprint.Size = New-Object System.Drawing.Size(170, 30)
$btnReprint.Anchor = 'Bottom,Left'
$form.Controls.Add($btnReprint)

$btnToken = New-Object System.Windows.Forms.Button
$btnToken.Text = 'API-sleutel wijzigen'
$btnToken.Location = New-Object System.Drawing.Point(190, 475); $btnToken.Size = New-Object System.Drawing.Size(150, 30)
$btnToken.Anchor = 'Bottom,Left'
$form.Controls.Add($btnToken)

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = 'Logbestand'
$btnLog.Location = New-Object System.Drawing.Point(348, 475); $btnLog.Size = New-Object System.Drawing.Size(100, 30)
$btnLog.Anchor = 'Bottom,Left'
$form.Controls.Add($btnLog)

$btnLogo = New-Object System.Windows.Forms.Button
$btnLogo.Text = 'Logo kiezen'
$btnLogo.Location = New-Object System.Drawing.Point(456, 475); $btnLogo.Size = New-Object System.Drawing.Size(105, 30)
$btnLogo.Anchor = 'Bottom,Left'
$form.Controls.Add($btnLogo)

$btnNoLogo = New-Object System.Windows.Forms.Button
$btnNoLogo.Text = 'Logo weghalen'
$btnNoLogo.Location = New-Object System.Drawing.Point(569, 475); $btnNoLogo.Size = New-Object System.Drawing.Size(115, 30)
$btnNoLogo.Anchor = 'Bottom,Left'
$form.Controls.Add($btnNoLogo)

$btnHelp = New-Object System.Windows.Forms.Button
$btnHelp.Text = 'Uitleg'
$btnHelp.Location = New-Object System.Drawing.Point(692, 475); $btnHelp.Size = New-Object System.Drawing.Size(90, 30)
$btnNewBreak = New-Object System.Windows.Forms.Button
$btnNewBreak.Text = 'Nieuwe break'
$btnNewBreak.Location = New-Object System.Drawing.Point(788, 475); $btnNewBreak.Size = New-Object System.Drawing.Size(94, 30)
$btnNewBreak.Anchor = 'Bottom,Left'
$btnNewBreak.add_Click({
    $script:Config.BreakRound = [int]$script:Config.BreakRound + 1; Save-Config
    Set-Status 'Nieuwe break gestart: elke koper krijgt weer 1 sticker.' 'ok'
})
$form.Controls.Add($btnNewBreak)
$btnHelp.Anchor = 'Bottom,Left'
$form.Controls.Add($btnHelp)

function Add-Row($d, $st) {
    $it = New-Object System.Windows.Forms.ListViewItem($d.Time.ToString('HH:mm:ss'))
    [void]$it.SubItems.Add($d.Buyer)
    [void]$it.SubItems.Add($d.Product)
    [void]$it.SubItems.Add($d.OrderId)
    [void]$it.SubItems.Add($st)
    $it.Tag = $d
    if ($st -eq 'FOUT') { $it.ForeColor = [System.Drawing.Color]::Firebrick }
    if ($st -eq 'Zelfde koper') { $it.ForeColor = [System.Drawing.Color]::Gray }
    $list.Items.Insert(0, $it) | Out-Null
}

$script:WaitShown = New-Object 'System.Collections.Generic.HashSet[string]'
function Add-WaitRow($o, $reason) {
    $key = "$($o.id)|$reason|$($o.status)"
    if (-not $script:WaitShown.Add($key)) { return }
    $nm = '(nog geen naam)'
    if ($o.buyer -and $o.buyer.username) { $nm = '@' + $o.buyer.username }
    Write-Log ("Wacht: order {0} status={1} naam={2} show={3} reden={4}" -f $o.id, $o.status, $nm, $o.show_id, $reason)
    $it = New-Object System.Windows.Forms.ListViewItem((Get-Date).ToString('HH:mm:ss'))
    [void]$it.SubItems.Add($nm)
    [void]$it.SubItems.Add("(wacht: $reason)")
    [void]$it.SubItems.Add([string]$o.id)
    [void]$it.SubItems.Add("Wacht ($($o.status))")
    $it.ForeColor = [System.Drawing.Color]::DarkOrange
    $list.Items.Insert(0, $it) | Out-Null
}

function Update-StartButton {
    if ($script:Running) { $btnStart.Text = 'Pauze' } else { $btnStart.Text = 'Start' }
}

# Printer kiezen: opgeslagen keuze, anders eerste DYMO
if ($script:Config.Printer -and $cmbPrinter.Items.Contains($script:Config.Printer)) {
    $cmbPrinter.SelectedItem = $script:Config.Printer
} else {
    foreach ($p in $cmbPrinter.Items) { if ($p -match 'DYMO') { $cmbPrinter.SelectedItem = $p; break } }
    if ($cmbPrinter.SelectedItem) { $script:Config.Printer = [string]$cmbPrinter.SelectedItem; Save-Config }
}

function Ask-Token {
    $t = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Plak hier je Whatnot API-sleutel.`n`nAanmaken: whatnot.com/dashboard/developer`nVink aan: Bestellingen lezen en Klanten lezen.",
        'Whatnot API-sleutel', '')
    $t = $t.Trim()
    if (-not $t) { return $false }
    Set-Token $t
    try {
        [void](Invoke-Whatnot '/orders?limit=1')
        [System.Windows.Forms.MessageBox]::Show('Sleutel werkt! De stickerprinter staat aan.', 'Gelukt') | Out-Null
        return $true
    } catch {
        $code = Get-HttpCode $_
        if ($code -eq 401 -or $code -eq 403) {
            [System.Windows.Forms.MessageBox]::Show("Whatnot accepteert deze sleutel niet (code $code). Controleer of hij klopt en of Bestellingen lezen en Klanten lezen zijn aangevinkt.", 'Sleutel werkt niet') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Kon Whatnot niet bereiken: $($_.Exception.Message)", 'Geen verbinding') | Out-Null
        }
        return $false
    }
}

# ---------- Controleren op nieuwe verkopen ----------
function Poll {
    if ($script:Busy -or -not $script:Running) { return }
    $script:Busy = $true
    try {
        $orders = @((Invoke-Whatnot '/orders?limit=25').data)

        if ($script:FirstRun) {
            # Eerste keer ooit: alle bestaande orders overslaan, alleen nieuwe printen
            foreach ($o in $orders) { [void]$script:Seen.Add([string]$o.id) }
            Save-Seen
            $script:Config.InstallTime = [DateTime]::UtcNow.ToString('o'); Save-Config
            $script:FirstRun = $false
            Set-Status 'Verbonden. Wacht op nieuwe verkopen...' 'ok'
            return
        }

        # Orders die nog op betaling wachten opnieuw ophalen als ze niet meer in de laatste 25 staan
        $byId = @{}
        foreach ($o in $orders) { $byId[[string]$o.id] = $o }
        foreach ($pendId in @($script:Pending.Keys)) {
            if ($byId.ContainsKey($pendId)) { continue }
            if ($script:Pending[$pendId] -lt (Get-Date).AddHours(-24)) { $script:Pending.Remove($pendId); [void]$script:Seen.Add($pendId); continue }
            try {
                $r = Invoke-Whatnot "/orders/$pendId"
                $oo = if ($r.data) { $r.data } else { $r }
                if ($oo.id) { $byId[[string]$oo.id] = $oo }
            } catch { Write-Log "Order $pendId opnieuw ophalen mislukt: $_" }
        }

        $new = @()
        $paidCount = 0; $namedCount = 0
        foreach ($o in $byId.Values) {
            if (@('pending_fulfillment','completed') -contains $o.status) { $paidCount++ }
            if ($o.buyer -and $o.buyer.username) { $namedCount++ }
        }
        # Geen enkele order heeft een naam? Dan mist de sleutel waarschijnlijk 'Klanten lezen'
        $script:BuyerMissing = ($paidCount -ge 2 -and $namedCount -eq 0)
        foreach ($o in $byId.Values) {
            $id = [string]$o.id
            if ($script:Seen.Contains($id)) { continue }
            if (@('cancelled','failed') -contains $o.status) {
                # Geannuleerd of betaling mislukt: nooit printen
                if ($script:Pending.ContainsKey($id)) { $script:Pending.Remove($id) }
                [void]$script:Seen.Add($id); continue
            }
            if ($script:Config.RequirePaid) {
                if (@('pending_fulfillment','completed') -notcontains $o.status) {
                    # Nog niet betaald: onthouden en later opnieuw bekijken
                    if (-not $script:Pending.ContainsKey($id)) { $script:Pending[$id] = Get-Date }
                    Add-WaitRow $o 'nog niet betaald'
                    continue
                }
                if (-not ($o.buyer -and $o.buyer.username)) {
                    # Pas printen als de order in Whatnot volledig is, met de naam van de koper erbij
                    if (-not $script:Pending.ContainsKey($id)) { $script:Pending[$id] = Get-Date }
                    Add-WaitRow $o 'nog geen naam'
                    continue
                }
            }
            if ($script:Pending.ContainsKey($id)) { $script:Pending.Remove($id) }
            $new += $o
        }
        if ($chkShows.Checked) { $new = @($new | Where-Object { $_.show_id }) }
        # Extra slot: nooit iets printen dat besteld is voor de installatie
        $cutoff = Get-InstallTime
        if ($cutoff) {
            $cutoff = $cutoff.AddMinutes(-2)
            $tooOld = @($new | Where-Object { $t = Get-UtcTime $_.created_at; $t -and $t -lt $cutoff })
            foreach ($o in $tooOld) { [void]$script:Seen.Add([string]$o.id) }
            $new = @($new | Where-Object { $t = Get-UtcTime $_.created_at; -not ($t -and $t -lt $cutoff) })
        }
        # Oudste eerst, zodat de stickers op volgorde komen
        $new = @($new | Sort-Object -Property @{ Expression = { $t = Get-UtcTime $_.created_at; if ($t) { $t } else { [DateTime]::MaxValue } } })

        $printError = $false
        foreach ($o in $new) {
            $data = New-LabelData $o
            # Zelfde koper in dezelfde show (bijv. meerdere spots in een break): geen extra sticker
            $bkey = $null
            if ($chkBuyer.Checked -and $data.IsBreak -and $o.show_id -and $o.buyer -and $o.buyer.username) { $bkey = ("B|{0}|{1}|{2}|{3}" -f $o.show_id, $script:Config.BreakRound, $data.BreakId, $o.buyer.username).ToLower() }
            if ($bkey -and $script:Seen.Contains($bkey)) {
                Add-Row $data 'Zelfde koper'
                Update-Buyer $o $data $false
                [void]$script:Seen.Add([string]$o.id)
                continue
            }
            if ($bkey) { [void]$script:Seen.Add($bkey) }
            try {
                Print-Label $data
                if ($data.IsBreak) { Add-Row $data 'Geprint (break)' } else { Add-Row $data 'Geprint' }
                Update-Buyer $o $data $true
            } catch {
                Add-Row $data 'FOUT'
                Update-Buyer $o $data $false
                Write-Log "Printen mislukt voor order $($o.id): $_"
                Set-Status "Printen mislukt: $($_.Exception.Message)" 'fout'
                $printError = $true
            }
            [void]$script:Seen.Add([string]$o.id)
        }
        if ($new.Count -gt 0) { Save-Seen; Save-Buyers; Refresh-Buyers }

        if ($script:BuyerMissing) {
            Set-Status 'Koper ontbreekt: maak een sleutel met ook Klanten lezen aangevinkt.' 'fout'
        } elseif (-not $printError) {
            $wait = ''
            if ($script:Pending.Count -gt 0) { $wait = " $($script:Pending.Count) order(s) wachten op betaling/naam." }
            Set-Status ("Actief - laatste check {0}. Wacht op verkopen...{1}" -f (Get-Date -Format 'HH:mm:ss'), $wait) 'ok'
        }
    } catch {
        $code = Get-HttpCode $_
        if ($_.Exception.Message -eq 'GEEN_SLEUTEL') {
            Set-Status 'Geen API-sleutel ingesteld. Klik op "API-sleutel wijzigen".' 'fout'
            $script:Running = $false; Update-StartButton
        } elseif ($code -eq 401 -or $code -eq 403) {
            Set-Status "Whatnot weigert de sleutel (code $code). Controleer sleutel en rechten." 'fout'
            $script:Running = $false; Update-StartButton
        } else {
            Set-Status ("Geen verbinding met Whatnot ({0}), ik probeer het opnieuw..." -f (Get-Date -Format 'HH:mm:ss')) 'fout'
            Write-Log "Poll fout: $_"
        }
    } finally { $script:Busy = $false }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]([Math]::Max(3, [int]$script:Config.IntervalSec) * 1000)
$timer.add_Tick({ Poll })

# ---------- Knoppen ----------
$btnStart.add_Click({
    $script:Running = -not $script:Running
    Update-StartButton
    if ($script:Running) { Set-Status 'Bezig met verbinden...' 'grijs'; Poll } else { Set-Status 'Gepauzeerd. Er wordt niets geprint.' 'grijs' }
})
$btnTest.add_Click({
    $d = [pscustomobject]@{ OrderId = '1048576213'; Buyer = '@testkoper'; Product = 'Charizard ex 199/165 - Pokemon 151 Special Illustration Rare'; Price = "$Euro 25,00"; Time = (Get-Date) }
    try { Print-Label $d; Add-Row $d 'Test' }
    catch { [System.Windows.Forms.MessageBox]::Show("Testlabel mislukt: $($_.Exception.Message)", 'Fout') | Out-Null }
})
$cmbPrinter.add_SelectedIndexChanged({ $script:Config.Printer = [string]$cmbPrinter.SelectedItem; Save-Config })
$chkShows.add_CheckedChanged({ $script:Config.OnlyShows = $chkShows.Checked; Save-Config })
$chkBuyer.add_CheckedChanged({ $script:Config.PerBuyer = $chkBuyer.Checked; Save-Config })
$btnReprint.add_Click({
    if ($list.SelectedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Selecteer eerst een of meer regels.', 'Herprinten') | Out-Null; return }
    foreach ($it in @($list.SelectedItems)) {
        try { Print-Label $it.Tag; $it.SubItems[4].Text = 'Herprint'; $it.ForeColor = [System.Drawing.Color]::Black }
        catch { [System.Windows.Forms.MessageBox]::Show("Herprinten mislukt: $($_.Exception.Message)", 'Fout') | Out-Null; break }
    }
})
$btnToken.add_Click({
    if (Ask-Token) { $script:BuyerMissing = $false; $script:Running = $true; Update-StartButton; Poll }
})
$btnLog.add_Click({ if (-not (Test-Path $LogPath)) { Write-Log 'Logbestand aangemaakt' }; Start-Process notepad.exe $LogPath })
$btnLogo.add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Kies je logo'
    $dlg.Filter = 'Afbeeldingen (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp'
    if ($dlg.ShowDialog() -eq 'OK') {
        # Kopie bewaren bij de instellingen, dan maakt verplaatsen van het origineel niet uit
        $ext = [System.IO.Path]::GetExtension($dlg.FileName)
        $dest = Join-Path $DataDir ('logo' + $ext)
        if ($script:LogoImg) { $script:LogoImg.Dispose(); $script:LogoImg = $null }
        Get-ChildItem $DataDir -Filter 'logo.*' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Copy-Item $dlg.FileName $dest -Force
        $script:Config.LogoPath = $dest; Save-Config; Load-Logo
        if ($script:LogoImg) { [System.Windows.Forms.MessageBox]::Show('Logo ingesteld. Klik op Testlabel om te kijken hoe het eruitziet.', 'Logo') | Out-Null }
        else { [System.Windows.Forms.MessageBox]::Show('Dit bestand kon ik niet als afbeelding openen.', 'Logo') | Out-Null }
    }
})
function Show-LabelSettings {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Instellingen'; $f.Size = New-Object System.Drawing.Size(560, 560)
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = $form.Font

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Labelformaat (kies het formaat met 11352 of 25 x 54 mm):"
    $l1.Location = New-Object System.Drawing.Point(15, 15); $l1.AutoSize = $true; $f.Controls.Add($l1)
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'; $cmb.Location = New-Object System.Drawing.Point(15, 40); $cmb.Size = New-Object System.Drawing.Size(510, 28)
    $f.Controls.Add($cmb)
    $pset = New-Object System.Drawing.Printing.PrinterSettings
    if ($script:Config.Printer) { $pset.PrinterName = $script:Config.Printer }
    [void]$cmb.Items.Add('Automatisch')
    $names = @()
    if ($pset.IsValid) {
        foreach ($p in $pset.PaperSizes) {
            $txt = "{0}   ({1:N0} x {2:N0} mm)" -f $p.PaperName, ($p.Width * 0.254), ($p.Height * 0.254)
            [void]$cmb.Items.Add($txt); $names += $p.PaperName
        }
    }
    $cmb.SelectedIndex = 0
    if ($script:Config.PaperName) { $i = [array]::IndexOf($names, [string]$script:Config.PaperName); if ($i -ge 0) { $cmb.SelectedIndex = $i + 1 } }

    function New-Num($label, $y, $min, $max, $inc, $dec, $val) {
        $lb = New-Object System.Windows.Forms.Label
        $lb.Text = $label; $lb.Location = New-Object System.Drawing.Point(15, ($y + 3)); $lb.AutoSize = $true; $f.Controls.Add($lb)
        $n = New-Object System.Windows.Forms.NumericUpDown
        $n.Location = New-Object System.Drawing.Point(330, $y); $n.Size = New-Object System.Drawing.Size(90, 28)
        $n.Minimum = $min; $n.Maximum = $max; $n.Increment = $inc; $n.DecimalPlaces = $dec
        $n.Value = [decimal][Math]::Min($max, [Math]::Max($min, [double]$val)); $f.Controls.Add($n)
        return $n
    }
    $nX = New-Num 'Naar rechts schuiven (mm, min = naar links):' 85 -15 15 0.5 1 $script:Config.OffXmm
    $nY = New-Num 'Naar beneden schuiven (mm, min = omhoog):' 120 -15 15 0.5 1 $script:Config.OffYmm
    $nS = New-Num 'Grootte (%):' 155 30 300 5 0 $script:Config.ScalePct

    $lbW = New-Object System.Windows.Forms.Label
    $lbW.Text = "Break herkennen aan deze woorden in de producttitel (met komma's):"
    $lbW.Location = New-Object System.Drawing.Point(15, 198); $lbW.AutoSize = $true; $f.Controls.Add($lbW)
    $txtW = New-Object System.Windows.Forms.TextBox
    $txtW.Location = New-Object System.Drawing.Point(15, 222); $txtW.Size = New-Object System.Drawing.Size(510, 28)
    $txtW.Text = [string]$script:Config.BreakWords; $f.Controls.Add($txtW)

    $lbU = New-Object System.Windows.Forms.Label
    $lbU.Text = "Update-adres (link naar de GitHub-pagina van de stickerprinter):"
    $lbU.Location = New-Object System.Drawing.Point(15, 258); $lbU.AutoSize = $true; $f.Controls.Add($lbU)
    $txtU = New-Object System.Windows.Forms.TextBox
    $txtU.Location = New-Object System.Drawing.Point(15, 282); $txtU.Size = New-Object System.Drawing.Size(380, 28)
    $txtU.Text = [string]$script:Config.UpdateUrl; $f.Controls.Add($txtU)
    $bUpd = New-Object System.Windows.Forms.Button
    $bUpd.Text = 'Nu controleren'; $bUpd.Location = New-Object System.Drawing.Point(405, 280); $bUpd.Size = New-Object System.Drawing.Size(120, 30)
    $bUpd.add_Click({ & $apply; Check-Update $true; if ($script:RestartAfterClose) { $f.Close() } })
    $f.Controls.Add($bUpd)

    $info = New-Object System.Windows.Forms.TextBox
    $info.Multiline = $true; $info.ReadOnly = $true; $info.ScrollBars = 'Vertical'
    $chkPaid = New-Object System.Windows.Forms.CheckBox
    $chkPaid.Text = 'Alleen printen als de order betaald is en de naam erbij staat'
    $chkPaid.Location = New-Object System.Drawing.Point(15, 318); $chkPaid.AutoSize = $true
    $chkPaid.Checked = [bool]$script:Config.RequirePaid; $f.Controls.Add($chkPaid)

    $info.Location = New-Object System.Drawing.Point(15, 348); $info.Size = New-Object System.Drawing.Size(510, 75)
    $info.Text = if ($script:LastDiag) { "Laatste print:`r`n" + ($script:LastDiag -replace ' \| ', "`r`n") } else { 'Druk op Testlabel om de printergegevens te zien.' }
    $f.Controls.Add($info)

    $apply = {
        if ($cmb.SelectedIndex -le 0) { $script:Config.PaperName = '' } else { $script:Config.PaperName = $names[$cmb.SelectedIndex - 1] }
        $script:Config.OffXmm = [double]$nX.Value; $script:Config.OffYmm = [double]$nY.Value; $script:Config.ScalePct = [double]$nS.Value
        $script:Config.BreakWords = $txtW.Text.Trim()
        $script:Config.UpdateUrl = $txtU.Text.Trim()
        $script:Config.RequirePaid = $chkPaid.Checked
        Save-Config
    }
    $bTest = New-Object System.Windows.Forms.Button
    $bTest.Text = 'Opslaan + testlabel'; $bTest.Location = New-Object System.Drawing.Point(15, 435); $bTest.Size = New-Object System.Drawing.Size(160, 32)
    $bTest.add_Click({
        & $apply
        $d = [pscustomobject]@{ OrderId = '1048576213'; Buyer = '@testkoper'; Product = 'Testlabel'; Price = ''; Time = (Get-Date) }
        try { Print-Label $d } catch { [System.Windows.Forms.MessageBox]::Show("Testlabel mislukt: $($_.Exception.Message)", 'Fout') | Out-Null }
        if ($script:LastDiag) { $info.Text = "Laatste print:`r`n" + ($script:LastDiag -replace ' \| ', "`r`n") }
    })
    $f.Controls.Add($bTest)
    $bReset = New-Object System.Windows.Forms.Button
    $bReset.Text = 'Standaard'; $bReset.Location = New-Object System.Drawing.Point(185, 435); $bReset.Size = New-Object System.Drawing.Size(100, 32)
    $bReset.add_Click({ $cmb.SelectedIndex = 0; $nX.Value = 0; $nY.Value = 0; $nS.Value = 100 })
    $f.Controls.Add($bReset)
    $bOk = New-Object System.Windows.Forms.Button
    $bOk.Text = 'Opslaan en sluiten'; $bOk.Location = New-Object System.Drawing.Point(375, 435); $bOk.Size = New-Object System.Drawing.Size(150, 32)
    $bOk.add_Click({ & $apply; $f.Close() })
    $f.Controls.Add($bOk)
    [void]$f.ShowDialog($form)
}
$btnCal.add_Click({ Show-LabelSettings })

$btnHelp.add_Click({
    $help = Join-Path $PSScriptRoot 'LEES MIJ EERST - Uitleg.txt'
    if (Test-Path $help) { Start-Process notepad.exe ('"' + $help + '"') }
    else { [System.Windows.Forms.MessageBox]::Show('Uitlegbestand niet gevonden. Het hoort in dezelfde map te staan als het programma.', 'Uitleg') | Out-Null }
})
$btnNoLogo.add_Click({
    if ($script:LogoImg) { $script:LogoImg.Dispose(); $script:LogoImg = $null }
    $script:Config.LogoPath = ''; Save-Config
    [System.Windows.Forms.MessageBox]::Show('Logo weggehaald. Stickers worden weer zonder logo geprint.', 'Logo') | Out-Null
})
$form.add_FormClosing({ $timer.Stop(); Save-Config; Save-Seen; Save-Buyers })

$form.add_Shown({
    Check-Update $false
    if ($script:RestartAfterClose) { return }
    if (-not $cmbPrinter.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show('Geen DYMO-printer gevonden in Windows. Kies je printer bovenin, of installeer eerst DYMO Connect.', 'Printer') | Out-Null
    }
    if (-not (Get-Token)) { [void](Ask-Token) }
    if (Get-Token) { $script:Running = $true; Update-StartButton; Poll }
    else { Update-StartButton; Set-Status 'Geen API-sleutel ingesteld. Klik op "API-sleutel wijzigen".' 'fout' }
    $timer.Start()
})

Load-Logo
Refresh-Buyers
Update-StartButton
[System.Windows.Forms.Application]::Run($form)
if ($script:RestartAfterClose) {
    Start-Process powershell.exe -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File', ('"' + (Join-Path $PSScriptRoot 'WhatnotStickerprinter.ps1') + '"'))
}
} catch {
    [System.Windows.Forms.MessageBox]::Show("Er ging iets mis bij het starten:`n`n$($_.Exception.Message)", 'Whatnot Stickerprinter') | Out-Null
}
