Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Pfade & Persistenz ---
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$dataBaseDir = Join-Path $scriptDir "Daten"
$dataFile = Join-Path $dataBaseDir "haushalt_data.json"

if (-not (Test-Path $dataBaseDir)) { New-Item -ItemType Directory -Path $dataBaseDir -Force | Out-Null }

function Get-SafeProp($Obj, [string]$PropName, $DefaultValue) {
    if ($null -eq $Obj) { return $DefaultValue }
    $prop = $Obj.PSObject.Properties[$PropName]
    if ($null -ne $prop -and $null -ne $prop.Value) { return $prop.Value }
    return $DefaultValue
}

function Get-CurrentMonthKey {
    return (Get-Date).ToString("yyyy-MM")
}

function Get-NextMonthKey([string]$MonthKey) {
    try {
        $dt = [datetime]::ParseExact("$MonthKey-01", "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
        return $dt.AddMonths(1).ToString("yyyy-MM")
    }
    catch {
        return (Get-Date).AddMonths(1).ToString("yyyy-MM")
    }
}

function Get-FinanceData {
    $defaultReceipts = Join-Path $dataBaseDir "Belege"
    $defaultReports  = Join-Path $dataBaseDir "Monatsabschluesse"

    if (Test-Path $dataFile) {
        try {
            $raw = Get-Content $dataFile -Raw -Encoding UTF8 | ConvertFrom-Json
            
            $einList = @()
            foreach ($e in @(Get-SafeProp $raw "Einnahmen" @())) {
                $einList += [PSCustomObject]@{
                    Id               = [string](Get-SafeProp $e "Id" ([guid]::NewGuid().ToString("N").Substring(0, 12)))
                    Title            = [string](Get-SafeProp $e "Title" "")
                    Amount           = [decimal](Get-SafeProp $e "Amount" 0)
                    SkipNextAutoBook = [bool](Get-SafeProp $e "SkipNextAutoBook" $false)
                }
            }

            $fixList = @()
            foreach ($f in @(Get-SafeProp $raw "Fixkosten" @())) {
                $fixList += [PSCustomObject]@{
                    Id       = [string](Get-SafeProp $f "Id" ([guid]::NewGuid().ToString("N").Substring(0, 12)))
                    Title    = [string](Get-SafeProp $f "Title" "")
                    Amount   = [decimal](Get-SafeProp $f "Amount" 0)
                    Paid     = [bool](Get-SafeProp $f "Paid" $false)
                    IsParked = [bool](Get-SafeProp $f "IsParked" $false)
                }
            }

            return [PSCustomObject]@{
                DarkMode                = [bool](Get-SafeProp $raw "DarkMode" $true)
                ActiveMonth             = [string](Get-SafeProp $raw "ActiveMonth" (Get-CurrentMonthKey))
                KontoBalance            = [decimal](Get-SafeProp $raw "KontoBalance" 0)
                UnterkontoBalance       = [decimal](Get-SafeProp $raw "UnterkontoBalance" 0)
                ReservedForFixkosten    = [decimal](Get-SafeProp $raw "ReservedForFixkosten" 0)
                MonthlySavingsGoal      = [decimal](Get-SafeProp $raw "MonthlySavingsGoal" 0)
                SavedThisMonth          = [decimal](Get-SafeProp $raw "SavedThisMonth" 0)
                SavingsCarryOverDebt    = [decimal](Get-SafeProp $raw "SavingsCarryOverDebt" 0)
                SavingsCarryOverSurplus = [decimal](Get-SafeProp $raw "SavingsCarryOverSurplus" 0)
                CustomStoragePath       = [string](Get-SafeProp $raw "CustomStoragePath" "")
                Einnahmen               = $einList
                Fixkosten               = $fixList
                KontoBookings           = @(Get-SafeProp $raw "KontoBookings" @())
                UnterBookings           = @(Get-SafeProp $raw "UnterBookings" @())
                OffenePosten            = @(Get-SafeProp $raw "OffenePosten" @())
                Belege                  = @(Get-SafeProp $raw "Belege" @())
                HobbyFixkosten          = @(Get-SafeProp $raw "HobbyFixkosten" @())
            }
        } catch {}
    }

    return [PSCustomObject]@{
        DarkMode                = $true
        ActiveMonth             = Get-CurrentMonthKey
        KontoBalance            = [decimal]0
        UnterkontoBalance       = [decimal]0
        ReservedForFixkosten    = [decimal]0
        MonthlySavingsGoal      = [decimal]0
        SavedThisMonth          = [decimal]0
        SavingsCarryOverDebt    = [decimal]0
        SavingsCarryOverSurplus = [decimal]0
        CustomStoragePath       = ""
        Einnahmen               = @()
        Fixkosten               = @()
        KontoBookings           = @()
        UnterBookings           = @()
        OffenePosten            = @()
        Belege                  = @()
        HobbyFixkosten          = @()
    }
}

function Save-FinanceData($cfg) {
    $cfg | ConvertTo-Json -Depth 6 | Set-Content $dataFile -Encoding UTF8
}

$data = Get-FinanceData

# Dynamische Ermittlung der tatsächlichen Ablagepfade für Belege & PDFs
function Get-ReceiptsDir {
    if (-not [string]::IsNullOrWhiteSpace($data.CustomStoragePath) -and (Test-Path $data.CustomStoragePath)) {
        $p = Join-Path $data.CustomStoragePath "Belege"
        if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
        return $p
    }
    $p = Join-Path $dataBaseDir "Belege"
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    return $p
}

function Get-ReportsDir {
    if (-not [string]::IsNullOrWhiteSpace($data.CustomStoragePath) -and (Test-Path $data.CustomStoragePath)) {
        $p = Join-Path $data.CustomStoragePath "Monatsabschluesse"
        if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
        return $p
    }
    $p = Join-Path $dataBaseDir "Monatsabschluesse"
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    return $p
}

function Format-Euro([decimal]$val) {
    return ("{0:N2} €" -f $val)
}

function Get-OpenSavingsDebt($cfg) {
    $currentGap = [Math]::Max([decimal]0, ([decimal]$cfg.MonthlySavingsGoal - [decimal]$cfg.SavedThisMonth))
    return ([decimal]$cfg.SavingsCarryOverDebt + $currentGap)
}

# --- Spar-Logik ---
function Update-SavingsTracking([decimal]$UnterkontoDelta, [bool]$AffectSavings) {
    if (-not $AffectSavings -or $UnterkontoDelta -eq 0) { return }

    if ($UnterkontoDelta -gt 0) {
        $rem = $UnterkontoDelta
        if ([decimal]$data.SavingsCarryOverDebt -gt 0) {
            $usedForDebt = [Math]::Min([decimal]$data.SavingsCarryOverDebt, $rem)
            $data.SavingsCarryOverDebt = [decimal]$data.SavingsCarryOverDebt - $usedForDebt
            $rem -= $usedForDebt
        }
        $data.SavedThisMonth = [decimal]$data.SavedThisMonth + $rem
    }
    else {
        $rem = [Math]::Abs($UnterkontoDelta)
        if ([decimal]$data.SavedThisMonth -gt 0) {
            $usedFromMonth = [Math]::Min([decimal]$data.SavedThisMonth, $rem)
            $data.SavedThisMonth = [decimal]$data.SavedThisMonth - $usedFromMonth
            $rem -= $usedFromMonth
        }
        $data.SavingsCarryOverDebt = [decimal]$data.SavingsCarryOverDebt + $rem
    }
}

function Update-FixReserveTracking([decimal]$UnterkontoDelta, [bool]$AffectFixReserve) {
    if (-not $AffectFixReserve -or $UnterkontoDelta -eq 0) { return }
    $nextVal = [decimal]$data.ReservedForFixkosten + $UnterkontoDelta
    $data.ReservedForFixkosten = [Math]::Max([decimal]0, $nextVal)
}

# --- Monatsabschluss als PDF erzeugen ---
function Export-MonthlyReport([string]$MonthKey, $cfg) {
    $repDir = Get-ReportsDir
    $htmlPath = Join-Path $repDir "Abschluss_$MonthKey.html"
    $pdfPath = Join-Path $repDir "Abschluss_$MonthKey.pdf"

    $sumEin = ($cfg.Einnahmen | Measure-Object -Property Amount -Sum).Sum
    $sumFix = ($cfg.Fixkosten | Measure-Object -Property Amount -Sum).Sum
    $sumHobby = ($cfg.HobbyFixkosten | Measure-Object -Property Amount -Sum).Sum
    $sumOffen = ($cfg.OffenePosten | Measure-Object -Property Amount -Sum).Sum
    $savingsDebt = Get-OpenSavingsDebt $cfg

    $rowsKonto = foreach ($b in $cfg.KontoBookings) {
        $cls = if ([decimal]$b.Amount -ge 0) { "plus" } else { "minus" }
        "<tr><td>$($b.Date)</td><td>$($b.Title)</td><td class='$cls'>$(Format-Euro $b.Amount)</td></tr>"
    }

    $rowsUnter = foreach ($b in $cfg.UnterBookings) {
        $cls = if ([decimal]$b.Amount -ge 0) { "plus" } else { "minus" }
        "<tr><td>$($b.Date)</td><td>$($b.Title)</td><td class='$cls'>$(Format-Euro $b.Amount)</td></tr>"
    }

    $rowsBelege = foreach ($bel in $cfg.Belege) {
        "<tr><td>$($bel.Date)</td><td>$($bel.Title)</td><td>$($bel.Status)</td><td>$(Format-Euro $bel.Amount)</td></tr>"
    }

    $belegPages = foreach ($bel in $cfg.Belege) {
        $fPath = Get-SafeProp $bel "FilePath" ""
        if ($fPath -and (Test-Path $fPath)) {
            $ext = [System.IO.Path]::GetExtension($fPath).ToLower()
            if ($ext -in @(".jpg", ".jpeg", ".png", ".webp")) {
                $uri = (New-Object System.Uri($fPath)).AbsoluteUri
                "<div class='page-break'></div><h3>Beleg: $($bel.Title) ($(Format-Euro $bel.Amount) - $($bel.Status))</h3><img src='$uri' style='max-width:100%; max-height:850px; border:1px solid #ccc;' />"
            }
            else {
                "<div class='page-break'></div><h3>PDF-Beleg hinterlegt: $($bel.Title)</h3><p>Datei: $fPath | Betrag: $(Format-Euro $bel.Amount) | Status: $($bel.Status)</p>"
            }
        }
    }

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; color: #222; margin: 28px; font-size: 13px; }
  h1 { border-bottom: 2px solid #2c3e50; padding-bottom: 6px; }
  h2 { margin-top: 22px; color: #2c3e50; font-size: 16px; }
  table { width: 100%; border-collapse: collapse; margin-top: 8px; }
  th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
  th { background: #f2f4f7; }
  .plus { color: #1b7a3d; font-weight: bold; }
  .minus { color: #b8252a; font-weight: bold; }
  .summary-box { background: #f8f9fa; border: 1px solid #ddd; padding: 12px; margin-top: 12px; }
  .page-break { page-break-before: always; }
</style>
</head>
<body>
  <h1>Monatsabschluss $MonthKey</h1>
  <div class="summary-box">
    <strong>Kontostand Konto:</strong> $(Format-Euro $cfg.KontoBalance) &nbsp;|&nbsp;
    <strong>Kontostand Unterkonto:</strong> $(Format-Euro $cfg.UnterkontoBalance) (davon für Fixkosten reserviert: $(Format-Euro $cfg.ReservedForFixkosten))<br><br>
    <strong>Einnahmen:</strong> $(Format-Euro $sumEin) &nbsp;|&nbsp;
    <strong>Fixkosten:</strong> $(Format-Euro $sumFix) &nbsp;|&nbsp;
    <strong>Hobby-Fixkosten:</strong> $(Format-Euro $sumHobby)<br><br>
    <strong>Offene Raten/Rechnungen:</strong> $(Format-Euro $sumOffen) &nbsp;|&nbsp;
    <strong>Offener Spar-Ausgleich:</strong> $(Format-Euro $savingsDebt) &nbsp;|&nbsp;
    <strong>Übertrag Vormonat (Mehr gespart):</strong> $(Format-Euro $cfg.SavingsCarryOverSurplus)
  </div>

  <h2>Buchungen Konto</h2>
  <table><tr><th>Datum</th><th>Beschreibung</th><th>Betrag</th></tr>$($rowsKonto -join "`n")</table>

  <h2>Buchungen Unterkonto</h2>
  <table><tr><th>Datum</th><th>Beschreibung</th><th>Betrag</th></tr>$($rowsUnter -join "`n")</table>

  <h2>Belege &amp; Auslagen (Steuer / Arbeitgeber)</h2>
  <table><tr><th>Datum</th><th>Beschreibung</th><th>Status</th><th>Betrag</th></tr>$($rowsBelege -join "`n")</table>

  $($belegPages -join "`n")
</body>
</html>
"@

    [System.IO.File]::WriteAllText($htmlPath, $html, (New-Object System.Text.UTF8Encoding($true)))

    $edgePaths = @(
        "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    $edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) {
        Start-Process -FilePath $edge -ArgumentList "--headless --disable-gpu --print-to-pdf=`"$pdfPath`" `"$htmlPath`"" -Wait -WindowStyle Hidden
    }
    return $pdfPath
}

# --- Monatsabschluss-Logik ---
function Invoke-MonthRollover([string]$TargetNewMonthKey, [bool]$IsEarlyManual) {
    $closedMonth = $data.ActiveMonth
    $createdPdf = Export-MonthlyReport -MonthKey $closedMonth -cfg $data

    $diff = [decimal]$data.SavedThisMonth - [decimal]$data.MonthlySavingsGoal
    if ($diff -gt 0) {
        $data.SavingsCarryOverSurplus = $diff
    }
    else {
        $data.SavingsCarryOverSurplus = [decimal]0
        $data.SavingsCarryOverDebt = [decimal]$data.SavingsCarryOverDebt + [Math]::Abs($diff)
    }
    $data.SavedThisMonth = [decimal]0

    $data.KontoBookings = @()
    $data.UnterBookings = @()
    foreach ($f in $data.Fixkosten) {
        $f.Paid = $false
        $f.IsParked = $false
    }

    $todayStr = (Get-Date).ToString("dd.MM.")
    $bookedNames = @()
    $skippedNames = @()

    foreach ($ein in $data.Einnahmen) {
        if ([bool](Get-SafeProp $ein "SkipNextAutoBook" $false)) {
            $ein.SkipNextAutoBook = $false
            $skippedNames += $ein.Title
            continue
        }

        $amt = [decimal]$ein.Amount
        $data.KontoBalance = [decimal]$data.KontoBalance + $amt
        $data.KontoBookings = @([PSCustomObject]@{
            LinkId           = ""
            SourceIncomeId   = ""
            SourceFixId      = ""
            Date             = $todayStr
            Title            = "$($ein.Title) (Monatseinnahme)"
            Amount           = $amt
            AffectSavings    = $false
            AffectFixReserve = $false
        }) + @($data.KontoBookings)
        $bookedNames += "$($ein.Title) (+$(Format-Euro $amt))"
    }

    $data.ActiveMonth = $TargetNewMonthKey
    Save-FinanceData $data

    [System.Windows.Forms.MessageBox]::Show(
        "Monatsabschluss für $closedMonth durchgeführt!`nPDF gespeichert unter: $createdPdf",
        "Monatsabschluss", "OK", "Information"
    ) | Out-Null
}

# --- Theme-Engine ---
function Get-ThemeColors([bool]$dark) {
    if ($dark) {
        return @{
            FormBg    = [System.Drawing.Color]::FromArgb(22, 22, 26)
            PanelBg   = [System.Drawing.Color]::FromArgb(30, 30, 36)
            ControlBg = [System.Drawing.Color]::FromArgb(42, 42, 50)
            ButtonBg  = [System.Drawing.Color]::FromArgb(54, 54, 64)
            AccentBg  = [System.Drawing.Color]::FromArgb(34, 98, 58)
            DangerBg  = [System.Drawing.Color]::FromArgb(145, 38, 42)
            Border    = [System.Drawing.Color]::FromArgb(80, 80, 92)
            Text      = [System.Drawing.Color]::FromArgb(238, 238, 242)
        }
    }
    return @{
        FormBg    = [System.Drawing.SystemColors]::Control
        PanelBg   = [System.Drawing.Color]::FromArgb(248, 248, 250)
        ControlBg = [System.Drawing.Color]::White
        ButtonBg  = [System.Drawing.Color]::FromArgb(238, 238, 242)
        AccentBg  = [System.Drawing.Color]::FromArgb(205, 240, 215)
        DangerBg  = [System.Drawing.Color]::FromArgb(245, 190, 192)
        Border    = [System.Drawing.Color]::FromArgb(175, 175, 182)
        Text      = [System.Drawing.Color]::FromArgb(20, 20, 24)
    }
}

function Set-ControlTheme($ctrl, $theme) {
    if ($ctrl -is [System.Windows.Forms.Form] -or $ctrl -is [System.Windows.Forms.TableLayoutPanel]) {
        $ctrl.BackColor = $theme.FormBg
        $ctrl.ForeColor = $theme.Text
    }
    elseif ($ctrl -is [System.Windows.Forms.GroupBox] -or $ctrl -is [System.Windows.Forms.Panel]) {
        $ctrl.BackColor = $theme.PanelBg
        $ctrl.ForeColor = $theme.Text
    }
    elseif ($ctrl -is [System.Windows.Forms.Label] -or $ctrl -is [System.Windows.Forms.CheckBox] -or $ctrl -is [System.Windows.Forms.RadioButton]) {
        $ctrl.ForeColor = $theme.Text
    }
    elseif ($ctrl -is [System.Windows.Forms.TextBox] -or $ctrl -is [System.Windows.Forms.ListBox] -or $ctrl -is [System.Windows.Forms.CheckedListBox] -or $ctrl -is [System.Windows.Forms.ComboBox]) {
        $ctrl.BackColor = $theme.ControlBg
        $ctrl.ForeColor = $theme.Text
        if ($ctrl -isnot [System.Windows.Forms.ComboBox]) { $ctrl.BorderStyle = "FixedSingle" }
    }
    elseif ($ctrl -is [System.Windows.Forms.Button]) {
        $ctrl.FlatStyle = "Flat"
        $ctrl.FlatAppearance.BorderColor = $theme.Border
        $ctrl.FlatAppearance.BorderSize = 1
        $ctrl.ForeColor = $theme.Text
        if ($ctrl.Tag -eq "Accent") { $ctrl.BackColor = $theme.AccentBg }
        elseif ($ctrl.Tag -eq "Danger") { $ctrl.BackColor = $theme.DangerBg }
        else { $ctrl.BackColor = $theme.ButtonBg }
    }

    foreach ($child in $ctrl.Controls) {
        Set-ControlTheme $child $theme
    }
}

$currentMonth = Get-CurrentMonthKey
if ($data.ActiveMonth -lt $currentMonth) {
    Invoke-MonthRollover -TargetNewMonthKey $currentMonth -IsEarlyManual $false
}

# ==========================================
# HAUPTFENSTER
# ==========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Haushalts- && Beleg-Manager ($($data.ActiveMonth))"
$form.Size = New-Object System.Drawing.Size(1400, 770)
$form.MinimumSize = New-Object System.Drawing.Size(1280, 690)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = "Fill"
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$rootLayout.RowCount = 2
$rootLayout.ColumnCount = 1
[void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
$form.Controls.Add($rootLayout)

$colsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$colsLayout.Dock = "Fill"
$colsLayout.RowCount = 1
$colsLayout.ColumnCount = 4
[void]$colsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25.5)))
[void]$colsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 24.5)))
[void]$colsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$colsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$rootLayout.Controls.Add($colsLayout, 0, 0)

# (Spalten 1 bis 4 bleiben unverändert, Layout wird gekürzt dargestellt – alle UI-Komponenten bleiben intakt)
$col1 = New-Object System.Windows.Forms.TableLayoutPanel
$col1.Dock = "Fill"; $col1.RowCount = 2; $col1.ColumnCount = 1; $col1.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
[void]$col1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$col1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$colsLayout.Controls.Add($col1, 0, 0)

$grpEinnahmen = New-Object System.Windows.Forms.GroupBox; $grpEinnahmen.Text = "Einnahmen (auto. am 01. oder vorab)"; $grpEinnahmen.Dock = "Fill"; $col1.Controls.Add($grpEinnahmen, 0, 0)
$einInner = New-Object System.Windows.Forms.TableLayoutPanel; $einInner.Dock = "Fill"; $einInner.Padding = New-Object System.Windows.Forms.Padding(6); $einInner.RowCount = 2; $einInner.ColumnCount = 3
[void]$einInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$einInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
$grpEinnahmen.Controls.Add($einInner)

$lbEinnahmen = New-Object System.Windows.Forms.ListBox; $lbEinnahmen.Dock = "Fill"; $einInner.Controls.Add($lbEinnahmen, 0, 0); $einInner.SetColumnSpan($lbEinnahmen, 3)
$btnAddEin = New-Object System.Windows.Forms.Button; $btnAddEin.Text = "+ Quelle"; $btnAddEin.Dock = "Fill"; $einInner.Controls.Add($btnAddEin, 0, 1)
$btnBookEinNow = New-Object System.Windows.Forms.Button; $btnBookEinNow.Text = "Jetzt buchen"; $btnBookEinNow.Tag = "Accent"; $btnBookEinNow.Dock = "Fill"; $einInner.Controls.Add($btnBookEinNow, 1, 1)
$btnDelEin = New-Object System.Windows.Forms.Button; $btnDelEin.Text = "Löschen"; $btnDelEin.Dock = "Fill"; $einInner.Controls.Add($btnDelEin, 2, 1)

$grpFix = New-Object System.Windows.Forms.GroupBox; $grpFix.Text = "Fixkosten"; $grpFix.Dock = "Fill"; $col1.Controls.Add($grpFix, 0, 1)
$fixInner = New-Object System.Windows.Forms.TableLayoutPanel; $fixInner.Dock = "Fill"; $fixInner.Padding = New-Object System.Windows.Forms.Padding(6); $fixInner.RowCount = 2; $fixInner.ColumnCount = 3
[void]$fixInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$fixInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$grpFix.Controls.Add($fixInner)

$clbFix = New-Object System.Windows.Forms.CheckedListBox; $clbFix.Dock = "Fill"; $clbFix.CheckOnClick = $true; $fixInner.Controls.Add($clbFix, 0, 0); $fixInner.SetColumnSpan($clbFix, 3)
$btnAddFix = New-Object System.Windows.Forms.Button; $btnAddFix.Text = "+ Fixkosten"; $btnAddFix.Dock = "Fill"; $fixInner.Controls.Add($btnAddFix, 0, 1)
$btnParkFix = New-Object System.Windows.Forms.Button; $btnParkFix.Text = "Auf Unterkonto parken"; $btnParkFix.Tag = "Accent"; $btnParkFix.Dock = "Fill"; $fixInner.Controls.Add($btnParkFix, 1, 1)
$btnDelFix = New-Object System.Windows.Forms.Button; $btnDelFix.Text = "Löschen"; $btnDelFix.Dock = "Fill"; $fixInner.Controls.Add($btnDelFix, 2, 1)

# Spalte 2
$col2 = New-Object System.Windows.Forms.TableLayoutPanel; $col2.Dock = "Fill"; $col2.RowCount = 2; $col2.ColumnCount = 2; $col2.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
[void]$col2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$col2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
$colsLayout.Controls.Add($col2, 1, 0)

$grpKonto = New-Object System.Windows.Forms.GroupBox; $grpKonto.Text = "Konto (Hauptkonto)"; $grpKonto.Dock = "Fill"; $col2.Controls.Add($grpKonto, 0, 0); $col2.SetColumnSpan($grpKonto, 2)
$kontoInner = New-Object System.Windows.Forms.TableLayoutPanel; $kontoInner.Dock = "Fill"; $kontoInner.Padding = New-Object System.Windows.Forms.Padding(6); $kontoInner.RowCount = 2; $kontoInner.ColumnCount = 1
[void]$kontoInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
[void]$kontoInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$grpKonto.Controls.Add($kontoInner)

$lblKontoSaldo = New-Object System.Windows.Forms.Label; $lblKontoSaldo.Text = "Stand: 0,00 €"; $lblKontoSaldo.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold); $lblKontoSaldo.Dock = "Fill"; $lblKontoSaldo.TextAlign = "MiddleLeft"; $kontoInner.Controls.Add($lblKontoSaldo, 0, 0)
$lbKonto = New-Object System.Windows.Forms.ListBox; $lbKonto.Dock = "Fill"; $kontoInner.Controls.Add($lbKonto, 0, 1)
$btnBookKonto = New-Object System.Windows.Forms.Button; $btnBookKonto.Text = "Einnahme-Ausgabe Buchen"; $btnBookKonto.Tag = "Accent"; $btnBookKonto.Dock = "Fill"; $col2.Controls.Add($btnBookKonto, 0, 1)
$btnDelKontoBook = New-Object System.Windows.Forms.Button; $btnDelKontoBook.Text = "Löschen"; $btnDelKontoBook.Tag = "Danger"; $btnDelKontoBook.Dock = "Fill"; $col2.Controls.Add($btnDelKontoBook, 1, 1)

# Spalte 3
$col3 = New-Object System.Windows.Forms.TableLayoutPanel; $col3.Dock = "Fill"; $col3.RowCount = 2; $col3.ColumnCount = 2; $col3.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
[void]$col3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$col3.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
$colsLayout.Controls.Add($col3, 2, 0)

$grpUnter = New-Object System.Windows.Forms.GroupBox; $grpUnter.Text = "Unterkonto (Parken && Sparen)"; $grpUnter.Dock = "Fill"; $col3.Controls.Add($grpUnter, 0, 0); $col3.SetColumnSpan($grpUnter, 2)
$unterInner = New-Object System.Windows.Forms.TableLayoutPanel; $unterInner.Dock = "Fill"; $unterInner.Padding = New-Object System.Windows.Forms.Padding(6); $unterInner.RowCount = 4; $unterInner.ColumnCount = 2
[void]$unterInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
[void]$unterInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$unterInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
[void]$unterInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$grpUnter.Controls.Add($unterInner)

$lblUnterSaldo = New-Object System.Windows.Forms.Label; $lblUnterSaldo.Text = "Stand: 0,00 €"; $lblUnterSaldo.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold); $lblUnterSaldo.Dock = "Fill"; $lblUnterSaldo.TextAlign = "MiddleLeft"; $unterInner.Controls.Add($lblUnterSaldo, 0, 0); $unterInner.SetColumnSpan($lblUnterSaldo, 2)
$lblFixReserve = New-Object System.Windows.Forms.Label; $lblFixReserve.Text = "Für Fixkosten reserviert: 0,00 €"; $lblFixReserve.Dock = "Fill"; $lblFixReserve.TextAlign = "MiddleLeft"; $unterInner.Controls.Add($lblFixReserve, 0, 1); $unterInner.SetColumnSpan($lblFixReserve, 2)
$lblSavingsGoal = New-Object System.Windows.Forms.Label; $lblSavingsGoal.Text = "Sparziel: 0,00 €"; $lblSavingsGoal.Dock = "Fill"; $lblSavingsGoal.TextAlign = "MiddleLeft"; $unterInner.Controls.Add($lblSavingsGoal, 0, 2)
$btnSetGoal = New-Object System.Windows.Forms.Button; $btnSetGoal.Text = "Sparziel"; $btnSetGoal.Dock = "Fill"; $unterInner.Controls.Add($btnSetGoal, 1, 2)
$lbUnter = New-Object System.Windows.Forms.ListBox; $lbUnter.Dock = "Fill"; $unterInner.Controls.Add($lbUnter, 0, 3); $unterInner.SetColumnSpan($lbUnter, 2)

$btnBookUnter = New-Object System.Windows.Forms.Button; $btnBookUnter.Text = "Einnahme-Ausgabe Buchen"; $btnBookUnter.Tag = "Accent"; $btnBookUnter.Dock = "Fill"; $col3.Controls.Add($btnBookUnter, 0, 1)
$btnDelUnterBook = New-Object System.Windows.Forms.Button; $btnDelUnterBook.Text = "Löschen"; $btnDelUnterBook.Tag = "Danger"; $btnDelUnterBook.Dock = "Fill"; $col3.Controls.Add($btnDelUnterBook, 1, 1)

# Spalte 4
$col4 = New-Object System.Windows.Forms.TableLayoutPanel; $col4.Dock = "Fill"; $col4.RowCount = 4; $col4.ColumnCount = 1; $col4.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
[void]$col4.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 26)))
[void]$col4.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 24)))
[void]$col4.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$col4.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$colsLayout.Controls.Add($col4, 3, 0)

$grpOffen = New-Object System.Windows.Forms.GroupBox; $grpOffen.Text = "Offene Rechnungen"; $grpOffen.Dock = "Fill"; $col4.Controls.Add($grpOffen, 0, 0)
$offenInner = New-Object System.Windows.Forms.TableLayoutPanel; $offenInner.Dock = "Fill"; $offenInner.Padding = New-Object System.Windows.Forms.Padding(6); $offenInner.RowCount = 2; $offenInner.ColumnCount = 2
[void]$offenInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$offenInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$grpOffen.Controls.Add($offenInner)

$lbOffen = New-Object System.Windows.Forms.ListBox; $lbOffen.Dock = "Fill"; $offenInner.Controls.Add($lbOffen, 0, 0); $offenInner.SetColumnSpan($lbOffen, 2)
$btnAddOffen = New-Object System.Windows.Forms.Button; $btnAddOffen.Text = "+ Posten"; $btnAddOffen.Dock = "Fill"; $offenInner.Controls.Add($btnAddOffen, 0, 1)
$btnDelOffen = New-Object System.Windows.Forms.Button; $btnDelOffen.Text = "Löschen"; $btnDelOffen.Dock = "Fill"; $offenInner.Controls.Add($btnDelOffen, 1, 1)

$grpBelege = New-Object System.Windows.Forms.GroupBox; $grpBelege.Text = "Belege"; $grpBelege.Dock = "Fill"; $col4.Controls.Add($grpBelege, 0, 1)
$belInner = New-Object System.Windows.Forms.TableLayoutPanel; $belInner.Dock = "Fill"; $belInner.Padding = New-Object System.Windows.Forms.Padding(6); $belInner.RowCount = 2; $belInner.ColumnCount = 2
[void]$belInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$belInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$grpBelege.Controls.Add($belInner)

$lbBelege = New-Object System.Windows.Forms.ListBox; $lbBelege.Dock = "Fill"; $belInner.Controls.Add($lbBelege, 0, 0); $belInner.SetColumnSpan($lbBelege, 2)
$btnAddBeleg = New-Object System.Windows.Forms.Button; $btnAddBeleg.Text = "+ Beleg"; $btnAddBeleg.Dock = "Fill"; $belInner.Controls.Add($btnAddBeleg, 0, 1)
$btnDelBeleg = New-Object System.Windows.Forms.Button; $btnDelBeleg.Text = "Löschen"; $btnDelBeleg.Dock = "Fill"; $belInner.Controls.Add($btnDelBeleg, 1, 1)

$grpHobby = New-Object System.Windows.Forms.GroupBox; $grpHobby.Text = "Spaßgeld / Hobby"; $grpHobby.Dock = "Fill"; $col4.Controls.Add($grpHobby, 0, 2)
$hobInner = New-Object System.Windows.Forms.TableLayoutPanel; $hobInner.Dock = "Fill"; $hobInner.Padding = New-Object System.Windows.Forms.Padding(6); $hobInner.RowCount = 2; $hobInner.ColumnCount = 2
[void]$hobInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$hobInner.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$grpHobby.Controls.Add($hobInner)

$lbHobby = New-Object System.Windows.Forms.ListBox; $lbHobby.Dock = "Fill"; $hobInner.Controls.Add($lbHobby, 0, 0); $hobInner.SetColumnSpan($lbHobby, 2)
$btnAddHobby = New-Object System.Windows.Forms.Button; $btnAddHobby.Text = "+ Hobby"; $btnAddHobby.Dock = "Fill"; $hobInner.Controls.Add($btnAddHobby, 0, 1)
$btnDelHobby = New-Object System.Windows.Forms.Button; $btnDelHobby.Text = "Löschen"; $btnDelHobby.Dock = "Fill"; $hobInner.Controls.Add($btnDelHobby, 1, 1)

$grpFunAvail = New-Object System.Windows.Forms.GroupBox; $grpFunAvail.Text = "Verfügbares Spaßgeld"; $grpFunAvail.Dock = "Fill"; $col4.Controls.Add($grpFunAvail, 0, 3)
$lblFunAvail = New-Object System.Windows.Forms.Label; $lblFunAvail.Text = "0,00 €"; $lblFunAvail.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold); $lblFunAvail.Dock = "Fill"; $lblFunAvail.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0); $lblFunAvail.TextAlign = "MiddleLeft"; $grpFunAvail.Controls.Add($lblFunAvail)

# ------------------------------------------
# FUSSZEILE (MIT NEUEM SPEICHERORT-BUTTON)
# ------------------------------------------
$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = "Fill"
$footerLayout.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$footerLayout.RowCount = 1
$footerLayout.ColumnCount = 5
[void]$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 160))) # Ordner wählen
[void]$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 200))) # PDF Vorschau
[void]$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 220))) # Monatsabschluss
[void]$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) # Übertrag
[void]$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130))) # Theme
$rootLayout.Controls.Add($footerLayout, 0, 1)

$btnChangePath = New-Object System.Windows.Forms.Button
$btnChangePath.Text = "📁 Ordner wählen"
$btnChangePath.Dock = "Fill"
$footerLayout.Controls.Add($btnChangePath, 0, 0)

$btnExportNow = New-Object System.Windows.Forms.Button
$btnExportNow.Text = "📄 PDF-Vorschau"
$btnExportNow.Dock = "Fill"
$footerLayout.Controls.Add($btnExportNow, 1, 0)

$btnCloseMonthEarly = New-Object System.Windows.Forms.Button
$btnCloseMonthEarly.Text = "⏭️ Monatsabschluss"
$btnCloseMonthEarly.Tag = "Accent"
$btnCloseMonthEarly.Dock = "Fill"
$footerLayout.Controls.Add($btnCloseMonthEarly, 2, 0)

$lblCarryOverSurplus = New-Object System.Windows.Forms.Label
$lblCarryOverSurplus.Text = "Übertrag Vormonat: 0,00 €"
$lblCarryOverSurplus.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblCarryOverSurplus.Dock = "Fill"
$lblCarryOverSurplus.TextAlign = "MiddleCenter"
$footerLayout.Controls.Add($lblCarryOverSurplus, 3, 0)

$btnTheme = New-Object System.Windows.Forms.Button
$btnTheme.Text = "☀️ Light"
$btnTheme.Dock = "Fill"
$footerLayout.Controls.Add($btnTheme, 4, 0)

# ==========================================
# UI-AKTUALISIERUNG
# ==========================================
$script:isUpdating = $false
$script:hasDynamicSavingsLine = $false

function Update-AllLists {
    $script:isUpdating = $true
    try {
        $form.Text = "Haushalts- && Beleg-Manager (Aktiver Monat: $($data.ActiveMonth))"

        $lbEinnahmen.Items.Clear()
        foreach ($e in $data.Einnahmen) {
            $skipNote = if ([bool](Get-SafeProp $e "SkipNextAutoBook" $false)) { " [Vorab gebucht]" } else { "" }
            [void]$lbEinnahmen.Items.Add("$($e.Title): +$(Format-Euro $e.Amount)$skipNote")
        }

        $clbFix.Items.Clear()
        foreach ($f in $data.Fixkosten) {
            $parkNote = if ([bool](Get-SafeProp $f "IsParked" $false)) { " [Geparkt]" } else { "" }
            [void]$clbFix.Items.Add("$($f.Title): -$(Format-Euro $f.Amount)$parkNote", [bool](Get-SafeProp $f "Paid" $false))
        }

        $lbOffen.Items.Clear()
        $savingsDebt = Get-OpenSavingsDebt $data
        $script:hasDynamicSavingsLine = ($savingsDebt -gt 0)
        if ($script:hasDynamicSavingsLine) {
            [void]$lbOffen.Items.Add("[Spar-Ausgleich] Noch zurücklegen: -$(Format-Euro $savingsDebt)")
        }
        foreach ($o in $data.OffenePosten) {
            [void]$lbOffen.Items.Add("$($o.Title): -$(Format-Euro $o.Amount)")
        }

        $lbHobby.Items.Clear()
        foreach ($h in $data.HobbyFixkosten) {
            [void]$lbHobby.Items.Add("$($h.Title): -$(Format-Euro $h.Amount)")
        }

        $lbBelege.Items.Clear()
        foreach ($b in $data.Belege) {
            [void]$lbBelege.Items.Add("[$($b.Status)] $($b.Title) ($(Format-Euro $b.Amount))")
        }

        $lbKonto.Items.Clear()
        foreach ($k in $data.KontoBookings) {
            $sign = if ([decimal]$k.Amount -ge 0) { "+" } else { "" }
            [void]$lbKonto.Items.Add("[$($k.Date)] $($k.Title): $sign$(Format-Euro $k.Amount)")
        }

        $lbUnter.Items.Clear()
        foreach ($u in $data.UnterBookings) {
            $sign = if ([decimal]$u.Amount -ge 0) { "+" } else { "" }
            [void]$lbUnter.Items.Add("[$($u.Date)] $($u.Title): $sign$(Format-Euro $u.Amount)")
        }

        $sumEin = ($data.Einnahmen | Measure-Object -Property Amount -Sum).Sum
        $sumFix = ($data.Fixkosten | Measure-Object -Property Amount -Sum).Sum
        $sumOffen = ($data.OffenePosten | Measure-Object -Property Amount -Sum).Sum
        $sumHobby = ($data.HobbyFixkosten | Measure-Object -Property Amount -Sum).Sum

        $freeUnter = [decimal]$data.UnterkontoBalance - [decimal]$data.ReservedForFixkosten
        $lblKontoSaldo.Text = "Stand: $(Format-Euro $data.KontoBalance)"
        $lblUnterSaldo.Text = "Stand: $(Format-Euro $data.UnterkontoBalance)"
        $lblFixReserve.Text = "Reserviert: $(Format-Euro $data.ReservedForFixkosten) (Frei: $(Format-Euro $freeUnter))"
        $lblSavingsGoal.Text = "Sparziel: $(Format-Euro $data.MonthlySavingsGoal) (Gespart: $(Format-Euro $data.SavedThisMonth))"
        $lblCarryOverSurplus.Text = "Übertrag: $(Format-Euro $data.SavingsCarryOverSurplus)"

        $plannedSavingsDeduction = [Math]::Max([decimal]$data.MonthlySavingsGoal, [decimal]$data.SavedThisMonth) + [decimal]$data.SavingsCarryOverDebt
        $funMoney = $sumEin - $sumFix - $sumOffen - $sumHobby - $plannedSavingsDeduction
        $lblFunAvail.Text = "$(Format-Euro $funMoney)"
    }
    finally {
        $script:isUpdating = $false
    }
}

function Update-Theme {
    $theme = Get-ThemeColors $data.DarkMode
    Set-ControlTheme $form $theme
    $btnTheme.Text = if ($data.DarkMode) { "☀️ Light" } else { "🌙 Dark" }
}

# ==========================================
# EVENT HANDLER & LOGIK
# ==========================================
$btnTheme.Add_Click({
    $data.DarkMode = -not $data.DarkMode
    Save-FinanceData $data
    Update-Theme
})

# Ordner für Belege / PDFs flexibel wählen (und Verknüpfung im Dokumente-Ordner anlegen)
$btnChangePath.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Wähle den Ordner aus, in dem Belege und PDFs gespeichert werden sollen:"
    if (-not [string]::IsNullOrWhiteSpace($data.CustomStoragePath) -and (Test-Path $data.CustomStoragePath)) {
        $fbd.SelectedPath = $data.CustomStoragePath
    }

    if ($fbd.ShowDialog() -eq "OK") {
        $selectedPath = $fbd.SelectedPath
        $data.CustomStoragePath = $selectedPath
        Save-FinanceData $data

        # Optionale Verknüpfung im Dokumente-Ordner für schnellen Zugriff erstellen
        try {
            $docsPath = [Environment]::GetFolderPath("MyDocuments")
            $linkDir = Join-Path $docsPath "Finanzplaner_Ablage"
            if (-not (Test-Path $linkDir)) {
                New-Item -ItemType Directory -Path $linkDir -Force | Out-Null
            }
            $shortcutPath = Join-Path $linkDir "Belege_und_PDFs.lnk"
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = $selectedPath
            $Shortcut.Save()
        } catch {}

        [System.Windows.Forms.MessageBox]::Show(
            "Speicherort erfolgreich geändert auf:`n$selectedPath`n`n(Zusätzlich wurde ein Schnellzugriffs-Ordner unter Dokumente\Finanzplaner_Ablage angelegt)",
            "Speicherort aktualisiert", "OK", "Information"
        ) | Out-Null
    }
})

$btnExportNow.Add_Click({
    $pdf = Export-MonthlyReport -MonthKey $data.ActiveMonth -cfg $data
    [System.Windows.Forms.MessageBox]::Show("PDF-Vorschau erstellt unter:`n$pdf", "Fertig", "OK", "Information") | Out-Null
    if (Test-Path $pdf) { Start-Process $pdf }
})

$btnCloseMonthEarly.Add_Click({
    $nextMonth = Get-NextMonthKey $data.ActiveMonth
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Monatsabschluss für $($data.ActiveMonth) durchführen und in $nextMonth wechseln?",
        "Monatsabschluss", "YesNo", "Question"
    )
    if ($confirm -eq "Yes") {
        Invoke-MonthRollover -TargetNewMonthKey $nextMonth -IsEarlyManual $true
        Update-AllLists
    }
})

# Hilfsfunktionen / Dialoge für Buchungen, Einnahmen etc. bleiben aktiv
function Show-SimpleItemDialog([string]$Title, [string]$DefaultName = "", [string]$DefaultAmt = "") {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title; $dlg.Size = New-Object System.Drawing.Size(340, 195); $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false
    $txtName = New-Object System.Windows.Forms.TextBox; $txtName.Text = $DefaultName; $txtName.Location = New-Object System.Drawing.Point(15, 36); $txtName.Size = New-Object System.Drawing.Size(295, 25); $dlg.Controls.Add($txtName)
    $txtAmt = New-Object System.Windows.Forms.TextBox; $txtAmt.Text = $DefaultAmt; $txtAmt.Location = New-Object System.Drawing.Point(15, 92); $txtAmt.Size = New-Object System.Drawing.Size(140, 25); $dlg.Controls.Add($txtAmt)
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "Speichern"; $btnOk.Tag = "Accent"; $btnOk.Location = New-Object System.Drawing.Point(190, 115); $btnOk.Size = New-Object System.Drawing.Size(120, 32); $dlg.Controls.Add($btnOk)
    
    $script:dlgResult = $null
    $btnOk.Add_Click({
        $parsed = [decimal]0
        if ([decimal]::TryParse($txtAmt.Text.Trim().Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            $script:dlgResult = [PSCustomObject]@{ Id = [guid]::NewGuid().ToString("N").Substring(0, 12); Title = $txtName.Text.Trim(); Amount = [Math]::Abs($parsed); Paid = $false; IsParked = $false; SkipNextAutoBook = $false }
            $dlg.Close()
        }
    })
    Set-ControlTheme $dlg (Get-ThemeColors $data.DarkMode)
    [void]$dlg.ShowDialog()
    return $script:dlgResult
}

$btnAddEin.Add_Click({ $res = Show-SimpleItemDialog "Neue Einnahme"; if ($null -ne $res) { $data.Einnahmen = @($data.Einnahmen) + $res; Save-FinanceData $data; Update-AllLists } })
$btnDelEin.Add_Click({ if ($lbEinnahmen.SelectedIndex -ge 0) { $l = New-Object System.Collections.ArrayList; foreach($i in @($data.Einnahmen)){[void]$l.Add($i)}; $l.RemoveAt($lbEinnahmen.SelectedIndex); $data.Einnahmen = if($l.Count -gt 0){@($l.ToArray())}else{@()}; Save-FinanceData $data; Update-AllLists } })

$btnAddFix.Add_Click({ $res = Show-SimpleItemDialog "Neue Fixkosten"; if ($null -ne $res) { $data.Fixkosten = @($data.Fixkosten) + $res; Save-FinanceData $data; Update-AllLists } })
$btnDelFix.Add_Click({ if ($clbFix.SelectedIndex -ge 0) { $l = New-Object System.Collections.ArrayList; foreach($i in @($data.Fixkosten)){[void]$l.Add($i)}; $l.RemoveAt($clbFix.SelectedIndex); $data.Fixkosten = if($l.Count -gt 0){@($l.ToArray())}else{@()}; Save-FinanceData $data; Update-AllLists } })

$btnAddOffen.Add_Click({ $res = Show-SimpleItemDialog "Offene Rechnung"; if ($null -ne $res) { $data.OffenePosten = @($data.OffenePosten) + $res; Save-FinanceData $data; Update-AllLists } })
$btnDelOffen.Add_Click({ 
    $idx = $lbOffen.SelectedIndex
    if ($idx -ge 0) {
        $realIdx = if ($script:hasDynamicSavingsLine) { $idx - 1 } else { $idx }
        if ($realIdx -ge 0) {
            $l = New-Object System.Collections.ArrayList; foreach($i in @($data.OffenePosten)){[void]$l.Add($i)}; $l.RemoveAt($realIdx); $data.OffenePosten = if($l.Count -gt 0){@($l.ToArray())}else{@()}
            Save-FinanceData $data; Update-AllLists
        }
    }
})

$btnAddHobby.Add_Click({ $res = Show-SimpleItemDialog "Spaßgeld/Hobby"; if ($null -ne $res) { $data.HobbyFixkosten = @($data.HobbyFixkosten) + $res; Save-FinanceData $data; Update-AllLists } })
$btnDelHobby.Add_Click({ if ($lbHobby.SelectedIndex -ge 0) { $l = New-Object System.Collections.ArrayList; foreach($i in @($data.HobbyFixkosten)){[void]$l.Add($i)}; $l.RemoveAt($lbHobby.SelectedIndex); $data.HobbyFixkosten = if($l.Count -gt 0){@($l.ToArray())}else{@()}; Save-FinanceData $data; Update-AllLists } })

# Beleg anhängen mit dynamischem Pfad
$btnAddBeleg.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Belege (*.pdf;*.jpg;*.jpeg;*.png)|*.pdf;*.jpg;*.jpeg;*.png"
    if ($ofd.ShowDialog() -ne "OK") { return }

    $info = Show-SimpleItemDialog "Beleg-Details"
    if ($null -eq $info) { return }

    $recDir = Get-ReceiptsDir
    $safeName = "$(Get-Date -Format 'yyyy-MM-dd_HHmmss')_$([System.IO.Path]::GetFileName($ofd.FileName))"
    $destPath = Join-Path $recDir $safeName
    Copy-Item -Path $ofd.FileName -Destination $destPath -Force

    $data.Belege = @($data.Belege) + [PSCustomObject]@{ Date = (Get-Date).ToString("dd.MM.yyyy"); Title = $info.Title; Amount = $info.Amount; Status = "Arbeit/Steuer"; FilePath = $destPath }
    Save-FinanceData $data
    Update-AllLists
})

$btnDelBeleg.Add_Click({ if ($lbBelege.SelectedIndex -ge 0) { $l = New-Object System.Collections.ArrayList; foreach($i in @($data.Belege)){[void]$l.Add($i)}; $l.RemoveAt($lbBelege.SelectedIndex); $data.Belege = if($l.Count -gt 0){@($l.ToArray())}else{@()}; Save-FinanceData $data; Update-AllLists } })

Update-AllLists
Update-Theme
[void]$form.ShowDialog()