# Dienstplan-Kalender Setup

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir "config.yaml"
$PythonScript = Join-Path $ScriptDir "dienstplan.py"

# ── Hilfsfunktionen ───────────────────────────────────────────────────────────

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "       Dienstplan - Kalender Konverter        " -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
}

function Read-Name {
    if (-not (Test-Path $ConfigPath)) { return "" }
    $lines = Get-Content $ConfigPath -Encoding UTF8
    foreach ($line in $lines) {
        if ($line -match '^name:\s*"(.+)"') { return $Matches[1] }
    }
    return ""
}

function Save-Name($name) {
    if (-not (Test-Path $ConfigPath)) { return }
    $content = Get-Content $ConfigPath -Raw -Encoding UTF8
    # Nur die erste Zeile die mit "name:" beginnt ersetzen (kein Einzug)
    $content = $content -replace '(?m)^name:\s*"[^"]*"', "name: `"$name`""
    [System.IO.File]::WriteAllText($ConfigPath, $content, [System.Text.Encoding]::UTF8)
}

function Ask($prompt, $default) {
    if ($default) {
        Write-Host "  $prompt [$default]: " -NoNewline -ForegroundColor White
    } else {
        Write-Host "  ${prompt}: " -NoNewline -ForegroundColor White
    }
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $default }
    return $val.Trim()
}

function Pick-PDF {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title  = "Dienstplan auswaehlen (PDF oder Excel)"
    $dialog.Filter = "Dienstplan Dateien (*.pdf;*.xlsx;*.xls)|*.pdf;*.xlsx;*.xls|PDF (*.pdf)|*.pdf|Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Alle Dateien (*.*)|*.*"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    if ($dialog.ShowDialog() -eq "OK") { return $dialog.FileName }
    return $null
}

# ── Schritt 1: Name ───────────────────────────────────────────────────────────

Show-Header

$currentName = Read-Name

Write-Host "  Schritt 1/3 -- Dein Name" -ForegroundColor Yellow
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Host "  Gib deinen Namen ein wie er im Dienstplan steht." -ForegroundColor Gray
Write-Host "  (Format: Nachname, Vorname)" -ForegroundColor Gray
Write-Host ""

$name = Ask "Name" $currentName

if ([string]::IsNullOrWhiteSpace($name)) {
    Write-Host ""
    Write-Host "  Fehler: Kein Name eingegeben." -ForegroundColor Red
    Read-Host "  Enter druecken zum Beenden"
    exit 1
}

Save-Name $name
Write-Host ""
Write-Host "  OK - Name gespeichert: $name" -ForegroundColor Green

# ── Schritt 2: Dienst-Zeiten anpassen (optional) ─────────────────────────────

Write-Host ""
Write-Host "  Schritt 2/3 -- Dienst-Zeiten (optional)" -ForegroundColor Yellow
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Moechtest du Dienst-Zeiten anpassen?" -ForegroundColor Gray
Write-Host ""

$changeShifts = Ask "Anpassen? (j/n)" "n"

if ($changeShifts -eq "j") {
    Write-Host ""
    Write-Host "  Verfuegbare Kuerzel:" -ForegroundColor White
    Write-Host ""
    Write-Host "    PPK   | 1 FT  | 1 K   | k+K   | 1 PP" -ForegroundColor Gray
    Write-Host "    KRD   | 1 ZT  | OSB   | O     | OA" -ForegroundColor Gray
    Write-Host ""

    $kuerzel = (Ask "Kuerzel eingeben (oder Enter zum Ueberspringen)" "").ToUpper()

    if ($kuerzel -ne "") {
        Write-Host ""

        $content = Get-Content $ConfigPath -Raw -Encoding UTF8
        $lines   = $content -split "`n"

        # Finde den Block des Kuerzesl und aendere start/end
        $inBlock  = $false
        $result   = @()
        $changed  = $false

        foreach ($line in $lines) {
            $trimmed = $line.TrimStart()

            # Neuer Block beginnt mit  "KUERZEL":
            if ($trimmed -match '^"(.+)":') {
                $inBlock = ($Matches[1] -eq $kuerzel)
            }

            if ($inBlock -and $trimmed -match '^start:') {
                $newVal = Ask "  Startzeit (HH:MM)" "07:15"
                $line   = $line -replace '"[^"]*"', "`"$newVal`""
                $changed = $true
            }
            elseif ($inBlock -and $trimmed -match '^end:') {
                $newVal = Ask "  Endzeit   (HH:MM)" "15:45"
                $line   = $line -replace '"[^"]*"', "`"$newVal`""
                $changed = $true
            }

            $result += $line
        }

        if ($changed) {
            [System.IO.File]::WriteAllText($ConfigPath, ($result -join "`n"), [System.Text.Encoding]::UTF8)
            Write-Host ""
            Write-Host "  OK - Zeiten aktualisiert." -ForegroundColor Green
        } else {
            Write-Host "  Kuerzel nicht gefunden oder keine einfachen Zeiten (KRD/k+K bitte direkt in config.yaml bearbeiten)." -ForegroundColor DarkYellow
        }
    }
}

# ── Schritt 3: PDF auswaehlen und ausfuehren ──────────────────────────────────

Write-Host ""
Write-Host "  Schritt 3/3 -- PDF auswaehlen" -ForegroundColor Yellow
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Ein Datei-Dialog oeffnet sich..." -ForegroundColor Gray
Write-Host ""

$pdfPath = Pick-PDF

if (-not $pdfPath) {
    Write-Host "  Kein PDF ausgewaehlt. Abbruch." -ForegroundColor Red
    Read-Host "  Enter druecken zum Beenden"
    exit 1
}

Write-Host "  OK - PDF: $pdfPath" -ForegroundColor Green
Write-Host ""
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

python $PythonScript $pdfPath $name

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  Name nicht gefunden oder Fehler. Starte Debug-Modus..." -ForegroundColor Yellow
    Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    $ext = [System.IO.Path]::GetExtension($pdfPath).ToLower()
    if ($ext -eq ".xlsx" -or $ext -eq ".xls" -or $ext -eq ".xlsm") {
        python $PythonScript $pdfPath --scan
    } else {
        python $PythonScript $pdfPath $name --debug
    }
    Write-Host ""
    Read-Host "  Enter druecken zum Beenden"
    exit 1
}

# --- Alle KollegInnen? ---
Write-Host ""
Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Moechtest du fuer ALLE KollegInnen" -ForegroundColor Yellow
Write-Host "  ebenfalls Kalenderdateien erstellen?" -ForegroundColor Yellow
Write-Host "  (je eine .ics Datei pro Person)" -ForegroundColor Gray
Write-Host ""

$alleKollegen = Ask "Fuer alle erstellen? (j/n)" "n"

if ($alleKollegen -eq "j") {
    Write-Host ""
    Write-Host "  --------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    python $PythonScript $pdfPath --alle
}

# Ausgabeordner im Explorer oeffnen
$pdfDir = Split-Path -Parent $pdfPath
Write-Host ""
Write-Host "  Oeffne Ordner mit den erstellten Dateien..." -ForegroundColor Gray
Start-Process explorer.exe -ArgumentList "`"$pdfDir`""

Write-Host ""
Read-Host "  Enter druecken zum Beenden"
