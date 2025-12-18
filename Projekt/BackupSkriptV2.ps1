# ================================================================

# Office Backup System - Final Edition v3.0

# ================================================================

# Professionelles Backup-System für Office-Dateien

# Mit intelligenter Änderungserkennung und Auto-Fix Funktionen

# Autor: Backup System Generator

# Version: 3.0 Final

# ================================================================



# .NET-Assemblies für GUI laden

Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing



# Globale Konfiguration

$script:Version = "3.0 Final"

$script:BackupConfig = @{

    SourcePaths = @()

    DestinationPath = ""

    FileTypes = @{

        Word = $true

        Excel = $true

        PowerPoint = $true

        PDF = $true

        Images = $false

        All = $false

    }

    AutoBackup = $false

    BackupInterval = 60

    BackupMode = "Smart"  # "Full", "Incremental" oder "Smart"

    ForceTimeCheck = $true

    LastBackupPath = ""

    MinFileAge = 2  # Minuten - Dateien müssen mindestens so alt sein

}



# ================================================================

# HILFSFUNKTIONEN

# ================================================================



# Funktion zum Aktualisieren von Dateizeitstempeln (Fix für Änderungserkennung)

function Update-FileTimestamp {

    param(

        [string]$FilePath

    )

    try {

        $file = Get-Item $FilePath

        $file.LastWriteTime = Get-Date

        return $true

    }

    catch {

        return $false

    }

}



# Verbesserte Funktion zum Abrufen des letzten Backup-Ordners

function Get-LastBackupFolder {

    param([string]$BackupRoot)

    

    if (Test-Path $BackupRoot) {

        $lastBackup = Get-ChildItem -Path $BackupRoot -Directory | 

            Where-Object { $_.Name -match "^Backup_\d{4}-\d{2}-\d{2}_\d{6}" } |

            Sort-Object Name -Descending | 

            Select-Object -First 1

        

        if ($lastBackup) {

            return $lastBackup.FullName

        }

    }

    return $null

}



# Erweiterte Prüfung ob Datei neu oder geändert ist

function Test-FileNeedsBackup {

    param(

        [System.IO.FileInfo]$File,

        [string]$LastBackupPath,

        [string]$Category,

        [string]$RelativePath,

        [bool]$ForceCheck = $false

    )

    

    # Neue Datei - immer sichern

    if (-not $LastBackupPath) { 

        return @{NeedsBackup = $true; Reason = "Kein vorheriges Backup"} 

    }

    

    # Konstruiere Backup-Dateipfad

    $backupFilePath = Join-Path $LastBackupPath $Category

    if ($RelativePath) {

        $backupFilePath = Join-Path $backupFilePath $RelativePath

    }

    $backupFilePath = Join-Path $backupFilePath $File.Name

    

    # Datei existiert nicht im Backup

    if (-not (Test-Path $backupFilePath)) {

        return @{NeedsBackup = $true; Reason = "Neue Datei"}

    }

    

    $backupFile = Get-Item $backupFilePath

    

    # Zeitvergleich mit Toleranz

    $timeDiff = ($File.LastWriteTime - $backupFile.LastWriteTime).TotalSeconds

    

    # Datei ist neuer (mit 2 Sekunden Toleranz für Zeitungenauigkeiten)

    if ($timeDiff -gt 2) {

        return @{NeedsBackup = $true; Reason = "Neuere Version (${timeDiff}s)"}

    }

    

    # Größenvergleich

    if ($File.Length -ne $backupFile.Length) {

        return @{NeedsBackup = $true; Reason = "Größe geändert"}

    }

    

    # Hash-Vergleich für kritische Dateien (optional, langsamer)

    if ($ForceCheck) {

        $sourceHash = (Get-FileHash $File.FullName -Algorithm MD5).Hash

        $backupHash = (Get-FileHash $backupFile.FullName -Algorithm MD5).Hash

        if ($sourceHash -ne $backupHash) {

            return @{NeedsBackup = $true; Reason = "Inhalt geändert (Hash)"}

        }

    }

    

    return @{NeedsBackup = $false; Reason = "Unverändert"}

}



# ================================================================

# HAUPTFUNKTION: BACKUP-PROZESS

# ================================================================



function Start-BackupProcess {

    param(

        [string[]]$SourcePaths,

        [string]$DestinationPath,

        [hashtable]$FileTypes,

        [string]$BackupMode = "Smart",

        [bool]$ForceBackup = $false

    )

    

    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan

    Write-Host " BACKUP-PROZESS GESTARTET" -ForegroundColor Green

    Write-Host ("=" * 60) -ForegroundColor Cyan

    Write-Host "Zeit: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" -ForegroundColor Yellow

    Write-Host "Modus: $BackupMode $(if($ForceBackup){'[ERZWUNGEN]'})" -ForegroundColor Yellow

    

    # Smart Mode: Entscheide automatisch zwischen Full und Incremental

    if ($BackupMode -eq "Smart") {

        $lastBackup = Get-LastBackupFolder -BackupRoot $DestinationPath

        if ($lastBackup) {

            $lastBackupDate = [DateTime]::ParseExact(

                ($lastBackup | Split-Path -Leaf).Substring(7, 17).Replace('_', ' '),

                'yyyy-MM-dd HHmmss',

                $null

            )

            $daysSinceBackup = (Get-Date).Subtract($lastBackupDate).TotalDays

            

            # Vollbackup wenn älter als 7 Tage

            if ($daysSinceBackup -gt 7) {

                $BackupMode = "Full"

                Write-Host "Smart-Mode: Vollbackup (letztes Backup vor $([Math]::Round($daysSinceBackup, 1)) Tagen)" -ForegroundColor Magenta

            } else {

                $BackupMode = "Incremental"

                Write-Host "Smart-Mode: Inkrementell (letztes Backup vor $([Math]::Round($daysSinceBackup, 1)) Tagen)" -ForegroundColor Magenta

            }

        } else {

            $BackupMode = "Full"

            Write-Host "Smart-Mode: Erstes Backup - Vollständig" -ForegroundColor Magenta

        }

    }

    

    # Letzten Backup-Ordner finden

    $lastBackupPath = $null

    if ($BackupMode -eq "Incremental" -and -not $ForceBackup) {

        $lastBackupPath = Get-LastBackupFolder -BackupRoot $DestinationPath

        if ($lastBackupPath) {

            Write-Host "Vergleiche mit: $(Split-Path $lastBackupPath -Leaf)" -ForegroundColor Gray

        } else {

            Write-Host "Kein vorheriges Backup - führe Vollbackup durch" -ForegroundColor Yellow

            $BackupMode = "Full"

        }

    }

    

    # Backup-Ordner vorbereiten

    $backupDate = Get-Date -Format "yyyy-MM-dd_HHmmss"

    $backupType = switch ($BackupMode) {

        "Full" { "Full" }

        "Incremental" { "Inc" }

        default { "Backup" }

    }

    $backupRoot = Join-Path $DestinationPath "Backup_${backupDate}_${backupType}"

    

    # Dateierweiterungen definieren

    $extensions = @{}

    if ($FileTypes.Word) { 

        $extensions['Word'] = @('*.docx', '*.doc', '*.docm', '*.dotx', '*.dotm', '*.rtf') 

    }

    if ($FileTypes.Excel) { 

        $extensions['Excel'] = @('*.xlsx', '*.xls', '*.xlsm', '*.xlsb', '*.xltx', '*.xltm', '*.csv') 

    }

    if ($FileTypes.PowerPoint) { 

        $extensions['PowerPoint'] = @('*.pptx', '*.ppt', '*.pptm', '*.potx', '*.potm', '*.ppsx', '*.ppsm') 

    }

    if ($FileTypes.PDF) { 

        $extensions['PDF'] = @('*.pdf') 

    }

    if ($FileTypes.Images) { 

        $extensions['Images'] = @('*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff', '*.svg', '*.ico') 

    }

    if ($FileTypes.All) { 

        $extensions['All'] = @('*.*') 

    }

    

    # Statistiken

    $stats = @{

        TotalFiles = 0

        CheckedFiles = 0

        CopiedFiles = 0

        SkippedFiles = 0

        NewFiles = 0

        ModifiedFiles = 0

        Errors = @()

        ProcessedPaths = @()

    }

    

    $backupNeeded = $false

    $fileList = @()

    

    Write-Host "`nAnalysiere Dateien..." -ForegroundColor Cyan

    

    # Durchsuche alle Quellpfade

    foreach ($sourcePath in $SourcePaths) {

        if (-not (Test-Path $sourcePath)) {

            Write-Host "  ✗ Pfad nicht gefunden: $sourcePath" -ForegroundColor Red

            $stats.Errors += "Pfad nicht gefunden: $sourcePath"

            continue

        }

        

        $stats.ProcessedPaths += $sourcePath

        Write-Host "  → Scanne: $sourcePath" -ForegroundColor Gray

        

        foreach ($category in $extensions.Keys) {

            foreach ($extension in $extensions[$category]) {

                try {

                    $files = Get-ChildItem -Path $sourcePath -Filter $extension -Recurse -File -ErrorAction SilentlyContinue

                    

                    foreach ($file in $files) {

                        $stats.TotalFiles++

                        

                        # Berechne relativen Pfad

                        $relativePath = ""

                        if ($file.DirectoryName.StartsWith($sourcePath)) {

                            $relativePath = $file.DirectoryName.Substring($sourcePath.Length).TrimStart('\')

                        }

                        

                        # Prüfe ob Backup nötig

                        if ($ForceBackup -or $BackupMode -eq "Full") {

                            $needsBackup = @{NeedsBackup = $true; Reason = "Vollbackup"}

                        } else {

                            $needsBackup = Test-FileNeedsBackup -File $file `

                                -LastBackupPath $lastBackupPath `

                                -Category $category `

                                -RelativePath $relativePath `

                                -ForceCheck $false

                        }

                        

                        if ($needsBackup.NeedsBackup) {

                            $fileList += @{

                                File = $file

                                Category = $category

                                RelativePath = $relativePath

                                Reason = $needsBackup.Reason

                            }

                            

                            if ($needsBackup.Reason -eq "Neue Datei") {

                                $stats.NewFiles++

                            } elseif ($needsBackup.Reason -like "*geändert*" -or $needsBackup.Reason -like "*Neuere*") {

                                $stats.ModifiedFiles++

                            }

                        } else {

                            $stats.SkippedFiles++

                        }

                        

                        $stats.CheckedFiles++

                        

                        # Fortschrittsanzeige

                        if ($stats.CheckedFiles % 100 -eq 0) {

                            Write-Host "  ... $($stats.CheckedFiles) Dateien geprüft" -ForegroundColor DarkGray

                        }

                    }

                }

                catch {

                    $stats.Errors += "Fehler beim Scannen ($extension): $_"

                }

            }

        }

    }

    

    Write-Host "`nAnalyse abgeschlossen:" -ForegroundColor Green

    Write-Host "  Geprüft: $($stats.CheckedFiles) Dateien" -ForegroundColor Yellow

    Write-Host "  Zu sichern: $($fileList.Count) Dateien" -ForegroundColor $(if($fileList.Count -gt 0){'Green'}else{'Gray'})

    Write-Host "  → Neu: $($stats.NewFiles)" -ForegroundColor Cyan

    Write-Host "  → Geändert: $($stats.ModifiedFiles)" -ForegroundColor Cyan

    Write-Host "  → Unverändert: $($stats.SkippedFiles)" -ForegroundColor Gray

    

    # Backup durchführen wenn nötig

    if ($fileList.Count -gt 0) {

        Write-Host "`nKopiere Dateien..." -ForegroundColor Green

        

        # Erstelle Backup-Ordner

        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null

        $backupNeeded = $true

        

        $progressCount = 0

        foreach ($item in $fileList) {

            $progressCount++

            $percent = [Math]::Round(($progressCount / $fileList.Count) * 100, 0)

            

            try {

                # Zielordner erstellen

                $categoryPath = Join-Path $backupRoot $item.Category

                $targetDir = Join-Path $categoryPath $item.RelativePath

                

                if (-not (Test-Path $targetDir)) {

                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

                }

                

                $targetFile = Join-Path $targetDir $item.File.Name

                

                # Datei kopieren

                Copy-Item -Path $item.File.FullName -Destination $targetFile -Force

                $stats.CopiedFiles++

                

                # Fortschritt anzeigen

                if ($fileList.Count -le 20 -or $progressCount % 10 -eq 0) {

                    Write-Host "  [$percent%] ✓ $($item.File.Name) [$($item.Reason)]" -ForegroundColor Green

                }

            }

            catch {

                Write-Host "  ✗ Fehler bei $($item.File.Name): $_" -ForegroundColor Red

                $stats.Errors += "Kopierfehler bei $($item.File.Name): $_"

            }

        }

    } else {

        Write-Host "`n→ Keine Änderungen gefunden - kein Backup erforderlich" -ForegroundColor Yellow

        

        if ($BackupMode -eq "Incremental" -and $lastBackupPath) {

            Write-Host "  Alle Dateien sind identisch mit: $(Split-Path $lastBackupPath -Leaf)" -ForegroundColor Gray

            Write-Host "`n  💡 Tipp: Falls du Änderungen gemacht hast die nicht erkannt wurden:" -ForegroundColor Cyan

            Write-Host "     - Nutze 'Vollständiges Backup' statt 'Inkrementell'" -ForegroundColor White

            Write-Host "     - Oder aktiviere 'Backup erzwingen' Option" -ForegroundColor White

        }

    }

    

    # Log-Datei erstellen wenn Backup durchgeführt

    if ($backupNeeded) {

        $logFile = Join-Path $backupRoot "backup_log.txt"

        $logContent = @"

================================================================================

BACKUP LOG - Office Backup System v$script:Version

================================================================================

Datum/Zeit: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')

Computer: $env:COMPUTERNAME

Benutzer: $env:USERNAME

Backup-Typ: $BackupMode

--------------------------------------------------------------------------------



QUELLPFADE:

$($stats.ProcessedPaths | ForEach-Object { "  • $_" } | Out-String)



ZIELORDNER:

  $backupRoot



STATISTIKEN:

  Geprüfte Dateien:     $($stats.CheckedFiles)

  Gesicherte Dateien:   $($stats.CopiedFiles)

    → Neue Dateien:     $($stats.NewFiles)

    → Geänderte:        $($stats.ModifiedFiles)

  Übersprungen:         $($stats.SkippedFiles)

  

DATEITYPEN:

$(($FileTypes.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object { "  • $($_.Key)" }) -join "`n")



$(if ($stats.Errors.Count -gt 0) {

"FEHLER:

$($stats.Errors | ForEach-Object { "  ! $_" } | Out-String)"

} else {

"STATUS: Backup erfolgreich abgeschlossen ohne Fehler"

})



================================================================================

Erstellt mit Office Backup System v$script:Version

================================================================================

"@

        $logContent | Out-File -FilePath $logFile -Encoding UTF8

    }

    

    # Zusammenfassung

    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan

    Write-Host " BACKUP ABGESCHLOSSEN" -ForegroundColor $(if($stats.Errors.Count -eq 0){'Green'}else{'Yellow'})

    Write-Host ("=" * 60) -ForegroundColor Cyan

    

    if ($backupNeeded) {

        Write-Host "Ordner: $backupRoot" -ForegroundColor Cyan

        Write-Host "Gesichert: $($stats.CopiedFiles) von $($stats.CheckedFiles) Dateien" -ForegroundColor Green

        

        # Ordnergröße berechnen

        try {

            $size = (Get-ChildItem $backupRoot -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB

            Write-Host "Größe: $([Math]::Round($size, 2)) MB" -ForegroundColor Yellow

        } catch {}

    }

    

    if ($stats.Errors.Count -gt 0) {

        Write-Host "`nFehler: $($stats.Errors.Count)" -ForegroundColor Red

    }

    

    return @{

        Success = ($stats.Errors.Count -eq 0)

        BackupPath = if ($backupNeeded) { $backupRoot } else { $null }

        TotalFiles = $stats.CheckedFiles

        CopiedFiles = $stats.CopiedFiles

        SkippedFiles = $stats.SkippedFiles

        NewFiles = $stats.NewFiles

        ModifiedFiles = $stats.ModifiedFiles

        BackupNeeded = $backupNeeded

        Errors = $stats.Errors

    }

}



# ================================================================

# GUI INTERFACE

# ================================================================



function Show-BackupGUI {

    # Hauptfenster

    $form = New-Object System.Windows.Forms.Form

    $form.Text = "Office Backup System v$script:Version"

    $form.Size = New-Object System.Drawing.Size(750, 700)

    $form.StartPosition = "CenterScreen"

    $form.FormBorderStyle = "FixedDialog"

    $form.MaximizeBox = $false

    $form.Icon = [System.Drawing.SystemIcons]::Shield

    

    # Tab Control

    $tabControl = New-Object System.Windows.Forms.TabControl

    $tabControl.Size = New-Object System.Drawing.Size(710, 550)

    $tabControl.Location = New-Object System.Drawing.Point(10, 10)

    

    # ================================================================

    # TAB 1: BACKUP-EINSTELLUNGEN

    # ================================================================

    $tabBackup = New-Object System.Windows.Forms.TabPage

    $tabBackup.Text = "📁 Backup-Einstellungen"

    $tabBackup.BackColor = [System.Drawing.SystemColors]::Control

    

    # Quellpfade Gruppe

    $grpSource = New-Object System.Windows.Forms.GroupBox

    $grpSource.Text = "Quellordner (zu sichernde Ordner)"

    $grpSource.Location = New-Object System.Drawing.Point(10, 10)

    $grpSource.Size = New-Object System.Drawing.Size(680, 130)

    

    $txtSources = New-Object System.Windows.Forms.TextBox

    $txtSources.Multiline = $true

    $txtSources.ScrollBars = "Vertical"

    $txtSources.Location = New-Object System.Drawing.Point(10, 20)

    $txtSources.Size = New-Object System.Drawing.Size(540, 70)

    $txtSources.Text = "C:\Users\$env:USERNAME\Documents`r`nC:\Users\$env:USERNAME\Desktop"

    $txtSources.Font = New-Object System.Drawing.Font("Consolas", 9)

    

    $btnAddSource = New-Object System.Windows.Forms.Button

    $btnAddSource.Text = "📂 Ordner hinzufügen..."

    $btnAddSource.Location = New-Object System.Drawing.Point(560, 20)

    $btnAddSource.Size = New-Object System.Drawing.Size(110, 30)

    

    $btnClearSources = New-Object System.Windows.Forms.Button

    $btnClearSources.Text = "❌ Leeren"

    $btnClearSources.Location = New-Object System.Drawing.Point(560, 55)

    $btnClearSources.Size = New-Object System.Drawing.Size(110, 30)

    

    $lblSourceTip = New-Object System.Windows.Forms.Label

    $lblSourceTip.Text = "💡 Tipp: Ein Pfad pro Zeile. Für ganzes Laufwerk: C:\"

    $lblSourceTip.Location = New-Object System.Drawing.Point(10, 95)

    $lblSourceTip.Size = New-Object System.Drawing.Size(400, 20)

    $lblSourceTip.ForeColor = [System.Drawing.Color]::DarkBlue

    

    $grpSource.Controls.AddRange(@($txtSources, $btnAddSource, $btnClearSources, $lblSourceTip))

    

    # Zielpfad Gruppe

    $grpDest = New-Object System.Windows.Forms.GroupBox

    $grpDest.Text = "Backup-Zielordner"

    $grpDest.Location = New-Object System.Drawing.Point(10, 145)

    $grpDest.Size = New-Object System.Drawing.Size(680, 80)

    

    $txtDest = New-Object System.Windows.Forms.TextBox

    $txtDest.Location = New-Object System.Drawing.Point(10, 25)

    $txtDest.Size = New-Object System.Drawing.Size(540, 25)

    $txtDest.Text = "D:\Backups"

    $txtDest.Font = New-Object System.Drawing.Font("Consolas", 10)

    

    $btnBrowseDest = New-Object System.Windows.Forms.Button

    $btnBrowseDest.Text = "📁 Durchsuchen..."

    $btnBrowseDest.Location = New-Object System.Drawing.Point(560, 23)

    $btnBrowseDest.Size = New-Object System.Drawing.Size(110, 27)

    

    $lblDestSpace = New-Object System.Windows.Forms.Label

    $lblDestSpace.Location = New-Object System.Drawing.Point(10, 53)

    $lblDestSpace.Size = New-Object System.Drawing.Size(400, 20)

    $lblDestSpace.Text = "Freier Speicher wird hier angezeigt..."

    $lblDestSpace.ForeColor = [System.Drawing.Color]::Gray

    

    $grpDest.Controls.AddRange(@($txtDest, $btnBrowseDest, $lblDestSpace))

    

    # Dateitypen Gruppe

    $grpFileTypes = New-Object System.Windows.Forms.GroupBox

    $grpFileTypes.Text = "Dateitypen für Backup"

    $grpFileTypes.Location = New-Object System.Drawing.Point(10, 230)

    $grpFileTypes.Size = New-Object System.Drawing.Size(680, 110)

    

    $chkWord = New-Object System.Windows.Forms.CheckBox

    $chkWord.Text = "📄 Word (.docx, .doc)"

    $chkWord.Location = New-Object System.Drawing.Point(15, 25)

    $chkWord.Size = New-Object System.Drawing.Size(200, 20)

    $chkWord.Checked = $true

    

    $chkExcel = New-Object System.Windows.Forms.CheckBox

    $chkExcel.Text = "📊 Excel (.xlsx, .xls)"

    $chkExcel.Location = New-Object System.Drawing.Point(230, 25)

    $chkExcel.Size = New-Object System.Drawing.Size(200, 20)

    $chkExcel.Checked = $true

    

    $chkPowerPoint = New-Object System.Windows.Forms.CheckBox

    $chkPowerPoint.Text = "📱 PowerPoint (.pptx, .ppt)"

    $chkPowerPoint.Location = New-Object System.Drawing.Point(445, 25)

    $chkPowerPoint.Size = New-Object System.Drawing.Size(220, 20)

    $chkPowerPoint.Checked = $true

    

    $chkPDF = New-Object System.Windows.Forms.CheckBox

    $chkPDF.Text = "📕 PDF-Dateien"

    $chkPDF.Location = New-Object System.Drawing.Point(15, 50)

    $chkPDF.Size = New-Object System.Drawing.Size(200, 20)

    $chkPDF.Checked = $true

    

    $chkImages = New-Object System.Windows.Forms.CheckBox

    $chkImages.Text = "🖼️ Bilder (.jpg, .png, etc.)"

    $chkImages.Location = New-Object System.Drawing.Point(230, 50)

    $chkImages.Size = New-Object System.Drawing.Size(200, 20)

    

    $chkAll = New-Object System.Windows.Forms.CheckBox

    $chkAll.Text = "💾 ALLE Dateien"

    $chkAll.Location = New-Object System.Drawing.Point(445, 50)

    $chkAll.Size = New-Object System.Drawing.Size(200, 20)

    $chkAll.ForeColor = [System.Drawing.Color]::DarkRed

    

    $btnSelectAll = New-Object System.Windows.Forms.Button

    $btnSelectAll.Text = "Alle"

    $btnSelectAll.Location = New-Object System.Drawing.Point(15, 75)

    $btnSelectAll.Size = New-Object System.Drawing.Size(60, 23)

    

    $btnSelectNone = New-Object System.Windows.Forms.Button

    $btnSelectNone.Text = "Keine"

    $btnSelectNone.Location = New-Object System.Drawing.Point(80, 75)

    $btnSelectNone.Size = New-Object System.Drawing.Size(60, 23)

    

    $grpFileTypes.Controls.AddRange(@($chkWord, $chkExcel, $chkPowerPoint, $chkPDF, $chkImages, $chkAll, $btnSelectAll, $btnSelectNone))

    

    # Backup-Modus Gruppe

    $grpMode = New-Object System.Windows.Forms.GroupBox

    $grpMode.Text = "Backup-Modus"

    $grpMode.Location = New-Object System.Drawing.Point(10, 345)

    $grpMode.Size = New-Object System.Drawing.Size(680, 120)

    

    $radioSmart = New-Object System.Windows.Forms.RadioButton

    $radioSmart.Text = "🧠 Smart (Automatisch entscheiden)"

    $radioSmart.Location = New-Object System.Drawing.Point(15, 25)

    $radioSmart.Size = New-Object System.Drawing.Size(300, 20)

    $radioSmart.Checked = $true

    

    $radioFull = New-Object System.Windows.Forms.RadioButton

    $radioFull.Text = "💯 Vollständig (Alle Dateien)"

    $radioFull.Location = New-Object System.Drawing.Point(15, 48)

    $radioFull.Size = New-Object System.Drawing.Size(300, 20)

    

    $radioIncremental = New-Object System.Windows.Forms.RadioButton

    $radioIncremental.Text = "📈 Inkrementell (Nur Änderungen)"

    $radioIncremental.Location = New-Object System.Drawing.Point(15, 71)

    $radioIncremental.Size = New-Object System.Drawing.Size(300, 20)

    

    $chkForceBackup = New-Object System.Windows.Forms.CheckBox

    $chkForceBackup.Text = "⚡ Backup erzwingen (ignoriert Zeitstempel)"

    $chkForceBackup.Location = New-Object System.Drawing.Point(350, 25)

    $chkForceBackup.Size = New-Object System.Drawing.Size(300, 20)

    $chkForceBackup.ForeColor = [System.Drawing.Color]::DarkBlue

    

    $lblModeInfo = New-Object System.Windows.Forms.Label

    $lblModeInfo.Text = "Smart-Modus: Vollbackup alle 7 Tage, sonst inkrementell"

    $lblModeInfo.Location = New-Object System.Drawing.Point(15, 95)

    $lblModeInfo.Size = New-Object System.Drawing.Size(650, 20)

    $lblModeInfo.ForeColor = [System.Drawing.Color]::Gray

    

    $grpMode.Controls.AddRange(@($radioSmart, $radioFull, $radioIncremental, $chkForceBackup, $lblModeInfo))

    

    # Auto-Backup

    $chkAutoBackup = New-Object System.Windows.Forms.CheckBox

    $chkAutoBackup.Text = "⏰ Automatisches Backup aktivieren"

    $chkAutoBackup.Location = New-Object System.Drawing.Point(25, 475)

    $chkAutoBackup.Size = New-Object System.Drawing.Size(250, 20)

    $chkAutoBackup.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)

    

    $lblInterval = New-Object System.Windows.Forms.Label

    $lblInterval.Text = "Intervall (Min):"

    $lblInterval.Location = New-Object System.Drawing.Point(280, 475)

    $lblInterval.Size = New-Object System.Drawing.Size(85, 20)

    

    $numInterval = New-Object System.Windows.Forms.NumericUpDown

    $numInterval.Location = New-Object System.Drawing.Point(365, 473)

    $numInterval.Size = New-Object System.Drawing.Size(60, 25)

    $numInterval.Minimum = 5

    $numInterval.Maximum = 1440

    $numInterval.Value = 60

    

    $tabBackup.Controls.AddRange(@(

        $grpSource, $grpDest, $grpFileTypes, $grpMode,

        $chkAutoBackup, $lblInterval, $numInterval

    ))

    

    # ================================================================

    # TAB 2: BACKUP-HISTORIE

    # ================================================================

    $tabHistory = New-Object System.Windows.Forms.TabPage

    $tabHistory.Text = "📜 Historie & Statistik"

    

    $lblHistoryTitle = New-Object System.Windows.Forms.Label

    $lblHistoryTitle.Text = "Vorhandene Backups:"

    $lblHistoryTitle.Location = New-Object System.Drawing.Point(10, 10)

    $lblHistoryTitle.Size = New-Object System.Drawing.Size(200, 20)

    $lblHistoryTitle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    

    $listHistory = New-Object System.Windows.Forms.ListBox

    $listHistory.Location = New-Object System.Drawing.Point(10, 35)

    $listHistory.Size = New-Object System.Drawing.Size(680, 300)

    $listHistory.Font = New-Object System.Drawing.Font("Consolas", 9)

    $listHistory.HorizontalScrollbar = $true

    

    $btnRefreshHistory = New-Object System.Windows.Forms.Button

    $btnRefreshHistory.Text = "🔄 Aktualisieren"

    $btnRefreshHistory.Location = New-Object System.Drawing.Point(10, 345)

    $btnRefreshHistory.Size = New-Object System.Drawing.Size(110, 30)

    

    $btnOpenBackup = New-Object System.Windows.Forms.Button

    $btnOpenBackup.Text = "📂 Öffnen"

    $btnOpenBackup.Location = New-Object System.Drawing.Point(130, 345)

    $btnOpenBackup.Size = New-Object System.Drawing.Size(110, 30)

    

    $btnDeleteBackup = New-Object System.Windows.Forms.Button

    $btnDeleteBackup.Text = "🗑️ Löschen"

    $btnDeleteBackup.Location = New-Object System.Drawing.Point(250, 345)

    $btnDeleteBackup.Size = New-Object System.Drawing.Size(110, 30)

    $btnDeleteBackup.BackColor = [System.Drawing.Color]::MistyRose

    

    $btnCleanupOld = New-Object System.Windows.Forms.Button

    $btnCleanupOld.Text = "🧹 Alte aufräumen"

    $btnCleanupOld.Location = New-Object System.Drawing.Point(370, 345)

    $btnCleanupOld.Size = New-Object System.Drawing.Size(110, 30)

    

    $grpStats = New-Object System.Windows.Forms.GroupBox

    $grpStats.Text = "Statistiken"

    $grpStats.Location = New-Object System.Drawing.Point(10, 385)

    $grpStats.Size = New-Object System.Drawing.Size(680, 120)

    

    $lblStats = New-Object System.Windows.Forms.Label

    $lblStats.Location = New-Object System.Drawing.Point(10, 20)

    $lblStats.Size = New-Object System.Drawing.Size(660, 90)

    $lblStats.Font = New-Object System.Drawing.Font("Consolas", 9)

    

    $grpStats.Controls.Add($lblStats)

    

    $tabHistory.Controls.AddRange(@(

        $lblHistoryTitle, $listHistory,

        $btnRefreshHistory, $btnOpenBackup, $btnDeleteBackup, $btnCleanupOld,

        $grpStats

    ))

    

    # ================================================================

    # TAB 3: TOOLS & HILFE

    # ================================================================

    $tabTools = New-Object System.Windows.Forms.TabPage

    $tabTools.Text = "🔧 Tools & Hilfe"

    

    $grpQuickFix = New-Object System.Windows.Forms.GroupBox

    $grpQuickFix.Text = "Quick-Fix Tools"

    $grpQuickFix.Location = New-Object System.Drawing.Point(10, 10)

    $grpQuickFix.Size = New-Object System.Drawing.Size(680, 150)

    

    $lblFixInfo = New-Object System.Windows.Forms.Label

    $lblFixInfo.Text = "Falls Änderungen nicht erkannt werden:"

    $lblFixInfo.Location = New-Object System.Drawing.Point(10, 25)

    $lblFixInfo.Size = New-Object System.Drawing.Size(660, 20)

    

    $btnFixTimestamps = New-Object System.Windows.Forms.Button

    $btnFixTimestamps.Text = "⏱️ Zeitstempel aktualisieren"

    $btnFixTimestamps.Location = New-Object System.Drawing.Point(10, 50)

    $btnFixTimestamps.Size = New-Object System.Drawing.Size(200, 35)

    $btnFixTimestamps.BackColor = [System.Drawing.Color]::LightYellow

    

    $btnTestFile = New-Object System.Windows.Forms.Button

    $btnTestFile.Text = "📝 Test-Datei erstellen"

    $btnTestFile.Location = New-Object System.Drawing.Point(220, 50)

    $btnTestFile.Size = New-Object System.Drawing.Size(200, 35)

    $btnTestFile.BackColor = [System.Drawing.Color]::LightCyan

    

    $btnClearLastBackup = New-Object System.Windows.Forms.Button

    $btnClearLastBackup.Text = "🔄 Letztes Backup zurücksetzen"

    $btnClearLastBackup.Location = New-Object System.Drawing.Point(430, 50)

    $btnClearLastBackup.Size = New-Object System.Drawing.Size(200, 35)

    $btnClearLastBackup.BackColor = [System.Drawing.Color]::LightSalmon

    

    $lblFixResult = New-Object System.Windows.Forms.Label

    $lblFixResult.Location = New-Object System.Drawing.Point(10, 95)

    $lblFixResult.Size = New-Object System.Drawing.Size(660, 45)

    $lblFixResult.BorderStyle = "FixedSingle"

    $lblFixResult.BackColor = [System.Drawing.Color]::WhiteSmoke

    

    $grpQuickFix.Controls.AddRange(@($lblFixInfo, $btnFixTimestamps, $btnTestFile, $btnClearLastBackup, $lblFixResult))

    

    $grpSchedule = New-Object System.Windows.Forms.GroupBox

    $grpSchedule.Text = "Windows-Aufgabenplanung"

    $grpSchedule.Location = New-Object System.Drawing.Point(10, 170)

    $grpSchedule.Size = New-Object System.Drawing.Size(680, 100)

    

    $btnCreateTask = New-Object System.Windows.Forms.Button

    $btnCreateTask.Text = "📅 Windows-Aufgabe erstellen"

    $btnCreateTask.Location = New-Object System.Drawing.Point(10, 25)

    $btnCreateTask.Size = New-Object System.Drawing.Size(200, 35)

    $btnCreateTask.BackColor = [System.Drawing.Color]::LightBlue

    

    $btnOpenTaskScheduler = New-Object System.Windows.Forms.Button

    $btnOpenTaskScheduler.Text = "⚙️ Aufgabenplanung öffnen"

    $btnOpenTaskScheduler.Location = New-Object System.Drawing.Point(220, 25)

    $btnOpenTaskScheduler.Size = New-Object System.Drawing.Size(200, 35)

    

    $lblTaskInfo = New-Object System.Windows.Forms.Label

    $lblTaskInfo.Text = "Erstellt eine Windows-Aufgabe für automatische Backups (täglich oder wöchentlich)"

    $lblTaskInfo.Location = New-Object System.Drawing.Point(10, 65)

    $lblTaskInfo.Size = New-Object System.Drawing.Size(660, 25)

    

    $grpSchedule.Controls.AddRange(@($btnCreateTask, $btnOpenTaskScheduler, $lblTaskInfo))

    

    $grpHelp = New-Object System.Windows.Forms.GroupBox

    $grpHelp.Text = "Hilfe & Information"

    $grpHelp.Location = New-Object System.Drawing.Point(10, 280)

    $grpHelp.Size = New-Object System.Drawing.Size(680, 220)

    

    $txtHelp = New-Object System.Windows.Forms.TextBox

    $txtHelp.Multiline = $true

    $txtHelp.ScrollBars = "Vertical"

    $txtHelp.Location = New-Object System.Drawing.Point(10, 20)

    $txtHelp.Size = New-Object System.Drawing.Size(660, 190)

    $txtHelp.ReadOnly = $true

    $txtHelp.BackColor = [System.Drawing.Color]::White

    $txtHelp.Text = @"

OFFICE BACKUP SYSTEM v$script:Version - Hilfe



BACKUP-MODI:

• Smart: Automatische Entscheidung (Vollbackup alle 7 Tage)

• Vollständig: Sichert ALLE Dateien

• Inkrementell: Nur neue/geänderte Dateien



PROBLEMBEHEBUNG:

• Änderungen werden nicht erkannt?

  → "Zeitstempel aktualisieren" klicken

  → Oder "Backup erzwingen" aktivieren



• Backup dauert zu lange?

  → Nutze "Inkrementell" statt "Vollständig"

  → Reduziere die Anzahl der Quellordner



• Zu viel Speicherplatz belegt?

  → Lösche alte Backups im Historie-Tab

  → Nutze "Alte aufräumen" Button



TIPPS:

• Tägliches Vollbackup + stündlich inkrementell

• Wichtige Ordner separat öfter sichern

• Regelmäßig alte Backups löschen (>30 Tage)



Support: Office Backup System v$script:Version

"@

    

    $grpHelp.Controls.Add($txtHelp)

    

    $tabTools.Controls.AddRange(@($grpQuickFix, $grpSchedule, $grpHelp))

    

    # Tabs zum Control hinzufügen

    $tabControl.TabPages.AddRange(@($tabBackup, $tabHistory, $tabTools))

    

    # ================================================================

    # HAUPTBUTTONS

    # ================================================================

    

    $btnBackupNow = New-Object System.Windows.Forms.Button

    $btnBackupNow.Text = "▶️ BACKUP STARTEN"

    $btnBackupNow.Location = New-Object System.Drawing.Point(10, 570)

    $btnBackupNow.Size = New-Object System.Drawing.Size(170, 50)

    $btnBackupNow.BackColor = [System.Drawing.Color]::LightGreen

    $btnBackupNow.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)

    $btnBackupNow.FlatStyle = "Popup"

    

    $btnSaveConfig = New-Object System.Windows.Forms.Button

    $btnSaveConfig.Text = "💾 Konfig speichern"

    $btnSaveConfig.Location = New-Object System.Drawing.Point(190, 570)

    $btnSaveConfig.Size = New-Object System.Drawing.Size(130, 50)

    

    $btnLoadConfig = New-Object System.Windows.Forms.Button

    $btnLoadConfig.Text = "📂 Konfig laden"

    $btnLoadConfig.Location = New-Object System.Drawing.Point(330, 570)

    $btnLoadConfig.Size = New-Object System.Drawing.Size(130, 50)

    

    $btnAbout = New-Object System.Windows.Forms.Button

    $btnAbout.Text = "ℹ️ Über"

    $btnAbout.Location = New-Object System.Drawing.Point(470, 570)

    $btnAbout.Size = New-Object System.Drawing.Size(100, 50)

    

    $btnExit = New-Object System.Windows.Forms.Button

    $btnExit.Text = "❌ Beenden"

    $btnExit.Location = New-Object System.Drawing.Point(580, 570)

    $btnExit.Size = New-Object System.Drawing.Size(140, 50)

    $btnExit.BackColor = [System.Drawing.Color]::LightCoral

    

    # Status-Bar

    $statusBar = New-Object System.Windows.Forms.StatusStrip

    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel

    $statusLabel.Text = "Bereit"

    $statusLabel.Spring = $true

    $statusLabel.TextAlign = "MiddleLeft"

    $statusBar.Items.Add($statusLabel) | Out-Null

    

    # ================================================================

    # HILFSFUNKTIONEN FÜR GUI

    # ================================================================

    

    # Historie aktualisieren

    $script:RefreshHistory = {

        $listHistory.Items.Clear()

        $lblStats.Text = "Lade..."

        

        if (Test-Path $txtDest.Text) {

            $backups = Get-ChildItem -Path $txtDest.Text -Directory | 

                Where-Object { $_.Name -match "^Backup_" } |

                Sort-Object Name -Descending

            

            $totalSize = 0

            $totalFiles = 0

            $fullBackups = 0

            $incBackups = 0

            

            foreach ($backup in $backups) {

                try {

                    $files = @(Get-ChildItem -Path $backup.FullName -File -Recurse -ErrorAction SilentlyContinue)

                    $fileCount = $files.Count

                    $size = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB

                    

                    $type = ""

                    if ($backup.Name -match "_Full$") { 

                        $type = "[VOLL]"

                        $fullBackups++

                    } elseif ($backup.Name -match "_Inc$") { 

                        $type = "[INKR]"

                        $incBackups++

                    } else {

                        $type = "[BACKUP]"

                    }

                    

                    $dateStr = $backup.CreationTime.ToString('dd.MM.yyyy HH:mm')

                    $listItem = "{0,-20} {1,-8} {2,6:N0} Dateien, {3,8:N1} MB - {4}" -f $backup.Name.Substring(0, [Math]::Min(20, $backup.Name.Length)), $type, $fileCount, $size, $dateStr

                    $listHistory.Items.Add($listItem)

                    

                    $totalFiles += $fileCount

                    $totalSize += $size

                }

                catch {

                    $listHistory.Items.Add("$($backup.Name) - [FEHLER]")

                }

            }

            

            $oldestBackup = if ($backups.Count -gt 0) { $backups[-1].CreationTime } else { $null }

            $newestBackup = if ($backups.Count -gt 0) { $backups[0].CreationTime } else { $null }

            

            $lblStats.Text = @"

Anzahl Backups:     $($backups.Count) ($fullBackups vollständig, $incBackups inkrementell)

Gesamte Dateien:    $("{0:N0}" -f $totalFiles)

Gesamtgröße:        $("{0:N1}" -f $totalSize) MB

Ältestes Backup:    $(if ($oldestBackup) { $oldestBackup.ToString('dd.MM.yyyy HH:mm') } else { '-' })

Neuestes Backup:    $(if ($newestBackup) { $newestBackup.ToString('dd.MM.yyyy HH:mm') } else { '-' })

"@

        } else {

            $lblStats.Text = "Backup-Zielordner existiert nicht!"

        }

    }

    

    # Speicherplatz prüfen

    $script:CheckDiskSpace = {

        if (Test-Path $txtDest.Text) {

            try {

                $drive = (Get-Item $txtDest.Text).PSDrive.Name

                $disk = Get-PSDrive $drive -ErrorAction SilentlyContinue

                if ($disk) {

                    $freeGB = [Math]::Round($disk.Free / 1GB, 2)

                    $totalGB = [Math]::Round(($disk.Used + $disk.Free) / 1GB, 2)

                    $lblDestSpace.Text = "💾 Freier Speicher: $freeGB GB von $totalGB GB"

                    $lblDestSpace.ForeColor = if ($freeGB -lt 10) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }

                } else {

                    $lblDestSpace.Text = "Netzlaufwerk oder unbekannt"

                }

            } catch {

                $lblDestSpace.Text = "Speicherplatz unbekannt"

            }

        }

    }

    

    # ================================================================

    # EVENT HANDLER

    # ================================================================

    

    # Quellordner hinzufügen

    $btnAddSource.Add_Click({

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog

        $folderDialog.Description = "Wählen Sie einen Ordner zum Sichern"

        $folderDialog.ShowNewFolderButton = $false

        

        if ($folderDialog.ShowDialog() -eq "OK") {

            if ($txtSources.Text.Trim() -ne "") {

                $txtSources.Text += "`r`n"

            }

            $txtSources.Text += $folderDialog.SelectedPath

            $statusLabel.Text = "Ordner hinzugefügt: $($folderDialog.SelectedPath)"

        }

    })

    

    $btnClearSources.Add_Click({

        $txtSources.Text = ""

        $statusLabel.Text = "Quellordner gelöscht"

    })

    

    # Zielordner wählen

    $btnBrowseDest.Add_Click({

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog

        $folderDialog.Description = "Wählen Sie den Backup-Zielordner"

        

        if ($folderDialog.ShowDialog() -eq "OK") {

            $txtDest.Text = $folderDialog.SelectedPath

            & $script:CheckDiskSpace

            & $script:RefreshHistory

            $statusLabel.Text = "Zielordner geändert: $($folderDialog.SelectedPath)"

        }

    })

    

    # Dateitypen Buttons

    $btnSelectAll.Add_Click({

        $chkWord.Checked = $true

        $chkExcel.Checked = $true

        $chkPowerPoint.Checked = $true

        $chkPDF.Checked = $true

        $chkImages.Checked = $true

    })

    

    $btnSelectNone.Add_Click({

        $chkWord.Checked = $false

        $chkExcel.Checked = $false

        $chkPowerPoint.Checked = $false

        $chkPDF.Checked = $false

        $chkImages.Checked = $false

        $chkAll.Checked = $false

    })

    

    $chkAll.Add_CheckedChanged({

        if ($chkAll.Checked) {

            $result = [System.Windows.Forms.MessageBox]::Show(

                "WARNUNG: 'Alle Dateien' kann sehr lange dauern und viel Speicherplatz benötigen!`n`nWirklich fortfahren?",

                "Warnung",

                [System.Windows.Forms.MessageBoxButtons]::YesNo,

                [System.Windows.Forms.MessageBoxIcon]::Warning

            )

            if ($result -eq "No") {

                $chkAll.Checked = $false

            }

        }

    })

    

    # Historie Buttons

    $btnRefreshHistory.Add_Click($script:RefreshHistory)

    

    $btnOpenBackup.Add_Click({

        if ($listHistory.SelectedItem) {

            $backupName = $listHistory.SelectedItem.Split(' ')[0].Trim()

            $backupPath = Join-Path $txtDest.Text $backupName

            if (Test-Path $backupPath) {

                Start-Process explorer.exe $backupPath

                $statusLabel.Text = "Ordner geöffnet: $backupName"

            }

        } else {

            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie ein Backup aus!", "Hinweis", "OK", "Information")

        }

    })

    

    $btnDeleteBackup.Add_Click({

        if ($listHistory.SelectedItem) {

            $backupName = $listHistory.SelectedItem.Split(' ')[0].Trim()

            $result = [System.Windows.Forms.MessageBox]::Show(

                "Backup '$backupName' wirklich löschen?`n`nDieser Vorgang kann nicht rückgängig gemacht werden!",

                "Backup löschen",

                [System.Windows.Forms.MessageBoxButtons]::YesNo,

                [System.Windows.Forms.MessageBoxIcon]::Warning

            )

            

            if ($result -eq "Yes") {

                $backupPath = Join-Path $txtDest.Text $backupName

                try {

                    Remove-Item -Path $backupPath -Recurse -Force

                    & $script:RefreshHistory

                    $statusLabel.Text = "Backup gelöscht: $backupName"

                } catch {

                    [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen: $_", "Fehler", "OK", "Error")

                }

            }

        }

    })

    

    $btnCleanupOld.Add_Click({

        $result = [System.Windows.Forms.MessageBox]::Show(

            "Alle Backups älter als 30 Tage löschen?",

            "Alte Backups aufräumen",

            [System.Windows.Forms.MessageBoxButtons]::YesNo,

            [System.Windows.Forms.MessageBoxIcon]::Question

        )

        

        if ($result -eq "Yes") {

            $cutoffDate = (Get-Date).AddDays(-30)

            $deleted = 0

            

            Get-ChildItem -Path $txtDest.Text -Directory | 

                Where-Object { $_.Name -match "^Backup_" -and $_.CreationTime -lt $cutoffDate } |

                ForEach-Object {

                    Remove-Item $_.FullName -Recurse -Force

                    $deleted++

                }

            

            & $script:RefreshHistory

            $statusLabel.Text = "$deleted alte Backups gelöscht"

            [System.Windows.Forms.MessageBox]::Show("$deleted alte Backups wurden gelöscht!", "Aufräumen abgeschlossen", "OK", "Information")

        }

    })

    

    # Tools Buttons

    $btnFixTimestamps.Add_Click({

        $sourcePaths = $txtSources.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

        $updated = 0

        

        foreach ($path in $sourcePaths) {

            if (Test-Path $path) {

                $files = Get-ChildItem -Path $path -Include "*.docx","*.xlsx","*.pptx","*.pdf" -Recurse -File | Select-Object -First 10

                foreach ($file in $files) {

                    $file.LastWriteTime = Get-Date

                    $updated++

                }

            }

        }

        

        $lblFixResult.Text = "✓ $updated Dateien aktualisiert!`nDiese werden beim nächsten Backup als 'geändert' erkannt."

        $lblFixResult.ForeColor = [System.Drawing.Color]::Green

        $statusLabel.Text = "Zeitstempel von $updated Dateien aktualisiert"

    })

    

    $btnTestFile.Add_Click({

        $sourcePaths = $txtSources.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

        if ($sourcePaths.Count -gt 0) {

            $testPath = $sourcePaths[0]

            if (Test-Path $testPath) {

                $testFile = Join-Path $testPath "BACKUP_TEST_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

                "Test-Datei erstellt am $(Get-Date)`nDiese Datei sollte beim nächsten Backup gefunden werden!" | Out-File $testFile

                $lblFixResult.Text = "✓ Test-Datei erstellt:`n$testFile"

                $lblFixResult.ForeColor = [System.Drawing.Color]::Blue

                $statusLabel.Text = "Test-Datei erstellt"

            }

        }

    })

    

    $btnClearLastBackup.Add_Click({

        $lastBackup = Get-LastBackupFolder -BackupRoot $txtDest.Text

        if ($lastBackup) {

            $result = [System.Windows.Forms.MessageBox]::Show(

                "Letztes Backup zurücksetzen?`n`n$lastBackup`n`nDas nächste Backup wird dann als Vollbackup durchgeführt.",

                "Zurücksetzen",

                [System.Windows.Forms.MessageBoxButtons]::YesNo,

                [System.Windows.Forms.MessageBoxIcon]::Question

            )

            

            if ($result -eq "Yes") {

                Remove-Item $lastBackup -Recurse -Force

                & $script:RefreshHistory

                $lblFixResult.Text = "✓ Letztes Backup zurückgesetzt!`nNächstes Backup wird vollständig sein."

                $lblFixResult.ForeColor = [System.Drawing.Color]::Orange

            }

        } else {

            $lblFixResult.Text = "Kein Backup zum Zurücksetzen gefunden."

        }

    })

    

    $btnCreateTask.Add_Click({

        $taskName = "Office-Backup-System"

        $scriptPath = $MyInvocation.MyCommand.Path

        

        if (-not $scriptPath) {

            [System.Windows.Forms.MessageBox]::Show(

                "Bitte speichern Sie das Skript zuerst!",

                "Fehler",

                "OK",

                "Error"

            )

            return

        }

        

        try {

            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `

                -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -Silent"

            

            $trigger = New-ScheduledTaskTrigger -Daily -At "14:00"

            

            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `

                -DontStopIfGoingOnBatteries -StartWhenAvailable

            

            Register-ScheduledTask -TaskName $taskName -Action $action `

                -Trigger $trigger -Settings $settings -Force

            

            [System.Windows.Forms.MessageBox]::Show(

                "Windows-Aufgabe '$taskName' wurde erstellt!`n`nTäglich um 14:00 Uhr",

                "Erfolg",

                "OK",

                "Information"

            )

            $statusLabel.Text = "Windows-Aufgabe erstellt"

        }

        catch {

            [System.Windows.Forms.MessageBox]::Show(

                "Fehler beim Erstellen der Aufgabe:`n$_",

                "Fehler",

                "OK",

                "Error"

            )

        }

    })

    

    $btnOpenTaskScheduler.Add_Click({

        Start-Process taskschd.msc

        $statusLabel.Text = "Aufgabenplanung geöffnet"

    })

    

    # Hauptbuttons

    $btnBackupNow.Add_Click({

        $statusLabel.Text = "Backup wird vorbereitet..."

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $form.Refresh()

        

        # Konfiguration sammeln

        $sourcePaths = $txtSources.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

        

        if ($sourcePaths.Count -eq 0) {

            [System.Windows.Forms.MessageBox]::Show(

                "Bitte geben Sie mindestens einen Quellordner an!",

                "Fehler",

                "OK",

                "Error"

            )

            $form.Cursor = [System.Windows.Forms.Cursors]::Default

            return

        }

        

        $destPath = $txtDest.Text

        if (-not (Test-Path $destPath)) {

            $create = [System.Windows.Forms.MessageBox]::Show(

                "Zielordner existiert nicht. Erstellen?",

                "Ordner erstellen",

                [System.Windows.Forms.MessageBoxButtons]::YesNo,

                [System.Windows.Forms.MessageBoxIcon]::Question

            )

            

            if ($create -eq "Yes") {

                New-Item -ItemType Directory -Path $destPath -Force | Out-Null

            } else {

                $form.Cursor = [System.Windows.Forms.Cursors]::Default

                return

            }

        }

        

        $fileTypes = @{

            Word = $chkWord.Checked

            Excel = $chkExcel.Checked

            PowerPoint = $chkPowerPoint.Checked

            PDF = $chkPDF.Checked

            Images = $chkImages.Checked

            All = $chkAll.Checked

        }

        

        $backupMode = if ($radioSmart.Checked) { "Smart" } 

                     elseif ($radioFull.Checked) { "Full" } 

                     else { "Incremental" }

        

        $forceBackup = $chkForceBackup.Checked

        

        # Backup durchführen

        $statusLabel.Text = "Backup läuft..."

        $result = Start-BackupProcess -SourcePaths $sourcePaths `

            -DestinationPath $destPath `

            -FileTypes $fileTypes `

            -BackupMode $backupMode `

            -ForceBackup $forceBackup

        

        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        

        # Ergebnis anzeigen

        if ($result.Success) {

            if ($result.BackupNeeded) {

                $statusLabel.Text = "✓ Backup erfolgreich! $($result.CopiedFiles) Dateien gesichert."

                

                [System.Windows.Forms.MessageBox]::Show(

                    "Backup erfolgreich abgeschlossen!`n`n" +

                    "Gesicherte Dateien: $($result.CopiedFiles)`n" +

                    "  → Neu: $($result.NewFiles)`n" +

                    "  → Geändert: $($result.ModifiedFiles)`n" +

                    "Übersprungen: $($result.SkippedFiles)`n`n" +

                    "Backup-Ordner:`n$($result.BackupPath)",

                    "Backup erfolgreich",

                    [System.Windows.Forms.MessageBoxButtons]::OK,

                    [System.Windows.Forms.MessageBoxIcon]::Information

                )

            } else {

                $statusLabel.Text = "ℹ️ Keine Änderungen gefunden."

                

                [System.Windows.Forms.MessageBox]::Show(

                    "Keine neuen oder geänderten Dateien gefunden.`n`n" +

                    "Geprüfte Dateien: $($result.TotalFiles)`n`n" +

                    "Alle Dateien sind bereits gesichert.`n`n" +

                    "💡 Tipp: Nutzen Sie 'Backup erzwingen' oder 'Vollständig'`n" +

                    "wenn Sie trotzdem ein Backup erstellen möchten.",

                    "Keine Änderungen",

                    [System.Windows.Forms.MessageBoxButtons]::OK,

                    [System.Windows.Forms.MessageBoxIcon]::Information

                )

            }

        } else {

            $statusLabel.Text = "✗ Backup fehlgeschlagen!"

            

            $errorMsg = if ($result.Errors.Count -gt 0) {

                "Fehler:`n" + ($result.Errors | Select-Object -First 5 | Out-String)

            } else {

                "Ein unbekannter Fehler ist aufgetreten."

            }

            

            [System.Windows.Forms.MessageBox]::Show(

                "Backup fehlgeschlagen!`n`n$errorMsg",

                "Fehler",

                [System.Windows.Forms.MessageBoxButtons]::OK,

                [System.Windows.Forms.MessageBoxIcon]::Error

            )

        }

        

        # Historie aktualisieren

        & $script:RefreshHistory

        & $script:CheckDiskSpace

    })

    

    $btnSaveConfig.Add_Click({

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog

        $saveDialog.Filter = "JSON-Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"

        $saveDialog.FileName = "backup_config_$(Get-Date -Format 'yyyyMMdd').json"

        

        if ($saveDialog.ShowDialog() -eq "OK") {

            $config = @{

                Version = $script:Version

                SourcePaths = $txtSources.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

                DestinationPath = $txtDest.Text

                FileTypes = @{

                    Word = $chkWord.Checked

                    Excel = $chkExcel.Checked

                    PowerPoint = $chkPowerPoint.Checked

                    PDF = $chkPDF.Checked

                    Images = $chkImages.Checked

                    All = $chkAll.Checked

                }

                BackupMode = if ($radioSmart.Checked) { "Smart" } 

                            elseif ($radioFull.Checked) { "Full" } 

                            else { "Incremental" }

                ForceBackup = $chkForceBackup.Checked

                AutoBackup = $chkAutoBackup.Checked

                BackupInterval = $numInterval.Value

            }

            

            $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $saveDialog.FileName -Encoding UTF8

            $statusLabel.Text = "Konfiguration gespeichert: $(Split-Path $saveDialog.FileName -Leaf)"

            

            [System.Windows.Forms.MessageBox]::Show(

                "Konfiguration wurde gespeichert!",

                "Gespeichert",

                "OK",

                "Information"

            )

        }

    })

    

    $btnLoadConfig.Add_Click({

        $openDialog = New-Object System.Windows.Forms.OpenFileDialog

        $openDialog.Filter = "JSON-Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"

        

        if ($openDialog.ShowDialog() -eq "OK") {

            try {

                $config = Get-Content -Path $openDialog.FileName -Raw | ConvertFrom-Json

                

                # Konfiguration laden

                $txtSources.Text = $config.SourcePaths -join "`r`n"

                $txtDest.Text = $config.DestinationPath

                $chkWord.Checked = $config.FileTypes.Word

                $chkExcel.Checked = $config.FileTypes.Excel

                $chkPowerPoint.Checked = $config.FileTypes.PowerPoint

                $chkPDF.Checked = $config.FileTypes.PDF

                $chkImages.Checked = $config.FileTypes.Images

                $chkAll.Checked = $config.FileTypes.All

                

                switch ($config.BackupMode) {

                    "Smart" { $radioSmart.Checked = $true }

                    "Full" { $radioFull.Checked = $true }

                    "Incremental" { $radioIncremental.Checked = $true }

                }

                

                if ($config.PSObject.Properties.Name -contains 'ForceBackup') {

                    $chkForceBackup.Checked = $config.ForceBackup

                }

                

                $chkAutoBackup.Checked = $config.AutoBackup

                $numInterval.Value = $config.BackupInterval

                

                & $script:RefreshHistory

                & $script:CheckDiskSpace

                

                $statusLabel.Text = "Konfiguration geladen: $(Split-Path $openDialog.FileName -Leaf)"

                

                [System.Windows.Forms.MessageBox]::Show(

                    "Konfiguration wurde geladen!",

                    "Geladen",

                    "OK",

                    "Information"

                )

            }

            catch {

                [System.Windows.Forms.MessageBox]::Show(

                    "Fehler beim Laden der Konfiguration:`n$_",

                    "Fehler",

                    "OK",

                    "Error"

                )

            }

        }

    })

    

    $btnAbout.Add_Click({

        [System.Windows.Forms.MessageBox]::Show(

            "OFFICE BACKUP SYSTEM`nVersion $script:Version`n`n" +

            "Ein professionelles Backup-System für Office-Dateien`n" +

            "mit intelligenter Änderungserkennung.`n`n" +

            "Features:`n" +

            "• Smart-Mode mit automatischer Entscheidung`n" +

            "• Inkrementelle Backups`n" +

            "• Automatische Sortierung nach Dateitypen`n" +

            "• Quick-Fix Tools für Problembehebung`n" +

            "• Windows-Aufgabenplanung Integration`n`n" +

            "© 2024 Office Backup System",

            "Über",

            "OK",

            "Information"

        )

    })

    

    $btnExit.Add_Click({

        if ($timer.Enabled) {

            $result = [System.Windows.Forms.MessageBox]::Show(

                "Automatisches Backup ist aktiv. Wirklich beenden?",

                "Beenden",

                [System.Windows.Forms.MessageBoxButtons]::YesNo,

                [System.Windows.Forms.MessageBoxIcon]::Question

            )

            

            if ($result -eq "No") { return }

        }

        $form.Close()

    })

    

    # Timer für automatisches Backup

    $timer = New-Object System.Windows.Forms.Timer

    $timer.Add_Tick({

        if ($chkAutoBackup.Checked) {

            $statusLabel.Text = "Automatisches Backup läuft..."

            

            $sourcePaths = $txtSources.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }

            $destPath = $txtDest.Text

            

            $fileTypes = @{

                Word = $chkWord.Checked

                Excel = $chkExcel.Checked

                PowerPoint = $chkPowerPoint.Checked

                PDF = $chkPDF.Checked

                Images = $chkImages.Checked

                All = $chkAll.Checked

            }

            

            # Auto-Backup nutzt immer Smart-Mode

            $result = Start-BackupProcess -SourcePaths $sourcePaths `

                -DestinationPath $destPath `

                -FileTypes $fileTypes `

                -BackupMode "Smart"

            

            if ($result.BackupNeeded) {

                $statusLabel.Text = "Auto-Backup: $($result.CopiedFiles) Dateien gesichert"

                

                # Balloon-Tipp anzeigen (wenn möglich)

                try {

                    $notification = New-Object System.Windows.Forms.NotifyIcon

                    $notification.Icon = [System.Drawing.SystemIcons]::Information

                    $notification.BalloonTipTitle = "Backup abgeschlossen"

                    $notification.BalloonTipText = "$($result.CopiedFiles) Dateien wurden gesichert"

                    $notification.Visible = $true

                    $notification.ShowBalloonTip(5000)

                    Start-Sleep -Seconds 5

                    $notification.Dispose()

                } catch {}

            } else {

                $statusLabel.Text = "Auto-Backup: Keine Änderungen"

            }

            

            & $script:RefreshHistory

        }

    })

    

    $chkAutoBackup.Add_CheckedChanged({

        if ($chkAutoBackup.Checked) {

            $timer.Interval = [int]($numInterval.Value * 60 * 1000)

            $timer.Start()

            $statusLabel.Text = "✓ Auto-Backup aktiviert (alle $($numInterval.Value) Minuten)"

        } else {

            $timer.Stop()

            $statusLabel.Text = "Auto-Backup deaktiviert"

        }

    })

    

    $numInterval.Add_ValueChanged({

        if ($timer.Enabled) {

            $timer.Interval = [int]($numInterval.Value * 60 * 1000)

            $statusLabel.Text = "Intervall geändert: $($numInterval.Value) Minuten"

        }

    })

    

    # Controls zum Formular hinzufügen

    $form.Controls.AddRange(@(

        $tabControl,

        $btnBackupNow, $btnSaveConfig, $btnLoadConfig, $btnAbout, $btnExit,

        $statusBar

    ))

    

    # Beim Start

    $form.Add_Shown({

        & $script:RefreshHistory

        & $script:CheckDiskSpace

        $statusLabel.Text = "Bereit - Version $script:Version"

    })

    

    # Formular anzeigen

    $form.ShowDialog() | Out-Null

    

    # Aufräumen

    if ($timer) { $timer.Dispose() }

}



# ================================================================

# HAUPTPROGRAMM

# ================================================================



# Prüfe ob als Administrator

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")



# Silent Mode für automatische Ausführung

if ($args -contains "-Silent") {

    # Konfiguration laden

    $configPath = Join-Path $PSScriptRoot "backup_config.json"

    

    if (Test-Path $configPath) {

        try {

            $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

            

            $backupMode = if ($config.PSObject.Properties.Name -contains 'BackupMode') { 

                $config.BackupMode 

            } else { 

                "Smart" 

            }

            

            Start-BackupProcess -SourcePaths $config.SourcePaths `

                -DestinationPath $config.DestinationPath `

                -FileTypes $config.FileTypes `

                -BackupMode $backupMode

        }

        catch {

            Write-Host "Fehler beim Laden der Konfiguration: $_" -ForegroundColor Red

        }

    } else {

        Write-Host "Keine Konfigurationsdatei gefunden!" -ForegroundColor Red

        Write-Host "Bitte führen Sie das Programm zuerst im GUI-Modus aus und speichern Sie eine Konfiguration." -ForegroundColor Yellow

    }

} else {

    # GUI-Modus

    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan

    Write-Host " OFFICE BACKUP SYSTEM v$script:Version" -ForegroundColor Green

    Write-Host ("=" * 60) -ForegroundColor Cyan

    Write-Host ""

    

    if (-not $isAdmin) {

        Write-Host " ⚠️  Läuft ohne Administrator-Rechte" -ForegroundColor Yellow

        Write-Host "    Einige Funktionen könnten eingeschränkt sein" -ForegroundColor Gray

        Write-Host ""

    }

    

    Write-Host " Lade Benutzeroberfläche..." -ForegroundColor White

    Write-Host ""

    

    # GUI starten

    Show-BackupGUI

    

    Write-Host "`n Programm beendet." -ForegroundColor Gray

    Write-Host ("=" * 60) -ForegroundColor Cyan

}



# Ende des Skripts