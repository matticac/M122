# ================================================================
# Office Backup System - Final Edition v3.3 (Stability Update)
# ================================================================
# Fix: Verhindert das erneute Kopieren aller Dateien bei Inkrementellen Backups
# Fix: Erhöhte Zeittoleranz für OneDrive/Cloud-Ordner
# Autor: Backup System Generator & M122
# Version: 3.3 Stability
# ================================================================

# .NET-Assemblies für GUI laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Konfiguration
$script:Version = "3.3 Stability"
$script:BackupConfig = @{
    SourcePaths = @()
    DestinationPath = ""
    ExcludePaths = @("TBZ") 
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
    BackupMode = "Smart"
    ForceTimeCheck = $true
    LastBackupPath = ""
}

# ================================================================
# HILFSFUNKTIONEN
# ================================================================

function Update-FileTimestamp {
    param([string]$FilePath)
    try {
        $file = Get-Item $FilePath
        $file.LastWriteTime = Get-Date
        return $true
    } catch { return $false }
}

# FIX: Funktion sucht jetzt gezielt nach dem letzten VOLL-Backup für den Vergleich
function Get-LastFullBackupFolder {
    param([string]$BackupRoot)
    if (Test-Path $BackupRoot) {
        # Suche nach Ordnern, die auf "_Full" enden
        $lastBackup = Get-ChildItem -Path $BackupRoot -Directory | 
            Where-Object { $_.Name -match "_Full$" } | 
            Sort-Object Name -Descending | Select-Object -First 1
        if ($lastBackup) { return $lastBackup.FullName }
    }
    return $null
}

# Allgemeine Funktion für das absolut letzte Backup (für Smart Mode Entscheidung)
function Get-AbsoluteLastBackupFolder {
    param([string]$BackupRoot)
    if (Test-Path $BackupRoot) {
        $lastBackup = Get-ChildItem -Path $BackupRoot -Directory | 
            Where-Object { $_.Name -match "^Backup_" } | 
            Sort-Object Name -Descending | Select-Object -First 1
        if ($lastBackup) { return $lastBackup.FullName }
    }
    return $null
}

function Test-FileNeedsBackup {
    param(
        [System.IO.FileInfo]$File,
        [string]$LastBackupPath,
        [string]$Category,
        [string]$RelativePath,
        [bool]$ForceCheck = $false
    )
    
    # Wenn kein Referenz-Backup existiert -> Neu sichern
    if (-not $LastBackupPath) { return @{NeedsBackup = $true; Reason = "Kein Basis-Backup"} }
    
    $backupFilePath = Join-Path $LastBackupPath $Category
    if ($RelativePath) { $backupFilePath = Join-Path $backupFilePath $RelativePath }
    $backupFilePath = Join-Path $backupFilePath $File.Name
    
    # Datei nicht im Vollbackup gefunden -> Neu
    if (-not (Test-Path $backupFilePath)) { return @{NeedsBackup = $true; Reason = "Neue Datei"} }
    
    $backupFile = Get-Item $backupFilePath
    
    # Größenvergleich
    if ($File.Length -ne $backupFile.Length) { return @{NeedsBackup = $true; Reason = "Größe geändert"} }

    # Zeitvergleich
    # FIX: Toleranz auf 3 Sekunden erhöht (hilft bei OneDrive/Windows Zeitunterschieden)
    $timeDiff = [Math]::Abs(($File.LastWriteTime - $backupFile.LastWriteTime).TotalSeconds)
    if ($timeDiff -gt 3) { return @{NeedsBackup = $true; Reason = "Zeitstempel geändert (${timeDiff}s)"} }
    
    return @{NeedsBackup = $false; Reason = "Unverändert"}
}

# ================================================================
# HAUPTFUNKTION: BACKUP-PROZESS
# ================================================================

function Start-BackupProcess {
    param(
        [string[]]$SourcePaths,
        [string]$DestinationPath,
        [string[]]$ExcludePaths, 
        [hashtable]$FileTypes,
        [string]$BackupMode = "Smart",
        [bool]$ForceBackup = $false
    )
    
    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
    Write-Host " BACKUP-PROZESS GESTARTET" -ForegroundColor Green
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "Zeit: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" -ForegroundColor Yellow
    Write-Host "Modus: $BackupMode" -ForegroundColor Yellow
    
    # Ausnahmen bereinigen
    $cleanExcludes = $ExcludePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($cleanExcludes.Count -gt 0) {
        Write-Host "Ausnahmen (ignoriere Ordner):" -ForegroundColor Magenta
        $cleanExcludes | ForEach-Object { Write-Host "  ⛔ $_" -ForegroundColor Magenta }
    }

    # Smart Mode Logik: Entscheidet ob wir Full brauchen
    if ($BackupMode -eq "Smart") {
        $lastFull = Get-LastFullBackupFolder -BackupRoot $DestinationPath
        if ($lastFull) {
            try {
                # Datum aus Ordnernamen extrahieren (Backup_YYYY-MM-DD_HHmmss_Full)
                $dateStr = ($lastFull | Split-Path -Leaf).Substring(7, 17).Replace('_', ' ')
                $lastBackupDate = [DateTime]::ParseExact($dateStr, 'yyyy-MM-dd HHmmss', $null)
                
                # Wenn letztes Vollbackup älter als 7 Tage -> Neues Vollbackup
                if ((Get-Date).Subtract($lastBackupDate).TotalDays -gt 7) { 
                    $BackupMode = "Full" 
                    Write-Host "Smart-Check: Letztes Vollbackup > 7 Tage alt -> Mache Vollbackup" -ForegroundColor Gray
                } else { 
                    $BackupMode = "Incremental" 
                    Write-Host "Smart-Check: Basis ist aktuell -> Mache Inkrementell" -ForegroundColor Gray
                }
            } catch { $BackupMode = "Full" }
        } else { 
            $BackupMode = "Full" 
            Write-Host "Smart-Check: Kein Vollbackup gefunden -> Mache Vollbackup" -ForegroundColor Gray
        }
    }
    
    # Referenz-Pfad finden (Womit vergleichen wir?)
    $referencePath = $null
    
    if ($BackupMode -eq "Incremental" -and -not $ForceBackup) {
        # FIX: Wir vergleichen IMMER mit dem letzten FULL Backup, nicht dem letzten Inkrementellen.
        # Das verhindert den "Ping-Pong" Effekt.
        $referencePath = Get-LastFullBackupFolder -BackupRoot $DestinationPath
        
        if (-not $referencePath) { 
            Write-Host "Kein Basis-Vollbackup gefunden -> Erzwinge Vollbackup" -ForegroundColor Red
            $BackupMode = "Full" 
        } else {
            Write-Host "Vergleiche Änderungen seit: $(Split-Path $referencePath -Leaf)" -ForegroundColor Cyan
        }
    }
    
    $backupDate = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $typeSuffix = if ($BackupMode -eq "Full") { "Full" } elseif ($BackupMode -eq "Incremental") { "Inc" } else { "Bkp" }
    $backupRoot = Join-Path $DestinationPath "Backup_${backupDate}_${typeSuffix}"
    
    # Erweiterungen
    $extensions = @{}
    if ($FileTypes.Word) { $extensions['Word'] = @('*.docx', '*.doc', '*.rtf') }
    if ($FileTypes.Excel) { $extensions['Excel'] = @('*.xlsx', '*.xls', '*.csv') }
    if ($FileTypes.PowerPoint) { $extensions['PowerPoint'] = @('*.pptx', '*.ppt') }
    if ($FileTypes.PDF) { $extensions['PDF'] = @('*.pdf') }
    if ($FileTypes.Images) { $extensions['Images'] = @('*.jpg', '*.png', '*.jpeg') }
    if ($FileTypes.All) { $extensions['All'] = @('*.*') }
    
    $stats = @{ TotalFiles=0; CheckedFiles=0; CopiedFiles=0; SkippedFiles=0; ExcludedFiles=0; Errors=@() }
    $fileList = @()
    
    Write-Host "`nAnalysiere Dateien..." -ForegroundColor Cyan
    
    foreach ($pathRaw in $SourcePaths) {
        $sourcePath = $pathRaw.TrimEnd('\')
        if (-not (Test-Path $sourcePath)) { $stats.Errors += "Pfad fehlt: $sourcePath"; continue }
        
        Write-Host "  → Scanne: $sourcePath" -ForegroundColor Gray
        
        foreach ($category in $extensions.Keys) {
            foreach ($extension in $extensions[$category]) {
                try {
                    $files = Get-ChildItem -Path $sourcePath -Filter $extension -Recurse -File -ErrorAction SilentlyContinue
                    
                    foreach ($file in $files) {
                        $stats.TotalFiles++
                        
                        # Ausnahmen Prüfung
                        $isExcluded = $false
                        foreach ($excl in $cleanExcludes) {
                            if ($file.FullName -like "*\$excl\*" -or $file.FullName -like "$excl\*") {
                                $isExcluded = $true
                                break
                            }
                        }
                        if ($isExcluded) {
                            $stats.ExcludedFiles++; $stats.SkippedFiles++; continue
                        }

                        $relativePath = ""
                        if ($file.DirectoryName.StartsWith($sourcePath)) {
                            $relativePath = $file.DirectoryName.Substring($sourcePath.Length).TrimStart('\')
                        }
                        
                        if ($ForceBackup -or $BackupMode -eq "Full") {
                            $needsBackup = @{NeedsBackup = $true; Reason = "Vollbackup"}
                        } else {
                            # Hier nutzen wir nun den referencePath (Last Full Backup)
                            $needsBackup = Test-FileNeedsBackup -File $file -LastBackupPath $referencePath -Category $category -RelativePath $relativePath
                        }
                        
                        if ($needsBackup.NeedsBackup) {
                            $fileList += @{ File=$file; Category=$category; RelativePath=$relativePath; Reason=$needsBackup.Reason }
                        } else {
                            $stats.SkippedFiles++
                        }
                        $stats.CheckedFiles++
                    }
                } catch { $stats.Errors += "Scan-Fehler: $_" }
            }
        }
    }
    
    Write-Host "`nErgebnis:" -ForegroundColor Green
    Write-Host "  Ignoriert (Ausnahme): $($stats.ExcludedFiles)" -ForegroundColor Magenta
    Write-Host "  Zu sichern:           $($fileList.Count)" -ForegroundColor Cyan
    
    if ($fileList.Count -gt 0) {
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
        $prog = 0
        foreach ($item in $fileList) {
            $prog++
            try {
                $targetDir = Join-Path (Join-Path $backupRoot $item.Category) $item.RelativePath
                if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                Copy-Item -Path $item.File.FullName -Destination (Join-Path $targetDir $item.File.Name) -Force
                $stats.CopiedFiles++
                if ($fileList.Count -le 20 -or $prog % 10 -eq 0) { Write-Host "  ✓ $($item.File.Name)" -ForegroundColor Green }
            } catch { $stats.Errors += "Kopierfehler $($item.File.Name): $_" }
        }
        
        $logPath = Join-Path $backupRoot "backup_log.txt"
        "Backup v$script:Version`nDate: $(Get-Date)`nMode: $BackupMode`nRef: $(if($referencePath){$referencePath}else{'NONE'})`nFiles: $($stats.CopiedFiles)" | Out-File $logPath
        
        return @{ Success=$true; BackupNeeded=$true; CopiedFiles=$stats.CopiedFiles; BackupPath=$backupRoot; Errors=$stats.Errors }
    } else {
        return @{ Success=$true; BackupNeeded=$false; CopiedFiles=0; Errors=$stats.Errors }
    }
}

# ================================================================
# GUI INTERFACE
# ================================================================

function Show-BackupGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office Backup System v$script:Version - Stability Update"
    $form.Size = New-Object System.Drawing.Size(750, 780)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $tabControl = New-Object System.Windows.Forms.TabControl; $tabControl.Size="710,630"; $tabControl.Location="10,10"
    
    # --- TAB 1 ---
    $tabBackup = New-Object System.Windows.Forms.TabPage; $tabBackup.Text = "📁 Backup-Einstellungen"
    
    # 1. Quellordner
    $grpSource = New-Object System.Windows.Forms.GroupBox; $grpSource.Text="1. Quellordner"; $grpSource.Location="10,10"; $grpSource.Size="680,100"
    $txtSources = New-Object System.Windows.Forms.TextBox; $txtSources.Multiline=$true; $txtSources.ScrollBars="Vertical"; $txtSources.Location="10,20"; $txtSources.Size="540,65"; $txtSources.Text="C:\Users\matti"
    $btnAddSource = New-Object System.Windows.Forms.Button; $btnAddSource.Text="Hinzufügen..."; $btnAddSource.Location="560,20"; $btnAddSource.Size="110,30"
    $btnAddSource.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog()-eq"OK"){$txtSources.Text+="`r`n"+$dlg.SelectedPath} })
    $grpSource.Controls.AddRange(@($txtSources, $btnAddSource))
    
    # 2. Zielordner
    $grpDest = New-Object System.Windows.Forms.GroupBox; $grpDest.Text="2. Zielordner"; $grpDest.Location="10,115"; $grpDest.Size="680,60"
    $txtDest = New-Object System.Windows.Forms.TextBox; $txtDest.Location="10,25"; $txtDest.Size="540,25"; $txtDest.Text="C:\Backup"
    $btnBrowseDest = New-Object System.Windows.Forms.Button; $btnBrowseDest.Text="Durchsuchen..."; $btnBrowseDest.Location="560,23"; $btnBrowseDest.Size="110,27"
    $btnBrowseDest.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog()-eq"OK"){$txtDest.Text=$dlg.SelectedPath} })
    $grpDest.Controls.AddRange(@($txtDest, $btnBrowseDest))
    
    # 3. Ausnahmen
    $grpExcl = New-Object System.Windows.Forms.GroupBox; $grpExcl.Text="3. Ausnahmen (Ordnernamen ignorieren)"; $grpExcl.Location="10,180"; $grpExcl.Size="680,80"; $grpExcl.ForeColor="DarkRed"
    $txtExcl = New-Object System.Windows.Forms.TextBox; $txtExcl.Location="10,25"; $txtExcl.Size="540,25"; $txtExcl.Text="TBZ"
    $lblExclInfo = New-Object System.Windows.Forms.Label; $lblExclInfo.Text="Trennzeichen: Komma oder neue Zeile"; $lblExclInfo.Location="10,50"; $lblExclInfo.Size="500,20"; $lblExclInfo.ForeColor="Gray"
    $grpExcl.Controls.AddRange(@($txtExcl, $lblExclInfo))
    
    # 4. Dateitypen
    $grpTypes = New-Object System.Windows.Forms.GroupBox; $grpTypes.Text="4. Dateitypen"; $grpTypes.Location="10,265"; $grpTypes.Size="680,100"
    $chkWord=New-Object System.Windows.Forms.CheckBox; $chkWord.Text="Word"; $chkWord.Location="15,25"; $chkWord.Checked=$true
    $chkExcel=New-Object System.Windows.Forms.CheckBox; $chkExcel.Text="Excel"; $chkExcel.Location="230,25"; $chkExcel.Checked=$true
    $chkPPT=New-Object System.Windows.Forms.CheckBox; $chkPPT.Text="PowerPoint"; $chkPPT.Location="445,25"; $chkPPT.Checked=$true
    $chkPDF=New-Object System.Windows.Forms.CheckBox; $chkPDF.Text="PDF"; $chkPDF.Location="15,50"; $chkPDF.Checked=$true
    $chkImg=New-Object System.Windows.Forms.CheckBox; $chkImg.Text="Bilder"; $chkImg.Location="230,50"
    $chkAll=New-Object System.Windows.Forms.CheckBox; $chkAll.Text="ALLE Dateien"; $chkAll.Location="445,50"; $chkAll.ForeColor="DarkRed"
    $grpTypes.Controls.AddRange(@($chkWord, $chkExcel, $chkPPT, $chkPDF, $chkImg, $chkAll))
    
    # 5. Modus
    $grpMode = New-Object System.Windows.Forms.GroupBox; $grpMode.Text="5. Modus"; $grpMode.Location="10,370"; $grpMode.Size="680,80"
    $radioSmart=New-Object System.Windows.Forms.RadioButton; $radioSmart.Text="Smart"; $radioSmart.Location="15,25"; $radioSmart.Checked=$true
    $radioFull=New-Object System.Windows.Forms.RadioButton; $radioFull.Text="Vollständig"; $radioFull.Location="150,25"
    $radioInc=New-Object System.Windows.Forms.RadioButton; $radioInc.Text="Inkrementell"; $radioInc.Location="300,25"
    $chkForce=New-Object System.Windows.Forms.CheckBox; $chkForce.Text="Erzwingen"; $chkForce.Location="450,25"
    $grpMode.Controls.AddRange(@($radioSmart, $radioFull, $radioInc, $chkForce))
    
    # Auto Backup
    $chkAuto=New-Object System.Windows.Forms.CheckBox; $chkAuto.Text="Auto-Backup aktiv"; $chkAuto.Location="20,460"; $chkAuto.Size="200,20"
    $numInt=New-Object System.Windows.Forms.NumericUpDown; $numInt.Location="250,460"; $numInt.Value=60; $numInt.Minimum=5
    
    $tabBackup.Controls.AddRange(@($grpSource, $grpDest, $grpExcl, $grpTypes, $grpMode, $chkAuto, $numInt))
    
    # --- TAB 2 ---
    $tabHist = New-Object System.Windows.Forms.TabPage; $tabHist.Text = "📜 Historie"
    $lstHist = New-Object System.Windows.Forms.ListBox; $lstHist.Location="10,10"; $lstHist.Size="680,400"; $lstHist.Font="Consolas,9"
    $btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="Aktualisieren"; $btnRef.Location="10,420"; $btnRef.Size="120,30"
    $btnOpen = New-Object System.Windows.Forms.Button; $btnOpen.Text="Öffnen"; $btnOpen.Location="140,420"; $btnOpen.Size="120,30"
    
    $refreshHist = {
        $lstHist.Items.Clear()
        if (Test-Path $txtDest.Text) {
            Get-ChildItem $txtDest.Text -Directory | Where Name -match "^Backup_" | Sort Name -Descending | ForEach {
                $count = (Get-ChildItem $_.FullName -Recurse -File -EA SilentlyContinue).Count
                $type = if($_.Name -match "_Full$"){"[VOLL]"}elseif($_.Name -match "_Inc$"){"[INKR]"}else{""}
                $lstHist.Items.Add("$($_.Name) $type - $count Dateien")
            }
        }
    }
    $btnRef.Add_Click($refreshHist)
    $btnOpen.Add_Click({ if($lstHist.SelectedItem){ Start-Process explorer (Join-Path $txtDest.Text ($lstHist.SelectedItem -split " ")[0]) } })
    $tabHist.Controls.AddRange(@($lstHist, $btnRef, $btnOpen))
    $tabControl.TabPages.AddRange(@($tabBackup, $tabHist))
    
    # --- BUTTONS ---
    $btnStart = New-Object System.Windows.Forms.Button; $btnStart.Text="▶️ STARTEN"; $btnStart.BackColor="LightGreen"; $btnStart.Location="10,650"; $btnStart.Size="150,40"
    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text="Speichern"; $btnSave.Location="170,650"; $btnSave.Size="100,40"
    $btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text="Laden"; $btnLoad.Location="280,650"; $btnLoad.Size="100,40"
    $btnExit = New-Object System.Windows.Forms.Button; $btnExit.Text="Beenden"; $btnExit.Location="590,650"; $btnExit.Size="100,40"; $btnExit.BackColor="LightCoral"
    $status = New-Object System.Windows.Forms.Label; $status.Text="Bereit"; $status.Location="10,700"; $status.Size="700,20"
    
    $btnStart.Add_Click({
        $status.Text = "Backup läuft..."
        $form.Cursor = "WaitCursor"; $form.Refresh()
        $src = $txtSources.Text -split "`r`n" | Where { $_.Trim() }
        $excl = $txtExcl.Text -split "[,`r`n]" | Where { $_.Trim() } | ForEach { $_.Trim() }
        $types = @{ Word=$chkWord.Checked; Excel=$chkExcel.Checked; PowerPoint=$chkPPT.Checked; PDF=$chkPDF.Checked; Images=$chkImg.Checked; All=$chkAll.Checked }
        $mode = if($radioFull.Checked){"Full"}elseif($radioInc.Checked){"Incremental"}else{"Smart"}
        
        if (-not (Test-Path $txtDest.Text)) { New-Item -ItemType Directory -Path $txtDest.Text -Force | Out-Null }
        $res = Start-BackupProcess -SourcePaths $src -DestinationPath $txtDest.Text -ExcludePaths $excl -FileTypes $types -BackupMode $mode -ForceBackup $chkForce.Checked
        $form.Cursor = "Default"
        if ($res.Success) {
            $status.Text = "Fertig! $($res.CopiedFiles) Dateien kopiert."
            [System.Windows.Forms.MessageBox]::Show("Backup fertig!`nKopiert: $($res.CopiedFiles)", "Erfolg")
            & $refreshHist
        } else { $status.Text = "Fehler." }
    })
    
    $btnSave.Add_Click({ 
        $cfg = @{ Sources=$txtSources.Text; Dest=$txtDest.Text; Exclude=$txtExcl.Text; Interval=$numInt.Value }
        $cfg | ConvertTo-Json | Out-File "backup_config.json"
        $status.Text = "Gespeichert."
    })
    $btnLoad.Add_Click({
        if (Test-Path "backup_config.json") {
            $cfg = Get-Content "backup_config.json" -Raw | ConvertFrom-Json
            $txtSources.Text = $cfg.Sources; $txtDest.Text = $cfg.Dest
            if ($cfg.Exclude) { $txtExcl.Text = $cfg.Exclude }
            $numInt.Value = $cfg.Interval
            $status.Text = "Geladen."
        }
    })
    $btnExit.Add_Click({ $form.Close() })
    $form.Controls.AddRange(@($tabControl, $btnStart, $btnSave, $btnLoad, $btnExit, $status))
    $form.Add_Shown({ & $refreshHist })
    $form.ShowDialog() | Out-Null
}

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) { Write-Host "Hinweis: Keine Admin-Rechte." -F Yellow }
Show-BackupGUI