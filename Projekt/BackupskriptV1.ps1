# Office Backup System mit GUI
# Version 1.0
# Automatisches Backup-System für Office-Dateien mit Sortierung und optionaler Cloud-Synchronisation

# .NET-Assemblies für GUI laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Variablen
$script:BackupConfig = @{
    SourcePaths = @()
    DestinationPath = ""
    FileTypes = @{
        Word = $true
        Excel = $true
        PowerPoint = $true
        PDF = $true
        Images = $true
        All = $false
    }
    AutoBackup = $false
    BackupInterval = 60  # Minuten
    AzureEnabled = $false
    AzureStorageAccount = ""
    AzureContainer = ""
}

# Funktion für das Backup
function Start-BackupProcess {
    param(
        [string[]]$SourcePaths,
        [string]$DestinationPath,
        [hashtable]$FileTypes
    )
    
    Write-Host "=== Backup-Prozess gestartet ===" -ForegroundColor Green
    Write-Host "Zeitstempel: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
    
    # Backup-Ordner mit Zeitstempel erstellen
    $backupDate = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $backupRoot = Join-Path $DestinationPath "Backup_$backupDate"
    
    if (-not (Test-Path $backupRoot)) {
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    }
    
    # Dateierweiterungen definieren
    $extensions = @{}
    if ($FileTypes.Word) { $extensions['Word'] = @('*.docx', '*.doc', '*.docm', '*.dotx', '*.dotm') }
    if ($FileTypes.Excel) { $extensions['Excel'] = @('*.xlsx', '*.xls', '*.xlsm', '*.xlsb', '*.xltx', '*.xltm') }
    if ($FileTypes.PowerPoint) { $extensions['PowerPoint'] = @('*.pptx', '*.ppt', '*.pptm', '*.potx', '*.potm', '*.ppsx') }
    if ($FileTypes.PDF) { $extensions['PDF'] = @('*.pdf') }
    if ($FileTypes.Images) { $extensions['Images'] = @('*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff', '*.svg') }
    if ($FileTypes.All) { $extensions['All'] = @('*.*') }
    
    $totalFiles = 0
    $copiedFiles = 0
    $errors = @()
    
    foreach ($sourcePath in $SourcePaths) {
        if (-not (Test-Path $sourcePath)) {
            Write-Host "Warnung: Pfad '$sourcePath' existiert nicht!" -ForegroundColor Red
            $errors += "Pfad nicht gefunden: $sourcePath"
            continue
        }
        
        Write-Host "`nVerarbeite: $sourcePath" -ForegroundColor Cyan
        
        foreach ($category in $extensions.Keys) {
            $categoryPath = Join-Path $backupRoot $category
            
            foreach ($extension in $extensions[$category]) {
                # Dateien rekursiv suchen
                $files = Get-ChildItem -Path $sourcePath -Filter $extension -Recurse -File -ErrorAction SilentlyContinue
                
                foreach ($file in $files) {
                    $totalFiles++
                    
                    try {
                        # Zielordner erstellen falls nicht vorhanden
                        if (-not (Test-Path $categoryPath)) {
                            New-Item -ItemType Directory -Path $categoryPath -Force | Out-Null
                        }
                        
                        # Relativen Pfad beibehalten
                        $relativePath = $file.DirectoryName.Substring($sourcePath.Length).TrimStart('\')
                        $targetDir = Join-Path $categoryPath $relativePath
                        
                        if (-not (Test-Path $targetDir)) {
                            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                        }
                        
                        $targetFile = Join-Path $targetDir $file.Name
                        
                        # Datei kopieren
                        Copy-Item -Path $file.FullName -Destination $targetFile -Force
                        $copiedFiles++
                        
                        Write-Host "  ✓ $($file.Name)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "  ✗ Fehler bei $($file.Name): $_" -ForegroundColor Red
                        $errors += "Fehler bei $($file.Name): $_"
                    }
                }
            }
        }
    }
    
    # Zusammenfassung
    Write-Host "`n=== Backup-Zusammenfassung ===" -ForegroundColor Green
    Write-Host "Gesamt gefunden: $totalFiles Dateien" -ForegroundColor Yellow
    Write-Host "Erfolgreich kopiert: $copiedFiles Dateien" -ForegroundColor Green
    Write-Host "Backup-Ordner: $backupRoot" -ForegroundColor Cyan
    
    if ($errors.Count -gt 0) {
        Write-Host "`nFehler während des Backups:" -ForegroundColor Red
        $errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    }
    
    # Log-Datei erstellen
    $logFile = Join-Path $backupRoot "backup_log.txt"
    $logContent = @"
Backup-Log
==========
Datum: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quellpfade: $($SourcePaths -join ', ')
Zielpfad: $backupRoot
Gesamt: $totalFiles Dateien
Kopiert: $copiedFiles Dateien
Fehler: $($errors.Count)

Fehlerdetails:
$($errors -join "`n")
"@
    $logContent | Out-File -FilePath $logFile -Encoding UTF8
    
    return @{
        Success = $true
        BackupPath = $backupRoot
        TotalFiles = $totalFiles
        CopiedFiles = $copiedFiles
        Errors = $errors
    }
}

# Azure Upload Funktion
function Sync-ToAzure {
    param(
        [string]$LocalPath,
        [string]$StorageAccount,
        [string]$Container,
        [string]$SasToken
    )
    
    Write-Host "`n=== Azure-Synchronisation ===" -ForegroundColor Blue
    
    try {
        # Prüfe ob Azure CLI installiert ist
        $azVersion = az version 2>$null
        if (-not $azVersion) {
            Write-Host "Azure CLI nicht installiert. Installiere mit: winget install Microsoft.AzureCLI" -ForegroundColor Red
            return $false
        }
        
        # Upload mit azcopy (falls verfügbar) oder az storage
        $uploadCmd = "az storage blob upload-batch --source `"$LocalPath`" --destination `"$Container`" --account-name `"$StorageAccount`" --sas-token `"$SasToken`""
        
        Write-Host "Lade Dateien nach Azure hoch..." -ForegroundColor Yellow
        Invoke-Expression $uploadCmd
        
        Write-Host "✓ Azure-Upload erfolgreich!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Azure-Upload fehlgeschlagen: $_" -ForegroundColor Red
        return $false
    }
}

# GUI erstellen
function Show-BackupGUI {
    # Hauptfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office Backup System"
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(660, 480)
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    
    # === Tab 1: Backup-Einstellungen ===
    $tabBackup = New-Object System.Windows.Forms.TabPage
    $tabBackup.Text = "Backup-Einstellungen"
    
    # Quellpfade
    $lblSource = New-Object System.Windows.Forms.Label
    $lblSource.Text = "Quellpfade (einer pro Zeile):"
    $lblSource.Location = New-Object System.Drawing.Point(10, 10)
    $lblSource.Size = New-Object System.Drawing.Size(200, 20)
    
    $txtSources = New-Object System.Windows.Forms.TextBox
    $txtSources.Multiline = $true
    $txtSources.ScrollBars = "Vertical"
    $txtSources.Location = New-Object System.Drawing.Point(10, 35)
    $txtSources.Size = New-Object System.Drawing.Size(500, 80)
    $txtSources.Text = "C:\Users\$env:USERNAME\Documents`r`nC:\Users\$env:USERNAME\Desktop"
    
    $btnAddSource = New-Object System.Windows.Forms.Button
    $btnAddSource.Text = "Ordner hinzufügen..."
    $btnAddSource.Location = New-Object System.Drawing.Point(520, 35)
    $btnAddSource.Size = New-Object System.Drawing.Size(120, 30)
    
    # Zielpfad
    $lblDest = New-Object System.Windows.Forms.Label
    $lblDest.Text = "Backup-Zielpfad:"
    $lblDest.Location = New-Object System.Drawing.Point(10, 125)
    $lblDest.Size = New-Object System.Drawing.Size(200, 20)
    
    $txtDest = New-Object System.Windows.Forms.TextBox
    $txtDest.Location = New-Object System.Drawing.Point(10, 150)
    $txtDest.Size = New-Object System.Drawing.Size(500, 25)
    $txtDest.Text = "D:\Backups"
    
    $btnBrowseDest = New-Object System.Windows.Forms.Button
    $btnBrowseDest.Text = "Durchsuchen..."
    $btnBrowseDest.Location = New-Object System.Drawing.Point(520, 148)
    $btnBrowseDest.Size = New-Object System.Drawing.Size(120, 25)
    
    # Dateitypen
    $grpFileTypes = New-Object System.Windows.Forms.GroupBox
    $grpFileTypes.Text = "Dateitypen für Backup"
    $grpFileTypes.Location = New-Object System.Drawing.Point(10, 185)
    $grpFileTypes.Size = New-Object System.Drawing.Size(630, 100)
    
    $chkWord = New-Object System.Windows.Forms.CheckBox
    $chkWord.Text = "Word-Dateien"
    $chkWord.Location = New-Object System.Drawing.Point(10, 25)
    $chkWord.Size = New-Object System.Drawing.Size(150, 20)
    $chkWord.Checked = $true
    
    $chkExcel = New-Object System.Windows.Forms.CheckBox
    $chkExcel.Text = "Excel-Dateien"
    $chkExcel.Location = New-Object System.Drawing.Point(170, 25)
    $chkExcel.Size = New-Object System.Drawing.Size(150, 20)
    $chkExcel.Checked = $true
    
    $chkPowerPoint = New-Object System.Windows.Forms.CheckBox
    $chkPowerPoint.Text = "PowerPoint-Dateien"
    $chkPowerPoint.Location = New-Object System.Drawing.Point(330, 25)
    $chkPowerPoint.Size = New-Object System.Drawing.Size(150, 20)
    $chkPowerPoint.Checked = $true
    
    $chkPDF = New-Object System.Windows.Forms.CheckBox
    $chkPDF.Text = "PDF-Dateien"
    $chkPDF.Location = New-Object System.Drawing.Point(10, 55)
    $chkPDF.Size = New-Object System.Drawing.Size(150, 20)
    $chkPDF.Checked = $true
    
    $chkImages = New-Object System.Windows.Forms.CheckBox
    $chkImages.Text = "Bilder"
    $chkImages.Location = New-Object System.Drawing.Point(170, 55)
    $chkImages.Size = New-Object System.Drawing.Size(150, 20)
    
    $chkAll = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text = "Alle Dateien"
    $chkAll.Location = New-Object System.Drawing.Point(330, 55)
    $chkAll.Size = New-Object System.Drawing.Size(150, 20)
    
    $grpFileTypes.Controls.AddRange(@($chkWord, $chkExcel, $chkPowerPoint, $chkPDF, $chkImages, $chkAll))
    
    # Automatisches Backup
    $chkAutoBackup = New-Object System.Windows.Forms.CheckBox
    $chkAutoBackup.Text = "Automatisches Backup aktivieren"
    $chkAutoBackup.Location = New-Object System.Drawing.Point(10, 295)
    $chkAutoBackup.Size = New-Object System.Drawing.Size(250, 20)
    
    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Text = "Intervall (Minuten):"
    $lblInterval.Location = New-Object System.Drawing.Point(270, 295)
    $lblInterval.Size = New-Object System.Drawing.Size(120, 20)
    
    $numInterval = New-Object System.Windows.Forms.NumericUpDown
    $numInterval.Location = New-Object System.Drawing.Point(390, 293)
    $numInterval.Size = New-Object System.Drawing.Size(80, 25)
    $numInterval.Minimum = 5
    $numInterval.Maximum = 1440
    $numInterval.Value = 60
    
    $tabBackup.Controls.AddRange(@(
        $lblSource, $txtSources, $btnAddSource,
        $lblDest, $txtDest, $btnBrowseDest,
        $grpFileTypes,
        $chkAutoBackup, $lblInterval, $numInterval
    ))
    
    # === Tab 2: Azure-Einstellungen ===
    $tabAzure = New-Object System.Windows.Forms.TabPage
    $tabAzure.Text = "Cloud-Synchronisation"
    
    $chkAzure = New-Object System.Windows.Forms.CheckBox
    $chkAzure.Text = "Azure-Synchronisation aktivieren"
    $chkAzure.Location = New-Object System.Drawing.Point(10, 10)
    $chkAzure.Size = New-Object System.Drawing.Size(250, 20)
    
    $lblStorageAccount = New-Object System.Windows.Forms.Label
    $lblStorageAccount.Text = "Storage Account Name:"
    $lblStorageAccount.Location = New-Object System.Drawing.Point(10, 40)
    $lblStorageAccount.Size = New-Object System.Drawing.Size(150, 20)
    
    $txtStorageAccount = New-Object System.Windows.Forms.TextBox
    $txtStorageAccount.Location = New-Object System.Drawing.Point(10, 65)
    $txtStorageAccount.Size = New-Object System.Drawing.Size(400, 25)
    
    $lblContainer = New-Object System.Windows.Forms.Label
    $lblContainer.Text = "Container Name:"
    $lblContainer.Location = New-Object System.Drawing.Point(10, 100)
    $lblContainer.Size = New-Object System.Drawing.Size(150, 20)
    
    $txtContainer = New-Object System.Windows.Forms.TextBox
    $txtContainer.Location = New-Object System.Drawing.Point(10, 125)
    $txtContainer.Size = New-Object System.Drawing.Size(400, 25)
    $txtContainer.Text = "office-backups"
    
    $lblSasToken = New-Object System.Windows.Forms.Label
    $lblSasToken.Text = "SAS Token (optional):"
    $lblSasToken.Location = New-Object System.Drawing.Point(10, 160)
    $lblSasToken.Size = New-Object System.Drawing.Size(150, 20)
    
    $txtSasToken = New-Object System.Windows.Forms.TextBox
    $txtSasToken.Location = New-Object System.Drawing.Point(10, 185)
    $txtSasToken.Size = New-Object System.Drawing.Size(600, 25)
    $txtSasToken.UseSystemPasswordChar = $true
    
    $lblAzureInfo = New-Object System.Windows.Forms.Label
    $lblAzureInfo.Text = @"
Hinweis: Für die Azure-Synchronisation benötigen Sie:
• Ein aktives Azure-Abonnement
• Einen Storage Account mit Blob-Container
• Azure CLI installiert (winget install Microsoft.AzureCLI)
• Optional: SAS Token für authentifizierten Zugriff

Die Synchronisation läuft nach jedem erfolgreichen Backup.
"@
    $lblAzureInfo.Location = New-Object System.Drawing.Point(10, 230)
    $lblAzureInfo.Size = New-Object System.Drawing.Size(600, 120)
    
    $tabAzure.Controls.AddRange(@(
        $chkAzure,
        $lblStorageAccount, $txtStorageAccount,
        $lblContainer, $txtContainer,
        $lblSasToken, $txtSasToken,
        $lblAzureInfo
    ))
    
    # === Tab 3: Zeitplan ===
    $tabSchedule = New-Object System.Windows.Forms.TabPage
    $tabSchedule.Text = "Zeitplan"
    
    $lblScheduleInfo = New-Object System.Windows.Forms.Label
    $lblScheduleInfo.Text = "Konfigurieren Sie automatische Backup-Zeiten:"
    $lblScheduleInfo.Location = New-Object System.Drawing.Point(10, 10)
    $lblScheduleInfo.Size = New-Object System.Drawing.Size(400, 20)
    
    $chkDaily = New-Object System.Windows.Forms.CheckBox
    $chkDaily.Text = "Tägliches Backup um:"
    $chkDaily.Location = New-Object System.Drawing.Point(10, 40)
    $chkDaily.Size = New-Object System.Drawing.Size(150, 20)
    
    $dtpTime = New-Object System.Windows.Forms.DateTimePicker
    $dtpTime.Location = New-Object System.Drawing.Point(170, 38)
    $dtpTime.Size = New-Object System.Drawing.Size(100, 25)
    $dtpTime.Format = "Time"
    $dtpTime.ShowUpDown = $true
    
    $chkWeekly = New-Object System.Windows.Forms.CheckBox
    $chkWeekly.Text = "Wöchentliches Backup"
    $chkWeekly.Location = New-Object System.Drawing.Point(10, 70)
    $chkWeekly.Size = New-Object System.Drawing.Size(200, 20)
    
    $grpWeekdays = New-Object System.Windows.Forms.GroupBox
    $grpWeekdays.Text = "Wochentage"
    $grpWeekdays.Location = New-Object System.Drawing.Point(10, 95)
    $grpWeekdays.Size = New-Object System.Drawing.Size(600, 60)
    
    $days = @("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So")
    $dayChecks = @()
    for ($i = 0; $i -lt $days.Count; $i++) {
        $dayCheck = New-Object System.Windows.Forms.CheckBox
        $dayCheck.Text = $days[$i]
        $dayCheck.Location = New-Object System.Drawing.Point((10 + $i * 85), 25)
        $dayCheck.Size = New-Object System.Drawing.Size(80, 20)
        if ($i -lt 5) { $dayCheck.Checked = $true }  # Mo-Fr standardmäßig aktiviert
        $dayChecks += $dayCheck
        $grpWeekdays.Controls.Add($dayCheck)
    }
    
    $btnCreateTask = New-Object System.Windows.Forms.Button
    $btnCreateTask.Text = "Windows-Aufgabe erstellen"
    $btnCreateTask.Location = New-Object System.Drawing.Point(10, 170)
    $btnCreateTask.Size = New-Object System.Drawing.Size(200, 30)
    $btnCreateTask.BackColor = [System.Drawing.Color]::LightBlue
    
    $lblTaskInfo = New-Object System.Windows.Forms.Label
    $lblTaskInfo.Text = @"
Die Windows-Aufgabe wird automatisch erstellt und führt dieses Skript
zu den konfigurierten Zeiten aus. Sie können die Aufgabe jederzeit
in der Windows-Aufgabenplanung (taskschd.msc) bearbeiten.
"@
    $lblTaskInfo.Location = New-Object System.Drawing.Point(10, 210)
    $lblTaskInfo.Size = New-Object System.Drawing.Size(600, 60)
    
    $tabSchedule.Controls.AddRange(@(
        $lblScheduleInfo,
        $chkDaily, $dtpTime,
        $chkWeekly, $grpWeekdays,
        $btnCreateTask,
        $lblTaskInfo
    ))
    
    # Tabs zum Control hinzufügen
    $tabControl.TabPages.AddRange(@($tabBackup, $tabAzure, $tabSchedule))
    
    # Buttons am unteren Rand
    $btnBackupNow = New-Object System.Windows.Forms.Button
    $btnBackupNow.Text = "Backup jetzt starten"
    $btnBackupNow.Location = New-Object System.Drawing.Point(10, 500)
    $btnBackupNow.Size = New-Object System.Drawing.Size(150, 40)
    $btnBackupNow.BackColor = [System.Drawing.Color]::LightGreen
    $btnBackupNow.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    
    $btnSaveConfig = New-Object System.Windows.Forms.Button
    $btnSaveConfig.Text = "Konfiguration speichern"
    $btnSaveConfig.Location = New-Object System.Drawing.Point(170, 500)
    $btnSaveConfig.Size = New-Object System.Drawing.Size(150, 40)
    
    $btnLoadConfig = New-Object System.Windows.Forms.Button
    $btnLoadConfig.Text = "Konfiguration laden"
    $btnLoadConfig.Location = New-Object System.Drawing.Point(330, 500)
    $btnLoadConfig.Size = New-Object System.Drawing.Size(150, 40)
    
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Beenden"
    $btnExit.Location = New-Object System.Drawing.Point(540, 500)
    $btnExit.Size = New-Object System.Drawing.Size(130, 40)
    
    # Status-Label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Bereit"
    $lblStatus.Location = New-Object System.Drawing.Point(10, 550)
    $lblStatus.Size = New-Object System.Drawing.Size(660, 20)
    $lblStatus.ForeColor = [System.Drawing.Color]::Blue
    
    # Event-Handler
    $btnAddSource.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Wählen Sie einen Quellordner"
        if ($folderDialog.ShowDialog() -eq "OK") {
            if ($txtSources.Text -ne "") {
                $txtSources.Text += "`r`n"
            }
            $txtSources.Text += $folderDialog.SelectedPath
        }
    })
    
    $btnBrowseDest.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Wählen Sie den Backup-Zielordner"
        if ($folderDialog.ShowDialog() -eq "OK") {
            $txtDest.Text = $folderDialog.SelectedPath
        }
    })
    
    $btnBackupNow.Add_Click({
        $lblStatus.Text = "Backup läuft..."
        $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        $form.Refresh()
        
        # Konfiguration sammeln
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
        
        # Backup durchführen
        $result = Start-BackupProcess -SourcePaths $sourcePaths -DestinationPath $destPath -FileTypes $fileTypes
        
        if ($result.Success) {
            $lblStatus.Text = "Backup erfolgreich! $($result.CopiedFiles) Dateien gesichert."
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
            
            # Azure-Sync wenn aktiviert
            if ($chkAzure.Checked -and $txtStorageAccount.Text -and $txtContainer.Text) {
                $lblStatus.Text = "Synchronisiere mit Azure..."
                $form.Refresh()
                
                $azureResult = Sync-ToAzure -LocalPath $result.BackupPath `
                    -StorageAccount $txtStorageAccount.Text `
                    -Container $txtContainer.Text `
                    -SasToken $txtSasToken.Text
                
                if ($azureResult) {
                    $lblStatus.Text = "Backup und Azure-Sync erfolgreich!"
                } else {
                    $lblStatus.Text = "Backup erfolgreich, Azure-Sync fehlgeschlagen."
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Backup erfolgreich abgeschlossen!`n`nGesicherte Dateien: $($result.CopiedFiles)`nBackup-Pfad: $($result.BackupPath)",
                "Backup erfolgreich",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } else {
            $lblStatus.Text = "Backup fehlgeschlagen!"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show(
                "Backup fehlgeschlagen! Überprüfen Sie die Einstellungen.",
                "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $btnSaveConfig.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "JSON-Dateien (*.json)|*.json"
        $saveDialog.FileName = "backup_config.json"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            $config = @{
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
                AutoBackup = $chkAutoBackup.Checked
                BackupInterval = $numInterval.Value
                AzureEnabled = $chkAzure.Checked
                AzureStorageAccount = $txtStorageAccount.Text
                AzureContainer = $txtContainer.Text
                DailyBackup = $chkDaily.Checked
                DailyTime = $dtpTime.Value.ToString("HH:mm")
                WeeklyBackup = $chkWeekly.Checked
            }
            
            $config | ConvertTo-Json | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            $lblStatus.Text = "Konfiguration gespeichert!"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
        }
    })
    
    $btnLoadConfig.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Filter = "JSON-Dateien (*.json)|*.json"
        
        if ($openDialog.ShowDialog() -eq "OK") {
            try {
                $config = Get-Content -Path $openDialog.FileName -Raw | ConvertFrom-Json
                
                $txtSources.Text = $config.SourcePaths -join "`r`n"
                $txtDest.Text = $config.DestinationPath
                $chkWord.Checked = $config.FileTypes.Word
                $chkExcel.Checked = $config.FileTypes.Excel
                $chkPowerPoint.Checked = $config.FileTypes.PowerPoint
                $chkPDF.Checked = $config.FileTypes.PDF
                $chkImages.Checked = $config.FileTypes.Images
                $chkAll.Checked = $config.FileTypes.All
                $chkAutoBackup.Checked = $config.AutoBackup
                $numInterval.Value = $config.BackupInterval
                $chkAzure.Checked = $config.AzureEnabled
                $txtStorageAccount.Text = $config.AzureStorageAccount
                $txtContainer.Text = $config.AzureContainer
                
                $lblStatus.Text = "Konfiguration geladen!"
                $lblStatus.ForeColor = [System.Drawing.Color]::Green
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Fehler beim Laden der Konfiguration: $_",
                    "Fehler",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })
    
    $btnCreateTask.Add_Click({
        $taskName = "Office-Backup-System"
        $taskPath = $MyInvocation.MyCommand.Path
        
        # Aufgabe erstellen
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$taskPath`" -Silent"
        
        $triggers = @()
        if ($chkDaily.Checked) {
            $triggers += New-ScheduledTaskTrigger -Daily -At $dtpTime.Value
        }
        
        if ($triggers.Count -gt 0) {
            try {
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Force
                [System.Windows.Forms.MessageBox]::Show(
                    "Windows-Aufgabe '$taskName' wurde erfolgreich erstellt!",
                    "Erfolg",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                $lblStatus.Text = "Aufgabe erstellt!"
                $lblStatus.ForeColor = [System.Drawing.Color]::Green
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Fehler beim Erstellen der Aufgabe: $_",
                    "Fehler",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Bitte wählen Sie mindestens einen Zeitplan aus!",
                "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    })
    
    $btnExit.Add_Click({
        $form.Close()
    })
    
    # Controls zum Formular hinzufügen
    $form.Controls.AddRange(@(
        $tabControl,
        $btnBackupNow, $btnSaveConfig, $btnLoadConfig, $btnExit,
        $lblStatus
    ))
    
    # Formular anzeigen
    $form.ShowDialog()
}

# Silent Mode für automatische Ausführung
if ($args -contains "-Silent") {
    # Konfiguration aus Standard-Datei laden
    $configPath = Join-Path $PSScriptRoot "backup_config.json"
    if (Test-Path $configPath) {
        $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
        Start-BackupProcess -SourcePaths $config.SourcePaths `
            -DestinationPath $config.DestinationPath `
            -FileTypes $config.FileTypes
        
        if ($config.AzureEnabled) {
            # Azure-Sync durchführen
            # ...
        }
    }
} else {
    # GUI starten
    Show-BackupGUI
}