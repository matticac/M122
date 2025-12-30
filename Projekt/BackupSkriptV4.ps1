# ================================================================
# Office Backup System - Final Edition v3.9 (UI Polish)
# ================================================================
# Fix: Layout verbessert (nicht mehr zusammengedrückt)
# Fix: Alle Abkürzungen ausgeschrieben (Benutzername, Zielpfad etc.)
# Autor: Backup System Generator & M122
# Version: 3.9 UI-Fix
# ================================================================

# .NET-Assemblies für GUI laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Konfiguration
$script:Version = "3.9 UI-Fix"
$script:BackupConfig = @{
    SourcePaths = @()
    DestinationPath = ""
    ExcludePaths = @("TBZ") 
    FileTypes = @{
        Word = $true; Excel = $true; PowerPoint = $true; PDF = $true; Images = $false; All = $false
    }
    AutoBackup = $false
    BackupInterval = 60
    BackupMode = "Smart"
    ForceTimeCheck = $true
    LastBackupPath = ""
    VmIP = "192.168.112.128"
    VmUser = "kali"
    VmDest = "/home/kali/Backup"
}

# ================================================================
# HILFSFUNKTIONEN
# ================================================================

function Update-FileTimestamp {
    param([string]$FilePath)
    try { (Get-Item $FilePath).LastWriteTime = Get-Date; return $true } catch { return $false }
}

function Get-BackupChain {
    param([string]$BackupRoot)
    if (-not (Test-Path $BackupRoot)) { return @() }
    $lastFull = Get-ChildItem -Path $BackupRoot -Directory | Where-Object { $_.Name -match "_Full$" } | Sort-Object CreationTime -Descending | Select-Object -First 1
    if (-not $lastFull) { return @() }
    return Get-ChildItem -Path $BackupRoot -Directory | Where-Object { $_.Name -match "^Backup_" -and $_.CreationTime -ge $lastFull.CreationTime } | Sort-Object CreationTime -Descending
}

function Test-FileNeedsBackup {
    param([System.IO.FileInfo]$File, [System.IO.DirectoryInfo[]]$BackupChain, [string]$Category, [string]$RelativePath)
    
    if (-not $BackupChain -or $BackupChain.Count -eq 0) { return @{NeedsBackup = $true; Reason = "Kein Basis-Backup"} }
    
    foreach ($backupFolder in $BackupChain) {
        $checkPath = Join-Path (Join-Path (Join-Path $backupFolder.FullName $Category) ($RelativePath -replace "^\\","")) $File.Name
        
        if (Test-Path $checkPath) {
            $backupFile = Get-Item $checkPath
            if ($File.Length -ne $backupFile.Length) { return @{NeedsBackup = $true; Reason = "Größe geändert"} }
            if ([Math]::Abs(($File.LastWriteTime - $backupFile.LastWriteTime).TotalSeconds) -gt 3) { return @{NeedsBackup = $true; Reason = "Zeitstempel geändert"} }
            return @{NeedsBackup = $false; Reason = "Bereits in $($backupFolder.Name)"}
        }
    }
    return @{NeedsBackup = $true; Reason = "Neue Datei"}
}

function Start-VmUpload {
    param([string]$LocalPath, [string]$IP, [string]$User, [string]$RemotePath)
    
    if (-not (Test-Connection -ComputerName $IP -Count 1 -Quiet)) {
        return @{ Success=$false; Message="VM ($IP) ist nicht erreichbar (Offline?)" }
    }
    
    try {
        ssh -o ConnectTimeout=5 "$User@$IP" "mkdir -p $RemotePath"
        if ($LASTEXITCODE -ne 0) { return @{ Success=$false; Message="SSH Verbindung fehlgeschlagen (Key?)" } }
        
        scp -r -p "$LocalPath" "$($User)@$($IP):$RemotePath"
        
        if ($LASTEXITCODE -eq 0) { return @{ Success=$true; Message="Upload erfolgreich!" } }
        else { return @{ Success=$false; Message="SCP Fehler (Code $LASTEXITCODE)" } }
    }
    catch {
        return @{ Success=$false; Message="Fehler: $_" }
    }
}

# ================================================================
# BACKUP PROZESS
# ================================================================

function Start-BackupProcess {
    param($SourcePaths, $DestinationPath, $ExcludePaths, $FileTypes, $BackupMode, $ForceBackup)
    
    $cleanExcludes = $ExcludePaths | Where { $_ -and $_.Trim() }
    
    $backupChain = @()
    if ($BackupMode -eq "Smart") {
        $backupChain = Get-BackupChain -BackupRoot $DestinationPath
        if ($backupChain.Count -gt 0) {
            $daysOld = (Get-Date).Subtract($backupChain[$backupChain.Count-1].CreationTime).TotalDays
            if ($daysOld -gt 7) { $BackupMode = "Full"; $backupChain = @() } else { $BackupMode = "Incremental" }
        } else { $BackupMode = "Full" }
    } elseif ($BackupMode -eq "Incremental") {
        $backupChain = Get-BackupChain -BackupRoot $DestinationPath
        if ($backupChain.Count -eq 0) { $BackupMode = "Full" }
    }
    
    $backupDate = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $typeSuffix = if ($BackupMode -eq "Full") { "Full" } elseif ($BackupMode -eq "Incremental") { "Inc" } else { "Bkp" }
    $backupRoot = Join-Path $DestinationPath "Backup_${backupDate}_${typeSuffix}"
    
    $extensions = @{}
    if ($FileTypes.Word) { $extensions['Word'] = @('*.docx', '*.doc') }
    if ($FileTypes.Excel) { $extensions['Excel'] = @('*.xlsx', '*.xls') }
    if ($FileTypes.PowerPoint) { $extensions['PowerPoint'] = @('*.pptx', '*.ppt') }
    if ($FileTypes.PDF) { $extensions['PDF'] = @('*.pdf') }
    if ($FileTypes.Images) { $extensions['Images'] = @('*.jpg', '*.png') }
    if ($FileTypes.All) { $extensions['All'] = @('*.*') }
    
    $stats = @{ Copied=0; Excluded=0; Checked=0; Errors=@() }
    $fileList = @()
    
    foreach ($pathRaw in $SourcePaths) {
        $src = $pathRaw.TrimEnd('\')
        if (-not (Test-Path $src)) { continue }
        
        foreach ($cat in $extensions.Keys) {
            foreach ($ext in $extensions[$cat]) {
                Get-ChildItem -Path $src -Filter $ext -Recurse -File -EA SilentlyContinue | ForEach-Object {
                    $stats.Checked++
                    $file = $_
                    
                    $isExcluded = $false
                    foreach ($ex in $cleanExcludes) { if ($file.FullName -like "*\$ex\*" -or $file.FullName -like "$ex\*") { $isExcluded=$true; break } }
                    if ($isExcluded) { $stats.Excluded++; return }
                    
                    $relPath = if ($file.DirectoryName.StartsWith($src)) { $file.DirectoryName.Substring($src.Length).TrimStart('\') } else { "" }
                    
                    $check = if ($ForceBackup -or $BackupMode -eq "Full") { @{NeedsBackup=$true; Reason="Full"} } 
                             else { Test-FileNeedsBackup -File $file -BackupChain $backupChain -Category $cat -RelativePath $relPath }
                    
                    if ($check.NeedsBackup) {
                        $fileList += @{ File=$file; Cat=$cat; Rel=$relPath; Reason=$check.Reason }
                    }
                }
            }
        }
    }
    
    if ($fileList.Count -gt 0) {
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
        $prog=0
        foreach ($item in $fileList) {
            $prog++
            $target = Join-Path (Join-Path (Join-Path $backupRoot $item.Cat) $item.Rel) $item.File.Name
            $parent = Split-Path $target
            if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            Copy-Item $item.File.FullName $target -Force
            $stats.Copied++
        }
        "Backup Info`nMode: $BackupMode`nFiles: $($stats.Copied)" | Out-File (Join-Path $backupRoot "info.txt")
        return @{ Success=$true; Files=$stats.Copied; Path=$backupRoot; Mode=$BackupMode } 
    }
    return @{ Success=$true; Files=0; Path=$null; Mode=$BackupMode }
}

# ================================================================
# GUI INTERFACE (NEU & AUFGERÄUMT)
# ================================================================

function Show-BackupGUI {
    # Fenster etwas größer machen damit Platz ist
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office Backup System v$script:Version"; 
    $form.Size = New-Object System.Drawing.Size(780, 850) # HIER: Länger gemacht
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Size = New-Object System.Drawing.Size(740, 700) # HIER: Tab auch größer
    $tabs.Location = New-Object System.Drawing.Point(10, 10)
    
    # --- TAB 1: BACKUP ---
    $tab1 = New-Object System.Windows.Forms.TabPage; $tab1.Text = "📁 Backup Einstellungen"
    
    # 1. Quelle
    $grpSrc = New-Object System.Windows.Forms.GroupBox; $grpSrc.Text="1. Quellordner (Was sichern?)"; $grpSrc.Location="15,15"; $grpSrc.Size="700,110"
    $txtSrc = New-Object System.Windows.Forms.TextBox; $txtSrc.Multiline=$true; $txtSrc.ScrollBars="Vertical"; $txtSrc.Location="15,25"; $txtSrc.Size="550,70"; $txtSrc.Text="C:\Users\matti"
    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text="Hinzufügen"; $btnAdd.Location="580,25"; $btnAdd.Size="100,30"
    $btnAdd.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtSrc.Text+="`r`n"+$d.SelectedPath}})
    $grpSrc.Controls.AddRange(@($txtSrc, $btnAdd))
    
    # 2. Ziel
    $grpDst = New-Object System.Windows.Forms.GroupBox; $grpDst.Text="2. Zielordner (Wohin speichern?)"; $grpDst.Location="15,140"; $grpDst.Size="700,70"
    $txtDst = New-Object System.Windows.Forms.TextBox; $txtDst.Location="15,30"; $txtDst.Size="550,25"; $txtDst.Text="C:\Backup"
    $btnDst = New-Object System.Windows.Forms.Button; $btnDst.Text="Suchen..."; $btnDst.Location="580,28"; $btnDst.Size="100,27"; 
    $btnDst.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtDst.Text=$d.SelectedPath}})
    $grpDst.Controls.AddRange(@($txtDst, $btnDst))
    
    # 3. Ausnahmen
    $grpEx = New-Object System.Windows.Forms.GroupBox; $grpEx.Text="3. Ausnahmen (Diese Ordnernamen ignorieren)"; $grpEx.Location="15,225"; $grpEx.Size="700,70"; $grpEx.ForeColor="DarkRed"
    $txtEx = New-Object System.Windows.Forms.TextBox; $txtEx.Location="15,30"; $txtEx.Size="550,25"; $txtEx.Text="TBZ"
    $btnEx = New-Object System.Windows.Forms.Button; $btnEx.Text="Suchen..."; $btnEx.Location="580,28"; $btnEx.Size="100,27"; 
    $btnEx.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$n=Split-Path $d.SelectedPath -Leaf;if($txtEx.Text){$txtEx.Text+=", "+$n}else{$txtEx.Text=$n}}})
    $grpEx.Controls.AddRange(@($txtEx, $btnEx))
    
    # 4. Typen
    $grpTyp = New-Object System.Windows.Forms.GroupBox; $grpTyp.Text="4. Dateitypen auswählen"; $grpTyp.Location="15,310"; $grpTyp.Size="700,100"
    # Layout breiter machen
    $chkW=New-Object System.Windows.Forms.CheckBox;$chkW.Text="Word Dokumente";$chkW.Location="20,25";$chkW.Size="150,20";$chkW.Checked=$true
    $chkE=New-Object System.Windows.Forms.CheckBox;$chkE.Text="Excel Tabellen";$chkE.Location="200,25";$chkE.Size="150,20";$chkE.Checked=$true
    $chkP=New-Object System.Windows.Forms.CheckBox;$chkP.Text="PowerPoint";$chkP.Location="380,25";$chkP.Size="150,20";$chkP.Checked=$true
    $chkD=New-Object System.Windows.Forms.CheckBox;$chkD.Text="PDF Dokumente";$chkD.Location="20,55";$chkD.Size="150,20";$chkD.Checked=$true
    $grpTyp.Controls.AddRange(@($chkW,$chkE,$chkP,$chkD))
    
    # 5. Modus
    $grpMod = New-Object System.Windows.Forms.GroupBox; $grpMod.Text="5. Backup Modus"; $grpMod.Location="15,425"; $grpMod.Size="700,80"
    $radS=New-Object System.Windows.Forms.RadioButton;$radS.Text="Smart (Automatisch)";$radS.Location="20,30";$radS.Size="150,20";$radS.Checked=$true
    $radF=New-Object System.Windows.Forms.RadioButton;$radF.Text="Vollständig";$radF.Location="200,30";$radF.Size="150,20"
    $radI=New-Object System.Windows.Forms.RadioButton;$radI.Text="Inkrementell";$radI.Location="380,30";$radI.Size="150,20"
    $grpMod.Controls.AddRange(@($radS,$radF,$radI))
    
    $tab1.Controls.AddRange(@($grpSrc, $grpDst, $grpEx, $grpTyp, $grpMod))
    
    # --- TAB 2: HISTORY ---
    $tab2 = New-Object System.Windows.Forms.TabPage; $tab2.Text = "📜 Verlauf & Historie"
    $lstH = New-Object System.Windows.Forms.ListBox; $lstH.Location="15,15"; $lstH.Size="700,600"; $lstH.Font="Consolas,9"
    $btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="Liste aktualisieren"; $btnRef.Location="15,630"; $btnRef.Size="150,30"
    $btnRef.Add_Click({$lstH.Items.Clear();if(Test-Path $txtDst.Text){Get-ChildItem $txtDst.Text -Dir|Where Name -match "^Backup_"|Sort Name -Desc|ForEach{$lstH.Items.Add($_.Name)}}})
    $tab2.Controls.AddRange(@($lstH, $btnRef))
    
    # --- TAB 3: TOOLS / VM (HIER DAS UI FIX) ---
    $tab3 = New-Object System.Windows.Forms.TabPage; $tab3.Text = "🔧 Tools & VM Upload"
    
    $grpVM = New-Object System.Windows.Forms.GroupBox; $grpVM.Text="Kali VM Verbindung (SCP/SSH)"; $grpVM.Location="15,15"; $grpVM.Size="700,200" # Höher gemacht
    
    # Zeile 1: IP
    $lblIP = New-Object System.Windows.Forms.Label; $lblIP.Text="IP Adresse:"; $lblIP.Location="20,40"; $lblIP.Size="80,20"
    $txtIP = New-Object System.Windows.Forms.TextBox; $txtIP.Text=$script:BackupConfig.VmIP; $txtIP.Location="110,37"; $txtIP.Size="150,20"
    
    # Zeile 1: User (rechts daneben)
    $lblUs = New-Object System.Windows.Forms.Label; $lblUs.Text="Benutzername:"; $lblUs.Location="280,40"; $lblUs.Size="90,20" # Breiter
    $txtUs = New-Object System.Windows.Forms.TextBox; $txtUs.Text=$script:BackupConfig.VmUser; $txtUs.Location="380,37"; $txtUs.Size="150,20"
    
    # Test Button (ganz rechts)
    $btnTestVM = New-Object System.Windows.Forms.Button; $btnTestVM.Text="Verbindung testen"; $btnTestVM.Location="550,35"; $btnTestVM.Size="130,25"
    
    # Zeile 2: Pfad
    $lblPa = New-Object System.Windows.Forms.Label; $lblPa.Text="Zielpfad auf VM:"; $lblPa.Location="20,90"; $lblPa.Size="100,20" # Breiter
    $txtPa = New-Object System.Windows.Forms.TextBox; $txtPa.Text=$script:BackupConfig.VmDest; $txtPa.Location="130,87"; $txtPa.Size="400,20"
    
    $btnTestVM.Add_Click({
        if(Test-Connection -ComputerName $txtIP.Text -Count 1 -Quiet){ [System.Windows.Forms.MessageBox]::Show("VM Online!", "Verbindung OK") }
        else { [System.Windows.Forms.MessageBox]::Show("VM nicht erreichbar.", "Verbindungsfehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    })
    
    $grpVM.Controls.AddRange(@($lblIP, $txtIP, $lblUs, $txtUs, $lblPa, $txtPa, $btnTestVM))
    $tab3.Controls.Add($grpVM)
    
    $tabs.TabPages.AddRange(@($tab1, $tab2, $tab3))
    
    # --- MAIN BUTTONS ---
    # Position angepasst an neues Fenster
    $btnStart = New-Object System.Windows.Forms.Button; $btnStart.Text="BACKUP STARTEN"; $btnStart.BackColor="LightGreen"; $btnStart.Location="15,730"; $btnStart.Size="200,50"; $btnStart.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    
    $status = New-Object System.Windows.Forms.Label; $status.Text="System Bereit"; $status.Location="230,745"; $status.Size="400,20"; $status.ForeColor="Blue"
    
    $btnStart.Add_Click({
        $status.Text="Backup läuft..."
        $form.Cursor="WaitCursor"; $form.Refresh()
        
        $src=$txtSrc.Text -split "`r`n"|Where{$_.Trim()}; $excl=$txtEx.Text -split "[,`r`n]"|Where{$_.Trim()}
        $types=@{Word=$chkW.Checked;Excel=$chkE.Checked;PowerPoint=$chkP.Checked;PDF=$chkD.Checked;Images=$false;All=$false}
        $mode=if($radF.Checked){"Full"}elseif($radI.Checked){"Incremental"}else{"Smart"}
        
        if(-not(Test-Path $txtDst.Text)){New-Item -ItemType Directory -Path $txtDst.Text -Force|Out-Null}
        
        $res = Start-BackupProcess -SourcePaths $src -DestinationPath $txtDst.Text -ExcludePaths $excl -FileTypes $types -BackupMode $mode -ForceBackup $false
        
        $form.Cursor="Default"
        
        if ($res.Success) {
            $status.Text="Fertig: $($res.Files) Dateien gesichert."
            $msg = "Backup erfolgreich abgeschlossen!`nGesicherte Dateien: $($res.Files)"
            
            # --- VM LOGIK ---
            if ($res.Mode -eq "Full" -and $res.Files -gt 0) {
                $upl = [System.Windows.Forms.MessageBox]::Show($msg + "`n`nDies war ein Vollständiges Backup.`nMöchten Sie es auf die Kali VM hochladen?", "VM Upload", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($upl -eq "Yes") {
                    $status.Text = "Lade auf VM hoch..."
                    $form.Refresh()
                    $vmRes = Start-VmUpload -LocalPath $res.Path -IP $txtIP.Text -User $txtUs.Text -RemotePath $txtPa.Text
                    if ($vmRes.Success) { [System.Windows.Forms.MessageBox]::Show($vmRes.Message, "Upload Erfolg") }
                    else { [System.Windows.Forms.MessageBox]::Show($vmRes.Message, "Upload Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
                    $status.Text = $vmRes.Message
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show($msg, "Erfolg")
            }
            # ----------------
            
            & $btnRef.PerformClick()
        } else { $status.Text="Fehler beim Backup." }
    })
    
    $form.Controls.AddRange(@($tabs, $btnStart, $status))
    $form.ShowDialog()|Out-Null
}

Show-BackupGUI