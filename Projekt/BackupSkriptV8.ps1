# ================================================================
# Office Backup System - Final Edition v5.2 (Bugfix Refresh)
# ================================================================
# Fix: Absturz nach erfolgreichem Backup behoben (Call Operator &)
# Fix: "Illegale Zeichen" Fehler behoben
# Feature: Robustes Überspringen von Problem-Dateien
# Feature: Speicherplatz-Check
# Autor: M122
# ================================================================

# .NET-Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- KONFIGURATIONSMANAGEMENT ---
if ($PSScriptRoot) { $BaseDir = $PSScriptRoot } else { $BaseDir = $PWD.Path }
$ConfigFile = Join-Path $BaseDir "backup_config.json"

$DefaultConfig = @{
    SourcePaths = @([Environment]::GetFolderPath("MyDocuments"))
    DestinationPath = "C:\Backup"
    ExcludePaths = @("TBZ", "Temp")
    FileTypes = @{ Word=$true; Excel=$true; PowerPoint=$true; PDF=$true; Images=$false; All=$false }
    BackupMode = "Smart"
    UseZip = $false
    UseVM = $false
    VmIP = ""
    VmUser = ""
    VmPass = "" 
    VmDest = "/home/kali/Backup"
}

if (Test-Path $ConfigFile) {
    try {
        $LoadedConfig = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        $script:BackupConfig = $DefaultConfig.Clone()
        foreach ($prop in $LoadedConfig.PSObject.Properties) {
            if ($script:BackupConfig.ContainsKey($prop.Name)) { $script:BackupConfig[$prop.Name] = $prop.Value }
        }
    } catch { $script:BackupConfig = $DefaultConfig.Clone() }
} else { $script:BackupConfig = $null }

# ================================================================
# FUNKTIONEN
# ================================================================

function Save-Config {
    try { $script:BackupConfig | ConvertTo-Json -Depth 5 | Out-File $ConfigFile -Encoding UTF8 -Force } catch {}
}

function Show-Notification {
    param([string]$Title, [string]$Message, [string]$Type = "Info")
    $icon = New-Object System.Windows.Forms.NotifyIcon
    if ($Type -eq "Error") { $icon.Icon = [System.Drawing.SystemIcons]::Error; $tipIcon = [System.Windows.Forms.ToolTipIcon]::Error } 
    elseif ($Type -eq "Warning") { $icon.Icon = [System.Drawing.SystemIcons]::Warning; $tipIcon = [System.Windows.Forms.ToolTipIcon]::Warning }
    else { $icon.Icon = [System.Drawing.SystemIcons]::Information; $tipIcon = [System.Windows.Forms.ToolTipIcon]::Info }
    $icon.Visible = $true
    $icon.ShowBalloonTip(3000, $Title, $Message, $tipIcon)
    Start-Sleep -Seconds 1
    $icon.Dispose()
}

function Get-BackupChain {
    param([string]$BackupRoot)
    if (-not (Test-Path $BackupRoot)) { return @() }
    $lastFull = Get-ChildItem -Path $BackupRoot -Directory | Where-Object { $_.Name -match "_Full$" } | Sort-Object CreationTime -Descending | Select-Object -First 1
    if (-not $lastFull) { return @() }
    return Get-ChildItem -Path $BackupRoot -Directory | Where-Object { $_.Name -match "^Backup_" -and $_.CreationTime -ge $lastFull.CreationTime } | Sort-Object CreationTime -Descending
}

function Test-FileNeedsBackup {
    param($File, $BackupChain, $Category, $RelativePath)
    if (-not $BackupChain) { return @{NeedsBackup=$true; Reason="Kein Basis-Backup"} }
    foreach ($fol in $BackupChain) {
        $checkPath = Join-Path (Join-Path (Join-Path $fol.FullName $Category) ($RelativePath -replace "^\\","")) $File.Name
        try {
            if (Test-Path $checkPath) {
                $bf = Get-Item $checkPath
                if ($File.Length -ne $bf.Length) { return @{NeedsBackup=$true; Reason="Größe"} }
                if ([Math]::Abs(($File.LastWriteTime - $bf.LastWriteTime).TotalSeconds) -gt 3) { return @{NeedsBackup=$true; Reason="Zeit"} }
                return @{NeedsBackup=$false; Reason="Vorhanden"}
            }
        } catch { return @{NeedsBackup=$true; Reason="Fehler"} }
    }
    return @{NeedsBackup=$true; Reason="Neu"}
}

function Check-DiskSpace {
    param($DestPath, $RequiredBytes)
    try {
        $root = Split-Path $DestPath -Qualifier
        if (-not $root) { return @{OK=$true; Free=0} } 
        $drive = Get-Volume -DriveLetter ($root.TrimEnd(':')) -ErrorAction Stop
        
        if ($drive.SizeRemaining -lt $RequiredBytes) {
            return @{OK=$false; Free=$drive.SizeRemaining}
        }
        return @{OK=$true; Free=$drive.SizeRemaining}
    } catch {
        return @{OK=$true; Free=0} 
    }
}

function Start-VmUpload {
    param($LocalPath, $IP, $User, $Pass, $RemotePath)
    if (-not (Test-Connection -ComputerName $IP -Count 1 -Quiet)) { return @{ Success=$false; Message="VM Offline" } }
    try {
        ssh -o ConnectTimeout=5 -o StrictHostKeyChecking=no "$User@$IP" "mkdir -p $RemotePath"
        if ($LASTEXITCODE -ne 0) { return @{ Success=$false; Message="SSH Login fehlgeschlagen" } }
        scp -r -p -o StrictHostKeyChecking=no "$LocalPath" "$($User)@$($IP):$RemotePath"
        if ($LASTEXITCODE -eq 0) { return @{ Success=$true; Message="Upload OK" } }
        else { return @{ Success=$false; Message="SCP Fehler" } }
    } catch { return @{ Success=$false; Message="Fehler: $_" } }
}

function Start-Cleanup {
    $dst = $script:BackupConfig.DestinationPath
    if (-not (Test-Path $dst)) { [System.Windows.Forms.MessageBox]::Show("Backup-Pfad nicht gefunden.", "Fehler"); return }
    $fulls = Get-ChildItem -Path $dst | Where-Object { $_.Name -match "_Full" } | Sort-Object Name -Descending
    if (-not $fulls) { [System.Windows.Forms.MessageBox]::Show("Kein Full-Backup gefunden.", "Info"); return }
    $cutoffName = $fulls[0].Name
    if ($cutoffName.EndsWith(".zip")) { $cutoffName = $cutoffName.Substring(0, $cutoffName.Length - 4) }
    $items = Get-ChildItem -Path $dst | Where-Object { $_.Name -match "^Backup_" }
    $delCount = 0
    foreach ($item in $items) {
        $itemName = $item.Name
        if ($itemName.EndsWith(".zip")) { $itemName = $itemName.Substring(0, $itemName.Length - 4) }
        if ($itemName -lt $cutoffName) { 
            try { Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop; $delCount++ } catch {}
        }
    }
    $msgLokal = "Lokal bereinigt: $delCount Objekte gelöscht."
    if ($script:BackupConfig.UseVM -and $script:BackupConfig.VmIP) {
        $ans = [System.Windows.Forms.MessageBox]::Show("$msgLokal`n`nAuch auf VM ($($script:BackupConfig.VmIP)) bereinigen?", "VM Bereinigung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -eq "Yes") {
             $user = $script:BackupConfig.VmUser; $ip = $script:BackupConfig.VmIP; $remPath = $script:BackupConfig.VmDest
             try {
                $cmdList = "ls -1 $remPath"
                $res = ssh -o ConnectTimeout=5 -o StrictHostKeyChecking=no "$user@$ip" $cmdList
                if ($LASTEXITCODE -ne 0) { throw "SSH Login fehlgeschlagen" }
                $remDeleted = 0
                foreach ($line in $res) {
                    $line = $line.Trim()
                    $compLine = $line
                    if ($compLine.EndsWith(".zip")) { $compLine = $compLine.Substring(0, $compLine.Length - 4) }
                    if ($line -match "^Backup_" -and ($line -match "_Full" -or $line -match "_Inc")) {
                        if ($compLine -lt $cutoffName) {
                             ssh -o StrictHostKeyChecking=no "$user@$ip" "rm -rf '$remPath/$line'"
                             $remDeleted++
                        }
                    }
                }
                [System.Windows.Forms.MessageBox]::Show("VM Bereinigung fertig.`nGelöscht: $remDeleted", "Erfolg")
             } catch { [System.Windows.Forms.MessageBox]::Show("Fehler VM: $_", "Fehler") }
        }
    } else { [System.Windows.Forms.MessageBox]::Show($msgLokal, "Fertig") }
}

function Start-BackupProcess {
    param($SourcePaths, $DestinationPath, $ExcludePaths, $FileTypes, $BackupMode, $Compress)
    $cleanExcludes = $ExcludePaths | Where { $_ -and $_.Trim() }
    $backupChain = @()
    
    if ($BackupMode -eq "Smart") {
        $backupChain = Get-BackupChain -BackupRoot $DestinationPath
        if ($backupChain.Count -gt 0) {
            if ((Get-Date).Subtract($backupChain[0].CreationTime).TotalDays -gt 7) { $BackupMode="Full"; $backupChain=@() } else { $BackupMode="Incremental" }
        } else { $BackupMode="Full" }
    } elseif ($BackupMode -eq "Incremental") {
        $backupChain = Get-BackupChain -BackupRoot $DestinationPath
        if ($backupChain.Count -eq 0) { $BackupMode="Full" }
    }
    
    $date = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $suffix = if($BackupMode -eq "Full"){"Full"}elseif($BackupMode -eq "Incremental"){"Inc"}else{"Bkp"}
    $rootName = "Backup_${date}_${suffix}"
    $root = Join-Path $DestinationPath $rootName
    
    $exts = @{}; if($FileTypes.Word){$exts['Word']=@('*.docx','*.doc')}; if($FileTypes.Excel){$exts['Excel']=@('*.xlsx','*.xls')}; if($FileTypes.PowerPoint){$exts['PowerPoint']=@('*.pptx','*.ppt')}; if($FileTypes.PDF){$exts['PDF']=@('*.pdf')}
    
    $stats = @{Copied=0; Checked=0; SizeBytes=0; Errors=0}
    $list = @()
    
    # 1. Dateien sammeln
    foreach($p in $SourcePaths){
        $src=$p.TrimEnd('\'); if(-not(Test-Path $src)){continue}
        foreach($cat in $exts.Keys){
            foreach($ext in $exts[$cat]){
                Get-ChildItem -Path $src -Filter $ext -Recurse -File -EA SilentlyContinue | ForEach {
                    $stats.Checked++; $f=$_; $skip=$false
                    foreach($ex in $cleanExcludes){if($f.FullName -like "*\$ex\*" -or $f.FullName -like "$ex\*"){$skip=$true;break}}
                    if($skip){return}
                    
                    try {
                        $rel=if($f.DirectoryName.StartsWith($src)){$f.DirectoryName.Substring($src.Length).TrimStart('\')}else{""}
                        $chk=if($BackupMode -eq "Full"){@{NeedsBackup=$true}}else{Test-FileNeedsBackup -File $f -BackupChain $backupChain -Category $cat -RelativePath $rel}
                        if($chk.NeedsBackup){
                            $list+=@{F=$f;C=$cat;R=$rel}
                            $stats.SizeBytes += $f.Length
                        }
                    } catch { }
                }
            }
        }
    }

    # 2. Speicherplatz Check
    if ($list.Count -gt 0) {
        if(-not(Test-Path $DestinationPath)){ New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null }
        
        $spaceCheck = Check-DiskSpace -DestPath $DestinationPath -RequiredBytes $stats.SizeBytes
        if (-not $spaceCheck.OK) {
            $mbNeeded = [math]::Round($stats.SizeBytes / 1MB, 2)
            $mbFree = [math]::Round($spaceCheck.Free / 1MB, 2)
            return @{Success=$false; Message="NICHT GENUG SPEICHER!`nBenötigt: $mbNeeded MB`nFrei: $mbFree MB"}
        }
    }
    
    # 3. Kopieren
    if($list.Count -gt 0){
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        
        foreach($i in $list){
            try {
                $dst = Join-Path (Join-Path (Join-Path $root $i.C) $i.R) $i.F.Name
                $parentDir = Split-Path $dst
                
                if (-not (Test-Path $parentDir)) { 
                    New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction SilentlyContinue | Out-Null 
                }
                
                Copy-Item -LiteralPath $i.F.FullName -Destination $dst -Force -ErrorAction Stop
                $stats.Copied++
            } catch {
                $stats.Errors++
                Write-Host "Übersprungen (Fehler): $($i.F.Name)" -ForegroundColor Yellow
            }
        }

        if ($Compress) {
            $zipFile = "$root.zip"
            $zipSpace = Check-DiskSpace -DestPath $DestinationPath -RequiredBytes ($stats.SizeBytes / 2)
             if (-not $zipSpace.OK) { return @{Success=$false; Message="Zu wenig Platz für ZIP Erstellung!"} }

            try {
                Compress-Archive -Path "$root\*" -DestinationPath $zipFile -CompressionLevel Optimal -Force
                Remove-Item $root -Recurse -Force
                return @{Success=$true; Files=$stats.Copied; Errors=$stats.Errors; Path=$zipFile; Mode=$BackupMode}
            } catch {
                return @{Success=$false; Message="Fehler beim Zippen: $_"}
            }
        }

        return @{Success=$true; Files=$stats.Copied; Errors=$stats.Errors; Path=$root; Mode=$BackupMode}
    }
    return @{Success=$true; Files=0; Errors=0; Path=$null; Mode=$BackupMode}
}

# ================================================================
# GUI
# ================================================================

function Show-MainGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office Backup System v5.2 (Bugfix Refresh)"; $form.Size="800,850"; $form.StartPosition="CenterScreen"; $form.FormBorderStyle="FixedDialog"; $form.MaximizeBox=$false
    
    $tabs = New-Object System.Windows.Forms.TabControl; $tabs.Size="760,680"; $tabs.Location="10,10"
    
    $btnStart = New-Object System.Windows.Forms.Button; $btnStart.Text="BACKUP STARTEN"; $btnStart.BackColor="LightGreen"; $btnStart.Location="15,710"; $btnStart.Size="200,50"; $btnStart.Font=New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text="Einstellungen Speichern"; $btnSave.Location="230,720"; $btnSave.Size="180,30"
    $status = New-Object System.Windows.Forms.Label; $status.Text="Bereit"; $status.Location="420,725"; $status.Size="350,20"; $status.ForeColor="Blue"

    # TAB 1: BACKUP
    $tab1 = New-Object System.Windows.Forms.TabPage; $tab1.Text = "📁 Backup Einstellungen"
    $grpSrc = New-Object System.Windows.Forms.GroupBox; $grpSrc.Text="1. Quellordner"; $grpSrc.Location="15,15"; $grpSrc.Size="720,110"
    $txtSrc = New-Object System.Windows.Forms.TextBox; $txtSrc.Multiline=$true; $txtSrc.ScrollBars="Vertical"; $txtSrc.Location="15,25"; $txtSrc.Size="570,70"; $txtSrc.Text = $script:BackupConfig.SourcePaths -join "`r`n"
    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text="Hinzufügen"; $btnAdd.Location="600,25"; $btnAdd.Size="100,30"; $btnAdd.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtSrc.Text+="`r`n"+$d.SelectedPath}})
    $grpSrc.Controls.AddRange(@($txtSrc, $btnAdd))
    
    $grpDst = New-Object System.Windows.Forms.GroupBox; $grpDst.Text="2. Zielordner"; $grpDst.Location="15,140"; $grpDst.Size="720,70"
    $txtDst = New-Object System.Windows.Forms.TextBox; $txtDst.Location="15,30"; $txtDst.Size="570,25"; $txtDst.Text = $script:BackupConfig.DestinationPath
    $btnDst = New-Object System.Windows.Forms.Button; $btnDst.Text="Suchen..."; $btnDst.Location="600,28"; $btnDst.Size="100,27"; $btnDst.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtDst.Text=$d.SelectedPath}})
    $grpDst.Controls.AddRange(@($txtDst, $btnDst))
    
    $grpEx = New-Object System.Windows.Forms.GroupBox; $grpEx.Text="3. Ausnahmen"; $grpEx.Location="15,225"; $grpEx.Size="720,70"; $grpEx.ForeColor="DarkRed"
    $txtEx = New-Object System.Windows.Forms.TextBox; $txtEx.Location="15,30"; $txtEx.Size="570,25"; $txtEx.Text = $script:BackupConfig.ExcludePaths -join ", "
    $btnEx = New-Object System.Windows.Forms.Button; $btnEx.Text="Suchen..."; $btnEx.Location="600,28"; $btnEx.Size="100,27"; $btnEx.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$n=Split-Path $d.SelectedPath -Leaf;if($txtEx.Text){$txtEx.Text+=", "+$n}else{$txtEx.Text=$n}}})
    $grpEx.Controls.AddRange(@($txtEx, $btnEx))
    
    $grpTyp = New-Object System.Windows.Forms.GroupBox; $grpTyp.Text="4. Dateitypen"; $grpTyp.Location="15,310"; $grpTyp.Size="720,80"
    $chkW=New-Object System.Windows.Forms.CheckBox;$chkW.Text="Word";$chkW.Location="20,25";$chkW.Size="100,20";$chkW.Checked=$script:BackupConfig.FileTypes.Word
    $chkE=New-Object System.Windows.Forms.CheckBox;$chkE.Text="Excel";$chkE.Location="130,25";$chkE.Size="100,20";$chkE.Checked=$script:BackupConfig.FileTypes.Excel
    $chkP=New-Object System.Windows.Forms.CheckBox;$chkP.Text="PowerPoint";$chkP.Location="240,25";$chkP.Size="120,20";$chkP.Checked=$script:BackupConfig.FileTypes.PowerPoint
    $chkD=New-Object System.Windows.Forms.CheckBox;$chkD.Text="PDF";$chkD.Location="370,25";$chkD.Size="80,20";$chkD.Checked=$script:BackupConfig.FileTypes.PDF
    $grpTyp.Controls.AddRange(@($chkW,$chkE,$chkP,$chkD))
    
    $grpMod = New-Object System.Windows.Forms.GroupBox; $grpMod.Text="5. Modus & Optionen"; $grpMod.Location="15,400"; $grpMod.Size="720,80"
    $radS=New-Object System.Windows.Forms.RadioButton;$radS.Text="Smart (Auto)";$radS.Location="20,30";$radS.Size="120,20";$radS.Checked=($script:BackupConfig.BackupMode -eq "Smart")
    $radF=New-Object System.Windows.Forms.RadioButton;$radF.Text="Vollständig";$radF.Location="150,30";$radF.Size="120,20";$radF.Checked=($script:BackupConfig.BackupMode -eq "Full")
    $radI=New-Object System.Windows.Forms.RadioButton;$radI.Text="Inkrementell";$radI.Location="280,30";$radI.Size="120,20";$radI.Checked=($script:BackupConfig.BackupMode -eq "Incremental")
    
    $chkZip=New-Object System.Windows.Forms.CheckBox;$chkZip.Text="Als ZIP speichern";$chkZip.Location="420,30";$chkZip.Size="180,20";$chkZip.Checked=$script:BackupConfig.UseZip
    $chkZip.ForeColor="DarkBlue"
    
    $grpMod.Controls.AddRange(@($radS,$radF,$radI,$chkZip))
    $tab1.Controls.AddRange(@($grpSrc, $grpDst, $grpEx, $grpTyp, $grpMod))
    
    # TAB 2: HISTORY
    $tab2 = New-Object System.Windows.Forms.TabPage; $tab2.Text = "📜 Historie"
    $lstH = New-Object System.Windows.Forms.ListBox; $lstH.Location="15,15"; $lstH.Size="720,600"; $lstH.Font="Consolas,9"
    $btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="Aktualisieren"; $btnRef.Location="15,620"; $btnRef.Size="150,30"; $btnRef.Add_Click({$lstH.Items.Clear();if(Test-Path $txtDst.Text){Get-ChildItem $txtDst.Text |Where Name -match "^Backup_"|Sort Name -Desc|ForEach{$lstH.Items.Add($_.Name)}}})
    $tab2.Controls.AddRange(@($lstH, $btnRef))
    
    # TAB 3: TOOLS / VM
    $tab3 = New-Object System.Windows.Forms.TabPage; $tab3.Text = "🔧 Tools & VM"
    $grpVM = New-Object System.Windows.Forms.GroupBox; $grpVM.Text="Kali VM Verbindung"; $grpVM.Location="15,15"; $grpVM.Size="720,250"
    $chkUseVM = New-Object System.Windows.Forms.CheckBox; $chkUseVM.Text="VM für Uploads nutzen"; $chkUseVM.Location="20,30"; $chkUseVM.Size="200,20"; $chkUseVM.Checked=$script:BackupConfig.UseVM
    $lblIP = New-Object System.Windows.Forms.Label; $lblIP.Text="IP Adresse:"; $lblIP.Location="20,70"; $lblIP.Size="80,20"
    $txtIP = New-Object System.Windows.Forms.TextBox; $txtIP.Text=$script:BackupConfig.VmIP; $txtIP.Location="110,67"; $txtIP.Size="150,20"
    $lblUs = New-Object System.Windows.Forms.Label; $lblUs.Text="Benutzer:"; $lblUs.Location="280,70"; $lblUs.Size="70,20"
    $txtUs = New-Object System.Windows.Forms.TextBox; $txtUs.Text=$script:BackupConfig.VmUser; $txtUs.Location="360,67"; $txtUs.Size="150,20"
    $lblPw = New-Object System.Windows.Forms.Label; $lblPw.Text="Passwort:"; $lblPw.Location="20,110"; $lblPw.Size="80,20"
    $txtPw = New-Object System.Windows.Forms.TextBox; $txtPw.Text=$script:BackupConfig.VmPass; $txtPw.Location="110,107"; $txtPw.Size="150,20"; $txtPw.PasswordChar="*"
    $lblWarn = New-Object System.Windows.Forms.Label; $lblWarn.Text="(SSH-Key empfohlen!)"; $lblWarn.Location="270,110"; $lblWarn.Size="250,20"; $lblWarn.ForeColor="Gray"
    $lblPa = New-Object System.Windows.Forms.Label; $lblPa.Text="Zielpfad:"; $lblPa.Location="20,150"; $lblPa.Size="80,20"
    $txtPa = New-Object System.Windows.Forms.TextBox; $txtPa.Text=$script:BackupConfig.VmDest; $txtPa.Location="110,147"; $txtPa.Size="400,20"
    
    $btnTestVM = New-Object System.Windows.Forms.Button; $btnTestVM.Text="Testen"; $btnTestVM.Location="550,65"; $btnTestVM.Size="100,60"
    $btnTestVM.Add_Click({
        $tIP = $txtIP.Text; $tUser = $txtUs.Text; $form.Cursor = "WaitCursor"
        if(-not (Test-Connection -ComputerName $tIP -Count 1 -Quiet)){ 
            $form.Cursor = "Default"; [System.Windows.Forms.MessageBox]::Show("Netzwerk-Fehler: Ping gescheitert.", "Fehler"); return
        }
        try {
            $testSSH = ssh -o BatchMode=yes -o ConnectTimeout=5 -o StrictHostKeyChecking=no "$tUser@$tIP" "echo ready" 2>&1
            $form.Cursor = "Default"
            if ($testSSH -match "ready") { [System.Windows.Forms.MessageBox]::Show("Verbindung Perfekt!`nPing & SSH OK.", "Erfolg") } 
            else { [System.Windows.Forms.MessageBox]::Show("Ping OK, aber SSH Login scheitert (Key fehlt?).`nFehler: $testSSH", "Warnung") }
        } catch { $form.Cursor="Default"; [System.Windows.Forms.MessageBox]::Show("Fehler: $_", "Error") }
    })
    $grpVM.Controls.AddRange(@($chkUseVM, $lblIP, $txtIP, $lblUs, $txtUs, $lblPw, $txtPw, $lblWarn, $lblPa, $txtPa, $btnTestVM))

    # CLEANUP BUTTON
    $grpClean = New-Object System.Windows.Forms.GroupBox; $grpClean.Text="Wartung"; $grpClean.Location="15,280"; $grpClean.Size="720,100"
    $lblClean = New-Object System.Windows.Forms.Label; $lblClean.Text="Löscht alle alten Backups (Ordner & ZIPs)."; $lblClean.Location="20,30"; $lblClean.Size="400,20"
    
    $btnCleanup = New-Object System.Windows.Forms.Button
    $btnCleanup.Text = "ALTE BACKUPS BEREINIGEN"
    $btnCleanup.Location = New-Object System.Drawing.Point(450, 25)
    $btnCleanup.Size = New-Object System.Drawing.Size(200, 50)
    $btnCleanup.BackColor = "DarkRed"
    $btnCleanup.ForeColor = "White"
    $btnCleanup.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $btnCleanup.Add_Click({ Start-Cleanup })
    
    $grpClean.Controls.AddRange(@($lblClean, $btnCleanup))
    
    $tab3.Controls.AddRange(@($grpVM, $grpClean))
    $tabs.TabPages.AddRange(@($tab1, $tab2, $tab3))
    
    $tabs.Add_SelectedIndexChanged({
        if ($tabs.SelectedIndex -eq 0) { $btnStart.Visible=$true; $btnSave.Visible=$true; $status.Visible=$true }
        else { $btnStart.Visible=$false; $btnSave.Visible=$false; $status.Visible=$false }
    })
    
    # ACTIONS
    $btnSave.Add_Click({
        $script:BackupConfig.SourcePaths = $txtSrc.Text -split "`r`n" | Where { $_.Trim() }
        $script:BackupConfig.DestinationPath = $txtDst.Text
        $script:BackupConfig.ExcludePaths = $txtEx.Text -split "[,`r`n]" | Where { $_.Trim() } | ForEach { $_.Trim() }
        $script:BackupConfig.FileTypes.Word = $chkW.Checked
        $script:BackupConfig.FileTypes.Excel = $chkE.Checked
        $script:BackupConfig.FileTypes.PowerPoint = $chkP.Checked
        $script:BackupConfig.FileTypes.PDF = $chkD.Checked
        $script:BackupConfig.UseVM = $chkUseVM.Checked
        $script:BackupConfig.VmIP = $txtIP.Text
        $script:BackupConfig.VmUser = $txtUs.Text
        $script:BackupConfig.VmPass = $txtPw.Text 
        $script:BackupConfig.VmDest = $txtPa.Text
        $script:BackupConfig.BackupMode = if($radF.Checked){"Full"}elseif($radI.Checked){"Incremental"}else{"Smart"}
        $script:BackupConfig.UseZip = $chkZip.Checked
        Save-Config
        $status.Text = "Einstellungen gespeichert."
    })
    
    $btnStart.Add_Click({
        $status.Text="Backup läuft..."; $form.Cursor="WaitCursor"; $form.Refresh()
        $src=$txtSrc.Text -split "`r`n"|Where{$_.Trim()}; $excl=$txtEx.Text -split "[,`r`n]"|Where{$_.Trim()}
        $types=@{Word=$chkW.Checked;Excel=$chkE.Checked;PowerPoint=$chkP.Checked;PDF=$chkD.Checked;Images=$false;All=$false}
        $mode=if($radF.Checked){"Full"}elseif($radI.Checked){"Incremental"}else{"Smart"}
        $compress = $chkZip.Checked 
        
        if(-not(Test-Path $txtDst.Text)){New-Item -ItemType Directory -Path $txtDst.Text -Force|Out-Null}
        
        $res = Start-BackupProcess -SourcePaths $src -DestinationPath $txtDst.Text -ExcludePaths $excl -FileTypes $types -BackupMode $mode -Compress $compress
        
        $form.Cursor="Default"
        if ($res.Success) {
            $msg = "Backup abgeschlossen!`nDateien: $($res.Files)"
            if ($res.Errors -gt 0) { $msg += "`nFehler (übersprungen): $($res.Errors)" }
            $msg += "`nPfad: $($res.Path)"
            
            $status.Text="Fertig: $($res.Files) Dateien."
            Show-Notification -Title "Backup Erfolgreich" -Message "Es wurden $($res.Files) Dateien gesichert." -Type "Info"
            
            if ($chkUseVM.Checked -and $res.Files -gt 0) {
                if ($res.Mode -eq "Full" -or $compress) {
                    $upl = [System.Windows.Forms.MessageBox]::Show($msg + "`n`nAuf VM (" + $txtIP.Text + ") hochladen?", "VM Upload", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    if ($upl -eq "Yes") {
                        $status.Text = "Lade auf VM hoch..."
                        $form.Refresh()
                        $vmRes = Start-VmUpload -LocalPath $res.Path -IP $txtIP.Text -User $txtUs.Text -Pass $txtPw.Text -RemotePath $txtPa.Text
                        [System.Windows.Forms.MessageBox]::Show($vmRes.Message, "Upload Status")
                        $status.Text = $vmRes.Message
                    }
                }
            } else { [System.Windows.Forms.MessageBox]::Show($msg, "Erfolg") }
            if ($btnRef -ne $null) { $btnRef.PerformClick() }
        } else { 
            $status.Text="Fehler beim Backup." 
            [System.Windows.Forms.MessageBox]::Show($res.Message, "Backup Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Show-Notification -Title "Backup Fehlgeschlagen" -Message "Fehler: $($res.Message)" -Type "Error"
        }
    })
    
    $form.Controls.AddRange(@($tabs, $btnStart, $btnSave, $status))
    $form.ShowDialog() | Out-Null
}

# SETUP (NUR BEIM ERSTEN MAL)
function Show-SetupWizard {
    $wiz = New-Object System.Windows.Forms.Form; $wiz.Text="Backup Setup"; $wiz.Size="500,600"; $wiz.StartPosition="CenterScreen"; $wiz.FormBorderStyle="FixedDialog"; $wiz.MaximizeBox=$false
    $l1=New-Object System.Windows.Forms.Label;$l1.Text="Backup Konfiguration";$l1.Font=New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold);$l1.Location="20,20";$l1.Size="400,30"
    
    $g1=New-Object System.Windows.Forms.GroupBox;$g1.Text="Quelle";$g1.Location="20,60";$g1.Size="440,70"
    $t1=New-Object System.Windows.Forms.TextBox;$t1.Text=[Environment]::GetFolderPath("MyDocuments");$t1.Location="20,30";$t1.Size="300,20"
    $b1=New-Object System.Windows.Forms.Button;$b1.Text="Wählen";$b1.Location="330,28";$b1.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$t1.Text=$d.SelectedPath}})
    $g1.Controls.AddRange(@($t1,$b1))
    
    $g2=New-Object System.Windows.Forms.GroupBox;$g2.Text="Ziel";$g2.Location="20,140";$g2.Size="440,70"
    $t2=New-Object System.Windows.Forms.TextBox;$t2.Text="C:\Backup";$t2.Location="20,30";$t2.Size="300,20"
    $b2=New-Object System.Windows.Forms.Button;$b2.Text="Wählen";$b2.Location="330,28";$b2.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$t2.Text=$d.SelectedPath}})
    $g2.Controls.AddRange(@($t2,$b2))
    
    $bs=New-Object System.Windows.Forms.Button;$bs.Text="Fertigstellen";$bs.BackColor="LightGreen";$bs.Location="300,500";$bs.Size="150,40"
    $bs.Add_Click({
        $script:BackupConfig=@{SourcePaths=@($t1.Text);DestinationPath=$t2.Text;ExcludePaths=@("TBZ");FileTypes=@{Word=$true;Excel=$true;PowerPoint=$true;PDF=$true;Images=$false;All=$false};BackupMode="Smart";UseZip=$false;UseVM=$false;VmIP="";VmUser="";VmPass="";VmDest="/home/kali/Backup"}
        Save-Config
        $wiz.Close()
    })
    $wiz.Controls.AddRange(@($l1,$g1,$g2,$bs))
    $wiz.ShowDialog()|Out-Null
}

if ($script:BackupConfig -eq $null) { Show-SetupWizard; if(Test-Path $ConfigFile){$LoadedConfig=Get-Content $ConfigFile -Raw|ConvertFrom-Json;$script:BackupConfig=$DefaultConfig.Clone();foreach($prop in $LoadedConfig.PSObject.Properties){if($script:BackupConfig.ContainsKey($prop.Name)){$script:BackupConfig[$prop.Name]=$prop.Value}};Show-MainGUI} } else { Show-MainGUI }