# ================================================================
# Office Backup System - Final Edition v4.2 (UI Polish & Logic)
# ================================================================
# Fix: Buttons nur auf dem Haupt-Tab sichtbar
# Fix: Button-Texte ausgeschrieben und Größen angepasst
# Feature: Passwort-Feld für VM (Hinweis: SSH-Key empfohlen)
# Autor: Backup System Generator & M122
# Version: 4.2 Final
# ================================================================

# .NET-Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- KONFIGURATIONSMANAGEMENT ---
if ($PSScriptRoot) { $BaseDir = $PSScriptRoot } else { $BaseDir = $PWD.Path }
$ConfigFile = Join-Path $BaseDir "backup_config.json"

$DefaultConfig = @{
    SourcePaths = @([Environment]::GetFolderPath("MyDocuments"))
    DestinationPath = "C:\Backup"
    ExcludePaths = @("TBZ", "Temp")
    FileTypes = @{ Word=$true; Excel=$true; PowerPoint=$true; PDF=$true; Images=$false; All=$false }
    BackupMode = "Smart"
    UseVM = $false
    VmIP = ""
    VmUser = ""
    VmPass = "" # Neu: Passwort
    VmDest = "/home/user/backup"
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

function Update-FileTimestamp { param([string]$f); try{(Get-Item $f).LastWriteTime=Get-Date;$true}catch{$false} }

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
        $path = Join-Path (Join-Path (Join-Path $fol.FullName $Category) ($RelativePath -replace "^\\","")) $File.Name
        if (Test-Path $path) {
            $bf = Get-Item $path
            if ($File.Length -ne $bf.Length) { return @{NeedsBackup=$true; Reason="Größe"} }
            if ([Math]::Abs(($File.LastWriteTime - $bf.LastWriteTime).TotalSeconds) -gt 3) { return @{NeedsBackup=$true; Reason="Zeit"} }
            return @{NeedsBackup=$false; Reason="Vorhanden"}
        }
    }
    return @{NeedsBackup=$true; Reason="Neu"}
}

function Start-VmUpload {
    param($LocalPath, $IP, $User, $Pass, $RemotePath)
    if (-not (Test-Connection -ComputerName $IP -Count 1 -Quiet)) { return @{ Success=$false; Message="VM Offline" } }
    
    # Hinweis: Windows OpenSSH unterstützt keine direkte Passwort-Übergabe per Skript (Sicherheitsfeature).
    # Wir nutzen hier den Standard-Weg. Wenn kein Key da ist, fragt Windows ggf. im Hintergrund oder scheitert.
    try {
        ssh -o ConnectTimeout=5 -o StrictHostKeyChecking=no "$User@$IP" "mkdir -p $RemotePath"
        if ($LASTEXITCODE -ne 0) { return @{ Success=$false; Message="SSH Login fehlgeschlagen (Key prüfen!)" } }
        
        scp -r -p -o StrictHostKeyChecking=no "$LocalPath" "$($User)@$($IP):$RemotePath"
        if ($LASTEXITCODE -eq 0) { return @{ Success=$true; Message="Upload OK" } }
        else { return @{ Success=$false; Message="SCP Fehler" } }
    } catch { return @{ Success=$false; Message="Fehler: $_" } }
}

function Start-BackupProcess {
    param($SourcePaths, $DestinationPath, $ExcludePaths, $FileTypes, $BackupMode)
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
    $root = Join-Path $DestinationPath "Backup_${date}_${suffix}"
    
    $exts = @{}
    if($FileTypes.Word){$exts['Word']=@('*.docx','*.doc')}
    if($FileTypes.Excel){$exts['Excel']=@('*.xlsx','*.xls')}
    if($FileTypes.PowerPoint){$exts['PowerPoint']=@('*.pptx','*.ppt')}
    if($FileTypes.PDF){$exts['PDF']=@('*.pdf')}
    
    $stats = @{Copied=0; Checked=0}
    $list = @()
    
    foreach($p in $SourcePaths){
        $src=$p.TrimEnd('\'); if(-not(Test-Path $src)){continue}
        foreach($cat in $exts.Keys){
            foreach($ext in $exts[$cat]){
                Get-ChildItem -Path $src -Filter $ext -Recurse -File -EA SilentlyContinue | ForEach {
                    $stats.Checked++
                    $f=$_; $skip=$false
                    foreach($ex in $cleanExcludes){if($f.FullName -like "*\$ex\*" -or $f.FullName -like "$ex\*"){$skip=$true;break}}
                    if($skip){return}
                    
                    $rel=if($f.DirectoryName.StartsWith($src)){$f.DirectoryName.Substring($src.Length).TrimStart('\')}else{""}
                    $chk=if($BackupMode -eq "Full"){@{NeedsBackup=$true}}else{Test-FileNeedsBackup -File $f -BackupChain $backupChain -Category $cat -RelativePath $rel}
                    if($chk.NeedsBackup){$list+=@{F=$f;C=$cat;R=$rel}}
                }
            }
        }
    }
    
    if($list.Count -gt 0){
        New-Item -ItemType Directory -Path $root -Force|Out-Null
        foreach($i in $list){
            $dst=Join-Path (Join-Path (Join-Path $root $i.C) $i.R) $i.F.Name
            $null=New-Item -ItemType Directory -Path (Split-Path $dst) -Force
            Copy-Item $i.F.FullName $dst -Force
            $stats.Copied++
        }
        return @{Success=$true; Files=$stats.Copied; Path=$root; Mode=$BackupMode}
    }
    return @{Success=$true; Files=0; Path=$null; Mode=$BackupMode}
}

# ================================================================
# GUI (LOGIK & DESIGN)
# ================================================================

function Show-MainGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office Backup System v4.2"; $form.Size="800,850"; $form.StartPosition="CenterScreen"; $form.FormBorderStyle="FixedDialog"; $form.MaximizeBox=$false
    
    $tabs = New-Object System.Windows.Forms.TabControl; $tabs.Size="760,680"; $tabs.Location="10,10"
    
    # BUTTONS DEFINIEREN (DAMIT WIR SIE SPÄTER EIN/AUSBLENDEN KÖNNEN)
    $btnStart = New-Object System.Windows.Forms.Button; $btnStart.Text="BACKUP STARTEN"; $btnStart.BackColor="LightGreen"; $btnStart.Location="15,710"; $btnStart.Size="200,50"; $btnStart.Font=New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text="Einstellungen Speichern"; $btnSave.Location="230,720"; $btnSave.Size="180,30" # HIER: Breiter gemacht!
    $status = New-Object System.Windows.Forms.Label; $status.Text="Bereit"; $status.Location="420,725"; $status.Size="350,20"; $status.ForeColor="Blue"

    # --- TAB 1: BACKUP ---
    $tab1 = New-Object System.Windows.Forms.TabPage; $tab1.Text = "📁 Backup Einstellungen"
    
    # 1. Quelle
    $grpSrc = New-Object System.Windows.Forms.GroupBox; $grpSrc.Text="1. Quellordner"; $grpSrc.Location="15,15"; $grpSrc.Size="720,110"
    $txtSrc = New-Object System.Windows.Forms.TextBox; $txtSrc.Multiline=$true; $txtSrc.ScrollBars="Vertical"; $txtSrc.Location="15,25"; $txtSrc.Size="570,70"
    $txtSrc.Text = $script:BackupConfig.SourcePaths -join "`r`n"
    $btnAdd = New-Object System.Windows.Forms.Button; $btnAdd.Text="Hinzufügen"; $btnAdd.Location="600,25"; $btnAdd.Size="100,30"
    $btnAdd.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtSrc.Text+="`r`n"+$d.SelectedPath}})
    $grpSrc.Controls.AddRange(@($txtSrc, $btnAdd))
    
    # 2. Ziel
    $grpDst = New-Object System.Windows.Forms.GroupBox; $grpDst.Text="2. Zielordner"; $grpDst.Location="15,140"; $grpDst.Size="720,70"
    $txtDst = New-Object System.Windows.Forms.TextBox; $txtDst.Location="15,30"; $txtDst.Size="570,25"; $txtDst.Text = $script:BackupConfig.DestinationPath
    $btnDst = New-Object System.Windows.Forms.Button; $btnDst.Text="Suchen..."; $btnDst.Location="600,28"; $btnDst.Size="100,27"
    $btnDst.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$txtDst.Text=$d.SelectedPath}})
    $grpDst.Controls.AddRange(@($txtDst, $btnDst))
    
    # 3. Ausnahmen
    $grpEx = New-Object System.Windows.Forms.GroupBox; $grpEx.Text="3. Ausnahmen"; $grpEx.Location="15,225"; $grpEx.Size="720,70"; $grpEx.ForeColor="DarkRed"
    $txtEx = New-Object System.Windows.Forms.TextBox; $txtEx.Location="15,30"; $txtEx.Size="570,25"
    $txtEx.Text = $script:BackupConfig.ExcludePaths -join ", "
    $btnEx = New-Object System.Windows.Forms.Button; $btnEx.Text="Suchen..."; $btnEx.Location="600,28"; $btnEx.Size="100,27"
    $btnEx.Add_Click({$d=New-Object System.Windows.Forms.FolderBrowserDialog;if($d.ShowDialog()-eq"OK"){$n=Split-Path $d.SelectedPath -Leaf;if($txtEx.Text){$txtEx.Text+=", "+$n}else{$txtEx.Text=$n}}})
    $grpEx.Controls.AddRange(@($txtEx, $btnEx))
    
    # 4. Typen & Modus
    $grpTyp = New-Object System.Windows.Forms.GroupBox; $grpTyp.Text="4. Dateitypen"; $grpTyp.Location="15,310"; $grpTyp.Size="720,80"
    $chkW=New-Object System.Windows.Forms.CheckBox;$chkW.Text="Word";$chkW.Location="20,25";$chkW.Checked=$script:BackupConfig.FileTypes.Word
    $chkE=New-Object System.Windows.Forms.CheckBox;$chkE.Text="Excel";$chkE.Location="150,25";$chkE.Checked=$script:BackupConfig.FileTypes.Excel
    $chkP=New-Object System.Windows.Forms.CheckBox;$chkP.Text="PowerPoint";$chkP.Location="300,25";$chkP.Checked=$script:BackupConfig.FileTypes.PowerPoint
    $chkD=New-Object System.Windows.Forms.CheckBox;$chkD.Text="PDF";$chkD.Location="450,25";$chkD.Checked=$script:BackupConfig.FileTypes.PDF
    $grpTyp.Controls.AddRange(@($chkW,$chkE,$chkP,$chkD))
    
    $grpMod = New-Object System.Windows.Forms.GroupBox; $grpMod.Text="5. Modus"; $grpMod.Location="15,400"; $grpMod.Size="720,80"
    $radS=New-Object System.Windows.Forms.RadioButton;$radS.Text="Smart (Auto)";$radS.Location="20,30";$radS.Size="120,20";$radS.Checked=($script:BackupConfig.BackupMode -eq "Smart")
    $radF=New-Object System.Windows.Forms.RadioButton;$radF.Text="Vollständig";$radF.Location="150,30";$radF.Size="120,20";$radF.Checked=($script:BackupConfig.BackupMode -eq "Full")
    $radI=New-Object System.Windows.Forms.RadioButton;$radI.Text="Inkrementell";$radI.Location="300,30";$radI.Size="120,20";$radI.Checked=($script:BackupConfig.BackupMode -eq "Incremental")
    $grpMod.Controls.AddRange(@($radS,$radF,$radI))
    
    $tab1.Controls.AddRange(@($grpSrc, $grpDst, $grpEx, $grpTyp, $grpMod))
    
    # --- TAB 2: HISTORY ---
    $tab2 = New-Object System.Windows.Forms.TabPage; $tab2.Text = "📜 Historie"
    $lstH = New-Object System.Windows.Forms.ListBox; $lstH.Location="15,15"; $lstH.Size="720,600"; $lstH.Font="Consolas,9"
    $btnRef = New-Object System.Windows.Forms.Button; $btnRef.Text="Aktualisieren"; $btnRef.Location="15,620"; $btnRef.Size="150,30"
    $btnRef.Add_Click({$lstH.Items.Clear();if(Test-Path $txtDst.Text){Get-ChildItem $txtDst.Text -Dir|Where Name -match "^Backup_"|Sort Name -Desc|ForEach{$lstH.Items.Add($_.Name)}}})
    $tab2.Controls.AddRange(@($lstH, $btnRef))
    
    # --- TAB 3: TOOLS / VM ---
    $tab3 = New-Object System.Windows.Forms.TabPage; $tab3.Text = "🔧 Tools & VM"
    
    $grpVM = New-Object System.Windows.Forms.GroupBox; $grpVM.Text="Kali VM Verbindung"; $grpVM.Location="15,15"; $grpVM.Size="720,250"
    $chkUseVM = New-Object System.Windows.Forms.CheckBox; $chkUseVM.Text="VM für Uploads nutzen"; $chkUseVM.Location="20,30"; $chkUseVM.Size="200,20"; $chkUseVM.Checked=$script:BackupConfig.UseVM
    
    # Zeile 1
    $lblIP = New-Object System.Windows.Forms.Label; $lblIP.Text="IP Adresse:"; $lblIP.Location="20,70"; $lblIP.Size="80,20"
    $txtIP = New-Object System.Windows.Forms.TextBox; $txtIP.Text=$script:BackupConfig.VmIP; $txtIP.Location="110,67"; $txtIP.Size="150,20"
    
    $lblUs = New-Object System.Windows.Forms.Label; $lblUs.Text="Benutzer:"; $lblUs.Location="280,70"; $lblUs.Size="70,20"
    $txtUs = New-Object System.Windows.Forms.TextBox; $txtUs.Text=$script:BackupConfig.VmUser; $txtUs.Location="360,67"; $txtUs.Size="150,20"
    
    # Zeile 2 (NEU: Passwort)
    $lblPw = New-Object System.Windows.Forms.Label; $lblPw.Text="Passwort:"; $lblPw.Location="20,110"; $lblPw.Size="80,20"
    $txtPw = New-Object System.Windows.Forms.TextBox; $txtPw.Text=$script:BackupConfig.VmPass; $txtPw.Location="110,107"; $txtPw.Size="150,20"; $txtPw.PasswordChar="*"
    $lblWarn = New-Object System.Windows.Forms.Label; $lblWarn.Text="(Optional. SSH-Key empfohlen!)"; $lblWarn.Location="270,110"; $lblWarn.Size="250,20"; $lblWarn.ForeColor="Gray"
    
    # Zeile 3
    $lblPa = New-Object System.Windows.Forms.Label; $lblPa.Text="Zielpfad:"; $lblPa.Location="20,150"; $lblPa.Size="80,20"
    $txtPa = New-Object System.Windows.Forms.TextBox; $txtPa.Text=$script:BackupConfig.VmDest; $txtPa.Location="110,147"; $txtPa.Size="400,20"
    
    $btnTestVM = New-Object System.Windows.Forms.Button; $btnTestVM.Text="Testen"; $btnTestVM.Location="550,65"; $btnTestVM.Size="100,60"
    
    $btnTestVM.Add_Click({
        if(Test-Connection -ComputerName $txtIP.Text -Count 1 -Quiet){ [System.Windows.Forms.MessageBox]::Show("VM Online! (SSH Login noch ungeprüft)", "OK") }
        else { [System.Windows.Forms.MessageBox]::Show("VM nicht erreichbar.", "Fehler") }
    })
    
    $grpVM.Controls.AddRange(@($chkUseVM, $lblIP, $txtIP, $lblUs, $txtUs, $lblPw, $txtPw, $lblWarn, $lblPa, $txtPa, $btnTestVM))
    $tab3.Controls.Add($grpVM)
    
    $tabs.TabPages.AddRange(@($tab1, $tab2, $tab3))
    
    # LOGIK: BUTTONS NUR AUF TAB 1 ANZEIGEN
    $tabs.Add_SelectedIndexChanged({
        if ($tabs.SelectedIndex -eq 0) {
            $btnStart.Visible = $true
            $btnSave.Visible = $true
            $status.Visible = $true
        } else {
            $btnStart.Visible = $false
            $btnSave.Visible = $false
            $status.Visible = $false
        }
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
        $script:BackupConfig.VmPass = $txtPw.Text # Passwort speichern
        $script:BackupConfig.VmDest = $txtPa.Text
        $script:BackupConfig.BackupMode = if($radF.Checked){"Full"}elseif($radI.Checked){"Incremental"}else{"Smart"}
        Save-Config
        $status.Text = "Einstellungen gespeichert."
    })
    
    $btnStart.Add_Click({
        $status.Text="Backup läuft..."; $form.Cursor="WaitCursor"; $form.Refresh()
        $src=$txtSrc.Text -split "`r`n"|Where{$_.Trim()}; $excl=$txtEx.Text -split "[,`r`n]"|Where{$_.Trim()}
        $types=@{Word=$chkW.Checked;Excel=$chkE.Checked;PowerPoint=$chkP.Checked;PDF=$chkD.Checked;Images=$false;All=$false}
        $mode=if($radF.Checked){"Full"}elseif($radI.Checked){"Incremental"}else{"Smart"}
        
        if(-not(Test-Path $txtDst.Text)){New-Item -ItemType Directory -Path $txtDst.Text -Force|Out-Null}
        
        $res = Start-BackupProcess -SourcePaths $src -DestinationPath $txtDst.Text -ExcludePaths $excl -FileTypes $types -BackupMode $mode
        
        $form.Cursor="Default"
        if ($res.Success) {
            $status.Text="Fertig: $($res.Files) Dateien."
            $msg = "Backup abgeschlossen!`nDateien: $($res.Files)"
            
            # VM Logic
            if ($chkUseVM.Checked -and $res.Mode -eq "Full" -and $res.Files -gt 0) {
                $upl = [System.Windows.Forms.MessageBox]::Show($msg + "`n`nFull-Backup auf VM (" + $txtIP.Text + ") hochladen?", "VM Upload", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($upl -eq "Yes") {
                    $status.Text = "Lade auf VM hoch..."
                    $form.Refresh()
                    $vmRes = Start-VmUpload -LocalPath $res.Path -IP $txtIP.Text -User $txtUs.Text -Pass $txtPw.Text -RemotePath $txtPa.Text
                    [System.Windows.Forms.MessageBox]::Show($vmRes.Message, "Upload Status")
                    $status.Text = $vmRes.Message
                }
            } else { [System.Windows.Forms.MessageBox]::Show($msg, "Erfolg") }
            & $btnRef.PerformClick()
        } else { $status.Text="Fehler beim Backup." }
    })
    
    $form.Controls.AddRange(@($tabs, $btnStart, $btnSave, $status))
    $form.ShowDialog() | Out-Null
}

# SETUP (NUR BEIM ERSTEN MAL)
function Show-SetupWizard {
    $wiz = New-Object System.Windows.Forms.Form; $wiz.Text="Backup Setup"; $wiz.Size="500,600"; $wiz.StartPosition="CenterScreen"; $wiz.FormBorderStyle="FixedDialog"; $wiz.MaximizeBox=$false
    $l1=New-Object System.Windows.Forms.Label;$l1.Text="Backup Konfiguration";$l1.Font=New-Object System.Drawing.Font("Arial",12,1);$l1.Location="20,20";$l1.Size="400,30"
    
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
        $script:BackupConfig=@{SourcePaths=@($t1.Text);DestinationPath=$t2.Text;ExcludePaths=@("TBZ");FileTypes=@{Word=$true;Excel=$true;PowerPoint=$true;PDF=$true;Images=$false;All=$false};BackupMode="Smart";UseVM=$false;VmIP="";VmUser="";VmPass="";VmDest="/home/user"}
        Save-Config
        $wiz.Close()
    })
    $wiz.Controls.AddRange(@($l1,$g1,$g2,$bs))
    $wiz.ShowDialog()|Out-Null
}

if ($script:BackupConfig -eq $null) { Show-SetupWizard; if(Test-Path $ConfigFile){$LoadedConfig=Get-Content $ConfigFile -Raw|ConvertFrom-Json;$script:BackupConfig=$DefaultConfig.Clone();foreach($prop in $LoadedConfig.PSObject.Properties){if($script:BackupConfig.ContainsKey($prop.Name)){$script:BackupConfig[$prop.Name]=$prop.Value}};Show-MainGUI} } else { Show-MainGUI }