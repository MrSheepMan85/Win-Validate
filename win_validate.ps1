<#
.SYNOPSIS
    Outil de Validation Matérielle.
    Version : 4.4 (Full Stress - Détection Secteur intelligente)
    Layout  : Rapport Détaillé
.DESCRIPTION
    Si sur Batterie : Teste l'autonomie et la chauffe.
    Si sur Secteur  : Teste la stabilité thermique uniquement (ignore la note batterie).
#>

# =========================================================
#  PARAMÈTRES
# =========================================================
$STRESS_DURATION = 120      # Durée du stress (2 min)
$REF_CPU_SCORE   = 9.2
$REF_RAM_SCORE   = 9.2
$REF_GPU_SCORE   = 6.2
$REF_DISK_SCORE  = 8.6
# =========================================================

Add-Type -AssemblyName System.Windows.Forms

# 0. DÉTECTION DU DOSSIER
function Get-SmartLocation {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if (Get-Variable "ps2exeExecPath" -ErrorAction SilentlyContinue) { return [System.IO.Path]::GetDirectoryName($ps2exeExecPath) }
    try {
        $Proc = [System.Diagnostics.Process]::GetCurrentProcess()
        if ($Proc.MainModule.FileName -notmatch "powershell") { return [System.IO.Path]::GetDirectoryName($Proc.MainModule.FileName) }
    } catch {}
    return [Environment]::GetFolderPath("Desktop")
}

$RootDir = Get-SmartLocation
Set-Location -Path $RootDir -ErrorAction SilentlyContinue

$Date = Get-Date -Format "yyyy-MM-dd HH:mm"

# Vérification Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERREUR : Lancez en tant qu'Administrateur." -ForegroundColor Red; Read-Host "Entrée..."; Exit
}

# 1. IDENTIFICATION
try { $BiosInfo = Get-WmiObject Win32_BIOS -ErrorAction Stop; $ServiceTag = $BiosInfo.SerialNumber; $FormattedName = "FRALW-$ServiceTag" } catch { $FormattedName = "FRALW-Inconnu" }
try { $SysInfo = Get-WmiObject Win32_ComputerSystem; $ModelePC = $SysInfo.Model.Trim(); $Fabricant = $SysInfo.Manufacturer.Trim() } catch { $ModelePC = "Inconnu"; $Fabricant = "" }

$SafeModelName = $ModelePC -replace '[\\/:\*\?"<>\|]', ''
$ModelDir = "$RootDir\$SafeModelName"
if (-not (Test-Path $ModelDir)) { try { New-Item -ItemType Directory -Force -Path $ModelDir | Out-Null } catch { $ModelDir = $RootDir } }

$RapportPath = "$ModelDir\$FormattedName.txt"
$CsvPath = "$RootDir\Inventaire_Parc_Global.csv"
$BatReportPath = "$ModelDir\battery_temp.xml"

Clear-Host
Write-Host "============================================================" -ForegroundColor Red
Write-Host "   DIAGNOSTIC MATÉRIEL (v4.4 - SMART STRESS)                " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Red
Write-Host "Machine : $FormattedName ($ModelePC)" -ForegroundColor Green
Write-Host "Démarrage..."

# 2. INFO MATÉRIEL
try {
    $ProcInfo = Get-WmiObject Win32_Processor | Select-Object -First 1; $InfoCPU = $ProcInfo.Name.Trim()
    $Cores = $ProcInfo.NumberOfLogicalProcessors
    $MemInfo = Get-WmiObject Win32_ComputerSystem; $RawRam = $MemInfo.TotalPhysicalMemory; $InfoRAM = "$([math]::Round($RawRam / 1GB, 1)) GB"
    $GpuInfo = Get-WmiObject Win32_VideoController | Select-Object -First 1; $InfoGPU = $GpuInfo.Name
    $DiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" } | Select-Object -First 1
    $InfoDisk = "$($DiskInfo.Model) ($([math]::Round($DiskInfo.Size / 1GB, 0)) GB)"
} catch { $InfoCPU="N/A"; $Cores=2; $InfoRAM="N/A"; $InfoGPU="N/A"; $InfoDisk="N/A" }

# 3. BASELINE (WinSAT)
Write-Host "[1/5] Benchmark de référence (WinSAT)..." -ForegroundColor Green
Start-Process winsat -ArgumentList "formal -restart" -Wait -WindowStyle Hidden
$WinSATPath = "$env:WINDIR\Performance\WinSAT\DataStore"
$LatestXML = Get-ChildItem -Path "$WinSATPath\*Formal.Assessment*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($LatestXML) {
    [xml]$Data = Get-Content $LatestXML.FullName
    try {
        $ScoreCPU = [double]$Data.WinSAT.WinSPR.CpuScore
        $ScoreRAM = [double]$Data.WinSAT.WinSPR.MemoryScore
        $ScoreGPU = [double]$Data.WinSAT.WinSPR.GraphicsScore
        $ScoreD3D = [double]$Data.WinSAT.WinSPR.GamingScore 
        $ScoreDisk = [double]$Data.WinSAT.WinSPR.DiskScore
    } catch { $ScoreCPU=0; $ScoreRAM=0; $ScoreGPU=0; $ScoreD3D=0; $ScoreDisk=0 }
} else { $ScoreCPU=0; $ScoreRAM=0; $ScoreGPU=0; $ScoreD3D=0; $ScoreDisk=0 }

function Calculate-RelativeScore ($current, $target) {
    if (!$current -or $current -eq 0) { return 0 }
    $val = ($current / $target) * 20
    if ($val -gt 20) { return 20 }
    return [math]::Round($val, 1)
}
$NoteCPU = Calculate-RelativeScore $ScoreCPU $REF_CPU_SCORE
$NoteRAM = Calculate-RelativeScore $ScoreRAM $REF_RAM_SCORE
$NoteGPU = Calculate-RelativeScore $ScoreD3D $REF_GPU_SCORE
$NoteDisk = Calculate-RelativeScore $ScoreDisk $REF_DISK_SCORE

# 4. SMART
Write-Host "[2/5] Santé Disque (S.M.A.R.T)..." -ForegroundColor Green
try {
    $PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1; $SmartStatus = $PhysicalDisk.HealthStatus
    if ($SmartStatus -eq "Healthy") { $NoteSmart = 20; $SmartMsg = "Disque Sain" } else { $NoteSmart = 0; $SmartMsg = "DANGER : $SmartStatus" }
} catch { $NoteSmart = "N/A"; $SmartMsg = "Inconnu" }

# 5. RÉSEAU
Write-Host "[3/5] Test Connectivité..." -ForegroundColor Green
if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $NetStatus = "Connecté Internet" } else { $NetStatus = "Pas d'Internet" }

# 6. FULL STRESS TEST INTELLIGENT
Write-Host "[4/5] FULL STRESS TEST (CPU+GPU+DISK)..." -ForegroundColor Green

# A. Batterie & État Secteur
if (Test-Path $BatReportPath) { Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue }
Start-Process powercfg -ArgumentList "/batteryreport /output `"$BatReportPath`" /XML" -Wait -WindowStyle Hidden
Start-Sleep 1
$SantePercent="N/A"; $NoteSanteChimique="N/A"; $InfoBat="Batterie non détectée"

if (Test-Path $BatReportPath) {
    try {
        [xml]$BatXML = Get-Content $BatReportPath
        $ActiveBat = $BatXML.BatteryReport.Batteries.Battery | Select-Object -First 1
        if ($ActiveBat -and $ActiveBat.DesignCapacity -gt 0) {
            $DesignCap = [double]$ActiveBat.DesignCapacity; $FullCap = [double]$ActiveBat.FullChargeCapacity
            $SantePercent = ($FullCap / $DesignCap) * 100
            $NoteSanteChimique = ($SantePercent / 5)
            $InfoBat = "$([math]::Round($FullCap,0)) mWh / $([math]::Round($DesignCap,0)) mWh"
        }
    } catch {}
    Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue
}

# Détection Secteur
$PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus # Online = Secteur, Offline = Batterie
$IsPluggedIn = ($PowerStatus -eq "Online")

# B. Lancement Stress (Commun aux deux cas pour tester la stabilité thermique)
$StartP = 0; try { $StartP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch {}
if ($IsPluggedIn) { Write-Host "      > SECTEUR DÉTECTÉ : Test stabilité thermique uniquement (Pas de note batterie)." -ForegroundColor Cyan }
else { Write-Host "      > BATTERIE DÉTECTÉE ($StartP%) : Test complet (Autonomie + Chauffe)." -ForegroundColor Yellow }

Write-Host "      > Charge maximale en cours ($STRESS_DURATION sec)..." -ForegroundColor Red

$Jobs = @()
if (!$Cores) { $Cores = 2 }
1..$Cores | ForEach-Object { $Jobs += Start-Job -ScriptBlock { $r=1; while($true) { $r += [math]::Tan($r) } } } # CPU
$Jobs += Start-Job -ScriptBlock { $f="$env:TEMP\io.dat"; $d=[byte[]]::new(50*1MB); (new-object Random).NextBytes($d); while($true){[IO.File]::WriteAllBytes($f,$d);$null=[IO.File]::ReadAllBytes($f)} } # DISK
$Jobs += Start-Job -ScriptBlock { while($true) { Start-Process winsat -ArgumentList "d3d -objs C(20) -duration 5" -Wait -WindowStyle Hidden } } # GPU

# Attente
for($i=0; $i -lt $STRESS_DURATION; $i++) { 
    $PercentDone = [math]::Round(($i / $STRESS_DURATION) * 100)
    Write-Progress -Activity "FULL STRESS" -Status "Ventilation au max..." -PercentComplete $PercentDone
    Start-Sleep 1 
}
Write-Progress -Activity "FULL STRESS" -Completed
Get-Job | Stop-Job | Remove-Job
try { Remove-Item "$env:TEMP\io.dat" -ErrorAction SilentlyContinue } catch {}

# C. Notation Différenciée
$DropPercent = 0
$NoteStress = "N/A"
$BatteryMessage = "Non testé"

if ($IsPluggedIn) {
    # Cas Secteur
    $BatteryMessage = "Ignoré (Sur Secteur)"
    $NoteStress = "N/A" # On ne note pas la batterie
} else {
    # Cas Batterie
    $EndP = 0; try { $EndP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch { $EndP = $StartP }
    $DropPercent = $StartP - $EndP
    $BatteryMessage = "Perte: $DropPercent% (Charge Max)"
    
    if ($DropPercent -le 2) { $NoteStress = 20 }
    elseif ($DropPercent -le 4) { $NoteStress = 15 }
    elseif ($DropPercent -le 6) { $NoteStress = 10 }
    elseif ($DropPercent -le 8) { $NoteStress = 5 }
    else { $NoteStress = 0 }
}

# Calcul Note Batterie Finale
if ($NoteSanteChimique -ne "N/A" -and $NoteStress -ne "N/A") {
    $NoteBatFinale = [math]::Round(($NoteSanteChimique + $NoteStress) / 2, 1)
} elseif ($NoteSanteChimique -ne "N/A") {
    $NoteBatFinale = $NoteSanteChimique # On garde que la chimie si sur secteur
} else {
    $NoteBatFinale = "N/A" # Pas de batterie du tout (PC Fixe)
}

# 7. VERDICT GLOBAL
$MoyennePerf = ($NoteCPU + $NoteRAM + $NoteGPU + $NoteDisk) / 4

if ($NoteBatFinale -eq "N/A") {
    # PC Fixe ou Secteur sans batterie lisible -> 100% Perf
    $NoteGlobale = [math]::Round($MoyennePerf, 1); $FlagProbleme = $false
} else {
    # Laptop -> 50% Perf / 50% Batterie
    $NoteGlobale = [math]::Round(($MoyennePerf * 0.5) + ($NoteBatFinale * 0.5), 1)
    $FlagProbleme = $false; $RaisonProbleme = ""
    
    # Critères
    if ($SantePercent -ne "N/A" -and $SantePercent -lt 50) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE HS" }
    if ($NoteStress -ne "N/A" -and $DropPercent -ge 7) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE INSTABLE" }
    if ($NoteSmart -eq 0) { $FlagProbleme = $true; $RaisonProbleme = "DISQUE HS" }
    
    if ($FlagProbleme) { 
        if ($NoteGlobale -gt 9) { $NoteGlobale = 9 }
        $Verdict = "RECALÉ ($RaisonProbleme)" 
    } else {
        if ($NoteGlobale -ge 16) { $Verdict = "PREMIUM" } 
        elseif ($NoteGlobale -ge 14) { $Verdict = "BON ÉTAT" } 
        elseif ($NoteGlobale -ge 11) { $Verdict = "STANDARD" } 
        else { $Verdict = "FAIBLE" }
    }
}

# 8. RAPPORT
Write-Host "[5/5] Génération Rapport..." -ForegroundColor Green
$SanteDisplay = if($SantePercent -ne "N/A"){"$([math]::Round($SantePercent,0))%"}else{"N/A"}
$StressDisplay = if($NoteStress -ne "N/A"){"$NoteStress / 20"}else{"Non Noté (Secteur)"}
$ChimiqueDisplay = if($NoteSanteChimique -ne "N/A"){"$([math]::Round($NoteSanteChimique,1))"}else{"N/A"}

$NewContent = @"
============================================================
       RAPPORT DIAGNOSTIC - $Date
============================================================
Modèle     : $Fabricant $ModelePC
Réseau     : $NetStatus
------------------------------------------------------------
   >>> NOTE GLOBALE : $NoteGlobale / 20
   >>> VERDICT      : $Verdict
------------------------------------------------------------

 1. PROCESSEUR (CPU)
------------------------------------------------------------
 Matériel : $InfoCPU
 Note     : $NoteCPU / 20 (Score: $ScoreCPU)

 2. MÉMOIRE (RAM)
------------------------------------------------------------
 Matériel : $InfoRAM
 Note     : $NoteRAM / 20 (Score: $ScoreRAM)

 3. GRAPHIQUE (GPU)
------------------------------------------------------------
 Matériel : $InfoGPU
 Note     : $NoteGPU / 20 (Score: $ScoreD3D)

 4. STOCKAGE & SANTÉ
------------------------------------------------------------
 Matériel : $InfoDisk
 Note Perf: $NoteDisk / 20
 État     : $SmartMsg (Note: $NoteSmart/20)

 5. BATTERIE & STRESS
------------------------------------------------------------
 Note Finale : $NoteBatFinale / 20
 Santé (SOH) : $SanteDisplay (Note Chimique: $ChimiqueDisplay/20)
 Tenue Stress: $StressDisplay ($BatteryMessage)
 Détail      : $InfoBat

============================================================
"@

try {
    if (Test-Path $RapportPath) { $OldContent = Get-Content $RapportPath -Raw; $FinalContent = $NewContent + "`r`n`r`n`r`n" + $OldContent } else { $FinalContent = $NewContent }
    $FinalContent | Out-File -FilePath $RapportPath -Encoding UTF8 -Force -ErrorAction Stop
    Invoke-Item $RapportPath
} catch { Write-Host "Erreur Rapport." -ForegroundColor Red }

if (-not (Test-Path $CsvPath)) { try { "Date;MachineID;Modele;Note_Globale;Verdict;CPU_Score;Bat_Sante;Bat_Drop" | Out-File $CsvPath -Encoding UTF8 -ErrorAction Stop } catch {} }
try { "$Date;$FormattedName;$ModelePC;$NoteGlobale;$Verdict;$ScoreCPU;$SanteDisplay;$DropPercent" | Out-File $CsvPath -Append -Encoding UTF8 -ErrorAction Stop } catch {}

Write-Host ""
Read-Host "Terminé. Entrée..."
