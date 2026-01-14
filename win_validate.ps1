<#
.SYNOPSIS
    Outil de Validation Matérielle.
    Version : 4.1 (Automatique - Historique Excel)
.DESCRIPTION
    Lance les tests sans poser de questions.
    Stocke l'historique dans Excel (ajout de ligne) et dans le TXT (ajout en haut).
#>

# =========================================================
#  CALIBRATION
# =========================================================
$REF_CPU_SCORE  = 9.2
$REF_RAM_SCORE  = 9.2
$REF_GPU_SCORE  = 6.2
$REF_DISK_SCORE = 8.6
# =========================================================

Add-Type -AssemblyName System.Windows.Forms

# 0. DÉTECTION INTELLIGENTE DU DOSSIER
function Get-SmartLocation {
    # 1. Mode Script
    if ($PSScriptRoot) { return $PSScriptRoot }
    # 2. Mode Compilé PS2EXE
    if (Get-Variable "ps2exeExecPath" -ErrorAction SilentlyContinue) { return [System.IO.Path]::GetDirectoryName($ps2exeExecPath) }
    # 3. Mode Processus
    try {
        $Proc = [System.Diagnostics.Process]::GetCurrentProcess()
        if ($Proc.MainModule.FileName -notmatch "powershell") { return [System.IO.Path]::GetDirectoryName($Proc.MainModule.FileName) }
    } catch {}
    # 4. Fallback Bureau
    return [Environment]::GetFolderPath("Desktop")
}

$RootDir = Get-SmartLocation
Set-Location -Path $RootDir -ErrorAction SilentlyContinue

$Date = Get-Date -Format "yyyy-MM-dd HH:mm"

# Vérification Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERREUR : Lancez en tant qu'Administrateur." -ForegroundColor Red; Read-Host "Entrée..."; Exit
}

# 1. IDENTIFICATION & ARBORESCENCE
try { $BiosInfo = Get-WmiObject Win32_BIOS -ErrorAction Stop; $ServiceTag = $BiosInfo.SerialNumber; $FormattedName = "FRALW-$ServiceTag" } catch { $FormattedName = "FRALW-Inconnu" }
try { $SysInfo = Get-WmiObject Win32_ComputerSystem; $ModelePC = $SysInfo.Model.Trim(); $Fabricant = $SysInfo.Manufacturer.Trim() } catch { $ModelePC = "Inconnu"; $Fabricant = "" }

# Création du dossier Modèle
$SafeModelName = $ModelePC -replace '[\\/:\*\?"<>\|]', ''
$ModelDir = "$RootDir\$SafeModelName"

if (-not (Test-Path $ModelDir)) { 
    try { New-Item -ItemType Directory -Force -Path $ModelDir | Out-Null } catch { $ModelDir = $RootDir } 
}

# CHEMINS
$RapportPath = "$ModelDir\$FormattedName.txt"
$CsvPath = "$RootDir\Inventaire_Parc_Global.csv"
$BatReportPath = "$ModelDir\battery_temp.xml"

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   DIAGNOSTIC MATÉRIEL (v4.1 - Auto)                        " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Machine : $FormattedName" -ForegroundColor Green
Write-Host "Modèle  : $ModelePC" -ForegroundColor Green
Write-Host "Dossier : $SafeModelName" -ForegroundColor Yellow
Write-Host "Démarrage..."

# 2. INFO MATÉRIEL
try {
    $ProcInfo = Get-WmiObject Win32_Processor | Select-Object -First 1; $InfoCPU = $ProcInfo.Name.Trim()
    $MemInfo = Get-WmiObject Win32_ComputerSystem; $RawRam = $MemInfo.TotalPhysicalMemory; $InfoRAM = "$([math]::Round($RawRam / 1GB, 1)) GB"
    $GpuInfo = Get-WmiObject Win32_VideoController | Select-Object -First 1; $InfoGPU = $GpuInfo.Name
    $DiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" } | Select-Object -First 1
    $InfoDisk = "$($DiskInfo.Model) ($([math]::Round($DiskInfo.Size / 1GB, 0)) GB)"
} catch { $InfoCPU="N/A"; $InfoRAM="N/A"; $InfoGPU="N/A"; $InfoDisk="N/A" }

# 3. WINSAT
Write-Host "[1/5] Tests de performance (WinSAT)..." -ForegroundColor Green
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
Write-Host "[2/5] Vérification Santé Disque (S.M.A.R.T)..." -ForegroundColor Green
try {
    $PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1
    $SmartStatus = $PhysicalDisk.HealthStatus
    $MediaType = $PhysicalDisk.MediaType
    if ($SmartStatus -eq "Healthy") { $NoteSmart = 20; $SmartMsg = "Disque Sain ($MediaType)" } 
    else { $NoteSmart = 0; $SmartMsg = "DANGER : État $SmartStatus" }
} catch { $NoteSmart = "N/A"; $SmartMsg = "Inconnu" }

# 5. RÉSEAU
Write-Host "[3/5] Test Connectivité..." -ForegroundColor Green
try {
    if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $NetStatus = "Connecté Internet" } 
    else { $NetStatus = "Pas d'Internet" }
} catch { $NetStatus = "Erreur Test" }

# 6. BATTERIE
Write-Host "[4/5] Analyse Batterie (Santé & Stress)..." -ForegroundColor Green
if (Test-Path $BatReportPath) { Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue }
Start-Process powercfg -ArgumentList "/batteryreport /output `"$BatReportPath`" /XML" -Wait -WindowStyle Hidden
Start-Sleep -Seconds 2

if (Test-Path $BatReportPath) {
    try {
        [xml]$BatXML = Get-Content $BatReportPath
        $ActiveBat = $BatXML.BatteryReport.Batteries.Battery | Select-Object -First 1
        if ($ActiveBat -and $ActiveBat.DesignCapacity -gt 0) {
            $DesignCap = [double]$ActiveBat.DesignCapacity
            $FullCap = [double]$ActiveBat.FullChargeCapacity
            $SantePercent = ($FullCap / $DesignCap) * 100
            $NoteSanteChimique = ($SantePercent / 5)
            $InfoBat = "$([math]::Round($FullCap,0)) mWh / $([math]::Round($DesignCap,0)) mWh"
        } else { $SantePercent=0; $NoteSanteChimique=0; $InfoBat="Données illisibles" }
    } catch { $SantePercent=0; $NoteSanteChimique=0; $InfoBat="Erreur lecture" }
    Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue
} else { $SantePercent="N/A"; $NoteSanteChimique="N/A"; $InfoBat="Non concerné (Fixe)" }

$BatteryMessage = "Test ignoré"; $DropPercent = 0; $NoteStress = 20
if ($SantePercent -ne "N/A" -and $SantePercent -gt 0) {
    $PowerStatus = Get-WmiObject -Class Win32_Battery
    if ($PowerStatus.BatteryStatus -ne 2) {
        Write-Host "      Stress Test Batterie (30s)..." -ForegroundColor Yellow
        $StartP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining
        $Job = Start-Job -ScriptBlock { $r=0; 1..10000000 | ForEach-Object { $r += [math]::Sqrt($_) } }
        for($i=0; $i -lt 30; $i++) { Write-Host "." -NoNewline; Start-Sleep 1 }; Write-Host ""
        Stop-Job $Job; Remove-Job $Job
        $EndP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining
        $DropPercent = $StartP - $EndP
        $BatteryMessage = "Perte: $DropPercent% en 30s ($StartP% -> $EndP%)"
        if ($DropPercent -eq 0) { $NoteStress = 20 } elseif ($DropPercent -eq 1) { $NoteStress = 15 } elseif ($DropPercent -eq 2) { $NoteStress = 10 } else { $NoteStress = 5 }
    }
}
if ($NoteSanteChimique -ne "N/A") { $NoteBatFinale = [math]::Round(($NoteSanteChimique + $NoteStress) / 2, 1) } else { $NoteBatFinale = "N/A" }

# 7. NOTE GLOBALE
$MoyennePerf = ($NoteCPU + $NoteRAM + $NoteGPU + $NoteDisk) / 4
if ($NoteBatFinale -eq "N/A") {
    $NoteGlobale = [math]::Round($MoyennePerf, 1); $FlagProbleme = $false
} else {
    $NoteGlobale = [math]::Round(($MoyennePerf * 0.6) + ($NoteBatFinale * 0.4), 1)
    $FlagProbleme = $false; $RaisonProbleme = ""
    if ($SantePercent -lt 50) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE HS (Santé Faible)" }
    if ($DropPercent -ge 3) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE INSTABLE (Chute Tension)" }
    if ($NoteSmart -eq 0) { $FlagProbleme = $true; $RaisonProbleme = "DISQUE DUR DÉFAILLANT" }

    if ($FlagProbleme) { if ($NoteGlobale -gt 10) { $NoteGlobale = 10 }; $Verdict = "À VÉRIFIER ($RaisonProbleme)" } 
    else {
        if ($NoteGlobale -ge 17) { $Verdict = "EXCELLENT" } elseif ($NoteGlobale -ge 15) { $Verdict = "TRÈS BON" } elseif ($NoteGlobale -ge 12) { $Verdict = "CORRECT" } else { $Verdict = "FAIBLE" }
    }
}

# 8. RAPPORT & EXPORT
Write-Host "[5/5] Génération des rapports..." -ForegroundColor Green
$SanteDisplay = if($SantePercent -ne "N/A"){"$([math]::Round($SantePercent,0))%"}else{"N/A"}

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

 5. BATTERIE
------------------------------------------------------------
 Note Finale : $NoteBatFinale / 20
 Santé (SOH) : $SanteDisplay (Note Chimique: $([math]::Round($NoteSanteChimique,1))/20)
 Tenue Stress: $NoteStress / 20 ($BatteryMessage)
 Détail      : $InfoBat

============================================================
"@

# Ecriture Rapport TXT (Historique inversé)
try {
    if (Test-Path $RapportPath) { $OldContent = Get-Content $RapportPath -Raw; $FinalContent = $NewContent + "`r`n`r`n`r`n" + $OldContent } 
    else { $FinalContent = $NewContent }
    $FinalContent | Out-File -FilePath $RapportPath -Encoding UTF8 -Force -ErrorAction Stop
    Write-Host "SUCCÈS RAPPORT : $RapportPath" -ForegroundColor Green
    Invoke-Item $RapportPath
} catch { Write-Host "Erreur écriture Rapport." -ForegroundColor Red }

# Ecriture CSV (Historique Append)
if (-not (Test-Path $CsvPath)) { 
    try { "Date;MachineID;Modele;Note_Globale;Verdict;CPU;RAM;GPU;Disk;Bat_Sante;Bat_Stress" | Out-File $CsvPath -Encoding UTF8 -ErrorAction Stop } catch {}
}

try {
    "$Date;$FormattedName;$ModelePC;$NoteGlobale;$Verdict;$NoteCPU;$NoteRAM;$NoteGPU;$NoteDisk;$SanteDisplay;$DropPercent" | Out-File $CsvPath -Append -Encoding UTF8 -ErrorAction Stop
    Write-Host "SUCCÈS CSV : Ajouté à $CsvPath" -ForegroundColor Green
} catch {
    Write-Host "ERREUR CSV : Fermez le fichier Excel !" -ForegroundColor Red
}

Write-Host ""
Read-Host "Terminé. Appuyez sur Entrée..."
