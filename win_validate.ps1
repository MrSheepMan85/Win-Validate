<#
.SYNOPSIS
    Outil de Validation Matérielle - v5.4
    Organisation : _RESULTATS > [Fabricant + Modèle] > Rapports
.DESCRIPTION
    - Dossiers nommés "Fabricant Modèle" (ex: Dell Latitude 5530).
    - Tout rangé dans _RESULTATS.
    - Features : Stress Test Sin/Cos, CSV Safe, WinSAT Retry.
#>

Add-Type -AssemblyName System.Windows.Forms

# =========================================================
#  1. INITIALISATION & CONFIGURATION
# =========================================================

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

# Configuration
$Config = @{
    Seuils = @{ CPU_Ref=9.2; RAM_Ref=9.2; GPU_Ref=6.2; Disk_Ref=8.6 }
    StressTest = @{ Duree_Secondes=120 }
    Criteres = @{ Batterie_Min_Sante=50; Batterie_Max_Perte=7 }
}

# Chargement JSON
$ConfigFile = Join-Path $RootDir "config.json"
if (Test-Path $ConfigFile) {
    try {
        $JsonContent = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($JsonContent.Seuils) { $Config.Seuils = $JsonContent.Seuils }
        if ($JsonContent.StressTest) { $Config.StressTest = $JsonContent.StressTest }
        if ($JsonContent.Criteres) { $Config.Criteres = $JsonContent.Criteres }
    } catch {}
}

# Fonction CSV Sécurisée
function Write-SafeCsv {
    param([string]$Path, [string]$Content)
    $MaxRetries = 5; $RetryCount = 0; $Success = $false
    while (-not $Success -and $RetryCount -lt $MaxRetries) {
        try {
            $Content | Out-File -FilePath $Path -Append -Encoding UTF8 -ErrorAction Stop
            $Success = $true
        } catch {
            $RetryCount++; $SleepTime = (Get-Random -Minimum 200 -Maximum 1000)
            Write-Warning "Fichier CSV verrouillé. Nouvelle tentative..."
            Start-Sleep -Milliseconds $SleepTime
        }
    }
}

$Date = Get-Date -Format "yyyy-MM-dd HH:mm"

# Check Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERREUR : Lancez en tant qu'Administrateur." -ForegroundColor Red; Read-Host "Entrée..."; Exit
}

# =========================================================
#  2. IDENTIFICATION & ARBORESCENCE (MODIFIÉ)
# =========================================================
try { $BiosInfo = Get-WmiObject Win32_BIOS -ErrorAction Stop; $ServiceTag = $BiosInfo.SerialNumber; $FormattedName = "FRALW-$ServiceTag" } catch { $FormattedName = "FRALW-Inconnu" }
try { $SysInfo = Get-WmiObject Win32_ComputerSystem; $ModelePC = $SysInfo.Model.Trim(); $Fabricant = $SysInfo.Manufacturer.Trim() } catch { $ModelePC = "Inconnu"; $Fabricant = "" }

# --- MODIFICATION v5.4 : Nom Dossier = Fabricant + Modèle ---
$RawFolderName = "$Fabricant $ModelePC"
$SafeModelName = $RawFolderName -replace '[\\/:\*\?"<>\|]', ''
# ------------------------------------------------------------

# Dossier Maître
$MasterFolder = Join-Path $RootDir "_RESULTATS"
if (-not (Test-Path $MasterFolder)) { try { New-Item -ItemType Directory -Force -Path $MasterFolder | Out-Null } catch {} }

# Dossier Modèle
$ModelDir = Join-Path $MasterFolder $SafeModelName
if (-not (Test-Path $ModelDir)) { try { New-Item -ItemType Directory -Force -Path $ModelDir | Out-Null } catch {} }

# Chemins finaux
$RapportPath = Join-Path $ModelDir "$FormattedName.txt"
$CsvPath = Join-Path $MasterFolder "Inventaire_Global.csv"
$BatReportPath = Join-Path $ModelDir "battery_temp.xml"

Clear-Host
Write-Host "============================================================" -ForegroundColor Red
Write-Host "   DIAGNOSTIC MATÉRIEL v5.4 (Fabricant + Modèle)            " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Red
Write-Host "Machine : $FormattedName" -ForegroundColor Green
Write-Host "Dossier : _RESULTATS\$SafeModelName" -ForegroundColor Yellow
Write-Host "Démarrage..."

# =========================================================
#  3. INFO MATÉRIEL
# =========================================================
try {
    $ProcInfo = Get-WmiObject Win32_Processor | Select-Object -First 1; $InfoCPU = $ProcInfo.Name.Trim(); $Cores = $ProcInfo.NumberOfLogicalProcessors
    $MemInfo = Get-WmiObject Win32_ComputerSystem; $RawRam = $MemInfo.TotalPhysicalMemory; $InfoRAM = "$([math]::Round($RawRam / 1GB, 1)) GB"
    $GpuInfo = Get-WmiObject Win32_VideoController | Select-Object -First 1; $InfoGPU = $GpuInfo.Name
    $DiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" } | Select-Object -First 1
    $InfoDisk = "$($DiskInfo.Model) ($([math]::Round($DiskInfo.Size / 1GB, 0)) GB)"
} catch { $InfoCPU="N/A"; $Cores=2; $InfoRAM="N/A"; $InfoGPU="N/A"; $InfoDisk="N/A" }

# =========================================================
#  4. WINSAT (RETRY AUTO)
# =========================================================
Write-Host "[1/5] Benchmark WinSAT..." -ForegroundColor Green
$WinSatAttempts = 0; $WinSatSuccess = $false; $WinSATPath = "$env:WINDIR\Performance\WinSAT\DataStore"
do {
    $WinSatAttempts++
    Start-Process winsat -ArgumentList "formal -restart" -Wait -WindowStyle Hidden
    $LatestXML = Get-ChildItem -Path "$WinSATPath\*Formal.Assessment*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($LatestXML) {
        [xml]$Data = Get-Content $LatestXML.FullName
        try {
            $ScoreCPU = [double]$Data.WinSAT.WinSPR.CpuScore
            if ($ScoreCPU -gt 0) { 
                $WinSatSuccess = $true; $ScoreRAM = [double]$Data.WinSAT.WinSPR.MemoryScore; $ScoreGPU = [double]$Data.WinSAT.WinSPR.GraphicsScore
                $ScoreD3D = [double]$Data.WinSAT.WinSPR.GamingScore; $ScoreDisk = [double]$Data.WinSAT.WinSPR.DiskScore
            }
        } catch {}
    }
} while (-not $WinSatSuccess -and $WinSatAttempts -lt 2)
if (-not $WinSatSuccess) { $ScoreCPU=0; $ScoreRAM=0; $ScoreGPU=0; $ScoreD3D=0; $ScoreDisk=0 }

function Calculate-RelativeScore ($current, $target) {
    if (!$current -or $current -eq 0) { return 0 }
    $val = ($current / $target) * 20; if ($val -gt 20) { return 20 }; return [math]::Round($val, 1)
}
$NoteCPU = Calculate-RelativeScore $ScoreCPU $Config.Seuils.CPU_Ref
$NoteRAM = Calculate-RelativeScore $ScoreRAM $Config.Seuils.RAM_Ref
$NoteGPU = Calculate-RelativeScore $ScoreD3D $Config.Seuils.GPU_Ref
$NoteDisk = Calculate-RelativeScore $ScoreDisk $Config.Seuils.Disk_Ref

# =========================================================
#  5. SMART & RÉSEAU
# =========================================================
Write-Host "[2/5] Santé Disque..." -ForegroundColor Green
try {
    $PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1; $SmartStatus = $PhysicalDisk.HealthStatus
    if ($SmartStatus -eq "Healthy") { $NoteSmart = 20; $SmartMsg = "Disque Sain" } else { $NoteSmart = 0; $SmartMsg = "DANGER : $SmartStatus" }
} catch { $NoteSmart = "N/A"; $SmartMsg = "Inconnu" }

Write-Host "[3/5] Réseau..." -ForegroundColor Green
if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $NetStatus = "Connecté Internet" } else { $NetStatus = "Pas d'Internet" }

# =========================================================
#  6. STRESS TEST (SIN/COS)
# =========================================================
$DureeStress = $Config.StressTest.Duree_Secondes
Write-Host "[4/5] FULL STRESS TEST ($DureeStress sec)..." -ForegroundColor Green

if (Test-Path $BatReportPath) { Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue }
Start-Process powercfg -ArgumentList "/batteryreport /output `"$BatReportPath`" /XML" -Wait -WindowStyle Hidden; Start-Sleep 1
$SantePercent="N/A"; $NoteSanteChimique="N/A"; $InfoBat="Non détectée"
if (Test-Path $BatReportPath) {
    try {
        [xml]$BatXML = Get-Content $BatReportPath; $ActiveBat = $BatXML.BatteryReport.Batteries.Battery | Select-Object -First 1
        if ($ActiveBat -and $ActiveBat.DesignCapacity -gt 0) {
            $DesignCap = [double]$ActiveBat.DesignCapacity; $FullCap = [double]$ActiveBat.FullChargeCapacity
            $SantePercent = ($FullCap / $DesignCap) * 100; $NoteSanteChimique = ($SantePercent / 5)
            $InfoBat = "$([math]::Round($FullCap,0)) mWh / $([math]::Round($DesignCap,0)) mWh"
        }
    } catch {}
    Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue
}

$IsPluggedIn = ([System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus -eq "Online")
$StartP = 0; try { $StartP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch {}
if ($IsPluggedIn) { Write-Host "      > SECTEUR : Test stabilité uniquement." -ForegroundColor Cyan } else { Write-Host "      > BATTERIE ($StartP%) : Test complet." -ForegroundColor Yellow }

$Jobs = @(); if (!$Cores) { $Cores = 2 }
1..$Cores | ForEach-Object { $Jobs += Start-Job -ScriptBlock { $x=0.1; while($true) { $x = [math]::Sin($x) * [math]::Cos($x) + [math]::Tan($x) } } }
$Jobs += Start-Job -ScriptBlock { $f="$env:TEMP\io.dat"; $d=[byte[]]::new(50*1MB); (new-object Random).NextBytes($d); while($true){[IO.File]::WriteAllBytes($f,$d);$null=[IO.File]::ReadAllBytes($f)} }
$Jobs += Start-Job -ScriptBlock { while($true) { Start-Process winsat -ArgumentList "d3d -objs C(20) -duration 5" -Wait -WindowStyle Hidden } }

for($i=0; $i -lt $DureeStress; $i++) { 
    $PercentDone = [math]::Round(($i / $DureeStress) * 100)
    Write-Progress -Activity "STRESS TEST" -Status "Progression: $PercentDone%" -PercentComplete $PercentDone; Start-Sleep 1 
}
Write-Progress -Activity "STRESS TEST" -Completed
Get-Job | Stop-Job | Remove-Job; try { Remove-Item "$env:TEMP\io.dat" -ErrorAction SilentlyContinue } catch {}

$DropPercent = 0; $NoteStress = "N/A"; $BatteryMessage = "Ignoré (Secteur)"
if (-not $IsPluggedIn) {
    $EndP = 0; try { $EndP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch { $EndP = $StartP }
    $DropPercent = $StartP - $EndP; $BatteryMessage = "Perte: $DropPercent%"
    if ($DropPercent -le 2) { $NoteStress = 20 } elseif ($DropPercent -le 4) { $NoteStress = 15 } elseif ($DropPercent -le 6) { $NoteStress = 10 } elseif ($DropPercent -le 8) { $NoteStress = 5 } else { $NoteStress = 0 }
}
if ($NoteSanteChimique -ne "N/A" -and $NoteStress -ne "N/A") { $NoteBatFinale = [math]::Round(($NoteSanteChimique + $NoteStress) / 2, 1) } elseif ($NoteSanteChimique -ne "N/A") { $NoteBatFinale = $NoteSanteChimique } else { $NoteBatFinale = "N/A" }

# =========================================================
#  7. VERDICT & RAPPORT
# =========================================================
$MoyennePerf = ($NoteCPU + $NoteRAM + $NoteGPU + $NoteDisk) / 4
if ($NoteBatFinale -eq "N/A") { $NoteGlobale = [math]::Round($MoyennePerf, 1); $FlagProbleme = $false } else {
    $NoteGlobale = [math]::Round(($MoyennePerf * 0.5) + ($NoteBatFinale * 0.5), 1); $FlagProbleme = $false; $RaisonProbleme = ""
    if ($SantePercent -ne "N/A" -and $SantePercent -lt $Config.Criteres.Batterie_Min_Sante) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE HS" }
    if ($NoteStress -ne "N/A" -and $DropPercent -ge $Config.Criteres.Batterie_Max_Perte) { $FlagProbleme = $true; $RaisonProbleme = "BATTERIE INSTABLE" }
    if ($NoteSmart -eq 0) { $FlagProbleme = $true; $RaisonProbleme = "DISQUE HS" }
    if ($FlagProbleme) { if ($NoteGlobale -gt 9) { $NoteGlobale = 9 }; $Verdict = "RECALÉ ($RaisonProbleme)" } else {
        if ($NoteGlobale -ge 16) { $Verdict = "PREMIUM" } elseif ($NoteGlobale -ge 14) { $Verdict = "BON ÉTAT" } elseif ($NoteGlobale -ge 11) { $Verdict = "STANDARD" } else { $Verdict = "FAIBLE" }
    }
}

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
 1. PROCESSEUR (CPU) : $NoteCPU / 20 (Score: $ScoreCPU) ($InfoCPU)
 2. MÉMOIRE (RAM)    : $NoteRAM / 20 (Score: $ScoreRAM) ($InfoRAM)
 3. GRAPHIQUE (GPU)  : $NoteGPU / 20 (Score: $ScoreD3D) ($InfoGPU)
 4. DISQUE           : $NoteDisk / 20 ($InfoDisk) - État: $SmartMsg
 5. BATTERIE         : $NoteBatFinale / 20 (Santé: $SanteDisplay / Stress: $StressDisplay)
============================================================
"@

try { $FinalContent = if (Test-Path $RapportPath) { $NewContent + "`r`n`r`n`r`n" + (Get-Content $RapportPath -Raw) } else { $NewContent }; $FinalContent | Out-File -FilePath $RapportPath -Encoding UTF8 -Force -ErrorAction Stop; Invoke-Item $RapportPath } catch { Write-Error "Erreur Rapport TXT" }

# CSV Global stocké dans _RESULTATS
if (-not (Test-Path $CsvPath)) { Write-SafeCsv -Path $CsvPath -Content "Date;MachineID;Modele;Note_Globale;Verdict;CPU_Score;Bat_Sante;Bat_Drop" }
$CsvLine = "$Date;$FormattedName;$ModelePC;$NoteGlobale;$Verdict;$ScoreCPU;$SanteDisplay;$DropPercent"
Write-SafeCsv -Path $CsvPath -Content $CsvLine

Write-Host ""
Read-Host "Terminé. Entrée..."
