<#
.SYNOPSIS
    Outil de Validation Matérielle - v6.9 (Admin Console)
    Features : Dashboard Dédoublonné, Option Quitter après Dashboard, Fix Arrays.
#>

Add-Type -AssemblyName System.Windows.Forms

# =========================================================
#  0. FONCTIONS SOCLES
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

# Config par défaut
$Config = @{
    Seuils = @{ CPU_Ref=9.2; RAM_Ref=9.2; GPU_Ref=6.2; Disk_Ref=8.6 }
    StressTest = @{ Duree_Secondes=120 }
    Criteres = @{ Batterie_Min_Sante=50; Batterie_Max_Perte=7 }
    Fichiers = @{ Nom_Csv="Inventaire_Parc_Global.csv"; Nom_Log="win_validate.log" }
}

# Dossiers
$MasterFolder = Join-Path $RootDir "_RESULTATS"
if (-not (Test-Path $MasterFolder)) { New-Item -ItemType Directory -Force -Path $MasterFolder | Out-Null }
$LogPath = Join-Path $MasterFolder $Config.Fichiers.Nom_Log
$CsvPath = Join-Path $MasterFolder $Config.Fichiers.Nom_Csv

# Logging
function Write-Log {
    param([string]$Message, [string]$Level="INFO")
    $Timestamp = Get-Date -Format "HH:mm:ss"
    $Color = switch($Level) { "INFO"{"DarkCyan"} "SUCCESS"{"Green"} "WARN"{"Yellow"} "ERROR"{"Red"} Default{"Gray"} }
    Write-Host "   [$Timestamp] [$Level] $Message" -ForegroundColor $Color
    try { "[$Timestamp] [$Level] $Message" | Out-File -FilePath $LogPath -Append -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# Chargement JSON
$ConfigFile = Join-Path $RootDir "config.json"
if (Test-Path $ConfigFile) {
    try {
        $JsonContent = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($JsonContent.Seuils) { $Config.Seuils = $JsonContent.Seuils }
        if ($JsonContent.StressTest) { $Config.StressTest = $JsonContent.StressTest }
        if ($JsonContent.Criteres) { $Config.Criteres = $JsonContent.Criteres }
        if ($JsonContent.Fichiers) { $Config.Fichiers = $JsonContent.Fichiers }
        $CsvPath = Join-Path $MasterFolder $Config.Fichiers.Nom_Csv 
    } catch { Write-Log "Erreur JSON. Défauts utilisés." -Level WARN }
}

# =========================================================
#  1. DASHBOARD WEB (CORRECTIF TABLEAU VIDE)
# =========================================================
function Show-Dashboard {
    if (-not (Test-Path $CsvPath)) {
        Write-Host "   [!] Aucun fichier CSV d'inventaire trouvé." -ForegroundColor Red
        return
    }

    Write-Host "   Traitement des données (Dédoublonnage)..." -ForegroundColor Cyan
    
    # FORCAGE EN TABLEAU @(...) POUR EVITER LE BUG SI 1 SEULE LIGNE
    $RawData = @(Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8)
    
    # Dédoublonnage : On ne garde que le plus récent par MachineID
    $UniqueData = @($RawData | Sort-Object Date -Descending | Group-Object MachineID | ForEach-Object { $_.Group[0] })
    
    $Total = $UniqueData.Count
    
    # Calculs statistiques (forcés en tableau aussi)
    $Notes = $UniqueData | ForEach-Object { $val = $_."Note_Globale" -replace ',', '.'; [double]$val }
    if ($Total -gt 0) {
        $Avg = ($Notes | Measure-Object -Average).Average
        $AvgDisplay = [math]::Round($Avg, 2)
    } else { $AvgDisplay = 0 }

    $PremiumCount = @($UniqueData | Where-Object { $_.Verdict -match "PREMIUM" }).Count
    $RecaleCount = @($UniqueData | Where-Object { $_.Verdict -match "RECALÉ" }).Count
    
    # Génération HTML
    $HtmlHeader = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Dashboard Parc (Vue Unique)</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; margin: 0; padding: 20px; color: #333; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .kpi-container { display: flex; gap: 20px; margin-bottom: 30px; }
        .card { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); flex: 1; text-align: center; border-top: 4px solid #3498db; }
        .card h3 { margin: 0 0 10px 0; color: #7f8c8d; font-size: 14px; text-transform: uppercase; }
        .card .value { font-size: 36px; font-weight: bold; color: #2c3e50; }
        .card.premium { border-top-color: #00b894; } .card.premium .value { color: #00b894; }
        .card.fail { border-top-color: #d63031; } .card.fail .value { color: #d63031; }
        .info-unique { background: #e1f5fe; color: #0277bd; padding: 15px; border-radius: 8px; margin-bottom: 20px; border-left: 5px solid #0277bd; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        th { background: #34495e; padding: 15px; text-align: left; font-weight: 600; color: #ecf0f1; }
        td { padding: 12px 15px; border-bottom: 1px solid #eee; }
        .badge { padding: 5px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; color: white; display: inline-block; }
        .bg-premium { background-color: #00b894; } .bg-bon { background-color: #0984e3; } .bg-std { background-color: #f1c40f; color: #333; } .bg-fail { background-color: #d63031; }
        .row-fail { background-color: #fff5f5; }
    </style>
</head>
<body>
    <div class="header">
        <div><h1>📊 Dashboard de Validation</h1><small>Gestion de Parc Automatisée</small></div>
        <div><strong>$(Get-Date -Format 'dd/MM/yyyy HH:mm')</strong></div>
    </div>

    <div class="info-unique">
        <strong>ℹ️ Mode Unique :</strong> Affichage du dernier scan connu pour chaque numéro de série (Dédoublonné).
    </div>

    <div class="kpi-container">
        <div class="card"><h3>Machines Uniques</h3><div class="value">$Total</div></div>
        <div class="card"><h3>Note Moyenne</h3><div class="value">$AvgDisplay <small style="font-size:16px;color:#999">/20</small></div></div>
        <div class="card premium"><h3>Premium</h3><div class="value">$PremiumCount</div></div>
        <div class="card fail"><h3>Recalés</h3><div class="value">$RecaleCount</div></div>
    </div>

    <table>
        <thead>
            <tr>
                <th>Dernier Scan</th>
                <th>Machine ID</th>
                <th>Modèle</th>
                <th>Note</th>
                <th>Verdict</th>
                <th>CPU Score</th>
                <th>Batterie</th>
                <th>Drop Stress</th>
            </tr>
        </thead>
        <tbody>
"@
    
    $HtmlRows = ""
    foreach ($row in $UniqueData) {
        $BadgeClass = "bg-std"
        $RowClass = ""
        if ($row.Verdict -match "PREMIUM") { $BadgeClass = "bg-premium" }
        elseif ($row.Verdict -match "BON") { $BadgeClass = "bg-bon" }
        elseif ($row.Verdict -match "RECALÉ") { $BadgeClass = "bg-fail"; $RowClass = "row-fail" }
        
        $HtmlRows += @"
            <tr class="$RowClass">
                <td>$($row.Date)</td>
                <td><strong>$($row.MachineID)</strong></td>
                <td>$($row.Modele)</td>
                <td><strong>$($row.Note_Globale)</strong></td>
                <td><span class="badge $BadgeClass">$($row.Verdict)</span></td>
                <td>$($row.CPU_Score)</td>
                <td>$($row.Bat_Sante)</td>
                <td>$($row.Bat_Drop)%</td>
            </tr>
"@
    }

    $HtmlFooter = @"
        </tbody>
    </table>
</body>
</html>
"@
    $FinalHtml = $HtmlHeader + $HtmlRows + $HtmlFooter
    $TempPath = "$env:TEMP\Dashboard_Parc.html"
    $FinalHtml | Out-File -FilePath $TempPath -Encoding UTF8 -Force
    Start-Process $TempPath
}

# =========================================================
#  2. MENU DÉMARRAGE (OPTION QUITTER)
# =========================================================
Clear-Host
Write-Host "============================================================" -ForegroundColor Red
Write-Host "   WIN-VALIDATE v6.9 (Admin Console)                        " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Red

Write-Host ""
# 1. QUESTION DASHBOARD
$ShowDb = Read-Host " [?] Voulez-vous voir le Dashboard du Parc ? (O/N) : "
if ($ShowDb -match "O|o") {
    Show-Dashboard
    Write-Host "   > Dashboard ouvert dans le navigateur." -ForegroundColor Green
    Write-Host ""
    
    # 2. QUESTION QUITTER (NOUVEAU)
    $Quit = Read-Host " [?] Voulez-vous QUITTER l'application maintenant ? (O = Quitter / Entrée = Continuer) : "
    if ($Quit -match "O|o") {
        Write-Host "Fermeture..."
        Start-Sleep 1
        Exit
    }
}

# Vérification Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERREUR : Lancez en tant qu'Administrateur." -ForegroundColor Red; Read-Host "Entrée..."; Exit
}

Write-Host ""
Write-Host "Appuyez sur Entrée pour lancer le DIAGNOSTIC de CETTE machine..." -ForegroundColor Yellow
Read-Host " > " 

# =========================================================
#  3. FONCTION CSV SÉCURISÉE
# =========================================================
function Write-SafeCsv {
    param([string]$Path, [string]$Content)
    $MutexName = "Global\WinValidateCSV"
    $Mutex = New-Object System.Threading.Mutex($false, $MutexName)
    Write-Log "Ecriture CSV..."
    try {
        if ($Mutex.WaitOne(5000)) {
            $MaxRetries = 5; $RetryCount = 0; $Success = $false
            while (-not $Success -and $RetryCount -lt $MaxRetries) {
                try {
                    $Content | Out-File -FilePath $Path -Append -Encoding UTF8 -ErrorAction Stop
                    $Success = $true
                    Write-Log "Ligne ajoutée au CSV." -Level SUCCESS
                } catch {
                    $RetryCount++
                    Write-Log "CSV Verrouillé ($RetryCount)..." -Level WARN
                    Start-Sleep -Milliseconds ($RetryCount * 500)
                }
            }
            if (-not $Success) { Write-Log "ECHEC FATAL CSV." -Level ERROR }
            $Mutex.ReleaseMutex()
        } else { Write-Log "Timeout Mutex CSV." -Level ERROR }
    } catch { Write-Log "Erreur système Mutex." -Level ERROR } finally { $Mutex.Dispose() }
}

# =========================================================
#  4. IDENTIFICATION
# =========================================================
try { $BiosInfo = Get-WmiObject Win32_BIOS -ErrorAction Stop; $ServiceTag = $BiosInfo.SerialNumber.Trim(); $FormattedName = "FRALW-$ServiceTag" } catch { $FormattedName = "FRALW-Inconnu" }
try { $SysInfo = Get-WmiObject Win32_ComputerSystem; $ModelePC = $SysInfo.Model.Trim(); $Fabricant = $SysInfo.Manufacturer.Trim() } catch { $ModelePC = "Inconnu"; $Fabricant = "" }

$SafeFolderName = "$Fabricant $ModelePC" -replace '[\\/:\*\?"<>\|]', ''
$ModelDir = Join-Path $MasterFolder $SafeFolderName
if (-not (Test-Path $ModelDir)) { New-Item -ItemType Directory -Force -Path $ModelDir | Out-Null; Write-Log "Dossier créé : $SafeFolderName" -Level SUCCESS }

$RapportPath = Join-Path $ModelDir "$FormattedName.txt"
$BatReportPath = Join-Path $ModelDir "battery_temp.xml"

Write-Host "Machine : $FormattedName" -ForegroundColor Green
Write-Host "Dossier : _RESULTATS\$SafeFolderName" -ForegroundColor Yellow
Write-Host "Démarrage..."

# =========================================================
#  5. DIAGNOSTIC
# =========================================================

# Matériel
Write-Log "Scan matériel..."
try {
    $ProcInfo = Get-WmiObject Win32_Processor | Select-Object -First 1; $InfoCPU = $ProcInfo.Name.Trim(); $Cores = $ProcInfo.NumberOfLogicalProcessors
    $MemInfo = Get-WmiObject Win32_ComputerSystem; $RawRam = $MemInfo.TotalPhysicalMemory; $InfoRAM = "$([math]::Round($RawRam / 1GB, 1)) GB"
    $GpuInfo = Get-WmiObject Win32_VideoController | Select-Object -First 1; $InfoGPU = $GpuInfo.Name.Trim()
    $DiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" } | Select-Object -First 1; $InfoDisk = "$($DiskInfo.Model.Trim()) ($([math]::Round($DiskInfo.Size / 1GB, 0)) GB)"
    Write-Log "> CPU: $InfoCPU | RAM: $InfoRAM"
} catch { Write-Log "Erreur WMI." -Level ERROR; $InfoCPU="N/A"; $Cores=2 }

# WinSAT
Write-Host "[1/5] Benchmark WinSAT..." -ForegroundColor Green
$WinSatAttempts = 0; $WinSatSuccess = $false; $WinSATPath = "$env:WINDIR\Performance\WinSAT\DataStore"
do {
    $WinSatAttempts++; Write-Log "WinSAT Tentative $WinSatAttempts..."
    Start-Process winsat -ArgumentList "formal -restart" -Wait -WindowStyle Hidden; Start-Sleep 2 
    $LatestXML = Get-ChildItem -Path "$WinSATPath\*Formal.Assessment*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($LatestXML) {
        [xml]$Data = Get-Content $LatestXML.FullName
        try {
            $ScoreCPU = [double]$Data.WinSAT.WinSPR.CpuScore
            if ($ScoreCPU -gt 0) { 
                $WinSatSuccess = $true; $ScoreRAM = [double]$Data.WinSAT.WinSPR.MemoryScore; $ScoreGPU = [double]$Data.WinSAT.WinSPR.GraphicsScore
                $ScoreD3D = [double]$Data.WinSAT.WinSPR.GamingScore; $ScoreDisk = [double]$Data.WinSAT.WinSPR.DiskScore
                Write-Log "Scores WinSAT OK." -Level SUCCESS
            }
        } catch {}
    }
} while (-not $WinSatSuccess -and $WinSatAttempts -lt 2)
if (-not $WinSatSuccess) { $ScoreCPU=0; $ScoreRAM=0; $ScoreGPU=0; $ScoreD3D=0; $ScoreDisk=0; Write-Log "ECHEC WinSAT." -Level ERROR }

function Calculate-RelativeScore ($current, $target) {
    if (!$current -or $current -eq 0) { return 0 }
    $val = ($current / $target) * 20; if ($val -gt 20) { return 20 }; return [math]::Round($val, 1)
}
$NoteCPU = Calculate-RelativeScore $ScoreCPU $Config.Seuils.CPU_Ref; $NoteRAM = Calculate-RelativeScore $ScoreRAM $Config.Seuils.RAM_Ref
$NoteGPU = Calculate-RelativeScore $ScoreD3D $Config.Seuils.GPU_Ref; $NoteDisk = Calculate-RelativeScore $ScoreDisk $Config.Seuils.Disk_Ref

# SMART & Réseau
Write-Host "[2/5] Santé Disque..." -ForegroundColor Green
try {
    $PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1; $SmartStatus = $PhysicalDisk.HealthStatus
    if ($SmartStatus -eq "Healthy") { $NoteSmart = 20; $SmartMsg = "Disque Sain" } else { $NoteSmart = 0; $SmartMsg = "DANGER : $SmartStatus" }
} catch { $NoteSmart = "N/A"; $SmartMsg = "Inconnu" }
Write-Host "[3/5] Réseau..." -ForegroundColor Green
if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $NetStatus = "Connecté Internet"; Write-Log "Ping OK." -Level SUCCESS } else { $NetStatus = "Pas d'Internet"; Write-Log "Ping ECHEC." -Level WARN }

# Stress Test
$DureeStress = $Config.StressTest.Duree_Secondes
Write-Host "[4/5] FULL STRESS TEST ($DureeStress sec)..." -ForegroundColor Green
if (Test-Path $BatReportPath) { Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue }
$StartP = 0; try { $StartP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch {}
$IsPluggedIn = ([System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus -eq "Online")
if ($IsPluggedIn) { Write-Log "Mode SECTEUR." } else { Write-Log "Mode BATTERIE ($StartP%)." }

$Jobs = @(); if (!$Cores) { $Cores = 2 }
try {
    Write-Log "Charges lancées..."
    1..$Cores | ForEach-Object { $Jobs += Start-Job -ScriptBlock { $x=0.1; while($true) { $x = [math]::Sin($x) * [math]::Cos($x) + [math]::Tan($x) + [math]::Sqrt($x) } } }
    $Jobs += Start-Job -ScriptBlock { $f="$env:TEMP\io.dat"; $d=[byte[]]::new(50*1MB); (new-object Random).NextBytes($d); while($true){[IO.File]::WriteAllBytes($f,$d);$null=[IO.File]::ReadAllBytes($f)} }
    $Jobs += Start-Job -ScriptBlock { while($true) { Start-Process winsat -ArgumentList "d3d -objs C(20) -duration 5" -Wait -WindowStyle Hidden } }
    for($i=0; $i -lt $DureeStress; $i++) { 
        $PercentDone = [math]::Round(($i / $DureeStress) * 100)
        Write-Progress -Activity "STRESS TEST" -Status "Progression: $PercentDone%" -PercentComplete $PercentDone; Start-Sleep 1 
    }
    Write-Progress -Activity "STRESS TEST" -Completed; Write-Log "Fin timer."
} finally {
    Write-Log "Nettoyage process..."
    Get-Job | Stop-Job | Remove-Job -Force; try { Remove-Item "$env:TEMP\io.dat" -ErrorAction SilentlyContinue } catch {}
}

$SantePercent="N/A"; $NoteSanteChimique="N/A"; $InfoBat="Non détectée"
Start-Process powercfg -ArgumentList "/batteryreport /output `"$BatReportPath`" /XML" -Wait -WindowStyle Hidden
if (Test-Path $BatReportPath) {
    try {
        [xml]$BatXML = Get-Content $BatReportPath; $ActiveBat = $BatXML.BatteryReport.Batteries.Battery | Select-Object -First 1
        if ($ActiveBat -and $ActiveBat.DesignCapacity -gt 0) {
            $DesignCap = [double]$ActiveBat.DesignCapacity; $FullCap = [double]$ActiveBat.FullChargeCapacity
            $SantePercent = ($FullCap / $DesignCap) * 100; $NoteSanteChimique = ($SantePercent / 5)
            $InfoBat = "$([math]::Round($FullCap,0)) mWh / $([math]::Round($DesignCap,0)) mWh"
            Write-Log "Santé Bat: $([math]::Round($SantePercent))%" -Level SUCCESS
        }
    } catch {}; Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue
}

$DropPercent = 0; $NoteStress = "N/A"
if (-not $IsPluggedIn) {
    $EndP = 0; try { $EndP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch { $EndP = $StartP }
    $DropPercent = $StartP - $EndP; Write-Log "Drop: $DropPercent%"
    if ($DropPercent -le 2) { $NoteStress = 20 } elseif ($DropPercent -le 4) { $NoteStress = 15 } elseif ($DropPercent -le 6) { $NoteStress = 10 } elseif ($DropPercent -le 8) { $NoteStress = 5 } else { $NoteStress = 0 }
}
if ($NoteSanteChimique -ne "N/A" -and $NoteStress -ne "N/A") { $NoteBatFinale = [math]::Round(($NoteSanteChimique + $NoteStress) / 2, 1) } elseif ($NoteSanteChimique -ne "N/A") { $NoteBatFinale = [math]::Round($NoteSanteChimique, 1) } else { $NoteBatFinale = "N/A" }

# Verdict
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
$ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm"

$NewContent = @"
============================================================
       RAPPORT DIAGNOSTIC - $ReportDate
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

try { 
    $FinalContent = if (Test-Path $RapportPath) { $NewContent + "`r`n`r`n`r`n" + (Get-Content $RapportPath -Raw) } else { $NewContent }
    $FinalContent | Out-File -FilePath $RapportPath -Encoding UTF8 -Force -ErrorAction Stop
    Invoke-Item $RapportPath
    Write-Log "TXT OK." -Level SUCCESS
} catch { Write-Log "Echec TXT." -Level ERROR }

if (-not (Test-Path $CsvPath)) { Write-SafeCsv -Path $CsvPath -Content "Date;MachineID;Modele;Note_Globale;Verdict;CPU_Score;Bat_Sante;Bat_Drop" }
$CsvLine = "$ReportDate;$FormattedName;$ModelePC;$NoteGlobale;$Verdict;$ScoreCPU;$SanteDisplay;$DropPercent"
Write-SafeCsv -Path $CsvPath -Content $CsvLine

Write-Host ""; Write-Log "Fin."
Read-Host "Terminé. Entrée : "
