<#
.SYNOPSIS
    Outil de Validation Matérielle - v6.4 (Version Complète Déployée)
    Toutes fonctionnalités incluses + Logs détaillés + Date dynamique.
#>

Add-Type -AssemblyName System.Windows.Forms

# =========================================================
#  0. FONCTIONS SOCLES (LOGGING & CONFIGURATION)
# =========================================================

function Get-SmartLocation {
    # Détermine l'emplacement du script ou de l'exe
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

# Configuration par défaut
$Config = @{
    Seuils = @{ 
        CPU_Ref=9.2
        RAM_Ref=9.2
        GPU_Ref=6.2
        Disk_Ref=8.6 
    }
    StressTest = @{ 
        Duree_Secondes=120 
    }
    Criteres = @{ 
        Batterie_Min_Sante=50
        Batterie_Max_Perte=7 
    }
    Fichiers = @{ 
        Nom_Csv="Inventaire_Parc_Global.csv"
        Nom_Log="win_validate.log" 
    }
}

# Initialisation Dossier Maître
$MasterFolder = Join-Path $RootDir "_RESULTATS"
if (-not (Test-Path $MasterFolder)) { 
    New-Item -ItemType Directory -Force -Path $MasterFolder | Out-Null 
}
$LogPath = Join-Path $MasterFolder $Config.Fichiers.Nom_Log

# Fonction de Logging Détaillée
function Write-Log {
    param([string]$Message, [string]$Level="INFO")
    $Timestamp = Get-Date -Format "HH:mm:ss"
    
    # Choix de la couleur selon le niveau
    $Color = switch($Level) { 
        "INFO"    {"DarkCyan"} 
        "SUCCESS" {"Green"}
        "WARN"    {"Yellow"} 
        "ERROR"   {"Red"} 
        Default   {"Gray"} 
    }
    
    # Affichage Console
    Write-Host "   [$Timestamp] [$Level] $Message" -ForegroundColor $Color
    
    # Enregistrement dans le fichier Log
    try { 
        "[$Timestamp] [$Level] $Message" | Out-File -FilePath $LogPath -Append -Encoding UTF8 -ErrorAction SilentlyContinue 
    } catch {}
}

# Chargement du fichier JSON externe
$ConfigFile = Join-Path $RootDir "config.json"
if (Test-Path $ConfigFile) {
    try {
        $JsonContent = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($JsonContent.Seuils) { $Config.Seuils = $JsonContent.Seuils }
        if ($JsonContent.StressTest) { $Config.StressTest = $JsonContent.StressTest }
        if ($JsonContent.Criteres) { $Config.Criteres = $JsonContent.Criteres }
        if ($JsonContent.Fichiers) { $Config.Fichiers = $JsonContent.Fichiers }
        Write-Log "Configuration chargée depuis : $(Split-Path $ConfigFile -Leaf)" -Level SUCCESS
    } catch { 
        Write-Log "Echec lecture JSON. Paramètres par défaut utilisés." -Level ERROR 
    }
} else { 
    Write-Log "Fichier config.json absent. Paramètres par défaut utilisés." -Level WARN 
}

# =========================================================
#  1. FONCTION CSV SÉCURISÉE (MUTEX)
# =========================================================
function Write-SafeCsv {
    param([string]$Path, [string]$Content)
    
    # Création du Mutex pour empêcher les conflits d'écriture
    $MutexName = "Global\WinValidateCSV"
    $Mutex = New-Object System.Threading.Mutex($false, $MutexName)
    
    Write-Log "Préparation écriture CSV..."
    
    try {
        # On attend jusqu'à 5 secondes pour avoir le droit d'écrire
        if ($Mutex.WaitOne(5000)) {
            $MaxRetries = 5
            $RetryCount = 0
            $Success = $false
            
            # Boucle de réessai en cas de fichier ouvert par Excel
            while (-not $Success -and $RetryCount -lt $MaxRetries) {
                try {
                    $Content | Out-File -FilePath $Path -Append -Encoding UTF8 -ErrorAction Stop
                    $Success = $true
                    Write-Log "Ligne ajoutée au CSV avec succès." -Level SUCCESS
                } catch {
                    $RetryCount++
                    Write-Log "CSV Verrouillé (Essai $RetryCount/$MaxRetries). Attente..." -Level WARN
                    Start-Sleep -Milliseconds ($RetryCount * 500)
                }
            }
            
            if (-not $Success) { 
                Write-Log "ECHEC FATAL CSV (Fichier bloqué par un autre programme)." -Level ERROR 
            }
            
            $Mutex.ReleaseMutex()
        } else { 
            Write-Log "Timeout Mutex CSV (Impossible d'obtenir l'accès)." -Level ERROR 
        }
    } catch { 
        Write-Log "Erreur système Mutex : $_" -Level ERROR 
    } finally { 
        $Mutex.Dispose() 
    }
}

# Vérification des droits Admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERREUR : Lancez en tant qu'Administrateur." -ForegroundColor Red
    Read-Host "Entrée..."
    Exit
}

# =========================================================
#  2. IDENTIFICATION ET DOSSIERS
# =========================================================
try { 
    $BiosInfo = Get-WmiObject Win32_BIOS -ErrorAction Stop
    $ServiceTag = $BiosInfo.SerialNumber.Trim()
    $FormattedName = "FRALW-$ServiceTag" 
} catch { 
    $FormattedName = "FRALW-Inconnu"
    Write-Log "Erreur lecture ServiceTag." -Level ERROR 
}

try { 
    $SysInfo = Get-WmiObject Win32_ComputerSystem
    $ModelePC = $SysInfo.Model.Trim()
    $Fabricant = $SysInfo.Manufacturer.Trim()
} catch { 
    $ModelePC = "Inconnu"
    $Fabricant = ""
    Write-Log "Erreur lecture Modèle." -Level ERROR 
}

# Création du dossier propre "Fabricant Modèle"
$SafeFolderName = "$Fabricant $ModelePC" -replace '[\\/:\*\?"<>\|]', ''
$ModelDir = Join-Path $MasterFolder $SafeFolderName

if (-not (Test-Path $ModelDir)) { 
    New-Item -ItemType Directory -Force -Path $ModelDir | Out-Null
    Write-Log "Dossier créé : $SafeFolderName" -Level SUCCESS
}

$RapportPath = Join-Path $ModelDir "$FormattedName.txt"
$CsvPath = Join-Path $MasterFolder $Config.Fichiers.Nom_Csv
$BatReportPath = Join-Path $ModelDir "battery_temp.xml"

Clear-Host
Write-Host "============================================================" -ForegroundColor Red
Write-Host "   DIAGNOSTIC MATÉRIEL v6.4 (Ultimate)                      " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Red
Write-Host "Machine : $FormattedName" -ForegroundColor Green
Write-Host "Dossier : _RESULTATS\$SafeFolderName" -ForegroundColor Yellow
Write-Host "Démarrage..."

# =========================================================
#  3. INFO MATÉRIEL
# =========================================================
Write-Log "Scan matériel en cours..."
try {
    $ProcInfo = Get-WmiObject Win32_Processor | Select-Object -First 1
    $InfoCPU = $ProcInfo.Name.Trim()
    $Cores = $ProcInfo.NumberOfLogicalProcessors
    
    $MemInfo = Get-WmiObject Win32_ComputerSystem
    $RawRam = $MemInfo.TotalPhysicalMemory
    $InfoRAM = "$([math]::Round($RawRam / 1GB, 1)) GB"
    
    $GpuInfo = Get-WmiObject Win32_VideoController | Select-Object -First 1
    $InfoGPU = $GpuInfo.Name.Trim()
    
    $DiskInfo = Get-WmiObject Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed" } | Select-Object -First 1
    $InfoDisk = "$($DiskInfo.Model.Trim()) ($([math]::Round($DiskInfo.Size / 1GB, 0)) GB)"
    
    Write-Log "   > CPU : $InfoCPU ($Cores coeurs)"
    Write-Log "   > RAM : $InfoRAM"
    Write-Log "   > GPU : $InfoGPU"
} catch { 
    Write-Log "Erreur WMI critique." -Level ERROR
    $InfoCPU="N/A"; $Cores=2; $InfoRAM="N/A"; $InfoGPU="N/A"; $InfoDisk="N/A" 
}

# =========================================================
#  4. WINSAT (RETRY AUTOMATIQUE)
# =========================================================
Write-Host "[1/5] Benchmark WinSAT..." -ForegroundColor Green
$WinSatAttempts = 0
$WinSatSuccess = $false
$WinSATPath = "$env:WINDIR\Performance\WinSAT\DataStore"

do {
    $WinSatAttempts++
    Write-Log "WinSAT Tentative $WinSatAttempts/2..."
    
    Start-Process winsat -ArgumentList "formal -restart" -Wait -WindowStyle Hidden
    Start-Sleep -Seconds 2 
    
    $LatestXML = Get-ChildItem -Path "$WinSATPath\*Formal.Assessment*.xml" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($LatestXML) {
        [xml]$Data = Get-Content $LatestXML.FullName
        try {
            $ScoreCPU = [double]$Data.WinSAT.WinSPR.CpuScore
            
            # Vérification que le score n'est pas 0 (bug WinSAT possible)
            if ($ScoreCPU -gt 0) { 
                $WinSatSuccess = $true
                $ScoreRAM = [double]$Data.WinSAT.WinSPR.MemoryScore
                $ScoreGPU = [double]$Data.WinSAT.WinSPR.GraphicsScore
                $ScoreD3D = [double]$Data.WinSAT.WinSPR.GamingScore
                $ScoreDisk = [double]$Data.WinSAT.WinSPR.DiskScore
                Write-Log "Scores WinSAT récupérés (CPU: $ScoreCPU)" -Level SUCCESS
            } else { 
                Write-Log "Score WinSAT 0.0 détecté (Invalide)." -Level WARN 
            }
        } catch { 
            Write-Log "XML WinSAT illisible." -Level WARN 
        }
    }
} while (-not $WinSatSuccess -and $WinSatAttempts -lt 2)

if (-not $WinSatSuccess) { 
    $ScoreCPU=0; $ScoreRAM=0; $ScoreGPU=0; $ScoreD3D=0; $ScoreDisk=0
    Write-Log "ECHEC WinSAT définitif." -Level ERROR 
}

function Calculate-RelativeScore ($current, $target) {
    if (!$current -or $current -eq 0) { return 0 }
    $val = ($current / $target) * 20
    if ($val -gt 20) { return 20 }
    return [math]::Round($val, 1)
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
    $PhysicalDisk = Get-PhysicalDisk | Select-Object -First 1
    $SmartStatus = $PhysicalDisk.HealthStatus
    if ($SmartStatus -eq "Healthy") { 
        $NoteSmart = 20; $SmartMsg = "Disque Sain" 
    } else { 
        $NoteSmart = 0; $SmartMsg = "DANGER : $SmartStatus" 
    }
    Write-Log "SMART : $SmartStatus"
} catch { 
    $NoteSmart = "N/A"; $SmartMsg = "Inconnu"
    Write-Log "Erreur lecture SMART." -Level WARN 
}

Write-Host "[3/5] Réseau..." -ForegroundColor Green
if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { 
    $NetStatus = "Connecté Internet"
    Write-Log "Ping Google OK." -Level SUCCESS
} else { 
    $NetStatus = "Pas d'Internet"
    Write-Log "Ping Google ECHEC." -Level WARN 
}

# =========================================================
#  6. STRESS TEST (LOGIQUE COMPLÈTE)
# =========================================================
$DureeStress = $Config.StressTest.Duree_Secondes
Write-Host "[4/5] FULL STRESS TEST ($DureeStress sec)..." -ForegroundColor Green

# Nettoyage préventif
if (Test-Path $BatReportPath) { Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue }

# Snapshot Batterie avant stress
$StartP = 0
try { $StartP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch {}
$IsPluggedIn = ([System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus -eq "Online")

if ($IsPluggedIn) { 
    Write-Log "Mode SECTEUR : Test stabilité (Pas de note batterie)." 
} else { 
    Write-Log "Mode BATTERIE ($StartP%) : Test autonomie actif." 
}

$Jobs = @()
if (!$Cores) { $Cores = 2 }

# Bloc Try/Finally pour garantir l'arrêt des processus (Anti-Zombie)
try {
    Write-Log "Lancement charges (CPU Sin/Cos + GPU + Disk IO)..."
    
    # Charge CPU Mathématique
    1..$Cores | ForEach-Object { 
        $Jobs += Start-Job -ScriptBlock { 
            $x=0.1
            while($true) { $x = [math]::Sin($x) * [math]::Cos($x) + [math]::Tan($x) + [math]::Sqrt($x) } 
        } 
    }
    # Charge Disque
    $Jobs += Start-Job -ScriptBlock { 
        $f="$env:TEMP\io.dat"; $d=[byte[]]::new(50*1MB); (new-object Random).NextBytes($d)
        while($true){ [IO.File]::WriteAllBytes($f,$d); $null=[IO.File]::ReadAllBytes($f) } 
    }
    # Charge GPU
    $Jobs += Start-Job -ScriptBlock { 
        while($true) { Start-Process winsat -ArgumentList "d3d -objs C(20) -duration 5" -Wait -WindowStyle Hidden } 
    }

    # Boucle d'attente avec barre de progression
    for($i=0; $i -lt $DureeStress; $i++) { 
        $PercentDone = [math]::Round(($i / $DureeStress) * 100)
        Write-Progress -Activity "STRESS TEST (Math Heavy)" -Status "Progression: $PercentDone%" -PercentComplete $PercentDone
        Start-Sleep 1 
    }
    Write-Progress -Activity "STRESS TEST" -Completed
    Write-Log "Fin du timer stress."

} finally {
    Write-Log "Nettoyage processus..."
    Get-Job | Stop-Job | Remove-Job -Force
    try { Remove-Item "$env:TEMP\io.dat" -ErrorAction SilentlyContinue } catch {}
}

# Analyse Batterie après stress
$SantePercent="N/A"
$NoteSanteChimique="N/A"
$InfoBat="Non détectée"

Start-Process powercfg -ArgumentList "/batteryreport /output `"$BatReportPath`" /XML" -Wait -WindowStyle Hidden
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
            Write-Log "Batterie Santé: $([math]::Round($SantePercent))%" -Level SUCCESS
        }
    } catch { 
        Write-Log "XML Batterie invalide." -Level WARN 
    }
    Remove-Item -LiteralPath $BatReportPath -Force -ErrorAction SilentlyContinue
}

# Calcul Perte et Note Stress
$DropPercent = 0
$NoteStress = "N/A"
$BatteryMessage = "Ignoré (Secteur)"

if (-not $IsPluggedIn) {
    $EndP = 0
    try { $EndP = (Get-WmiObject Win32_Battery).EstimatedChargeRemaining } catch { $EndP = $StartP }
    $DropPercent = $StartP - $EndP
    $BatteryMessage = "Perte: $DropPercent%"
    
    Write-Log "Drop batterie: $StartP% -> $EndP% (Perte: $DropPercent%)"
    
    if ($DropPercent -le 2) { $NoteStress = 20 } 
    elseif ($DropPercent -le 4) { $NoteStress = 15 } 
    elseif ($DropPercent -le 6) { $NoteStress = 10 } 
    elseif ($DropPercent -le 8) { $NoteStress = 5 } 
    else { $NoteStress = 0 }
}

# Note Finale Batterie (Moyenne Santé + Stress)
if ($NoteSanteChimique -ne "N/A" -and $NoteStress -ne "N/A") { 
    $NoteBatFinale = [math]::Round(($NoteSanteChimique + $NoteStress) / 2, 1) 
} elseif ($NoteSanteChimique -ne "N/A") { 
    $NoteBatFinale = [math]::Round($NoteSanteChimique, 1) 
} else { 
    $NoteBatFinale = "N/A" 
}

# =========================================================
#  7. VERDICT & RAPPORT FINAL
# =========================================================
$MoyennePerf = ($NoteCPU + $NoteRAM + $NoteGPU + $NoteDisk) / 4

if ($NoteBatFinale -eq "N/A") { 
    $NoteGlobale = [math]::Round($MoyennePerf, 1)
    $FlagProbleme = $false 
} else {
    $NoteGlobale = [math]::Round(($MoyennePerf * 0.5) + ($NoteBatFinale * 0.5), 1)
    $FlagProbleme = $false
    $RaisonProbleme = ""
    
    # Vérification des critères d'échec
    if ($SantePercent -ne "N/A" -and $SantePercent -lt $Config.Criteres.Batterie_Min_Sante) { 
        $FlagProbleme = $true; $RaisonProbleme = "BATTERIE HS" 
    }
    if ($NoteStress -ne "N/A" -and $DropPercent -ge $Config.Criteres.Batterie_Max_Perte) { 
        $FlagProbleme = $true; $RaisonProbleme = "BATTERIE INSTABLE" 
    }
    if ($NoteSmart -eq 0) { 
        $FlagProbleme = $true; $RaisonProbleme = "DISQUE HS" 
    }
    
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

Write-Host "[5/5] Génération Rapport..." -ForegroundColor Green

$SanteDisplay = if($SantePercent -ne "N/A"){"$([math]::Round($SantePercent,0))%"}else{"N/A"}
$StressDisplay = if($NoteStress -ne "N/A"){"$NoteStress / 20"}else{"Non Noté (Secteur)"}

# Capture de l'heure exacte de FIN pour le rapport
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
    Write-Log "Rapport TXT écrit : $RapportPath" -Level SUCCESS
} catch { 
    Write-Log "Echec écriture Rapport TXT." -Level ERROR 
}

# Écriture dans le CSV Global
if (-not (Test-Path $CsvPath)) { 
    Write-SafeCsv -Path $CsvPath -Content "Date;MachineID;Modele;Note_Globale;Verdict;CPU_Score;Bat_Sante;Bat_Drop" 
}
$CsvLine = "$ReportDate;$FormattedName;$ModelePC;$NoteGlobale;$Verdict;$ScoreCPU;$SanteDisplay;$DropPercent"
Write-SafeCsv -Path $CsvPath -Content $CsvLine

Write-Host ""
Write-Log "Session terminée."
Read-Host "Terminé. Entrée..."
