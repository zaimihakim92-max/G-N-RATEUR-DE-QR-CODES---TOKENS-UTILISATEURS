<#
.SYNOPSIS
    Déplace un très grand nombre de fichiers SharePoint Online par paquets, via Start-PnPCopyJob,
    avec validation après chaque paquet et reprise possible en cas d'interruption.

.DESCRIPTION
    - Lit la liste des noms de fichiers depuis un fichier texte (un nom par ligne)
    - Traite par lots de $BatchSize (défaut 100)
    - Pour chaque lot : lance le job de déplacement, attend sa fin, vérifie que les fichiers
      sont bien présents dans le dossier cible et absents du dossier source
    - Journalise les succès/échecs dans un CSV
    - Reprend automatiquement là où il s'est arrêté si relancé (skip des lots déjà validés)

.NOTES
    Nécessite le module PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$SourceFolderServerRelativeUrl,   # ex: /sites/ayvens_INDIA/Shared Documents/General/Document Upload/TAXINVOICE

    [Parameter(Mandatory = $true)]
    [string]$DestinationFolderUrl,           # ex: https://tenant.sharepoint.com/sites/ayvens_INDIA/Shared Documents/General/Document Upload/C1/TAXINVOICE

    [Parameter(Mandatory = $true)]
    [string]$FileListPath,                   # ex: C:\Temp\TaxInvoice100.txt

    [int]$BatchSize = 100,

    [string]$LogPath = "C:\Temp\MoveLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [string]$ProgressStatePath = "C:\Temp\MoveProgress.json",

    [int]$JobPollSeconds = 5,

    [int]$JobTimeoutMinutes = 15
)

# ---------------------------------------------------------------------------
# Connexion
# ---------------------------------------------------------------------------
Write-Host "Connexion à $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# ---------------------------------------------------------------------------
# Chargement de la liste des fichiers
# ---------------------------------------------------------------------------
if (-not (Test-Path $FileListPath)) {
    throw "Fichier de liste introuvable : $FileListPath"
}
$allFiles = Get-Content $FileListPath | Where-Object { $_.Trim() -ne "" }
$total = $allFiles.Count
Write-Host "Total de fichiers à déplacer : $total" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Découpage en lots
# ---------------------------------------------------------------------------
$batches = for ($i = 0; $i -lt $allFiles.Count; $i += $BatchSize) {
    ,@($allFiles[$i..([Math]::Min($i + $BatchSize - 1, $allFiles.Count - 1))])
}
$totalBatches = $batches.Count
Write-Host "Nombre de lots ($BatchSize fichiers/lot) : $totalBatches" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Reprise : lots déjà validés
# ---------------------------------------------------------------------------
$startBatchIndex = 0
if (Test-Path $ProgressStatePath) {
    try {
        $state = Get-Content $ProgressStatePath -Raw | ConvertFrom-Json
        if ($state.LastCompletedBatch -is [int]) {
            $startBatchIndex = $state.LastCompletedBatch + 1
            Write-Host "Reprise détectée : reprise au lot $($startBatchIndex + 1)/$totalBatches" -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Impossible de lire l'état de progression, redémarrage depuis le début."
    }
}

# Init du log CSV si nouveau
if (-not (Test-Path $LogPath)) {
    "BatchNumber,FileName,Status,Detail,Timestamp" | Out-File -FilePath $LogPath -Encoding UTF8
}

function Write-LogEntry {
    param($BatchNumber, $FileName, $Status, $Detail)
    $line = '{0},{1},{2},"{3}",{4}' -f $BatchNumber, $FileName, $Status, $Detail, (Get-Date -Format 'o')
    Add-Content -Path $LogPath -Value $line
}

function Save-Progress {
    param([int]$BatchIndex)
    @{ LastCompletedBatch = $BatchIndex; UpdatedAt = (Get-Date -Format 'o') } |
        ConvertTo-Json | Out-File -FilePath $ProgressStatePath -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Boucle principale
# ---------------------------------------------------------------------------
$globalSuccess = 0
$globalFailed  = 0

for ($b = $startBatchIndex; $b -lt $totalBatches; $b++) {

    $batchNumber = $b + 1
    $batchFiles  = $batches[$b]
    Write-Host "`n=== Lot $batchNumber / $totalBatches ($($batchFiles.Count) fichiers) ===" -ForegroundColor Green

    $sourceUrls = $batchFiles | ForEach-Object {
        "$SourceFolderServerRelativeUrl/$_".Replace("//", "/")
    }

    # --- Lancement du job de déplacement ---
    try {
        $job = Start-PnPCopyJob -SourceUrl $sourceUrls `
                                 -DestinationUrl $DestinationFolderUrl `
                                 -Overwrite `
                                 -Move `
                                 -ErrorAction Stop
    } catch {
        Write-Warning "Échec du lancement du job pour le lot $batchNumber : $($_.Exception.Message)"
        foreach ($f in $batchFiles) { Write-LogEntry $batchNumber $f "JOB_START_FAILED" $_.Exception.Message }
        $globalFailed += $batchFiles.Count
        continue   # on passe au lot suivant, celui-ci ne sera pas marqué comme validé
    }

    # --- Attente de la fin du job ---
    $elapsedSeconds = 0
    $timeoutSeconds = $JobTimeoutMinutes * 60
    $jobState = $null

    do {
        Start-Sleep -Seconds $JobPollSeconds
        $elapsedSeconds += $JobPollSeconds
        try {
            $status = Get-PnPCopyJobStatus -CopyJobInfo $job -ErrorAction Stop
            $jobState = $status.JobState
        } catch {
            Write-Warning "Erreur en interrogeant le statut du job (lot $batchNumber) : $($_.Exception.Message)"
            $jobState = -1
            break
        }
        Write-Host "  Statut job : $jobState (écoulé : ${elapsedSeconds}s)" -ForegroundColor DarkGray
    } while ($jobState -ne 0 -and $elapsedSeconds -lt $timeoutSeconds)

    if ($elapsedSeconds -ge $timeoutSeconds) {
        Write-Warning "Timeout dépassé pour le lot $batchNumber, passage à la validation quand même."
    }

    # --- Validation : présence dans la destination, absence dans la source ---
    $batchSuccess = 0
    $batchFailed  = 0

    foreach ($f in $batchFiles) {
        $destCheckUrl   = "$DestinationFolderUrl/$f".Replace("//", "/") -replace "^https://[^/]+", ""
        $sourceCheckUrl = "$SourceFolderServerRelativeUrl/$f".Replace("//", "/")

        $existsAtDest = $null
        $existsAtSrc  = $null

        try { $existsAtDest = Get-PnPFile -Url $destCheckUrl -AsFile -ErrorAction Stop } catch { $existsAtDest = $null }
        try { $existsAtSrc  = Get-PnPFile -Url $sourceCheckUrl -AsFile -ErrorAction Stop } catch { $existsAtSrc = $null }

        if ($existsAtDest -and -not $existsAtSrc) {
            Write-LogEntry $batchNumber $f "OK" "Déplacé et validé"
            $batchSuccess++
        }
        elseif ($existsAtDest -and $existsAtSrc) {
            Write-LogEntry $batchNumber $f "DUPLICATE" "Présent aux deux emplacements"
            $batchFailed++
        }
        else {
            Write-LogEntry $batchNumber $f "MISSING" "Introuvable à destination"
            $batchFailed++
        }
    }

    $globalSuccess += $batchSuccess
    $globalFailed  += $batchFailed

    Write-Host "  Lot $batchNumber : $batchSuccess OK / $batchFailed échec(s)" -ForegroundColor $(if ($batchFailed -eq 0) { "Green" } else { "Red" })

    # On ne marque le lot comme "complété" (pour la reprise) que s'il n'y a aucun échec
    if ($batchFailed -eq 0) {
        Save-Progress -BatchIndex $b
    } else {
        Write-Warning "Lot $batchNumber contient des échecs : il ne sera PAS marqué comme complété. Corrige puis relance le script (il refera ce lot)."
    }
}

# ---------------------------------------------------------------------------
# Résumé final
# ---------------------------------------------------------------------------
Write-Host "`n===================== RÉSUMÉ =====================" -ForegroundColor Cyan
Write-Host "Total fichiers      : $total"
Write-Host "Réussis (validés)   : $globalSuccess" -ForegroundColor Green
Write-Host "Échecs / à revoir   : $globalFailed" -ForegroundColor $(if ($globalFailed -gt 0) { "Red" } else { "Green" })
Write-Host "Log détaillé        : $LogPath"
Write-Host "État de reprise     : $ProgressStatePath"
Write-Host "===================================================" -ForegroundColor Cyan
