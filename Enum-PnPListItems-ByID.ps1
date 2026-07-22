<#
.SYNOPSIS
    Enumère tous les items d'une bibliothèque SharePoint volumineuse via la pagination
    native -PageSize de Get-PnPListItem (sans CAML), pour éviter le List View Threshold (5000).

.DESCRIPTION
    - Utilise Get-PnPListItem -PageSize + -ScriptBlock : PnP.PowerShell pagine en interne
      via ListItemCollectionPosition, ce qui contourne fiablement le threshold même sur
      des bibliothèques de 100 000+ éléments (contrairement à un filtre CAML <Where> sur
      un champ indexé, qui peut encore déclencher l'erreur dans certains cas).
    - Le ScriptBlock est invoqué à chaque page reçue : export CSV incrémental + mise à
      jour du resume file, sans attendre la fin de l'énumération complète.
    - Reprise : les items dont l'ID est <= au dernier ID connu (StartId ou resume file)
      sont filtrés côté client et ignorés, sans être ré-exportés.
    - Gestion des erreurs transitoires (throttling / timeout) avec retry et backoff.


.PARAMETER SiteUrl
    URL du site SharePoint (ex: https://tenant.sharepoint.com/sites/MonSite)

.PARAMETER ListName
    Nom de la bibliothèque (ex: "ALD Documents")

.PARAMETER OutputCsv
    Chemin du fichier CSV de sortie

.PARAMETER ResumeFile
    Chemin du fichier JSON de reprise (créé/mis à jour automatiquement)

.PARAMETER PageSize
    Taille de page CAML (1000 recommandé)

.PARAMETER StartId
    Force un ID de départ manuel (prioritaire sur le fichier resume). Utile pour
    redémarrer depuis 13158346 comme dans ton cas.

.EXAMPLE
    .\Enum-PnPListItems-ByID.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MonSite" `
        -ListName "ALD Documents" -StartId 13158346
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ListName,

    [string]$OutputCsv = ".\Export_$($ListName -replace '\s','_').csv",

    [string]$ResumeFile = ".\Resume_$($ListName -replace '\s','_').json",

    [int]$PageSize = 1000,

    [Nullable[int]]$StartId = $null,

    [string[]]$Fields = @('ID', 'FileLeafRef', 'FileDirRef', 'Modified', 'File_x0020_Size'),

    [int]$MaxRetries = 5
)

# ---------- Connexion ----------
# Réutilise la connexion PnP déjà active dans la session (si elle existe et pointe
# vers le bon site) plutôt que de forcer une reconnexion interactive à chaque lancement.
$needsConnection = $true

try {
    $currentConnection = Get-PnPConnection -ErrorAction Stop
    if ($currentConnection -and $currentConnection.Url -eq $SiteUrl) {
        # Test rapide que la connexion est toujours valide (token non expiré, etc.)
        Get-PnPWeb -ErrorAction Stop | Out-Null
        Write-Host "Connexion PnP existante réutilisée ($SiteUrl)." -ForegroundColor Green
        $needsConnection = $false
    }
    elseif ($currentConnection) {
        Write-Host "Une connexion PnP existe mais pointe vers un autre site ($($currentConnection.Url)). Reconnexion nécessaire." -ForegroundColor Yellow
    }
}
catch {
    # Pas de connexion active, ou connexion expirée/invalide
    $needsConnection = $true
}

if ($needsConnection) {
    Write-Host "Connexion à $SiteUrl ..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
    Write-Host "Connecté." -ForegroundColor Green
}

# ---------- Reprise ----------
$lastId = 0

if ($StartId) {
    $lastId = $StartId
    Write-Host "Démarrage forcé à partir de l'ID $lastId (paramètre -StartId)." -ForegroundColor Yellow
}
elseif (Test-Path $ResumeFile) {
    try {
        $resumeData = Get-Content $ResumeFile -Raw | ConvertFrom-Json
        $lastId = $resumeData.LastProcessedId
        Write-Host "Fichier de reprise trouvé. Reprise à partir de l'ID $lastId." -ForegroundColor Yellow
    }
    catch {
        Write-Warning "Fichier de reprise illisible, redémarrage depuis 0."
        $lastId = 0
    }
}
else {
    Write-Host "Aucune reprise trouvée, démarrage depuis l'ID 0." -ForegroundColor Yellow
}

# ---------- Préparation CSV ----------
$totalItems = 0
$skippedAlreadyDone = 0
$script:pageNumber = 0

Write-Host "`nDébut de l'énumération (méthode -PageSize native, sans filtre CAML)..." -ForegroundColor Cyan
if ($lastId -gt 0) {
    Write-Host "Les items avec ID <= $lastId seront ignorés (déjà traités précédemment)." -ForegroundColor Yellow
}

# IMPORTANT : on n'utilise plus de CAML <Where> sur ID pour filtrer/paginer.
# Retour d'expérience terrain (SharePoint Diary, MS Q&A) : même un <Where><Gt> sur un
# champ indexé (ID, Created...) peut encore déclencher le list view threshold.
# La méthode fiable et documentée est d'appeler Get-PnPListItem SANS -Query, avec
# uniquement -PageSize : PnP.PowerShell gère alors la pagination en interne via
# ListItemCollectionPosition (mécanisme différent du rendu de vue CAML), ce qui
# contourne réellement le threshold, y compris sur des bibliothèques de 100 000+ items.
#
# -ScriptBlock est appelé une fois PAR PAGE reçue : on l'utilise pour exporter en CSV
# au fur et à mesure et mettre à jour le resume file, sans attendre la fin de
# l'énumération complète (qui peut prendre longtemps sur 130 000 fichiers).

$scriptBlock = {
    param($pageItems)

    $script:pageNumber++
    $pageCount = @($pageItems).Count
    if ($pageCount -eq 0) { return }

    # Filtre côté client : on saute les IDs déjà traités lors d'une exécution précédente.
    $newItems = $pageItems | Where-Object { $_.Id -gt $lastId }
    $script:skippedAlreadyDone += ($pageCount - @($newItems).Count)

    if (@($newItems).Count -eq 0) {
        Write-Host "Page $($script:pageNumber) : $pageCount items, tous déjà traités (skip)." -ForegroundColor DarkGray
        return
    }

    $rows = foreach ($item in $newItems) {
        [PSCustomObject]@{
            ID          = $item.Id
            FileLeafRef = $item["FileLeafRef"]
            FileDirRef  = $item["FileDirRef"]
            Modified    = $item["Modified"]
            SizeBytes   = $item["File_x0020_Size"]
        }
    }

    if (-not (Test-Path $OutputCsv)) {
        $rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    }
    else {
        $rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Append
    }

    $script:totalItems += @($newItems).Count
    $maxIdInPage = ($newItems | Measure-Object -Property Id -Maximum).Maximum

    # Sauvegarde du resume file après CHAQUE page, pas seulement à la fin
    @{
        LastProcessedId = $maxIdInPage
        LastUpdated     = (Get-Date).ToString("o")
        TotalSoFar      = $script:totalItems
        ListName        = $ListName
    } | ConvertTo-Json | Set-Content -Path $ResumeFile -Encoding UTF8

    Write-Host "Page $($script:pageNumber) : $($newItems.Count) nouveaux items exportés (total cumulé: $($script:totalItems)) - max ID vu: $maxIdInPage" -ForegroundColor Gray
}

$attempt = 0
$success = $false

while (-not $success -and $attempt -lt $MaxRetries) {
    $attempt++
    try {
        # Appel unique : PnP.PowerShell boucle en interne sur toutes les pages de $PageSize
        # jusqu'à la fin de la bibliothèque, en invoquant $scriptBlock à chaque page.
        Get-PnPListItem -List $ListName -PageSize $PageSize -Fields $Fields -ScriptBlock $scriptBlock -ErrorAction Stop | Out-Null
        $success = $true
    }
    catch {
        Write-Warning "Erreur (tentative $attempt/$MaxRetries) : $($_.Exception.Message)"
        if ($attempt -lt $MaxRetries) {
            $waitSeconds = [Math]::Pow(2, $attempt)
            Write-Host "Nouvelle tentative dans $waitSeconds s (reprendra après le dernier ID sauvegardé : $lastId)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $waitSeconds
            # Recharge le dernier ID connu pour que le filtre côté client saute bien
            # ce qui a déjà été exporté avant le plantage
            if (Test-Path $ResumeFile) {
                $lastId = (Get-Content $ResumeFile -Raw | ConvertFrom-Json).LastProcessedId
            }
        }
        else {
            Write-Error "Échec définitif après $MaxRetries tentatives. Arrêt du script. Relance-le tel quel : le resume file permettra de reprendre ici."
            throw
        }
    }
}

Write-Host "`n=== Terminé ===" -ForegroundColor Green
Write-Host "Total d'items énumérés : $totalItems"
Write-Host "CSV : $OutputCsv"
Write-Host "Resume file : $ResumeFile (peut être supprimé si l'énumération est complète)"
