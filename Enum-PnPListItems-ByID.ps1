<#
.SYNOPSIS
    Enumère tous les items d'une bibliothèque SharePoint volumineuse en paginant
    sur le champ ID (indexe natif), pour contourner le List View Threshold (5000).

.DESCRIPTION
    - Filtre CAML sur <Gt> ID + <OrderBy> ID Ascending => pas de blocage threshold,
      même sur des bibliothèques de 100 000+ éléments.
    - Reprise automatique : le dernier ID traité est sauvegardé dans un fichier
      resume (.json). Si le script plante ou est interrompu, le relancer reprend
      exactement là où il s'était arrêté.
    - Export progressif en CSV (append), pour ne rien perdre en cas de coupure.
    - Gestion des erreurs transitoires (throttling / timeout) avec retry.

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
Write-Host "Connexion à $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
Write-Host "Connecté." -ForegroundColor Green

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
$csvExists = Test-Path $OutputCsv
if (-not $csvExists) {
    # En-tête créé au premier batch pour matcher les colonnes demandées
}

$totalItems = 0
$batchNumber = 0
$fieldsRefXml = ($Fields | ForEach-Object { "<FieldRef Name='$_'/>" }) -join ""

Write-Host "`nDébut de l'énumération..." -ForegroundColor Cyan

do {
    $camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Gt>
        <FieldRef Name='ID'/>
        <Value Type='Number'>$lastId</Value>
      </Gt>
    </Where>
    <OrderBy>
      <FieldRef Name='ID' Ascending='TRUE'/>
    </OrderBy>
  </Query>
  <RowLimit Paged='TRUE'>$PageSize</RowLimit>
</View>
"@

    $attempt = 0
    $items = $null
    $success = $false

    while (-not $success -and $attempt -lt $MaxRetries) {
        $attempt++
        try {
            $items = Get-PnPListItem -List $ListName -Query $camlQuery -PageSize $PageSize -Fields $Fields -ErrorAction Stop
            $success = $true
        }
        catch {
            Write-Warning "Erreur (tentative $attempt/$MaxRetries) à partir de l'ID $lastId : $($_.Exception.Message)"
            if ($attempt -lt $MaxRetries) {
                $waitSeconds = [Math]::Pow(2, $attempt)  # backoff exponentiel : 2, 4, 8, 16, 32s
                Write-Host "Nouvelle tentative dans $waitSeconds s..." -ForegroundColor Yellow
                Start-Sleep -Seconds $waitSeconds
            }
            else {
                Write-Error "Échec définitif après $MaxRetries tentatives à partir de l'ID $lastId. Arrêt du script. Relance-le : le resume file permettra de reprendre ici."
                throw
            }
        }
    }

    $count = @($items).Count
    if ($count -eq 0) {
        Write-Host "Aucun item supplémentaire. Enumération terminée." -ForegroundColor Green
        break
    }

    $batchNumber++
    $totalItems += $count

    # Construction des lignes CSV
    $rows = foreach ($item in $items) {
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

    # Dernier ID du batch = nouveau point de reprise
    $lastId = ($items | Select-Object -Last 1).Id

    # Sauvegarde du fichier de reprise après CHAQUE batch (pas seulement à la fin)
    @{
        LastProcessedId = $lastId
        LastUpdated     = (Get-Date).ToString("o")
        TotalSoFar      = $totalItems
        ListName        = $ListName
    } | ConvertTo-Json | Set-Content -Path $ResumeFile -Encoding UTF8

    Write-Host "Batch $batchNumber : $count items (total cumulé: $totalItems) - dernier ID: $lastId" -ForegroundColor Gray

} while ($count -eq $PageSize)  # si le batch renvoie moins que PageSize, c'est le dernier

Write-Host "`n=== Terminé ===" -ForegroundColor Green
Write-Host "Total d'items énumérés : $totalItems"
Write-Host "CSV : $OutputCsv"
Write-Host "Resume file : $ResumeFile (peut être supprimé si l'énumération est complète)"
