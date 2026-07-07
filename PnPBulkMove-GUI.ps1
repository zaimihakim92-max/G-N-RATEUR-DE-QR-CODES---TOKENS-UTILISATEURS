<#
.SYNOPSIS
    Outil GUI - Déplacement massif de fichiers SharePoint Online (PnP.PowerShell)
    par paquets, avec validation et reprise automatique.

.DESCRIPTION
    - Champs : URL du site, dossier source, dossier destination, fichier liste (.txt, un nom/ligne)
    - Taille de paquet configurable
    - Validation après chaque paquet (présence à destination / absence à la source)
    - Reprise automatique basée sur un fichier d'état (hash source+destination)
    - Log CSV détaillé, mode simulation disponible
    - Traitement en arrière-plan (runspace) : l'interface reste réactive

.NOTES
    Prérequis : module PnP.PowerShell installé (Install-Module PnP.PowerShell -Scope CurrentUser)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# Détection d'une connexion PnP existante dans CETTE fenêtre PowerShell
# (optionnelle : l'outil permet aussi d'en créer une nouvelle depuis l'interface)
# ============================================================================
$script:ReusedConnection = $null
try { $script:ReusedConnection = Get-PnPConnection -ErrorAction Stop } catch { $script:ReusedConnection = $null }
$script:CurrentPnPConnection = $script:ReusedConnection

# ============================================================================
# Fonctions utilitaires
# ============================================================================
function Get-StateHash {
    param([string]$Source, [string]$Destination)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("$Source|$Destination")
    $md5   = [System.Security.Cryptography.MD5]::Create()
    ($md5.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ''
}

function Get-CompletedBatches {
    param($ResumeFile)
    if (Test-Path $ResumeFile) {
        try {
            $state = Get-Content $ResumeFile -Raw | ConvertFrom-Json
            if ($state.CompletedBatches) { return @($state.CompletedBatches | ForEach-Object { [int]$_ }) }
            elseif ($state.LastCompletedBatch -is [int]) { return @(1..($state.LastCompletedBatch + 1)) } # ancien format (0-based)
        } catch {}
    }
    return @()
}

function Get-NextSuggestedBatch {
    param($CompletedBatches, [int]$TotalBatches)
    for ($i = 1; $i -le $TotalBatches; $i++) {
        if ($CompletedBatches -notcontains $i) { return $i }
    }
    return $TotalBatches + 1
}

# ============================================================================
# État partagé entre l'UI et le traitement en arrière-plan
# ============================================================================
$sync = [hashtable]::Synchronized(@{
    LogQueue      = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    CurrentBatch  = 0
    TotalBatches  = 0
    GlobalSuccess = 0
    GlobalFailed  = 0
    Done          = $false
    AllDone       = $false
    Cancel        = $false
    AuditSummary  = $null
    Params        = @{}
})

# ============================================================================
# Formulaire principal
# ============================================================================
$form                     = New-Object System.Windows.Forms.Form
$form.Text                = "PnP Bulk Move - Déplacement massif SharePoint"
$form.Size                = New-Object System.Drawing.Size(780, 905)
$form.StartPosition       = "CenterScreen"
$form.FormBorderStyle     = 'FixedDialog'
$form.MaximizeBox         = $false
$form.Font                = New-Object System.Drawing.Font("Segoe UI", 9)

# --- GroupBox Connexion / Emplacements ---
$gbConn = New-Object System.Windows.Forms.GroupBox
$gbConn.Text = "Connexion et emplacements"
$gbConn.Location = New-Object System.Drawing.Point(15, 15)
$gbConn.Size = New-Object System.Drawing.Size(735, 290)
$form.Controls.Add($gbConn)

function New-LabeledTextBox {
    param($Parent, $LabelText, $Y, $Width = 700)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Location = New-Object System.Drawing.Point(10, $Y)
    $lbl.Size = New-Object System.Drawing.Size(700, 18)
    $Parent.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(10, ($Y + 20))
    $tb.Size = New-Object System.Drawing.Size($Width, 22)
    $Parent.Controls.Add($tb)
    return $tb
}

$radioReuse = New-Object System.Windows.Forms.RadioButton
$radioReuse.Text = "Réutiliser la connexion PnP active de cette fenêtre PowerShell"
$radioReuse.Location = New-Object System.Drawing.Point(10, 20)
$radioReuse.Size = New-Object System.Drawing.Size(700, 20)
$gbConn.Controls.Add($radioReuse)

$radioNew = New-Object System.Windows.Forms.RadioButton
$radioNew.Text = "Se connecter maintenant (nouvelle connexion)"
$radioNew.Location = New-Object System.Drawing.Point(10, 42)
$radioNew.Size = New-Object System.Drawing.Size(700, 20)
$gbConn.Controls.Add($radioNew)

$txtSite = New-LabeledTextBox -Parent $gbConn -LabelText "URL du site" -Y 68 -Width 590
$txtSource = New-LabeledTextBox -Parent $gbConn -LabelText "Dossier SOURCE (chemin relatif serveur, ex: /sites/MonSite/Shared Documents/Dossier)" -Y 168
$txtDest   = New-LabeledTextBox -Parent $gbConn -LabelText "Dossier DESTINATION (URL complète ou chemin relatif serveur - les deux fonctionnent)" -Y 216

$lblClientId = New-Object System.Windows.Forms.Label
$lblClientId.Text = "Client ID Entra ID (app enregistrée, requis pour une nouvelle connexion - le client public PnP a été retiré par Microsoft)"
$lblClientId.Location = New-Object System.Drawing.Point(10, 115)
$lblClientId.Size = New-Object System.Drawing.Size(700, 18)
$gbConn.Controls.Add($lblClientId)

$txtClientId = New-Object System.Windows.Forms.TextBox
$txtClientId.Location = New-Object System.Drawing.Point(10, 135)
$txtClientId.Size = New-Object System.Drawing.Size(460, 22)
$gbConn.Controls.Add($txtClientId)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Se connecter"
$btnConnect.Location = New-Object System.Drawing.Point(480, 134)
$btnConnect.Size = New-Object System.Drawing.Size(120, 24)
$gbConn.Controls.Add($btnConnect)

$lblConnInfo = New-Object System.Windows.Forms.Label
$lblConnInfo.Location = New-Object System.Drawing.Point(10, 262)
$lblConnInfo.Size = New-Object System.Drawing.Size(715, 18)
$lblConnInfo.ForeColor = [System.Drawing.Color]::DarkGreen
$gbConn.Controls.Add($lblConnInfo)

function Set-ConnectionModeUI {
    if ($radioReuse.Checked) {
        $txtSite.ReadOnly = $true
        $txtSite.BackColor = [System.Drawing.Color]::WhiteSmoke
        $txtSite.Text = $script:ReusedConnection.Url
        $txtClientId.Enabled = $false
        $btnConnect.Enabled = $false
        $script:CurrentPnPConnection = $script:ReusedConnection
        try { $u = (Get-PnPContext -Connection $script:ReusedConnection).Credentials.UserPrincipalName } catch { $u = $null }
        if (-not $u) { $u = "session active" }
        $lblConnInfo.Text = "Connexion réutilisée : $u"
        $lblConnInfo.ForeColor = [System.Drawing.Color]::DarkGreen
    } else {
        $txtSite.ReadOnly = $false
        $txtSite.BackColor = [System.Drawing.Color]::White
        $txtClientId.Enabled = $true
        $btnConnect.Enabled = $true
        $lblConnInfo.Text = "Renseigne l'URL du site + le Client ID, puis clique sur 'Se connecter'."
        $lblConnInfo.ForeColor = [System.Drawing.Color]::DarkOrange
    }
}

if ($script:ReusedConnection) {
    $radioReuse.Checked = $true
} else {
    $radioReuse.Enabled = $false
    $radioNew.Checked = $true
}
$radioReuse.Add_CheckedChanged({ if ($radioReuse.Checked) { Set-ConnectionModeUI } })
$radioNew.Add_CheckedChanged({ if ($radioNew.Checked) { Set-ConnectionModeUI } })
Set-ConnectionModeUI

$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Renseigne l'URL du site.", "Champ manquant", 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtClientId.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Un Client ID d'application Entra ID est requis.`n`nLe client public multi-tenant 'PnP Management Shell' a été retiré par Microsoft en septembre 2024 : il faut désormais ta propre application enregistrée.`n`nTu peux en créer une automatiquement avec :`nRegister-PnPEntraIDAppForInteractiveLogin -ApplicationName ""MonApp"" -Tenant tontenant.onmicrosoft.com",
            "Client ID requis", 'OK', 'Warning') | Out-Null
        return
    }
    $lblConnInfo.Text = "Connexion en cours..."
    $lblConnInfo.ForeColor = [System.Drawing.Color]::DarkOrange
    $btnConnect.Enabled = $false
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $newConn = Connect-PnPOnline -Url $txtSite.Text -ClientId $txtClientId.Text -Interactive -ReturnConnection -ErrorAction Stop
    } catch {
        try {
            $lblConnInfo.Text = "Authentification interactive impossible, tentative DeviceLogin (suis les instructions dans une fenêtre de console)..."
            [System.Windows.Forms.Application]::DoEvents()
            $newConn = Connect-PnPOnline -Url $txtSite.Text -ClientId $txtClientId.Text -DeviceLogin -ReturnConnection -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Échec de connexion :`n$($_.Exception.Message)", "Erreur", 'OK', 'Error') | Out-Null
            $lblConnInfo.Text = "Échec de connexion."
            $lblConnInfo.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
            return
        }
    }
    $script:CurrentPnPConnection = $newConn
    try { $u = (Get-PnPContext -Connection $newConn).Credentials.UserPrincipalName } catch { $u = $null }
    if (-not $u) { $u = "connecté" }
    $lblConnInfo.Text = "Nouvelle connexion établie : $u"
    $lblConnInfo.ForeColor = [System.Drawing.Color]::DarkGreen
    $btnConnect.Enabled = $true
})

# --- GroupBox Fichiers / paramètres de traitement ---
$gbParams = New-Object System.Windows.Forms.GroupBox
$gbParams.Text = "Liste des fichiers et paramètres"
$gbParams.Location = New-Object System.Drawing.Point(15, 315)
$gbParams.Size = New-Object System.Drawing.Size(735, 190)
$form.Controls.Add($gbParams)

$lblFileList = New-Object System.Windows.Forms.Label
$lblFileList.Text = "Fichier liste (.txt, un nom de fichier par ligne, présents dans le dossier SOURCE)"
$lblFileList.Location = New-Object System.Drawing.Point(10, 15)
$lblFileList.Size = New-Object System.Drawing.Size(600, 18)
$gbParams.Controls.Add($lblFileList)

$txtFileList = New-Object System.Windows.Forms.TextBox
$txtFileList.Location = New-Object System.Drawing.Point(10, 35)
$txtFileList.Size = New-Object System.Drawing.Size(590, 22)
$gbParams.Controls.Add($txtFileList)

$btnBrowseFile = New-Object System.Windows.Forms.Button
$btnBrowseFile.Text = "Parcourir..."
$btnBrowseFile.Location = New-Object System.Drawing.Point(610, 34)
$btnBrowseFile.Size = New-Object System.Drawing.Size(115, 24)
$gbParams.Controls.Add($btnBrowseFile)
$btnBrowseFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    if ($dlg.ShowDialog() -eq 'OK') { $txtFileList.Text = $dlg.FileName }
})

$lblLogFolder = New-Object System.Windows.Forms.Label
$lblLogFolder.Text = "Dossier de logs / reprise"
$lblLogFolder.Location = New-Object System.Drawing.Point(10, 65)
$lblLogFolder.Size = New-Object System.Drawing.Size(590, 18)
$gbParams.Controls.Add($lblLogFolder)

$txtLogFolder = New-Object System.Windows.Forms.TextBox
$txtLogFolder.Location = New-Object System.Drawing.Point(10, 85)
$txtLogFolder.Size = New-Object System.Drawing.Size(590, 22)
$txtLogFolder.Text = "C:\Temp\PnPBulkMove"
$gbParams.Controls.Add($txtLogFolder)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Text = "Parcourir..."
$btnBrowseLog.Location = New-Object System.Drawing.Point(610, 84)
$btnBrowseLog.Size = New-Object System.Drawing.Size(115, 24)
$gbParams.Controls.Add($btnBrowseLog)
$btnBrowseLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq 'OK') { $txtLogFolder.Text = $dlg.SelectedPath }
})

$lblBatch = New-Object System.Windows.Forms.Label
$lblBatch.Text = "Taille des paquets"
$lblBatch.Location = New-Object System.Drawing.Point(10, 115)
$lblBatch.Size = New-Object System.Drawing.Size(120, 18)
$gbParams.Controls.Add($lblBatch)

$numBatch = New-Object System.Windows.Forms.NumericUpDown
$numBatch.Location = New-Object System.Drawing.Point(10, 135)
$numBatch.Size = New-Object System.Drawing.Size(80, 22)
$numBatch.Minimum = 1
$numBatch.Maximum = 5000
$numBatch.Value = 100
$gbParams.Controls.Add($numBatch)

$chkSimulate = New-Object System.Windows.Forms.CheckBox
$chkSimulate.Text = "Mode simulation (n'exécute pas le déplacement, journalise seulement ce qui serait fait)"
$chkSimulate.Location = New-Object System.Drawing.Point(150, 137)
$chkSimulate.Size = New-Object System.Drawing.Size(560, 20)
$gbParams.Controls.Add($chkSimulate)

$chkVerbose = New-Object System.Windows.Forms.CheckBox
$chkVerbose.Text = "Mode verbose (détail fichier par fichier : lancement des jobs, résultats individuels, chronométrage)"
$chkVerbose.Location = New-Object System.Drawing.Point(10, 160)
$chkVerbose.Size = New-Object System.Drawing.Size(700, 20)
$gbParams.Controls.Add($chkVerbose)

# --- Sélection manuelle du paquet à traiter ---
$gbBatchSelect = New-Object System.Windows.Forms.GroupBox
$gbBatchSelect.Text = "Paquet à traiter"
$gbBatchSelect.Location = New-Object System.Drawing.Point(15, 515)
$gbBatchSelect.Size = New-Object System.Drawing.Size(735, 55)
$form.Controls.Add($gbBatchSelect)

$lblBatchNum = New-Object System.Windows.Forms.Label
$lblBatchNum.Text = "N° du paquet :"
$lblBatchNum.Location = New-Object System.Drawing.Point(10, 24)
$lblBatchNum.Size = New-Object System.Drawing.Size(90, 20)
$gbBatchSelect.Controls.Add($lblBatchNum)

$numBatchToRun = New-Object System.Windows.Forms.NumericUpDown
$numBatchToRun.Location = New-Object System.Drawing.Point(105, 22)
$numBatchToRun.Size = New-Object System.Drawing.Size(80, 22)
$numBatchToRun.Minimum = 1
$numBatchToRun.Maximum = 999999
$numBatchToRun.Value = 1
$gbBatchSelect.Controls.Add($numBatchToRun)

$btnSuggest = New-Object System.Windows.Forms.Button
$btnSuggest.Text = "Suggérer le prochain non traité"
$btnSuggest.Location = New-Object System.Drawing.Point(195, 21)
$btnSuggest.Size = New-Object System.Drawing.Size(210, 24)
$gbBatchSelect.Controls.Add($btnSuggest)

$lblBatchInfo = New-Object System.Windows.Forms.Label
$lblBatchInfo.Text = "Clique sur 'Suggérer' pour voir le total de paquets et l'avancement."
$lblBatchInfo.Location = New-Object System.Drawing.Point(415, 24)
$lblBatchInfo.Size = New-Object System.Drawing.Size(310, 20)
$gbBatchSelect.Controls.Add($lblBatchInfo)

function Update-BatchSuggestion {
    if (-not (Test-Path $txtFileList.Text)) { return }
    try {
        $logFolder = $txtLogFolder.Text
        if (-not (Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
        $hash = Get-StateHash -Source $txtSource.Text -Destination $txtDest.Text
        $resumeFile = Join-Path $logFolder "Resume_$hash.json"

        $count = (Get-Content $txtFileList.Text | Where-Object { $_.Trim() -ne "" }).Count
        $batchSize = [int]$numBatch.Value
        $totalBatches = [Math]::Max([Math]::Ceiling($count / $batchSize), 1)
        $completed = Get-CompletedBatches -ResumeFile $resumeFile
        $next = Get-NextSuggestedBatch -CompletedBatches $completed -TotalBatches $totalBatches

        $numBatchToRun.Maximum = $totalBatches
        if ($next -gt $totalBatches) {
            $numBatchToRun.Value = $totalBatches
            $lblBatchInfo.Text = "Total : $totalBatches paquet(s) — tous déjà validés."
        } else {
            $numBatchToRun.Value = $next
            $lblBatchInfo.Text = "Total : $totalBatches paquet(s) — $($completed.Count) validé(s) — prochain : $next"
        }
    } catch {
        $lblBatchInfo.Text = "Impossible de calculer (vérifie le fichier liste / dossier de logs)."
    }
}

$btnSuggest.Add_Click({ Update-BatchSuggestion })

# --- Boutons d'action ---
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Traiter le paquet sélectionné"
$btnStart.Location = New-Object System.Drawing.Point(15, 580)
$btnStart.Size = New-Object System.Drawing.Size(200, 32)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(46, 125, 50)
$btnStart.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnStart)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Arrêter"
$btnStop.Location = New-Object System.Drawing.Point(225, 580)
$btnStop.Size = New-Object System.Drawing.Size(120, 32)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$btnResetResume = New-Object System.Windows.Forms.Button
$btnResetResume.Text = "Réinitialiser reprise"
$btnResetResume.Location = New-Object System.Drawing.Point(355, 580)
$btnResetResume.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($btnResetResume)

$btnAudit = New-Object System.Windows.Forms.Button
$btnAudit.Text = "Auditer le dossier destination"
$btnAudit.Location = New-Object System.Drawing.Point(15, 617)
$btnAudit.Size = New-Object System.Drawing.Size(490, 30)
$btnAudit.BackColor = [System.Drawing.Color]::FromArgb(21, 101, 192)
$btnAudit.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnAudit)

# --- Progression ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 657)
$progressBar.Size = New-Object System.Drawing.Size(735, 22)
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "En attente..."
$lblStatus.Location = New-Object System.Drawing.Point(15, 682)
$lblStatus.Size = New-Object System.Drawing.Size(735, 20)
$form.Controls.Add($lblStatus)

# --- Journal ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Journal :"
$lblLog.Location = New-Object System.Drawing.Point(15, 707)
$lblLog.Size = New-Object System.Drawing.Size(200, 18)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 727)
$txtLog.Size = New-Object System.Drawing.Size(735, 140)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::Black
$txtLog.ForeColor = [System.Drawing.Color]::LightGreen
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtLog)

# ============================================================================
# Script exécuté en arrière-plan (runspace)
# ============================================================================
$scriptBlock = {
    param($sync)

    $p = $sync.Params
    function Log { param($msg) $sync.LogQueue.Enqueue("$(Get-Date -Format 'HH:mm:ss')  $msg") }
    function VLog { param($msg) if ($p.Verbose) { $sync.LogQueue.Enqueue("$(Get-Date -Format 'HH:mm:ss')      [verbose] $msg") } }

    try {
        Import-Module PnP.PowerShell -ErrorAction Stop

        $conn = $p.Connection
        Log "Réutilisation de la connexion existante : $($conn.Url)"

        # Normalisation : accepte indifféremment une URL complète ou un chemin relatif serveur
        function Get-RelativeUrl { param($Url) ($Url -replace '^https://[^/]+', '').TrimEnd('/') }
        $srcFolder  = Get-RelativeUrl $p.SourceFolder
        $destFolder = Get-RelativeUrl $p.DestFolder
        VLog "Dossier source normalisé   : $srcFolder"
        VLog "Dossier destination normalisé : $destFolder"

        $allFiles = Get-Content $p.FileListPath | Where-Object { $_.Trim() -ne "" }
        $total = $allFiles.Count
        Log "Total de fichiers à traiter : $total"

        $batches = for ($i = 0; $i -lt $allFiles.Count; $i += $p.BatchSize) {
            ,@($allFiles[$i..([Math]::Min($i + $p.BatchSize - 1, $allFiles.Count - 1))])
        }
        $totalBatches = $batches.Count
        $sync.TotalBatches = $totalBatches
        Log "Nombre de paquets ($($p.BatchSize) fichiers/paquet) : $totalBatches"

        function Get-CompletedBatches {
            param($ResumeFile)
            if (Test-Path $ResumeFile) {
                try {
                    $state = Get-Content $ResumeFile -Raw | ConvertFrom-Json
                    if ($state.CompletedBatches) { return @($state.CompletedBatches | ForEach-Object { [int]$_ }) }
                    elseif ($state.LastCompletedBatch -is [int]) { return @(1..($state.LastCompletedBatch + 1)) }
                } catch {}
            }
            return @()
        }

        # Audit "intelligent" : UN SEUL appel de listing par dossier au lieu d'un Get-PnPFile
        # par fichier (c'est ce qui rendait la validation par paquet peu fiable : centaines
        # d'appels réseau successifs = plus de risques de timeout/throttling/erreurs transitoires).
        function Get-FolderFileSet {
            param($FolderRelativeUrl, $Connection)
            $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            try {
                $items = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderRelativeUrl -ItemType File -Connection $Connection -ErrorAction Stop
                foreach ($it in $items) { [void]$set.Add($it.Name) }
            } catch {
                throw "Impossible de lister le dossier '$FolderRelativeUrl' : $($_.Exception.Message)"
            }
            return $set
        }

        $completedBatches = Get-CompletedBatches -ResumeFile $p.ResumeFile

        if (-not (Test-Path $p.LogFile)) {
            "BatchNumber,FileName,Status,Detail,Timestamp" | Out-File -FilePath $p.LogFile -Encoding UTF8
        }

        if ($p.Mode -eq 'Audit') {
            Log "=== AUDIT COMPLET du dossier destination en cours (listing intégral source + destination) ==="
            $swList = [System.Diagnostics.Stopwatch]::StartNew()
            $destSet = Get-FolderFileSet -FolderRelativeUrl $destFolder -Connection $conn
            Log "  Destination : $($destSet.Count) fichier(s) trouvé(s)."
            VLog "Listing destination effectué en $($swList.Elapsed.TotalSeconds.ToString('0.0'))s"
            $swList.Restart()
            $srcSet = Get-FolderFileSet -FolderRelativeUrl $srcFolder -Connection $conn
            Log "  Source : $($srcSet.Count) fichier(s) trouvé(s)."
            VLog "Listing source effectué en $($swList.Elapsed.TotalSeconds.ToString('0.0'))s"

            $okCount = 0; $pendingCount = 0; $dupCount = 0; $missingCount = 0
            $auditFile = $p.AuditFile
            "FileName,InSource,InDestination,Status,Timestamp" | Out-File -FilePath $auditFile -Encoding UTF8

            $swCompare = [System.Diagnostics.Stopwatch]::StartNew()
            $processed = 0
            foreach ($f in $allFiles) {
                $inDest = $destSet.Contains($f)
                $inSrc  = $srcSet.Contains($f)
                if ($inDest -and -not $inSrc) { $status = "OK_DEPLACE"; $okCount++ }
                elseif ($inDest -and $inSrc)  { $status = "DUPLIQUE_AUX_DEUX_ENDROITS"; $dupCount++ }
                elseif (-not $inDest -and $inSrc) { $status = "PAS_ENCORE_DEPLACE"; $pendingCount++ }
                else { $status = "INTROUVABLE_AUX_DEUX_ENDROITS"; $missingCount++ }
                Add-Content -Path $auditFile -Value ('{0},{1},{2},{3},{4}' -f $f, $inSrc, $inDest, $status, (Get-Date -Format 'o'))
                $processed++
                if ($p.Verbose -and ($processed % 5000 -eq 0)) {
                    VLog "Comparaison : $processed / $($allFiles.Count) fichiers traités ($($swCompare.Elapsed.TotalSeconds.ToString('0.0'))s écoulées)"
                }
            }
            VLog "Comparaison complète effectuée en $($swCompare.Elapsed.TotalSeconds.ToString('0.0'))s"

            $sync.AuditSummary = @{
                Total      = $allFiles.Count
                OK         = $okCount
                Pending    = $pendingCount
                Duplicate  = $dupCount
                Missing    = $missingCount
                ReportPath = $auditFile
            }
            Log "=== AUDIT TERMINÉ === OK: $okCount | Pas encore déplacés: $pendingCount | Doublons: $dupCount | Introuvables: $missingCount"
            Log "Rapport détaillé : $auditFile"
        }
        elseif ($p.BatchNumber -lt 1 -or $p.BatchNumber -gt $totalBatches) {
            Log "ERREUR : le numéro de paquet demandé ($($p.BatchNumber)) est hors limites (1 à $totalBatches)."
            $sync.CurrentBatch = 0
        }
        elseif ($sync.Cancel) {
            Log "Arrêt demandé par l'utilisateur avant démarrage du paquet."
        }
        else {
            $batchNumber = $p.BatchNumber
            $b = $batchNumber - 1
            $batchFiles  = $batches[$b]
            $sync.CurrentBatch = $batchNumber
            Log "=== Paquet $batchNumber / $totalBatches ($($batchFiles.Count) fichiers) ==="

            if ($p.Simulate) {
                foreach ($f in $batchFiles) {
                    Add-Content -Path $p.LogFile -Value ('{0},{1},{2},"{3}",{4}' -f $batchNumber, $f, "SIMULATION", "Serait déplacé vers $destFolder", (Get-Date -Format 'o'))
                }
                $sync.GlobalSuccess += $batchFiles.Count
                if ($completedBatches -notcontains $batchNumber) { $completedBatches += $batchNumber }
                @{ CompletedBatches = ($completedBatches | Sort-Object -Unique); UpdatedAt = (Get-Date -Format 'o') } | ConvertTo-Json | Out-File -FilePath $p.ResumeFile -Encoding UTF8
                Log "  [Simulation] Paquet $batchNumber : $($batchFiles.Count) fichier(s) journalisé(s), aucune action réelle."
                if (@(1..$totalBatches | Where-Object { $completedBatches -notcontains $_ }).Count -eq 0) {
                    Log "=== TOUS LES PAQUETS SONT TRAITÉS (SIMULATION) ==="
                    $sync.AllDone = $true
                } else {
                    Log "=== Paquet $batchNumber terminé (simulation). Choisis le prochain paquet à traiter. ==="
                }
            }
            else {
                # --- Lancement asynchrone d'un job de déplacement PAR FICHIER (NoWait) ---
                $swJobs = [System.Diagnostics.Stopwatch]::StartNew()
                $jobs = @{}
                foreach ($f in $batchFiles) {
                    if ($sync.Cancel) { break }
                    $srcUrl = "$srcFolder/$f".Replace("//", "/")
                    try {
                        $job = Move-PnPFile -SourceUrl $srcUrl -TargetUrl $destFolder -Overwrite -Force -NoWait -Connection $conn -ErrorAction Stop
                        $jobs[$f] = $job
                        VLog "Job lancé : $srcUrl -> $destFolder"
                    } catch {
                        Add-Content -Path $p.LogFile -Value ('{0},{1},{2},"{3}",{4}' -f $batchNumber, $f, "JOB_START_FAILED", $_.Exception.Message, (Get-Date -Format 'o'))
                        $sync.GlobalFailed++
                        VLog "ÉCHEC lancement pour $f : $($_.Exception.Message)"
                    }
                }
                VLog "Tous les jobs du paquet lancés en $($swJobs.Elapsed.TotalSeconds.ToString('0.0'))s ($($jobs.Count) job(s))."

                # --- Attente de la fin de chaque job du paquet ---
                $swWait = [System.Diagnostics.Stopwatch]::StartNew()
                foreach ($f in $jobs.Keys) {
                    try {
                        Receive-PnPCopyMoveJobStatus -Job $jobs[$f] -Wait -Connection $conn -ErrorAction Stop | Out-Null
                        VLog "Job terminé : $f"
                    } catch {
                        Log "  Erreur statut job pour $f : $($_.Exception.Message)"
                    }
                }
                VLog "Attente des jobs terminée en $($swWait.Elapsed.TotalSeconds.ToString('0.0'))s."

                $swAudit = [System.Diagnostics.Stopwatch]::StartNew()
                $batchSuccess = 0; $batchFailed = 0
                try {
                    $destSet = Get-FolderFileSet -FolderRelativeUrl $destFolder -Connection $conn
                    $srcSet  = Get-FolderFileSet -FolderRelativeUrl $srcFolder -Connection $conn
                    VLog "Listing destination : $($destSet.Count) fichier(s) — Listing source : $($srcSet.Count) fichier(s) (en $($swAudit.Elapsed.TotalSeconds.ToString('0.0'))s)"

                    foreach ($f in $batchFiles) {
                        $inDest = $destSet.Contains($f)
                        $inSrc  = $srcSet.Contains($f)

                        if ($inDest -and -not $inSrc) {
                            Add-Content -Path $p.LogFile -Value ('{0},{1},{2},"{3}",{4}' -f $batchNumber, $f, "OK", "Déplacé et validé", (Get-Date -Format 'o'))
                            $batchSuccess++
                            VLog "OK        : $f"
                        } elseif ($inDest -and $inSrc) {
                            Add-Content -Path $p.LogFile -Value ('{0},{1},{2},"{3}",{4}' -f $batchNumber, $f, "DUPLICATE", "Présent aux deux emplacements", (Get-Date -Format 'o'))
                            $batchFailed++
                            VLog "DOUBLON   : $f (présent source ET destination)"
                        } else {
                            Add-Content -Path $p.LogFile -Value ('{0},{1},{2},"{3}",{4}' -f $batchNumber, $f, "MISSING", "Introuvable à destination", (Get-Date -Format 'o'))
                            $batchFailed++
                            VLog "MANQUANT  : $f (introuvable aux deux emplacements)"
                        }
                    }
                } catch {
                    Log "  ERREUR lors de l'audit du paquet : $($_.Exception.Message)"
                    $batchFailed = $batchFiles.Count
                }

                $sync.GlobalSuccess += $batchSuccess
                $sync.GlobalFailed  += $batchFailed
                Log "  Paquet $batchNumber : $batchSuccess OK / $batchFailed échec(s)"

                if ($batchFailed -eq 0) {
                    if ($completedBatches -notcontains $batchNumber) { $completedBatches += $batchNumber }
                    @{ CompletedBatches = ($completedBatches | Sort-Object -Unique); UpdatedAt = (Get-Date -Format 'o') } | ConvertTo-Json | Out-File -FilePath $p.ResumeFile -Encoding UTF8
                    if (@(1..$totalBatches | Where-Object { $completedBatches -notcontains $_ }).Count -eq 0) {
                        Log "=== TOUS LES PAQUETS SONT TRAITÉS ET VALIDÉS ==="
                        $sync.AllDone = $true
                    } else {
                        Log "=== Paquet $batchNumber terminé. Choisis le prochain paquet à traiter. ==="
                    }
                } else {
                    Log "  ATTENTION : paquet $batchNumber non marqué comme complété (échecs présents). Corrige puis relance ce même paquet."
                }
            }
        }

        Log "--- Bilan cumulé --- OK: $($sync.GlobalSuccess) / Échecs: $($sync.GlobalFailed)"
    } catch {
        Log "ERREUR FATALE : $($_.Exception.Message)"
    } finally {
        $sync.Done = $true
    }
}

# ============================================================================
# Timer UI (poll de l'état partagé)
# ============================================================================
$script:runspace = $null
$script:ps = $null
$script:handle = $null

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 400
$timer.Add_Tick({
    try {
        while ($sync.LogQueue.Count -gt 0) {
            $line = $sync.LogQueue.Dequeue()
            $txtLog.AppendText("$line`r`n")
            $txtLog.SelectionStart = $txtLog.Text.Length
            $txtLog.ScrollToCaret()
        }
        if ($sync.TotalBatches -gt 0) {
            $progressBar.Maximum = $sync.TotalBatches
            $progressBar.Value = [Math]::Min($sync.CurrentBatch, $sync.TotalBatches)
        }
        $lblStatus.Text = "Paquet $($sync.CurrentBatch) / $($sync.TotalBatches)  —  OK: $($sync.GlobalSuccess)   Échecs: $($sync.GlobalFailed)"

        if ($sync.Done -and -not $script:notifiedDone) {
            $script:notifiedDone = $true
            $timer.Stop()
            $btnStart.Enabled = $true
            $btnStop.Enabled = $false
            $btnAudit.Enabled = $true
            try { $script:ps.EndInvoke($script:handle) } catch {}
            try { $script:ps.Dispose() } catch {}
            try { $script:runspace.Close() } catch {}

            if ($sync.Params.Mode -eq 'Audit') {
                if ($sync.AuditSummary) {
                    $s = $sync.AuditSummary
                    [System.Windows.Forms.MessageBox]::Show(
                        "Audit terminé.`n`nTotal fichiers de la liste : $($s.Total)`nDéplacés et validés (OK) : $($s.OK)`nPas encore déplacés : $($s.Pending)`nDoublons (présents aux 2 endroits) : $($s.Duplicate)`nIntrouvables (aux 2 endroits) : $($s.Missing)`n`nRapport détaillé : $($s.ReportPath)",
                        "Audit terminé", 'OK', 'Information') | Out-Null
                } else {
                    [System.Windows.Forms.MessageBox]::Show("L'audit a échoué. Consulte le journal pour le détail.", "Audit incomplet", 'OK', 'Warning') | Out-Null
                }
            }
            else {
                $btnStart.Text = "Traiter le paquet sélectionné"
                if ($sync.Cancel) {
                    # rien de plus, le numéro reste tel quel pour reprendre ce paquet
                }
                elseif ($sync.AllDone) {
                    $btnStart.Enabled = $false
                    [System.Windows.Forms.MessageBox]::Show(
                        "Tous les paquets ont été traités.`nOK cumulé : $($sync.GlobalSuccess)`nÉchecs cumulés : $($sync.GlobalFailed)`nLog : $($sync.Params.LogFile)",
                        "Terminé", 'OK', 'Information') | Out-Null
                }
                Update-BatchSuggestion
            }
        }
    } catch {
        # Ne jamais laisser une exception du Timer remonter (évite la cascade de popups d'erreur Windows)
        $timer.Stop()
    }
})

# ============================================================================
# Démarrage (un clic = un seul paquet, celui sélectionné dans le champ ci-dessus)
# ============================================================================
$script:hasStartedOnce = $false

$btnStart.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text) -or
        [string]::IsNullOrWhiteSpace($txtSource.Text) -or
        [string]::IsNullOrWhiteSpace($txtDest.Text) -or
        [string]::IsNullOrWhiteSpace($txtFileList.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Merci de renseigner tous les champs (site, source, destination, fichier liste).", "Champs manquants", 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Test-Path $txtFileList.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Le fichier liste est introuvable :`n$($txtFileList.Text)", "Fichier introuvable", 'OK', 'Error') | Out-Null
        return
    }
    if (-not $script:CurrentPnPConnection) {
        [System.Windows.Forms.MessageBox]::Show("Aucune connexion active. Choisis 'Réutiliser' ou clique sur 'Se connecter'.", "Connexion requise", 'OK', 'Warning') | Out-Null
        return
    }

    $logFolder = $txtLogFolder.Text
    if (-not (Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }

    $hash       = Get-StateHash -Source $txtSource.Text -Destination $txtDest.Text
    $resumeFile = Join-Path $logFolder "Resume_$hash.json"
    $logFile    = Join-Path $logFolder "MoveLog_$hash.csv"

    if (-not $script:hasStartedOnce) {
        # Réinitialisation des compteurs UNIQUEMENT au tout premier paquet de la session,
        # pour que les cumuls OK/Échecs s'additionnent correctement entre paquets manuels.
        $sync.CurrentBatch  = 0
        $sync.TotalBatches  = 0
        $sync.GlobalSuccess = 0
        $sync.GlobalFailed  = 0
        $sync.AllDone       = $false
        $txtLog.Clear()
        $progressBar.Value = 0
        $script:hasStartedOnce = $true
    }
    $sync.Done   = $false
    $sync.Cancel = $false
    $sync.AuditSummary = $null
    $script:notifiedDone = $false
    $sync.Params = @{
        Mode         = 'Batch'
        SiteUrl      = $txtSite.Text
        SourceFolder = $txtSource.Text
        DestFolder   = $txtDest.Text
        FileListPath = $txtFileList.Text
        BatchSize    = [int]$numBatch.Value
        BatchNumber  = [int]$numBatchToRun.Value
        ResumeFile   = $resumeFile
        LogFile      = $logFile
        Connection   = $script:CurrentPnPConnection
        Simulate     = $chkSimulate.Checked
        Verbose      = $chkVerbose.Checked
    }

    $btnStart.Enabled = $false
    $btnStop.Enabled  = $true
    $btnAudit.Enabled = $false

    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions  = "ReuseThread"
    $script:runspace.Open()

    $script:ps = [powershell]::Create()
    $script:ps.Runspace = $script:runspace
    [void]$script:ps.AddScript($scriptBlock).AddArgument($sync)
    $script:handle = $script:ps.BeginInvoke()

    $timer.Start()
})

$btnAudit.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSite.Text) -or
        [string]::IsNullOrWhiteSpace($txtSource.Text) -or
        [string]::IsNullOrWhiteSpace($txtDest.Text) -or
        [string]::IsNullOrWhiteSpace($txtFileList.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Merci de renseigner tous les champs (site, source, destination, fichier liste).", "Champs manquants", 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Test-Path $txtFileList.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Le fichier liste est introuvable :`n$($txtFileList.Text)", "Fichier introuvable", 'OK', 'Error') | Out-Null
        return
    }
    if (-not $script:CurrentPnPConnection) {
        [System.Windows.Forms.MessageBox]::Show("Aucune connexion active. Choisis 'Réutiliser' ou clique sur 'Se connecter'.", "Connexion requise", 'OK', 'Warning') | Out-Null
        return
    }

    $logFolder = $txtLogFolder.Text
    if (-not (Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
    $hash      = Get-StateHash -Source $txtSource.Text -Destination $txtDest.Text
    $auditFile = Join-Path $logFolder "Audit_$hash`_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $logFile   = Join-Path $logFolder "MoveLog_$hash.csv"

    $txtLog.AppendText("--- Lancement d'un audit complet (indépendant des paquets) ---`r`n")
    $sync.Done   = $false
    $sync.Cancel = $false
    $sync.AuditSummary = $null
    $script:notifiedDone = $false
    $sync.Params = @{
        Mode         = 'Audit'
        SourceFolder = $txtSource.Text
        DestFolder   = $txtDest.Text
        FileListPath = $txtFileList.Text
        BatchSize    = [int]$numBatch.Value
        BatchNumber  = 1
        ResumeFile   = (Join-Path $logFolder "Resume_$hash.json")
        LogFile      = $logFile
        AuditFile    = $auditFile
        Connection   = $script:CurrentPnPConnection
        Simulate     = $false
        Verbose      = $chkVerbose.Checked
    }

    $btnStart.Enabled = $false
    $btnAudit.Enabled = $false
    $btnStop.Enabled  = $true

    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions  = "ReuseThread"
    $script:runspace.Open()

    $script:ps = [powershell]::Create()
    $script:ps.Runspace = $script:runspace
    [void]$script:ps.AddScript($scriptBlock).AddArgument($sync)
    $script:handle = $script:ps.BeginInvoke()

    $timer.Start()
})

$btnStop.Add_Click({
    $sync.Cancel = $true
    $txtLog.AppendText("Demande d'arrêt envoyée, patientez la fin du paquet en cours...`r`n")
    $btnStop.Enabled = $false
})

$btnResetResume.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSource.Text) -or [string]::IsNullOrWhiteSpace($txtDest.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Renseigne d'abord source et destination pour identifier la reprise à réinitialiser.", "Info", 'OK', 'Information') | Out-Null
        return
    }
    $hash = Get-StateHash -Source $txtSource.Text -Destination $txtDest.Text
    $resumeFile = Join-Path $txtLogFolder.Text "Resume_$hash.json"
    if (Test-Path $resumeFile) {
        $confirm = [System.Windows.Forms.MessageBox]::Show("Supprimer le fichier de reprise ?`n$resumeFile", "Confirmation", 'YesNo', 'Warning')
        if ($confirm -eq 'Yes') {
            Remove-Item $resumeFile -Force
            $script:hasStartedOnce = $false
            $btnStart.Text = "Traiter le paquet sélectionné"
            $btnStart.Enabled = $true
            Update-BatchSuggestion
            [System.Windows.Forms.MessageBox]::Show("Reprise réinitialisée.", "OK", 'OK', 'Information') | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Aucun fichier de reprise trouvé pour cette combinaison.", "Info", 'OK', 'Information') | Out-Null
    }
})

$form.Add_FormClosing({
    # Empêche toute exécution du Timer après destruction des contrôles :
    # c'était la cause du bug des milliers de fenêtres qui apparaissaient à la fermeture.
    $timer.Stop()
    $sync.Cancel = $true
    if ($script:ps) {
        try {
            if ($script:handle -and -not $script:handle.IsCompleted) {
                # Laisse une seconde au paquet en cours pour se terminer proprement
                $script:handle.AsyncWaitHandle.WaitOne(1000) | Out-Null
            }
            $script:ps.Stop()
        } catch {}
        try { $script:ps.Dispose() } catch {}
    }
    if ($script:runspace) {
        try { $script:runspace.Close() } catch {}
    }
})

[void]$form.ShowDialog()
