<#
=============================================================================================
Nom          : SPO Sharing Links Explorer - GUI
Description  : Interface graphique pour analyser (récursivement) et révoquer les liens de
               partage SharePoint Online sur une collection/chemin donné.
               Réutilise IMPÉRATIVEMENT une connexion PnP.PowerShell déjà établie dans la
               session courante (Connect-PnPOnline effectué AVANT de lancer ce script).

Prérequis    : - Module PnP.PowerShell installé
               - Etre déjà connecté : Connect-PnPOnline -Url <SiteUrl> ... (dans la même
                 session PowerShell, ou dot-sourcer ce script après connexion)
               - Lancer idéalement avec : powershell.exe -STA -File .\SPO_SharingLinks_GUI.ps1

Fonctionnement:
  1. Saisir le chemin de la collection à analyser (ex: /sites/HR/Documents partages ou un
     sous-dossier /sites/HR/Documents partages/ProjetX). Laisser vide = tout le site.
  2. Cliquer "Analyser" -> scan récursif en tâche de fond (runspace), la fenêtre reste fluide.
  3. Tous les sharing links trouvés s'affichent au fur et à mesure dans la grille.
  4. Cocher les lignes voulues puis cliquer "Revoke" pour révoquer les liens sélectionnés.
  5. "Exporter CSV" pour sauvegarder l'intégralité des résultats.
=============================================================================================
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------------------------------------
# Vérification de la connexion PnP existante (on ne se reconnecte JAMAIS ici)
# ------------------------------------------------------------------------------------------
try {
    $script:Conn = Get-PnPConnection -ErrorAction Stop
    $script:ConnectedWeb = Get-PnPWeb -Connection $script:Conn -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Aucune connexion PnP active détectée.`n`nConnectez-vous d'abord avec Connect-PnPOnline dans cette même session PowerShell, puis relancez ce script.",
        "Connexion PnP requise",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

# ------------------------------------------------------------------------------------------
# Etat partagé thread-safe entre le runspace de scan et le thread UI
# ------------------------------------------------------------------------------------------
$Global:SyncHash = [hashtable]::Synchronized(@{
    Queue         = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    Status        = "Prêt"
    ItemsScanned  = 0
    LinksFound    = 0
    Done          = $true
    Cancelled     = $false
    Error         = $null
})

$script:AllResults  = New-Object System.Collections.ArrayList
$script:Runspace     = $null
$script:PS           = $null
$script:AsyncResult  = $null

# ------------------------------------------------------------------------------------------
# ScriptBlock exécuté dans le runspace de fond : scan récursif + collecte des sharing links
# ------------------------------------------------------------------------------------------
$ScanScriptBlock = {
    param($SyncHash, $Conn, $ScopePathRaw)

    Import-Module PnP.PowerShell -DisableNameChecking -ErrorAction Stop

    try {
        $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages",
                            "Preservation Hold Library","Pages","Images",
                            "Site Collection Documents","Site Collection Images")

        $NormalizedScope = $ScopePathRaw.Trim().TrimEnd('/')

        $Lists = Get-PnPList -Connection $Conn | Where-Object {
            $_.Hidden -eq $false -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"
        }

        foreach ($List in $Lists) {
            if ($SyncHash.Cancelled) { break }

            $RootFolder = Get-PnPProperty -Connection $Conn -ClientObject $List -Property RootFolder
            $ListRoot   = $RootFolder.ServerRelativeUrl

            # Filtre de portée : ignore les bibliothèques totalement hors du chemin demandé
            if ($NormalizedScope -ne "" -and
                -not ($ListRoot -like "$NormalizedScope*") -and
                -not ($NormalizedScope -like "$ListRoot*")) {
                continue
            }

            $Items = Get-PnPListItem -Connection $Conn -List $List -PageSize 2000

            foreach ($Item in $Items) {
                if ($SyncHash.Cancelled) { break }

                $FileRef = $Item.FieldValues.FileRef
                if ($NormalizedScope -ne "" -and $FileRef -notlike "$NormalizedScope*") { continue }

                $ObjectType = $Item.FileSystemObjectType
                $FileName   = $Item.FieldValues.FileLeafRef
                $SyncHash.ItemsScanned++
                $SyncHash.Status = "Analyse: $FileName"

                try {
                    $HasUnique = Get-PnPProperty -Connection $Conn -ClientObject $Item -Property HasUniqueRoleAssignments
                } catch { continue }
                if (-not $HasUnique) { continue }

                try {
                    if ($ObjectType -eq "File") {
                        $Links = Get-PnPFileSharingLink -Connection $Conn -Identity $FileRef
                    } elseif ($ObjectType -eq "Folder") {
                        $Links = Get-PnPFolderSharingLink -Connection $Conn -Folder $FileRef
                    } else { continue }
                } catch { continue }

                foreach ($L in $Links) {
                    $Link = $L.Link
                    $ExpirationDate = $L.ExpirationDateTime
                    $CurrentDateTime = (Get-Date).Date

                    if ($ExpirationDate) {
                        $ExpiryDate = ([DateTime]$ExpirationDate).ToLocalTime()
                        $ExpiryDays = (New-TimeSpan -Start $CurrentDateTime -End $ExpiryDate).Days
                        if ($ExpiryDate -lt $CurrentDateTime) {
                            $LinkStatus     = "Expiré"
                            $FriendlyExpiry = "Expiré depuis $([math]::Abs($ExpiryDays)) jours"
                        } else {
                            $LinkStatus     = "Actif"
                            $FriendlyExpiry = "Expire dans $ExpiryDays jours"
                        }
                    } else {
                        $LinkStatus     = "Actif"
                        $ExpiryDate     = $null
                        $FriendlyExpiry = "N'expire jamais"
                    }

                    $Obj = [PSCustomObject]@{
                        Library           = $List.Title
                        ObjectType        = $ObjectType
                        Name              = $FileName
                        FileUrl           = $FileRef
                        LinkId            = $L.Id
                        Scope             = $Link.Scope
                        AccessType        = $Link.Type
                        Roles             = ($L.Roles -join ",")
                        Users             = ($L.GrantedToIdentitiesV2.User.Email -join ",")
                        LinkStatus        = $LinkStatus
                        ExpiryDate        = $ExpiryDate
                        FriendlyExpiry    = $FriendlyExpiry
                        PasswordProtected = $L.HasPassword
                        BlockDownload     = $Link.PreventsDownload
                        SharedLink        = $Link.WebUrl
                        RevokeStatus      = ""
                    }
                    $SyncHash.Queue.Enqueue($Obj)
                    $SyncHash.LinksFound++
                }
            }
        }
    }
    catch {
        $SyncHash.Error = $_.Exception.Message
    }
    finally {
        $SyncHash.Done = $true
        if ($SyncHash.Cancelled) { $SyncHash.Status = "Analyse annulée" }
        else { $SyncHash.Status = "Terminé" }
    }
}

# ------------------------------------------------------------------------------------------
# Construction du formulaire
# ------------------------------------------------------------------------------------------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "SPO Sharing Links Explorer - $($script:ConnectedWeb.Title)"
$Form.Size = New-Object System.Drawing.Size(1300, 760)
$Form.StartPosition = "CenterScreen"
$Form.MinimumSize = New-Object System.Drawing.Size(1000, 600)

# -- Panneau haut : saisie du chemin + actions --
$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = "Chemin de la collection à analyser (vide = tout le site) :"
$lblPath.Location = New-Object System.Drawing.Point(10, 15)
$lblPath.Size = New-Object System.Drawing.Size(400, 20)
$Form.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(10, 38)
$txtPath.Size = New-Object System.Drawing.Size(650, 24)
$txtPath.Text = $script:ConnectedWeb.ServerRelativeUrl
$Form.Controls.Add($txtPath)

$btnAnalyze = New-Object System.Windows.Forms.Button
$btnAnalyze.Text = "Analyser"
$btnAnalyze.Location = New-Object System.Drawing.Point(670, 36)
$btnAnalyze.Size = New-Object System.Drawing.Size(100, 28)
$btnAnalyze.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnAnalyze.ForeColor = [System.Drawing.Color]::White
$Form.Controls.Add($btnAnalyze)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Annuler"
$btnCancel.Location = New-Object System.Drawing.Point(778, 36)
$btnCancel.Size = New-Object System.Drawing.Size(90, 28)
$btnCancel.Enabled = $false
$Form.Controls.Add($btnCancel)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Tout cocher"
$btnSelectAll.Location = New-Object System.Drawing.Point(880, 36)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 28)
$Form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Tout décocher"
$btnDeselectAll.Location = New-Object System.Drawing.Point(986, 36)
$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 28)
$Form.Controls.Add($btnDeselectAll)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exporter CSV"
$btnExport.Location = New-Object System.Drawing.Point(1092, 36)
$btnExport.Size = New-Object System.Drawing.Size(110, 28)
$Form.Controls.Add($btnExport)

$btnRevoke = New-Object System.Windows.Forms.Button
$btnRevoke.Text = "Revoke"
$btnRevoke.Location = New-Object System.Drawing.Point(10, 72)
$btnRevoke.Size = New-Object System.Drawing.Size(150, 32)
$btnRevoke.BackColor = [System.Drawing.Color]::FromArgb(196,43,28)
$btnRevoke.ForeColor = [System.Drawing.Color]::White
$btnRevoke.Font = New-Object System.Drawing.Font($btnRevoke.Font, [System.Drawing.FontStyle]::Bold)
$Form.Controls.Add($btnRevoke)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Prêt."
$lblStatus.Location = New-Object System.Drawing.Point(170, 78)
$lblStatus.Size = New-Object System.Drawing.Size(1100, 20)
$Form.Controls.Add($lblStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 112)
$progressBar.Size = New-Object System.Drawing.Size(1260, 18)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 0
$Form.Controls.Add($progressBar)

# -- Grille de résultats --
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 140)
$grid.Size = New-Object System.Drawing.Size(1265, 560)
$grid.Anchor = "Top,Bottom,Left,Right"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $false
$grid.SelectionMode = "FullRowSelect"
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = "None"
$grid.AllowUserToResizeColumns = $true

$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.Name = "Selection"; $colCheck.HeaderText = ""; $colCheck.Width = 30
$grid.Columns.Add($colCheck) | Out-Null

function New-TextCol($name, $header, $width, [switch]$hidden) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $name; $c.HeaderText = $header; $c.Width = $width; $c.ReadOnly = $true
    if ($hidden) { $c.Visible = $false }
    return $c
}

$grid.Columns.Add((New-TextCol "Library" "Bibliothèque" 130)) | Out-Null
$grid.Columns.Add((New-TextCol "ObjectType" "Type" 60)) | Out-Null
$grid.Columns.Add((New-TextCol "Name" "Nom" 180)) | Out-Null
$grid.Columns.Add((New-TextCol "Scope" "Type de lien" 100)) | Out-Null
$grid.Columns.Add((New-TextCol "AccessType" "Accès" 70)) | Out-Null
$grid.Columns.Add((New-TextCol "Users" "Utilisateurs" 200)) | Out-Null
$grid.Columns.Add((New-TextCol "LinkStatus" "Statut" 70)) | Out-Null
$grid.Columns.Add((New-TextCol "FriendlyExpiry" "Expiration" 150)) | Out-Null
$grid.Columns.Add((New-TextCol "PasswordProtected" "Mot de passe" 90)) | Out-Null
$grid.Columns.Add((New-TextCol "BlockDownload" "Bloque téléch." 90)) | Out-Null
$grid.Columns.Add((New-TextCol "SharedLink" "Lien" 220)) | Out-Null
$grid.Columns.Add((New-TextCol "RevokeStatus" "Statut Revoke" 110)) | Out-Null
$grid.Columns.Add((New-TextCol "FileUrl" "FileUrl" 0 -hidden)) | Out-Null
$grid.Columns.Add((New-TextCol "LinkId" "LinkId" 0 -hidden)) | Out-Null

$Form.Controls.Add($grid)

# ------------------------------------------------------------------------------------------
# Timer UI : draine la queue alimentée par le runspace, met à jour la grille et le statut
# ------------------------------------------------------------------------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250

$timer.Add_Tick({
    # Flush des nouveaux éléments trouvés
    while ($Global:SyncHash.Queue.Count -gt 0) {
        $item = $Global:SyncHash.Queue.Dequeue()
        [void]$script:AllResults.Add($item)

        $rowIndex = $grid.Rows.Add()
        $row = $grid.Rows[$rowIndex]
        $row.Cells["Library"].Value           = $item.Library
        $row.Cells["ObjectType"].Value         = $item.ObjectType
        $row.Cells["Name"].Value               = $item.Name
        $row.Cells["Scope"].Value              = $item.Scope
        $row.Cells["AccessType"].Value          = $item.AccessType
        $row.Cells["Users"].Value               = $item.Users
        $row.Cells["LinkStatus"].Value           = $item.LinkStatus
        $row.Cells["FriendlyExpiry"].Value      = $item.FriendlyExpiry
        $row.Cells["PasswordProtected"].Value   = $item.PasswordProtected
        $row.Cells["BlockDownload"].Value       = $item.BlockDownload
        $row.Cells["SharedLink"].Value           = $item.SharedLink
        $row.Cells["RevokeStatus"].Value         = ""
        $row.Cells["FileUrl"].Value              = $item.FileUrl
        $row.Cells["LinkId"].Value               = $item.LinkId
    }

    $lblStatus.Text = "$($Global:SyncHash.Status)  |  Eléments analysés: $($Global:SyncHash.ItemsScanned)  |  Liens trouvés: $($Global:SyncHash.LinksFound)"

    if ($Global:SyncHash.Done) {
        $timer.Stop()
        $progressBar.Style = "Blocks"
        $progressBar.MarqueeAnimationSpeed = 0
        $btnAnalyze.Enabled = $true
        $btnCancel.Enabled = $false
        $txtPath.Enabled = $true

        if ($script:PS) {
            try { $script:PS.EndInvoke($script:AsyncResult) | Out-Null } catch {}
            $script:PS.Dispose()
            $script:Runspace.Close()
            $script:Runspace.Dispose()
            $script:PS = $null
        }

        if ($Global:SyncHash.Error) {
            [System.Windows.Forms.MessageBox]::Show("Erreur durant l'analyse :`n$($Global:SyncHash.Error)", "Erreur", "OK", "Error") | Out-Null
        }
        elseif (-not $Global:SyncHash.Cancelled) {
            [System.Windows.Forms.MessageBox]::Show("Analyse terminée.`n$($Global:SyncHash.ItemsScanned) éléments analysés, $($Global:SyncHash.LinksFound) liens de partage trouvés.", "Analyse terminée", "OK", "Information") | Out-Null
        }
    }
})

# ------------------------------------------------------------------------------------------
# Lancement de l'analyse
# ------------------------------------------------------------------------------------------
$btnAnalyze.Add_Click({
    $grid.Rows.Clear()
    $script:AllResults.Clear()

    $Global:SyncHash.Queue.Clear()
    $Global:SyncHash.Status = "Démarrage..."
    $Global:SyncHash.ItemsScanned = 0
    $Global:SyncHash.LinksFound = 0
    $Global:SyncHash.Done = $false
    $Global:SyncHash.Cancelled = $false
    $Global:SyncHash.Error = $null

    $btnAnalyze.Enabled = $false
    $btnCancel.Enabled = $true
    $txtPath.Enabled = $false
    $progressBar.Style = "Marquee"
    $progressBar.MarqueeAnimationSpeed = 30

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $script:Runspace = [runspacefactory]::CreateRunspace($iss)
    $script:Runspace.ApartmentState = "STA"
    $script:Runspace.ThreadOptions = "ReuseThread"
    $script:Runspace.Open()

    $script:PS = [powershell]::Create()
    $script:PS.Runspace = $script:Runspace
    $script:PS.AddScript($ScanScriptBlock) | Out-Null
    $script:PS.AddArgument($Global:SyncHash) | Out-Null
    $script:PS.AddArgument($script:Conn) | Out-Null
    $script:PS.AddArgument($txtPath.Text) | Out-Null

    $script:AsyncResult = $script:PS.BeginInvoke()
    $timer.Start()
})

$btnCancel.Add_Click({
    $Global:SyncHash.Cancelled = $true
    $lblStatus.Text = "Annulation en cours..."
})

$btnSelectAll.Add_Click({
    foreach ($row in $grid.Rows) { $row.Cells["Selection"].Value = $true }
})

$btnDeselectAll.Add_Click({
    foreach ($row in $grid.Rows) { $row.Cells["Selection"].Value = $false }
})

# ------------------------------------------------------------------------------------------
# Revoke des liens sélectionnés
# ------------------------------------------------------------------------------------------
$btnRevoke.Add_Click({
    $selectedRows = @($grid.Rows | Where-Object { $_.Cells["Selection"].Value -eq $true })

    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun lien sélectionné.", "Revoke", "OK", "Warning") | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Vous êtes sur le point de révoquer $($selectedRows.Count) lien(s) de partage.`nCette action est irréversible. Continuer ?",
        "Confirmation Revoke",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $success = 0
    $failed  = 0

    foreach ($row in $selectedRows) {
        $fileUrl    = $row.Cells["FileUrl"].Value
        $linkId     = $row.Cells["LinkId"].Value
        $objectType = $row.Cells["ObjectType"].Value

        try {
            if ($objectType -eq "File") {
                Remove-PnPFileSharingLink -Connection $script:Conn -FileUrl $fileUrl -Identity $linkId -Force -ErrorAction Stop
            }
            elseif ($objectType -eq "Folder") {
                Remove-PnPFolderSharingLink -Connection $script:Conn -Folder $fileUrl -Identity $linkId -Force -ErrorAction Stop
            }
            $row.Cells["RevokeStatus"].Value = "Révoqué"
            $row.Cells["RevokeStatus"].Style.ForeColor = [System.Drawing.Color]::Green
            $row.Cells["Selection"].Value = $false
            $success++
        }
        catch {
            $row.Cells["RevokeStatus"].Value = "Erreur: $($_.Exception.Message)"
            $row.Cells["RevokeStatus"].Style.ForeColor = [System.Drawing.Color]::Red
            $failed++
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Révocation terminée.`nSuccès: $success`nEchecs: $failed", "Revoke", "OK", "Information") | Out-Null
})

# ------------------------------------------------------------------------------------------
# Export CSV
# ------------------------------------------------------------------------------------------
$btnExport.Add_Click({
    if ($script:AllResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucune donnée à exporter. Lancez une analyse d'abord.", "Export", "OK", "Warning") | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Fichier CSV (*.csv)|*.csv"
    $sfd.FileName = "SPO_SharingLinks_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').csv"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:AllResults | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Force
        [System.Windows.Forms.MessageBox]::Show("Export terminé :`n$($sfd.FileName)", "Export", "OK", "Information") | Out-Null
    }
})

$Form.Add_FormClosing({
    $Global:SyncHash.Cancelled = $true
    if ($script:PS) {
        try { $script:PS.Stop() } catch {}
        try { $script:Runspace.Close() } catch {}
    }
})

[void]$Form.ShowDialog()
