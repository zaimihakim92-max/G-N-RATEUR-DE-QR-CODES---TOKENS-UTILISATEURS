#Requires -Version 5.1
<#
.SYNOPSIS
    Ajout / retrait en masse d'appareils dans une collection SCCM (regles d'appartenance DIRECTE).

.DESCRIPTION
    Interface graphique WinForms. A lancer depuis une console PowerShell ou la connexion SCCM
    est deja etablie (module ConfigurationManager importe + lecteur PSDrive du site present).

    Concu pour ne laisser AUCUNE place a une action indesirable :
      - le contexte SCCM est detecte, jamais reconnecte a l'aveugle ;
      - la collection cible doit etre VALIDEE avant toute action ;
      - toute modification de la cible (collection ou liste d'appareils) re-arme les verrous ;
      - une RECHERCHE prealable est obligatoire avant toute execution ;
      - les noms sont sanitises : aucun caractere generique ( * ? % [ ] ) n'est accepte ;
      - resolution stricte : un nom ambigu (plusieurs ResourceID) est IGNORE, jamais devine ;
      - confirmation detaillee avant l'ajout ; confirmation par SAISIE DE L'ID avant le retrait ;
      - les collections integrees (SMS*) sont refusees en ecriture ;
      - journalisation horodatee (GUI + console + fichier) et export CSV.

    Le script n'agit QUE sur les regles d'appartenance DIRECTE. Il ne touche jamais aux
    regles de requete d'une collection.

.NOTES
    Lancer de preference via :  powershell.exe -STA -File .\SCCM_CollectionBulk_GUI.ps1
    (ou depuis la console ConfigMgr, apres avoir importe le module et le lecteur de site).
#>

# ============================================================================
#  0. Pre-requis d'assembly / apartment
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "Le thread courant n'est pas en mode STA. L'interface peut mal se comporter."
    Write-Warning "Relancez avec :  powershell.exe -STA -File .\SCCM_CollectionBulk_GUI.ps1"
}

# ============================================================================
#  1. Etat global
# ============================================================================
$script:CMDrive             = $null      # nom du lecteur PSDrive du site (= code de site)
$script:ValidatedCollection = $null      # objet collection valide (Id / Name / MemberCount / ...)
$script:SimData             = $null      # resultats de la derniere simulation (List)
$script:SimulationDone      = $false     # simulation faite pour la cible courante ?
$script:IsProcessing        = $false     # une operation d'ecriture est-elle en cours ?
$script:CancelRequested     = $false     # demande d'arret par l'utilisateur
$script:LogFile             = $null      # chemin du fichier de log courant

$script:WILDCARD_CHARS      = @('*','?','%','[',']')

# ============================================================================
#  2. Journalisation (GUI + console + fichier)
# ============================================================================
function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO','OK','WARN','ERR','RECH')] [string] $Level = 'INFO'
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Message"

    if ($script:txtLog) {
        $script:txtLog.AppendText($line + "`r`n")
        $script:txtLog.SelectionStart = $script:txtLog.Text.Length
        $script:txtLog.ScrollToCaret()
    }

    $color = switch ($Level) { 'OK'{'Green'} 'WARN'{'Yellow'} 'ERR'{'Red'} 'RECH'{'Cyan'} default{'Gray'} }
    Write-Host $line -ForegroundColor $color

    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
    }
}

function Initialize-LogFile {
    if ($script:LogFile) { return }
    $folder = $script:txtLogFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        $folder = Join-Path ([Environment]::GetFolderPath('Desktop')) 'SCCM_BulkCollection_Logs'
    }
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $script:LogFile = Join-Path $folder ("SCCM_BulkCollection_{0}.log" -f $stamp)
    Write-Log "Journal : $($script:LogFile)" 'INFO'
    Write-Log "Operateur : $env:USERNAME  /  Machine : $env:COMPUTERNAME" 'INFO'
}

# ============================================================================
#  3. Contexte SCCM
# ============================================================================
function Initialize-CMContext {
    $drive = Get-PSDrive -PSProvider CMSite -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $drive) { return $false }
    $script:CMDrive = $drive.Name
    return $true
}

# Execute un scriptblock dans le contexte du lecteur CM, puis restaure l'emplacement.
function Invoke-InCMDrive {
    param([Parameter(Mandatory)][scriptblock] $Script)
    $prev = Get-Location
    try {
        Set-Location "$($script:CMDrive):\" -ErrorAction Stop
        & $Script
    }
    finally {
        Set-Location $prev -ErrorAction SilentlyContinue
    }
}

# ============================================================================
#  4. Validation / sanitisation des entrees
# ============================================================================
function Test-NameSafe {
    # Autorise lettres, chiffres, tiret, point (FQDN), underscore. Rejette tout caractere generique.
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    foreach ($c in $script:WILDCARD_CHARS) { if ($Name.Contains($c)) { return $false } }
    return ($Name -match '^[A-Za-z0-9][A-Za-z0-9._-]{0,62}$')
}

function Get-InputDeviceNames {
    $raw = $script:txtDevices.Text
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $parts = $raw -split "[\r\n,; `t]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

    $seen   = @{}
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($p in $parts) {
        $key = $p.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $result.Add($p) }
    }
    return $result
}

# ============================================================================
#  5. Gestion de l'etat des controles (verrous)
# ============================================================================
function Reset-Simulation {
    $script:SimData        = $null
    $script:SimulationDone = $false
    if ($script:grid) { $script:grid.Rows.Clear() }
    Update-ButtonsState
}

function Invalidate-Collection {
    $script:ValidatedCollection = $null
    if ($script:lblCollInfo) {
        $script:lblCollInfo.Text      = "Collection non validee."
        $script:lblCollInfo.ForeColor = [System.Drawing.Color]::DimGray
    }
    Reset-Simulation
}

function Update-ButtonsState {
    $hasColl    = ($null -ne $script:ValidatedCollection)
    $hasDevices = ((Get-InputDeviceNames).Count -gt 0)
    $busy       = $script:IsProcessing

    if ($script:btnValidate) { $script:btnValidate.Enabled = (-not $busy) }
    if ($script:btnSimulate) { $script:btnSimulate.Enabled = ($hasColl -and $hasDevices -and -not $busy) }
    if ($script:btnAdd)      { $script:btnAdd.Enabled      = ($hasColl -and $script:SimulationDone -and -not $busy) }
    if ($script:btnRemove)   { $script:btnRemove.Enabled   = ($hasColl -and $script:SimulationDone -and -not $busy) }
    if ($script:btnListMembers) { $script:btnListMembers.Enabled = ($hasColl -and -not $busy) }
    if ($script:btnStop)     { $script:btnStop.Enabled     = $busy }
    if ($script:txtCollection){ $script:txtCollection.Enabled = (-not $busy) }
    if ($script:txtDevices)  { $script:txtDevices.Enabled  = (-not $busy) }
    if ($script:btnLoadTxt)  { $script:btnLoadTxt.Enabled  = (-not $busy) }
    if ($script:btnClear)    { $script:btnClear.Enabled    = (-not $busy) }
}

function Set-ProcessingState {
    param([bool] $On)
    $script:IsProcessing = $On
    $script:MainForm.Cursor = if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    Update-ButtonsState
}

# ============================================================================
#  6. Validation de la collection
# ============================================================================
function Validate-Collection {
    $target = $script:txtCollection.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($target)) {
        [System.Windows.Forms.MessageBox]::Show("Saisissez un nom ou un ID de collection.","Validation",'OK','Warning') | Out-Null
        return
    }
    foreach ($c in $script:WILDCARD_CHARS) {
        if ($target.Contains($c)) {
            [System.Windows.Forms.MessageBox]::Show("Les caracteres generiques ( * ? % [ ] ) sont interdits dans la cible.","Securite",'OK','Error') | Out-Null
            return
        }
    }

    Invalidate-Collection
    Write-Log "Validation de la collection '$target'..." 'INFO'

    $col = $null
    try {
        $col = Invoke-InCMDrive {
            # 1) Tentative par ID exact (format 8 caracteres alphanumeriques)
            if ($target -match '^[A-Za-z0-9]{8}$') {
                $c = Get-CMDeviceCollection -CollectionId $target -ErrorAction SilentlyContinue
                if ($c) { return $c }
            }
            # 2) Tentative par nom EXACT (filtre cote client pour eviter tout comportement generique)
            @(Get-CMDeviceCollection -Name $target -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -eq $target })
        }
    }
    catch {
        Write-Log "Erreur lors de la resolution : $($_.Exception.Message)" 'ERR'
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la resolution de la collection :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        return
    }

    $matches = @($col)
    if ($matches.Count -eq 0) {
        Write-Log "Aucune collection ne correspond a '$target'." 'ERR'
        $script:lblCollInfo.Text      = "Aucune collection trouvee."
        $script:lblCollInfo.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }
    if ($matches.Count -gt 1) {
        $ids = ($matches | ForEach-Object { $_.CollectionID }) -join ', '
        Write-Log "Nom ambigu ($($matches.Count) collections) : $ids. Utilisez l'ID exact." 'ERR'
        $script:lblCollInfo.Text      = "Nom ambigu ($($matches.Count) collections). Utilisez l'ID exact : $ids"
        $script:lblCollInfo.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }

    $c  = $matches[0]
    $id = $c.CollectionID

    # Verrou : refus des collections integrees (non modifiables par regle directe).
    if ($id -like 'SMS*') {
        Write-Log "Collection integree ($id) : ecriture refusee." 'ERR'
        $script:lblCollInfo.Text      = "Collection integree ($id) : ajout/retrait de regle directe interdit."
        $script:lblCollInfo.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }

    # Informations complementaires (nombre de regles de requete, pour avertissement).
    $queryCount = 0
    try {
        $queryCount = Invoke-InCMDrive {
            @(Get-CMDeviceCollectionQueryMembershipRule -CollectionId $id -ErrorAction SilentlyContinue).Count
        }
    } catch { }

    $script:ValidatedCollection = [pscustomobject]@{
        Id          = $id
        Name        = $c.Name
        MemberCount = $c.MemberCount
        QueryRules  = $queryCount
    }

    $info = "VALIDEE  |  Nom : $($c.Name)  |  ID : $id  |  Membres actuels : $($c.MemberCount)  |  Regles de requete : $queryCount"
    $script:lblCollInfo.Text      = $info
    $script:lblCollInfo.ForeColor = [System.Drawing.Color]::DarkGreen
    Write-Log $info 'OK'
    if ($queryCount -gt 0) {
        Write-Log "Attention : cette collection contient des regles de requete. L'ajout direct reste valable mais l'appartenance peut aussi provenir d'une requete." 'WARN'
    }

    $script:SimulationDone = $false
    Update-ButtonsState
}

# ============================================================================
#  7. Recherche (aucune modification)
# ============================================================================
function Invoke-Simulation {
    if (-not $script:ValidatedCollection) { return }
    $names = Get-InputDeviceNames
    if ($names.Count -eq 0) { return }

    if ($names.Count -gt 5000) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Vous avez saisi $($names.Count) entrees. Continuer la recherche ?",
            "Volume important",'YesNo','Warning')
        if ($r -ne 'Yes') { return }
    }

    Initialize-LogFile
    Set-ProcessingState $true
    $script:CancelRequested = $false
    $script:grid.Rows.Clear()
    Write-Log "----- RECHERCHE (aucune modification) -----" 'RECH'
    Write-Log "Cible : $($script:ValidatedCollection.Name) ($($script:ValidatedCollection.Id))  |  $($names.Count) nom(s) unique(s)." 'RECH'

    $data = New-Object System.Collections.Generic.List[object]
    $cid  = $script:ValidatedCollection.Id

    $prev = Get-Location
    try {
        Set-Location "$($script:CMDrive):\" -ErrorAction Stop

        # Ensemble frais des ResourceID deja membres DIRECTS.
        $directSet = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($r in @(Get-CMDeviceCollectionDirectMembershipRule -CollectionId $cid -ErrorAction SilentlyContinue)) {
            [void]$directSet.Add([int]$r.ResourceID)
        }

        # Index nom -> liste de ResourceID, construit UNE SEULE FOIS via Get-CMResource.
        # SMS_R_System n'est pas restreint par le RBAC, contrairement a la vue par defaut de
        # Get-CMDevice : cela fiabilise la resolution avec des droits partiels. Une seule requete
        # serveur (moins de throttling), filtrage exact en memoire (aucun comportement generique).
        Write-Log "Construction de l'index des ressources (Get-CMResource -ResourceType System)..." 'INFO'
        $devIndex = @{}
        $resCount = 0
        foreach ($res in (Get-CMResource -ResourceType System -Fast -ErrorAction SilentlyContinue | Where-Object { $_.Name })) {
            $resCount++
            $k = ([string]$res.Name).ToLowerInvariant()
            if (-not $devIndex.ContainsKey($k)) { $devIndex[$k] = New-Object System.Collections.Generic.List[object] }
            [void]$devIndex[$k].Add([pscustomobject]@{ Rid = [int]$res.ResourceId; LastLogon = [string]$res.LastLogonUserName })
        }
        Write-Log "Index pret : $resCount ressource(s), $($devIndex.Keys.Count) nom(s) distinct(s)." 'OK'

        $script:progress.Minimum = 0
        $script:progress.Maximum = $names.Count
        $script:progress.Value   = 0
        $i = 0

        foreach ($name in $names) {
            if ($script:CancelRequested) { Write-Log "Recherche interrompue par l'utilisateur." 'WARN'; break }
            $i++
            $script:progress.Value = $i

            if (-not (Test-NameSafe $name)) {
                $data.Add([pscustomobject]@{ Name=$name; Category='INVALIDE'; ResourceID=$null; LastLogon=''; Detail='Nom rejete (format ou caractere interdit)' })
                Write-Log "INVALIDE : '$name' (format ou caractere interdit)" 'WARN'
                Add-GridRow $name 'INVALIDE' '' '' 'Nom rejete'
                [System.Windows.Forms.Application]::DoEvents(); continue
            }

            $k    = $name.ToLowerInvariant()
            $hits = if ($devIndex.ContainsKey($k)) { $devIndex[$k] } else { @() }
            if ($hits.Count -eq 0) {
                $data.Add([pscustomobject]@{ Name=$name; Category='NON_TROUVE'; ResourceID=$null; LastLogon=''; Detail='Aucun enregistrement' })
                Write-Log "NON TROUVE : '$name'" 'WARN'
                Add-GridRow $name 'NON_TROUVE' '' '' 'Aucun enregistrement'
            }
            elseif ($hits.Count -gt 1) {
                $ridList = ($hits | ForEach-Object { $_.Rid }) -join ','
                $data.Add([pscustomobject]@{ Name=$name; Category='AMBIGU'; ResourceID=$null; LastLogon=''; Detail="ResourceID multiples : $ridList" })
                Write-Log "AMBIGU : '$name' -> $($hits.Count) enregistrements ($ridList). IGNORE." 'WARN'
                Add-GridRow $name 'AMBIGU' '' '' "Multiples : $ridList"
            }
            else {
                $rid  = [int]$hits[0].Rid
                $llu  = [string]$hits[0].LastLogon
                if ($directSet.Contains($rid)) {
                    $data.Add([pscustomobject]@{ Name=$name; Category='MEMBRE_DIRECT'; ResourceID=$rid; LastLogon=$llu; Detail='Deja membre direct' })
                    Add-GridRow $name 'MEMBRE_DIRECT' $rid $llu 'Deja membre direct'
                } else {
                    $data.Add([pscustomobject]@{ Name=$name; Category='HORS_COLLECTION'; ResourceID=$rid; LastLogon=$llu; Detail='Non membre direct' })
                    Add-GridRow $name 'HORS_COLLECTION' $rid $llu 'Non membre direct'
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        Write-Log "Erreur de recherche : $($_.Exception.Message)" 'ERR'
    }
    finally {
        Set-Location $prev -ErrorAction SilentlyContinue
        $script:progress.Value = 0
    }

    $script:SimData = $data
    $addC = @($data | Where-Object Category -eq 'HORS_COLLECTION').Count
    $memC = @($data | Where-Object Category -eq 'MEMBRE_DIRECT').Count
    $nfC  = @($data | Where-Object Category -eq 'NON_TROUVE').Count
    $amC  = @($data | Where-Object Category -eq 'AMBIGU').Count
    $ivC  = @($data | Where-Object Category -eq 'INVALIDE').Count

    Write-Log "Bilan recherche -> a ajouter : $addC | deja membres : $memC | non trouves : $nfC | ambigus : $amC | invalides : $ivC" 'RECH'
    $script:SimulationDone = $true
    Set-ProcessingState $false
}

function Add-GridRow {
    param([string]$Name,[string]$Category,[string]$ResourceID,[string]$LastLogon,[string]$Detail)
    $idx = $script:grid.Rows.Add($Name, $Category, $ResourceID, $LastLogon, $Detail)
    $row = $script:grid.Rows[$idx]
    switch ($Category) {
        'HORS_COLLECTION' { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220,245,220) }
        'MEMBRE_DIRECT'   { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220,235,250) }
        'NON_TROUVE'      { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,225,225) }
        'AMBIGU'          { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(252,235,200) }
        'INVALIDE'        { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240,220,240) }
    }
}

# ============================================================================
#  8. Confirmation par saisie (pour le retrait)
# ============================================================================
function Confirm-ByTyping {
    param([string]$Expected,[string]$Prompt)

    $margin = 18
    $width  = 520

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Confirmation requise"
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.AutoScaleMode = 'Font'
    $f.Font = New-Object System.Drawing.Font('Segoe UI',9)

    # Message : label auto-dimensionne avec retour a la ligne (robuste au scaling DPI).
    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $true
    $l.MaximumSize = New-Object System.Drawing.Size($width, 0)
    $l.Location = New-Object System.Drawing.Point($margin, $margin)
    $l.Text = $Prompt

    $lHint = New-Object System.Windows.Forms.Label
    $lHint.AutoSize = $true
    $lHint.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    $lHint.Text = "Recopiez l'identifiant ci-dessus pour deverrouiller :"

    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $width
    $t.Font = New-Object System.Drawing.Font('Consolas',10)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Confirmer"; $ok.Size = New-Object System.Drawing.Size(110,32)
    $ok.DialogResult = 'OK'; $ok.Enabled = $false
    $ko = New-Object System.Windows.Forms.Button
    $ko.Text = "Annuler"; $ko.Size = New-Object System.Drawing.Size(110,32); $ko.DialogResult = 'Cancel'

    $f.Controls.AddRange(@($l,$lHint,$t,$ok,$ko))
    $f.PerformLayout()

    # Positionnement vertical dynamique en fonction de la hauteur reelle du message.
    $y = $l.Bottom + 14
    $lHint.Location = New-Object System.Drawing.Point($margin, $y)
    $y = $lHint.Bottom + 4
    $t.Location = New-Object System.Drawing.Point($margin, $y)
    $y = $t.Bottom + 16
    $ok.Location = New-Object System.Drawing.Point(($margin + $width - $ok.Width - $ko.Width - 10), $y)
    $ko.Location = New-Object System.Drawing.Point(($margin + $width - $ko.Width), $y)

    $f.ClientSize = New-Object System.Drawing.Size(($width + 2*$margin), ($ko.Bottom + $margin))

    $t.Add_TextChanged({ $ok.Enabled = ($t.Text.Trim() -ceq $Expected) })
    $f.Add_Shown({ $t.Focus() })
    $f.AcceptButton = $ok; $f.CancelButton = $ko
    return ($f.ShowDialog($script:MainForm) -eq 'OK')
}

# ============================================================================
#  8bis. Lister les membres d'une collection (lecture + extraction CSV)
# ============================================================================
function Show-CollectionMembers {
    if (-not $script:ValidatedCollection) { return }
    $cid  = $script:ValidatedCollection.Id
    $name = $script:ValidatedCollection.Name

    if ($script:btnListMembers) { $script:btnListMembers.Enabled = $false }

    # Etat partage (references stables, mutees en place pour rester visibles des closures).
    $rows      = New-Object System.Collections.Generic.List[object]
    $directSet = New-Object 'System.Collections.Generic.HashSet[int]'
    $extraCols = New-Object System.Collections.Generic.List[object]   # elements : @{ Prop=..; Header=.. }
    $resMap    = @{}   # ResourceID -> LastLogonUser (construit une seule fois par ouverture)

    # --- Chargement / rechargement des membres ---
    $loadData = {
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Write-Log "Lecture des membres de '$name' ($cid)..." 'INFO'
        $rows.Clear(); $directSet.Clear()
        $members = @()
        try {
            $prev = Get-Location
            Set-Location "$($script:CMDrive):\" -ErrorAction Stop
            foreach ($r in @(Get-CMDeviceCollectionDirectMembershipRule -CollectionId $cid -ErrorAction SilentlyContinue)) {
                [void]$directSet.Add([int]$r.ResourceID)
            }
            $members = @(Get-CMCollectionMember -CollectionId $cid -ErrorAction SilentlyContinue)

            # Table ResourceID -> derniere ouverture de session, construite une seule fois.
            # SMS_CollectionMember n'expose pas LastLogonUser : on l'obtient via SMS_R_System.
            if ($resMap.Count -eq 0) {
                Write-Log "Construction de la table 'dernier utilisateur' (Get-CMResource)..." 'INFO'
                foreach ($res in (Get-CMResource -ResourceType System -Fast -ErrorAction SilentlyContinue | Where-Object { $_.ResourceId })) {
                    $resMap[[int]$res.ResourceId] = [string]$res.LastLogonUserName
                }
            }
        }
        catch { Write-Log "Erreur lecture des membres : $($_.Exception.Message)" 'ERR' }
        finally {
            Set-Location $prev -ErrorAction SilentlyContinue
            $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
        foreach ($m in $members) {
            $rid = [int]$m.ResourceID
            $rows.Add([pscustomobject]@{
                Member       = $m                # objet complet, pour colonnes additionnelles
                Name         = [string]$m.Name
                ResourceID   = $rid
                Domaine      = [string]$m.Domain
                OS           = [string]$m.DeviceOS
                LastLogon    = if ($resMap.ContainsKey($rid)) { $resMap[$rid] } else { '' }
                Appartenance = if ($directSet.Contains($rid)) { 'Direct' } else { 'Requete' }
            })
        }
        Write-Log "Membres charges : $($rows.Count) (dont directs : $($directSet.Count))." 'OK'
    }

    & $loadData

    # --- Fenetre ---
    $w = New-Object System.Windows.Forms.Form
    $w.Text = "Membres de $name ($cid)"
    $w.ClientSize = New-Object System.Drawing.Size(820,560)
    $w.MinimumSize = New-Object System.Drawing.Size(680,440)
    $w.StartPosition = 'CenterParent'
    $w.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.SetBounds(12,10,640,20); $lblInfo.Anchor = 'Top,Left,Right'

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtre (nom) :"; $lblFilter.SetBounds(12,36,80,22)
    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.SetBounds(95,33,713,24); $txtFilter.Anchor = 'Top,Left,Right'

    $g = New-Object System.Windows.Forms.DataGridView
    $g.SetBounds(12,66,796,444); $g.Anchor = 'Top,Left,Right,Bottom'
    $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $true; $g.AutoSizeColumnsMode = 'Fill'

    # Bande de boutons (positionnee via ClientSize -> toujours visible).
    $btnRemoveSel = New-Object System.Windows.Forms.Button
    $btnRemoveSel.Text = "Retirer la selection (direct)"; $btnRemoveSel.SetBounds(12,518,240,32); $btnRemoveSel.Anchor = 'Bottom,Left'
    $btnRemoveSel.BackColor = [System.Drawing.Color]::FromArgb(245,215,215)
    $btnAddCol = New-Object System.Windows.Forms.Button
    $btnAddCol.Text = "Ajouter une colonne..."; $btnAddCol.SetBounds(430,518,160,32); $btnAddCol.Anchor = 'Bottom,Right'
    $btnCsv = New-Object System.Windows.Forms.Button
    $btnCsv.Text = "Extraire CSV"; $btnCsv.SetBounds(598,518,120,32); $btnCsv.Anchor = 'Bottom,Right'
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Fermer"; $btnClose.SetBounds(726,518,82,32); $btnClose.Anchor = 'Bottom,Right'; $btnClose.DialogResult = 'Cancel'

    $w.Controls.AddRange(@($lblInfo,$lblFilter,$txtFilter,$g,$btnRemoveSel,$btnAddCol,$btnCsv,$btnClose))

    # --- (Re)construction des colonnes : base fixe + colonnes additionnelles ---
    $rebuildColumns = {
        $g.Columns.Clear()
        [void]$g.Columns.Add('m1','Nom')
        [void]$g.Columns.Add('m2','ResourceID')
        [void]$g.Columns.Add('m3','Domaine')
        [void]$g.Columns.Add('m4','OS')
        [void]$g.Columns.Add('m6','Dernier utilisateur')
        [void]$g.Columns.Add('m5','Appartenance')
        foreach ($ec in $extraCols) { [void]$g.Columns.Add(('x_' + $ec.Prop), $ec.Header) }
    }

    # --- Rendu des lignes (avec filtre) ---
    $render = {
        param($flt)
        $g.Rows.Clear()
        $shown = 0
        foreach ($r in $rows) {
            if ($flt -and ($r.Name -notmatch [regex]::Escape($flt))) { continue }
            $vals = New-Object System.Collections.Generic.List[object]
            $vals.Add($r.Name); $vals.Add($r.ResourceID); $vals.Add($r.Domaine); $vals.Add($r.OS); $vals.Add($r.LastLogon); $vals.Add($r.Appartenance)
            foreach ($ec in $extraCols) {
                $v = $null
                try { $v = $r.Member.$($ec.Prop) } catch { }
                if ($v -is [array]) { $v = ($v -join '; ') }
                $vals.Add([string]$v)
            }
            $idx = $g.Rows.Add($vals.ToArray())
            if ($r.Appartenance -eq 'Direct') {
                $g.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220,235,250)
            }
            $shown++
        }
        $lblInfo.Text = "$($rows.Count) membre(s)  -  direct (retirable) : $($directSet.Count)  -  affiche(s) : $shown"
    }

    & $rebuildColumns
    & $render $null
    $txtFilter.Add_TextChanged({ & $render $txtFilter.Text.Trim() })

    # --- Ajouter une colonne : choix d'une propriete du membre ---
    $btnAddCol.Add_Click({
        if ($rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun membre : rien a lister comme propriete.","Colonne",'OK','Information') | Out-Null
            return
        }
        $already = @('Name','ResourceID','Domain','DeviceOS','LastLogonUserName') + ($extraCols | ForEach-Object { $_.Prop })
        $props = $rows[0].Member.PSObject.Properties.Name | Sort-Object -Unique | Where-Object { $_ -notin $already }
        if (-not $props) {
            [System.Windows.Forms.MessageBox]::Show("Aucune autre propriete disponible.","Colonne",'OK','Information') | Out-Null
            return
        }
        $d = New-Object System.Windows.Forms.Form
        $d.Text = "Ajouter une colonne"; $d.ClientSize = New-Object System.Drawing.Size(360,120)
        $d.FormBorderStyle = 'FixedDialog'; $d.StartPosition = 'CenterParent'; $d.MaximizeBox=$false; $d.MinimizeBox=$false
        $dl = New-Object System.Windows.Forms.Label
        $dl.Text = "Propriete a afficher :"; $dl.SetBounds(15,15,330,20)
        $cb = New-Object System.Windows.Forms.ComboBox
        $cb.SetBounds(15,40,330,24); $cb.DropDownStyle = 'DropDownList'
        [void]$cb.Items.AddRange(@($props)); $cb.SelectedIndex = 0
        $dok = New-Object System.Windows.Forms.Button
        $dok.Text = "Ajouter"; $dok.SetBounds(160,80,90,28); $dok.DialogResult='OK'
        $dko = New-Object System.Windows.Forms.Button
        $dko.Text = "Annuler"; $dko.SetBounds(255,80,90,28); $dko.DialogResult='Cancel'
        $d.Controls.AddRange(@($dl,$cb,$dok,$dko)); $d.AcceptButton=$dok; $d.CancelButton=$dko
        if ($d.ShowDialog($w) -eq 'OK' -and $cb.SelectedItem) {
            $p = [string]$cb.SelectedItem
            $extraCols.Add(@{ Prop = $p; Header = $p })
            & $rebuildColumns
            & $render $txtFilter.Text.Trim()
            Write-Log "Colonne ajoutee dans la liste des membres : $p" 'INFO'
        }
    })

    # --- Extraction CSV (inclut les colonnes additionnelles) ---
    $btnCsv.Add_Click({
        if ($rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun membre a extraire.","Extraction",'OK','Information') | Out-Null
            return
        }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV (*.csv)|*.csv"
        $sfd.FileName = ("Membres_{0}_{1}.csv" -f $cid, (Get-Date).ToString('yyyyMMdd_HHmmss'))
        if ($sfd.ShowDialog() -eq 'OK') {
            try {
                $out = foreach ($r in $rows) {
                    $o = [ordered]@{ Name=$r.Name; ResourceID=$r.ResourceID; Domaine=$r.Domaine; OS=$r.OS; 'Dernier utilisateur'=$r.LastLogon; Appartenance=$r.Appartenance }
                    foreach ($ec in $extraCols) {
                        $v = $null
                        try { $v = $r.Member.$($ec.Prop) } catch { }
                        if ($v -is [array]) { $v = ($v -join '; ') }
                        $o[$ec.Header] = [string]$v
                    }
                    [pscustomobject]$o
                }
                $out | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
                Write-Log "Extraction des membres : $($sfd.FileName)" 'OK'
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Extraction impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
            }
        }
    })

    # --- Retirer la selection (garde-fous : direct uniquement + saisie de l'ID) ---
    $btnRemoveSel.Add_Click({
        $sel = @($g.SelectedRows)
        if ($sel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Selectionnez au moins une ligne.","Retrait",'OK','Information') | Out-Null
            return
        }
        $toRemove = New-Object System.Collections.Generic.List[int]
        $skipQuery = 0
        foreach ($rowSel in $sel) {
            $app = [string]$rowSel.Cells['m5'].Value
            $rid = [int]$rowSel.Cells['m2'].Value
            if ($app -eq 'Direct') { [void]$toRemove.Add($rid) } else { $skipQuery++ }
        }
        if ($toRemove.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucun membre DIRECT dans la selection.`nLes membres issus d'une regle de requete ne sont pas retirables ici.","Retrait",'OK','Warning') | Out-Null
            return
        }
        $extra = if ($skipQuery -gt 0) { "$skipQuery membre(s) par requete ignore(s).`n`n" } else { "" }
        $prompt = "RETRAIT de $($toRemove.Count) membre(s) DIRECT(s) de la collection :`n  $name  ($cid)`n`n" +
                  $extra + "Operation potentiellement destructrice.`nPour confirmer, saisissez exactement l'ID de collection :`n`n$cid"
        if (-not (Confirm-ByTyping -Expected $cid -Prompt $prompt)) {
            Write-Log "Retrait (liste des membres) annule (ID non confirme)." 'WARN'
            return
        }

        $w.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnRemoveSel.Enabled = $false; $btnCsv.Enabled = $false; $btnAddCol.Enabled = $false; $btnClose.Enabled = $false
        Write-Log "===== RETRAIT (liste des membres) : $($toRemove.Count) sur '$name' ($cid) =====" 'INFO'
        $okc = 0; $failc = 0
        $prev = Get-Location
        try {
            Set-Location "$($script:CMDrive):\" -ErrorAction Stop
            foreach ($rid in $toRemove) {
                try {
                    Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $cid -ResourceId $rid -Force -ErrorAction Stop -Confirm:$false | Out-Null
                    $okc++; Write-Log "OK  RETRAIT  RID $rid" 'OK'
                } catch {
                    $failc++; Write-Log "ECHEC  RID $rid : $($_.Exception.Message)" 'ERR'
                }
            }
        }
        catch { Write-Log "Erreur generale (retrait liste) : $($_.Exception.Message)" 'ERR' }
        finally { Set-Location $prev -ErrorAction SilentlyContinue }

        Write-Log "----- Bilan retrait (liste) : reussite $okc / echec $failc -----" $(if($failc -gt 0){'WARN'}else{'OK'})

        # L'etat a change : on invalide la recherche principale (verrou coherent).
        $script:SimulationDone = $false
        Update-ButtonsState

        & $loadData
        & $render $txtFilter.Text.Trim()
        $w.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRemoveSel.Enabled = $true; $btnCsv.Enabled = $true; $btnAddCol.Enabled = $true; $btnClose.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Retrait termine.`nReussite : $okc`nEchec : $failc","Bilan",'OK','Information') | Out-Null
    })

    $w.CancelButton = $btnClose
    [void]$w.ShowDialog($script:MainForm)
    if ($script:btnListMembers) { $script:btnListMembers.Enabled = ($null -ne $script:ValidatedCollection -and -not $script:IsProcessing) }
}

# ============================================================================
#  9. Execution (Ajout / Retrait)
# ============================================================================
function Invoke-Operation {
    param([ValidateSet('Add','Remove')][string]$Action)

    if (-not $script:ValidatedCollection -or -not $script:SimulationDone -or -not $script:SimData) { return }

    $cid  = $script:ValidatedCollection.Id
    $name = $script:ValidatedCollection.Name

    if ($Action -eq 'Add') {
        $targets = @($script:SimData | Where-Object Category -eq 'HORS_COLLECTION')
    } else {
        $targets = @($script:SimData | Where-Object Category -eq 'MEMBRE_DIRECT')
    }

    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun appareil eligible pour cette operation dans la recherche courante.","Rien a faire",'OK','Information') | Out-Null
        return
    }

    $skipped = @($script:SimData | Where-Object Category -in @('NON_TROUVE','AMBIGU','INVALIDE')).Count
    $other   = $script:SimData.Count - $targets.Count - $skipped

    if ($Action -eq 'Add') {
        $msg = "AJOUT de $($targets.Count) appareil(s) dans :`n`n  $name`n  $cid`n`n" +
               "Ignores : deja membres ($other) / non resolus ($skipped)`n`nConfirmer l'ajout ?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg,"Confirmation d'ajout",'YesNo','Warning')
        if ($r -ne 'Yes') { Write-Log "Ajout annule par l'utilisateur." 'WARN'; return }
    } else {
        $prompt = "RETRAIT de $($targets.Count) appareil(s) de la collection :`n  $name  ($cid)`n`n" +
                  "Operation potentiellement destructrice.`nPour confirmer, saisissez exactement l'ID de collection :`n`n$cid"
        if (-not (Confirm-ByTyping -Expected $cid -Prompt $prompt)) {
            Write-Log "Retrait annule (ID non confirme)." 'WARN'; return
        }
    }

    Initialize-LogFile
    Set-ProcessingState $true
    $script:CancelRequested = $false
    Write-Log "===== $($Action.ToUpper()) : $($targets.Count) appareil(s) sur '$name' ($cid) =====" 'INFO'

    $ok = 0; $fail = 0
    $script:progress.Minimum = 0; $script:progress.Maximum = $targets.Count; $script:progress.Value = 0

    $prev = Get-Location
    try {
        Set-Location "$($script:CMDrive):\" -ErrorAction Stop
        $i = 0
        foreach ($t in $targets) {
            if ($script:CancelRequested) { Write-Log "Operation interrompue par l'utilisateur." 'WARN'; break }
            $i++; $script:progress.Value = $i
            try {
                if ($Action -eq 'Add') {
                    Add-CMDeviceCollectionDirectMembershipRule -CollectionId $cid -ResourceId $t.ResourceID -ErrorAction Stop -Confirm:$false | Out-Null
                    Write-Log "OK  AJOUT  $($t.Name) (RID $($t.ResourceID))" 'OK'
                } else {
                    Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $cid -ResourceId $t.ResourceID -Force -ErrorAction Stop -Confirm:$false | Out-Null
                    Write-Log "OK  RETRAIT  $($t.Name) (RID $($t.ResourceID))" 'OK'
                }
                $ok++
                Update-GridRowStatus $t.Name ($(if($Action -eq 'Add'){'AJOUTE'}else{'RETIRE'})) ([System.Drawing.Color]::FromArgb(200,235,200))
            }
            catch {
                $fail++
                Write-Log "ECHEC  $($t.Name) (RID $($t.ResourceID)) : $($_.Exception.Message)" 'ERR'
                Update-GridRowStatus $t.Name 'ECHEC' ([System.Drawing.Color]::FromArgb(250,210,210))
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        Write-Log "Erreur generale : $($_.Exception.Message)" 'ERR'
    }
    finally {
        Set-Location $prev -ErrorAction SilentlyContinue
        $script:progress.Value = 0
    }

    Write-Log "----- Bilan $($Action.ToUpper()) : reussite $ok / echec $fail -----" $(if($fail -gt 0){'WARN'}else{'OK'})

    # Mise a jour optionnelle de l'appartenance de la collection.
    if ($script:chkUpdate.Checked -and $ok -gt 0) {
        try {
            Invoke-InCMDrive { Invoke-CMCollectionUpdate -CollectionId $cid -ErrorAction Stop }
            Write-Log "Mise a jour de l'appartenance de la collection declenchee." 'OK'
        } catch { Write-Log "Mise a jour de collection impossible : $($_.Exception.Message)" 'WARN' }
    }

    # L'etat a change : on re-arme le verrou (nouvelle recherche obligatoire).
    $script:SimulationDone = $false
    Set-ProcessingState $false
    [System.Windows.Forms.MessageBox]::Show("Termine.`nReussite : $ok`nEchec : $fail`n`nRelancez une recherche pour toute autre action.","Bilan",'OK','Information') | Out-Null
}

function Update-GridRowStatus {
    param([string]$Name,[string]$Status,[System.Drawing.Color]$Color)
    foreach ($row in $script:grid.Rows) {
        if ($row.Cells['c1'].Value -eq $Name) {
            $row.Cells['c4'].Value = $Status
            $row.DefaultCellStyle.BackColor = $Color
            break
        }
    }
}

# ============================================================================
#  10. Interface
# ============================================================================
$script:MainForm = New-Object System.Windows.Forms.Form
$script:MainForm.Text = "SCCM - Ajout / retrait d'appareils en collection (regle directe)"
$script:MainForm.Size = New-Object System.Drawing.Size(960,860)
$script:MainForm.MinimumSize = New-Object System.Drawing.Size(880,760)
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI',9)

# --- Contexte ---
$lblCtx = New-Object System.Windows.Forms.Label
$lblCtx.SetBounds(15,12,910,20); $lblCtx.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)

# --- Collection ---
$grpColl = New-Object System.Windows.Forms.GroupBox
$grpColl.Text = "1. Collection cible"; $grpColl.SetBounds(15,40,910,95)
$grpColl.Anchor = 'Top,Left,Right'

$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Text = "Nom ou ID :"; $lbl1.SetBounds(15,28,75,20)
$script:txtCollection = New-Object System.Windows.Forms.TextBox
$script:txtCollection.SetBounds(95,25,430,24); $script:txtCollection.Anchor = 'Top,Left,Right'
$script:btnValidate = New-Object System.Windows.Forms.Button
$script:btnValidate.Text = "Valider la collection"; $script:btnValidate.SetBounds(535,24,175,26); $script:btnValidate.Anchor = 'Top,Right'
$script:btnListMembers = New-Object System.Windows.Forms.Button
$script:btnListMembers.Text = "Lister les membres"; $script:btnListMembers.SetBounds(715,24,175,26); $script:btnListMembers.Anchor = 'Top,Right'; $script:btnListMembers.Enabled = $false
$script:lblCollInfo = New-Object System.Windows.Forms.Label
$script:lblCollInfo.Text = "Collection non validee."; $script:lblCollInfo.SetBounds(15,58,875,28)
$script:lblCollInfo.ForeColor = [System.Drawing.Color]::DimGray; $script:lblCollInfo.Anchor = 'Top,Left,Right'
$grpColl.Controls.AddRange(@($lbl1,$script:txtCollection,$script:btnValidate,$script:btnListMembers,$script:lblCollInfo))

# --- Appareils ---
$grpDev = New-Object System.Windows.Forms.GroupBox
$grpDev.Text = "2. Appareils (un par ligne, ou colles ; separateurs espace/virgule/point-virgule acceptes)"
$grpDev.SetBounds(15,145,910,170); $grpDev.Anchor = 'Top,Left,Right'

$script:txtDevices = New-Object System.Windows.Forms.TextBox
$script:txtDevices.Multiline = $true; $script:txtDevices.ScrollBars = 'Vertical'
$script:txtDevices.SetBounds(15,25,720,130); $script:txtDevices.Font = New-Object System.Drawing.Font('Consolas',9)
$script:txtDevices.Anchor = 'Top,Left,Right,Bottom'

$script:btnLoadTxt = New-Object System.Windows.Forms.Button
$script:btnLoadTxt.Text = "Charger .txt"; $script:btnLoadTxt.SetBounds(750,25,145,30); $script:btnLoadTxt.Anchor = 'Top,Right'
$script:btnClear = New-Object System.Windows.Forms.Button
$script:btnClear.Text = "Vider la liste"; $script:btnClear.SetBounds(750,62,145,30); $script:btnClear.Anchor = 'Top,Right'
$script:lblCount = New-Object System.Windows.Forms.Label
$script:lblCount.Text = "0 nom unique"; $script:lblCount.SetBounds(750,105,145,40); $script:lblCount.Anchor = 'Top,Right'
$grpDev.Controls.AddRange(@($script:txtDevices,$script:btnLoadTxt,$script:btnClear,$script:lblCount))

# --- Options + actions ---
$script:chkUpdate = New-Object System.Windows.Forms.CheckBox
$script:chkUpdate.Text = "Declencher une mise a jour de l'appartenance apres l'operation"
$script:chkUpdate.SetBounds(20,325,420,22)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Dossier de log :"; $lblLog.SetBounds(20,352,95,20)
$script:txtLogFolder = New-Object System.Windows.Forms.TextBox
$script:txtLogFolder.SetBounds(120,349,470,24)
$script:txtLogFolder.Text = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'SCCM_BulkCollection_Logs')
$script:btnBrowseLog = New-Object System.Windows.Forms.Button
$script:btnBrowseLog.Text = "Parcourir"; $script:btnBrowseLog.SetBounds(600,348,90,26)

$script:btnSimulate = New-Object System.Windows.Forms.Button
$script:btnSimulate.Text = "RECHERCHER"; $script:btnSimulate.SetBounds(15,385,220,38); $script:btnSimulate.Enabled = $false
$script:btnSimulate.BackColor = [System.Drawing.Color]::FromArgb(210,225,245)
$script:btnAdd = New-Object System.Windows.Forms.Button
$script:btnAdd.Text = "AJOUTER"; $script:btnAdd.SetBounds(245,385,220,38); $script:btnAdd.Enabled = $false
$script:btnAdd.BackColor = [System.Drawing.Color]::FromArgb(210,240,210)
$script:btnRemove = New-Object System.Windows.Forms.Button
$script:btnRemove.Text = "RETIRER"; $script:btnRemove.SetBounds(475,385,220,38); $script:btnRemove.Enabled = $false
$script:btnRemove.BackColor = [System.Drawing.Color]::FromArgb(245,215,215)
$script:btnStop = New-Object System.Windows.Forms.Button
$script:btnStop.Text = "STOP"; $script:btnStop.SetBounds(705,385,220,38); $script:btnStop.Enabled = $false

# --- Progression ---
$script:progress = New-Object System.Windows.Forms.ProgressBar
$script:progress.SetBounds(15,432,910,18); $script:progress.Anchor = 'Top,Left,Right'

# --- Grille de resultats ---
$script:grid = New-Object System.Windows.Forms.DataGridView
$script:grid.SetBounds(15,458,910,180); $script:grid.Anchor = 'Top,Left,Right'
$script:grid.ReadOnly = $true; $script:grid.AllowUserToAddRows = $false
$script:grid.SelectionMode = 'FullRowSelect'; $script:grid.RowHeadersVisible = $false
$script:grid.AutoSizeColumnsMode = 'Fill'
[void]$script:grid.Columns.Add('c1','Nom demande')
[void]$script:grid.Columns.Add('c2','Categorie')
[void]$script:grid.Columns.Add('c3','ResourceID')
[void]$script:grid.Columns.Add('c5','Dernier utilisateur')
[void]$script:grid.Columns.Add('c4','Detail / Statut')

# --- Journal ---
$lblJ = New-Object System.Windows.Forms.Label
$lblJ.Text = "Journal :"; $lblJ.SetBounds(15,645,80,18); $lblJ.Anchor = 'Bottom,Left'
$script:txtLog = New-Object System.Windows.Forms.TextBox
$script:txtLog.Multiline = $true; $script:txtLog.ScrollBars = 'Vertical'; $script:txtLog.ReadOnly = $true
$script:txtLog.SetBounds(15,665,790,140); $script:txtLog.Font = New-Object System.Drawing.Font('Consolas',8.5)
$script:txtLog.BackColor = [System.Drawing.Color]::FromArgb(30,30,30); $script:txtLog.ForeColor = [System.Drawing.Color]::Gainsboro
$script:txtLog.Anchor = 'Bottom,Left,Right'
$script:btnExport = New-Object System.Windows.Forms.Button
$script:btnExport.Text = "Exporter CSV"; $script:btnExport.SetBounds(815,665,110,32); $script:btnExport.Anchor = 'Bottom,Right'
$script:btnOpenLog = New-Object System.Windows.Forms.Button
$script:btnOpenLog.Text = "Ouvrir le log"; $script:btnOpenLog.SetBounds(815,702,110,32); $script:btnOpenLog.Anchor = 'Bottom,Right'

$script:MainForm.Controls.AddRange(@(
    $lblCtx,$grpColl,$grpDev,
    $script:chkUpdate,$lblLog,$script:txtLogFolder,$script:btnBrowseLog,
    $script:btnSimulate,$script:btnAdd,$script:btnRemove,$script:btnStop,
    $script:progress,$script:grid,$lblJ,$script:txtLog,$script:btnExport,$script:btnOpenLog
))

# ============================================================================
#  11. Evenements
# ============================================================================
$script:btnValidate.Add_Click({ Validate-Collection })
$script:btnListMembers.Add_Click({ Show-CollectionMembers })
$script:txtCollection.Add_TextChanged({ Invalidate-Collection })

$script:txtDevices.Add_TextChanged({
    $n = (Get-InputDeviceNames).Count
    $script:lblCount.Text = "$n nom(s) unique(s)"
    Reset-Simulation
})

$script:btnLoadTxt.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Get-Content -LiteralPath $ofd.FileName -ErrorAction Stop
            if ($script:txtDevices.Text.Trim() -ne '') { $script:txtDevices.AppendText("`r`n") }
            $script:txtDevices.AppendText(($lines -join "`r`n"))
            Write-Log "Fichier charge : $($ofd.FileName) ($($lines.Count) ligne(s))" 'INFO'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Lecture impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        }
    }
})

$script:btnClear.Add_Click({ $script:txtDevices.Clear() })

$script:btnBrowseLog.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') { $script:txtLogFolder.Text = $fbd.SelectedPath; $script:LogFile = $null }
})

$script:btnSimulate.Add_Click({ Invoke-Simulation })
$script:btnAdd.Add_Click({ Invoke-Operation -Action 'Add' })
$script:btnRemove.Add_Click({ Invoke-Operation -Action 'Remove' })
$script:btnStop.Add_Click({ $script:CancelRequested = $true; Write-Log "Arret demande..." 'WARN' })

$script:btnExport.Add_Click({
    if (-not $script:SimData -or $script:SimData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucune donnee a exporter (lancez d'abord une recherche).","Export",'OK','Information') | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV (*.csv)|*.csv"; $sfd.FileName = "SCCM_Recherche_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $script:SimData | Select-Object Name,Category,ResourceID,LastLogon,Detail |
                Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            Write-Log "Export CSV : $($sfd.FileName)" 'OK'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export impossible :`n$($_.Exception.Message)","Erreur",'OK','Error') | Out-Null
        }
    }
})

$script:btnOpenLog.Add_Click({
    if ($script:LogFile -and (Test-Path -LiteralPath $script:LogFile)) { Start-Process notepad.exe $script:LogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Aucun fichier de log pour l'instant.","Log",'OK','Information') | Out-Null }
})

# Empeche la fermeture pendant un traitement.
$script:MainForm.Add_FormClosing({
    if ($script:IsProcessing) {
        $_.Cancel = $true
        $script:CancelRequested = $true
        [System.Windows.Forms.MessageBox]::Show("Un traitement est en cours. Il a ete interrompu ; reessayez de fermer.","Fermeture",'OK','Warning') | Out-Null
    }
})

# ============================================================================
#  12. Demarrage
# ============================================================================
if (-not (Initialize-CMContext)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Aucune connexion SCCM detectee (lecteur PSDrive de type CMSite introuvable).`n`n" +
        "Lancez ce script depuis une console ou le module ConfigurationManager est importe " +
        "et le lecteur du site est monte.",
        "Contexte SCCM manquant",'OK','Error') | Out-Null
    $lblCtx.Text = "Contexte SCCM : NON DETECTE - actions bloquees."
    $lblCtx.ForeColor = [System.Drawing.Color]::Firebrick
    $script:btnValidate.Enabled = $false
} else {
    $lblCtx.Text = "Contexte SCCM detecte - Code de site : $($script:CMDrive)   |   Operateur : $env:USERNAME"
    $lblCtx.ForeColor = [System.Drawing.Color]::DarkGreen
}

Update-ButtonsState
[void]$script:MainForm.ShowDialog()
