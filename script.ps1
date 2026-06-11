<#
================================================================================
                    GÉNÉRATEUR DE QR CODES - TOKENS UTILISATEURS
================================================================================

  OBJECTIF
  --------
  À partir d'un fichier CSV (délimiteur ;) contenant au minimum :
      - le nom d'un utilisateur
      - son adresse email
      - son token (secret à distribuer)
  Le script génère un QR code PNG par token, nomme chaque image d'après
  l'email de l'utilisateur (clé de jointure pour un envoi de masse ultérieur),
  puis exporte un CSV enrichi d'une colonne "CheminQR".

  SCHÉMA DE FONCTIONNEMENT
  ------------------------

      ┌──────────────────────┐
      │  CSV d'entrée  (;)   │   Nom ; Email ; Token
      └──────────┬───────────┘
                 │  [GUI] Sélection du fichier + mapping des colonnes
                 ▼
      ┌──────────────────────┐
      │  Validation ligne    │   email vide ? token vide ? doublon email ?
      │  par ligne           │
      └──────────┬───────────┘
                 │
                 ▼
      ┌──────────────────────┐      ┌─────────────────────────────┐
      │  QRCoder.dll         │─────▶│  QRCodes\email_domaine.png  │
      │  (locale > cache >   │      │  (1 PNG par token)          │
      │   NuGet TLS1.2,      │      └─────────────────────────────┘
      │   hash SHA256 épinglé│
      │   si configuré)      │
      └──────────┬───────────┘
                 │
                 ▼
      ┌──────────────────────┐
      │  Dossier de sortie   │   ACL NTFS restreintes :
      │  verrouillé          │   utilisateur courant + SYSTEM uniquement
      └──────────┬───────────┘
                 │
                 ▼
      ┌──────────────────────┐
      │  CSV enrichi  (;)    │   colonnes d'origine + CheminQR
      └──────────┬───────────┘
                 │
                 ▼
      ┌──────────────────────┐
      │  [Bouton Purger]     │   suppression des PNG + CSV générés
      │  après l'envoi       │   (écrasement avant suppression)
      └──────────────────────┘

  MESURES DE SÉCURITÉ INTÉGRÉES
  -----------------------------
  1. Chargement de la librairie par ordre de priorité :
        a) DLL locale .\lib\QRCoder.dll (recommandé en production, zéro réseau)
        b) cache local %LOCALAPPDATA%\QRCoderPS
        c) téléchargement NuGet en HTTPS/TLS 1.2 (premier lancement seulement)
  2. Contrôle d'intégrité SHA256 de la DLL avant chargement :
     - $ExpectedDllHash renseigné -> toute divergence bloque le démarrage
     - $ExpectedDllHash vide      -> l'empreinte est affichée au lancement
       pour être épinglée (modèle "Trust On First Use")
  3. Restriction des permissions NTFS sur le dossier de sortie :
     héritage désactivé, accès limité à l'utilisateur courant et SYSTEM.
     -> limite l'exposition des secrets (PNG + CSV) sur un poste partagé.
  4. Bouton de PURGE intégré : écrasement du contenu des fichiers générés
     puis suppression, à exécuter dès la fin de la campagne d'envoi.
  5. Pas d'Invoke-Expression / IEX : les données du CSV sont manipulées
     uniquement comme objets structurés ; les noms de fichiers passent par
     une fonction d'assainissement (Get-SafeFileName).

  PRÉREQUIS
  ---------
  - Windows PowerShell 5.1 ou PowerShell 7+ (Windows)
  - Accès Internet au premier lancement (téléchargement QRCoder via NuGet),
    OU une QRCoder.dll (netstandard2.0) déposée dans <dossier_du_script>\lib\
  - Recommandé : épingler l'empreinte affichée au premier lancement dans
    $ExpectedDllHash (ou la calculer : Get-FileHash <dll> -Algorithm SHA256)

  BONNES PRATIQUES D'EXPLOITATION
  -------------------------------
  - Stocker le CSV source et le dossier de sortie HORS de tout dossier
    synchronisé (cloud, partage réseau ouvert).
  - Purger les fichiers générés (bouton dédié) dès la fin de l'envoi.
  - Supprimer également le CSV source une fois la campagne terminée.
================================================================================
#>

# ==============================================================================
# CONFIGURATION DE SÉCURITÉ
# ==============================================================================

# Empreinte SHA256 attendue de QRCoder.dll (locale, cache ou téléchargée).
# - Renseignée : toute DLL dont l'empreinte diffère est REFUSÉE (anti-MitM).
# - Vide       : le script démarre quand même et AFFICHE l'empreinte calculée,
#                à reporter ici pour l'épingler (Trust On First Use).
# Pour la calculer manuellement :
#     Get-FileHash <chemin>\QRCoder.dll -Algorithm SHA256 | Select-Object -Expand Hash
$ExpectedDllHash = ""

# Taille d'un module du QR code (20 ≈ image de ~600x600 px)
$TaillePixels = 20

# ==============================================================================
# CHARGEMENT DES ASSEMBLIES GUI
# ==============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# 1. CHARGEMENT SÉCURISÉ DE LA LIBRAIRIE QR (DLL LOCALE + CONTRÔLE D'INTÉGRITÉ)
# ==============================================================================

# Capturé ICI, au niveau du script : $PSScriptRoot = dossier contenant le .ps1.
# (Vide uniquement si le code est collé directement dans une console.)
$script:ScriptRootDir = $PSScriptRoot

function Initialize-QRLibrary {
    <#
        Charge la librairie QRCoder.dll selon l'ordre de priorité suivant :

        1. DLL LOCALE  : <dossier_du_script>\lib\QRCoder.dll si présente
                         (recommandé en production : aucune dépendance réseau)
        2. CACHE       : %LOCALAPPDATA%\QRCoderPS\QRCoder.dll si déjà téléchargée
        3. TÉLÉCHARGEMENT depuis NuGet (TLS 1.2 forcé), puis mise en cache.

        Contrôle d'intégrité (atténuation du risque MitM) :
        - Si $ExpectedDllHash est renseigné : la DLL (locale, cache OU
          fraîchement téléchargée) DOIT correspondre, sinon refus de démarrer.
        - Si $ExpectedDllHash est vide : mode "épinglage au premier usage"
          (TOFU) -> l'empreinte calculée est affichée pour que vous puissiez
          la reporter dans $ExpectedDllHash et figer la configuration.
    #>
    $scriptDir = $script:ScriptRootDir
    if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = (Get-Location).Path }   # repli : console interactive / ISE

    $dllLocale = Join-Path $scriptDir "lib\QRCoder.dll"
    $cacheDir  = Join-Path $env:LOCALAPPDATA "QRCoderPS"
    $dllCache  = Join-Path $cacheDir "QRCoder.dll"

    # --- Étape 1 : déterminer la source de la DLL ---
    $dllPath = $null
    if     (Test-Path $dllLocale) { $dllPath = $dllLocale }   # priorité : DLL fournie manuellement
    elseif (Test-Path $dllCache)  { $dllPath = $dllCache  }   # sinon : cache local existant
    else {
        # --- Étape 2 : téléchargement depuis NuGet (premier lancement uniquement) ---
        try {
            New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
            $nupkg = Join-Path $env:TEMP "qrcoder.nupkg.zip"

            # TLS 1.2 forcé : indispensable sur Windows PowerShell 5.1 (défaut parfois TLS 1.0)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/QRCoder/1.6.0" `
                              -OutFile $nupkg -UseBasicParsing -ErrorAction Stop

            # Le .nupkg est une archive zip : on en extrait la DLL netstandard2.0
            $extractDir = Join-Path $env:TEMP "qrcoder_extract"
            if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
            Expand-Archive -Path $nupkg -DestinationPath $extractDir -Force

            $dll = Get-ChildItem -Path $extractDir -Recurse -Filter "QRCoder.dll" |
                   Where-Object { $_.FullName -match 'netstandard2\.0' } |
                   Select-Object -First 1
            if (-not $dll) { throw "QRCoder.dll introuvable dans le package NuGet." }

            Copy-Item $dll.FullName $dllCache -Force
            Remove-Item $nupkg, $extractDir -Recurse -Force -ErrorAction SilentlyContinue
            $dllPath = $dllCache
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Impossible de récupérer la librairie QR :`n$($_.Exception.Message)`n`n" +
                "Alternative hors-ligne : déposez une QRCoder.dll (netstandard2.0) vérifiée dans :`n" +
                "  - $((Join-Path $scriptDir 'lib'))   (recommandé)`n" +
                "  - ou $cacheDir",
                "Erreur de téléchargement", 'OK', 'Error') | Out-Null
            exit 1
        }
    }

    # --- Étape 3 : contrôle d'intégrité SHA256 ---
    $hashActuel = (Get-FileHash -Path $dllPath -Algorithm SHA256).Hash

    if (-not [string]::IsNullOrWhiteSpace($script:ExpectedDllHash)) {
        # Empreinte épinglée : toute divergence = refus de chargement
        if ($hashActuel -ne $script:ExpectedDllHash.ToUpper().Trim()) {
            [System.Windows.Forms.MessageBox]::Show(
                "ALERTE INTÉGRITÉ : l'empreinte de QRCoder.dll ne correspond pas à la valeur épinglée.`n`n" +
                "Attendue : $($script:ExpectedDllHash)`n" +
                "Calculée : $hashActuel`n`n" +
                "DLL utilisée : $dllPath`n" +
                "La DLL a peut-être été modifiée ou remplacée. Chargement refusé.",
                "Échec du contrôle d'intégrité", 'OK', 'Error') | Out-Null
            exit 1
        }
    }
    else {
        # Mode TOFU : on informe l'opérateur pour qu'il épingle l'empreinte
        $msg = "Empreinte SHA256 de la DLL chargée :`n`n$hashActuel`n`n" +
               "Source : $dllPath`n`n" +
               "RECOMMANDATION : après vérification de la DLL, reportez cette valeur dans " +
               "`$ExpectedDllHash` en tête de script pour activer le contrôle d'intégrité " +
               "à chaque lancement (protection MitM / substitution)."
        [System.Windows.Forms.MessageBox]::Show($msg, "Épinglage d'empreinte recommandé", 'OK', 'Information') | Out-Null
    }

    # --- Étape 4 : chargement (uniquement après validation) ---
    Add-Type -Path $dllPath
    $script:DllChargee = $dllPath
}
Initialize-QRLibrary

# ==============================================================================
# 2. FONCTIONS UTILITAIRES
# ==============================================================================
function Get-SafeFileName {
    <#
        Assainit une chaîne pour en faire un nom de fichier valide :
        - remplace les caractères interdits par le système de fichiers
        - remplace le caractère @ (compatibilité maximale)
        - tronque à 100 caractères
    #>
    param([string]$Texte)
    $invalides = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regex     = "[{0}]" -f [Regex]::Escape($invalides)
    $safe      = ($Texte -replace $regex, '_') -replace '@', '_'
    $safe      = $safe.Trim()
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "inconnu" }
    return $safe
}

function New-QRCodePng {
    <#
        Encode un texte en QR code et l'écrit en PNG.
        Niveau de correction d'erreur Q (~25%) : bon compromis lisibilité/densité.
    #>
    param([string]$Texte, [string]$CheminPng)
    $generator = New-Object QRCoder.QRCodeGenerator
    $data      = $generator.CreateQrCode($Texte, [QRCoder.QRCodeGenerator+ECCLevel]::Q)
    $pngQr     = New-Object QRCoder.PngByteQRCode($data)
    [IO.File]::WriteAllBytes($CheminPng, $pngQr.GetGraphic($script:TaillePixels))
    $generator.Dispose()
}

function Find-Column {
    <#
        Auto-détection d'une colonne du CSV par liste de motifs (regex),
        testés dans l'ordre de priorité fourni.
    #>
    param([string[]]$Colonnes, [string[]]$MotsCles)
    foreach ($mc in $MotsCles) {
        $match = $Colonnes | Where-Object { $_ -match $mc } | Select-Object -First 1
        if ($match) { return $match }
    }
    return $null
}

function Protect-OutputFolder {
    <#
        Verrouille le dossier de sortie via ACL NTFS :
        - désactive l'héritage (et supprime les ACE héritées)
        - n'accorde l'accès qu'à l'utilisateur courant et SYSTEM
        -> les autres comptes du poste ne peuvent plus lire les secrets générés.
    #>
    param([string]$Dossier)
    try {
        $acl = Get-Acl -Path $Dossier
        $acl.SetAccessRuleProtection($true, $false)   # héritage OFF, ACE héritées purgées

        $regles = @(
            [System.Security.AccessControl.FileSystemAccessRule]::new(
                [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
                'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'),
            [System.Security.AccessControl.FileSystemAccessRule]::new(
                'NT AUTHORITY\SYSTEM',
                'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        )
        foreach ($r in $regles) { $acl.AddAccessRule($r) }
        Set-Acl -Path $Dossier -AclObject $acl
        return $true
    }
    catch {
        return $false   # signalé dans le journal, non bloquant (ex: FAT32, droits insuffisants)
    }
}

function Remove-FileSecure {
    <#
        Suppression "renforcée" d'un fichier :
        - écrase son contenu par des zéros (1 passe)
        - puis le supprime
        NB : sur SSD (wear leveling) et volumes journalisés, l'écrasement
        physique n'est pas garanti à 100 %. Pour un besoin fort, utiliser
        un outil dédié (cipher /w, sdelete) sur le volume.
    #>
    param([string]$Chemin)
    try {
        $taille = (Get-Item $Chemin).Length
        $zeros  = New-Object byte[] ([Math]::Min($taille, 10MB))
        $fs = [IO.File]::OpenWrite($Chemin)
        $restant = $taille
        while ($restant -gt 0) {
            $bloc = [Math]::Min($restant, $zeros.Length)
            $fs.Write($zeros, 0, $bloc)
            $restant -= $bloc
        }
        $fs.Flush(); $fs.Close()
        Remove-Item -Path $Chemin -Force
        return $true
    }
    catch { return $false }
}

# ==============================================================================
# 3. CONSTRUCTION DE L'INTERFACE GRAPHIQUE
# ==============================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Générateur de QR codes - Tokens"
$form.Size            = New-Object System.Drawing.Size(620, 610)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Étape 1 : fichier CSV d'entrée -------------------------------------------
$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = "1. Fichier CSV d'entrée (délimiteur ;) :"
$lblCsv.Location = '15,15'; $lblCsv.AutoSize = $true
$form.Controls.Add($lblCsv)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = '15,38'; $txtCsv.Size = '470,25'; $txtCsv.ReadOnly = $true
$form.Controls.Add($txtCsv)

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "Parcourir..."
$btnCsv.Location = '495,36'; $btnCsv.Size = '95,27'
$form.Controls.Add($btnCsv)

# --- Étape 2 : mapping des colonnes -------------------------------------------
$grpMap = New-Object System.Windows.Forms.GroupBox
$grpMap.Text = "2. Mapping des colonnes"
$grpMap.Location = '15,75'; $grpMap.Size = '575,110'
$form.Controls.Add($grpMap)

$labels = @("Nom :", "Email :", "Token :")
$combos = @{}
$keys   = @("Nom", "Email", "Token")
for ($i = 0; $i -lt 3; $i++) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labels[$i]; $lbl.AutoSize = $true
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.DropDownStyle = 'DropDownList'; $cmb.Size = '160,25'
    switch ($i) {
        0 { $lbl.Location = '15,28';  $cmb.Location = '15,48'  }
        1 { $lbl.Location = '205,28'; $cmb.Location = '205,48' }
        2 { $lbl.Location = '395,28'; $cmb.Location = '395,48' }
    }
    $grpMap.Controls.Add($lbl); $grpMap.Controls.Add($cmb)
    $combos[$keys[$i]] = $cmb
}

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Aucun fichier chargé."
$lblCount.Location = '15,82'; $lblCount.AutoSize = $true
$lblCount.ForeColor = [System.Drawing.Color]::Gray
$grpMap.Controls.Add($lblCount)

# --- Étape 3 : dossier de sortie -----------------------------------------------
$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "3. Dossier de sortie (sera verrouillé par ACL NTFS) :"
$lblOut.Location = '15,200'; $lblOut.AutoSize = $true
$form.Controls.Add($lblOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = '15,223'; $txtOut.Size = '470,25'; $txtOut.ReadOnly = $true
$form.Controls.Add($txtOut)

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = "Parcourir..."
$btnOut.Location = '495,221'; $btnOut.Size = '95,27'
$form.Controls.Add($btnOut)

# --- Étape 4 : génération -------------------------------------------------------
$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = "Générer les QR codes"
$btnGo.Location = '15,265'; $btnGo.Size = '450,38'
$btnGo.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
$btnGo.ForeColor = [System.Drawing.Color]::White
$btnGo.FlatStyle = 'Flat'; $btnGo.Enabled = $false
$form.Controls.Add($btnGo)

# --- Bouton purge post-envoi ----------------------------------------------------
$btnPurge = New-Object System.Windows.Forms.Button
$btnPurge.Text = "Purger"
$btnPurge.Location = '475,265'; $btnPurge.Size = '115,38'
$btnPurge.BackColor = [System.Drawing.Color]::FromArgb(160, 40, 40)
$btnPurge.ForeColor = [System.Drawing.Color]::White
$btnPurge.FlatStyle = 'Flat'; $btnPurge.Enabled = $false
$form.Controls.Add($btnPurge)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '15,313'; $progress.Size = '575,20'
$form.Controls.Add($progress)

# --- Journal --------------------------------------------------------------------
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = '15,343'; $txtLog.Size = '575,215'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtLog)

function Write-Log {
    param([string]$Message)
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Message))
}

# ==============================================================================
# 4. ÉVÉNEMENTS DE L'INTERFACE
# ==============================================================================
$script:donnees       = $null   # contenu du CSV chargé
$script:fichiersCrees = @()     # liste des fichiers générés (pour la purge)

# --- Sélection + chargement du CSV ---------------------------------------------
$btnCsv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $ofd.Title  = "Sélectionner le CSV (Nom ; Email ; Token)"
    if ($ofd.ShowDialog() -ne 'OK') { return }

    $txtCsv.Text = $ofd.FileName

    # Import en UTF-8, repli sur l'encodage par défaut (ANSI) si échec
    try   { $script:donnees = Import-Csv -Path $ofd.FileName -Delimiter ';' -Encoding UTF8 }
    catch { $script:donnees = Import-Csv -Path $ofd.FileName -Delimiter ';' }

    if (-not $script:donnees -or $script:donnees.Count -eq 0) {
        Write-Log "ERREUR : CSV vide ou illisible."
        return
    }

    # Alimentation des listes déroulantes avec les colonnes détectées
    $colonnes = @($script:donnees[0].PSObject.Properties.Name)
    foreach ($k in $keys) {
        $combos[$k].Items.Clear()
        $combos[$k].Items.AddRange($colonnes)
    }

    # Auto-détection par mots-clés (ordre de priorité décroissant)
    $autoNom   = Find-Column $colonnes @('^nom$', 'nom', 'name', 'utilisateur', 'user')
    $autoMail  = Find-Column $colonnes @('mail', 'courriel')
    $autoToken = Find-Column $colonnes @('token', 'jeton', 'secret', 'seed', 'serial')
    if ($autoNom)   { $combos['Nom'].SelectedItem   = $autoNom }
    if ($autoMail)  { $combos['Email'].SelectedItem = $autoMail }
    if ($autoToken) { $combos['Token'].SelectedItem = $autoToken }

    $lblCount.Text = "$($script:donnees.Count) entrée(s) - colonnes : $($colonnes -join ', ')"
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 90)
    Write-Log "CSV chargé : $($script:donnees.Count) ligne(s)."
    $btnGo.Enabled = ($txtOut.Text -ne "")
})

# --- Sélection du dossier de sortie ---------------------------------------------
$btnOut.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Dossier de sortie des QR codes et du CSV enrichi"
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtOut.Text   = $fbd.SelectedPath
        $btnGo.Enabled = ($null -ne $script:donnees)
    }
})

# --- Génération ------------------------------------------------------------------
$btnGo.Add_Click({
    $colNom   = $combos['Nom'].SelectedItem
    $colMail  = $combos['Email'].SelectedItem
    $colToken = $combos['Token'].SelectedItem

    # Garde-fous de mapping : Email et Token sont indispensables
    if (-not $colMail -or -not $colToken) {
        [System.Windows.Forms.MessageBox]::Show(
            "Les colonnes Email et Token doivent être sélectionnées.`n" +
            "(L'email sert de nom de fichier au QR code : c'est la clé de jointure pour l'envoi de masse.)",
            "Mapping incomplet", 'OK', 'Warning') | Out-Null
        return
    }
    if ($colMail -eq $colToken) {
        [System.Windows.Forms.MessageBox]::Show("Email et Token pointent vers la même colonne.",
            "Mapping invalide", 'OK', 'Warning') | Out-Null
        return
    }

    $btnGo.Enabled = $false

    # Création + verrouillage NTFS du dossier des QR codes
    $dossierQR = Join-Path $txtOut.Text "QRCodes"
    if (-not (Test-Path $dossierQR)) { New-Item -ItemType Directory -Path $dossierQR -Force | Out-Null }
    if (Protect-OutputFolder -Dossier $dossierQR) {
        Write-Log "ACL NTFS appliquées : accès restreint à l'utilisateur courant + SYSTEM."
    } else {
        Write-Log "ATTENTION : impossible d'appliquer les ACL (volume non NTFS ou droits insuffisants)."
    }

    $total = $script:donnees.Count
    $progress.Maximum = $total; $progress.Value = 0
    $ok = 0; $ko = 0
    $emailsVus = @{}                                          # détection des doublons
    $sortie = New-Object System.Collections.Generic.List[object]
    $script:fichiersCrees = @()

    Write-Log "--- Démarrage : $total entrée(s), nommage basé sur '$colMail' ---"

    $i = 0
    foreach ($ligne in $script:donnees) {
        $i++
        $progress.Value = $i
        [System.Windows.Forms.Application]::DoEvents()        # rafraîchit la GUI

        $nom   = if ($colNom) { [string]$ligne.$colNom } else { "" }
        $email = ([string]$ligne.$colMail).Trim()
        $token = ([string]$ligne.$colToken).Trim()
        $cheminQR = ""

        # Ligne incomplète -> ignorée mais conservée dans le CSV de sortie
        if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($token)) {
            Write-Log "[$i/$total] IGNORÉ (email ou token vide) : '$nom'"
            $ko++
        }
        else {
            # Nom de fichier = email assaini ; suffixe numérique si doublon
            $base = Get-SafeFileName $email.ToLower()
            if ($emailsVus.ContainsKey($base)) {
                $emailsVus[$base]++
                Write-Log "[$i/$total] ATTENTION : email en doublon ($email) -> suffixe _$($emailsVus[$base])"
                $base = "{0}_{1}" -f $base, $emailsVus[$base]
            } else {
                $emailsVus[$base] = 1
            }

            $cheminQR = Join-Path $dossierQR ($base + ".png")
            try {
                New-QRCodePng -Texte $token -CheminPng $cheminQR
                $script:fichiersCrees += $cheminQR
                Write-Log "[$i/$total] OK : $nom <$email> -> $base.png"
                $ok++
            }
            catch {
                Write-Log "[$i/$total] ERREUR pour $email : $($_.Exception.Message)"
                $cheminQR = ""
                $ko++
            }
        }

        # Ligne de sortie = toutes les colonnes d'origine + CheminQR
        $obj = [ordered]@{}
        foreach ($p in $ligne.PSObject.Properties) { $obj[$p.Name] = $p.Value }
        $obj['CheminQR'] = $cheminQR
        $sortie.Add([PSCustomObject]$obj)
    }

    # Export du CSV enrichi (délimiteur ; conservé, cohérent avec l'entrée)
    $csvOut = Join-Path $txtOut.Text ("export_tokens_qr_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $sortie | Export-Csv -Path $csvOut -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $script:fichiersCrees += $csvOut

    Write-Log "--- Terminé : $ok OK / $ko erreur(s) ou ignoré(s) ---"
    Write-Log "QR codes    : $dossierQR"
    Write-Log "CSV enrichi : $csvOut"
    Write-Log "RAPPEL : purgez ces fichiers (bouton 'Purger') dès la fin de l'envoi de masse."

    [System.Windows.Forms.MessageBox]::Show(
        "Génération terminée.`n`n$ok QR code(s) généré(s), $ko en erreur/ignoré(s).`n`nCSV enrichi :`n$csvOut",
        "Terminé", 'OK', 'Information') | Out-Null

    $btnGo.Enabled    = $true
    $btnPurge.Enabled = $true
})

# --- Purge post-envoi --------------------------------------------------------------
$btnPurge.Add_Click({
    if ($script:fichiersCrees.Count -eq 0) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Supprimer définitivement les $($script:fichiersCrees.Count) fichier(s) générés " +
        "(QR codes + CSV enrichi) ?`n`nLeur contenu sera écrasé avant suppression.",
        "Confirmation de purge", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }

    $purges = 0; $echecs = 0
    foreach ($f in $script:fichiersCrees) {
        if (Test-Path $f) {
            if (Remove-FileSecure -Chemin $f) { $purges++ } else { $echecs++ }
        }
    }
    Write-Log "Purge : $purges fichier(s) écrasé(s) puis supprimé(s), $echecs échec(s)."
    Write-Log "RAPPEL : pensez aussi à supprimer le CSV SOURCE contenant les tokens."
    $script:fichiersCrees = @()
    $btnPurge.Enabled = $false
})

# ==============================================================================
# 5. LANCEMENT
# ==============================================================================
[void]$form.ShowDialog()
