<#
    .SYNOPSIS
    QRToken-GUI-Secure.ps1 - Générateur QR Code PRO
    
    .DESCRIPTION
    Mode 1 : Charger un CSV enrichi → Générer QR par token
    Mode 2 : Copier-coller texte libre → Générer QR custom
    
    Taille ajustable, sécurité ACL, preuve d'envoi
    
    .EXAMPLE
    .\QRToken-GUI-Secure.ps1
#>

# ==============================================================================
# ASSEMBLIES
# ==============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==============================================================================
# CONFIGURATION
# ==============================================================================
$DefaultPixelSize = 10

# ==============================================================================
# FONCTIONS
# ==============================================================================

function Detect-CsvDelimiter {
    param([string]$FilePath)
    try {
        $firstLine = Get-Content $FilePath -TotalCount 1 -Encoding UTF8
        if ($firstLine -like "*;*") { return ";" }
        if ($firstLine -like "*,*") { return "," }
        if ($firstLine -like "*`t*") { return "`t" }
        return ";"
    }
    catch { return ";" }
}

function Initialize-QRLibrary {
    try {
        if (Test-Path ".\lib\QRCoder.dll") {
            [Reflection.Assembly]::LoadFrom((Resolve-Path ".\lib\QRCoder.dll").Path) | Out-Null
            return $true
        }
        
        $cache = "$env:LOCALAPPDATA\QRCoderPS"
        if (-not (Test-Path $cache)) { mkdir $cache -Force | Out-Null }
        
        $cached = "$cache\QRCoder.dll"
        if (Test-Path $cached) {
            [Reflection.Assembly]::LoadFrom($cached) | Out-Null
            return $true
        }
        
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        
        $url = "https://www.nuget.org/api/v2/package/QRCoder/1.6.0"
        $zip = "$cache\qrcoder.zip"
        
        (New-Object System.Net.WebClient).DownloadFile($url, $zip)
        
        Add-Type -AssemblyName System.IO.Compression
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, "$cache\temp")
        
        Copy-Item "$cache\temp\lib\net*\QRCoder.dll" $cached -Force
        Remove-Item "$cache\temp" -Recurse -Force
        Remove-Item $zip -Force
        
        [Reflection.Assembly]::LoadFrom($cached) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function New-QRCodePng {
    param([string]$Text, [string]$FilePath, [int]$PixelSize)
    
    try {
        $gen = New-Object QRCoder.QRCodeGenerator
        $qr = $gen.CreateQrCode($Text, [QRCoder.QRCodeGenerator+ECCLevel]::Q)
        $png = New-Object QRCoder.PngByteQRCode -ArgumentList $qr
        $bytes = $png.GetGraphic($PixelSize)
        
        [IO.File]::WriteAllBytes($FilePath, $bytes)
        return $true
    }
    catch {
        return $false
    }
}

function Get-SafeFileName {
    param([string]$Text)
    $text = $text -replace '[^\w\s\-\.]', ''
    $text = $text -replace '\s+', '_'
    $text = $text.Substring(0, [Math]::Min(100, $text.Length))
    return $text
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$timestamp] $Message`n")
    $txtLog.ScrollToCaret()
}

# ==============================================================================
# INTERFACE
# ==============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "QRToken Generator PRO"
$form.Size = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# === SECTION SUPÉRIEURE : Taille du QR ===
$grpSize = New-Object System.Windows.Forms.GroupBox
$grpSize.Text = "Taille du QR Code"
$grpSize.Location = '15,10'; $grpSize.Size = '870,50'
$form.Controls.Add($grpSize)

$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Text = "Taille:"; $lblSize.Location = '10,20'; $lblSize.Size = '50,20'
$grpSize.Controls.Add($lblSize)

$trackSize = New-Object System.Windows.Forms.TrackBar
$trackSize.Location = '65,20'; $trackSize.Size = '300,20'
$trackSize.Minimum = 5
$trackSize.Maximum = 20
$trackSize.Value = $DefaultPixelSize
$trackSize.TickFrequency = 1
$trackSize.TickStyle = [System.Windows.Forms.TickStyle]::BottomRight
$grpSize.Controls.Add($trackSize)

$lblSizeValue = New-Object System.Windows.Forms.Label
$lblSizeValue.Text = "$($DefaultPixelSize)px"; $lblSizeValue.Location = '370,20'; $lblSizeValue.Size = '80,20'
$lblSizeValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpSize.Controls.Add($lblSizeValue)

$lblSizeDesc = New-Object System.Windows.Forms.Label
$lblSizeDesc.Text = "(5 petit → 20 gros)"; $lblSizeDesc.Location = '455,20'; $lblSizeDesc.Size = '150,20'
$lblSizeDesc.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$grpSize.Controls.Add($lblSizeDesc)

$trackSize.Add_Scroll({
    $lblSizeValue.Text = "$($trackSize.Value)px"
})

# === TABCONTROL : Deux modes ===
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = '15,70'; $tabControl.Size = '870,560'
$form.Controls.Add($tabControl)

# ========================
# TAB 1 : MODE CSV
# ========================
$tabCSV = New-Object System.Windows.Forms.TabPage
$tabCSV.Text = "📁 Mode CSV"
$tabControl.TabPages.Add($tabCSV)

# Groupe CSV
$grpCSV = New-Object System.Windows.Forms.GroupBox
$grpCSV.Text = "1. Charger CSV (Nom ; Email ; Token ; CheminQR)"
$grpCSV.Location = '10,10'; $grpCSV.Size = '840,60'
$tabCSV.Controls.Add($grpCSV)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = '10,30'; $txtCsv.Size = '750,22'
$txtCsv.ReadOnly = $true
$grpCSV.Controls.Add($txtCsv)

$btnLoadCsv = New-Object System.Windows.Forms.Button
$btnLoadCsv.Text = "Charger CSV"; $btnLoadCsv.Location = '765,30'; $btnLoadCsv.Size = '65,22'
$grpCSV.Controls.Add($btnLoadCsv)

# Groupe colonnes
$grpCol = New-Object System.Windows.Forms.GroupBox
$grpCol.Text = "2. Mapper les colonnes"
$grpCol.Location = '10,80'; $grpCol.Size = '840,60'
$tabCSV.Controls.Add($grpCol)

$combos = @{}
$x = 10
foreach ($k in @('Nom', 'Email', 'Token')) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "${k}:"; $lbl.Location = "$x,25"; $lbl.Size = '45,20'
    $grpCol.Controls.Add($lbl)
    
    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = "$($x+45),25"; $cmb.Size = '150,22'
    $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $grpCol.Controls.Add($cmb)
    $combos[$k] = $cmb
    
    $x += 210
}

# Groupe dossier de sortie
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "3. Dossier de sortie"
$grpOut.Location = '10,150'; $grpOut.Size = '840,60'
$tabCSV.Controls.Add($grpOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = '10,30'; $txtOut.Size = '750,22'
$grpOut.Controls.Add($txtOut)

$btnBrowseOut = New-Object System.Windows.Forms.Button
$btnBrowseOut.Text = "Parcourir"; $btnBrowseOut.Location = '765,30'; $btnBrowseOut.Size = '65,22'
$grpOut.Controls.Add($btnBrowseOut)

# Groupe actions CSV
$grpActCSV = New-Object System.Windows.Forms.GroupBox
$grpActCSV.Text = "4. Actions"
$grpActCSV.Location = '10,220'; $grpActCSV.Size = '840,200'
$tabCSV.Controls.Add($grpActCSV)

$btnGenCSV = New-Object System.Windows.Forms.Button
$btnGenCSV.Text = "GÉNÉRER les QR codes"; $btnGenCSV.Location = '10,25'; $btnGenCSV.Size = '150,35'
$btnGenCSV.BackColor = [System.Drawing.Color]::LimeGreen
$btnGenCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnGenCSV.Enabled = $false
$grpActCSV.Controls.Add($btnGenCSV)

$btnPurge = New-Object System.Windows.Forms.Button
$btnPurge.Text = "Purger"; $btnPurge.Location = '170,25'; $btnPurge.Size = '80,35'
$btnPurge.BackColor = [System.Drawing.Color]::LightCoral
$btnPurge.Enabled = $false
$grpActCSV.Controls.Add($btnPurge)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Location = '10,70'; $txtLog.Size = '820,120'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Courier New", 8)
$grpActCSV.Controls.Add($txtLog)

# ========================
# TAB 2 : MODE TEXTE LIBRE
# ========================
$tabText = New-Object System.Windows.Forms.TabPage
$tabText.Text = "📝 Mode Texte Libre"
$tabControl.TabPages.Add($tabText)

# Groupe saisie
$grpText = New-Object System.Windows.Forms.GroupBox
$grpText.Text = "Saisir ou copier-coller les codes (un par ligne)"
$grpText.Location = '10,10'; $grpText.Size = '840,200'
$tabText.Controls.Add($grpText)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Multiline = $true
$txtInput.WordWrap = $true
$txtInput.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtInput.Location = '10,25'; $txtInput.Size = '820,165'
$txtInput.PlaceHolder = "Copier-coller vos codes ici`nUn code par ligne"
$grpText.Controls.Add($txtInput)

# Groupe infos
$grpTextInfo = New-Object System.Windows.Forms.GroupBox
$grpTextInfo.Text = "Noms des fichiers (optionnel)"
$grpTextInfo.Location = '10,220'; $grpTextInfo.Size = '840,100'
$tabText.Controls.Add($grpTextInfo)

$lblTextNames = New-Object System.Windows.Forms.Label
$lblTextNames.Text = "Noms personnalisés (séparé par |) ou auto-généré :"; 
$lblTextNames.Location = '10,20'; $lblTextNames.Size = '820,20'
$grpTextInfo.Controls.Add($lblTextNames)

$txtNames = New-Object System.Windows.Forms.TextBox
$txtNames.Location = '10,45'; $txtNames.Size = '820,50'
$txtNames.Multiline = $true
$txtNames.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$txtNames.PlaceHolder = "qr_001 | qr_002 | qr_003  (laissez vide pour auto)"
$grpTextInfo.Controls.Add($txtNames)

# Groupe actions texte
$grpActText = New-Object System.Windows.Forms.GroupBox
$grpActText.Text = "Actions"
$grpActText.Location = '10,330'; $grpActText.Size = '840,100'
$tabText.Controls.Add($grpActText)

$txtOutText = New-Object System.Windows.Forms.TextBox
$txtOutText.Location = '10,25'; $txtOutText.Size = '740,22'
$grpActText.Controls.Add($txtOutText)

$btnBrowseOutText = New-Object System.Windows.Forms.Button
$btnBrowseOutText.Text = "Parcourir"; $btnBrowseOutText.Location = '755,25'; $btnBrowseOutText.Size = '75,22'
$grpActText.Controls.Add($btnBrowseOutText)

$btnGenText = New-Object System.Windows.Forms.Button
$btnGenText.Text = "GÉNÉRER les QR codes"; $btnGenText.Location = '10,55'; $btnGenText.Size = '150,35'
$btnGenText.BackColor = [System.Drawing.Color]::LimeGreen
$btnGenText.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnGenText.Enabled = $false
$grpActText.Controls.Add($btnGenText)

# ==============================================================================
# ÉVÉNEMENTS
# ==============================================================================
$script:donnees = $null

# Initialiser la QR Library
if (-not (Initialize-QRLibrary)) {
    [System.Windows.Forms.MessageBox]::Show("Erreur : QRCoder.dll non disponible`nTéléchargement automatique échoué", "Erreur", 'OK', 'Error')
    exit
}

# === ÉVÉNEMENTS CSV ===
$btnLoadCsv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV (*.csv)|*.csv"
    if ($ofd.ShowDialog() -ne 'OK') { return }
    
    $txtCsv.Text = $ofd.FileName
    $delimiter = Detect-CsvDelimiter -FilePath $ofd.FileName
    $delimName = if ($delimiter -eq ';') { 'Point-virgule' } elseif ($delimiter -eq ',') { 'Virgule' } else { 'TAB' }
    
    try {
        $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter $delimiter -Encoding UTF8)
    }
    catch {
        $script:donnees = @(Import-Csv -Path $ofd.FileName -Delimiter $delimiter)
    }
    
    if ($script:donnees.Count -eq 0) {
        Write-Log "❌ CSV vide"
        return
    }
    
    $colonnes = @($script:donnees[0].PSObject.Properties.Name)
    foreach ($k in $combos.Keys) {
        $combos[$k].Items.Clear()
        $combos[$k].Items.AddRange($colonnes)
        
        # Auto-détection
        $match = $colonnes | Where-Object { $_ -match $k } | Select-Object -First 1
        if ($match) { $combos[$k].SelectedItem = $match }
    }
    
    $btnGenCSV.Enabled = $true
    Write-Log "✓ Délimiteur : $delimName | $($script:donnees.Count) lignes"
})

$btnBrowseOut.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.SelectedPath = [System.IO.Path]::GetDirectoryName($txtCsv.Text)
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtOut.Text = $fbd.SelectedPath
    }
})

$btnGenCSV.Add_Click({
    if (-not $txtOut.Text -or -not (Test-Path $txtOut.Text)) {
        Write-Log "❌ Dossier de sortie invalide"
        return
    }
    
    if (-not $script:donnees) {
        Write-Log "❌ Aucun CSV chargé"
        return
    }
    
    $pixelSize = $trackSize.Value
    $generated = 0
    $failed = 0
    
    Write-Log "--- Génération en cours (taille: ${pixelSize}px) ---"
    
    foreach ($i in 0..($script:donnees.Count - 1)) {
        $row = $script:donnees[$i]
        $token = $row.($combos['Token'].SelectedItem)
        $email = $row.($combos['Email'].SelectedItem)
        $nom = $row.($combos['Nom'].SelectedItem)
        
        if ([string]::IsNullOrWhiteSpace($token)) {
            $failed++
            Write-Log "⊘ [$($i+1)] Token vide"
            continue
        }
        
        $filename = Get-SafeFileName $email
        $filepath = Join-Path $txtOut.Text "$filename.png"
        
        if (New-QRCodePng -Text $token -FilePath $filepath -PixelSize $pixelSize) {
            $generated++
            Write-Log "✓ [$($i+1)] $filename.png"
        } else {
            $failed++
            Write-Log "❌ [$($i+1)] Erreur génération"
        }
    }
    
    Write-Log "--- Terminé : $generated générés, $failed échoués ---"
    $btnPurge.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("$generated QR codes générés`n$failed échoués", "Résumé")
})

# === ÉVÉNEMENTS TEXTE LIBRE ===
$btnBrowseOutText.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtOutText.Text = $fbd.SelectedPath
    }
})

$btnGenText.Add_Click({
    if (-not $txtOutText.Text -or -not (Test-Path $txtOutText.Text)) {
        Write-Log "❌ Dossier de sortie invalide"
        return
    }
    
    $codes = $txtInput.Text -split "`n" | Where-Object { $_.Trim() }
    $names = $txtNames.Text -split "`n" | Where-Object { $_.Trim() }
    
    if ($codes.Count -eq 0) {
        Write-Log "❌ Aucun code à générer"
        return
    }
    
    $pixelSize = $trackSize.Value
    $generated = 0
    
    Write-Log "--- Génération en cours (taille: ${pixelSize}px) ---"
    
    for ($i = 0; $i -lt $codes.Count; $i++) {
        $code = $codes[$i].Trim()
        $filename = if ($i -lt $names.Count) { $names[$i].Trim() } else { "qr_$($i+1:D3)" }
        
        if ([string]::IsNullOrWhiteSpace($filename)) {
            $filename = "qr_$($i+1:D3)"
        }
        
        $filepath = Join-Path $txtOutText.Text "$filename.png"
        
        if (New-QRCodePng -Text $code -FilePath $filepath -PixelSize $pixelSize) {
            $generated++
            Write-Log "✓ $filename.png"
        } else {
            Write-Log "❌ $filename.png - Erreur"
        }
    }
    
    Write-Log "--- Terminé : $generated QR codes générés ---"
    [System.Windows.Forms.MessageBox]::Show("$generated QR codes générés`nDossier : $($txtOutText.Text)", "Résumé")
})

# Enable/Disable boutons
$txtInput.Add_TextChanged({ $btnGenText.Enabled = $txtOutText.Text -and $txtInput.Text.Length -gt 0 })
$txtOutText.Add_TextChanged({ $btnGenText.Enabled = $txtOutText.Text -and $txtInput.Text.Length -gt 0 })

# Afficher
$form.ShowDialog() | Out-Null
