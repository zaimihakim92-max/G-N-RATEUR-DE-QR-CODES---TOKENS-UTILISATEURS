<#
    .SYNOPSIS
    Install-QRCodeMaker3.0.ps1
    Installation automatique et complète de QRCode Maker 3.0
#>

Write-Host "╔════════════════════════════════════════════════════════════╗"
Write-Host "║       INSTALLATION QRCODE MAKER 3.0                       ║"
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Configuration
$basePath = "C:\qrcode_maker_3.0"
$libPath = "$basePath\lib"
$dllPath = "$libPath\QRCoder.dll"

Write-Host "📁 Répertoire d'installation : $basePath" -ForegroundColor Yellow
Write-Host ""

# ÉTAPE 1 : Créer la structure
Write-Host "1️⃣  CRÉER LA STRUCTURE DE RÉPERTOIRES" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

if (-not (Test-Path $basePath)) {
    mkdir $basePath -Force | Out-Null
    Write-Host "   ✓ Dossier créé : $basePath"
} else {
    Write-Host "   ✓ Dossier existe déjà : $basePath"
}

if (-not (Test-Path $libPath)) {
    mkdir $libPath -Force | Out-Null
    Write-Host "   ✓ Dossier créé : $libPath"
} else {
    Write-Host "   ✓ Dossier existe déjà : $libPath"
}

Write-Host ""

# ÉTAPE 2 : Télécharger QRCoder.dll
Write-Host "2️⃣  TÉLÉCHARGER QRCODER.DLL" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

if (Test-Path $dllPath) {
    $size = (Get-Item $dllPath).Length / 1KB
    Write-Host "   ✓ QRCoder.dll existe déjà ($($size)KB)"
} else {
    Write-Host "   Téléchargement de QRCoder v1.6.0..."
    
    try {
        $zipPath = "$basePath\qrcoder.zip"
        $tempPath = "$basePath\temp"
        
        # Activer TLS 1.2
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        
        # Télécharger
        $url = "https://www.nuget.org/api/v2/package/QRCoder/1.6.0"
        (New-Object System.Net.WebClient).DownloadFile($url, $zipPath)
        Write-Host "   ✓ Téléchargé : $zipPath"
        
        # Extraire
        Add-Type -AssemblyName System.IO.Compression
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempPath)
        Write-Host "   ✓ Extrait dans : $tempPath"
        
        # Copier DLL
        $sourceDll = Get-ChildItem "$tempPath\lib\net*\QRCoder.dll" -ErrorAction Stop | Select-Object -First 1
        Copy-Item $sourceDll.FullName $dllPath -Force
        Write-Host "   ✓ DLL copiée : $dllPath"
        
        # Nettoyer
        Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Write-Host "   ✓ Nettoyage terminé"
    }
    catch {
        Write-Host "   ✗ Erreur téléchargement: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "   SOLUTION: Télécharger manuellement" -ForegroundColor Yellow
        Write-Host "   1. Aller sur https://www.nuget.org/packages/QRCoder"
        Write-Host "   2. Cliquer 'Download package'"
        Write-Host "   3. Renommer en .zip"
        Write-Host "   4. Extraire et copier lib\net*\QRCoder.dll dans $libPath\"
        Write-Host ""
    }
}

Write-Host ""

# ÉTAPE 3 : Copier les scripts
Write-Host "3️⃣  COPIER LES SCRIPTS" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

$scripts = @(
    "QRToken-GUI-Secure.ps1",
    "Test-QRCoderDLL.ps1",
    "Test-QRCoderSignature.ps1"
)

foreach ($script in $scripts) {
    $sourcePath = ".\$script"
    $destPath = "$basePath\$script"
    
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $destPath -Force
        Write-Host "   ✓ Copié : $script"
    } else {
        Write-Host "   ⚠ Introuvable : $script" -ForegroundColor Yellow
    }
}

Write-Host ""

# ÉTAPE 4 : Créer le raccourci
Write-Host "4️⃣  CRÉER UN RACCOURCI RAPIDE" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Créer un script de lancement rapide
$launchScript = @"
# QRCode Maker 3.0 - Lancement rapide
Set-Location '$basePath'

Write-Host ""
Write-Host "╔════════════════════════════════════════════╗"
Write-Host "║   QRCODE MAKER 3.0                        ║"
Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Localisation : $basePath" -ForegroundColor Green
Write-Host ""
Write-Host "Disponible :" -ForegroundColor Yellow
Write-Host "  1. .\QRToken-GUI-Secure.ps1     → Lancer l'app"
Write-Host "  2. .\Test-QRCoderDLL.ps1        → Tester la DLL"
Write-Host "  3. .\Test-QRCoderSignature.ps1  → Tester la signature"
Write-Host ""
`$input = Read-Host "Choix (1/2/3) ou appuyer sur Entrée pour lancer l'app"

if (`$input -eq "2") {
    .\Test-QRCoderDLL.ps1
} elseif (`$input -eq "3") {
    .\Test-QRCoderSignature.ps1
} else {
    .\QRToken-GUI-Secure.ps1
}
"@

$launchPath = "$basePath\Lancer.ps1"
Set-Content -Path $launchPath -Value $launchScript -Encoding UTF8
Write-Host "   ✓ Créé : Lancer.ps1"

Write-Host ""

# ÉTAPE 5 : Vérification finale
Write-Host "5️⃣  VÉRIFICATION FINALE" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

$allGood = $true

# Vérifier DLL
if (Test-Path $dllPath) {
    $size = (Get-Item $dllPath).Length / 1KB
    Write-Host "   ✓ QRCoder.dll : $($size)KB"
} else {
    Write-Host "   ✗ QRCoder.dll : MANQUANT" -ForegroundColor Red
    $allGood = $false
}

# Vérifier scripts
foreach ($script in $scripts) {
    $path = "$basePath\$script"
    if (Test-Path $path) {
        Write-Host "   ✓ $script"
    } else {
        Write-Host "   ⚠ $script : Non copié" -ForegroundColor Yellow
    }
}

Write-Host ""

if ($allGood) {
    Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "✓✓✓ INSTALLATION TERMINÉE AVEC SUCCÈS ! ✓✓✓" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "📂 Structure créée :" -ForegroundColor Green
    Write-Host ""
    Write-Host "   $basePath\" -ForegroundColor Cyan
    Write-Host "   ├── QRToken-GUI-Secure.ps1" -ForegroundColor Cyan
    Write-Host "   ├── Test-QRCoderDLL.ps1" -ForegroundColor Cyan
    Write-Host "   ├── Test-QRCoderSignature.ps1" -ForegroundColor Cyan
    Write-Host "   ├── Lancer.ps1" -ForegroundColor Cyan
    Write-Host "   └── lib\" -ForegroundColor Cyan
    Write-Host "       └── QRCoder.dll" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "🚀 POUR LANCER :" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Option 1 (Recommandé):" -ForegroundColor Cyan
    Write-Host "   cd C:\qrcode_maker_3.0" -ForegroundColor Gray
    Write-Host "   .\Lancer.ps1" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   Option 2 (Direct):" -ForegroundColor Cyan
    Write-Host "   cd C:\qrcode_maker_3.0" -ForegroundColor Gray
    Write-Host "   .\QRToken-GUI-Secure.ps1" -ForegroundColor Gray
    Write-Host ""
} else {
    Write-Host "⚠️  INSTALLATION PARTIELLE" -ForegroundColor Yellow
    Write-Host "Certains éléments manquent, voir ci-dessus"
    Write-Host ""
}

Write-Host "Appuyer sur une touche pour terminer..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
