# Script PowerShell pour créer la version portable de Slim_pptx
# Usage: .\build_portable.ps1

param(
    [switch]$SkipEnvCheck,
    [switch]$Verbose
)

# Configuration de l'encodage pour les caractères spéciaux
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   Création de la version portable" -ForegroundColor Cyan
Write-Host "   Slim PPTX" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Vérification de Python
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Python détecté: $pythonVersion" -ForegroundColor Green
    } else {
        throw "Python non trouvé"
    }
} catch {
    Write-Host "❌ Python n'est pas installé ou n'est pas dans le PATH" -ForegroundColor Red
    Write-Host "Veuillez installer Python et réessayer" -ForegroundColor Yellow
    Read-Host "Appuyez sur Entrée pour quitter"
    exit 1
}

Write-Host ""

# Activation de l'environnement virtuel si disponible
$envPath = "venv\unified_analyzer_env\Scripts\Activate.ps1"
if (Test-Path $envPath) {
    Write-Host "🔧 Activation de l'environnement virtuel..." -ForegroundColor Yellow
    & $envPath
    Write-Host "✅ Environnement virtuel activé" -ForegroundColor Green
    Write-Host ""
} elseif (-not $SkipEnvCheck) {
    Write-Host "⚠️ Aucun environnement virtuel trouvé" -ForegroundColor Yellow
    Write-Host "Le build continuera avec Python système" -ForegroundColor Yellow
    Write-Host ""
}

# Vérification des fichiers requis
$requiredFiles = @("slim_pptx.py", "logging_config.py", "env_config.json")
$missingFiles = @()

foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "❌ Fichiers manquants:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "   - $file" -ForegroundColor Red
    }
    Read-Host "Appuyez sur Entrée pour quitter"
    exit 1
}

Write-Host "✅ Tous les fichiers requis sont présents" -ForegroundColor Green
Write-Host ""

# Lancement du script de build
Write-Host "🚀 Lancement du build portable..." -ForegroundColor Cyan

try {
    if ($Verbose) {
        python build_portable.py
    } else {
        python build_portable.py 2>&1 | Tee-Object -FilePath "build_log.txt"
    }
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "✅ Build terminé avec succès!" -ForegroundColor Green
        Write-Host ""
        Write-Host "📁 Dossier créé: Slim_PPTX_Portable" -ForegroundColor Cyan
        Write-Host "📦 Archive créée: Slim_PPTX_Portable.zip" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Vous pouvez maintenant distribuer la version portable." -ForegroundColor Green
        
        # Ouverture du dossier si il existe
        if (Test-Path "Slim_PPTX_Portable") {
            $openFolder = Read-Host "Voulez-vous ouvrir le dossier de la version portable? (o/n)"
            if ($openFolder -eq "o" -or $openFolder -eq "O") {
                Start-Process "Slim_PPTX_Portable"
            }
        }
    } else {
        Write-Host ""
        Write-Host "❌ Erreur lors du build" -ForegroundColor Red
        if (Test-Path "build_log.txt") {
            Write-Host "Consultez le fichier build_log.txt pour plus de détails" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host ""
    Write-Host "❌ Erreur inattendue: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Read-Host "Appuyez sur Entrée pour quitter" 