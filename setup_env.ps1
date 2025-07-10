# Script PowerShell pour configurer automatiquement l'environnement virtuel
# Usage: .\setup_env.ps1

Write-Host "Configuration de l'environnement virtuel Slim_pptx..." -ForegroundColor Cyan

# Vérifier si Python est installé
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python détecté: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Erreur: Python n'est pas installé ou n'est pas dans le PATH" -ForegroundColor Red
    exit 1
}

# Créer le dossier venv s'il n'existe pas
if (-not (Test-Path "venv")) {
    New-Item -ItemType Directory -Name "venv" | Out-Null
    Write-Host "Dossier venv créé" -ForegroundColor Yellow
}

# Créer l'environnement virtuel pdf_analyzer_env
$envPath = "venv\pdf_analyzer_env"
if (-not (Test-Path $envPath)) {
    Write-Host "Création de l'environnement virtuel pdf_analyzer_env..." -ForegroundColor Yellow
    python -m venv $envPath
    Write-Host "Environnement virtuel créé avec succès!" -ForegroundColor Green
} else {
    Write-Host "L'environnement virtuel existe déjà" -ForegroundColor Green
}

# Activer l'environnement virtuel
Write-Host "Activation de l'environnement virtuel..." -ForegroundColor Yellow
& "$envPath\Scripts\Activate.ps1"

# Mettre à jour pip
Write-Host "Mise à jour de pip..." -ForegroundColor Yellow
python -m pip install --upgrade pip

# Installer les dépendances
Write-Host "Installation des dépendances..." -ForegroundColor Yellow
$dependencies = @(
    "PyPDF2>=3.0.0",
    "Pillow>=10.0.0",
    "tkinterdnd2==0.3.0"
)

foreach ($dep in $dependencies) {
    Write-Host "Installation de $dep..." -ForegroundColor Cyan
    pip install $dep
}

# Sauvegarder l'environnement actif
"pdf_analyzer" | Out-File -FilePath "active_env.txt" -Encoding UTF8
Write-Host "Environnement pdf_analyzer configuré comme actif" -ForegroundColor Green

Write-Host "Configuration terminée avec succès!" -ForegroundColor Green
Write-Host "Vous pouvez maintenant utiliser: .\run_slim_pptx.ps1" -ForegroundColor Cyan 