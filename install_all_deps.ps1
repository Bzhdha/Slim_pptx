# Script PowerShell pour installer toutes les dépendances nécessaires
# Usage: .\install_all_deps.ps1

Write-Host "Installation de toutes les dépendances pour Slim_pptx..." -ForegroundColor Cyan

# Vérifier que l'environnement virtuel est activé
if (-not $env:VIRTUAL_ENV) {
    Write-Host "Erreur: Aucun environnement virtuel activé" -ForegroundColor Red
    Write-Host "Activez d'abord l'environnement avec: .\activate_env.ps1" -ForegroundColor Yellow
    exit 1
}

Write-Host "Environnement virtuel activé: $env:VIRTUAL_ENV" -ForegroundColor Green

# Mettre à jour pip
Write-Host "Mise à jour de pip..." -ForegroundColor Yellow
python -m pip install --upgrade pip

# Installer toutes les dépendances nécessaires
Write-Host "Installation des dépendances..." -ForegroundColor Yellow

$dependencies = @(
    "PyPDF2>=3.0.0",
    "python-pptx>=0.6.22", 
    "Pillow>=10.0.0",
    "tkinterdnd2"
)

foreach ($dep in $dependencies) {
    Write-Host "Installation de $dep..." -ForegroundColor Cyan
    pip install $dep
}

# Vérifier les installations
Write-Host "`nVérification des installations..." -ForegroundColor Yellow
pip list | Select-String -Pattern "PyPDF2|pptx|Pillow|tkinterdnd2"

Write-Host "`nInstallation terminée!" -ForegroundColor Green
Write-Host "Vous pouvez maintenant lancer: python slim_pptx.py" -ForegroundColor Cyan 