@echo off
echo Vérification de l'environnement Python...

REM Vérification si Python est installé
python --version >nul 2>&1
if errorlevel 1 (
    echo Python n'est pas installé ou n'est pas dans le PATH
    echo Veuillez installer Python 3.8 ou supérieur
    pause
    exit /b 1
)

REM Lancement du gestionnaire d'environnements
python env_manager.py

if errorlevel 1 (
    echo Une erreur s'est produite lors du lancement de l'application
    pause
    exit /b 1
) 