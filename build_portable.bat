@echo off
chcp 65001 >nul
echo ========================================
echo    Création de la version portable
echo    Slim PPTX
echo ========================================
echo.

REM Vérification de Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python n'est pas installé ou n'est pas dans le PATH
    echo Veuillez installer Python et réessayer
    pause
    exit /b 1
)

echo ✅ Python détecté
echo.

REM Activation de l'environnement virtuel si disponible
if exist "venv\unified_analyzer_env\Scripts\activate.bat" (
    echo 🔧 Activation de l'environnement virtuel...
    call "venv\unified_analyzer_env\Scripts\activate.bat"
    echo ✅ Environnement virtuel activé
    echo.
)

REM Lancement du script de build
echo 🚀 Lancement du build portable...
python build_portable.py

if errorlevel 1 (
    echo.
    echo ❌ Erreur lors du build
    pause
    exit /b 1
)

echo.
echo ✅ Build terminé avec succès!
echo.
echo 📁 Dossier créé: Slim_PPTX_Portable
echo 📦 Archive créée: Slim_PPTX_Portable.zip
echo.
echo Vous pouvez maintenant distribuer la version portable.
echo.
pause 