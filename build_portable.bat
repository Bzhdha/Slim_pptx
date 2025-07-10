@echo off
setlocal

echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat

echo Installation des dépendances...
pip install -r requirements.txt

REM === Téléchargement automatique de la DLL tkdnd2.8 si absente ===
set TKDND_URL=https://sourceforge.net/projects/tkdnd/files/tkdnd/2.8/tkdnd2.8-win64-20200224.zip/download
set TKDND_ZIP=tkdnd2.8-win64.zip

if not exist "tkdnd2.8" (
    mkdir tkdnd2.8
)

if not exist "tkdnd2.8\tkdnd2.8.dll" (
    echo Téléchargement de tkdnd2.8.dll...
    curl -L -o "%TKDND_ZIP%" "%TKDND_URL%"
    echo Extraction de tkdnd2.8.dll...
    powershell -Command "Expand-Archive -Path '%TKDND_ZIP%' -DestinationPath 'tkdnd2.8' -Force"
    del "%TKDND_ZIP%"
)

if not exist "tkdnd2.8\tkdnd2.8.dll" (
    echo ERREUR : La DLL tkdnd2.8.dll n'a pas pu être téléchargée ou extraite !
    pause
    exit /b 1
)

echo Création de la version portable...
pyinstaller --noconfirm --onefile --windowed --name "Slim_PPTX" --icon=NONE ^
    --add-binary "tkdnd2.8\\tkdnd2.8.dll;tkdnd2.8/" slim_pptx.py

echo Copie des fichiers nécessaires...
mkdir "dist\Slim_PPTX_Portable"
copy "dist\Slim_PPTX.exe" "dist\Slim_PPTX_Portable\"
copy "README.md" "dist\Slim_PPTX_Portable\"
xcopy "tkdnd2.8" "dist\Slim_PPTX_Portable\tkdnd2.8\" /E /Y

echo Nettoyage...
rmdir /s /q build
del /q "Slim_PPTX.spec"

echo Version portable créée dans le dossier dist\Slim_PPTX_Portable
pause 