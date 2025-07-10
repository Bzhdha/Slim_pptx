@echo off
echo Création de l'environnement virtuel...
python -m venv venv

echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat

echo Installation des dépendances...
pip install -r requirements.txt

echo Configuration terminée !
echo Pour lancer l'application, utilisez run.bat
pause 