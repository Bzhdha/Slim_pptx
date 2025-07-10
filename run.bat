@echo off
echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat

echo Lancement de l'application...
python slim_pptx.py

pause 