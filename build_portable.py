#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de build pour créer la version portable de Slim_pptx
"""

import os
import sys
import subprocess
import shutil
import zipfile
from pathlib import Path

def run_command(command, description):
    """Exécute une commande et affiche le résultat."""
    print(f"\n{description}...")
    print(f"Commande: {command}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, 
                              capture_output=True, text=True)
        print("✅ Succès!")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erreur: {e}")
        if e.stdout:
            print("Sortie standard:", e.stdout)
        if e.stderr:
            print("Erreur:", e.stderr)
        return False

def create_portable_version():
    """Crée la version portable de l'application."""
    
    print("🚀 Création de la version portable de Slim_pptx")
    print("=" * 50)
    
    # Vérification de PyInstaller
    if not run_command("pyinstaller --version", "Vérification de PyInstaller"):
        print("❌ PyInstaller n'est pas installé. Installation...")
        if not run_command("pip install pyinstaller", "Installation de PyInstaller"):
            print("❌ Impossible d'installer PyInstaller")
            return False
    
    # Nettoyage des builds précédents
    print("\n🧹 Nettoyage des builds précédents...")
    for dir_to_clean in ["build", "dist", "__pycache__"]:
        if os.path.exists(dir_to_clean):
            shutil.rmtree(dir_to_clean)
            print(f"   Supprimé: {dir_to_clean}")
    
    # Création du fichier spec pour PyInstaller
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['slim_pptx.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('logging_config.py', '.'),
        ('env_config.json', '.'),
        ('tkdnd2.8', 'tkdnd2.8'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL.ImageDraw',
        'pptx',
        'pptx.presentation',
        'pptx.slide',
        'pptx.shapes',
        'xml.etree.ElementTree',
        'zipfile',
        'zlib',
        'io',
        'tempfile',
        'shutil',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Slim_PPTX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
'''
    
    with open('slim_pptx.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ Fichier spec créé")
    
    # Build avec PyInstaller
    if not run_command("pyinstaller slim_pptx.spec", "Build de l'application"):
        print("❌ Échec du build")
        return False
    
    # Création du dossier portable
    portable_dir = "Slim_PPTX_Portable"
    if os.path.exists(portable_dir):
        shutil.rmtree(portable_dir)
    
    os.makedirs(portable_dir)
    
    # Copie de l'exécutable
    exe_source = "dist/Slim_PPTX.exe"
    exe_dest = os.path.join(portable_dir, "Slim_PPTX.exe")
    
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, exe_dest)
        print(f"✅ Exécutable copié: {exe_dest}")
    else:
        print(f"❌ Exécutable non trouvé: {exe_source}")
        return False
    
    # Création du README portable
    readme_content = """# Slim PPTX - Version Portable

## Utilisation

1. Double-cliquez sur `Slim_PPTX.exe` pour lancer l'application
2. Glissez-déposez votre fichier PowerPoint dans la fenêtre
3. Suivez les instructions à l'écran pour optimiser votre fichier

## Fonctionnalités

- Analyse des images dans les présentations PowerPoint
- Détection des images non utilisées
- Création de versions allégées
- Gestion des images rognées
- Interface graphique intuitive

## Support

Pour toute question ou problème, consultez la documentation complète.

## Version

Cette version portable ne nécessite aucune installation.
"""
    
    with open(os.path.join(portable_dir, "README.txt"), 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print("✅ README créé")
    
    # Création du fichier de lancement batch (optionnel)
    batch_content = """@echo off
echo Lancement de Slim PPTX...
start Slim_PPTX.exe
"""
    
    with open(os.path.join(portable_dir, "Lancer_Slim_PPTX.bat"), 'w', encoding='utf-8') as f:
        f.write(batch_content)
    
    print("✅ Fichier batch de lancement créé")
    
    # Création de l'archive ZIP
    zip_filename = "Slim_PPTX_Portable.zip"
    print(f"\n📦 Création de l'archive {zip_filename}...")
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(portable_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, portable_dir)
                zipf.write(file_path, arcname)
                print(f"   Ajouté: {arcname}")
    
    print(f"✅ Archive créée: {zip_filename}")
    
    # Nettoyage des fichiers temporaires
    print("\n🧹 Nettoyage des fichiers temporaires...")
    for file_to_clean in ["slim_pptx.spec"]:
        if os.path.exists(file_to_clean):
            os.remove(file_to_clean)
            print(f"   Supprimé: {file_to_clean}")
    
    print("\n🎉 Version portable créée avec succès!")
    print(f"📁 Dossier: {portable_dir}")
    print(f"📦 Archive: {zip_filename}")
    print("\nVous pouvez maintenant distribuer le dossier ou l'archive.")
    
    return True

if __name__ == "__main__":
    try:
        success = create_portable_version()
        if success:
            print("\n✅ Build terminé avec succès!")
        else:
            print("\n❌ Build échoué!")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️ Build interrompu par l'utilisateur")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Erreur inattendue: {e}")
        sys.exit(1) 