#!/usr/bin/env python3
"""
Script de test pour vérifier si un fichier PowerPoint est valide
Usage: python test_pptx.py <chemin_vers_fichier.pptx>
"""

import os
import sys
import zipfile
from pptx import Presentation

def test_pptx_file(file_path):
    """Teste si un fichier PowerPoint est valide."""
    print(f"Test du fichier : {file_path}")
    print("=" * 50)
    
    # Vérification de l'existence
    if not os.path.exists(file_path):
        print("❌ ERREUR : Le fichier n'existe pas")
        return False
    
    if not os.path.isfile(file_path):
        print("❌ ERREUR : Le chemin ne correspond pas à un fichier")
        return False
    
    # Vérification de la taille
    file_size = os.path.getsize(file_path)
    print(f"📁 Taille du fichier : {file_size:,} octets ({file_size/1024/1024:.2f} MB)")
    
    if file_size == 0:
        print("❌ ERREUR : Le fichier est vide")
        return False
    
    # Vérification de l'extension
    if not file_path.lower().endswith('.pptx'):
        print("⚠️  ATTENTION : Le fichier n'a pas l'extension .pptx")
    
    # Test ZIP
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            print("✅ Fichier ZIP valide")
            
            # Vérification de la structure PPTX
            file_list = zip_ref.namelist()
            print(f"📋 Nombre de fichiers dans le ZIP : {len(file_list)}")
            
            # Fichiers essentiels pour un PPTX
            required_files = [
                'ppt/presentation.xml',
                'ppt/slides/slide1.xml',
                '[Content_Types].xml'
            ]
            
            missing_files = []
            for required_file in required_files:
                if required_file in file_list:
                    print(f"✅ {required_file}")
                else:
                    print(f"❌ {required_file} - MANQUANT")
                    missing_files.append(required_file)
            
            if missing_files:
                print(f"❌ ERREUR : Fichiers essentiels manquants : {missing_files}")
                return False
            
            # Afficher quelques fichiers pour diagnostic
            print("\n📂 Structure du fichier PPTX :")
            ppt_files = [f for f in file_list if f.startswith('ppt/')]
            for f in sorted(ppt_files)[:10]:  # Afficher les 10 premiers
                print(f"  {f}")
            if len(ppt_files) > 10:
                print(f"  ... et {len(ppt_files) - 10} autres fichiers")
                
    except zipfile.BadZipFile:
        print("❌ ERREUR : Le fichier n'est pas un fichier ZIP valide")
        return False
    except Exception as e:
        print(f"❌ ERREUR lors de la lecture du ZIP : {str(e)}")
        return False
    
    # Test avec python-pptx
    try:
        print("\n🔍 Test avec python-pptx...")
        prs = Presentation(file_path)
        print(f"✅ Présentation ouverte avec succès")
        print(f"📊 Nombre de diapositives : {len(prs.slides)}")
        
        # Informations sur les layouts
        layouts = prs.slide_layouts
        print(f"🎨 Nombre de layouts disponibles : {len(layouts)}")
        
        # Informations sur les masters
        masters = prs.slide_masters
        print(f"👑 Nombre de masters : {len(masters)}")
        
        return True
        
    except Exception as e:
        print(f"❌ ERREUR lors de l'ouverture avec python-pptx : {str(e)}")
        return False

def main():
    if len(sys.argv) != 2:
        print("Usage: python test_pptx.py <chemin_vers_fichier.pptx>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    if test_pptx_file(file_path):
        print("\n🎉 Le fichier PowerPoint est VALIDE !")
        sys.exit(0)
    else:
        print("\n💥 Le fichier PowerPoint est INVALIDE ou CORROMPU !")
        sys.exit(1)

if __name__ == "__main__":
    main() 