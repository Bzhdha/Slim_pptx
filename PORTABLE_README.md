# Slim PPTX - Version Portable

## Création de la version portable

### Prérequis

- Python 3.8 ou supérieur
- Windows 10 ou supérieur
- Environnement virtuel configuré (optionnel mais recommandé)

### Méthodes de création

#### Méthode 1 : Script batch (recommandé)
```batch
build_portable.bat
```

#### Méthode 2 : Script PowerShell
```powershell
.\build_portable.ps1
```

#### Méthode 3 : Script Python direct
```bash
python build_portable.py
```

### Processus de build

Le script de build va :

1. **Vérifier les prérequis** : Python, PyInstaller
2. **Nettoyer** les builds précédents
3. **Créer un fichier spec** pour PyInstaller avec toutes les dépendances
4. **Compiler l'application** en exécutable autonome
5. **Organiser les fichiers** dans un dossier portable
6. **Créer une archive ZIP** pour la distribution

### Fichiers générés

Après le build, vous obtiendrez :

- `Slim_PPTX_Portable/` - Dossier contenant la version portable
  - `Slim_PPTX.exe` - Exécutable principal
  - `README.txt` - Instructions d'utilisation
  - `Lancer_Slim_PPTX.bat` - Script de lancement optionnel
- `Slim_PPTX_Portable.zip` - Archive pour la distribution

## Distribution de la version portable

### Contenu du package

La version portable contient tout ce qui est nécessaire pour exécuter l'application :

- **Exécutable autonome** : Inclut Python et toutes les dépendances
- **Bibliothèques requises** : tkinterdnd2, python-pptx, Pillow, etc.
- **Fichiers de configuration** : logging_config.py, env_config.json
- **Documentation** : README.txt avec instructions

### Avantages de la version portable

✅ **Aucune installation requise**  
✅ **Fonctionne sur n'importe quel PC Windows**  
✅ **Pas de conflit avec d'autres versions de Python**  
✅ **Distribution facile**  
✅ **Pas de droits administrateur nécessaires**  

### Utilisation

1. **Extraction** : Décompressez l'archive ZIP
2. **Lancement** : Double-cliquez sur `Slim_PPTX.exe`
3. **Utilisation** : Glissez-déposez vos fichiers PowerPoint

## Dépannage

### Problèmes courants

#### L'application ne se lance pas
- Vérifiez que votre antivirus ne bloque pas l'exécution
- Assurez-vous d'avoir les droits d'écriture dans le dossier
- Vérifiez que tous les fichiers sont présents

#### Erreur de dépendance manquante
- Reconstruisez la version portable avec `build_portable.py`
- Vérifiez que toutes les dépendances sont dans `requirements.txt`

#### Performance lente
- La première exécution peut être plus lente (décompression)
- Les exécutions suivantes seront plus rapides

### Logs et débogage

Si vous rencontrez des problèmes :

1. **Consultez les logs** : L'application crée des fichiers de log
2. **Mode verbose** : Utilisez `.\build_portable.ps1 -Verbose`
3. **Log de build** : Consultez `build_log.txt` après un build échoué

## Personnalisation

### Ajout d'une icône

Pour ajouter une icône personnalisée :

1. Placez votre fichier `.ico` dans le dossier du projet
2. Modifiez `build_portable.py` ligne 108 :
   ```python
   icon='votre_icone.ico',
   ```

### Modification du nom de l'exécutable

Modifiez la ligne 103 dans `build_portable.py` :
```python
name='VotreNom',
```

### Ajout de fichiers supplémentaires

Pour inclure d'autres fichiers dans la version portable, ajoutez-les dans la section `datas` du fichier spec (lignes 15-19).

## Support

Pour toute question concernant la version portable :

1. Consultez les logs de build
2. Vérifiez que tous les prérequis sont satisfaits
3. Testez sur un PC propre si possible
4. Consultez la documentation complète du projet 