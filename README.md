# Slim PPTX

Application de réduction de taille des fichiers PowerPoint.

## Prérequis

- Python 3.6 ou supérieur
- Windows 10 ou supérieur

## Installation

### Version Standard
1. Téléchargez ou clonez ce dépôt sur votre ordinateur
2. Exécutez le fichier `setup.bat` en double-cliquant dessus
   - Ce script va :
     - Créer un environnement virtuel Python
     - Installer toutes les dépendances nécessaires
3. Attendez que l'installation soit terminée

### Version Portable
1. **Création** : Exécutez `build_portable.bat` ou `build_portable.ps1`
2. **Distribution** : Utilisez le dossier `Slim_PPTX_Portable` ou l'archive `Slim_PPTX_Portable.zip`
3. **Utilisation** : Double-cliquez sur `Slim_PPTX.exe` pour lancer l'application
4. **Aucune installation** n'est nécessaire sur la machine cible

## Utilisation

### Version Standard
1. Double-cliquez sur le fichier `run.bat` pour lancer l'application
2. L'interface graphique s'ouvrira automatiquement
3. Sélectionnez votre fichier PowerPoint à optimiser
4. Suivez les instructions à l'écran

### Version Portable
1. Double-cliquez sur `Slim_PPTX.exe` ou `Lancer_Slim_PPTX.bat`
2. L'interface graphique s'ouvrira automatiquement
3. Glissez-déposez votre fichier PowerPoint dans la fenêtre
4. Suivez les instructions à l'écran pour optimiser votre fichier

## Désinstallation

### Version Standard
Pour désinstaller l'application :
1. Supprimez simplement le dossier du projet
2. L'environnement virtuel et toutes les dépendances seront supprimés automatiquement

### Version Portable
Aucune désinstallation n'est nécessaire. Supprimez simplement le dossier `Slim_PPTX_Portable` ou l'archive pour retirer l'application.

## Dépannage

Si vous rencontrez des problèmes :

### Version Standard
- Assurez-vous que Python est bien installé sur votre système
- Vérifiez que vous avez les droits administrateur pour l'installation
- Relancez `setup.bat` si l'installation échoue

### Version Portable
- Assurez-vous que votre antivirus ne bloque pas l'exécution
- Vérifiez que tous les fichiers sont présents dans le dossier
- Consultez `PORTABLE_README.md` pour plus de détails sur la création
- Si le build échoue, consultez `build_log.txt` pour les détails

## Support

Pour toute question ou problème, n'hésitez pas à ouvrir une issue sur le dépôt. 