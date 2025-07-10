# Script PowerShell pour activer l'environnement virtuel Slim_pptx
# Usage: .\activate_env.ps1

# Vérifier si le fichier active_env.txt existe
if (Test-Path "active_env.txt") {
    $activeEnv = Get-Content "active_env.txt" -Raw
    $activeEnv = $activeEnv.Trim()
    
    Write-Host "Environnement actif détecté: $activeEnv" -ForegroundColor Green
    
    # Charger la configuration des environnements
    if (Test-Path "env_config.json") {
        $config = Get-Content "env_config.json" | ConvertFrom-Json
        
        if ($config.environments.$activeEnv) {
            $envName = $config.environments.$activeEnv.name
            $envPath = "venv\$envName"
            
            if (Test-Path $envPath) {
                Write-Host "Activation de l'environnement virtuel: $envName" -ForegroundColor Yellow
                
                # Activer l'environnement virtuel
                & "$envPath\Scripts\Activate.ps1"
                
                Write-Host "Environnement virtuel activé avec succès!" -ForegroundColor Green
                Write-Host "Vous pouvez maintenant exécuter: python slim_pptx.py" -ForegroundColor Cyan
            } else {
                Write-Host "Erreur: L'environnement virtuel '$envPath' n'existe pas." -ForegroundColor Red
                Write-Host "Exécutez d'abord: python env_manager.py" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Erreur: Environnement '$activeEnv' non trouvé dans la configuration." -ForegroundColor Red
        }
    } else {
        Write-Host "Erreur: Fichier env_config.json non trouvé." -ForegroundColor Red
    }
} else {
    Write-Host "Aucun environnement actif détecté." -ForegroundColor Yellow
    Write-Host "Exécutez d'abord: python env_manager.py" -ForegroundColor Cyan
} 