import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import venv

class EnvironmentManager:
    def __init__(self):
        self.config = self.load_config()
        self.root = None
        
    def load_config(self):
        """Charge la configuration des environnements."""
        try:
            with open('env_config.json', 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Erreur lors du chargement de la configuration: {e}")
            return None
            
    def check_python_version(self, required_version):
        """Vérifie si la version de Python est compatible."""
        current_version = sys.version_info
        required = tuple(map(int, required_version.split('.')))
        return current_version >= required
        
    def create_environment(self, env_name):
        """Crée un nouvel environnement virtuel."""
        env_path = os.path.join('venv', env_name)
        if not os.path.exists(env_path):
            print(f"Création de l'environnement {env_name}...")
            venv.create(env_path, with_pip=True)
        return env_path
        
    def install_dependencies(self, env_name, dependencies):
        """Installe les dépendances dans l'environnement spécifié."""
        env_path = os.path.join('venv', env_name)
        pip_path = os.path.join(env_path, 'Scripts', 'pip')
        
        # Création du fichier requirements temporaire
        temp_req = 'temp_requirements.txt'
        with open(temp_req, 'w') as f:
            f.write('\n'.join(dependencies))
            
        try:
            subprocess.run([pip_path, 'install', '-r', temp_req], check=True)
        finally:
            if os.path.exists(temp_req):
                os.remove(temp_req)
                
    def show_environment_selector(self):
        """Affiche une interface pour sélectionner l'environnement."""
        self.root = tk.Tk()
        self.root.title("Sélection de l'environnement")
        self.root.geometry("400x300")
        
        # Style
        style = ttk.Style()
        style.configure('TButton', padding=10)
        style.configure('TLabel', padding=5)
        
        # Titre
        title = ttk.Label(self.root, text="Choisissez votre environnement d'analyse", font=('Helvetica', 12))
        title.pack(pady=20)
        
        # Frame pour les boutons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(expand=True)
        
        # Boutons pour chaque environnement
        for env_key, env_info in self.config['environments'].items():
            btn = ttk.Button(
                button_frame,
                text=f"Analyseur {env_key.replace('_', ' ').title()}",
                command=lambda k=env_key: self.setup_environment(k)
            )
            btn.pack(pady=10, padx=20, fill='x')
            
        self.root.mainloop()
        
    def setup_environment(self, env_key):
        """Configure l'environnement sélectionné."""
        env_info = self.config['environments'][env_key]
        
        # Vérification de la version de Python
        if not self.check_python_version(env_info['python_version']):
            tk.messagebox.showerror(
                "Erreur",
                f"Version de Python requise: {env_info['python_version']}\n"
                f"Version actuelle: {sys.version.split()[0]}"
            )
            return
            
        # Création de l'environnement
        env_path = self.create_environment(env_info['name'])
        
        # Installation des dépendances
        self.install_dependencies(env_info['name'], env_info['dependencies'])
        
        # Sauvegarde de l'environnement actif
        with open('active_env.txt', 'w') as f:
            f.write(env_key)
            
        # Fermeture de la fenêtre de sélection
        if self.root:
            self.root.destroy()
            
        # Lancement de l'application
        self.launch_application(env_path)
        
    def launch_application(self, env_path):
        """Lance l'application dans l'environnement spécifié."""
        python_path = os.path.join(env_path, 'Scripts', 'python')
        subprocess.run([python_path, 'slim_pptx.py'])
        
if __name__ == "__main__":
    manager = EnvironmentManager()
    manager.show_environment_selector() 