# TKINTER-APP
Une petite application d'exploration de données développée avec Python et `tkinter` (interface fournie dans un notebook Jupyter).

**Description :**
- **But :** fournir une interface graphique simple pour charger, visualiser et exporter des données (prototype pédagogique pour le cours AS3).
- **Technos :** Python, tkinter, ttkbootstrap (optionnel), Jupyter Notebook.

**Fonctionnalités principales :**
- Chargement de fichiers via une interface graphique (dans le notebook).
- Visualisation basique et export de modèles/templates.

**Prérequis**
- Python 3.10 ou supérieur
- Jupyter Notebook / JupyterLab pour exécuter le notebook interactif
- Bibliothèques Python courantes (voir section Installation)

**Installation (rapide)**
1. Créez un environnement virtuel (recommandé) :

```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

2. Installez Jupyter et dépendances (exemple) :

```powershell
pip install --upgrade pip
pip install jupyter notebook pandas matplotlib
# ttkbootstrap est optionnel : pip install ttkbootstrap
```

Si un `requirements.txt` est ajouté au projet, vous pouvez lancer :

```powershell
pip install -r requirements.txt
```

**Lancer le projet**
- Ouvrir et exécuter le notebook interactif : [final.ipynb](final.ipynb)
	- Lancez `jupyter notebook` ou `jupyter lab`, puis ouvrez le fichier et exécutez les cellules.
- Le notebook contient l'interface `tkinter` et des exemples d'utilisation.

**Structure du dépôt**
- [final.ipynb](final.ipynb) : Notebook principal contenant l'interface et le prototype.
- [export_templates_masterclass.py](export_templates_masterclass.py) : Script utilitaire pour l'export (si utilisé séparément).
- [README.md](TKINTER-APP/README.md) : Ce document.

**Bonnes pratiques**
- Travailler dans un environnement virtuel pour isoler les dépendances.
- Lorsqu'on veut transformer le prototype du notebook en script exécutable, extraire les cellules contenant l'interface et les placer dans un fichier `main.py`.

**Prochaines étapes suggérées**
- Ajouter un `requirements.txt` listant les dépendances exactes.
- Extraire le code du notebook dans un script `app.py` pour lancer l'interface sans Jupyter.

**Contact / Auteurs**
- Projet réalisé par le groupe 2 (TP Python AS3). Pour questions, proposer un issue ou contact par messagerie du cours.

Bonne exploration !
