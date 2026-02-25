# Outil d'analyse OSCAN

Outil Python pour l'analyse des données de contrôles OSCAN (Office Français de la Biodiversité), avec :

- une interface graphique (`analyse_oscean_gui.py`) ;
- un script d'analyse en ligne de commande (`analyse_oscean.py`) ;
- des scripts utilitaires (génération de rapports, tests, etc.).

Ce dépôt contient **uniquement le code de l'outil**. Les **données sources** et les **résultats générés** ne doivent pas être ajoutés à Git : ils sont exclus via le fichier `.gitignore`.

---

## Installation

1. Cloner le dépôt :

```bash
git clone https://github.com/<votre-organisation-ou-compte>/OSCAN.git
cd OSCAN
```

2. Créer et activer un environnement virtuel (recommandé) :

```bash
python -m venv .venv
.venv\Scripts\activate  # sous Windows
```

3. Installer les dépendances :

```bash
pip install -r requirements.txt
```

---

## Utilisation

### Interface graphique

Voir la notice détaillée dans `README_GUI.md`.

En résumé, pour lancer l'interface graphique :

```bash
python analyse_oscean_gui.py
```

Les rapports sont générés dans le dossier `resultats/` (ignoré par Git).

### Ligne de commande

Le script `analyse_oscean.py` permet d'automatiser certaines opérations d'analyse sans interface graphique.  
Les paramètres exacts peuvent évoluer : se référer au code et / ou à la documentation spécifique si elle est ajoutée.

---

## Structure du dépôt

- `analyse_oscean_gui.py` : interface graphique principale.
- `analyse_oscean.py` : script d'analyse (mode script / batch).
- `rapport_pdf_oscean.py` : génération de rapports PDF.
- `test_gui.py` : scripts de test / développement de l'interface.
- `db_test/` : jeux de données de test (exemples internes).
- `requirements.txt` : dépendances Python.
- `README_GUI.md` : documentation détaillée de l'interface graphique.
- `resultats/` : dossier de sortie des analyses (créé à l'exécution, ignoré par Git).

---

## Données et résultats (non versionnés)

Par principe, **aucun fichier de données opérationnelles** ni **fichier de résultat** ne doit être committé dans le dépôt.

Le fichier `.gitignore` exclut notamment :

- les dossiers de données possibles (`donnees/`, `data/`, `sources/`) ;
- le dossier des résultats `resultats/` ;
- les fichiers temporaires, caches Python et environnements virtuels.

---

## Licence

Ce projet est distribué sous la **Licence Apache 2.0**.

Les conditions complètes de licence se trouvent dans le fichier `LICENSE`.  
Les éventuelles mentions légales supplémentaires sont indiquées dans le fichier `NOTICE`.

---

## Auteur

- **Auteur** : Aguirre MAURIN  
- **Service** : Service départemental de la Côte d'Or  
- **Organisme** : Office Français de la Biodiversité  
- **Courriel** : aguirre.maurin@ofb.gouv.fr

