# Interface Graphique - Outil d'analyse OSCAN

## Installation

1. Installer les dépendances :
```bash
pip install -r requirements.txt
```

## Lancement

Pour lancer l'interface graphique :
```bash
python analyse_oscean_gui.py
```

## Utilisation

### 1. Sélection des données
- Cliquez sur **"Parcourir..."** pour sélectionner le dossier contenant vos fichiers
- Cliquez sur **"Scanner"** pour détecter les fichiers les plus récents
- Cochez/décochez les fichiers à analyser dans le tableau
- Utilisez **"Ne garder que les plus récents par type"** pour une sélection automatique

### 2. Paramètres

#### Onglet "Encodage"
- Laisser sur **"Auto (recommandé)"** sauf problème spécifique

#### Onglet "Sorties"
- Cochez **"Générer rapport Excel"** (recommandé)
- Cochez **"Générer rapport CSV"** si besoin
- Modifiez le nom de base du rapport si nécessaire

#### Onglet "Filtres métier"
- **Filtre par année** : Activez et sélectionnez une année
- **Filtre par domaine** : Activez et sélectionnez un ou plusieurs domaines (multi-sélection)
- **Filtre par thème** : Activez et sélectionnez un ou plusieurs thèmes (multi-sélection)

### 3. Options de filtrage
- **Limiter aux données du département de la Côte-d'Or (21)** : Filtre automatique sur le code 21
- **Exclure les contrôles avec type d'usager vide** : Ignore les enregistrements sans type d'usager

### 4. Lancer l'analyse
- Cliquez sur **"Lancer l'analyse"**
- Suivez la progression dans l'onglet **"Journal"**
- Les rapports sont générés dans le dossier `resultats/`

### 5. Consulter les résultats
- Cliquez sur **"Ouvrir le dossier des résultats"** pour accéder aux fichiers générés
- Consultez l'onglet **"Résumé"** pour voir les statistiques des fichiers analysés

## Notes

- Les fichiers les plus récents sont détectés automatiquement selon la date dans le nom (format YYYYMMDD) ou la date de modification
- Le dossier `resultats` est automatiquement exclu de la recherche
- Les problèmes d'encodage sont corrigés automatiquement
- Les valeurs vides sont remplacées par "Information non rapportée"
