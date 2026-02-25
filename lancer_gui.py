"""Script de lancement de l'interface graphique OSCAN.

Usage (depuis la racine du projet) :
    python lancer_gui.py
    python3 lancer_gui.py
"""
import sys
from pathlib import Path

# S'assurer que le répertoire du script est dans le path
script_dir = Path(__file__).resolve().parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

if __name__ == "__main__":
    print("Lancement de l'interface graphique OSCAN...")
    import analyse_oscean_gui
    analyse_oscean_gui.main()
