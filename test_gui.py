"""Script de test minimal pour vérifier que PySide6 fonctionne"""

import sys

try:
    from PySide6.QtWidgets import QApplication, QLabel, QWidget
    print("✓ PySide6 importé avec succès")
except ImportError as e:
    print(f"✗ ERREUR: PySide6 non disponible")
    print(f"   Erreur: {e}")
    print("\n   Installez PySide6 avec: python -m pip install --user PySide6")
    sys.exit(1)

try:
    app = QApplication(sys.argv)
    print("✓ QApplication créé avec succès")
    
    window = QWidget()
    window.setWindowTitle("Test PySide6")
    label = QLabel("Si tu vois cette fenêtre, PySide6 fonctionne !", window)
    window.resize(400, 100)
    window.show()
    print("✓ Fenêtre de test affichée")
    print("\nFermez la fenêtre pour terminer le test.")
    
    sys.exit(app.exec())
except Exception as e:
    print(f"✗ ERREUR lors du lancement:")
    print(f"   {type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()
    input("\nAppuyez sur Entrée pour fermer...")
    sys.exit(1)
