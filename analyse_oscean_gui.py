"""Interface graphique pour l'outil d'analyse OSCAN (PySide6).

Les gros fichiers Excel génèrent de nombreux UserWarning d'openpyxl
sur des dates invalides. On les masque ici pour éviter l'impression
de plantage dans la console.
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional
import warnings

import pandas as pd

from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QIcon, QFont, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QCheckBox,
    QGroupBox,
    QTabWidget,
    QTextEdit,
    QComboBox,
    QProgressBar,
    QMessageBox,
    QFileDialog,
    QHeaderView,
    QSpinBox,
    QListWidget,
    QFrame,
    QSplitter,
    QSizePolicy,
    QScrollArea,
    QDialog,
)

# Importer les fonctions du script principal
try:
    from analyse_oscean import (
        trouver_fichiers_candidats,
        selectionner_plus_recents_par_extension,
        charger_fichier,
        trouver_colonnes,
        generer_tableaux,
        generer_tableaux_generiques,
        remplacer_valeurs_vides_tableau,
        ajouter_ligne_total,
        filtrer_cote_d_or,
        exporter_rapport_excel,
        corriger_encodage_texte,
        charger_sources_fixes,
        enrichir_avec_sources_fixes,
        SUPPORTED_EXT,
    )
except ImportError as e:
    print(f"ERREUR: Impossible d'importer le module analyse_oscean")
    print(f"Erreur: {e}")
    print("\nAssurez-vous que le fichier analyse_oscean.py est dans le même dossier.")
    input("Appuyez sur Entrée pour fermer...")
    sys.exit(1)

# Masquer les avertissements verbeux d'openpyxl sur les dates hors limites
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    module="openpyxl.worksheet._reader",
)


class AnalyseThread(QThread):
    """Thread pour exécuter l'analyse en arrière-plan"""
    progress = Signal(str)
    finished = Signal(dict, str)
    error = Signal(str)

    def __init__(self, fichiers_selectionnes, options):
        super().__init__()
        self.fichiers_selectionnes = fichiers_selectionnes
        self.options = options

    def run(self):
        try:
            self.progress.emit("Chargement des fichiers...")
            dataframes = []
            resume_sources = []

            for ext, path in self.fichiers_selectionnes:
                try:
                    df_tmp = charger_fichier(path)
                    if df_tmp.empty:
                        continue
                    df_tmp = df_tmp.copy()
                    df_tmp["__source_fichier"] = os.path.basename(path)
                    df_tmp["__source_extension"] = ext
                    dataframes.append(df_tmp)
                    resume_sources.append({
                        "Fichier": os.path.basename(path),
                        "Chemin_complet": path,
                        "Extension": ext,
                        "Nb_lignes": len(df_tmp),
                        "Nb_colonnes": len(df_tmp.columns),
                    })
                except Exception as e:
                    self.progress.emit(f"Erreur sur {os.path.basename(path)}: {str(e)}")

            if not dataframes:
                self.error.emit("Aucun fichier valide chargé")
                return

            self.progress.emit("Combinaison des données...")
            df_total = pd.concat(dataframes, ignore_index=True)

            # Filtrer Côte-d'Or si demandé
            if self.options.get("filtre_cote_dor", False):
                self.progress.emit("Filtrage Côte-d'Or...")
                df_total = filtrer_cote_d_or(df_total)

            # Enrichir avec les sources fixes (NATINF, communes TUB, ...)
            try:
                sources_fixes = charger_sources_fixes()
                if sources_fixes:
                    self.progress.emit("Enrichissement avec les sources fixes (NATINF, communes TUB)...")
                    df_total = enrichir_avec_sources_fixes(df_total, sources_fixes)
            except Exception as e:
                self.progress.emit(f"Attention : enrichissement avec les sources fixes impossible ({type(e).__name__}: {e})")

            # Appliquer les filtres métier
            if self.options.get("filtre_annee"):
                annee = self.options["filtre_annee"]
                # Chercher une colonne de date
                for col in df_total.columns:
                    if "date" in str(col).lower() or "annee" in str(col).lower():
                        try:
                            df_total[col] = pd.to_datetime(df_total[col], errors="coerce")
                            df_total = df_total[df_total[col].dt.year == annee]
                            break
                        except:
                            pass

            if self.options.get("filtre_domaine"):
                colonnes = trouver_colonnes(df_total)
                if "domaine" in colonnes:
                    domaines = self.options["filtre_domaine"]
                    df_total = df_total[df_total[colonnes["domaine"]].isin(domaines)]

            if self.options.get("filtre_theme"):
                colonnes = trouver_colonnes(df_total)
                if "theme" in colonnes:
                    themes = self.options["filtre_theme"]
                    df_total = df_total[df_total[colonnes["theme"]].isin(themes)]

            # Filtrer les types d'usagers vides si demandé
            if self.options.get("exclure_usagers_vides", False):
                colonnes = trouver_colonnes(df_total)
                if "type_usage" in colonnes:
                    col_type = colonnes["type_usage"]
                    df_total = df_total[df_total[col_type].notna()]
                    df_total = df_total[df_total[col_type].astype(str).str.strip() != ""]

            self.progress.emit("Génération des tableaux...")
            colonnes = trouver_colonnes(df_total)
            try:
                resultats = generer_tableaux(df_total, colonnes)
            except Exception as e:
                self.progress.emit(f"Erreur lors de la génération des tableaux: {str(e)}")
                # Essayer avec les tableaux génériques en fallback
                resultats = generer_tableaux_generiques(df_total)

            if not resultats:
                resultats = generer_tableaux_generiques(df_total)

            # Ajouter résumé des sources
            resultats["Resume_sources"] = pd.DataFrame(resume_sources)

            # Journal : liste des tableaux générés (utile avec 2+ sources)
            self.progress.emit("Tableaux générés : " + ", ".join(resultats.keys()))
            # Ne pas exporter les rapports dans le thread : laisser le choix des graphiques à l'utilisateur
            self.finished.emit(resultats, "")

        except Exception as e:
            self.error.emit(f"Erreur lors de l'analyse: {str(e)}")


class FenetreResumeDialog(QDialog):
    """Fenêtre secondaire pour consulter le Résumé, l'Aperçu et le Journal."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Résumé / Aperçu / Journal")
        self.setMinimumSize(700, 450)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        tabs = QTabWidget()

        # Onglet Résumé
        tab_resume = QWidget()
        layout_resume = QVBoxLayout(tab_resume)
        layout_resume.setContentsMargins(5, 5, 5, 5)
        self.table_resume = QTableWidget()
        self.table_resume.setColumnCount(6)
        self.table_resume.setHorizontalHeaderLabels(
            ["Fichier", "Type", "Lignes", "Colonnes", "Période", "État"]
        )
        self.table_resume.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_resume.setMinimumHeight(200)
        self.table_resume.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout_resume.addWidget(self.table_resume)
        tabs.addTab(tab_resume, "Résumé")

        # Onglet Aperçu
        tab_apercu = QWidget()
        layout_apercu = QVBoxLayout(tab_apercu)
        layout_apercu.setContentsMargins(5, 5, 5, 5)
        layout_apercu.setSpacing(10)
        self.table_apercu = QTableWidget()
        self.table_apercu.setMinimumHeight(200)
        self.table_apercu.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout_apercu.addWidget(self.table_apercu)
        btn_actualiser = QPushButton("Actualiser l'aperçu")
        btn_actualiser.setMinimumHeight(34)
        if parent and hasattr(parent, "actualiser_apercu"):
            btn_actualiser.clicked.connect(parent.actualiser_apercu)
        layout_apercu.addWidget(btn_actualiser)
        tabs.addTab(tab_apercu, "Aperçu")

        # Onglet Journal
        tab_journal = QWidget()
        layout_journal = QVBoxLayout(tab_journal)
        layout_journal.setContentsMargins(5, 5, 5, 5)
        self.journal = QTextEdit()
        self.journal.setReadOnly(True)
        self.journal.setFont(QFont("Consolas", 9))
        self.journal.setMinimumHeight(200)
        self.journal.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout_journal.addWidget(self.journal)
        tabs.addTab(tab_journal, "Journal")

        layout.addWidget(tabs)


class GraphConfigDialog(QDialog):
    """Fenêtre de configuration des graphiques (choix des tableaux et du style)."""

    def __init__(self, noms_tableaux: list[str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Configuration des graphiques")
        self.setMinimumSize(600, 400)
        self.noms_tableaux = noms_tableaux
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        label_info = QLabel(
            "Sélectionnez les tableaux pour lesquels un graphique sera généré,\n"
            "et choisissez le style de graphique souhaité."
        )
        label_info.setWordWrap(True)
        layout.addWidget(label_info)

        self.table = QTableWidget(len(self.noms_tableaux), 3)
        self.table.setHorizontalHeaderLabels(["Inclure", "Tableau", "Style de graphique"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)

        styles = [
            ("bar", "Histogramme (barres)"),
            ("pie", "Camembert"),
            ("line", "Courbe"),
        ]

        for row, nom in enumerate(self.noms_tableaux):
            # Proposition par défaut : inclure les tableaux de type "Nombre de contrôles" ou "Conformité"
            lower = str(nom).lower()
            inclure_par_defaut = lower.startswith("nombre de contrôles") or lower.startswith("conformité")

            chk = QCheckBox()
            chk.setChecked(inclure_par_defaut)
            self.table.setCellWidget(row, 0, chk)

            item_nom = QTableWidgetItem(str(nom))
            item_nom.setFlags(item_nom.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, item_nom)

            combo = QComboBox()
            for code, label in styles:
                combo.addItem(label, userData=code)
            # Style par défaut : histogramme
            combo.setCurrentIndex(0)
            self.table.setCellWidget(row, 2, combo)

        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Boutons OK / Annuler
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_ok = QPushButton("Valider")
        btn_cancel = QPushButton("Annuler")
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_ok)
        layout.addLayout(btn_layout)

    def get_config(self) -> dict[str, str]:
        """Retourne un dict {nom_tableau: style_code} pour les tableaux cochés."""
        config: dict[str, str] = {}
        for row, nom in enumerate(self.noms_tableaux):
            chk = self.table.cellWidget(row, 0)
            combo = self.table.cellWidget(row, 2)
            if isinstance(chk, QCheckBox) and chk.isChecked() and isinstance(combo, QComboBox):
                style_code = combo.currentData() or "bar"
                config[nom] = style_code
        return config


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        try:
            self.settings = QSettings("OSCAN", "AnalyseTool")
            self.fichiers_candidats = []
            self.thread_analyse: Optional[AnalyseThread] = None
            self.derniere_options: dict | None = None
            self.dernier_resultats: dict | None = None
            self.init_ui()
        except Exception as e:
            print(f"ERREUR lors de l'initialisation de MainWindow:")
            print(f"{type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            raise

    def init_ui(self):
        self.setWindowTitle("Outil d'analyse OSCAN")
        # Taille minimale raisonnable pour tenir sur un écran standard (fenêtre redimensionnable)
        self.setMinimumSize(1000, 650)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # En-tête avec logo
        header_layout = self.create_header()
        main_layout.addLayout(header_layout)

        # Zone principale: Layout horizontal avec 2 colonnes (Sélection | Filtres + Résumé)
        zone_principale_layout = QHBoxLayout()
        zone_principale_layout.setSpacing(15)
        
        # Colonne gauche: Sélection des données
        zone_selection = self.create_zone_selection()
        zone_principale_layout.addWidget(zone_selection, 2)  # stretch=2 (plus large)
        
        # Colonne droite: Filtres métier uniquement (Résumé/Aperçu/Journal dans une fenêtre dédiée)
        zone_filtres = self.create_zone_filtres()
        zone_filtres.setMinimumHeight(320)
        zone_filtres.setMinimumWidth(300)
        
        zone_principale_layout.addWidget(zone_filtres, 1)  # stretch=1 pour la colonne droite
        
        main_layout.addLayout(zone_principale_layout, 1)  # stretch=1 pour prendre l'espace disponible

        # Bande horizontale : Paramètres (sous les autres zones)
        zone_parametres = self.create_zone_parametres()
        main_layout.addWidget(zone_parametres)

        # Barre d'action en bas
        barre_action = self.create_barre_action()
        # create_barre_action retourne un QHBoxLayout, on l'ajoute donc comme layout
        main_layout.addLayout(barre_action)

        # Charger le dernier dossier utilisé (sans scanner automatiquement)
        dernier_dossier = self.settings.value("dernier_dossier", "")
        if dernier_dossier and os.path.isdir(dernier_dossier):
            self.dossier_input.setText(dernier_dossier)
            # Ne pas scanner automatiquement pour accélérer le lancement
            # L'utilisateur cliquera sur "Scanner" quand il sera prêt

        # Fenêtre secondaire Résumé / Aperçu / Journal (créée mais non affichée)
        self.fenetre_resume = FenetreResumeDialog(self)

        # Restaurer taille et position de la fenêtre (comme l'Explorateur Windows)
        geom = self.settings.value("geometry")
        if geom is not None:
            self.restoreGeometry(geom)
        state = self.settings.value("windowState")
        if state is not None:
            self.restoreState(state)

    def closeEvent(self, event):
        """Sauvegarde taille et position de la fenêtre à la fermeture."""
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        if hasattr(self, "fenetre_resume") and self.fenetre_resume.isVisible():
            self.fenetre_resume.close()
        event.accept()

    def create_header(self):
        """Crée l'en-tête avec logo et titre"""
        layout = QHBoxLayout()
        
        # Logo
        logo_path = os.path.join(os.path.dirname(__file__), "logo-ofb-intranet.png")
        if os.path.exists(logo_path):
            logo_label = QLabel()
            pixmap = QPixmap(logo_path)
            pixmap = pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            layout.addWidget(logo_label)
        else:
            # Placeholder si logo absent
            layout.addWidget(QLabel("LOGO"))

        layout.addSpacing(15)

        # Titre
        titre_label = QLabel("Outil d'analyse OSCAN")
        font = QFont()
        font.setPointSize(18)
        font.setBold(True)
        titre_label.setFont(font)
        titre_label.setStyleSheet("color: #005E8F;")
        layout.addWidget(titre_label)

        layout.addStretch()

        # Info utilisateur/version (optionnel)
        info_label = QLabel(f"Version 1.0\n{datetime.now().strftime('%d/%m/%Y')}")
        info_label.setStyleSheet("color: #666; font-size: 9pt;")
        layout.addWidget(info_label)

        # Bouton Quitter avec icône
        btn_quitter = QPushButton("Quitter")
        btn_quitter.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_DialogCloseButton))
        btn_quitter.setStyleSheet("padding: 5px 15px;")
        btn_quitter.clicked.connect(self.close)
        layout.addWidget(btn_quitter)

        return layout

    def create_zone_selection(self):
        """Zone de sélection des fichiers - Réorganisée pour meilleure lisibilité"""
        group = QGroupBox("Sélection des données")
        layout = QVBoxLayout(group)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 20, 15, 15)

        # Sélection du dossier
        dossier_layout = QHBoxLayout()
        dossier_layout.setSpacing(10)
        dossier_label = QLabel("Dossier source:")
        dossier_label.setMinimumWidth(100)
        self.dossier_input = QLineEdit()
        self.dossier_input.setPlaceholderText("Sélectionnez un dossier...")
        btn_parcourir = QPushButton("Parcourir...")
        btn_parcourir.setMinimumWidth(100)
        btn_parcourir.clicked.connect(self.choisir_dossier)
        btn_scanner = QPushButton("Scanner")
        btn_scanner.setMinimumWidth(100)
        btn_scanner.clicked.connect(self.scanner_dossier)
        dossier_layout.addWidget(dossier_label)
        dossier_layout.addWidget(self.dossier_input, 1)
        dossier_layout.addWidget(btn_parcourir)
        dossier_layout.addWidget(btn_scanner)
        layout.addLayout(dossier_layout)

        # Tableau des fichiers : peut rétrécir en hauteur pour éviter le chevauchement quand la fenêtre n'est pas maximisée
        self.table_fichiers = QTableWidget()
        self.table_fichiers.setColumnCount(6)
        self.table_fichiers.setHorizontalHeaderLabels(
            ["Inclure", "Type", "Nom", "Date (nom)", "Date fichier", "Taille"]
        )
        self.table_fichiers.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Colonne "Nom" : largeur min pour éviter la troncature des noms de fichiers
        self.table_fichiers.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table_fichiers.horizontalHeader().setMinimumSectionSize(80)
        self.table_fichiers.setColumnWidth(2, 220)
        self.table_fichiers.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_fichiers.setMinimumHeight(80)  # hauteur min faible pour permettre le rétrécissement
        self.table_fichiers.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.table_fichiers, 1)  # stretch=1 : le tableau absorbe l'espace variable

        # Boutons d'action côte à côte, cadres compacts
        btn_label = QLabel("Actions:")
        btn_label.setStyleSheet("font-weight: bold; color: #005E8F; margin-top: 10px;")
        layout.addWidget(btn_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        
        btn_tout_cocher = QPushButton("Tout cocher")
        btn_tout_cocher.setFixedHeight(28)
        btn_tout_cocher.setMaximumWidth(110)
        btn_tout_cocher.setStyleSheet("padding: 2px 8px; min-width: 0;")
        btn_tout_cocher.clicked.connect(lambda: self.cocher_fichiers(True))
        
        btn_tout_decocher = QPushButton("Tout décocher")
        btn_tout_decocher.setFixedHeight(28)
        btn_tout_decocher.setMaximumWidth(120)
        btn_tout_decocher.setStyleSheet("padding: 2px 8px; min-width: 0;")
        btn_tout_decocher.clicked.connect(lambda: self.cocher_fichiers(False))
        
        btn_plus_recents = QPushButton("Plus récents par type")
        btn_plus_recents.setFixedHeight(28)
        btn_plus_recents.setMaximumWidth(160)
        btn_plus_recents.setStyleSheet("padding: 2px 8px; min-width: 0;")
        btn_plus_recents.clicked.connect(self.selectionner_plus_recents)
        
        btn_layout.addWidget(btn_tout_cocher)
        btn_layout.addWidget(btn_tout_decocher)
        btn_layout.addWidget(btn_plus_recents)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Options de filtrage - simplifié
        options_label = QLabel("Options:")
        options_label.setStyleSheet("font-weight: bold; color: #005E8F; margin-top: 15px;")
        layout.addWidget(options_label)
        
        options_layout = QVBoxLayout()
        options_layout.setSpacing(10)
        
        self.check_cote_dor = QCheckBox("Limiter aux données du département de la Côte-d'Or (21)")
        self.check_cote_dor.setChecked(True)
        self.check_usagers_vides = QCheckBox("Exclure les contrôles avec type d'usager vide")
        self.check_usagers_vides.setChecked(True)
        
        options_layout.addWidget(self.check_cote_dor)
        options_layout.addWidget(self.check_usagers_vides)
        layout.addLayout(options_layout)

        return group

    def create_zone_parametres(self):
        """Zone des paramètres en bande horizontale (Encodage | Sorties | Nom du rapport)."""
        group = QGroupBox("Paramètres")
        layout = QHBoxLayout(group)
        layout.setSpacing(25)
        layout.setContentsMargins(15, 12, 15, 12)

        # Encodage
        encodage_label = QLabel("Encodage:")
        encodage_label.setStyleSheet("font-weight: bold; color: #005E8F;")
        layout.addWidget(encodage_label)
        self.radio_encodage_auto = QCheckBox("Auto (recommandé)")
        self.radio_encodage_auto.setChecked(True)
        layout.addWidget(self.radio_encodage_auto)

        # Séparateur vertical
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.VLine)
        sep1.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(sep1)

        # Sorties
        sorties_label = QLabel("Sorties:")
        sorties_label.setStyleSheet("font-weight: bold; color: #005E8F;")
        layout.addWidget(sorties_label)
        self.check_excel = QCheckBox("Générer rapport Excel")
        self.check_excel.setChecked(True)
        self.check_pdf = QCheckBox("Générer rapport PDF")
        self.check_pdf.setChecked(True)
        self.check_csv = QCheckBox("Générer rapport CSV")
        layout.addWidget(self.check_excel)
        layout.addWidget(self.check_pdf)
        layout.addWidget(self.check_csv)

        # Séparateur vertical
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(sep2)
        
        # Nom du rapport
        nom_label = QLabel("Nom de base du rapport:")
        nom_label.setStyleSheet("color: #333;")
        layout.addWidget(nom_label)
        # Préfixe par défaut : "oscan_" (donne oscan_YYYYMMDD_HHMMSS.ext)
        self.nom_rapport_input = QLineEdit("oscan_")
        self.nom_rapport_input.setMinimumWidth(180)
        layout.addWidget(self.nom_rapport_input)
        
        layout.addStretch()

        return group

    def create_zone_filtres(self):
        """Zone Filtres métier : largeur min, listes avec retours à la ligne pour une lecture optimale."""
        group = QGroupBox("Filtres métier")
        group.setMinimumWidth(280)
        layout = QVBoxLayout(group)
        layout.setSpacing(10)
        layout.setContentsMargins(12, 16, 12, 12)
        
        # Ligne année + bouton actualiser
        annee_label = QLabel("Filtre par année:")
        annee_label.setStyleSheet("font-weight: bold; color: #005E8F;")
        annee_label.setMinimumHeight(annee_label.fontMetrics().height() + 2)
        layout.addWidget(annee_label)
        annee_layout = QHBoxLayout()
        annee_layout.setSpacing(10)
        self.check_annee = QCheckBox("Activer")
        self.spin_annee = QSpinBox()
        self.spin_annee.setRange(2000, 2100)
        self.spin_annee.setValue(datetime.now().year)
        self.spin_annee.setMinimumHeight(26)
        annee_layout.addWidget(self.check_annee)
        annee_layout.addWidget(QLabel("Année:"))
        annee_layout.addWidget(self.spin_annee)
        annee_layout.addStretch()
        layout.addLayout(annee_layout)

        # ----- Zone 1 : Champs accessibles (liste des champs selon les sources sélectionnées) -----
        label_champs = QLabel("Champs accessibles (selon les sources sélectionnées):")
        label_champs.setStyleSheet("font-weight: bold; color: #005E8F;")
        layout.addWidget(label_champs)
        btn_actualiser_champs = QPushButton("Actualiser les champs")
        btn_actualiser_champs.setMinimumHeight(28)
        btn_actualiser_champs.clicked.connect(self.charger_champs_accessibles)
        layout.addWidget(btn_actualiser_champs)
        self.list_champs_accessibles = QListWidget()
        self.list_champs_accessibles.setMinimumWidth(260)
        self.list_champs_accessibles.setMinimumHeight(60)
        self.list_champs_accessibles.setMaximumHeight(100)
        self.list_champs_accessibles.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        scroll_champs = QScrollArea()
        scroll_champs.setWidget(self.list_champs_accessibles)
        scroll_champs.setWidgetResizable(True)
        scroll_champs.setMaximumHeight(110)
        scroll_champs.setFrameShape(QFrame.Shape.NoFrame)
        scroll_champs.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_champs.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        layout.addWidget(scroll_champs)

        sep_champs = QFrame()
        sep_champs.setFrameShape(QFrame.Shape.HLine)
        sep_champs.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(sep_champs)

        # ----- Zone 2 : Champs utilisés comme filtres (sélection par l'utilisateur) -----
        label_filtres = QLabel("Champs utilisés comme filtres :")
        label_filtres.setStyleSheet("font-weight: bold; color: #005E8F;")
        layout.addWidget(label_filtres)
        aide_filtres = QLabel("Astuce : double-cliquez sur un champ de la liste du dessus pour l'ajouter ici ;\n"
                              "double-cliquez sur un champ de cette liste pour le retirer des filtres.")
        aide_filtres.setStyleSheet("font-size: 8pt; color: #666666;")
        layout.addWidget(aide_filtres)

        self.list_champs_filtres = QListWidget()
        # Dans cette liste, la sélection visuelle n'est pas obligatoire : tous les éléments présents
        # seront utilisés comme champs de filtrage.
        self.list_champs_filtres.setSelectionMode(QListWidget.MultiSelection)
        self.list_champs_filtres.setMinimumWidth(260)
        self.list_champs_filtres.setMinimumHeight(60)
        self.list_champs_filtres.setMaximumHeight(100)
        self.list_champs_filtres.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        scroll_filtres = QScrollArea()
        scroll_filtres.setWidget(self.list_champs_filtres)
        scroll_filtres.setWidgetResizable(True)
        scroll_filtres.setMaximumHeight(110)
        scroll_filtres.setFrameShape(QFrame.Shape.NoFrame)
        scroll_filtres.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_filtres.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        layout.addWidget(scroll_filtres)

        # Gestion des ajouts / suppressions par double-clic
        self.list_champs_accessibles.itemDoubleClicked.connect(self._ajouter_champ_filtre)
        self.list_champs_filtres.itemDoubleClicked.connect(self._retirer_champ_filtre)

        layout.addStretch(1)
        return group

    def create_barre_action(self):
        """Barre d'action en bas - Améliorée"""
        layout = QHBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(10, 10, 10, 10)

        # Statut et progression
        self.label_statut = QLabel("Prêt")
        self.label_statut.setStyleSheet("color: #005E8F; font-weight: bold; font-size: 10pt;")
        self.label_statut.setMinimumWidth(120)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumWidth(300)
        self.progress_bar.setMinimumHeight(25)

        # Boutons compacts (cadre réduit)
        btn_ouvrir_resultats = QPushButton("Ouvrir résultats")
        btn_ouvrir_resultats.setFixedHeight(28)
        btn_ouvrir_resultats.setMaximumWidth(120)
        btn_ouvrir_resultats.setStyleSheet("padding: 2px 10px; min-width: 0;")
        btn_ouvrir_resultats.clicked.connect(self.ouvrir_dossier_resultats)
        
        self.btn_lancer = QPushButton("Lancer l'analyse")
        self.btn_lancer.setStyleSheet(
            "background-color: #005E8F; color: white; font-weight: bold; padding: 4px 14px; font-size: 9pt; min-width: 0;"
        )
        self.btn_lancer.setFixedHeight(28)
        self.btn_lancer.setMaximumWidth(140)
        self.btn_lancer.clicked.connect(self.lancer_analyse)

        btn_resume = QPushButton("Résumé / Aperçu / Journal")
        btn_resume.setFixedHeight(28)
        btn_resume.setMaximumWidth(180)
        btn_resume.setStyleSheet("padding: 2px 10px; min-width: 0;")
        btn_resume.clicked.connect(self.ouvrir_fenetre_resume)

        layout.addWidget(self.label_statut)
        layout.addWidget(self.progress_bar, 1)
        layout.addStretch()
        layout.addWidget(btn_resume)
        layout.addWidget(btn_ouvrir_resultats)
        layout.addWidget(self.btn_lancer)

        return layout

    def ouvrir_fenetre_resume(self):
        """Ouvre la fenêtre secondaire Résumé / Aperçu / Journal."""
        self.fenetre_resume.show()
        self.fenetre_resume.raise_()
        self.fenetre_resume.activateWindow()

    def choisir_dossier(self):
        dossier = QFileDialog.getExistingDirectory(self, "Sélectionner un dossier")
        if dossier:
            self.dossier_input.setText(dossier)
            self.settings.setValue("dernier_dossier", dossier)
            self.scanner_dossier()

    def scanner_dossier(self):
        dossier = self.dossier_input.text().strip()
        if not dossier or not os.path.isdir(dossier):
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un dossier valide.")
            return

        self.ajouter_journal(f"Scan du dossier: {dossier}")
        candidats = trouver_fichiers_candidats(dossier)
        # Afficher TOUS les fichiers trouvés, pas seulement les plus récents par extension
        # Trier par date décroissante pour avoir les plus récents en premier
        candidats_tries = sorted(candidats, key=lambda x: x[2], reverse=True)

        self.fichiers_candidats = []
        self.table_fichiers.setRowCount(0)

        for path, ext, ref_date in candidats_tries:
            row = self.table_fichiers.rowCount()
            self.table_fichiers.insertRow(row)

            # Case à cocher
            checkbox = QTableWidgetItem()
            checkbox.setCheckState(Qt.Checked)
            self.table_fichiers.setItem(row, 0, checkbox)

            # Type
            self.table_fichiers.setItem(row, 1, QTableWidgetItem(ext))

            # Nom
            self.table_fichiers.setItem(row, 2, QTableWidgetItem(os.path.basename(path)))

            # Date nom
            date_nom = extraire_date_nom_fichier(path)
            date_nom_str = date_nom.strftime("%Y-%m-%d") if date_nom else "N/A"
            self.table_fichiers.setItem(row, 3, QTableWidgetItem(date_nom_str))

            # Date fichier
            date_fichier = datetime.fromtimestamp(os.path.getmtime(path))
            self.table_fichiers.setItem(row, 4, QTableWidgetItem(date_fichier.strftime("%Y-%m-%d %H:%M")))

            # Taille
            taille = os.path.getsize(path)
            taille_mb = taille / (1024 * 1024)
            self.table_fichiers.setItem(row, 5, QTableWidgetItem(f"{taille_mb:.2f} MB"))

            self.fichiers_candidats.append((ext, path))

        self.ajouter_journal(f"{len(self.fichiers_candidats)} fichier(s) détecté(s)")
        
        # Charger les champs accessibles selon les fichiers détectés
        QApplication.processEvents()
        self.charger_champs_accessibles()

    def cocher_fichiers(self, cocher: bool):
        state = Qt.Checked if cocher else Qt.Unchecked
        for row in range(self.table_fichiers.rowCount()):
            item = self.table_fichiers.item(row, 0)
            if item:
                item.setCheckState(state)

    def selectionner_plus_recents(self):
        self.cocher_fichiers(False)
        # Cocher seulement les plus récents par type
        types_vus = {}
        for row in range(self.table_fichiers.rowCount()):
            type_item = self.table_fichiers.item(row, 1)
            date_item = self.table_fichiers.item(row, 3)
            if type_item and date_item:
                ext = type_item.text()
                date_str = date_item.text()
                if ext not in types_vus or date_str > types_vus[ext][1]:
                    types_vus[ext] = (row, date_str)

        for row, _ in types_vus.values():
            item = self.table_fichiers.item(row, 0)
            if item:
                item.setCheckState(Qt.Checked)

    def actualiser_apercu(self):
        """Met à jour l'onglet Aperçu avec un échantillon des données sélectionnées."""
        try:
            fichiers = self._fichiers_selectionnes()
            if not fichiers:
                self.ajouter_journal("Aperçu : aucun fichier sélectionné.")
                self.fenetre_resume.table_apercu.clear()
                self.fenetre_resume.table_apercu.setRowCount(0)
                self.fenetre_resume.table_apercu.setColumnCount(0)
                return

            # Charger un petit échantillon de chaque fichier sélectionné
            df_list = []
            max_lignes_par_fichier = 50
            for ext, path in fichiers:
                try:
                    df_tmp = charger_fichier(path)
                    if df_tmp.empty:
                        continue
                    df_tmp = df_tmp.head(max_lignes_par_fichier).copy()
                    df_tmp["__source_fichier"] = os.path.basename(path)
                    df_tmp["__source_extension"] = ext
                    df_list.append(df_tmp)
                except Exception as e:
                    self.ajouter_journal(
                        f"Aperçu : erreur lors du chargement de '{os.path.basename(path)}' ({type(e).__name__}: {e})"
                    )

            if not df_list:
                self.ajouter_journal("Aperçu : aucun échantillon exploitable (fichiers vides ou illisibles).")
                self.fenetre_resume.table_apercu.clear()
                self.fenetre_resume.table_apercu.setRowCount(0)
                self.fenetre_resume.table_apercu.setColumnCount(0)
                return

            df_total = pd.concat(df_list, ignore_index=True)

            # Remplir le QTableWidget d'aperçu
            table = self.fenetre_resume.table_apercu
            table.clear()
            table.setRowCount(len(df_total))
            table.setColumnCount(len(df_total.columns))
            table.setHorizontalHeaderLabels([str(c) for c in df_total.columns])

            for i in range(len(df_total)):
                for j, col in enumerate(df_total.columns):
                    val = df_total.iloc[i, j]
                    item = QTableWidgetItem("" if pd.isna(val) else str(val))
                    table.setItem(i, j, item)

            table.resizeColumnsToContents()
            self.ajouter_journal(
                f"Aperçu mis à jour : {len(df_total)} ligne(s), {len(df_total.columns)} colonne(s) "
                f"à partir de {len(df_list)} fichier(s)."
            )
        except Exception as e:
            self.ajouter_journal(
                f"Aperçu : erreur inattendue ({type(e).__name__}: {e})"
            )

    def ouvrir_dossier_resultats(self):
        dossier = os.path.join(os.path.dirname(__file__), "resultats")
        if os.path.exists(dossier):
            os.startfile(dossier)
        else:
            QMessageBox.information(self, "Information", "Le dossier de résultats n'existe pas encore.")

    def lancer_analyse(self):
        # Récupérer les fichiers sélectionnés
        fichiers_selectionnes = []
        for row in range(self.table_fichiers.rowCount()):
            item = self.table_fichiers.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                ext_item = self.table_fichiers.item(row, 1)
                nom_item = self.table_fichiers.item(row, 2)
                if ext_item and nom_item:
                    # Trouver le chemin complet
                    for ext, path in self.fichiers_candidats:
                        if os.path.basename(path) == nom_item.text():
                            fichiers_selectionnes.append((ext, path))
                            break

        if not fichiers_selectionnes:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner au moins un fichier.")
            return

        # Récupérer les options
        options = {
            "filtre_cote_dor": self.check_cote_dor.isChecked(),
            "exclure_usagers_vides": self.check_usagers_vides.isChecked(),
            "generer_excel": self.check_excel.isChecked(),
            "generer_pdf": getattr(self, "check_pdf", None).isChecked() if hasattr(self, "check_pdf") else True,
            "generer_csv": self.check_csv.isChecked(),
            "nom_rapport": self.nom_rapport_input.text() or "oscan_",
        }

        # Filtres métier
        if self.check_annee.isChecked():
            options["filtre_annee"] = self.spin_annee.value()

        # Zone 2 : tous les champs présents dans la liste sont considérés comme filtres,
        # pas uniquement ceux qui sont sélectionnés visuellement.
        champs_filtres = [
            self.list_champs_filtres.item(i).text()
            for i in range(self.list_champs_filtres.count())
        ]
        if champs_filtres:
            options["filtre_champs"] = champs_filtres

        # Mémoriser les options pour l'étape d'export
        self.derniere_options = options

        # Lancer l'analyse dans un thread
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indéterminé
        self.label_statut.setText("Analyse en cours...")
        if self.btn_lancer:
            self.btn_lancer.setEnabled(False)

        self.thread_analyse = AnalyseThread(fichiers_selectionnes, options)
        self.thread_analyse.progress.connect(self.ajouter_journal)
        self.thread_analyse.finished.connect(self.analyse_terminee)
        self.thread_analyse.error.connect(self.analyse_erreur)
        self.thread_analyse.start()

    def analyse_terminee(self, resultats, chemins):
        # L'analyse (chargement + tableaux) est terminée, on propose maintenant
        # à l'utilisateur de choisir quels tableaux auront des graphiques.
        self.progress_bar.setVisible(False)
        self.label_statut.setText("Analyse terminée (tableaux prêts)")
        if self.btn_lancer:
            self.btn_lancer.setEnabled(True)

        if not isinstance(resultats, dict) or not resultats:
            QMessageBox.warning(self, "Information", "Aucun tableau généré, rien à exporter.")
            return

        self.dernier_resultats = resultats

        # Construire la liste des tableaux proposés pour les graphiques
        noms_tableaux = [
            nom
            for nom, df in resultats.items()
            if isinstance(df, pd.DataFrame) and not df.empty and nom != "Resume_sources"
        ]
        if not noms_tableaux:
            # Pas de tableaux pertinents, on exporte directement sans graphiques
            self.exporter_rapports(resultats, {})
            return

        dialog = GraphConfigDialog(noms_tableaux, self)
        if dialog.exec() != QDialog.Accepted:
            self.ajouter_journal("Export annulé par l'utilisateur (configuration des graphiques).")
            return

        graph_config = dialog.get_config()
        self.exporter_rapports(resultats, graph_config)

    def analyse_erreur(self, message):
        self.progress_bar.setVisible(False)
        self.label_statut.setText("Erreur")
        if self.btn_lancer:
            self.btn_lancer.setEnabled(True)
        self.ajouter_journal(f"ERREUR: {message}")
        QMessageBox.critical(self, "Erreur", message)

    def ajouter_journal(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.fenetre_resume.journal.append(f"[{timestamp}] {message}")

    def exporter_rapports(self, resultats: dict, graph_config: dict[str, str]) -> None:
        """Exporte les rapports (Excel, PDF, CSV) en utilisant les dernières options et la configuration de graphiques."""
        options = self.derniere_options or {}

        dossier_resultats = os.path.join(os.path.dirname(__file__), "resultats")
        base_name = options.get("nom_rapport", "oscan_") or "oscan_"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        self.ajouter_journal("Export des rapports...")

        chemins_rapports: list[str] = []

        # Excel
        if options.get("generer_excel", True):
            chemin_excel = os.path.join(dossier_resultats, f"{base_name}{timestamp}.xlsx")
            os.makedirs(dossier_resultats, exist_ok=True)
            noms_feuilles_utilises: set[str] = set()
            with pd.ExcelWriter(chemin_excel, engine="openpyxl") as writer:
                for nom, df in resultats.items():
                    df_export = remplacer_valeurs_vides_tableau(df.copy())
                    safe_name = re.sub(r'[:\\\\/\\?\\*\\[\\]<>"]', "_", str(nom))
                    safe_name = safe_name.strip()
                    if not safe_name:
                        safe_name = "Feuille"
                    short_name = safe_name[:31]
                    sheet_name = short_name
                    compteur = 1
                    while sheet_name in noms_feuilles_utilises:
                        suffix = f"_{compteur}"
                        sheet_name = (short_name[: 31 - len(suffix)] + suffix)
                        compteur += 1
                    noms_feuilles_utilises.add(sheet_name)
                    df_export.to_excel(writer, sheet_name=sheet_name, index=False)
            chemins_rapports.append(chemin_excel)

        # PDF
        if options.get("generer_pdf", False):
            try:
                from rapport_pdf_oscean import generer_pdf_oscan

                chemin_pdf = generer_pdf_oscan(dossier_resultats, resultats, base_name, graph_config)
                chemins_rapports.append(chemin_pdf)
            except Exception as e:
                self.ajouter_journal(
                    f"Attention : échec de la génération du PDF ({type(e).__name__}: {e})"
                )

        # CSV
        if options.get("generer_csv", False):
            for nom, df in resultats.items():
                safe_name = re.sub(r'[:\\\\/\\?\\*\\[\\]<>"]', "_", str(nom))[:50]
                chemin_csv = os.path.join(
                    dossier_resultats, f"{base_name}{safe_name}_{timestamp}.csv"
                )
                os.makedirs(dossier_resultats, exist_ok=True)
                df_export = remplacer_valeurs_vides_tableau(df.copy())
                df_export.to_csv(chemin_csv, index=False, encoding="utf-8-sig")
                chemins_rapports.append(chemin_csv)

        chemins_str = "\n".join(chemins_rapports)
        self.ajouter_journal(f"Rapports générés:\n{chemins_str}")
        QMessageBox.information(self, "Succès", f"Analyse terminée.\n\nRapports:\n{chemins_str}")

    def _fichiers_selectionnes(self):
        """Retourne la liste (ext, path) des fichiers cochés dans le tableau."""
        selection = []
        for row in range(self.table_fichiers.rowCount()):
            item = self.table_fichiers.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                ext_item = self.table_fichiers.item(row, 1)
                nom_item = self.table_fichiers.item(row, 2)
                if ext_item and nom_item:
                    for ext, path in self.fichiers_candidats:
                        if os.path.basename(path) == nom_item.text():
                            selection.append((ext, path))
                            break
        return selection

    def charger_champs_accessibles(self):
        """Affiche la liste des champs accessibles selon les sources sélectionnées.

        La liste des champs filtres n'est PAS remplie automatiquement : elle est
        gérée par l'utilisateur (double-clic pour ajouter/retirer)."""
        try:
            fichiers = self._fichiers_selectionnes()
            if not fichiers:
                fichiers = self.fichiers_candidats[:5]  # fallback : premiers candidats
            if not fichiers:
                self.list_champs_accessibles.clear()
                self.list_champs_filtres.clear()
                return
            champs_set = set()
            for ext, path in fichiers:
                try:
                    # Réutiliser la logique robuste de chargement principale
                    df = charger_fichier(path)
                    if df.empty:
                        continue
                    champs_set.update(df.columns.astype(str))
                except Exception as e:
                    # On loggue dans le journal mais sans interrompre l'IHM
                    self.ajouter_journal(
                        f"Champs accessibles : erreur sur '{os.path.basename(path)}' "
                        f"({type(e).__name__}: {e})"
                    )
            champs_tries = sorted(champs_set)
            self.list_champs_accessibles.clear()
            self.list_champs_filtres.clear()
            for c in champs_tries:
                self.list_champs_accessibles.addItem(c)
            self.ajouter_journal(f"Champs accessibles: {len(champs_tries)} (sources: {len(fichiers)} fichier(s))")
        except Exception:
            # Ne pas remonter les erreurs à l'utilisateur ici
            pass

    def _ajouter_champ_filtre(self, item):
        """Ajoute un champ de la liste 'champs accessibles' vers la liste des filtres."""
        if not item:
            return
        texte = item.text()
        # Ne pas dupliquer un champ déjà présent dans la liste des filtres
        existants = {self.list_champs_filtres.item(i).text() for i in range(self.list_champs_filtres.count())}
        if texte and texte not in existants:
            self.list_champs_filtres.addItem(texte)

    def _retirer_champ_filtre(self, item):
        """Retire un champ de la liste des filtres (double-clic)."""
        if not item:
            return
        row = self.list_champs_filtres.row(item)
        if row >= 0:
            self.list_champs_filtres.takeItem(row)


def extraire_date_nom_fichier(filename: str) -> Optional[datetime]:
    """Helper pour extraire la date du nom de fichier"""
    import re
    base = os.path.basename(filename)
    match = re.search(r"(19|20)\d{6}", base)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(0), "%Y%m%d")
    except ValueError:
        return None


def main():
    try:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")  # Style moderne
        
        # Stylesheet personnalisé - Interface épurée et agréable
        app.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 10pt;
                border: 1px solid #C0C0C0;
                border-radius: 4px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: #FAFAFA;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #005E8F;
            }
            QPushButton {
                padding: 6px 16px;
                border-radius: 4px;
                background-color: #F5F5F5;
                border: 1px solid #D0D0D0;
                min-height: 25px;
            }
            QPushButton:hover {
                background-color: #E8E8E8;
                border: 1px solid #B0B0B0;
            }
            QPushButton:pressed {
                background-color: #D0D0D0;
            }
            QTableWidget {
                border: 1px solid #D0D0D0;
                gridline-color: #E5E5E5;
                background-color: white;
                alternate-background-color: #F9F9F9;
            }
            QTableWidget::item:selected {
                background-color: #8DBDD8;
                color: #000000;
            }
            QCheckBox {
                spacing: 5px;
                padding: 2px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1px solid #808080;
                border-radius: 3px;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #005E8F;
                border: 1px solid #005E8F;
            }
            QLineEdit {
                padding: 4px;
                border: 1px solid #C0C0C0;
                border-radius: 3px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #005E8F;
            }
            QListWidget {
                border: 1px solid #D0D0D0;
                border-radius: 3px;
                background-color: white;
            }
            QListWidget::item:selected {
                background-color: #8DBDD8;
            }
            QSpinBox {
                padding: 4px;
                border: 1px solid #C0C0C0;
                border-radius: 3px;
                background-color: white;
            }
        """)

        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"ERREUR lors du lancement de l'interface graphique:")
        print(f"{type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        input("Appuyez sur Entrée pour fermer...")
        sys.exit(1)


if __name__ == "__main__":
    main()
