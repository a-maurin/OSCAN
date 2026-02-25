import os
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd

# dbfread : import paresseux dans charger_fichier() pour accélérer le démarrage
HAS_DBFREAD: bool | None = None
DBF = None


def _charger_dbfread():
    """Charge dbfread à la demande (première lecture d'un fichier .dbf)."""
    global HAS_DBFREAD, DBF
    if HAS_DBFREAD is not None:
        return
    try:
        from dbfread import DBF as _DBF  # type: ignore
        DBF = _DBF
        HAS_DBFREAD = True
    except ImportError:
        HAS_DBFREAD = False
        DBF = None


# Extensions prises en charge
SUPPORTED_EXT = [".xlsx", ".csv", ".dbf"]


###############################################################################
# 1. Localisation des fichiers et choix des fichiers les plus récents
###############################################################################


def extraire_date_nom_fichier(filename: str) -> datetime | None:
    """
    Cherche une date au format YYYYMMDD dans le nom de fichier.
    Exemple: sd21_point_ctrl_20251231_wgs84.xlsx -> 2025-12-31
    """
    base = os.path.basename(filename)
    match = re.search(r"(19|20)\d{6}", base)  # 8 chiffres commençant par 19 ou 20
    if not match:
        return None
    try:
        return datetime.strptime(match.group(0), "%Y%m%d")
    except ValueError:
        return None


def trouver_fichiers_candidats(dossier_racine: str) -> list[tuple[str, str, datetime]]:
    """
    Parcourt dossier_racine et retourne une liste de tuples
    (chemin_complet, extension, date_reference)
    où date_reference est d'abord la date extraite du nom, sinon la date de modification.
    """
    candidats: list[tuple[str, str, datetime]] = []

    for root, dirs, files in os.walk(dossier_racine):
        # Exclure le dossier 'resultats' pour éviter de ré-analyser les rapports produits
        if "resultats" in dirs:
            dirs.remove("resultats")
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext not in SUPPORTED_EXT:
                continue

            path = os.path.join(root, f)
            date_nom = extraire_date_nom_fichier(f)
            if date_nom is not None:
                ref_date = date_nom
            else:
                # fallback : date de modification du fichier
                ts = os.path.getmtime(path)
                ref_date = datetime.fromtimestamp(ts)

            candidats.append((path, ext, ref_date))

    return candidats


def selectionner_plus_recents_par_extension(
    candidats: list[tuple[str, str, datetime]]
) -> dict[str, str]:
    """
    Pour chaque extension (.xlsx, .csv, .dbf), garde uniquement le fichier
    avec la date de référence la plus récente.
    Retourne un dict {extension: chemin_fichier}.
    """
    selection: dict[str, tuple[str, datetime]] = {}

    for path, ext, ref_date in candidats:
        if ext not in selection:
            selection[ext] = (path, ref_date)
        else:
            _, current_date = selection[ext]
            if ref_date > current_date:
                selection[ext] = (path, ref_date)

    return {ext: path for ext, (path, _) in selection.items()}


###############################################################################
# 2. Chargement des fichiers selon leur format
###############################################################################


def corriger_encodage_texte(serie: pd.Series) -> pd.Series:
    """
    Corrige les problèmes d'encodage courants où UTF-8 a été mal interprété comme Latin-1.
    Exemple : "ContrÃ'les" -> "Contrôles", "protÃ©gÃ©es" -> "protégées"
    """
    if not pd.api.types.is_string_dtype(serie) and serie.dtype != "object":
        return serie

    def corriger_valeur(val):
        if pd.isna(val) or not isinstance(val, str):
            return val
        try:
            # Si la chaîne contient des séquences typiques de mauvaise interprétation UTF-8->Latin-1
            # On essaie de la réencoder correctement
            if "Ã" in val or "â€™" in val or "â€œ" in val or "Ã©" in val or "Ã¨" in val:
                # Essayer de récupérer l'UTF-8 original
                # Encoder en latin-1 puis décoder en UTF-8
                val_bytes = val.encode("latin-1", errors="ignore")
                val_corrigee = val_bytes.decode("utf-8", errors="ignore")
                # Si la correction a fonctionné, retourner la valeur corrigée
                if val_corrigee != val and len(val_corrigee) > 0:
                    return val_corrigee
        except (UnicodeEncodeError, UnicodeDecodeError, AttributeError):
            pass
        return val

    return serie.map(corriger_valeur)


def corriger_encodage_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Applique la correction d'encodage sur toutes les colonnes texte du DataFrame.
    """
    df_corrige = df.copy()
    for col in df_corrige.columns:
        if pd.api.types.is_string_dtype(df_corrige[col]) or df_corrige[col].dtype == "object":
            df_corrige[col] = corriger_encodage_texte(df_corrige[col])
    return df_corrige


def charger_fichier(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        # Séparateur par défaut ; ajustable selon tes fichiers
        # Essayer plusieurs encodages
        encodages = ["utf-8", "cp1252", "latin-1", "iso-8859-1"]
        for enc in encodages:
            try:
                df = pd.read_csv(path, sep=";", encoding=enc, engine="python")
                df = corriger_encodage_dataframe(df)
                return df
            except (UnicodeDecodeError, UnicodeError):
                continue
        # Dernier essai avec virgule comme séparateur
        try:
            df = pd.read_csv(path, sep=",", encoding="utf-8", engine="python")
            df = corriger_encodage_dataframe(df)
            return df
        except Exception:
            raise ValueError(f"Impossible de lire le fichier CSV {path} avec les encodages testés")

    if ext in (".xls", ".xlsx"):
        df = pd.read_excel(path)
        df = corriger_encodage_dataframe(df)
        return df

    if ext == ".dbf":
        _charger_dbfread()
        if not HAS_DBFREAD or DBF is None:
            raise RuntimeError(
                "Lecture .dbf demandée mais le module 'dbfread' n'est pas installé.\n"
                "Installe-le avec par exemple : pip install dbfread"
            )
        # Essayer plusieurs encodages pour les fichiers DBF
        encodages = ["cp1252", "utf-8", "latin-1", "iso-8859-1"]
        for enc in encodages:
            try:
                table = DBF(path, load=True, encoding=enc, char_decode_errors="ignore")
                df = pd.DataFrame(iter(table))
                df = corriger_encodage_dataframe(df)
                return df
            except (UnicodeDecodeError, UnicodeError, ValueError):
                continue
        # Si aucun encodage ne fonctionne, essayer avec cp1252 par défaut
        table = DBF(path, load=True, encoding="cp1252", char_decode_errors="ignore")
        df = pd.DataFrame(iter(table))
        df = corriger_encodage_dataframe(df)
        return df

    raise ValueError(f"Extension non gérée: {ext}")


###############################################################################
# 3. Détection de colonnes et analyses génériques (inspirées de tes scripts)
###############################################################################


def trouver_colonnes(df: pd.DataFrame) -> dict:
    """
    Essaye de retrouver les colonnes clés en se basant sur des motifs dans les noms.
    """
    colonnes: dict[str, str] = {}

    mappings = {
        "domaine": ["domaine"],
        "theme": ["theme", "thème"],
        "type_usage": ["type_usage", "type usage", "type d'usager", "type usager"],
        "resultat": ["resultat", "résultat", "resultat_ctrl"],
        "fc_type": ["fc_type", "type fiche", "type_fiche"],
    }

    for cible, patterns in mappings.items():
        for col in df.columns:
            lower = str(col).lower()
            if any(p in lower for p in patterns):
                colonnes[cible] = col
                break

    return colonnes


def nettoyer_type_usager(type_usager: str) -> str:
    """
    Nettoie un type d'usager en supprimant les suffixes numériques.
    Exemple: "Particulier ... 1" -> "Particulier ..."
    """
    return re.sub(r'\s+\d+$', '', str(type_usager)).strip()


def decomposer_types_usagers(type_usagers: str) -> list[str]:
    """
    Décompose une chaîne contenant plusieurs types d'usagers séparés par des virgules.
    Exemple: "Agriculteur 1, Collectivité 1" -> ["Agriculteur", "Collectivité"]
    Exclut les valeurs vides ou nulles.
    """
    if pd.isna(type_usagers):
        return []
    type_str = str(type_usagers).strip()
    if not type_str:
        return []
    types = [t.strip() for t in type_str.split(',')]
    # Nettoyer et filtrer les valeurs vides
    types_nettoyes = [nettoyer_type_usager(t) for t in types if t.strip()]
    # Exclure les valeurs vides après nettoyage
    return [t for t in types_nettoyes if t and t.strip()]


def decomposer_dataframe(df: pd.DataFrame, colonne_type_usage: str) -> pd.DataFrame:
    """
    Décompose un DataFrame en créant une ligne par type d'usager quand plusieurs
    types sont combinés dans une même cellule (séparés par des virgules).
    """
    nouvelles_lignes = []
    for _, row in df.iterrows():
        if colonne_type_usage in df.columns:
            types_usagers = decomposer_types_usagers(row[colonne_type_usage])
            if types_usagers:
                # Créer une ligne pour chaque type d'usager
                for type_usager in types_usagers:
                    nouvelle_ligne = row.copy()
                    nouvelle_ligne[colonne_type_usage] = type_usager
                    nouvelles_lignes.append(nouvelle_ligne)
            else:
                # Pas de type d'usager valide, garder la ligne originale
                nouvelles_lignes.append(row.copy())
        else:
            nouvelles_lignes.append(row.copy())
    return pd.DataFrame(nouvelles_lignes)


def remplacer_valeurs_vides_tableau(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remplace toutes les valeurs vides (colonnes et lignes) par "Information non rapportée".
    S'applique à tous les tableaux générés.
    """
    df_renomme = df.copy()
    
    # 1. Renommer les colonnes vides
    nouvelles_colonnes = []
    for col in df_renomme.columns:
        if pd.isna(col) or str(col).strip() == "":
            nouvelles_colonnes.append("Information non rapportée")
        else:
            nouvelles_colonnes.append(col)
    df_renomme.columns = nouvelles_colonnes
    
    # 2. Remplacer les valeurs vides dans les données (colonnes texte/object)
    for col in df_renomme.columns:
        if col == "TOTAL":  # Ne pas modifier la ligne TOTAL
            continue
        if pd.api.types.is_string_dtype(df_renomme[col]) or df_renomme[col].dtype == "object":
            # Remplacer les valeurs NaN, None, et chaînes vides
            mask = df_renomme[col].isna() | (df_renomme[col].astype(str).str.strip() == "")
            df_renomme.loc[mask, col] = "Information non rapportée"
    
    return df_renomme


def ajouter_ligne_total(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    total_row: dict[str, object] = {}
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            total_row[col] = df[col].sum()
        else:
            total_row[col] = "TOTAL"

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def generer_tableaux(df: pd.DataFrame, colonnes: dict) -> dict[str, pd.DataFrame]:
    """
    Génère quelques tableaux de base :
    - nombre de contrôles par domaine / thème / type_usager / type_fiche
    - conformité par domaine / thème / type_usager (si possible)
    """
    resultats: dict[str, pd.DataFrame] = {}
    
    # Copie du DataFrame pour ne pas modifier l'original
    df_travail = df.copy()

    # 1. Nombre de contrôles par domaine
    if "domaine" in colonnes:
        counts = df_travail[colonnes["domaine"]].value_counts().reset_index()
        counts.columns = ["Domaine", "Nombre de contrôles"]
        counts = remplacer_valeurs_vides_tableau(counts)
        resultats["Nombre de contrôles par domaine"] = ajouter_ligne_total(counts)

    # 2. Nombre de contrôles par thème
    if "theme" in colonnes:
        counts = df_travail[colonnes["theme"]].value_counts().reset_index()
        counts.columns = ["Thème", "Nombre de contrôles"]
        counts = remplacer_valeurs_vides_tableau(counts)
        resultats["Nombre de contrôles par thème"] = ajouter_ligne_total(counts)

    # 3. Nombre de contrôles par type usager
    # IMPORTANT: Décomposer et nettoyer les types d'usagers avant comptage
    if "type_usage" in colonnes:
        try:
            col_type_usage = colonnes["type_usage"]
            # Vérifier que la colonne existe bien dans le DataFrame
            if col_type_usage not in df_travail.columns:
                print(f"Attention: Colonne '{col_type_usage}' introuvable dans le DataFrame")
            else:
                # Compter les valeurs vides avant filtrage pour information
                nb_vides_avant = df_travail[col_type_usage].isna().sum() + (
                    df_travail[col_type_usage].astype(str).str.strip() == ""
                ).sum()
                # Décomposer les types d'usagers combinés (séparés par virgules)
                df_decompose = decomposer_dataframe(df_travail, col_type_usage)
                # Nettoyer les suffixes numériques avec gestion d'erreur
                try:
                    df_decompose[col_type_usage] = df_decompose[col_type_usage].map(
                        lambda x: nettoyer_type_usager(x) if pd.notna(x) else x
                    )
                except Exception as e:
                    print(f"Erreur lors du nettoyage des types d'usagers: {e}")
                    # Fallback: utiliser directement les valeurs sans nettoyage
                    pass
                
                # Filtrer les valeurs vides ou nulles (ne pas les compter)
                df_decompose = df_decompose[df_decompose[col_type_usage].notna()]
                df_decompose = df_decompose[df_decompose[col_type_usage].astype(str).str.strip() != ""]
                if nb_vides_avant > 0:
                    print(
                        f"Attention : {nb_vides_avant} enregistrement(s) avec type d'usager vide "
                        f"ont été exclus du tableau."
                    )
                # Compter
                if not df_decompose.empty and col_type_usage in df_decompose.columns:
                    counts = df_decompose[col_type_usage].value_counts().reset_index()
                    counts.columns = ["Type d'usager", "Nombre de contrôles"]
                    # Trier par nombre décroissant
                    counts = counts.sort_values(by="Nombre de contrôles", ascending=False)
                    counts = remplacer_valeurs_vides_tableau(counts)
                    resultats["Nombre de contrôles par type usager"] = ajouter_ligne_total(counts)
        except Exception as e:
            print(f"Erreur lors de l'analyse des types d'usagers: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            # Ne pas faire planter tout le programme, continuer avec les autres tableaux

    # 4. Nombre de contrôles par type de fiche
    if "fc_type" in colonnes:
        counts = df_travail[colonnes["fc_type"]].value_counts().reset_index()
        counts.columns = ["Type de fiche contrôle", "Nombre de contrôles"]
        counts = remplacer_valeurs_vides_tableau(counts)
        resultats["Nombre de contrôles par type de fiche"] = ajouter_ligne_total(counts)

    # 5. Conformité par domaine / thème / type_usager
    if "resultat" in colonnes:
        if "domaine" in colonnes:
            tab = pd.crosstab(df_travail[colonnes["domaine"]], df_travail[colonnes["resultat"]]).reset_index()
            tab = remplacer_valeurs_vides_tableau(tab)
            resultats["Conformité par domaine"] = ajouter_ligne_total(tab)
        if "theme" in colonnes:
            tab = pd.crosstab(df_travail[colonnes["theme"]], df_travail[colonnes["resultat"]]).reset_index()
            tab = remplacer_valeurs_vides_tableau(tab)
            resultats["Conformité par thème"] = ajouter_ligne_total(tab)
        if "type_usage" in colonnes:
            try:
                # Pour la conformité par type usager, utiliser aussi la décomposition
                col_type_usage = colonnes["type_usage"]
                if col_type_usage in df_travail.columns:
                    df_decompose = decomposer_dataframe(df_travail, col_type_usage)
                    try:
                        df_decompose[col_type_usage] = df_decompose[col_type_usage].map(
                            lambda x: nettoyer_type_usager(x) if pd.notna(x) else x
                        )
                    except Exception as e:
                        print(f"Erreur lors du nettoyage des types d'usagers (conformité): {e}")
                    # Filtrer les valeurs vides ou nulles
                    df_decompose = df_decompose[df_decompose[col_type_usage].notna()]
                    df_decompose = df_decompose[df_decompose[col_type_usage].astype(str).str.strip() != ""]
                    if not df_decompose.empty:
                        tab = pd.crosstab(df_decompose[col_type_usage], df_decompose[colonnes["resultat"]]).reset_index()
                        tab = remplacer_valeurs_vides_tableau(tab)
                        resultats["Conformité par type usager"] = ajouter_ligne_total(tab)
            except Exception as e:
                print(f"Erreur lors de l'analyse de conformité par type usager: {type(e).__name__}: {e}")
                import traceback
                traceback.print_exc()
                # Ne pas faire planter tout le programme

    # 6. Tableaux spécifiques liés aux sources fixes (lib_natinf, _zone_tub)
    #    Ces colonnes peuvent avoir été ajoutées par enrichir_avec_sources_fixes().
    if "lib_natinf" in df_travail.columns:
        try:
            counts_natinf = (
                df_travail["lib_natinf"]
                .fillna("Information non rapportée")
                .astype(str)
                .str.strip()
                .value_counts()
                .reset_index()
            )
            counts_natinf.columns = ["Libellé NATINF", "Nombre de contrôles"]
            counts_natinf = counts_natinf.sort_values(
                by="Nombre de contrôles", ascending=False
            )
            counts_natinf = remplacer_valeurs_vides_tableau(counts_natinf)
            resultats["Nombre de contrôles par NATINF (lib_natinf)"] = ajouter_ligne_total(
                counts_natinf
            )
        except Exception as e:
            print(f"Erreur lors de la génération du tableau NATINF: {type(e).__name__}: {e}")

    if "_zone_tub" in df_travail.columns:
        try:
            # Comptage simple par zone (TUB / hors TUB / non renseigné)
            counts_tub = (
                df_travail["_zone_tub"]
                .fillna("")
                .astype(str)
                .str.strip()
                .replace("", "Hors zone TUB / non renseigné")
                .value_counts()
                .reset_index()
            )
            counts_tub.columns = ["Zone TUB", "Nombre de contrôles"]
            counts_tub = remplacer_valeurs_vides_tableau(counts_tub)
            resultats["Nombre de contrôles par zone TUB"] = ajouter_ligne_total(counts_tub)

            # Conformité par zone TUB si une colonne 'resultat' a été détectée
            if "resultat" in colonnes:
                col_res = colonnes["resultat"]
                df_tmp = df_travail.copy()
                df_tmp["_zone_tub_tmp"] = (
                    df_tmp["_zone_tub"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace("", "Hors zone TUB / non renseigné")
                )
                tab_tub = (
                    pd.crosstab(df_tmp["_zone_tub_tmp"], df_tmp[col_res])
                    .reset_index()
                )
                tab_tub = remplacer_valeurs_vides_tableau(tab_tub)
                resultats["Conformité par zone TUB"] = ajouter_ligne_total(tab_tub)
        except Exception as e:
            print(f"Erreur lors de la génération des tableaux zone TUB: {type(e).__name__}: {e}")

    return resultats


def generer_tableaux_generiques(df: pd.DataFrame, max_modalites: int = 20) -> dict[str, pd.DataFrame]:
    """
    Génère des tableaux génériques quand aucune colonne métier n'a été détectée :
    - descriptif des variables numériques
    - fréquences pour les variables catégorielles (texte)
    """
    resultats: dict[str, pd.DataFrame] = {}

    # 1. Statistiques descriptives sur les colonnes numériques
    num = df.select_dtypes(include=["number"])
    if not num.empty:
        desc = num.describe().T  # lignes = variables
        tab_desc = desc.reset_index().rename(columns={"index": "Variable"})
        tab_desc = remplacer_valeurs_vides_tableau(tab_desc)
        resultats["Statistiques numériques"] = tab_desc

    # 2. Fréquences des modalités pour les colonnes texte / catégories
    obj = df.select_dtypes(include=["object", "string"])
    for col in obj.columns:
        vc = obj[col].value_counts(dropna=False).head(max_modalites).reset_index()
        vc.columns = [col, "Effectif"]
        vc = remplacer_valeurs_vides_tableau(vc)
        resultats[f"Fréquences - {col}"] = vc

    return resultats


def _normaliser_texte(s: str) -> str:
    """
    Normalise une chaîne :
    - passage en minuscules
    - suppression des accents
    - suppression des espaces superflus
    """
    s = str(s).lower().strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s


def filtrer_cote_d_or(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tente de filtrer le DataFrame sur les données relatives au département
    de la Côte-d'Or, en utilisant plusieurs heuristiques :
    - recherche de valeurs contenant 'cote d or' / 'cote-d-or'
    - sinon, utilisation d'un code département / INSEE commençant par '21'
    Si aucun critère n'est trouvé, retourne le DataFrame inchangé.
    """
    if df.empty:
        return df

    nb_initial = len(df)

    # 1. Recherche d'un libellé texte mentionnant clairement la Côte-d'Or
    patterns_cote_dor = ["cote d or", "cote-d-or", "cote dor", "cote d'or"]

    obj_cols = df.select_dtypes(include=["object", "string"]).columns
    for col in obj_cols:
        try:
            series_norm = df[col].map(_normaliser_texte)
        except Exception:
            continue
        mask = series_norm.apply(
            lambda x: any(p in x for p in patterns_cote_dor) if isinstance(x, str) else False
        )
        if mask.any():
            df_filtre = df[mask].copy()
            print(
                f"Filtrage sur la colonne '{col}' pour le département de la Côte-d'Or : "
                f"{len(df_filtre)}/{nb_initial} lignes conservées"
            )
            return df_filtre

    # 2. Recherche d'un code département / INSEE commençant par '21'
    # Colonnes candidates par leur nom
    colonnes_candidates = [
        c
        for c in df.columns
        if any(kw in str(c).lower() for kw in ["depart", "dept", "dep", "code_insee", "insee"])
    ]

    for col in colonnes_candidates:
        serie = df[col]
        serie_str = serie.astype(str).str.strip()
        # Conserver : code 21 (Côte-d'Or) OU valeur vide/NaN (multi-sources : ne pas exclure les lignes sans code)
        mask_keep = (
            serie.isna()
            | (serie_str == "")
            | (serie_str.str.startswith("21", na=False))
        )
        if mask_keep.any():
            df_filtre = df[mask_keep].copy()
            print(
                f"Filtrage sur la colonne '{col}' (code '21' ou non renseigné) : "
                f"{len(df_filtre)}/{nb_initial} lignes conservées"
            )
            return df_filtre

    print("Aucun critère de filtrage spécifique à la Côte-d'Or trouvé, données non filtrées.")
    return df


###############################################################################
# 4. Sources fixes (référentiels NATINF, communes TUB, …)
###############################################################################


def charger_sources_fixes() -> dict[str, pd.DataFrame]:
    """
    Charge les sources fixes présentes dans le dossier 'sources' à côté du script :
    - fichier NATINF (codes + libellés)
    - fichier communes TUB (codes INSEE des communes de la zone TUB Côte-d'Or)

    Retourne un dict éventuellement partiel, par ex.:
    {
        "natinf": df_natinf,
        "tub_communes": df_tub,
    }
    """
    sources: dict[str, pd.DataFrame] = {}
    base_dir = os.path.dirname(__file__)
    dossier_sources = os.path.join(base_dir, "sources")
    if not os.path.isdir(dossier_sources):
        return sources

    # Parcourir les fichiers Excel du dossier sources
    for nom_fichier in os.listdir(dossier_sources):
        path = os.path.join(dossier_sources, nom_fichier)
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xls", ".xlsx"):
            continue

        try:
            df = pd.read_excel(path)
        except Exception:
            continue

        cols_lower = [str(c).lower() for c in df.columns]

        # Heuristique : fichier NATINF → une colonne contenant 'natinf'
        if any("natinf" in c for c in cols_lower) and "natinf" not in sources:
            sources["natinf"] = df
            continue

        # Heuristique : fichier communes TUB → une colonne contenant 'insee'
        if any("insee" in c for c in cols_lower) and "tub_communes" not in sources:
            sources["tub_communes"] = df
            continue

    return sources


def enrichir_avec_sources_fixes(df: pd.DataFrame, sources: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Enrichit le DataFrame d'analyse avec les sources fixes si possible :
    - jointure sur NATINF pour ajouter un libellé (colonne 'lib_natinf')
    - jointure sur codes INSEE pour marquer la zone TUB (colonne '_zone_tub')
    """
    if df.empty or not sources:
        return df

    df_enrichi = df.copy()

    # ------------------------------------------------------------------
    # 1) Référentiel NATINF : ajouter un libellé sur la base du numéro
    # ------------------------------------------------------------------
    df_natinf = sources.get("natinf")
    if isinstance(df_natinf, pd.DataFrame) and not df_natinf.empty:
        # Colonnes candidates côté référentiel
        cols_ref_lower = [str(c).lower() for c in df_natinf.columns]
        try:
            idx_code = next(i for i, c in enumerate(cols_ref_lower) if "natinf" in c)
        except StopIteration:
            idx_code = None
        # Libellé : colonne contenant 'lib' ou 'libell' ou, à défaut, la deuxième colonne
        idx_lib = None
        for i, c in enumerate(cols_ref_lower):
            if "lib" in c or "libell" in c:
                idx_lib = i
                break
        if idx_lib is None and len(df_natinf.columns) >= 2:
            idx_lib = 1
        if idx_code is not None and idx_lib is not None:
            col_code_ref = df_natinf.columns[idx_code]
            col_lib_ref = df_natinf.columns[idx_lib]

            # Côté données variables : chercher une colonne contenant 'natinf'
            cols_df_lower = [str(c).lower() for c in df_enrichi.columns]
            col_code_df = None
            for i, c in enumerate(cols_df_lower):
                if "natinf" in c:
                    col_code_df = df_enrichi.columns[i]
                    break

            if col_code_df is not None:
                ref_natinf = df_natinf[[col_code_ref, col_lib_ref]].dropna()
                ref_natinf = ref_natinf.drop_duplicates(subset=[col_code_ref])
                ref_natinf[col_code_ref] = ref_natinf[col_code_ref].astype(str).str.strip()
                df_enrichi[col_code_df] = df_enrichi[col_code_df].astype(str).str.strip()

                df_enrichi = df_enrichi.merge(
                    ref_natinf.rename(
                        columns={
                            col_code_ref: col_code_df,
                            col_lib_ref: "lib_natinf",
                        }
                    ),
                    on=col_code_df,
                    how="left",
                )

    # ------------------------------------------------------------------
    # 2) Communes TUB : marquer les enregistrements présents dans la zone
    # ------------------------------------------------------------------
    df_tub = sources.get("tub_communes")
    if isinstance(df_tub, pd.DataFrame) and not df_tub.empty:
        cols_tub_lower = [str(c).lower() for c in df_tub.columns]
        try:
            idx_insee = next(i for i, c in enumerate(cols_tub_lower) if "insee" in c)
        except StopIteration:
            idx_insee = None

        if idx_insee is not None:
            col_insee_ref = df_tub.columns[idx_insee]
            # Ensemble des codes INSEE de la zone TUB
            codes_tub = (
                df_tub[col_insee_ref].dropna().astype(str).str.strip().unique().tolist()
            )
            codes_tub_set = {c for c in codes_tub if c}

            # Chercher une colonne INSEE côté données variables
            cols_df_lower = [str(c).lower() for c in df_enrichi.columns]
            col_insee_df = None
            for i, c in enumerate(cols_df_lower):
                if "insee" in c:
                    col_insee_df = df_enrichi.columns[i]
                    break

            if col_insee_df is not None:
                serie_codes = df_enrichi[col_insee_df].astype(str).str.strip()
                df_enrichi["_zone_tub"] = serie_codes.apply(
                    lambda x: "Zone TUB Côte-d'Or" if x in codes_tub_set else ""
                )

    return df_enrichi


###############################################################################
# 4. Export vers Excel
###############################################################################


def exporter_rapport_excel(resultats: dict[str, pd.DataFrame], dossier_sortie: str, base_name: str) -> str:
    os.makedirs(dossier_sortie, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    chemin = os.path.join(dossier_sortie, f"rapport_{base_name}_{timestamp}.xlsx")

    # Noms de feuilles déjà utilisés (Excel limite 31 car. et interdit les doublons)
    noms_feuilles_utilises: set[str] = set()

    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        for nom, df in resultats.items():
            # S'assurer que les données sont bien corrigées avant export
            df_export = corriger_encodage_dataframe(df.copy())
            
            # Nettoyage du nom de feuille pour Excel :
            # - suppression uniquement des caractères strictement interdits par Excel : : \ / ? * [ ]
            # - préservation des accents et caractères français (é, è, ê, à, ç, etc.)
            # - limitation à 31 caractères, unicité (suffixe _1, _2 si doublon)
            raw_name = str(nom)
            safe_name = re.sub(r'[:\\\\/\\?\\*\\[\\]]', "_", raw_name)
            safe_name = re.sub(r'[<>"]', "_", safe_name)
            safe_name = safe_name.strip()
            if not safe_name:
                safe_name = "Feuille"
            base_name = safe_name[:31]
            sheet_name = base_name
            compteur = 1
            while sheet_name in noms_feuilles_utilises:
                suffix = f"_{compteur}"
                sheet_name = (base_name[: 31 - len(suffix)] + suffix)
                compteur += 1
            noms_feuilles_utilises.add(sheet_name)
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)

    return chemin


###############################################################################
# 5. Programme principal
###############################################################################


def main() -> None:
    print("=== Programme d'analyse OSCAN (version générique) ===")

    # 1) Déterminer le dossier à analyser
    #    - si un argument est fourni : l'utiliser en priorité
    #    - sinon : poser la question à l'utilisateur avec un défaut = dossier courant
    if len(sys.argv) > 1:
        dossier = sys.argv[1]
        print(f"Dossier fourni en argument : {dossier}")
    else:
        dossier = input(
            "Chemin du dossier à analyser (laisser vide pour le dossier courant) : "
        ).strip() or os.getcwd()

    if not os.path.isdir(dossier):
        print(f"Erreur : le chemin '{dossier}' n'est pas un dossier.")
        return

    print(f"Recherche de fichiers dans : {dossier}")
    candidats = trouver_fichiers_candidats(dossier)

    if not candidats:
        print("Aucun fichier compatible trouvé (.xlsx, .csv, .dbf).")
        return

    selection = selectionner_plus_recents_par_extension(candidats)

    if not selection:
        print("Aucun fichier récent trouvé pour les extensions supportées.")
        return

    print("\nFichiers retenus par extension (les plus récents) :")
    items = list(selection.items())  # liste de tuples (ext, path)
    for idx, (ext, path) in enumerate(items, start=1):
        print(f"  {idx}. {ext}: {path}")

    # Choix des fichiers par l'utilisateur (multi-sélection)
    choix = input(
        "\nEntrez les numéros des fichiers à analyser séparés par des virgules "
        "(laisser vide pour tous) : "
    ).strip()

    index_choisis: list[int] = []
    if not choix:
        index_choisis = list(range(1, len(items) + 1))
    else:
        for morceau in choix.split(","):
            morceau = morceau.strip()
            if not morceau:
                continue
            if not morceau.isdigit():
                print(f"Ignoré (non numérique) : '{morceau}'")
                continue
            val = int(morceau)
            if 1 <= val <= len(items):
                index_choisis.append(val)
            else:
                print(f"Ignoré (hors plage) : {val}")

    index_choisis = sorted(set(index_choisis))
    if not index_choisis:
        print("Aucun fichier sélectionné, arrêt.")
        return

    fichiers_choisis: list[tuple[str, str]] = []  # (ext, path)
    for idx in index_choisis:
        ext, path = items[idx - 1]
        fichiers_choisis.append((ext, path))

    print("\nFichiers sélectionnés pour l'analyse :")
    for ext, path in fichiers_choisis:
        print(f"  - {ext}: {path}")

    # Chargement de tous les fichiers sélectionnés et concaténation
    dataframes: list[pd.DataFrame] = []
    resume_sources: list[dict[str, object]] = []

    for ext, path in fichiers_choisis:
        try:
            df_tmp = charger_fichier(path)
        except Exception as e:
            print(f"Erreur lors du chargement de '{path}': {e} (fichier ignoré)")
            continue

        if df_tmp.empty:
            print(f"Fichier vide ignoré : {path}")
            continue

        # Ajouter une colonne d'origine
        df_tmp = df_tmp.copy()
        df_tmp["__source_fichier"] = os.path.basename(path)
        df_tmp["__source_extension"] = ext

        dataframes.append(df_tmp)
        resume_sources.append(
            {
                "Fichier": os.path.basename(path),
                "Chemin_complet": path,
                "Extension": ext,
                "Nb_lignes": len(df_tmp),
                "Nb_colonnes": len(df_tmp.columns),
            }
        )

    if not dataframes:
        print("Aucun fichier valide chargé, arrêt.")
        return

    df_total = pd.concat(dataframes, ignore_index=True)
    print(f"\nDonnées combinées avant filtrage : {len(df_total)} lignes, {len(df_total.columns)} colonnes")

    # Filtrer sur le département de la Côte-d'Or si possible
    df_total = filtrer_cote_d_or(df_total)
    print(f"Données après filtrage (Côte-d'Or si détectable) : {len(df_total)} lignes")

    # Enrichir avec les sources fixes (NATINF, communes TUB, ...)
    sources_fixes = charger_sources_fixes()
    if sources_fixes:
        df_total = enrichir_avec_sources_fixes(df_total, sources_fixes)
        print(f"Données après enrichissement avec les sources fixes : {len(df_total)} lignes, {len(df_total.columns)} colonnes")
    else:
        print("Aucune source fixe trouvée dans le dossier 'sources'.")

    colonnes = trouver_colonnes(df_total)
    print(f"Colonnes détectées sur l'ensemble : {colonnes}")

    # Si des colonnes métier sont détectées, on privilégie ces tableaux-là
    resultats = generer_tableaux(df_total, colonnes)

    # Si rien n'a été généré, on bascule sur des tableaux génériques
    if not resultats:
        print("Aucune colonne clé détectée ou aucun tableau spécifique généré, "
              "génération de tableaux génériques.")
        resultats = generer_tableaux_generiques(df_total)
        if not resultats:
            print("Aucun tableau généré (données vides ou non exploitables).")
            return

    # Ajouter un onglet de résumé des sources
    resultats["Resume_sources"] = pd.DataFrame(resume_sources)

    print("Tableaux générés :", list(resultats.keys()))

    dossier_resultats = os.path.join(os.path.dirname(__file__), "resultats")
    # nom de base à partir du premier fichier choisi
    base_name = os.path.splitext(os.path.basename(fichiers_choisis[0][1]))[0]

    chemin_excel = exporter_rapport_excel(resultats, dossier_resultats, base_name)
    print(f"\nClasseur Excel créé : {chemin_excel}")

    from rapport_pdf_oscean import generer_pdf_oscan
    try:
        chemin_pdf = generer_pdf_oscan(dossier_resultats, resultats, base_name)
        print(f"Rapport PDF créé   : {chemin_pdf}")
    except Exception as e:
        print(f"Attention : échec de la génération du PDF ({type(e).__name__}: {e})")

    print("\nAnalyse terminée.")


if __name__ == "__main__":
    main()

