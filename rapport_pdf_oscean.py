import os
from pathlib import Path
from datetime import datetime
from typing import Dict

import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
    Image,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import matplotlib.pyplot as plt
import pandas as pd


###############################################################################
# Paramètres de mise en forme (charte OFB adaptée au programme OSCAN)
###############################################################################

OFB_BLEU_PRINCIPAL = "#003A76"
OFB_VERT = "#0E823A"
OFB_BLANC = "#FFFFFF"
OFB_NOIR = "#000000"

LARGEUR_PAGE, HAUTEUR_PAGE = A4
MARGE_MM = 17
MARGE = MARGE_MM * mm
LARGEUR_UTILE = LARGEUR_PAGE - 2 * MARGE


def _pied_page(canvas, doc):
    """Ajoute un pied de page discret avec titre du rapport et numérotation."""
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))

    texte_gauche = "OFB – Bilan OSCAN – Synthèse"
    page_num = canvas.getPageNumber()
    texte_droite = f"Page {page_num}"

    canvas.drawString(MARGE, MARGE - 6, texte_gauche)
    w = canvas.stringWidth(texte_droite, "Helvetica", 8)
    canvas.drawString(LARGEUR_PAGE - MARGE - w, MARGE - 6, texte_droite)
    canvas.restoreState()


def _enregistrer_police_arial() -> str:
    """Enregistre Arial si disponible (charte OFB : Marianne ou Arial)."""
    for base in ("C:/Windows/Fonts/", "", "/usr/share/fonts/truetype/msttcorefonts/"):
        try:
            pdfmetrics.registerFont(TTFont("Arial", base + "arial.ttf"))
            pdfmetrics.registerFont(TTFont("Arial-Bold", base + "arialbd.ttf"))
            return "Arial"
        except Exception:
            continue
    return "Helvetica"


def _style_entete_table(font_bold: str = "Helvetica-Bold", font_body: str = "Helvetica") -> TableStyle:
    """Style de tableau type OFB : en-tête bleue, alternance de gris, grille fine."""
    try:
        pdfmetrics.getFont(font_bold)
    except Exception:
        font_bold = "Helvetica-Bold"
    try:
        pdfmetrics.getFont(font_body)
    except Exception:
        font_body = "Helvetica"

    return TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(OFB_BLEU_PRINCIPAL)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(OFB_BLANC)),
            ("FONTNAME", (0, 0), (-1, 0), font_bold),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [OFB_BLANC, colors.HexColor("#f5f5f5")]),
            ("FONTNAME", (0, 1), (-1, -1), font_body),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
        ]
    )


def _build_styles(police: str):
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="OFBTitle",
            fontName=police,
            fontSize=13,
            textColor=colors.HexColor(OFB_BLEU_PRINCIPAL),
            spaceBefore=12,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="OFBHeading2",
            fontName=police,
            fontSize=11,
            textColor=colors.HexColor(OFB_BLEU_PRINCIPAL),
            spaceBefore=10,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="OFBBody",
            fontName=police,
            fontSize=9,
            textColor=colors.HexColor(OFB_NOIR),
            spaceAfter=4,
        )
    )
    return styles


def _normaliser_largeurs(col_widths, ncols: int):
    """Ajuste les largeurs pour que la somme = LARGEUR_UTILE (évite débordement)."""
    if not col_widths or len(col_widths) != ncols:
        return [LARGEUR_UTILE / ncols] * ncols
    total = sum(col_widths)
    if total <= 0:
        return [LARGEUR_UTILE / ncols] * ncols
    if total > LARGEUR_UTILE:
        return [w * LARGEUR_UTILE / total for w in col_widths]
    return list(col_widths)


def _table_from_dataframe(
    df: pd.DataFrame,
    font_bold: str,
    font_body: str,
    wrap_first_column: bool = True,
):
    """
    Convertit un DataFrame en tableau ReportLab, avec une gestion simple
    du retour à la ligne sur la première colonne si souhaité.
    """
    if df.empty:
        return None

    data = [list(df.columns)] + df.astype(str).fillna("").values.tolist()
    ncols = len(data[0])
    col_widths = _normaliser_largeurs([LARGEUR_UTILE / ncols] * ncols, ncols)

    # Style pour les cellules (retour à la ligne)
    try:
        pdfmetrics.getFont(font_body)
    except Exception:
        font_body = "Helvetica"
    cell_style = ParagraphStyle(
        name="TableCell",
        fontName=font_body,
        fontSize=8,
        leading=9,
        leftIndent=0,
        rightIndent=0,
        spaceBefore=0,
        spaceAfter=0,
        wordWrap="CJK",
    )

    out_data = []
    for i, row in enumerate(data):
        out_row = []
        for j, cell in enumerate(row):
            s = str(cell).strip()
            if i > 0 and wrap_first_column and j == 0 and len(s) > 1:
                s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                out_row.append(Paragraph(s, cell_style))
            else:
                out_row.append(s)
        out_data.append(out_row)

    t = Table(out_data, colWidths=col_widths)
    t.setStyle(_style_entete_table(font_bold, font_body))
    return t


def _generer_graphiques(
    resultats: Dict[str, pd.DataFrame],
    dossier_sortie: Path,
    base_name: str,
    horodatage: str,
    graph_config: Dict[str, str] | None = None,
) -> Dict[str, Path]:
    """
    Génère des graphiques simples (barres horizontales) pour les tableaux pertinents.
    Retourne un dict {nom_tableau: chemin_fichier_png}.
    """
    figures: Dict[str, Path] = {}
    fig_dir = dossier_sortie / "figures"
    fig_dir.mkdir(parents=True, exist_ok=True)

    def _slugify(nom: str) -> str:
        s = str(nom).lower()
        s = "".join(c if c.isalnum() or c in "-_" else "_" for c in s)
        while "__" in s:
            s = s.replace("__", "_")
        return s.strip("_") or "graphique"

    for cle, df in resultats.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue

        # Si une configuration explicite est fournie, ne garder que les tableaux sélectionnés
        if graph_config is not None and cle not in graph_config:
            continue

        # Chercher une colonne numérique
        numeric_col = None
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                numeric_col = col
                break
        if numeric_col is None:
            continue

        # Colonne de catégories : première non numérique
        category_col = None
        for col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                category_col = col
                break
        if category_col is None:
            continue

        # Limiter le nombre de catégories pour lisibilité
        df_plot = df[[category_col, numeric_col]].copy()
        df_plot = df_plot.dropna(subset=[numeric_col])
        if df_plot.empty:
            continue
        df_plot = df_plot.sort_values(by=numeric_col, ascending=True).tail(20)

        try:
            # Style demandé (par défaut : histogramme/barres)
            style = "bar"
            if graph_config is not None:
                style = graph_config.get(cle, "bar") or "bar"

            fig, ax = plt.subplots(figsize=(8, 4))

            if style == "pie":
                # Camembert : on utilise les valeurs numériques comme tailles
                labels = df_plot[category_col].astype(str)
                sizes = df_plot[numeric_col]
                # Limiter le nombre de parts explose parfois la lisibilité, mais déjà limité à 20
                ax.pie(
                    sizes,
                    labels=labels,
                    autopct="%1.1f%%",
                    startangle=90,
                    textprops={"fontsize": 8},
                )
                ax.set_title(str(cle), color=OFB_BLEU_PRINCIPAL)
                ax.axis("equal")
            elif style == "line":
                # Courbe simple : catégories sur l'axe X, valeurs sur Y
                x = range(len(df_plot))
                ax.plot(
                    x,
                    df_plot[numeric_col],
                    marker="o",
                    color=colors.HexColor(OFB_BLEU_PRINCIPAL),
                )
                ax.set_xticks(x)
                ax.set_xticklabels(df_plot[category_col].astype(str), rotation=45, ha="right")
                ax.set_ylabel(str(numeric_col))
                ax.set_title(str(cle), color=OFB_BLEU_PRINCIPAL)
                ax.grid(axis="y", linestyle=":", color="#cccccc", alpha=0.7)
            else:
                # Histogramme / barres horizontales (par défaut)
                ax.barh(
                    df_plot[category_col].astype(str),
                    df_plot[numeric_col],
                    color=colors.HexColor(OFB_BLEU_PRINCIPAL),
                )
                ax.set_xlabel(str(numeric_col))
                ax.set_ylabel(str(category_col))
                ax.set_title(str(cle), color=OFB_BLEU_PRINCIPAL)
                ax.grid(axis="x", linestyle=":", color="#cccccc", alpha=0.7)

            plt.tight_layout()

            slug = _slugify(cle)
            fig_path = fig_dir / f"{base_name}{slug}_{horodatage}.png"
            fig.savefig(fig_path, dpi=150)
            plt.close(fig)

            figures[cle] = fig_path
        except Exception:
            # Ne pas faire échouer le PDF si un graphique pose problème
            plt.close("all")
            continue

    return figures


def generer_pdf_oscan(
    dossier_sortie: str | Path,
    resultats: Dict[str, pd.DataFrame],
    base_name: str,
    graph_config: Dict[str, str] | None = None,
) -> str:
    """
    Génère un rapport PDF de synthèse OSCAN, en complément du classeur Excel.

    - dossier_sortie : dossier 'resultats'
    - resultats : dictionnaire de tableaux d'analyse (y compris éventuellement 'Resume_sources')
    - base_name : nom de base (utilisé pour nommer le fichier)
    """
    dossier = Path(dossier_sortie)
    dossier.mkdir(parents=True, exist_ok=True)

    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Nom du PDF : base_name puis horodatage (ex : oscan_YYYYMMDD_HHMMSS.pdf)
    nom_pdf = dossier / f"{base_name}{horodatage}.pdf"

    # Générer les graphiques à partir des tableaux (respect de la charte OFB)
    figures = _generer_graphiques(resultats, dossier, base_name, horodatage, graph_config)

    police = _enregistrer_police_arial()
    font_table_header = "Arial-Bold" if police == "Arial" else "Helvetica-Bold"
    font_table_body = police
    styles = _build_styles(police)

    doc = SimpleDocTemplate(
        str(nom_pdf),
        pagesize=A4,
        leftMargin=MARGE,
        rightMargin=MARGE,
        topMargin=MARGE,
        bottomMargin=MARGE,
    )

    story = []

    # ------------------------------------------------------------------
    # Page de garde
    # ------------------------------------------------------------------
    # Bandeau de couverture inspiré des gabarits OFB :
    # grand titre centré, sous-titre et date, avec marges généreuses.
    story.append(Spacer(1, 45 * mm))
    story.append(
        Paragraph(
            "Bilan OSCAN – Synthèse",
            ParagraphStyle(
                name="CoverTitle",
                fontName=police,
                fontSize=18,
                textColor=colors.HexColor(OFB_BLEU_PRINCIPAL),
                alignment=1,
                spaceAfter=20,
            ),
        )
    )
    story.append(
        Paragraph(
            "Office français de la biodiversité",
            ParagraphStyle(
                name="CoverSub",
                fontName=police,
                fontSize=11,
                textColor=colors.HexColor(OFB_VERT),
                alignment=1,
                spaceAfter=8,
            ),
        )
    )
    story.append(
        Paragraph(
            f"Rapport généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
            ParagraphStyle(
                name="CoverDate",
                fontName=police,
                fontSize=9,
                textColor=colors.HexColor(OFB_NOIR),
                alignment=1,
                spaceAfter=6,
            ),
        )
    )
    story.append(PageBreak())

    section_num = 1

    # ------------------------------------------------------------------
    # 1. Résumé des sources (si présent)
    # ------------------------------------------------------------------
    resume_df = resultats.get("Resume_sources")
    if isinstance(resume_df, pd.DataFrame) and not resume_df.empty:
        tab_resume = _table_from_dataframe(
            resume_df, font_table_header, font_table_body, wrap_first_column=False
        )
        if tab_resume is not None:
            story.append(
                KeepTogether(
                    [
                        Paragraph(f"{section_num}. Résumé des sources analysées", styles["OFBTitle"]),
                        Spacer(1, 4 * mm),
                        Paragraph(
                            "Tableau récapitulatif des fichiers sources pris en compte dans ce bilan.",
                            styles["OFBBody"],
                        ),
                        Spacer(1, 4 * mm),
                        tab_resume,
                        Spacer(1, 10 * mm),
                    ]
                )
            )
            section_num += 1

    # ------------------------------------------------------------------
    # 2. Zoom sur les analyses NATINF / zone TUB si disponibles
    # ------------------------------------------------------------------
    cle_natinf = "Nombre de contrôles par NATINF (lib_natinf)"
    cle_tub_counts = "Nombre de contrôles par zone TUB"
    cle_tub_conf = "Conformité par zone TUB"

    if cle_natinf in resultats and isinstance(resultats[cle_natinf], pd.DataFrame):
        df_nat = resultats[cle_natinf]
        if not df_nat.empty:
            tab_nat = _table_from_dataframe(
                df_nat, font_table_header, font_table_body, wrap_first_column=True
            )
            if tab_nat is not None:
                elements = [
                    Paragraph(f"{section_num}. Répartition des contrôles par NATINF", styles["OFBTitle"]),
                    Spacer(1, 4 * mm),
                    Paragraph(
                        "Nombre de contrôles regroupés par libellé NATINF (référentiel fixe).",
                        styles["OFBBody"],
                    ),
                    Spacer(1, 4 * mm),
                    tab_nat,
                ]
                # Ajouter le graphique associé si disponible
                fig_path = figures.get(cle_natinf)
                if fig_path and fig_path.exists():
                    elements.extend(
                        [
                            Spacer(1, 6 * mm),
                            Image(str(fig_path), width=LARGEUR_UTILE, preserveAspectRatio=True, hAlign="CENTER"),
                        ]
                    )
                elements.append(Spacer(1, 10 * mm))
                story.append(KeepTogether(elements))
                section_num += 1

    if cle_tub_counts in resultats and isinstance(resultats[cle_tub_counts], pd.DataFrame):
        df_tub = resultats[cle_tub_counts]
        if not df_tub.empty:
            tab_tub = _table_from_dataframe(
                df_tub, font_table_header, font_table_body, wrap_first_column=False
            )
            if tab_tub is not None:
                elements = [
                    Paragraph(f"{section_num}. Contrôles par zone TUB", styles["OFBTitle"]),
                    Spacer(1, 4 * mm),
                    Paragraph(
                        "Répartition des contrôles selon l'appartenance ou non à la zone TUB de la Côte-d'Or.",
                        styles["OFBBody"],
                    ),
                    Spacer(1, 4 * mm),
                    tab_tub,
                ]
                fig_path = figures.get(cle_tub_counts)
                if fig_path and fig_path.exists():
                    elements.extend(
                        [
                            Spacer(1, 6 * mm),
                            Image(str(fig_path), width=LARGEUR_UTILE, preserveAspectRatio=True, hAlign="CENTER"),
                        ]
                    )
                elements.append(Spacer(1, 10 * mm))
                story.append(KeepTogether(elements))
                section_num += 1

    if cle_tub_conf in resultats and isinstance(resultats[cle_tub_conf], pd.DataFrame):
        df_tub_conf = resultats[cle_tub_conf]
        if not df_tub_conf.empty:
            tab_tub_conf = _table_from_dataframe(
                df_tub_conf, font_table_header, font_table_body, wrap_first_column=False
            )
            if tab_tub_conf is not None:
                elements = [
                    Paragraph(f"{section_num}. Conformité par zone TUB", styles["OFBTitle"]),
                    Spacer(1, 4 * mm),
                    Paragraph(
                        "Tableau de conformité croisant la zone TUB et le résultat des contrôles.",
                        styles["OFBBody"],
                    ),
                    Spacer(1, 4 * mm),
                    tab_tub_conf,
                ]
                fig_path = figures.get(cle_tub_conf)
                if fig_path and fig_path.exists():
                    elements.extend(
                        [
                            Spacer(1, 6 * mm),
                            Image(str(fig_path), width=LARGEUR_UTILE, preserveAspectRatio=True, hAlign="CENTER"),
                        ]
                    )
                elements.append(Spacer(1, 10 * mm))
                story.append(KeepTogether(elements))
                section_num += 1

    # ------------------------------------------------------------------
    # 3. Autres tableaux d'analyse
    # ------------------------------------------------------------------
    cles_ignorees = {"Resume_sources", cle_natinf, cle_tub_counts, cle_tub_conf}
    autres_cles = [k for k in resultats.keys() if k not in cles_ignorees]
    if autres_cles:
        for cle in sorted(autres_cles):
            df = resultats[cle]
            if not isinstance(df, pd.DataFrame) or df.empty:
                continue

            titre = f"{section_num}. {cle}"
            tab = _table_from_dataframe(df, font_table_header, font_table_body, wrap_first_column=True)
            if tab is None:
                continue

            elements = [
                Paragraph(titre, styles["OFBTitle"]),
                Spacer(1, 4 * mm),
                tab,
            ]
            fig_path = figures.get(cle)
            if fig_path and fig_path.exists():
                elements.extend(
                    [
                        Spacer(1, 6 * mm),
                        Image(str(fig_path), width=LARGEUR_UTILE, preserveAspectRatio=True, hAlign="CENTER"),
                    ]
                )
            elements.append(Spacer(1, 10 * mm))
            story.append(KeepTogether(elements))
            section_num += 1

    # Si aucun contenu, ajouter un message simple
    if not story:
        story.append(
            Paragraph(
                "Aucune donnée exploitable pour générer le rapport PDF.",
                styles["OFBBody"],
            )
        )

    doc.build(story, onFirstPage=_pied_page, onLaterPages=_pied_page)
    return str(nom_pdf)


