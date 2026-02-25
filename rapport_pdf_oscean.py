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
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


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


def generer_pdf_oscan(
    dossier_sortie: str | Path,
    resultats: Dict[str, pd.DataFrame],
    base_name: str,
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
    nom_pdf = dossier / f"rapport_{base_name}_{horodatage}.pdf"

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
    story.append(Spacer(1, 40 * mm))
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
                story.append(
                    KeepTogether(
                        [
                            Paragraph(f"{section_num}. Répartition des contrôles par NATINF", styles["OFBTitle"]),
                            Spacer(1, 4 * mm),
                            Paragraph(
                                "Nombre de contrôles regroupés par libellé NATINF (référentiel fixe).",
                                styles["OFBBody"],
                            ),
                            Spacer(1, 4 * mm),
                            tab_nat,
                            Spacer(1, 10 * mm),
                        ]
                    )
                )
                section_num += 1

    if cle_tub_counts in resultats and isinstance(resultats[cle_tub_counts], pd.DataFrame):
        df_tub = resultats[cle_tub_counts]
        if not df_tub.empty:
            tab_tub = _table_from_dataframe(
                df_tub, font_table_header, font_table_body, wrap_first_column=False
            )
            if tab_tub is not None:
                story.append(
                    KeepTogether(
                        [
                            Paragraph(f"{section_num}. Contrôles par zone TUB", styles["OFBTitle"]),
                            Spacer(1, 4 * mm),
                            Paragraph(
                                "Répartition des contrôles selon l'appartenance ou non à la zone TUB de la Côte-d'Or.",
                                styles["OFBBody"],
                            ),
                            Spacer(1, 4 * mm),
                            tab_tub,
                            Spacer(1, 10 * mm),
                        ]
                    )
                )
                section_num += 1

    if cle_tub_conf in resultats and isinstance(resultats[cle_tub_conf], pd.DataFrame):
        df_tub_conf = resultats[cle_tub_conf]
        if not df_tub_conf.empty:
            tab_tub_conf = _table_from_dataframe(
                df_tub_conf, font_table_header, font_table_body, wrap_first_column=False
            )
            if tab_tub_conf is not None:
                story.append(
                    KeepTogether(
                        [
                            Paragraph(f"{section_num}. Conformité par zone TUB", styles["OFBTitle"]),
                            Spacer(1, 4 * mm),
                            Paragraph(
                                "Tableau de conformité croisant la zone TUB et le résultat des contrôles.",
                                styles["OFBBody"],
                            ),
                            Spacer(1, 4 * mm),
                            tab_tub_conf,
                            Spacer(1, 10 * mm),
                        ]
                    )
                )
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

            story.append(
                KeepTogether(
                    [
                        Paragraph(titre, styles["OFBTitle"]),
                        Spacer(1, 4 * mm),
                        tab,
                        Spacer(1, 10 * mm),
                    ]
                )
            )
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


