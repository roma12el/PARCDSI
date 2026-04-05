"""
pdf_export.py — Génération de rapports PDF professionnels pour Amendis DSI
"""
import io
from datetime import datetime
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF

# ─── Couleurs Amendis ─────────────────────────────────────────────────────────
ROUGE       = colors.HexColor("#C8102E")
ROUGE_CLAIR = colors.HexColor("#E8415B")
ROUGE_PALE  = colors.HexColor("#F5A0AE")
ROUGE_BG    = colors.HexColor("#FFF0F2")
GRIS_FONCE  = colors.HexColor("#2D2D2D")
GRIS_MOY    = colors.HexColor("#6B6B6B")
GRIS_CLAIR  = colors.HexColor("#F2F2F2")
BLANC       = colors.white
VERT        = colors.HexColor("#2ECC71")
ORANGE      = colors.HexColor("#E67E22")
BLEU        = colors.HexColor("#2980B9")
JAUNE       = colors.HexColor("#F1C40F")

PRIO_COLORS = {
    "Urgent":       ROUGE,
    "Prioritaire":  ORANGE,
    "surveiller":   JAUNE,
    "Opérationnel": VERT,
    "Operationnel": VERT,
}

DATE_STR = datetime.now().strftime("%d/%m/%Y")
YEAR_STR = "2026"


def _get_prio_color(prio_str):
    for key, col in PRIO_COLORS.items():
        if key.lower() in str(prio_str).lower():
            return col
    return GRIS_MOY


def _val(v, default="N/A"):
    """Valeur sûre pour affichage."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return default
    s = str(v).strip()
    return s if s not in ("nan", "NaT", "None", "") else default


# ─── Styles ───────────────────────────────────────────────────────────────────
def _make_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        "AmendisTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=22,
        textColor=BLANC,
        alignment=TA_CENTER,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        "AmendisSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        textColor=BLANC,
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "SectionTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=13,
        textColor=ROUGE,
        spaceBefore=14,
        spaceAfter=6,
        borderPad=4,
    ))
    styles.add(ParagraphStyle(
        "FieldLabel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=GRIS_MOY,
    ))
    styles.add(ParagraphStyle(
        "FieldValue",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=ROUGE,
    ))
    styles.add(ParagraphStyle(
        "BodyText2",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=GRIS_FONCE,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=BLANC,
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "Caption",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8,
        textColor=GRIS_MOY,
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "SmallBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        textColor=GRIS_FONCE,
    ))
    return styles


# ─── En-tête de page ──────────────────────────────────────────────────────────
def _header_block(styles, title, subtitle=""):
    """Bloc en-tête rouge avec titre."""
    w = 17 * cm

    data = [[
        Paragraph(f"<b>{title}</b>", styles["AmendisTitle"]),
    ]]
    if subtitle:
        data.append([Paragraph(subtitle, styles["AmendisSubtitle"])])

    t = Table(data, colWidths=[w])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ROUGE),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    return t


def _meta_line(styles, left="", right=""):
    """Ligne de métadonnées sous l'en-tête."""
    data = [[
        Paragraph(f"<font size='8' color='#6B6B6B'>{left}</font>", styles["Normal"]),
        Paragraph(f"<font size='8' color='#6B6B6B'>{right}</font>",
                  ParagraphStyle("R", parent=styles["Normal"], alignment=TA_RIGHT)),
    ]]
    t = Table(data, colWidths=[9 * cm, 8 * cm])
    t.setStyle(TableStyle([
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


# ─── KPI row ──────────────────────────────────────────────────────────────────
def _kpi_row(items):
    """Ligne de KPI: liste de (label, value, color)."""
    n = len(items)
    if n == 0:
        return Spacer(1, 6)

    col_w = 17 * cm / n
    data_row = []
    for label, value, col in items:
        cell = Table(
            [[Paragraph(f"<b>{value}</b>",
                        ParagraphStyle("KV", fontName="Helvetica-Bold",
                                       fontSize=20, textColor=col, alignment=TA_CENTER))],
             [Paragraph(label,
                        ParagraphStyle("KL", fontName="Helvetica",
                                       fontSize=8, textColor=GRIS_MOY, alignment=TA_CENTER))]],
            colWidths=[col_w - 8],
        )
        cell.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BLANC),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
            ("LINEBELOW", (0, 0), (-1, 0), 2, col),
        ]))
        data_row.append(cell)

    outer = Table([data_row], colWidths=[col_w] * n)
    outer.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLANC),
        ("GRID", (0, 0), (-1, -1), 0.5, GRIS_CLAIR),
        ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return outer


# ─── Tableau de données ───────────────────────────────────────────────────────
def _data_table(df: pd.DataFrame, col_labels: dict, max_rows=None):
    """Crée un tableau Reportlab à partir d'un DataFrame."""
    if df is None or len(df) == 0:
        return Paragraph("Aucune donnée disponible.", _make_styles()["BodyText2"])

    cols = [c for c in col_labels.keys() if c in df.columns]
    labels = [col_labels[c] for c in cols]

    if max_rows:
        df = df.head(max_rows)

    # Header
    header = [Paragraph(f"<b>{lbl}</b>",
                         ParagraphStyle("TH", fontName="Helvetica-Bold",
                                        fontSize=7, textColor=BLANC, alignment=TA_CENTER))
              for lbl in labels]

    # Rows
    rows = [header]
    for _, row in df.iterrows():
        r = []
        for c in cols:
            v = _val(row.get(c))
            # Coloration priorité
            if c == "priorite":
                pc = _get_prio_color(v)
                cell = Paragraph(f"<b>{v}</b>",
                                  ParagraphStyle("PC", fontName="Helvetica-Bold",
                                                 fontSize=7, textColor=pc))
            else:
                cell = Paragraph(str(v),
                                  ParagraphStyle("TD", fontName="Helvetica",
                                                 fontSize=7, textColor=GRIS_FONCE))
            r.append(cell)
        rows.append(r)

    # Largeurs de colonnes
    total_w = 17 * cm
    col_w = total_w / len(cols)
    col_widths = [col_w] * len(cols)

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), ROUGE),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#E0E0E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANC, GRIS_CLAIR]),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    t.setStyle(TableStyle(style))
    return t


# ─── Fiche info-card ──────────────────────────────────────────────────────────
def _info_card(fields, ncols=3):
    """Carte d'informations en grille ncols colonnes."""
    styles = _make_styles()
    n = len(fields)
    row_size = ncols
    data_rows = []
    for i in range(0, n, row_size):
        chunk = fields[i:i + row_size]
        # Pad to ncols
        while len(chunk) < row_size:
            chunk.append(("", ""))
        row = []
        for label, value in chunk:
            cell = Table(
                [[Paragraph(label, styles["FieldLabel"])],
                 [Paragraph(str(value), styles["FieldValue"])]],
                colWidths=[(17 * cm / row_size) - 8],
            )
            cell.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), BLANC),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, -1), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("BOX", (0, 0), (-1, -1), 0.5, ROUGE_PALE),
                ("ROUNDEDCORNERS", [4, 4, 4, 4]),
            ]))
            row.append(cell)
        data_rows.append(row)

    t = Table(data_rows, colWidths=[(17 * cm / ncols)] * ncols)
    t.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


# ─── Barre de priorité ────────────────────────────────────────────────────────
def _priority_badge(prio_str):
    """Badge coloré pour la priorité."""
    col = _get_prio_color(prio_str)
    style = ParagraphStyle(
        "PB", fontName="Helvetica-Bold", fontSize=11,
        textColor=col, alignment=TA_LEFT,
        borderPad=6,
    )
    return Paragraph(f"<b>{prio_str}</b>", style)


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT: Fiche Utilisateur PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_fiche_utilisateur_pdf(row: pd.Series, imp_data: pd.DataFrame = None) -> bytes:
    """Génère une fiche utilisateur PDF professionnelle."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    styles = _make_styles()
    story = []

    nom = _val(row.get("utilisateur"))
    direction = _val(row.get("direction"))
    site = _val(row.get("site"))
    type_pc = _val(row.get("type_pc"))
    n_inv = _val(row.get("n_inventaire"))
    n_serie = _val(row.get("n_serie"))
    descriptions = _val(row.get("descriptions"))
    annee = _val(row.get("annee_acquisition"))
    anciennete = _val(row.get("anciennete"))
    cat_age = _val(row.get("categorie_age"))
    ram = _val(row.get("ram"))
    proc = _val(row.get("processeur"))
    hdd = _val(row.get("hdd"))
    prio = _val(row.get("priorite"), "🟢 Opérationnel")
    rec = _val(row.get("recommandation"), "Aucune action immédiate")
    raisons = _val(row.get("raisons"), "Équipement conforme")
    waterp = _val(row.get("waterp"), "NON")
    sensible = _val(row.get("poste_sensible"), "NON")
    obs = _val(row.get("observation"))

    # ── En-tête
    story.append(_header_block(styles,
        "DSI AMENDIS — FICHE TECHNIQUE UTILISATEUR",
        f"Site Tétouan {YEAR_STR}  |  Direction des Systèmes d'Information"))
    story.append(_meta_line(styles,
        f"Généré le: {DATE_STR}",
        "CONFIDENTIEL — Usage interne DSI"))
    story.append(Spacer(1, 4))

    # ── Identité
    story.append(Paragraph("IDENTITÉ UTILISATEUR", styles["SectionTitle"]))
    story.append(_info_card([
        ("NOM COMPLET", nom),
        ("DIRECTION", direction),
        ("SITE", site),
        ("TYPE DE POSTE", type_pc),
        ("POSTE SENSIBLE", sensible),
        ("WATERP", waterp),
    ], ncols=3))
    story.append(Spacer(1, 8))

    # ── PC Technique
    story.append(Paragraph("FICHE TECHNIQUE PC", styles["SectionTitle"]))
    story.append(_info_card([
        ("N° INVENTAIRE", n_inv),
        ("N° SÉRIE", n_serie),
        ("MODÈLE / DESCRIPTION", descriptions),
        ("ANNÉE D'ACQUISITION", annee),
        ("ANCIENNETÉ", f"{anciennete} ans"),
        ("CATÉGORIE ÂGE", cat_age),
        ("RAM", ram),
        ("PROCESSEUR", proc),
        ("DISQUE DUR (HDD)", hdd),
    ], ncols=3))
    story.append(Spacer(1, 8))

    # ── Statut DSI
    story.append(Paragraph("STATUT DSI & RECOMMANDATION", styles["SectionTitle"]))
    prio_color = _get_prio_color(prio)

    status_data = [
        [
            Paragraph("PRIORITÉ DE REMPLACEMENT", styles["FieldLabel"]),
            Paragraph("RECOMMANDATION DSI", styles["FieldLabel"]),
            Paragraph("JUSTIFICATION", styles["FieldLabel"]),
        ],
        [
            Paragraph(f"<b>{prio}</b>",
                       ParagraphStyle("PS", fontName="Helvetica-Bold",
                                      fontSize=12, textColor=prio_color)),
            Paragraph(rec, styles["BodyText2"]),
            Paragraph(raisons, styles["BodyText2"]),
        ],
    ]
    status_t = Table(status_data, colWidths=[5 * cm, 6 * cm, 6 * cm])
    status_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ROUGE),
        ("TEXTCOLOR", (0, 0), (-1, 0), BLANC),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("BACKGROUND", (0, 1), (-1, -1), ROUGE_BG),
        ("GRID", (0, 0), (-1, -1), 0.5, ROUGE_PALE),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(status_t)
    story.append(Spacer(1, 8))

    # ── Recommandation matériel
    story.append(Paragraph("PC RECOMMANDÉ PAR LA DSI", styles["SectionTitle"]))

    is_sensible = "OUI" in str(sensible).upper()
    is_portable = "PORTABLE" in str(type_pc).upper()

    if is_sensible:
        pc_rec = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs = "Intel Core i7, 16 GO RAM, 512 GO SSD, Windows 11 Pro"
        justif = "Poste sensible — sécurité renforcée requise"
    elif is_portable:
        pc_rec = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        justif = "Utilisateur nomade — PC portable recommandé"
    else:
        pc_rec = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        justif = "Poste fixe standard"

    rec_data = [
        ["PC RECOMMANDÉ", "SPÉCIFICATIONS", "JUSTIFICATION"],
        [pc_rec, specs, justif],
    ]
    rec_t = Table(rec_data, colWidths=[5.5 * cm, 6.5 * cm, 5 * cm])
    rec_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), GRIS_FONCE),
        ("TEXTCOLOR", (0, 0), (-1, 0), BLANC),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, 1), 8),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("BACKGROUND", (0, 1), (-1, -1), BLANC),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(rec_t)

    # ── Observation
    if obs and obs != "N/A":
        story.append(Spacer(1, 8))
        story.append(Paragraph("OBSERVATION", styles["SectionTitle"]))
        obs_t = Table([[Paragraph(obs, styles["BodyText2"])]], colWidths=[17 * cm])
        obs_t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFFDE7")),
            ("BOX", (0, 0), (-1, -1), 1, JAUNE),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(obs_t)

    # ── Imprimantes
    if imp_data is not None and len(imp_data) > 0:
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"IMPRIMANTES — DIRECTION {direction}", styles["SectionTitle"]))
        imp_cols = {
            "modele": "Modèle",
            "type_imp": "Type",
            "site": "Site",
            "n_inventaire": "N° Inv.",
            "n_serie": "N° Série",
            "annee_acquisition": "Année",
        }
        story.append(_data_table(imp_data, imp_cols, max_rows=20))

    # ── Pied de page
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=1, color=ROUGE_PALE))
    footer_data = [[
        Paragraph(f"DSI Amendis — Site Tétouan {YEAR_STR}",
                   ParagraphStyle("FL", fontName="Helvetica", fontSize=7, textColor=GRIS_MOY)),
        Paragraph(f"Généré le {DATE_STR}  |  Document confidentiel",
                   ParagraphStyle("FR", fontName="Helvetica", fontSize=7,
                                  textColor=GRIS_MOY, alignment=TA_RIGHT)),
    ]]
    ft = Table(footer_data, colWidths=[8.5 * cm, 8.5 * cm])
    ft.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(ft)

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT: Rapport Global PC PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_pc_pdf(df: pd.DataFrame) -> bytes:
    """Génère un rapport global PC en PDF."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    styles = _make_styles()
    story = []

    total = len(df)
    bureau = df[df.get("type_pc", pd.Series()).str.upper() == "BUREAU"].shape[0] if "type_pc" in df.columns else 0
    portable = df[df.get("type_pc", pd.Series()).str.upper() == "PORTABLE"].shape[0] if "type_pc" in df.columns else 0
    urgent = df[df.get("priorite", pd.Series()) == "🔴 Urgent"].shape[0] if "priorite" in df.columns else 0
    n_dir = df["direction"].nunique() if "direction" in df.columns else 0

    # En-tête
    story.append(_header_block(styles,
        "DSI AMENDIS — RAPPORT INVENTAIRE PC",
        f"Site Tétouan {YEAR_STR}  |  Direction des Systèmes d'Information"))
    story.append(_meta_line(styles,
        f"Date: {DATE_STR}",
        f"Total: {total} équipements analysés"))
    story.append(Spacer(1, 6))

    # KPIs
    story.append(Paragraph("INDICATEURS CLÉS DU PARC", styles["SectionTitle"]))
    story.append(_kpi_row([
        ("Total PC", str(total), ROUGE),
        ("PC Bureau", str(bureau), BLEU),
        ("PC Portables", str(portable), GRIS_FONCE),
        ("Remplacement Urgent", str(urgent), ROUGE),
        ("Directions", str(n_dir), VERT),
    ]))
    story.append(Spacer(1, 10))

    # Répartition par direction
    if "direction" in df.columns:
        story.append(Paragraph("RÉPARTITION PAR DIRECTION", styles["SectionTitle"]))
        dir_df = df.groupby("direction").size().reset_index(name="Nombre")
        if "type_pc" in df.columns:
            bureau_df = df[df["type_pc"] == "BUREAU"].groupby("direction").size().reset_index(name="Bureau")
            port_df = df[df["type_pc"] == "PORTABLE"].groupby("direction").size().reset_index(name="Portable")
            dir_df = dir_df.merge(bureau_df, on="direction", how="left")
            dir_df = dir_df.merge(port_df, on="direction", how="left")
            dir_df = dir_df.fillna(0).astype({"Bureau": int, "Portable": int})
        dir_df = dir_df.sort_values("Nombre", ascending=False)

        dir_cols = {"direction": "Direction", "Nombre": "Total"}
        if "Bureau" in dir_df.columns:
            dir_cols.update({"Bureau": "Bureau", "Portable": "Portable"})
        story.append(_data_table(dir_df, dir_cols))
        story.append(Spacer(1, 10))

    # Analyse par âge
    if "categorie_age" in df.columns:
        story.append(Paragraph("ANALYSE PAR ANCIENNETÉ", styles["SectionTitle"]))
        age_df = df["categorie_age"].value_counts().reset_index()
        age_df.columns = ["categorie_age", "Nombre"]
        age_df["Pourcentage"] = (age_df["Nombre"] / total * 100).round(1).astype(str) + "%"
        story.append(_data_table(age_df, {
            "categorie_age": "Catégorie d'âge",
            "Nombre": "Nombre",
            "Pourcentage": "% du parc",
        }))
        story.append(Spacer(1, 10))

    # Recommandations
    if "priorite" in df.columns:
        story.append(Paragraph("RECOMMANDATIONS DE REMPLACEMENT", styles["SectionTitle"]))
        prio_df = df["priorite"].value_counts().reset_index()
        prio_df.columns = ["priorite", "Nombre"]
        prio_df["Pourcentage"] = (prio_df["Nombre"] / total * 100).round(1).astype(str) + "%"
        story.append(_data_table(prio_df, {
            "priorite": "Priorité",
            "Nombre": "Nombre de PC",
            "Pourcentage": "% du parc",
        }))
        story.append(Spacer(1, 10))

        # Détail urgent
        urgent_df = df[df["priorite"] == "🔴 Urgent"]
        if len(urgent_df) > 0:
            story.append(PageBreak())
            story.append(Paragraph("DÉTAIL — PC À REMPLACER EN URGENCE", styles["SectionTitle"]))
            story.append(Paragraph(
                f"<font color='#C8102E'><b>{len(urgent_df)} PC</b></font> nécessitent un remplacement immédiat.",
                styles["BodyText2"]
            ))
            story.append(Spacer(1, 6))
            urg_cols = {
                "utilisateur": "Utilisateur",
                "direction": "Direction",
                "type_pc": "Type",
                "annee_acquisition": "Année",
                "ram": "RAM",
                "processeur": "Processeur",
                "raisons": "Raisons",
            }
            story.append(_data_table(urgent_df, urg_cols, max_rows=50))

    # RAM
    if "ram_go" in df.columns:
        story.append(Spacer(1, 10))
        story.append(Paragraph("DISTRIBUTION RAM", styles["SectionTitle"]))
        ram_df = df["ram_go"].dropna().value_counts().sort_index().reset_index()
        ram_df.columns = ["ram_go", "Nombre"]
        ram_df["RAM"] = ram_df["ram_go"].apply(lambda x: f"{int(x)} GO")
        story.append(_data_table(ram_df[["RAM", "Nombre"]], {"RAM": "RAM", "Nombre": "Nombre"}))

    # Footer
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=1, color=ROUGE_PALE))
    story.append(Paragraph(
        f"DSI Amendis Tétouan {YEAR_STR}  —  Document confidentiel  —  Généré le {DATE_STR}",
        styles["Caption"]
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT: Rapport Imprimantes PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_imp_pdf(df: pd.DataFrame) -> bytes:
    """Génère un rapport imprimantes en PDF."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    styles = _make_styles()
    story = []

    total = len(df)
    n_dir = df["direction"].nunique() if "direction" in df.columns else 0

    story.append(_header_block(styles,
        "DSI AMENDIS — RAPPORT INVENTAIRE IMPRIMANTES",
        f"Site Tétouan {YEAR_STR}  |  Direction des Systèmes d'Information"))
    story.append(_meta_line(styles, f"Date: {DATE_STR}", f"Total: {total} imprimantes"))
    story.append(Spacer(1, 6))

    # KPIs
    story.append(_kpi_row([
        ("Total Imprimantes", str(total), ROUGE),
        ("Directions", str(n_dir), BLEU),
    ]))
    story.append(Spacer(1, 10))

    # Par direction
    if "direction" in df.columns:
        story.append(Paragraph("RÉPARTITION PAR DIRECTION", styles["SectionTitle"]))
        dir_df = df.groupby("direction").size().reset_index(name="Nombre")
        dir_df = dir_df.sort_values("Nombre", ascending=False)
        story.append(_data_table(dir_df, {"direction": "Direction", "Nombre": "Nombre"}))
        story.append(Spacer(1, 10))

    # Par type
    if "type_imp" in df.columns:
        story.append(Paragraph("RÉPARTITION PAR TYPE", styles["SectionTitle"]))
        type_df = df["type_imp"].value_counts().reset_index()
        type_df.columns = ["type_imp", "Nombre"]
        story.append(_data_table(type_df, {"type_imp": "Type", "Nombre": "Nombre"}))
        story.append(Spacer(1, 10))

    # Inventaire complet
    story.append(PageBreak())
    story.append(Paragraph("INVENTAIRE COMPLET", styles["SectionTitle"]))
    imp_cols = {
        "utilisateur": "Utilisateur",
        "direction": "Direction",
        "site": "Site",
        "modele": "Modèle",
        "type_imp": "Type",
        "n_inventaire": "N° Inv.",
        "annee_acquisition": "Année",
    }
    story.append(_data_table(df, imp_cols))

    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=1, color=ROUGE_PALE))
    story.append(Paragraph(
        f"DSI Amendis Tétouan {YEAR_STR}  —  Document confidentiel  —  Généré le {DATE_STR}",
        styles["Caption"]
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT: Rapport Général PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_general_pdf(df_inv: pd.DataFrame, df_imp: pd.DataFrame) -> bytes:
    """Génère un rapport général PC + Imprimantes."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
    )
    styles = _make_styles()
    story = []

    total_pc = len(df_inv) if df_inv is not None else 0
    total_imp = len(df_imp) if df_imp is not None else 0

    story.append(_header_block(styles,
        "DSI AMENDIS — RAPPORT GÉNÉRAL INVENTAIRE",
        f"Site Tétouan {YEAR_STR}  |  Parc Informatique Complet"))
    story.append(_meta_line(styles, f"Date: {DATE_STR}", "CONFIDENTIEL DSI"))
    story.append(Spacer(1, 8))

    story.append(_kpi_row([
        ("Total PC", str(total_pc), ROUGE),
        ("Total Imprimantes", str(total_imp), BLEU),
        ("Total Équipements", str(total_pc + total_imp), GRIS_FONCE),
    ]))
    story.append(Spacer(1, 12))

    # Synthèse par direction
    if df_inv is not None and "direction" in df_inv.columns:
        story.append(Paragraph("SYNTHÈSE PAR DIRECTION — PC", styles["SectionTitle"]))
        syn = df_inv.groupby("direction").agg(
            Total=("utilisateur", "count"),
        ).reset_index()
        if "priorite" in df_inv.columns:
            urg = df_inv[df_inv["priorite"] == "🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            syn = syn.merge(urg, on="direction", how="left").fillna(0).astype({"Urgents": int})
        if "categorie_age" in df_inv.columns:
            old = df_inv[df_inv["categorie_age"] == "Plus de 10 ans"].groupby("direction").size().reset_index(name=">10 ans")
            syn = syn.merge(old, on="direction", how="left").fillna(0).astype({">10 ans": int})
        syn = syn.sort_values("Total", ascending=False)
        story.append(_data_table(syn, {c: c if c != "direction" else "Direction" for c in syn.columns}))

    if df_imp is not None and "direction" in df_imp.columns:
        story.append(Spacer(1, 10))
        story.append(Paragraph("SYNTHÈSE PAR DIRECTION — IMPRIMANTES", styles["SectionTitle"]))
        imp_syn = df_imp.groupby("direction").size().reset_index(name="Total")
        imp_syn = imp_syn.sort_values("Total", ascending=False)
        story.append(_data_table(imp_syn, {"direction": "Direction", "Total": "Total Imprimantes"}))

    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=1, color=ROUGE_PALE))
    story.append(Paragraph(
        f"DSI Amendis Tétouan {YEAR_STR}  —  Document confidentiel  —  Généré le {DATE_STR}",
        styles["Caption"]
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()
