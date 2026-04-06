"""
pdf_export.py — Génération PDF professionnelle Amendis DSI v4
- Rapport PC global
- Rapport Imprimantes global
- Rapport filtré par direction
- Fiche technique utilisateur
"""
import io
from datetime import datetime
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)

ROUGE       = colors.HexColor("#C8102E")
ROUGE_CLAIR = colors.HexColor("#E8415B")
ROUGE_PALE  = colors.HexColor("#F5A0AE")
ROUGE_BG    = colors.HexColor("#FFF0F2")
GRIS_FONCE  = colors.HexColor("#2D2D2D")
GRIS_MOY    = colors.HexColor("#6B6B6B")
GRIS_CLAIR  = colors.HexColor("#F2F2F2")
BLANC       = colors.white
VERT        = colors.HexColor("#27AE60")
ORANGE      = colors.HexColor("#E67E22")
BLEU        = colors.HexColor("#2980B9")
JAUNE       = colors.HexColor("#F1C40F")
DARK_TEAL   = colors.HexColor("#1B4F72")

DATE_STR  = datetime.now().strftime("%d/%m/%Y")
YEAR_STR  = "2026"
PAGE_W    = 17 * cm


def _prio_color(p):
    p = str(p).lower()
    if "urgent"      in p: return ROUGE
    if "prioritaire" in p: return ORANGE
    if "surveiller"  in p: return JAUNE
    if "opérationnel" in p or "operationnel" in p: return VERT
    return GRIS_MOY


def _v(v, default="—"):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return s if s not in ("nan", "NaT", "None", "") else default


def _styles():
    s = getSampleStyleSheet()
    def add(name, **kwargs):
        if name not in s.byName:
            s.add(ParagraphStyle(name, **kwargs))
    add("Title2",    fontName="Helvetica-Bold",  fontSize=20, textColor=BLANC,      alignment=TA_CENTER, spaceAfter=4)
    add("Sub2",      fontName="Helvetica",        fontSize=9,  textColor=BLANC,      alignment=TA_CENTER)
    add("Sec",       fontName="Helvetica-Bold",  fontSize=12, textColor=ROUGE,      spaceBefore=12, spaceAfter=6, borderPad=4)
    add("FL",        fontName="Helvetica",        fontSize=7,  textColor=GRIS_MOY,   alignment=TA_LEFT)
    add("FR",        fontName="Helvetica",        fontSize=7,  textColor=GRIS_MOY,   alignment=TA_RIGHT)
    add("FLabel",    fontName="Helvetica-Bold",  fontSize=7,  textColor=GRIS_MOY)
    add("FVal",      fontName="Helvetica-Bold",  fontSize=10, textColor=ROUGE)
    add("Body2",     fontName="Helvetica",        fontSize=8,  textColor=GRIS_FONCE, spaceAfter=3)
    add("Cap",       fontName="Helvetica-Oblique",fontSize=7, textColor=GRIS_MOY,   alignment=TA_CENTER)
    add("TH",        fontName="Helvetica-Bold",  fontSize=7,  textColor=BLANC,      alignment=TA_CENTER)
    add("TD",        fontName="Helvetica",        fontSize=7,  textColor=GRIS_FONCE, alignment=TA_LEFT)
    add("TDC",       fontName="Helvetica",        fontSize=7,  textColor=GRIS_FONCE, alignment=TA_CENTER)
    return s


def _header(title, subtitle=""):
    s = _styles()
    data = [[Paragraph(title, s["Title2"])]]
    if subtitle:
        data.append([Paragraph(subtitle, s["Sub2"])])
    t = Table(data, colWidths=[PAGE_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), ROUGE),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING", (0,0), (-1,-1), 14),
        ("BOTTOMPADDING", (0,-1), (-1,-1), 12),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
    ]))
    return t


def _meta(left, right):
    s = _styles()
    t = Table([[Paragraph(left, s["FL"]), Paragraph(right, s["FR"])]],
              colWidths=[PAGE_W/2, PAGE_W/2])
    t.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    return t


def _kpi_row(items):
    """items = list of (label, value, color)"""
    if not items: return Spacer(1,4)
    n = len(items)
    cw = PAGE_W / n
    row = []
    for label, val, col in items:
        cell = Table([
            [Paragraph(f"<b>{val}</b>",
                       ParagraphStyle("kv", fontName="Helvetica-Bold", fontSize=18,
                                      textColor=col, alignment=TA_CENTER))],
            [Paragraph(label,
                       ParagraphStyle("kl", fontName="Helvetica", fontSize=7,
                                      textColor=GRIS_MOY, alignment=TA_CENTER))],
        ], colWidths=[cw-8])
        cell.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLANC),
            ("LINEBELOW",(0,0),(-1,0),2,col),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,-1),(-1,-1),8),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        row.append(cell)
    outer = Table([row], colWidths=[cw]*n)
    outer.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),.5,GRIS_CLAIR),
        ("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    return outer


def _data_table(df, col_labels, max_rows=None):
    s = _styles()
    if df is None or len(df) == 0:
        return Paragraph("Aucune donnée.", s["Body2"])
    cols = [c for c in col_labels if c in df.columns]
    labels = [col_labels[c] for c in cols]
    if max_rows: df = df.head(max_rows)

    header = [Paragraph(f"<b>{l}</b>", s["TH"]) for l in labels]
    rows = [header]
    for _, row in df.iterrows():
        r = []
        for c in cols:
            v = _v(row.get(c))
            if c == "priorite":
                pc = _prio_color(v)
                cell = Paragraph(f"<b>{v}</b>",
                                  ParagraphStyle("pc", fontName="Helvetica-Bold",
                                                 fontSize=7, textColor=pc))
            else:
                cell = Paragraph(v, s["TD"])
            r.append(cell)
        rows.append(r)

    n = len(cols)
    cw = PAGE_W / n
    t = Table(rows, colWidths=[cw]*n, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),ROUGE),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#E0E0E0")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[BLANC,GRIS_CLAIR]),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),3),
        ("RIGHTPADDING",(0,0),(-1,-1),3),
    ]))
    return t


def _info_grid(fields, ncols=3):
    s = _styles()
    n = len(fields); cw = PAGE_W / ncols
    data_rows = []
    for i in range(0, n, ncols):
        chunk = fields[i:i+ncols]
        while len(chunk) < ncols: chunk.append(("",""))
        row = []
        for label, val in chunk:
            cell = Table([
                [Paragraph(str(label), s["FLabel"])],
                [Paragraph(str(val),   s["FVal"])],
            ], colWidths=[cw-8])
            cell.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),BLANC),
                ("BOX",(0,0),(-1,-1),.5,ROUGE_PALE),
                ("TOPPADDING",(0,0),(-1,-1),6),
                ("BOTTOMPADDING",(0,-1),(-1,-1),6),
                ("LEFTPADDING",(0,0),(-1,-1),8),
                ("RIGHTPADDING",(0,0),(-1,-1),8),
            ]))
            row.append(cell)
        data_rows.append(row)
    t = Table(data_rows, colWidths=[cw]*ncols)
    t.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),3),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),3),
        ("RIGHTPADDING",(0,0),(-1,-1),3),
    ]))
    return t


def _footer(story):
    s = _styles()
    story.append(Spacer(1,12))
    story.append(HRFlowable(width="100%", thickness=.5, color=ROUGE_PALE))
    t = Table([[
        Paragraph(f"DSI Amendis — Site Tétouan {YEAR_STR}", s["FL"]),
        Paragraph(f"Généré le {DATE_STR}  |  Confidentiel", s["FR"]),
    ]], colWidths=[PAGE_W/2, PAGE_W/2])
    t.setStyle(TableStyle([("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.append(t)


def _doc(buf):
    return SimpleDocTemplate(buf, pagesize=A4,
                              leftMargin=2*cm, rightMargin=2*cm,
                              topMargin=2*cm, bottomMargin=2*cm)


# ═══════════════════════════════════════════════════════════════════════════════
#  FICHE UTILISATEUR PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_fiche_utilisateur_pdf(row: pd.Series, imp_data=None,
                                  ancien_pc: pd.Series = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    nom      = _v(row.get("utilisateur"))
    dir_     = _v(row.get("direction"))
    site     = _v(row.get("site"))
    type_pc  = _v(row.get("type_pc"))
    n_inv    = _v(row.get("n_inventaire"))
    n_serie  = _v(row.get("n_serie"))
    desc     = _v(row.get("descriptions"))
    annee    = _v(row.get("annee_acquisition"))
    anc      = _v(row.get("anciennete"))
    cat_age  = _v(row.get("categorie_age"))
    ram      = _v(row.get("ram"))
    proc     = _v(row.get("processeur"))
    hdd      = _v(row.get("hdd"))
    prio     = _v(row.get("priorite"), "🟢 Opérationnel")
    rec      = _v(row.get("recommandation"), "Aucune action")
    raisons  = _v(row.get("raisons"), "Équipement conforme")
    waterp   = _v(row.get("waterp"), "NON")
    sensible = _v(row.get("poste_sensible"), "NON")
    obs      = _v(row.get("observation"))

    story.append(_header(f"FICHE TECHNIQUE — {nom}",
                         f"DSI Amendis Tétouan {YEAR_STR}"))
    story.append(_meta(f"Direction: {dir_}  |  Site: {site}", f"Généré: {DATE_STR}"))
    story.append(Spacer(1,6))

    # Identité
    story.append(Paragraph("IDENTITÉ", s["Sec"]))
    story.append(_info_grid([
        ("NOM COMPLET", nom), ("DIRECTION", dir_), ("SITE", site),
        ("TYPE PC", type_pc), ("POSTE SENSIBLE", sensible), ("WATERP", waterp),
    ]))
    story.append(Spacer(1,8))

    # PC actuel
    story.append(Paragraph("PC ACTUEL — FICHE TECHNIQUE", s["Sec"]))
    story.append(_info_grid([
        ("N° INVENTAIRE", n_inv), ("N° SÉRIE", n_serie), ("MODÈLE", desc),
        ("ANNÉE ACQUISITION", annee), ("ANCIENNETÉ", f"{anc} ans"), ("CATÉGORIE", cat_age),
        ("RAM", ram), ("PROCESSEUR", proc), ("DISQUE DUR", hdd),
    ]))
    story.append(Spacer(1,8))

    # Statut
    story.append(Paragraph("STATUT DSI", s["Sec"]))
    pc = _prio_color(prio)
    status = Table([
        [Paragraph("PRIORITÉ", ParagraphStyle("p1",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER)),
         Paragraph("RECOMMANDATION", ParagraphStyle("p2",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER)),
         Paragraph("JUSTIFICATION", ParagraphStyle("p3",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER))],
        [Paragraph(f"<b>{prio}</b>", ParagraphStyle("pv",fontName="Helvetica-Bold",fontSize=10,textColor=pc)),
         Paragraph(rec, s["Body2"]),
         Paragraph(raisons, s["Body2"])],
    ], colWidths=[5*cm, 6*cm, 6*cm])
    status.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),ROUGE), ("GRID",(0,0),(-1,-1),.5,ROUGE_PALE),
        ("BACKGROUND",(0,1),(-1,-1),ROUGE_BG),
        ("TOPPADDING",(0,0),(-1,-1),8), ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),8), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(status)
    story.append(Spacer(1,8))

    # Ancien PC si disponible
    if ancien_pc is not None:
        story.append(Paragraph("ANCIEN PC (HISTORIQUE)", s["Sec"]))
        story.append(_info_grid([
            ("ANCIEN MODÈLE", _v(ancien_pc.get("descriptions"))),
            ("ANCIEN N° INV",  _v(ancien_pc.get("n_inventaire"))),
            ("ANNÉE",          _v(ancien_pc.get("annee_acquisition"))),
            ("RAM",            _v(ancien_pc.get("ram"))),
            ("PROCESSEUR",     _v(ancien_pc.get("processeur"))),
            ("HDD",            _v(ancien_pc.get("hdd"))),
        ]))
        story.append(Spacer(1,8))

    # PC recommandé
    story.append(Paragraph("PC RECOMMANDÉ PAR LA DSI", s["Sec"]))
    is_sens = "OUI" in str(sensible).upper()
    is_port = "PORTABLE" in str(type_pc).upper()
    if is_sens:
        pc_rec   = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_r  = "Intel Core i7-1355U, 16 GO DDR5, 512 GO SSD NVMe, Win 11 Pro"
        justif   = "Poste sensible — sécurité renforcée"
    elif is_port:
        pc_rec   = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_r  = "Intel Core i5-1335U, 8 GO DDR4, 256 GO SSD, Win 11"
        justif   = "Utilisateur nomade — portable recommandé"
    else:
        pc_rec   = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_r  = "Intel Core i5-12500, 8 GO DDR4, 256 GO SSD, Win 11"
        justif   = "Poste fixe standard"

    is_urg = any(x in prio for x in ["Urgent", "Prioritaire"])
    rec_t = Table([
        [Paragraph("PC RECOMMANDÉ", ParagraphStyle("r1",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER)),
         Paragraph("SPÉCIFICATIONS", ParagraphStyle("r2",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER)),
         Paragraph("JUSTIFICATION", ParagraphStyle("r3",fontName="Helvetica-Bold",fontSize=7,textColor=BLANC,alignment=TA_CENTER))],
        [Paragraph(pc_rec, s["Body2"]), Paragraph(specs_r, s["Body2"]), Paragraph(justif, s["Body2"])],
    ], colWidths=[5.5*cm, 6.5*cm, 5*cm])
    rec_t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),GRIS_FONCE), ("BACKGROUND",(0,1),(-1,-1),BLANC),
        ("GRID",(0,0),(-1,-1),.3,GRIS_CLAIR),
        ("TOPPADDING",(0,0),(-1,-1),8), ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),8), ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(rec_t)

    # Observation
    if obs and obs != "—":
        story.append(Spacer(1,6))
        obs_t = Table([[Paragraph(f"<b>Observation:</b> {obs}", s["Body2"])]],
                       colWidths=[PAGE_W])
        obs_t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FFFDE7")),
            ("BOX",(0,0),(-1,-1),1,JAUNE),
            ("TOPPADDING",(0,0),(-1,-1),8), ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING",(0,0),(-1,-1),10),
        ]))
        story.append(obs_t)

    # Imprimantes
    if imp_data is not None and len(imp_data) > 0:
        story.append(Spacer(1,8))
        story.append(Paragraph(f"IMPRIMANTES — DIRECTION {dir_}", s["Sec"]))
        story.append(_data_table(imp_data, {
            "modele":"Modèle","type_imp":"Type","site":"Site",
            "n_inventaire":"N° Inv.","n_serie":"N° Série","annee_acquisition":"Année"
        }, max_rows=15))

    _footer(story)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT PC GLOBAL
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_pc_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_header(f"RAPPORT INVENTAIRE PC{titre_dir}",
                         f"DSI Amendis Tétouan {YEAR_STR}"))
    story.append(_meta(f"Date: {DATE_STR}", f"Total: {len(df)} PC analysés"))
    story.append(Spacer(1,8))

    total   = len(df)
    bureau  = df[df.get("type_pc", pd.Series(dtype=str)).str.upper() == "BUREAU"].shape[0]  if "type_pc" in df.columns else 0
    port    = df[df.get("type_pc", pd.Series(dtype=str)).str.upper() == "PORTABLE"].shape[0] if "type_pc" in df.columns else 0
    chrom   = df[df.get("type_pc", pd.Series(dtype=str)).str.upper() == "CHROMEBOOK"].shape[0] if "type_pc" in df.columns else 0
    urgent  = df[df.get("priorite", pd.Series(dtype=str)) == "🔴 Urgent"].shape[0]           if "priorite" in df.columns else 0
    n_dir   = df["direction"].nunique() if "direction" in df.columns else 0

    story.append(_kpi_row([
        ("Total PC", str(total), ROUGE), ("Bureau", str(bureau), BLEU),
        ("Portable", str(port), GRIS_FONCE), ("Chromebook", str(chrom), DARK_TEAL),
        ("Urgents", str(urgent), ROUGE), ("Directions", str(n_dir), VERT),
    ]))
    story.append(Spacer(1,10))

    # Par direction
    if "direction" in df.columns and not direction:
        story.append(Paragraph("RÉPARTITION PAR DIRECTION", s["Sec"]))
        agg = df.groupby("direction").size().reset_index(name="Total")
        if "type_pc" in df.columns:
            for tp, col_name in [("BUREAU","Bureau"),("PORTABLE","Portable"),("CHROMEBOOK","Chromebook")]:
                sub = df[df["type_pc"]==tp].groupby("direction").size().reset_index(name=col_name)
                agg = agg.merge(sub, on="direction", how="left")
            agg = agg.fillna(0)
            for c in ["Bureau","Portable","Chromebook"]:
                if c in agg.columns: agg[c] = agg[c].astype(int)
        if "priorite" in df.columns:
            urg = df[df["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)
        agg = agg.sort_values("Total", ascending=False)
        cols_map = {"direction":"Direction","Total":"Total","Bureau":"Bureau","Portable":"Portable","Chromebook":"Chromebook","Urgents":"Urgents"}
        story.append(_data_table(agg, {k:v for k,v in cols_map.items() if k in agg.columns}))
        story.append(Spacer(1,10))

    # Par ancienneté
    if "categorie_age" in df.columns:
        story.append(Paragraph("ANALYSE PAR ANCIENNETÉ", s["Sec"]))
        age_df = df["categorie_age"].value_counts().reset_index()
        age_df.columns = ["categorie_age", "Nombre"]
        age_df["Pourcentage"] = (age_df["Nombre"]/total*100).round(1).astype(str)+"%"
        story.append(_data_table(age_df, {"categorie_age":"Catégorie","Nombre":"Nombre","Pourcentage":"%"}))
        story.append(Spacer(1,10))

    # Recommandations
    if "priorite" in df.columns:
        story.append(Paragraph("RECOMMANDATIONS DE REMPLACEMENT", s["Sec"]))
        prio_df = df["priorite"].value_counts().reset_index()
        prio_df.columns = ["priorite","Nombre"]
        prio_df["Pourcentage"] = (prio_df["Nombre"]/total*100).round(1).astype(str)+"%"
        story.append(_data_table(prio_df, {"priorite":"Priorité","Nombre":"Nombre","Pourcentage":"%"}))
        story.append(Spacer(1,10))

        # Détail urgents
        urg_df = df[df["priorite"]=="🔴 Urgent"]
        if len(urg_df) > 0:
            story.append(PageBreak())
            story.append(Paragraph(f"DÉTAIL PC URGENTS ({len(urg_df)})", s["Sec"]))
            story.append(_data_table(urg_df, {
                "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
                "type_pc":"Type","annee_acquisition":"Année","ram":"RAM",
                "processeur":"Processeur","raisons":"Raisons"
            }, max_rows=100))

    # Inventaire complet par direction (si filtré)
    if direction:
        story.append(PageBreak())
        story.append(Paragraph(f"INVENTAIRE COMPLET — {direction}", s["Sec"]))
        story.append(_data_table(df, {
            "utilisateur":"Utilisateur","site":"Site","type_pc":"Type",
            "n_inventaire":"N° Inv.","annee_acquisition":"Année","ram":"RAM",
            "processeur":"Processeur","hdd":"HDD","priorite":"Priorité","observation":"Observation"
        }))

    _footer(story)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT IMPRIMANTES GLOBAL
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_imp_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_header(f"RAPPORT IMPRIMANTES{titre_dir}",
                         f"DSI Amendis Tétouan {YEAR_STR}"))
    story.append(_meta(f"Date: {DATE_STR}", f"Total: {len(df)} imprimantes"))
    story.append(Spacer(1,8))

    total = len(df)
    n_dir = df["direction"].nunique() if "direction" in df.columns else 0
    story.append(_kpi_row([
        ("Total Imprimantes", str(total), ROUGE),
        ("Directions", str(n_dir), BLEU),
    ]))
    story.append(Spacer(1,10))

    if "direction" in df.columns and not direction:
        story.append(Paragraph("RÉPARTITION PAR DIRECTION", s["Sec"]))
        dir_df = df.groupby("direction").size().reset_index(name="Nombre")
        dir_df = dir_df.sort_values("Nombre", ascending=False)
        if "type_imp" in df.columns:
            for tp in df["type_imp"].dropna().unique():
                sub = df[df["type_imp"]==tp].groupby("direction").size().reset_index(name=str(tp)[:15])
                dir_df = dir_df.merge(sub, on="direction", how="left")
            dir_df = dir_df.fillna(0)
        story.append(_data_table(dir_df, {c:c if c!="direction" else "Direction" for c in dir_df.columns}))
        story.append(Spacer(1,10))

    if "type_imp" in df.columns:
        story.append(Paragraph("RÉPARTITION PAR TYPE", s["Sec"]))
        type_df = df["type_imp"].value_counts().reset_index()
        type_df.columns = ["type_imp","Nombre"]
        story.append(_data_table(type_df, {"type_imp":"Type","Nombre":"Nombre"}))
        story.append(Spacer(1,10))

    story.append(PageBreak())
    titre_inv = f"INVENTAIRE — {direction}" if direction else "INVENTAIRE COMPLET"
    story.append(Paragraph(titre_inv, s["Sec"]))
    story.append(_data_table(df, {
        "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
        "modele":"Modèle","type_imp":"Type","n_inventaire":"N° Inv.","annee_acquisition":"Année"
    }))

    _footer(story)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT GÉNÉRAL PDF
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_general_pdf(df_inv: pd.DataFrame, df_imp: pd.DataFrame,
                                direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Parc Complet"
    story.append(_header(f"RAPPORT GÉNÉRAL{titre_dir}",
                         f"DSI Amendis Tétouan {YEAR_STR}"))
    story.append(_meta(f"Date: {DATE_STR}", "CONFIDENTIEL DSI"))
    story.append(Spacer(1,8))

    total_pc  = len(df_inv) if df_inv is not None else 0
    total_imp = len(df_imp) if df_imp is not None else 0
    urgent    = df_inv[df_inv.get("priorite", pd.Series(dtype=str))=="🔴 Urgent"].shape[0] if (df_inv is not None and "priorite" in df_inv.columns) else 0

    story.append(_kpi_row([
        ("Total PC",           str(total_pc),         ROUGE),
        ("Total Imprimantes",  str(total_imp),        BLEU),
        ("Total Équipements",  str(total_pc+total_imp), GRIS_FONCE),
        ("PC Urgents",         str(urgent),            ROUGE if urgent>0 else VERT),
    ]))
    story.append(Spacer(1,12))

    if df_inv is not None and "direction" in df_inv.columns:
        story.append(Paragraph("SYNTHÈSE PAR DIRECTION — PC", s["Sec"]))
        agg = df_inv.groupby("direction").agg(Total_PC=("utilisateur","count")).reset_index()
        if "priorite" in df_inv.columns:
            urg = df_inv[df_inv["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)
        if "categorie_age" in df_inv.columns:
            old = df_inv[df_inv["categorie_age"]=="Plus de 10 ans"].groupby("direction").size().reset_index(name=">10 ans")
            agg = agg.merge(old, on="direction", how="left").fillna(0)
            agg[">10 ans"] = agg[">10 ans"].astype(int)
        if df_imp is not None and "direction" in df_imp.columns:
            ia = df_imp.groupby("direction").size().reset_index(name="Imprimantes")
            agg = agg.merge(ia, on="direction", how="left").fillna(0)
            agg["Imprimantes"] = agg["Imprimantes"].astype(int)
        agg = agg.sort_values("Total_PC", ascending=False)
        story.append(_data_table(agg, {c: c.replace("_"," ") if c!="direction" else "Direction" for c in agg.columns}))

    _footer(story)
    doc.build(story)
    buf.seek(0)
    return buf.read()
