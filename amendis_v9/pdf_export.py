"""
pdf_export.py — DSI Amendis Tetouan v9
Palette grise uniquement — aucun noir pur, aucune couleur vive
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
    PageBreak, HRFlowable
)

# ─── Palette grise — aucun noir pur, aucune couleur ──────────────────────────
ENTETE   = colors.HexColor("#2C3E50")   # fond en-tete : gris-bleu fonce
TEXTE    = colors.HexColor("#333333")   # texte principal : gris fonce
TEXTE_S  = colors.HexColor("#666666")   # texte secondaire
FOND_SEC = colors.HexColor("#EEEEEE")   # fond titres section
FOND_ALT = colors.HexColor("#F7F7F7")   # lignes alternees
BORDURE  = colors.HexColor("#CCCCCC")   # bordures
BLANC    = colors.white

DATE_STR = datetime.now().strftime("%d/%m/%Y")
YEAR_STR = "2026"
W        = 17 * cm


def _v(v, default="—"):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return s if s not in ("nan","NaT","None","") else default


def _clean(v):
    s = _v(v)
    for ch in ("🔴","🟠","🟡","🟢","🔵","⚠️","✅","🚀","📅","💾","🏢",
               "🖥️","🖨️","💻","📋","👤","🔄","🛒","📦","📝","🌐","📍"):
        s = s.replace(ch, "")
    return s.strip()


def _doc(buf):
    return SimpleDocTemplate(buf, pagesize=A4,
                              leftMargin=2*cm, rightMargin=2*cm,
                              topMargin=1.5*cm, bottomMargin=2*cm)


def _styles():
    s = getSampleStyleSheet()
    def add(name, **kw):
        if name not in s.byName:
            s.add(ParagraphStyle(name, **kw))

    add("Titre",   fontName="Helvetica-Bold",    fontSize=16, textColor=BLANC,   alignment=TA_LEFT,  leading=20)
    add("SsTitre", fontName="Helvetica",         fontSize=8,  textColor=colors.HexColor("#AAAAAA"), alignment=TA_LEFT, leading=12)
    add("LogoTxt", fontName="Helvetica-Bold",    fontSize=15, textColor=BLANC,   alignment=TA_RIGHT, leading=19)
    add("LogoSub", fontName="Helvetica",         fontSize=7,  textColor=colors.HexColor("#999999"), alignment=TA_RIGHT, leading=10)

    add("Corps",   fontName="Helvetica",         fontSize=8,  textColor=TEXTE,   leading=12, spaceAfter=3)
    add("CorpsIt", fontName="Helvetica-Oblique", fontSize=8,  textColor=TEXTE_S, leading=12)
    add("CorpsB",  fontName="Helvetica-Bold",    fontSize=8,  textColor=TEXTE,   leading=12)

    add("SecTxt",  fontName="Helvetica-Bold",    fontSize=8,  textColor=TEXTE,   leading=11)

    add("TH",      fontName="Helvetica-Bold",    fontSize=7,  textColor=BLANC,   alignment=TA_CENTER, leading=9)
    add("TD",      fontName="Helvetica",         fontSize=7,  textColor=TEXTE,   alignment=TA_LEFT,   leading=9)

    add("SK",      fontName="Helvetica-Bold",    fontSize=8,  textColor=TEXTE_S)
    add("SV",      fontName="Helvetica",         fontSize=8,  textColor=TEXTE)

    add("KV",      fontName="Helvetica-Bold",    fontSize=20, textColor=TEXTE,   alignment=TA_CENTER, leading=24, wordWrap='CJK')
    add("KL",      fontName="Helvetica",         fontSize=6,  textColor=TEXTE_S, alignment=TA_CENTER, leading=9)

    add("PiedG",   fontName="Helvetica",         fontSize=7,  textColor=colors.HexColor("#999999"), alignment=TA_LEFT)
    add("PiedD",   fontName="Helvetica",         fontSize=7,  textColor=colors.HexColor("#999999"), alignment=TA_RIGHT)

    return s


# ─── EN-TETE ──────────────────────────────────────────────────────────────────
def _entete(titre, sous_titre=""):
    s = _styles()

    gauche = [[Paragraph(titre, s["Titre"])]]
    if sous_titre:
        gauche.append([Paragraph(sous_titre, s["SsTitre"])])

    # Logo "Amendis" en un seul mot blanc — simple et propre
    droite = [
        [Paragraph("Amendis", s["LogoTxt"])],
        [Paragraph("Direction des Systemes d'Information", s["LogoSub"])],
        [Paragraph(f"Site Tetouan  |  {DATE_STR}", s["LogoSub"])],
    ]

    t_g = Table(gauche, colWidths=[W * 0.58])
    t_d = Table(droite, colWidths=[W * 0.42])

    t_g.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    t_d.setStyle(TableStyle([
        ("ALIGN",(0,0),(-1,-1),"RIGHT"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),2),
        ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))

    outer = Table([[t_g, t_d]], colWidths=[W*0.58, W*0.42])
    outer.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),ENTETE),
        ("TOPPADDING",(0,0),(-1,-1),18),
        ("BOTTOMPADDING",(0,0),(-1,-1),18),
        ("LEFTPADDING",(0,0),(0,0),20),
        ("RIGHTPADDING",(0,0),(0,0),10),
        ("LEFTPADDING",(1,0),(1,0),8),
        ("RIGHTPADDING",(1,0),(1,0),18),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return outer


def _filet():
    return HRFlowable(width=W, thickness=1.5, color=BORDURE,
                      spaceAfter=10, spaceBefore=0)


# ─── INTRO ────────────────────────────────────────────────────────────────────
def _intro(texte):
    s = _styles()
    t = Table([[Paragraph(texte, s["CorpsIt"])]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),FOND_ALT),
        ("BOX",(0,0),(-1,-1),.5,BORDURE),
        ("TOPPADDING",(0,0),(-1,-1),9),
        ("BOTTOMPADDING",(0,0),(-1,-1),9),
        ("LEFTPADDING",(0,0),(-1,-1),14),
        ("RIGHTPADDING",(0,0),(-1,-1),14),
    ]))
    return t


# ─── SECTION ──────────────────────────────────────────────────────────────────
def _section(texte):
    s = _styles()
    t = Table([[Paragraph(texte.upper(), s["SecTxt"])]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),FOND_SEC),
        ("BOX",(0,0),(-1,-1),.5,BORDURE),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),10),
    ]))
    return t


# ─── KPI — gris sobres, aucune couleur ───────────────────────────────────────
def _kpis(items, ncols=4):
    s = _styles()
    cw = W / ncols
    row = []
    for (val, label) in items:
        val_clean = _clean(str(val))
        card = Table([
            [Paragraph(val_clean, s["KV"])],
            [Paragraph(label.upper(), s["KL"])],
        ], colWidths=[cw - 10])
        card.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLANC),
            ("BOX",(0,0),(-1,-1),.5,BORDURE),
            ("TOPPADDING",(0,0),(-1,-1),12),
            ("BOTTOMPADDING",(0,0),(-1,-1),12),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        row.append(card)
    while len(row) < ncols:
        row.append(Spacer(cw-10, 1))
    outer = Table([row], colWidths=[cw]*ncols)
    outer.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),5),
        ("RIGHTPADDING",(0,0),(-1,-1),5),
    ]))
    return outer


# ─── SPEC ─────────────────────────────────────────────────────────────────────
def _spec(champs):
    s = _styles()
    data = []
    for k, v in champs:
        data.append([Paragraph(str(k), s["SK"]),
                     Paragraph(_clean(str(v)), s["SV"])])
    if not data: return Spacer(1,1)
    t = Table(data, colWidths=[W*0.36, W*0.64])
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BLANC, FOND_ALT]),
        ("GRID",(0,0),(-1,-1),.4,BORDURE),
        ("BACKGROUND",(0,0),(0,-1),FOND_SEC),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),10),
        ("RIGHTPADDING",(0,0),(-1,-1),10),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return t


# ─── TABLEAU ──────────────────────────────────────────────────────────────────
def _tableau(df, col_map, max_rows=None):
    s = _styles()
    if df is None or len(df) == 0:
        return Paragraph("Aucune donnee.", s["Corps"])
    cols   = [c for c in col_map if c in df.columns]
    labels = [col_map[c] for c in cols]
    if max_rows: df = df.head(max_rows)

    entete = [Paragraph(f"<b>{l}</b>", s["TH"]) for l in labels]
    lignes = [entete]
    for _, r in df.iterrows():
        row = []
        for c in cols:
            row.append(Paragraph(_clean(r.get(c, "—")), s["TD"]))
        lignes.append(row)

    cw = W / len(cols)
    t  = Table(lignes, colWidths=[cw]*len(cols), repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),ENTETE),
        ("GRID",(0,0),(-1,-1),.4,BORDURE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[BLANC, FOND_ALT]),
        ("ALIGN",(0,0),(-1,0),"CENTER"),
        ("ALIGN",(0,1),(-1,-1),"LEFT"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    return t


# ─── BARRES ───────────────────────────────────────────────────────────────────
def _barres(data_dict, max_items=14):
    s = _styles()
    if not data_dict: return Paragraph("Aucune donnee.", s["Corps"])
    NOM_W   = W * 0.32
    BARRE_W = W * 0.52
    VAL_W   = W * 0.16
    max_v   = max(data_dict.values(), default=1)
    lignes  = []
    for nom, val in sorted(data_dict.items(), key=lambda x:-x[1])[:max_items]:
        pct  = min(1.0, val/max_v) if max_v > 0 else 0
        fill = max(3, int(BARRE_W * pct))
        rest = max(2, int(BARRE_W) - fill)
        barre = Table([["",""]], colWidths=[fill, rest], rowHeights=[8])
        barre.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(0,0),colors.HexColor("#888888")),
            ("BACKGROUND",(1,0),(1,0),BORDURE),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ]))
        lignes.append([
            Paragraph(str(nom)[:35], s["CorpsB"]),
            barre,
            Paragraph(str(val), ParagraphStyle("vb", fontName="Helvetica-Bold",
                                                fontSize=8, textColor=TEXTE,
                                                alignment=TA_RIGHT)),
        ])
    t = Table(lignes, colWidths=[NOM_W, BARRE_W, VAL_W])
    t.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BLANC, FOND_ALT]),
        ("BOX",(0,0),(-1,-1),.4,BORDURE),
        ("LINEBELOW",(0,0),(-1,-1),.3,BORDURE),
    ]))
    return t


# ─── PIED DE PAGE ─────────────────────────────────────────────────────────────
def _pied(story, s=None):
    if s is None: s = _styles()
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width=W, thickness=.5, color=BORDURE,
                             spaceBefore=0, spaceAfter=0))
    pied = Table([[
        Paragraph("DSI — Direction des Systemes d'Information  |  Amendis Tetouan  |  Confidentiel", s["PiedG"]),
        Paragraph(f"Genere le {DATE_STR}", s["PiedD"]),
    ]], colWidths=[W*0.72, W*0.28])
    pied.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),ENTETE),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    story.append(pied)


# ═══════════════════════════════════════════════════════════════════════════════
#  FICHE UTILISATEUR
# ═══════════════════════════════════════════════════════════════════════════════
def export_fiche_utilisateur_pdf(row: pd.Series, imp_data=None,
                                  ancien_pc: pd.Series = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s   = _styles()
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
    prio     = _clean(_v(row.get("priorite"), "Operationnel"))
    rec      = _v(row.get("recommandation"), "Aucune action")
    raisons  = _v(row.get("raisons"), "Equipement conforme")
    waterp   = _v(row.get("waterp"), "NON")
    sensible = _v(row.get("poste_sensible"), "NON")
    obs      = _v(row.get("observation"))

    story.append(_entete(f"Fiche Utilisateur — {nom}",
                         f"Direction {dir_}  |  Site {site}  |  {type_pc}"))
    story.append(_filet())
    story.append(_intro(
        f"{nom} — Direction {dir_}, Site {site}. "
        f"Statut DSI : <b>{prio}</b>.  Recommandation : <b>{rec}</b>."
    ))
    story.append(Spacer(1, 12))

    story.append(_kpis([
        (dir_,    "Direction"),
        (site,    "Site"),
        (type_pc, "Type de poste"),
        (prio,    "Statut DSI"),
    ], ncols=4))
    story.append(Spacer(1, 14))

    story.append(_section("Specifications du PC Actuel"))
    story.append(Spacer(1, 5))
    story.append(_spec([
        ("N Inventaire",          n_inv),
        ("Modele / Description",  desc),
        ("N Serie",               n_serie),
        ("Annee d acquisition",   annee),
        ("Anciennete",            f"{anc} ans"),
        ("Categorie age",         cat_age),
        ("Memoire RAM",           ram),
        ("Processeur",            proc),
        ("Disque dur / SSD",      hdd),
        ("Direction",             dir_),
        ("Site",                  site),
        ("Type de poste",         type_pc),
        ("Poste sensible",        sensible),
        ("WATERP",                waterp),
    ]))
    story.append(Spacer(1, 14))

    story.append(_section("Statut et Recommandation DSI"))
    story.append(Spacer(1, 5))
    story.append(_spec([
        ("Priorite",       prio),
        ("Recommandation", rec),
        ("Raisons",        raisons),
        ("Observation",    obs),
    ]))
    story.append(Spacer(1, 14))

    is_sens = "OUI" in str(sensible).upper()
    is_port = "PORTABLE" in str(type_pc).upper()
    if is_sens:
        pc_rec = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_r = "Intel Core i7-1355U | 16 GO DDR5 | 512 GO SSD NVMe | Windows 11 Pro"
        justif  = "Poste sensible — securite renforcee"
    elif is_port:
        pc_rec = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_r = "Intel Core i5-1335U | 8 GO DDR4 | 256 GO SSD | Windows 11"
        justif  = "Utilisateur nomade — portable recommande"
    else:
        pc_rec = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_r = "Intel Core i5-12500 | 8 GO DDR4 | 256 GO SSD | Windows 11"
        justif  = "Poste fixe standard"

    story.append(_section("PC Recommande par la DSI"))
    story.append(Spacer(1, 5))
    parts = specs_r.split("|")
    story.append(_spec([
        ("Modele recommande",      pc_rec),
        ("Processeur",             parts[0].strip() if len(parts)>0 else "—"),
        ("Memoire RAM",            parts[1].strip() if len(parts)>1 else "—"),
        ("Stockage",               parts[2].strip() if len(parts)>2 else "—"),
        ("Systeme exploitation",   parts[3].strip() if len(parts)>3 else "Windows 11"),
        ("Justification",          justif),
        ("Urgence",                prio),
    ]))
    story.append(Spacer(1, 14))

    if ancien_pc is not None:
        story.append(_section("Historique — Ancien PC"))
        story.append(Spacer(1, 5))
        story.append(_spec([
            ("Ancien modele",   _v(ancien_pc.get("descriptions"))),
            ("Ancien N Inv.",   _v(ancien_pc.get("n_inventaire"))),
            ("Annee",           _v(ancien_pc.get("annee_acquisition"))),
            ("RAM",             _v(ancien_pc.get("ram"))),
            ("Processeur",      _v(ancien_pc.get("processeur"))),
            ("Disque",          _v(ancien_pc.get("hdd"))),
            ("Observation",     _v(ancien_pc.get("observation"))),
        ]))
        story.append(Spacer(1, 14))

    if imp_data is not None and len(imp_data) > 0:
        story.append(_section(f"Imprimantes — Direction {dir_}"))
        story.append(Spacer(1, 5))
        story.append(_tableau(imp_data, {
            "modele":"Modele","type_imp":"Type","site":"Site",
            "n_inventaire":"N Inv.","annee_acquisition":"Annee"
        }, max_rows=10))

    _pied(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT PC
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_pc_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s   = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else ""
    story.append(_entete(f"Rapport Inventaire PC{titre_dir}",
                         f"Parc informatique — Site Tetouan {YEAR_STR}"))
    story.append(_filet())

    total   = len(df)
    bureau  = df[df["type_pc"].str.upper()=="BUREAU"].shape[0]    if "type_pc"      in df.columns else 0
    port    = df[df["type_pc"].str.upper()=="PORTABLE"].shape[0]  if "type_pc"      in df.columns else 0
    urgent  = df[df["priorite"]=="🔴 Urgent"].shape[0]            if "priorite"     in df.columns else 0
    ok_n    = df[df["priorite"]=="🟢 Opérationnel"].shape[0]      if "priorite"     in df.columns else 0
    n_dir   = df["direction"].nunique()                            if "direction"    in df.columns else 0
    avg_age = round(df["anciennete"].mean(),1) if "anciennete" in df.columns and df["anciennete"].notna().any() else 0
    old_n   = df[df["categorie_age"]=="Plus de 10 ans"].shape[0]  if "categorie_age" in df.columns else 0

    story.append(_intro(
        f"Inventaire de <b>{total} PC actifs</b> dans <b>{n_dir} directions</b>. "
        f"Age moyen du parc : <b>{avg_age} ans</b>. "
        f"PC superieurs a 10 ans : <b>{old_n}</b>. "
        f"Urgents a remplacer : <b>{urgent}</b>."
    ))
    story.append(Spacer(1, 12))

    story.append(_kpis([
        (total,   "Total PC"),
        (urgent,  "Urgents"),
        (ok_n,    "Operationnels"),
        (n_dir,   "Directions"),
    ], ncols=4))
    story.append(Spacer(1, 8))
    story.append(_kpis([
        (bureau,  "Bureau"),
        (port,    "Portable"),
        (avg_age, "Age moyen (ans)"),
        (old_n,   "PC sup. 10 ans"),
    ], ncols=4))
    story.append(Spacer(1, 16))

    if "direction" in df.columns and not direction:
        story.append(_section("Repartition par Direction"))
        story.append(Spacer(1, 5))
        story.append(_barres(df.groupby("direction").size().to_dict()))
        story.append(Spacer(1, 14))

    if "categorie_age" in df.columns:
        story.append(_section("Anciennete du Parc"))
        story.append(Spacer(1, 5))
        story.append(_barres(df["categorie_age"].value_counts().to_dict()))
        story.append(Spacer(1, 14))

    if "ram_go" in df.columns:
        story.append(_section("Distribution de la RAM"))
        story.append(Spacer(1, 5))
        story.append(_barres({f"{int(k)} GO": v for k,v in df["ram_go"].dropna().value_counts().sort_index().items()}))
        story.append(Spacer(1, 14))

    if "priorite" in df.columns and urgent > 0:
        story.append(PageBreak())
        story.append(_section("PC Urgents a Remplacer"))
        story.append(Spacer(1, 5))
        story.append(_tableau(df[df["priorite"]=="🔴 Urgent"], {
            "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
            "type_pc":"Type","annee_acquisition":"Annee","ram":"RAM","raisons":"Raisons",
        }))
        story.append(Spacer(1, 14))

    story.append(PageBreak())
    story.append(_section("Inventaire Complet"))
    story.append(Spacer(1, 5))
    story.append(_tableau(df, {
        "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
        "type_pc":"Type","n_inventaire":"N Inv.","annee_acquisition":"Annee",
        "ram":"RAM","priorite":"Statut",
    }))

    _pied(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT IMPRIMANTES
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_imp_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s   = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else ""
    story.append(_entete(f"Rapport Imprimantes{titre_dir}",
                         f"Inventaire imprimantes — Site Tetouan {YEAR_STR}"))
    story.append(_filet())

    total  = len(df)
    n_dir  = df["direction"].nunique() if "direction" in df.columns else 0
    n_site = df["site"].nunique()      if "site"      in df.columns else 0
    reseau = df[df["type_imp"].str.upper().str.contains("RÉSEAU|RESEAU", na=False)].shape[0] if "type_imp" in df.columns else 0

    story.append(_intro(
        f"Inventaire de <b>{total} imprimantes</b> dans "
        f"<b>{n_dir} directions</b> et <b>{n_site} sites</b>. "
        f"Reseau : <b>{reseau}</b>.  Monoposte : <b>{total-reseau}</b>."
    ))
    story.append(Spacer(1, 12))

    story.append(_kpis([
        (total,        "Total imprimantes"),
        (n_dir,        "Directions"),
        (reseau,       "Reseau"),
        (total-reseau, "Monoposte"),
    ], ncols=4))
    story.append(Spacer(1, 16))

    if "direction" in df.columns and not direction:
        story.append(_section("Repartition par Direction"))
        story.append(Spacer(1, 5))
        story.append(_barres(df.groupby("direction").size().to_dict()))
        story.append(Spacer(1, 14))

    if "modele" in df.columns:
        import re as _re2
        story.append(_section("Top Modeles"))
        story.append(Spacer(1, 5))
        df_m = df[df["modele"].notna()].copy()
        df_m["mc_raw"] = df_m["modele"].astype(str).str.strip()
        df_m = df_m[~df_m["mc_raw"].str.match(r'^\d+\.?\d*$', na=False)]
        df_m = df_m[df_m["mc_raw"].str.len() > 3]
        df_m = df_m[~df_m["mc_raw"].str.upper().str.strip().isin(["SANS","N/A","—","NAN","NONE",""])]
        # Clé sans espaces pour regrouper HP LASER JET P400 = HP LASERJET P400 = HP LASER JET P 400
        df_m["cle"]   = df_m["mc_raw"].str.upper().str.strip().apply(lambda s: _re2.sub(r'[\s\-\./]+','',s))
        df_m["label"] = df_m["mc_raw"].str.upper().str.strip().apply(lambda s: _re2.sub(r'\s+',' ',_re2.sub(r'[\s\-\.]+', ' ', s)).strip())
        grp_m = df_m.groupby("cle").agg(count=("cle","count"), label=("label", lambda x: x.value_counts().index[0])).reset_index()
        grp_m = grp_m.sort_values("count", ascending=False).head(10)
        story.append(_barres({r["label"]: r["count"] for _, r in grp_m.iterrows()}))
        story.append(Spacer(1, 14))

    story.append(PageBreak())
    story.append(_section("Inventaire Complet"))
    story.append(Spacer(1, 5))
    story.append(_tableau(df, {
        "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
        "modele":"Modele","type_imp":"Type","n_inventaire":"N Inv.",
        "annee_acquisition":"Annee",
    }))

    _pied(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT GENERAL
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_general_pdf(df_inv: pd.DataFrame, df_imp: pd.DataFrame,
                                direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s   = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else ""
    story.append(_entete(f"Rapport General{titre_dir}",
                         f"Synthese informatique — Site Tetouan {YEAR_STR}"))
    story.append(_filet())

    total_pc  = len(df_inv) if df_inv is not None else 0
    total_imp = len(df_imp) if df_imp is not None else 0
    urgent    = df_inv[df_inv["priorite"]=="🔴 Urgent"].shape[0] if (df_inv is not None and "priorite" in df_inv.columns) else 0
    n_dir     = df_inv["direction"].nunique()                     if (df_inv is not None and "direction" in df_inv.columns) else 0
    avg_age   = round(df_inv["anciennete"].mean(),1)              if (df_inv is not None and "anciennete" in df_inv.columns and df_inv["anciennete"].notna().any()) else 0
    old_pct   = round(df_inv[df_inv["categorie_age"]=="Plus de 10 ans"].shape[0]/total_pc*100) if (df_inv is not None and "categorie_age" in df_inv.columns and total_pc>0) else 0

    story.append(_intro(
        f"Synthese : <b>{total_pc} PC</b> et <b>{total_imp} imprimantes</b> "
        f"sur <b>{n_dir} directions</b>. "
        f"Age moyen : <b>{avg_age} ans</b>. "
        f"PC superieurs a 10 ans : <b>{old_pct}%</b>. "
        f"Urgents : <b>{urgent}</b>."
    ))
    story.append(Spacer(1, 12))

    story.append(_kpis([
        (total_pc,  "Total PC"),
        (total_imp, "Imprimantes"),
        (urgent,    "Urgents"),
        (n_dir,     "Directions"),
    ], ncols=4))
    story.append(Spacer(1, 16))

    if df_inv is not None and "direction" in df_inv.columns:
        story.append(_section("Synthese par Direction"))
        story.append(Spacer(1, 5))
        agg = df_inv.groupby("direction").agg(Total_PC=("utilisateur","count")).reset_index()
        if "priorite" in df_inv.columns:
            urg = df_inv[df_inv["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)
        if "categorie_age" in df_inv.columns:
            old_ = df_inv[df_inv["categorie_age"]=="Plus de 10 ans"].groupby("direction").size().reset_index(name="Sup 10 ans")
            agg  = agg.merge(old_, on="direction", how="left").fillna(0)
            agg["Sup 10 ans"] = agg["Sup 10 ans"].astype(int)
        if df_imp is not None and "direction" in df_imp.columns:
            ia  = df_imp.groupby("direction").size().reset_index(name="Imprimantes")
            agg = agg.merge(ia, on="direction", how="left").fillna(0)
            agg["Imprimantes"] = agg["Imprimantes"].astype(int)
        agg = agg.sort_values("Total_PC", ascending=False)
        story.append(_tableau(agg, {c: c.replace("_"," ") if c!="direction" else "Direction" for c in agg.columns}))
        story.append(Spacer(1, 14))

    if df_inv is not None and "categorie_age" in df_inv.columns:
        story.append(_section("Anciennete du Parc PC"))
        story.append(Spacer(1, 5))
        story.append(_barres(df_inv["categorie_age"].value_counts().to_dict()))
        story.append(Spacer(1, 14))

    if df_imp is not None and "modele" in df_imp.columns:
        import re as _re3
        story.append(_section("Top Modeles Imprimantes"))
        story.append(Spacer(1, 5))
        df_m = df_imp[df_imp["modele"].notna()].copy()
        df_m["mc_raw"] = df_m["modele"].astype(str).str.strip()
        df_m = df_m[~df_m["mc_raw"].str.match(r'^\d+\.?\d*$', na=False)]
        df_m = df_m[~df_m["mc_raw"].str.upper().str.strip().isin(["SANS","N/A","—","NAN","NONE",""])]
        df_m["cle"]   = df_m["mc_raw"].str.upper().str.strip().apply(lambda s: _re3.sub(r'[\s\-\./]+','',s))
        df_m["label"] = df_m["mc_raw"].str.upper().str.strip().apply(lambda s: _re3.sub(r'\s+',' ',_re3.sub(r'[\s\-\.]+', ' ', s)).strip())
        grp_m2 = df_m.groupby("cle").agg(count=("cle","count"), label=("label", lambda x: x.value_counts().index[0])).reset_index()
        grp_m2 = grp_m2.sort_values("count", ascending=False).head(8)
        story.append(_barres({r["label"]: r["count"] for _, r in grp_m2.iterrows()}))
        story.append(Spacer(1, 14))

    story.append(_section("Conclusions et Recommandations"))
    story.append(Spacer(1, 5))
    concs = [
        f"Plan de renouvellement : {urgent} PC urgents identifies, {old_pct}% du parc a plus de 10 ans.",
        f"Imprimantes : {total_imp} actives — migration vers modeles reseau partage recommandee.",
        f"Taux de restitution satisfaisant sur la majorite des directions.",
        f"Priorite au deploiement IMF sur les directions a fort parc.",
    ]
    rows_c = []
    for i, c in enumerate(concs):
        rows_c.append([
            Paragraph(str(i+1), ParagraphStyle("cn", fontName="Helvetica-Bold",
                                                fontSize=8, textColor=TEXTE_S,
                                                alignment=TA_CENTER)),
            Paragraph(c, s["Corps"]),
        ])
    t_c = Table(rows_c, colWidths=[1*cm, W-1*cm])
    t_c.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BLANC, FOND_ALT]),
        ("GRID",(0,0),(-1,-1),.4,BORDURE),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(t_c)

    _pied(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()
