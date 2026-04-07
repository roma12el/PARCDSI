"""
pdf_export.py — DSI Amendis Tétouan v6
Design SIMPLE et CLAIR — Charte couleurs Amendis (#C8102E rouge, #1B2A3B navy)
Logo Amendis dans l'en-tête, émojis conservés, mise en page épurée
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
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.platypus.flowables import Flowable

# ─── Palette Amendis ──────────────────────────────────────────────────────────
ROUGE  = colors.HexColor("#C8102E")
ROUGE2 = colors.HexColor("#E8415B")
NAVY   = colors.HexColor("#1B2A3B")
NAVY2  = colors.HexColor("#2D3E50")
BLEU   = colors.HexColor("#2980B9")
VERT   = colors.HexColor("#27AE60")
ORANGE = colors.HexColor("#E67E22")
GRIS   = colors.HexColor("#64748B")
GRIS_C = colors.HexColor("#F2F2F2")
BORDER = colors.HexColor("#E2E8F0")
BLANC  = colors.white
TEXTE  = colors.HexColor("#2D2D2D")

DATE_STR = datetime.now().strftime("%d/%m/%Y")
YEAR_STR = "2026"
W = 17 * cm


class BandeAmendis(Flowable):
    def __init__(self, width=W, h=5):
        super().__init__()
        self.width = width
        self._height = h
    def draw(self):
        self.canv.setFillColor(ROUGE)
        self.canv.rect(0, 0, self.width, self._height, fill=1, stroke=0)
    def wrap(self, *args): return self.width, self._height


def _v(v, default="—"):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return s if s not in ("nan","NaT","None","") else default


def _prio_color(p):
    p = str(p).lower()
    if "urgent"       in p: return ROUGE
    if "prioritaire"  in p: return ORANGE
    if "surveiller"   in p: return colors.HexColor("#D97706")
    if "operationnel" in p or "opérationnel" in p: return VERT
    return GRIS


def _styles():
    s = getSampleStyleSheet()
    def add(name, **kw):
        if name not in s.byName: s.add(ParagraphStyle(name, **kw))

    add("TitreNom",  fontName="Helvetica-Bold",  fontSize=18, textColor=BLANC, alignment=TA_LEFT,  leading=22)
    add("TitreSub",  fontName="Helvetica",        fontSize=9,  textColor=colors.HexColor("#BBBBBB"), alignment=TA_LEFT, leading=13)
    add("Logo",      fontName="Helvetica-Bold",   fontSize=14, textColor=ROUGE, alignment=TA_RIGHT, leading=18)
    add("LogoSub",   fontName="Helvetica",        fontSize=7,  textColor=GRIS,  alignment=TA_RIGHT, leading=10)
    add("Sec",       fontName="Helvetica-Bold",   fontSize=10, textColor=BLANC, spaceBefore=4, spaceAfter=2)
    add("Corps",     fontName="Helvetica",        fontSize=8,  textColor=TEXTE, spaceAfter=3, leading=12)
    add("Italic",    fontName="Helvetica-Oblique",fontSize=8,  textColor=GRIS,  spaceAfter=3)
    add("TH",        fontName="Helvetica-Bold",   fontSize=7,  textColor=BLANC, alignment=TA_CENTER)
    add("TD",        fontName="Helvetica",        fontSize=7,  textColor=TEXTE, alignment=TA_LEFT)
    add("SK",        fontName="Helvetica-Bold",   fontSize=8,  textColor=NAVY)
    add("SV",        fontName="Helvetica",        fontSize=8,  textColor=TEXTE)
    add("PiedG",     fontName="Helvetica",        fontSize=7,  textColor=GRIS,  alignment=TA_LEFT)
    add("PiedD",     fontName="Helvetica-Bold",   fontSize=7,  textColor=ROUGE, alignment=TA_RIGHT)
    return s


def _doc(buf):
    return SimpleDocTemplate(buf, pagesize=A4,
                              leftMargin=2*cm, rightMargin=2*cm,
                              topMargin=1.5*cm, bottomMargin=2*cm)


def _entete(titre, sous_titre=""):
    s = _styles()
    gauche = [[Paragraph(titre, s["TitreNom"])]]
    if sous_titre:
        gauche.append([Paragraph(sous_titre, s["TitreSub"])])
    droite = [
        [Paragraph("Amendis", s["Logo"])],
        [Paragraph("Direction des Systèmes d'Information", s["LogoSub"])],
        [Paragraph(f"Site Tetouan  |  {DATE_STR}", s["LogoSub"])],
    ]
    t_g = Table(gauche, colWidths=[W*0.6])
    t_d = Table(droite, colWidths=[W*0.4])
    t_g.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    t_d.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
    outer = Table([[t_g, t_d]], colWidths=[W*0.6, W*0.4])
    outer.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),NAVY),
        ("TOPPADDING",(0,0),(-1,-1),16),("BOTTOMPADDING",(0,0),(-1,-1),16),
        ("LEFTPADDING",(0,0),(0,0),18),("RIGHTPADDING",(0,0),(0,0),10),
        ("LEFTPADDING",(1,0),(1,0),8),("RIGHTPADDING",(1,0),(1,0),16),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return outer


def _section(texte, couleur=None):
    if couleur is None: couleur = NAVY
    s = _styles()
    t = Table([[Paragraph(f"<b>{texte}</b>", s["Sec"])]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),couleur),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),12),
    ]))
    return t


def _intro(texte):
    s = _styles()
    t = Table([[Paragraph(texte, s["Italic"])]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#EFF6FF")),
        ("LINEBEFORE",(0,0),(0,-1),4,BLEU),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),14),("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    return t


def _kpis(items, ncols=4):
    s = _styles()
    cw = W / ncols
    row = []
    for (val, label, col) in items:
        card = Table([
            [Paragraph(str(val), ParagraphStyle("kv3", fontName="Helvetica-Bold",
                                                 fontSize=22, textColor=col,
                                                 alignment=TA_CENTER, leading=26))],
            [Paragraph(label, ParagraphStyle("kl3", fontName="Helvetica-Bold",
                                              fontSize=6, textColor=GRIS,
                                              alignment=TA_CENTER))],
        ], colWidths=[cw-8])
        card.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BLANC),
            ("BOX",(0,0),(-1,-1),.5,BORDER),
            ("LINEABOVE",(0,0),(-1,0),3,col),
            ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        row.append(card)
    while len(row) < ncols:
        row.append(Spacer(cw-8, 1))
    outer = Table([row], colWidths=[cw]*ncols)
    outer.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
    ]))
    return outer


def _spec(champs):
    s = _styles()
    data = [[Paragraph(str(k), s["SK"]), Paragraph(str(v), s["SV"])] for k,v in champs]
    if not data: return Spacer(1,1)
    t = Table(data, colWidths=[W*0.36, W*0.64])
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BLANC, GRIS_C]),
        ("GRID",(0,0),(-1,-1),.3,BORDER),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#F8FAFC")),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return t


def _tableau(df, col_map, max_rows=None, marquer_prio=True):
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
            val = _v(r.get(c))
            if c == "priorite" and marquer_prio:
                pc = _prio_color(val)
                cell = Paragraph(f"<b>{val}</b>",
                                  ParagraphStyle("pv2", fontName="Helvetica-Bold",
                                                 fontSize=7, textColor=pc))
            else:
                cell = Paragraph(val, s["TD"])
            row.append(cell)
        lignes.append(row)
    n  = len(cols)
    cw = W / n
    t  = Table(lignes, colWidths=[cw]*n, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),NAVY),
        ("GRID",(0,0),(-1,-1),.3,BORDER),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[BLANC, GRIS_C]),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,1),(0,-1),"LEFT"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
    ]))
    return t


def _barres(data_dict, couleur=None, max_items=12):
    if couleur is None: couleur = ROUGE
    s = _styles()
    if not data_dict: return Paragraph("Aucune donnee.", s["Corps"])
    NOM_W   = W * 0.32
    BARRE_W = W * 0.52
    VAL_W   = W * 0.16
    max_v   = max(data_dict.values(), default=1)
    lignes  = []
    for nom, val in sorted(data_dict.items(), key=lambda x:-x[1])[:max_items]:
        pct   = min(1.0, val/max_v) if max_v > 0 else 0
        fill  = max(3, int(BARRE_W * pct))
        rest  = max(2, int(BARRE_W) - fill)
        barre = Table([["",""]], colWidths=[fill, rest], rowHeights=[10])
        barre.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(0,0),couleur),
            ("BACKGROUND",(1,0),(1,0),BORDER),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ]))
        lignes.append([
            Paragraph(str(nom)[:32], ParagraphStyle("nb_", fontName="Helvetica-Bold", fontSize=8, textColor=TEXTE)),
            barre,
            Paragraph(f"<b>{val}</b>", ParagraphStyle("vb_", fontName="Helvetica-Bold", fontSize=9, textColor=couleur, alignment=TA_RIGHT)),
        ])
    t = Table(lignes, colWidths=[NOM_W, BARRE_W, VAL_W])
    t.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BLANC, GRIS_C]),
        ("BOX",(0,0),(-1,-1),.3,BORDER),
        ("LINEBELOW",(0,0),(-1,-1),.3,BORDER),
    ]))
    return t


def _pied(story, s=None):
    if s is None: s = _styles()
    story.append(Spacer(1,12))
    barre = Table([[""]], colWidths=[W], rowHeights=[2])
    barre.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),ROUGE),
                                ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.append(barre)
    pied = Table([[
        Paragraph("DSI — Direction des Systemes d'Information  |  Amendis Tetouan  |  Confidentiel", s["PiedG"]),
        Paragraph(f"Genere le {DATE_STR}", s["PiedD"]),
    ]], colWidths=[W*0.70, W*0.30])
    pied.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),NAVY),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
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
    prio     = _v(row.get("priorite"), "Operationnel")
    rec      = _v(row.get("recommandation"), "Aucune action")
    raisons  = _v(row.get("raisons"), "Equipement conforme")
    waterp   = _v(row.get("waterp"), "NON")
    sensible = _v(row.get("poste_sensible"), "NON")
    obs      = _v(row.get("observation"))
    pc_col   = _prio_color(prio)

    story.append(_entete(f"Fiche Utilisateur — {nom}",
                         f"Direction {dir_}  |  Site {site}  |  {type_pc}"))
    story.append(BandeAmendis(W))
    story.append(Spacer(1, 10))

    story.append(_intro(
        f"<b>{nom}</b> — Direction {dir_}, Site {site}. "
        f"Statut DSI : <b>{prio}</b>.  Recommandation : <b>{rec}</b>."
    ))
    story.append(Spacer(1, 10))

    story.append(_kpis([
        (dir_,    "Direction",  NAVY),
        (site,    "Site",       BLEU),
        (type_pc, "Type PC",    NAVY2),
        (prio.replace("🔴","").replace("🟠","").replace("🟡","").replace("🟢","").strip(),
         "Statut DSI", pc_col),
    ], ncols=4))
    story.append(Spacer(1, 12))

    story.append(_section("Specifications du PC Actuel"))
    story.append(Spacer(1, 4))
    story.append(_spec([
        ("N Inventaire",        n_inv),
        ("Modele / Description", desc),
        ("N Serie",             n_serie),
        ("Annee d'acquisition", annee),
        ("Anciennete",          f"{anc} ans"),
        ("Categorie age",       cat_age),
        ("Memoire RAM",         ram),
        ("Processeur",          proc),
        ("Disque dur / SSD",    hdd),
        ("Direction",           dir_),
        ("Site",                site),
        ("Type de poste",       type_pc),
        ("Poste sensible",      sensible),
        ("WATERP",              waterp),
    ]))
    story.append(Spacer(1, 12))

    story.append(_section("Statut & Recommandation DSI", couleur=ROUGE))
    story.append(Spacer(1, 4))
    story.append(_spec([
        ("Priorite",      prio),
        ("Recommandation",rec),
        ("Raisons",       raisons),
        ("Observation",   obs),
    ]))
    story.append(Spacer(1, 12))

    is_sens = "OUI" in str(sensible).upper()
    is_port = "PORTABLE" in str(type_pc).upper()
    if is_sens:
        pc_rec  = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_r = "Intel Core i7-1355U | 16 GO DDR5 | 512 GO SSD NVMe | Windows 11 Pro"
        justif  = "Poste sensible - securite renforcee"
    elif is_port:
        pc_rec  = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_r = "Intel Core i5-1335U | 8 GO DDR4 | 256 GO SSD | Windows 11"
        justif  = "Utilisateur nomade - portable recommande"
    else:
        pc_rec  = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_r = "Intel Core i5-12500 | 8 GO DDR4 | 256 GO SSD | Windows 11"
        justif  = "Poste fixe standard"

    story.append(_section("PC Recommande par la DSI", couleur=BLEU))
    story.append(Spacer(1, 4))
    parts = specs_r.split("|")
    story.append(_spec([
        ("Modele recommande",      pc_rec),
        ("Processeur",             parts[0].strip() if len(parts)>0 else "—"),
        ("Memoire RAM",            parts[1].strip() if len(parts)>1 else "—"),
        ("Stockage",               parts[2].strip() if len(parts)>2 else "—"),
        ("Systeme",                parts[3].strip() if len(parts)>3 else "Windows 11"),
        ("Justification",          justif),
        ("Urgence remplacement",   prio),
    ]))
    story.append(Spacer(1, 12))

    if ancien_pc is not None:
        story.append(_section("Historique — Ancien PC", couleur=NAVY2))
        story.append(Spacer(1, 4))
        story.append(_spec([
            ("Ancien modele",  _v(ancien_pc.get("descriptions"))),
            ("Ancien N Inv.",  _v(ancien_pc.get("n_inventaire"))),
            ("Annee",          _v(ancien_pc.get("annee_acquisition"))),
            ("RAM",            _v(ancien_pc.get("ram"))),
            ("Processeur",     _v(ancien_pc.get("processeur"))),
            ("Disque",         _v(ancien_pc.get("hdd"))),
            ("Observation",    _v(ancien_pc.get("observation"))),
        ]))
        story.append(Spacer(1, 12))

    if imp_data is not None and len(imp_data) > 0:
        story.append(_section(f"Imprimantes — Direction {dir_}", couleur=colors.HexColor("#7D3C98")))
        story.append(Spacer(1, 4))
        story.append(_tableau(imp_data, {
            "modele":"Modele","type_imp":"Type","site":"Site",
            "n_inventaire":"N Inv.","annee_acquisition":"Annee"
        }, max_rows=10, marquer_prio=False))

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

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_entete(f"Rapport Inventaire PC{titre_dir}",
                         f"Parc informatique Site Tetouan {YEAR_STR}"))
    story.append(BandeAmendis(W))
    story.append(Spacer(1, 10))

    total  = len(df)
    bureau = df[df["type_pc"].str.upper()=="BUREAU"].shape[0]    if "type_pc"  in df.columns else 0
    port   = df[df["type_pc"].str.upper()=="PORTABLE"].shape[0]  if "type_pc"  in df.columns else 0
    urgent = df[df["priorite"]=="🔴 Urgent"].shape[0]            if "priorite" in df.columns else 0
    ok_n   = df[df["priorite"]=="🟢 Opérationnel"].shape[0]      if "priorite" in df.columns else 0
    n_dir  = df["direction"].nunique()                            if "direction" in df.columns else 0
    avg_age= round(df["anciennete"].mean(),1) if "anciennete" in df.columns and df["anciennete"].notna().any() else 0
    old_n  = df[df["categorie_age"]=="Plus de 10 ans"].shape[0]  if "categorie_age" in df.columns else 0

    story.append(_intro(
        f"Inventaire de <b>{total} PC actifs</b> dans <b>{n_dir} directions</b>. "
        f"Age moyen : <b>{avg_age} ans</b>. "
        f"PC sup. 10 ans : <b>{old_n}</b>.  Urgents : <b>{urgent}</b>."
    ))
    story.append(Spacer(1, 10))

    story.append(_kpis([
        (total,  "Total PC",       ROUGE),
        (urgent, "Urgents",        ROUGE),
        (ok_n,   "Operationnels",  VERT),
        (n_dir,  "Directions",     NAVY),
    ], ncols=4))
    story.append(Spacer(1, 8))
    story.append(_kpis([
        (bureau,  "Bureau",        BLEU),
        (port,    "Portable",      NAVY2),
        (avg_age, "Age Moyen ans", ORANGE),
        (old_n,   "PC sup 10 ans", ROUGE),
    ], ncols=4))
    story.append(Spacer(1, 14))

    if "direction" in df.columns and not direction:
        story.append(_section("PC par Direction"))
        story.append(Spacer(1, 4))
        story.append(_barres(df.groupby("direction").size().to_dict(), couleur=ROUGE))
        story.append(Spacer(1, 12))

    if "categorie_age" in df.columns:
        story.append(_section("Anciennete du Parc", couleur=ORANGE))
        story.append(Spacer(1, 4))
        story.append(_barres(df["categorie_age"].value_counts().to_dict(), couleur=ORANGE))
        story.append(Spacer(1, 12))

    if "ram_go" in df.columns:
        story.append(_section("Distribution RAM", couleur=BLEU))
        story.append(Spacer(1, 4))
        story.append(_barres({f"{int(k)} GO": v for k,v in df["ram_go"].dropna().value_counts().sort_index().items()}, couleur=BLEU))
        story.append(Spacer(1, 12))

    if "priorite" in df.columns and urgent > 0:
        story.append(PageBreak())
        story.append(_section("PC Urgents a Remplacer", couleur=ROUGE))
        story.append(Spacer(1, 4))
        story.append(_tableau(df[df["priorite"]=="🔴 Urgent"], {
            "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
            "type_pc":"Type","annee_acquisition":"Annee","ram":"RAM",
            "processeur":"Processeur","raisons":"Raisons",
        }, marquer_prio=False))
        story.append(Spacer(1, 12))

    story.append(PageBreak())
    story.append(_section("Inventaire Complet", couleur=NAVY))
    story.append(Spacer(1, 4))
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

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_entete(f"Rapport Imprimantes{titre_dir}",
                         f"Inventaire Imprimantes Site Tetouan {YEAR_STR}"))
    story.append(BandeAmendis(W))
    story.append(Spacer(1, 10))

    total  = len(df)
    n_dir  = df["direction"].nunique() if "direction" in df.columns else 0
    n_site = df["site"].nunique()      if "site"      in df.columns else 0
    reseau = df[df["type_imp"].str.upper().str.contains("RÉSEAU|RESEAU",na=False)].shape[0] if "type_imp" in df.columns else 0

    story.append(_intro(
        f"Inventaire de <b>{total} imprimantes</b> dans "
        f"<b>{n_dir} directions</b> et <b>{n_site} sites</b>. "
        f"Reseau : <b>{reseau}</b>.  Monoposte : <b>{total-reseau}</b>."
    ))
    story.append(Spacer(1, 10))

    story.append(_kpis([
        (total,        "Total Imprimantes", colors.HexColor("#7D3C98")),
        (n_dir,        "Directions",        NAVY),
        (reseau,       "Reseau",            VERT),
        (total-reseau, "Monoposte",         BLEU),
    ], ncols=4))
    story.append(Spacer(1, 14))

    if "direction" in df.columns and not direction:
        story.append(_section("Par Direction", couleur=colors.HexColor("#7D3C98")))
        story.append(Spacer(1, 4))
        story.append(_barres(df.groupby("direction").size().to_dict(), couleur=colors.HexColor("#9B59B6")))
        story.append(Spacer(1, 12))

    if "modele" in df.columns:
        story.append(_section("Top Modeles", couleur=BLEU))
        story.append(Spacer(1, 4))
        df_m = df[df["modele"].notna()].copy()
        df_m["mc"] = df_m["modele"].astype(str).str.strip()
        df_m = df_m[~df_m["mc"].str.match(r'^\d+\.?\d*$', na=False)]
        df_m = df_m[df_m["mc"].str.len() > 3]
        story.append(_barres(df_m["mc"].value_counts().head(10).to_dict(), couleur=BLEU))
        story.append(Spacer(1, 12))

    story.append(PageBreak())
    story.append(_section("Inventaire Complet", couleur=NAVY))
    story.append(Spacer(1, 4))
    story.append(_tableau(df, {
        "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
        "modele":"Modele","type_imp":"Type","n_inventaire":"N Inv.",
        "annee_acquisition":"Annee",
    }, marquer_prio=False))

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

    titre_dir = f" — Direction {direction}" if direction else " — Parc Complet"
    story.append(_entete(f"Rapport General{titre_dir}",
                         f"Synthese Informatique Complete — Site Tetouan {YEAR_STR}"))
    story.append(BandeAmendis(W))
    story.append(Spacer(1, 10))

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
        f"PC sup. 10 ans : <b>{old_pct}%</b>. "
        f"Urgents : <b>{urgent}</b>."
    ))
    story.append(Spacer(1, 10))

    story.append(_kpis([
        (total_pc,  "Total PC",     ROUGE),
        (total_imp, "Imprimantes",  colors.HexColor("#9B59B6")),
        (urgent,    "Urgents",      ROUGE),
        (n_dir,     "Directions",   NAVY),
    ], ncols=4))
    story.append(Spacer(1, 14))

    if df_inv is not None and "direction" in df_inv.columns:
        story.append(_section("Synthese par Direction", couleur=NAVY))
        story.append(Spacer(1, 4))
        agg = df_inv.groupby("direction").agg(Total_PC=("utilisateur","count")).reset_index()
        if "priorite" in df_inv.columns:
            urg = df_inv[df_inv["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)
        if "categorie_age" in df_inv.columns:
            old_ = df_inv[df_inv["categorie_age"]=="Plus de 10 ans"].groupby("direction").size().reset_index(name="Sup10ans")
            agg  = agg.merge(old_, on="direction", how="left").fillna(0)
            agg["Sup10ans"] = agg["Sup10ans"].astype(int)
        if df_imp is not None and "direction" in df_imp.columns:
            ia  = df_imp.groupby("direction").size().reset_index(name="Imprimantes")
            agg = agg.merge(ia, on="direction", how="left").fillna(0)
            agg["Imprimantes"] = agg["Imprimantes"].astype(int)
        agg = agg.sort_values("Total_PC", ascending=False)
        story.append(_tableau(agg, {c: c.replace("_"," ") if c!="direction" else "Direction" for c in agg.columns}, marquer_prio=False))
        story.append(Spacer(1, 12))

    if df_inv is not None and "categorie_age" in df_inv.columns:
        story.append(_section("Anciennete du Parc PC", couleur=ORANGE))
        story.append(Spacer(1, 4))
        story.append(_barres(df_inv["categorie_age"].value_counts().to_dict(), couleur=ORANGE))
        story.append(Spacer(1, 12))

    if df_imp is not None and "modele" in df_imp.columns:
        story.append(_section("Top Modeles Imprimantes", couleur=colors.HexColor("#7D3C98")))
        story.append(Spacer(1, 4))
        df_m = df_imp[df_imp["modele"].notna()].copy()
        df_m["mc"] = df_m["modele"].astype(str).str.strip()
        df_m = df_m[~df_m["mc"].str.match(r'^\d+\.?\d*$', na=False)]
        story.append(_barres(df_m["mc"].value_counts().head(8).to_dict(), couleur=colors.HexColor("#9B59B6")))
        story.append(Spacer(1, 12))

    story.append(_section("Conclusions & Recommandations", couleur=VERT))
    story.append(Spacer(1, 4))
    for c in [
        f"Plan de renouvellement urgent : {urgent} PC urgents, {old_pct}% du parc a plus de 10 ans.",
        f"Imprimantes : {total_imp} actives — migrer vers modeles reseau partages recommande.",
        f"Taux de restitution eleve sur la majorite des directions.",
        f"Priorite deploiement IMF sur les directions a fort parc.",
    ]:
        t_c = Table([[Paragraph(c, s["Corps"])]], colWidths=[W])
        t_c.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),GRIS_C),
            ("LINEBEFORE",(0,0),(0,-1),3,ROUGE),
            ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
            ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),
        ]))
        story.append(t_c)
        story.append(Spacer(1,4))

    _pied(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()
