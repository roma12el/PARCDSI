"""
pdf_export.py — DSI Amendis Tétouan v5
Design Navy/Teal/Accent — style rapport professionnel HTML
"""
import io
from datetime import datetime
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import Flowable

# ─── Palette couleurs (same as HTML template) ─────────────────────────────────
NAVY   = colors.HexColor("#0A1628")
NAVY2  = colors.HexColor("#1a3a5c")
TEAL   = colors.HexColor("#0D6B8C")
ACCENT = colors.HexColor("#00C2A8")
CORAL  = colors.HexColor("#E8533A")
GOLD   = colors.HexColor("#F0A500")
GREEN  = colors.HexColor("#22C55E")
ORANGE = colors.HexColor("#F59E0B")
RED    = colors.HexColor("#EF4444")
LIGHT  = colors.HexColor("#F0F4F8")
BORDER = colors.HexColor("#E2E8F0")
TEXT   = colors.HexColor("#1E293B")
MUTED  = colors.HexColor("#64748B")
WHITE  = colors.white
LIGHT_TEAL = colors.HexColor("#E0F7FA")
LIGHT_CORAL= colors.HexColor("#FEE8E4")
LIGHT_GREEN= colors.HexColor("#DCFCE7")
LIGHT_NAVY = colors.HexColor("#E8EDF5")

DATE_STR = datetime.now().strftime("%d/%m/%Y")
YEAR_STR = "2026"
PAGE_W   = 17 * cm

# ─── Accent bar flowable ───────────────────────────────────────────────────────
class AccentBar(Flowable):
    def __init__(self, width=PAGE_W, height=4):
        super().__init__()
        self.width = width
        self._height = height
    def draw(self):
        from reportlab.lib.colors import HexColor
        self.canv.setFillColor(ACCENT)
        self.canv.rect(0, 0, self.width/3, self._height, fill=1, stroke=0)
        self.canv.setFillColor(TEAL)
        self.canv.rect(self.width/3, 0, self.width/3, self._height, fill=1, stroke=0)
        self.canv.setFillColor(GOLD)
        self.canv.rect(2*self.width/3, 0, self.width/3, self._height, fill=1, stroke=0)
    def wrap(self, *args): return self.width, self._height


def _v(v, default="—"):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return s if s not in ("nan","NaT","None","") else default

def _prio_color(p):
    p = str(p).lower()
    if "urgent"      in p: return RED
    if "prioritaire" in p: return ORANGE
    if "surveiller"  in p: return GOLD
    if "opérationnel" in p or "operationnel" in p: return GREEN
    return MUTED

def _styles():
    s = getSampleStyleSheet()
    def add(name, **kw):
        if name not in s.byName: s.add(ParagraphStyle(name, **kw))

    # Header styles
    add("H_Title", fontName="Helvetica-Bold",  fontSize=20, textColor=WHITE,  alignment=TA_LEFT,  leading=26)
    add("H_Title2",fontName="Helvetica-Bold",  fontSize=14, textColor=ACCENT, alignment=TA_LEFT,  leading=18)
    add("H_Sub",   fontName="Helvetica",       fontSize=9,  textColor=colors.HexColor("#AAAAAA"), alignment=TA_LEFT)
    add("H_Date",  fontName="Helvetica-Bold",  fontSize=8,  textColor=ACCENT, alignment=TA_RIGHT)
    add("H_Org",   fontName="Helvetica-Bold",  fontSize=9,  textColor=WHITE,  alignment=TA_RIGHT)
    add("H_Dept",  fontName="Helvetica",       fontSize=8,  textColor=colors.HexColor("#888888"), alignment=TA_RIGHT)

    # Section title
    add("Sec",     fontName="Helvetica-Bold",  fontSize=10, textColor=NAVY,   spaceBefore=10, spaceAfter=5)
    add("SecAccent",fontName="Helvetica-Bold", fontSize=10, textColor=ACCENT, spaceBefore=8,  spaceAfter=4)

    # Body
    add("Body2",   fontName="Helvetica",       fontSize=8,  textColor=TEXT,   spaceAfter=3)
    add("BodyBold",fontName="Helvetica-Bold",  fontSize=8,  textColor=TEXT)
    add("Cap",     fontName="Helvetica-Oblique",fontSize=7, textColor=MUTED,  alignment=TA_CENTER)
    add("Intro",   fontName="Helvetica-Oblique",fontSize=8, textColor=MUTED,  spaceAfter=6, spaceBefore=4)

    # Table
    add("TH",      fontName="Helvetica-Bold",  fontSize=7,  textColor=WHITE,  alignment=TA_CENTER)
    add("TD",      fontName="Helvetica",       fontSize=7,  textColor=TEXT,   alignment=TA_LEFT)
    add("TDC",     fontName="Helvetica",       fontSize=7,  textColor=TEXT,   alignment=TA_CENTER)
    add("TDB",     fontName="Helvetica-Bold",  fontSize=7,  textColor=TEXT,   alignment=TA_LEFT)

    # Footer
    add("FL",      fontName="Helvetica",       fontSize=7,  textColor=colors.HexColor("#AAAAAA"), alignment=TA_LEFT)
    add("FR",      fontName="Helvetica",       fontSize=7,  textColor=ACCENT, alignment=TA_RIGHT)

    # KPI
    add("KV",      fontName="Helvetica-Bold",  fontSize=22, alignment=TA_CENTER, leading=26)
    add("KL",      fontName="Helvetica-Bold",  fontSize=7,  textColor=MUTED,    alignment=TA_CENTER, textTransform="uppercase")
    add("KS",      fontName="Helvetica",       fontSize=7,  textColor=MUTED,    alignment=TA_CENTER)

    # Spec
    add("SpecK",   fontName="Helvetica-Bold",  fontSize=7,  textColor=NAVY)
    add("SpecV",   fontName="Helvetica",       fontSize=8,  textColor=TEXT)

    return s


def _doc(buf):
    return SimpleDocTemplate(buf, pagesize=A4,
                              leftMargin=2*cm, rightMargin=2*cm,
                              topMargin=1.5*cm, bottomMargin=2*cm)


def _header(title, subtitle="", org_line="DSI — Direction des SI", date_line=None):
    """Full-width header with gradient background approximation."""
    s = _styles()
    if date_line is None:
        date_line = datetime.now().strftime("%B %Y")

    left_items = [[Paragraph(title, s["H_Title"])]]
    if subtitle:
        left_items.append([Paragraph(subtitle, s["H_Sub"])])

    right_items = [
        [Paragraph(date_line.upper(), s["H_Date"])],
        [Paragraph(org_line, s["H_Org"])],
        [Paragraph(f"Site Tétouan · Inventaire 2025-{YEAR_STR}", s["H_Dept"])],
    ]

    left_t  = Table(left_items, colWidths=[PAGE_W*0.65])
    right_t = Table(right_items, colWidths=[PAGE_W*0.35])

    left_t.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]))
    right_t.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#0F2040")),
        ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
        ("BOX",(0,0),(-1,-1),.5,colors.HexColor("#AACCDD")),
    ]))

    outer = Table([[left_t, right_t]], colWidths=[PAGE_W*0.65, PAGE_W*0.35])
    outer.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,0),NAVY),
        ("BACKGROUND",(1,0),(1,0),NAVY),
        ("TOPPADDING",(0,0),(-1,-1),14),("BOTTOMPADDING",(0,0),(-1,-1),14),
        ("LEFTPADDING",(0,0),(0,0),18),("RIGHTPADDING",(0,0),(0,0),10),
        ("LEFTPADDING",(1,0),(1,0),0),("RIGHTPADDING",(1,0),(1,0),0),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return outer


def _section_title(text, s):
    """Section title with left accent bar style."""
    bar_cell = Table([[""]], colWidths=[4], rowHeights=[14])
    bar_cell.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),ACCENT),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
    ]))
    label = Table([[Paragraph(f"<b>{text}</b>",
                               ParagraphStyle("st", fontName="Helvetica-Bold", fontSize=10,
                                              textColor=NAVY))]], colWidths=[PAGE_W-12])
    label.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    row = Table([[bar_cell, label]], colWidths=[8, PAGE_W-8])
    row.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return row


def _intro_box(text, s):
    """Intro box with light teal background."""
    t = Table([[Paragraph(text, s["Intro"])]], colWidths=[PAGE_W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#EFF6FF")),
        ("LINEBEFORE",(0,0),(0,-1),4,TEAL),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    return t


def _kpi_grid(items, ncols=4):
    """
    items = list of (icon_emoji, label, value, color, sub_text)
    """
    s = _styles()
    if not items: return Spacer(1,4)
    cw = PAGE_W / ncols
    row = []

    COLOR_BG = {
        TEAL:   LIGHT_TEAL,
        NAVY:   LIGHT_NAVY,
        CORAL:  LIGHT_CORAL,
        GREEN:  LIGHT_GREEN,
        RED:    colors.HexColor("#FEE8E4"),
        ORANGE: colors.HexColor("#FEF3C7"),
        GOLD:   colors.HexColor("#FFFBEB"),
        MUTED:  LIGHT,
    }

    for (icon, label, value, col, sub) in items:
        bg = COLOR_BG.get(col, LIGHT)
        top_cell = Table([
            [Paragraph(icon, ParagraphStyle("ico", fontName="Helvetica", fontSize=18, alignment=TA_CENTER))],
            [Paragraph(f"<b>{value}</b>", ParagraphStyle("kv2", fontName="Helvetica-Bold",
                                                           fontSize=22, textColor=col, alignment=TA_CENTER, leading=26))],
            [Paragraph(label, ParagraphStyle("kl2", fontName="Helvetica-Bold", fontSize=7,
                                              textColor=col, alignment=TA_CENTER,
                                              textTransform="uppercase"))],
        ], colWidths=[cw-8])
        top_cell.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),bg),
            ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        sub_cell = Table([[Paragraph(sub, ParagraphStyle("ks2", fontName="Helvetica",
                                                          fontSize=7, textColor=MUTED,
                                                          alignment=TA_CENTER))]], colWidths=[cw-8])
        sub_cell.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),WHITE),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LINEABOVE",(0,0),(-1,0),.5,BORDER),
        ]))
        card = Table([[top_cell],[sub_cell]], colWidths=[cw-8])
        card.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1),.3,BORDER),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ]))
        row.append(card)

    # Pad row if needed
    while len(row) < ncols:
        row.append(Spacer(cw-8, 1))

    outer = Table([row], colWidths=[cw]*ncols)
    outer.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
    ]))
    return outer


def _data_table(df, col_labels, max_rows=None, highlight_prio=True):
    s = _styles()
    if df is None or len(df)==0:
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
            if c == "priorite" and highlight_prio:
                pc = _prio_color(v)
                cell = Paragraph(f"<b>{v}</b>",
                                  ParagraphStyle("pc", fontName="Helvetica-Bold",
                                                 fontSize=7, textColor=pc))
            elif c in ("Avancement","%") and v.endswith("%"):
                try:
                    pct = float(v.replace("%","").strip())
                    col_p = GREEN if pct>=100 else (ORANGE if pct>=80 else RED)
                except: col_p = MUTED
                cell = Paragraph(f"<b>{v}</b>",
                                  ParagraphStyle("pct", fontName="Helvetica-Bold",
                                                 fontSize=7, textColor=col_p, alignment=TA_CENTER))
            else:
                cell = Paragraph(v, s["TD"])
            r.append(cell)
        rows.append(r)

    n = len(cols)
    # Smart column widths
    cw = PAGE_W / n
    t = Table(rows, colWidths=[cw]*n, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),NAVY),
        ("GRID",(0,0),(-1,-1),.3,BORDER),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE, LIGHT]),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ALIGN",(0,1),(0,-1),"LEFT"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
        # Footer row style if last row is TOTAL
    ]))
    return t


def _bar_chart_table(data_dict, max_val=None, color=TEAL, title=""):
    """Horizontal bar chart - single colored cell spanning pct of bar width."""
    s = _styles()
    if not data_dict:
        return Paragraph("Aucune donnée.", s["Body2"])

    NAME_W = PAGE_W * 0.28
    BAR_W  = PAGE_W * 0.55
    VAL_W  = PAGE_W * 0.12
    max_v  = max_val or max(data_dict.values(), default=1)

    rows_data = []
    style_cmds = [
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
        ("RIGHTPADDING",  (0,0), (-1,-1), 6),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0,0), (-1,-1), [WHITE, LIGHT]),
        ("BOX",           (0,0), (-1,-1), .3, BORDER),
        ("LINEBELOW",     (0,0), (-1,-1), .3, BORDER),
    ]

    for i, (name, val) in enumerate(sorted(data_dict.items(), key=lambda x:-x[1])[:12]):
        pct = min(1.0, val / max_v) if max_v > 0 else 0
        # Use BACKGROUND style command on bar cell to simulate fill
        rows_data.append([
            Paragraph(str(name)[:30], ParagraphStyle(f"bn{i}", fontName="Helvetica-Bold",
                                                      fontSize=8, textColor=TEXT)),
            Paragraph("",             ParagraphStyle(f"bar{i}", fontName="Helvetica", fontSize=1)),
            Paragraph(f"<b>{val}</b>",ParagraphStyle(f"bv{i}", fontName="Helvetica-Bold",
                                                      fontSize=9, textColor=color, alignment=TA_RIGHT)),
        ])
        # Color the bar column based on percentage using BACKGROUND
        bar_fill_w = max(4, BAR_W * pct)
        # We simulate the bar by splitting: fill part colored, rest light
        # Use a sub-table for the bar cell
        fill_t = Table([["",""]],
                        colWidths=[bar_fill_w, max(2, BAR_W - bar_fill_w)],
                        rowHeights=[10])
        fill_t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(0,0), color),
            ("BACKGROUND",(1,0),(1,0), BORDER),
            ("TOPPADDING",   (0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("LEFTPADDING",  (0,0),(-1,-1),0),("RIGHTPADDING",  (0,0),(-1,-1),0),
        ]))
        rows_data[-1][1] = fill_t

    t = Table(rows_data, colWidths=[NAME_W, BAR_W, VAL_W])
    t.setStyle(TableStyle(style_cmds))
    return t


def _spec_table(fields, title=""):
    """HP-style spec table: key | value, alternating rows."""
    s = _styles()
    rows_data = []
    for k, v in fields:
        rows_data.append([
            Paragraph(str(k), s["SpecK"]),
            Paragraph(str(v), s["SpecV"]),
        ])
    if not rows_data:
        return Spacer(1,1)
    t = Table(rows_data, colWidths=[PAGE_W*0.36, PAGE_W*0.64])
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[WHITE, LIGHT]),
        ("GRID",(0,0),(-1,-1),.3,BORDER),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#F8FAFC")),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    return t


def _conclusion_box(icon, title, text, color=GREEN, s=None):
    """Conclusion card — simplified to avoid height=None issues."""
    if s is None: s = _styles()
    half_w = PAGE_W / 2 - 6

    content = Table([
        [Paragraph(f"{icon}  <b>{title}</b>",
                   ParagraphStyle("ct", fontName="Helvetica-Bold", fontSize=9, textColor=NAVY))],
        [Paragraph(text, s["Body2"])],
    ], colWidths=[half_w - 10])
    content.setStyle(TableStyle([
        ("TOPPADDING",   (0,0),(-1,-1),3),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",  (0,0),(-1,-1),8),
        ("RIGHTPADDING", (0,0),(-1,-1),6),
    ]))

    card = Table([[content]], colWidths=[half_w])
    card.setStyle(TableStyle([
        ("BACKGROUND",  (0,0),(-1,-1),WHITE),
        ("BOX",         (0,0),(-1,-1),.4,BORDER),
        ("LINEBEFORE",  (0,0),(0,-1),5,color),
        ("TOPPADDING",  (0,0),(-1,-1),10),
        ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING", (0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),0),
    ]))
    return card


def _footer(story, s=None):
    if s is None: s = _styles()
    story.append(Spacer(1, 10))
    bar = Table([[""]], colWidths=[PAGE_W], rowHeights=[1])
    bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BORDER),
                              ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.append(bar)
    foot = Table([[
        Paragraph(f"<b>DSI</b> — Direction des Systèmes d'Information  |  Site Tétouan  |  Inventaire {YEAR_STR}", s["FL"]),
        Paragraph(f"Généré le {DATE_STR}  |  Confidentiel", s["FR"]),
    ]], colWidths=[PAGE_W*0.65, PAGE_W*0.35])
    foot.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),NAVY),
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),
    ]))
    story.append(foot)


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

    story.append(_header(f"Fiche Technique — {nom}",
                         f"Direction {dir_}  ·  Site {site}  ·  {type_pc}"))
    story.append(AccentBar(PAGE_W))
    story.append(Spacer(1, 10))

    # Intro box
    story.append(_intro_box(
        f"Fiche technique complète pour <b>{nom}</b> — Direction {dir_}, Site {site}. "
        f"Poste sensible: <b>{sensible}</b>  ·  WATERP: <b>{waterp}</b>  ·  Statut: <b>{prio}</b>",
        s))
    story.append(Spacer(1, 8))

    # KPIs identité
    story.append(_section_title("Identité & Statut", s))
    story.append(Spacer(1, 4))
    pc_col = _prio_color(prio)
    story.append(_kpi_grid([
        ("🏢", "Direction",    dir_,   TEAL,  site),
        ("📍", "Site",         site,   NAVY,  "Localisation"),
        ("🖥️", "Type Poste",  type_pc, NAVY2, "Matériel"),
        ("🔴" if "Urgent" in prio else "🟢", "Statut DSI", prio.split()[-1] if len(prio.split())>1 else prio,
         pc_col, rec[:30]),
    ], ncols=4))
    story.append(Spacer(1, 10))

    # Fiche technique PC — style HP datasheet
    story.append(_section_title("🖥️  Fiche Technique PC — Spécifications", s))
    story.append(Spacer(1, 4))
    story.append(_spec_table([
        ("N° Inventaire",      n_inv),
        ("Modèle / Description", desc),
        ("N° Série",           n_serie),
        ("Année d'acquisition", annee),
        ("Ancienneté",         f"{anc} ans"),
        ("Catégorie âge",      cat_age),
        ("Mémoire RAM",        ram),
        ("Processeur",         proc),
        ("Disque dur / SSD",   hdd),
        ("Poste sensible",     sensible),
        ("WATERP",             waterp),
    ]))
    story.append(Spacer(1, 10))

    # Statut DSI
    story.append(_section_title("📋  Statut DSI & Recommandation", s))
    story.append(Spacer(1, 4))
    status_t = Table([
        [Paragraph("<b>PRIORITÉ</b>",    ParagraphStyle("sh",fontName="Helvetica-Bold",fontSize=7,textColor=WHITE,alignment=TA_CENTER)),
         Paragraph("<b>RECOMMANDATION</b>", ParagraphStyle("sh2",fontName="Helvetica-Bold",fontSize=7,textColor=WHITE,alignment=TA_CENTER)),
         Paragraph("<b>JUSTIFICATION</b>",  ParagraphStyle("sh3",fontName="Helvetica-Bold",fontSize=7,textColor=WHITE,alignment=TA_CENTER))],
        [Paragraph(f"<b>{prio}</b>", ParagraphStyle("pv",fontName="Helvetica-Bold",fontSize=10,textColor=pc_col)),
         Paragraph(rec, s["Body2"]),
         Paragraph(raisons, s["Body2"])],
    ], colWidths=[5*cm, 6*cm, 6*cm])
    status_t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),NAVY),
        ("GRID",(0,0),(-1,-1),.4,BORDER),
        ("BACKGROUND",(0,1),(-1,-1),LIGHT_TEAL),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),10),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(status_t)
    story.append(Spacer(1, 10))

    # PC recommandé — style HP spec sheet
    is_sens = "OUI" in str(sensible).upper()
    is_port = "PORTABLE" in str(type_pc).upper()
    if is_sens:
        pc_rec = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_r = "Intel Core i7-1355U · 16 GO DDR5 · 512 GO SSD NVMe · Windows 11 Pro"
        justif  = "Poste sensible — sécurité renforcée"
    elif is_port:
        pc_rec = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_r = "Intel Core i5-1335U · 8 GO DDR4 · 256 GO SSD · Windows 11"
        justif  = "Utilisateur nomade — portable recommandé"
    else:
        pc_rec = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_r = "Intel Core i5-12500 · 8 GO DDR4 · 256 GO SSD · Windows 11"
        justif  = "Poste fixe standard"

    story.append(_section_title("🛒  PC Recommandé par la DSI", s))
    story.append(Spacer(1, 4))
    # Header de la spec table avec style HP bleu
    hdr_rec = Table([[Paragraph(f"🖥️  Ordinateur {'Portable' if is_port else 'Bureau'} Professionnel — {pc_rec}",
                                 ParagraphStyle("rh", fontName="Helvetica-Bold", fontSize=9,
                                                textColor=WHITE))]],
                     colWidths=[PAGE_W])
    hdr_rec.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),TEAL),
                                  ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
                                  ("LEFTPADDING",(0,0),(-1,-1),12)]))
    story.append(hdr_rec)
    story.append(_spec_table([
        ("Modèle recommandé", pc_rec),
        ("Processeur",        specs_r.split("·")[0].strip()),
        ("Mémoire RAM",       specs_r.split("·")[1].strip() if len(specs_r.split("·"))>1 else "—"),
        ("Stockage",          specs_r.split("·")[2].strip() if len(specs_r.split("·"))>2 else "—"),
        ("Système",           specs_r.split("·")[3].strip() if len(specs_r.split("·"))>3 else "Windows 11"),
        ("Justification",     justif),
        ("Urgence",           prio),
    ]))
    story.append(Spacer(1, 8))

    # Ancien PC
    if ancien_pc is not None:
        story.append(_section_title("📦  Ancien PC (Historique Mouvements)", s))
        story.append(Spacer(1, 4))
        story.append(_spec_table([
            ("Ancien modèle",   _v(ancien_pc.get("descriptions"))),
            ("Ancien N° Inv.",  _v(ancien_pc.get("n_inventaire"))),
            ("Année",           _v(ancien_pc.get("annee_acquisition"))),
            ("RAM",             _v(ancien_pc.get("ram"))),
            ("Processeur",      _v(ancien_pc.get("processeur"))),
            ("HDD",             _v(ancien_pc.get("hdd"))),
        ]))
        story.append(Spacer(1, 8))

    # Observation
    if obs and obs != "—":
        obs_t = Table([[Paragraph(f"📝  <b>Observation :</b> {obs}", s["Body2"])]],
                       colWidths=[PAGE_W])
        obs_t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#FFFBEB")),
            ("LINEBEFORE",(0,0),(0,-1),4,GOLD),
            ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING",(0,0),(-1,-1),14),
        ]))
        story.append(obs_t)
        story.append(Spacer(1, 8))

    # Imprimantes direction
    if imp_data is not None and len(imp_data) > 0:
        story.append(_section_title(f"🖨️  Imprimantes — Direction {dir_}", s))
        story.append(Spacer(1, 4))
        story.append(_data_table(imp_data, {
            "modele":"Modèle","type_imp":"Type","site":"Site",
            "n_inventaire":"N° Inv.","n_serie":"N° Série","annee_acquisition":"Année"
        }, max_rows=12))

    _footer(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT PC
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_pc_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_header(f"Rapport Inventaire PC{titre_dir}",
                         f"Parc Informatique Site Tétouan {YEAR_STR}"))
    story.append(AccentBar(PAGE_W))
    story.append(Spacer(1, 10))

    total  = len(df)
    bureau = df[df["type_pc"].str.upper()=="BUREAU"].shape[0]   if "type_pc" in df.columns else 0
    port   = df[df["type_pc"].str.upper()=="PORTABLE"].shape[0] if "type_pc" in df.columns else 0
    chrom  = df[df["type_pc"].str.upper()=="CHROMEBOOK"].shape[0] if "type_pc" in df.columns else 0
    urgent = df[df["priorite"]=="🔴 Urgent"].shape[0]           if "priorite" in df.columns else 0
    n_dir  = df["direction"].nunique() if "direction" in df.columns else 0
    avg_age= round(df["anciennete"].mean(), 1) if "anciennete" in df.columns and df["anciennete"].notna().any() else 0

    # Intro
    story.append(_intro_box(
        f"Ce rapport présente l'inventaire complet de <b>{total} PC</b> actifs "
        f"répartis dans <b>{n_dir} directions</b>. "
        f"Âge moyen du parc: <b>{avg_age} ans</b>. "
        f"PC urgents à remplacer: <b>{urgent}</b>.",
        s))
    story.append(Spacer(1, 8))

    # KPI grid
    story.append(_section_title("Indicateurs Clés", s))
    story.append(Spacer(1, 4))
    story.append(_kpi_grid([
        ("🖥️", "Total PC",    str(total),  TEAL,   f"Inventaire {YEAR_STR}"),
        ("🖥️", "Bureau",      str(bureau), NAVY,   f"{round(bureau/total*100) if total else 0}% du parc"),
        ("💻", "Portable",    str(port),   NAVY2,  f"{round(port/total*100) if total else 0}% du parc"),
        ("🔴", "Urgents",     str(urgent), CORAL,  f"{round(urgent/total*100) if total else 0}% — Remplacement"),
    ], ncols=4))
    story.append(Spacer(1, 12))

    # Répartition par direction
    if "direction" in df.columns and not direction:
        story.append(_section_title("01  Répartition par Direction", s))
        story.append(Spacer(1, 4))
        dir_counts = df.groupby("direction").size().to_dict()
        story.append(_bar_chart_table(dir_counts, color=TEAL))
        story.append(Spacer(1, 10))

    # Ancienneté
    if "categorie_age" in df.columns:
        story.append(_section_title("02  Analyse par Ancienneté", s))
        story.append(Spacer(1, 4))
        age_df = df["categorie_age"].value_counts().reset_index()
        age_df.columns = ["Catégorie","Nombre"]
        age_df["Pourcentage"] = (age_df["Nombre"]/total*100).round(1).astype(str)+"%"
        story.append(_data_table(age_df, {"Catégorie":"Catégorie","Nombre":"Nombre","Pourcentage":"%"},
                                  highlight_prio=False))
        story.append(Spacer(1, 10))

    # RAM
    if "ram_go" in df.columns:
        story.append(_section_title("03  Distribution RAM", s))
        story.append(Spacer(1, 4))
        ram_counts = {f"{int(v)} GO": cnt for v, cnt in
                      df["ram_go"].dropna().value_counts().sort_index().items()}
        story.append(_bar_chart_table(ram_counts, color=ACCENT))
        story.append(Spacer(1, 10))

    # Recommandations
    if "priorite" in df.columns:
        story.append(_section_title("04  Recommandations de Remplacement", s))
        story.append(Spacer(1, 4))
        prio_df = df["priorite"].value_counts().reset_index()
        prio_df.columns = ["priorite","Nombre"]
        prio_df["Pourcentage"] = (prio_df["Nombre"]/total*100).round(1).astype(str)+"%"
        story.append(_data_table(prio_df, {"priorite":"Priorité","Nombre":"Nombre","Pourcentage":"%"}))
        story.append(Spacer(1, 10))

        # Urgents détail
        urg_df = df[df["priorite"]=="🔴 Urgent"]
        if len(urg_df) > 0:
            story.append(PageBreak())
            story.append(_section_title(f"05  Détail PC Urgents ({len(urg_df)})", s))
            story.append(Spacer(1, 4))
            story.append(_data_table(urg_df, {
                "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
                "type_pc":"Type","annee_acquisition":"Année","ram":"RAM",
                "processeur":"Processeur","raisons":"Raisons"
            }, max_rows=80))

    # Inventaire complet si filtré par direction
    if direction:
        story.append(PageBreak())
        story.append(_section_title(f"Inventaire Complet — {direction}", s))
        story.append(Spacer(1, 4))
        story.append(_data_table(df, {
            "utilisateur":"Utilisateur","site":"Site","type_pc":"Type",
            "n_inventaire":"N° Inv.","annee_acquisition":"Année","ram":"RAM",
            "processeur":"Processeur","hdd":"HDD","priorite":"Priorité"
        }))

    # Conclusions
    story.append(Spacer(1, 10))
    story.append(_section_title("Conclusions & Recommandations", s))
    story.append(Spacer(1, 6))
    c1 = _conclusion_box("⚠️", "Plan de Renouvellement Urgent",
                          f"{urgent} PC urgents à remplacer immédiatement. "
                          f"Priorité aux postes de plus de 10 ans ({df['categorie_age'].value_counts().get('Plus de 10 ans',0) if 'categorie_age' in df.columns else 0}).",
                          RED, s)
    c2 = _conclusion_box("✅", "Parc Opérationnel",
                          f"{df['priorite'].value_counts().get('🟢 Opérationnel',0) if 'priorite' in df.columns else 0} PC "
                          f"opérationnels. Âge moyen: {avg_age} ans.",
                          GREEN, s)
    conclusions = Table([[c1, c2]], colWidths=[PAGE_W/2, PAGE_W/2])
    conclusions.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
    ]))
    story.append(conclusions)

    _footer(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT IMPRIMANTES
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_imp_pdf(df: pd.DataFrame, direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Toutes Directions"
    story.append(_header(f"Rapport Imprimantes{titre_dir}",
                         f"Parc Imprimantes Site Tétouan {YEAR_STR}"))
    story.append(AccentBar(PAGE_W))
    story.append(Spacer(1, 10))

    total = len(df)
    n_dir = df["direction"].nunique() if "direction" in df.columns else 0
    reseau = df[df["type_imp"].str.upper().str.contains("RÉSEAU|RESEAU",na=False)].shape[0] if "type_imp" in df.columns else 0

    story.append(_intro_box(
        f"Inventaire complet de <b>{total} imprimantes</b> actives réparties dans "
        f"<b>{n_dir} directions</b>. "
        f"Imprimantes réseau: <b>{reseau}</b> · Monoposte: <b>{total-reseau}</b>.",
        s))
    story.append(Spacer(1, 8))

    story.append(_section_title("Indicateurs Clés", s))
    story.append(Spacer(1, 4))
    story.append(_kpi_grid([
        ("🖨️","Total Imprimantes",str(total),  colors.HexColor("#7D3C98"), f"Inventaire {YEAR_STR}"),
        ("🏢","Directions",       str(n_dir),  TEAL,                        "Couverture"),
        ("🌐","Réseau",           str(reseau), GREEN,                        f"{round(reseau/total*100) if total else 0}% du parc"),
        ("🖨️","Monoposte",       str(total-reseau), NAVY,                   f"{round((total-reseau)/total*100) if total else 0}% du parc"),
    ], ncols=4))
    story.append(Spacer(1, 12))

    if "direction" in df.columns and not direction:
        story.append(_section_title("01  Répartition par Direction", s))
        story.append(Spacer(1, 4))
        dir_counts = df.groupby("direction").size().to_dict()
        story.append(_bar_chart_table(dir_counts, color=colors.HexColor("#9B59B6")))
        story.append(Spacer(1, 10))

    if "modele" in df.columns:
        story.append(_section_title("02  Top Modèles d'Imprimantes", s))
        story.append(Spacer(1, 4))
        df_mod = df[df["modele"].notna()].copy()
        df_mod["mc"] = df_mod["modele"].astype(str).str.strip()
        df_mod = df_mod[~df_mod["mc"].str.match(r'^\d+\.?\d*$', na=False)]
        df_mod = df_mod[df_mod["mc"].str.len()>3]
        mod_counts = df_mod["mc"].value_counts().head(10).to_dict()
        story.append(_bar_chart_table(mod_counts, color=TEAL))
        story.append(Spacer(1, 10))

    story.append(PageBreak())
    titre_inv = f"Inventaire — {direction}" if direction else "Inventaire Complet"
    story.append(_section_title(f"03  {titre_inv}", s))
    story.append(Spacer(1, 4))
    story.append(_data_table(df, {
        "utilisateur":"Utilisateur","direction":"Direction","site":"Site",
        "modele":"Modèle","type_imp":"Type","n_inventaire":"N° Inv.","annee_acquisition":"Année"
    }))

    _footer(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  RAPPORT GÉNÉRAL
# ═══════════════════════════════════════════════════════════════════════════════
def export_rapport_general_pdf(df_inv: pd.DataFrame, df_imp: pd.DataFrame,
                                 direction: str = None) -> bytes:
    buf = io.BytesIO()
    doc = _doc(buf)
    s = _styles()
    story = []

    titre_dir = f" — Direction {direction}" if direction else " — Parc Complet"
    story.append(_header(f"Rapport Général{titre_dir}",
                          f"Inventaire Informatique Complet — Site Tétouan {YEAR_STR}"))
    story.append(AccentBar(PAGE_W))
    story.append(Spacer(1, 10))

    total_pc  = len(df_inv) if df_inv is not None else 0
    total_imp = len(df_imp) if df_imp is not None else 0
    urgent    = df_inv[df_inv["priorite"]=="🔴 Urgent"].shape[0] if (df_inv is not None and "priorite" in df_inv.columns) else 0
    n_dir     = df_inv["direction"].nunique() if (df_inv is not None and "direction" in df_inv.columns) else 0
    avg_age   = round(df_inv["anciennete"].mean(),1) if (df_inv is not None and "anciennete" in df_inv.columns and df_inv["anciennete"].notna().any()) else 0
    old_pct   = round(df_inv[df_inv["categorie_age"]=="Plus de 10 ans"].shape[0]/total_pc*100) if (df_inv is not None and "categorie_age" in df_inv.columns and total_pc>0) else 0

    story.append(_intro_box(
        f"Rapport synthèse du parc informatique complet de <b>{total_pc} PC actifs</b> et "
        f"<b>{total_imp} imprimantes</b> sur <b>{n_dir} directions</b>. "
        f"Âge moyen du parc: <b>{avg_age} ans</b>. "
        f"PC &gt; 10 ans: <b>{old_pct}%</b>. "
        f"Urgents: <b>{urgent}</b>.",
        s))
    story.append(Spacer(1, 8))

    story.append(_section_title("Indicateurs Clés du Parc Global", s))
    story.append(Spacer(1, 4))
    story.append(_kpi_grid([
        ("🖥️","Total PC",          str(total_pc),        TEAL,  f"Inventaire {YEAR_STR}"),
        ("🖨️","Imprimantes",       str(total_imp),       colors.HexColor("#9B59B6"), "Actives"),
        ("🔴","Urgents Remplacement",str(urgent),         CORAL, f"{round(urgent/total_pc*100) if total_pc else 0}% du parc"),
        ("🏢","Directions",         str(n_dir),           NAVY,  "Couverture"),
    ], ncols=4))
    story.append(Spacer(1, 12))

    # Synthèse par direction
    if df_inv is not None and "direction" in df_inv.columns:
        story.append(_section_title("01  Synthèse par Direction", s))
        story.append(Spacer(1, 4))
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
        story.append(_data_table(agg, {c: c.replace("_"," ") if c!="direction" else "Direction" for c in agg.columns}, highlight_prio=False))
        story.append(Spacer(1, 10))

    # Ancienneté
    if df_inv is not None and "categorie_age" in df_inv.columns:
        story.append(_section_title("02  Ancienneté du Parc PC", s))
        story.append(Spacer(1, 4))
        age_counts = df_inv["categorie_age"].value_counts().to_dict()
        story.append(_bar_chart_table(age_counts, color=TEAL))
        story.append(Spacer(1, 10))

    # Top modèles imprimantes
    if df_imp is not None and "modele" in df_imp.columns:
        story.append(_section_title("03  Top Modèles Imprimantes", s))
        story.append(Spacer(1, 4))
        df_mod = df_imp[df_imp["modele"].notna()].copy()
        df_mod["mc"] = df_mod["modele"].astype(str).str.strip()
        df_mod = df_mod[~df_mod["mc"].str.match(r'^\d+\.?\d*$', na=False)]
        mod_counts = df_mod["mc"].value_counts().head(8).to_dict()
        story.append(_bar_chart_table(mod_counts, color=colors.HexColor("#9B59B6")))
        story.append(Spacer(1, 10))

    # Conclusions
    story.append(_section_title("Conclusions & Recommandations", s))
    story.append(Spacer(1, 6))
    c1 = _conclusion_box("⚠️", "Plan de Renouvellement Urgent",
                          f"{urgent} PC urgents. {old_pct}% du parc a plus de 10 ans. "
                          f"Renouvellement prioritaire recommandé.",
                          RED, s)
    c2 = _conclusion_box("🖨️", "Consolidation des Imprimantes",
                          f"{total_imp} imprimantes actives. Recommandation: migrer vers "
                          f"modèles réseau partagés pour réduire les coûts.",
                          TEAL, s)
    c3 = _conclusion_box("✅", "Restitution IMF — Objectif atteint",
                          f"Taux de restitution élevé sur la majorité des directions. "
                          f"HP EliteBook 840 G5 comme matériel de référence.",
                          GREEN, s)
    c4 = _conclusion_box("🚀", "Déploiement & Modernisation",
                          f"La DCL concentre le plus grand parc ({agg['Total_PC'].max() if df_inv is not None and 'direction' in df_inv.columns else '—'} PC). "
                          f"Plan de déploiement IMF à accélérer.",
                          GOLD, s)
    row1 = Table([[c1, c2]], colWidths=[PAGE_W/2, PAGE_W/2])
    row2 = Table([[c3, c4]], colWidths=[PAGE_W/2, PAGE_W/2])
    for row in [row1, row2]:
        row.setStyle(TableStyle([
            ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ]))
    story.append(row1)
    story.append(Spacer(1, 4))
    story.append(row2)

    _footer(story, s)
    doc.build(story)
    buf.seek(0)
    return buf.read()
