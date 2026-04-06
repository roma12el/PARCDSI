"""
app.py — DSI Amendis Tétouan v4
Design moderne style dashboard (image référence)
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime

from data_loader import load_inventaire, load_imprimantes
from pdf_export import (
    export_fiche_utilisateur_pdf, export_rapport_pc_pdf,
    export_rapport_imp_pdf, export_rapport_general_pdf,
)

CURRENT_YEAR = 2026

st.set_page_config(
    page_title="DSI Amendis — Inventaire Tétouan",
    page_icon="🖥️", layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS DESIGN MODERNE ───────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono&display=swap');

*, *::before, *::after { box-sizing: border-box; font-family: 'DM Sans', sans-serif; }

/* ── Main BG */
.stApp { background: #F0F4F8; }
.main .block-container { padding: 1.5rem 2rem 3rem; max-width: 1500px; }

/* ── Sidebar */
[data-testid="stSidebar"] {
    background: #1B2A3B !important;
    border-right: none !important;
    box-shadow: 4px 0 20px rgba(0,0,0,0.15);
}
[data-testid="stSidebar"] * { color: #CBD5E1 !important; font-family: 'DM Sans', sans-serif !important; }
[data-testid="stSidebar"] .stButton button {
    background: linear-gradient(135deg, #C8102E, #E8415B) !important;
    color: white !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important;
    width: 100%; padding: 12px !important; font-size: 14px !important;
    box-shadow: 0 4px 15px rgba(200,16,46,0.4) !important;
    transition: all 0.3s !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(200,16,46,0.5) !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stFileUploader label { color: #94A3B8 !important; font-size: 12px !important; font-weight: 500 !important; }
[data-testid="stSidebar"] .stRadio label { color: #CBD5E1 !important; font-size: 13px !important; }
[data-testid="stSidebar"] hr { border-color: #2D3E50 !important; }

/* ── Top bar */
.top-bar {
    background: white;
    border-radius: 16px;
    padding: 18px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    margin-bottom: 20px;
}
.top-bar-title { font-size: 22px; font-weight: 700; color: #1B2A3B; margin: 0; }
.top-bar-sub   { font-size: 12px; color: #94A3B8; margin: 2px 0 0; }

/* ── KPI Cards */
.kpi-grid { display: grid; gap: 16px; margin-bottom: 20px; }
.kpi-card {
    background: white;
    border-radius: 16px;
    padding: 22px 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border-top: 4px solid #C8102E;
    position: relative;
    overflow: hidden;
    transition: all 0.3s;
}
.kpi-card:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.1); }
.kpi-card::after {
    content: '';
    position: absolute;
    bottom: -20px;
    right: -20px;
    width: 80px;
    height: 80px;
    background: rgba(200,16,46,0.05);
    border-radius: 50%;
}
.kpi-icon  { font-size: 28px; margin-bottom: 8px; }
.kpi-value { font-size: 32px; font-weight: 700; color: #1B2A3B; line-height: 1; }
.kpi-label { font-size: 12px; color: #94A3B8; font-weight: 500; margin-top: 6px; }
.kpi-delta { font-size: 11px; color: #27AE60; font-weight: 600; margin-top: 4px; }

/* ── Section Title */
.section-title {
    font-size: 16px; font-weight: 700; color: #1B2A3B;
    margin: 24px 0 14px;
    display: flex; align-items: center; gap: 8px;
}
.section-title::after {
    content: '';
    flex: 1;
    height: 1px;
    background: #E2E8F0;
    margin-left: 8px;
}

/* ── Chart Cards */
.chart-card {
    background: white;
    border-radius: 16px;
    padding: 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    margin-bottom: 16px;
}

/* ── Data Table */
.dataframe { border-radius: 10px !important; overflow: hidden; }
.dataframe thead tr th {
    background: #1B2A3B !important;
    color: white !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    padding: 12px 16px !important;
}
.dataframe tbody tr:hover td { background: #FFF5F7 !important; }

/* ── Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: white;
    border-radius: 12px 12px 0 0;
    padding: 4px 4px 0;
    gap: 2px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 10px 10px 0 0 !important;
    font-weight: 600 !important;
    color: #64748B !important;
    padding: 10px 24px !important;
    font-size: 13px !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: #C8102E !important;
    color: white !important;
}

/* ── Buttons */
.stButton > button {
    background: linear-gradient(135deg, #C8102E, #E8415B);
    color: white; border: none; border-radius: 10px;
    font-weight: 600; padding: 10px 24px;
    box-shadow: 0 4px 12px rgba(200,16,46,0.3);
    transition: all 0.3s;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(200,16,46,0.4);
}

/* ── Download buttons */
.stDownloadButton > button {
    background: linear-gradient(135deg, #1B4F72, #2980B9) !important;
    color: white !important; border: none !important; border-radius: 10px !important;
    font-weight: 600 !important; padding: 10px 20px !important;
    box-shadow: 0 4px 12px rgba(41,128,185,0.3) !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 20px rgba(41,128,185,0.4) !important;
}

/* ── Profile card */
.profile-hero {
    background: linear-gradient(135deg, #1B2A3B 0%, #C8102E 100%);
    border-radius: 20px;
    padding: 28px 32px;
    color: white;
    margin-bottom: 20px;
    position: relative;
    overflow: hidden;
}
.profile-hero::before {
    content: ''; position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: rgba(255,255,255,0.05);
    border-radius: 50%;
}
.profile-name   { font-size: 26px; font-weight: 700; margin: 0 0 4px; }
.profile-sub    { font-size: 13px; opacity: 0.8; }
.profile-badge  {
    display: inline-block;
    background: rgba(255,255,255,0.2);
    backdrop-filter: blur(10px);
    color: white; padding: 4px 14px;
    border-radius: 20px; font-size: 12px;
    font-weight: 600; margin: 4px 4px 0 0;
    border: 1px solid rgba(255,255,255,0.3);
}
.badge-sensible { background: rgba(200,16,46,0.8) !important; }
.badge-waterp   { background: rgba(39,174,96,0.8) !important; }

/* ── Info tiles */
.info-tile {
    background: white;
    border-radius: 12px;
    padding: 16px;
    border: 1px solid #F1F5F9;
    margin: 6px 0;
    transition: all 0.2s;
}
.info-tile:hover { border-color: #C8102E; box-shadow: 0 0 0 3px rgba(200,16,46,0.08); }
.info-tile-label { font-size: 10px; color: #94A3B8; font-weight: 600; text-transform: uppercase; letter-spacing: .5px; }
.info-tile-value { font-size: 15px; font-weight: 700; color: #1B2A3B; margin-top: 4px; }

/* ── Priority badges */
.prio-urgent  { color: #C8102E; font-weight: 700; }
.prio-prio    { color: #E67E22; font-weight: 700; }
.prio-surv    { color: #F1C40F; font-weight: 700; }
.prio-ok      { color: #27AE60; font-weight: 700; }

/* ── PDF export box */
.pdf-export-box {
    background: linear-gradient(135deg, #F8FAFC, #EEF2FF);
    border: 1px solid #E0E7FF;
    border-radius: 16px;
    padding: 18px 22px;
    margin: 12px 0 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

/* ── Navigation radio */
[data-testid="stSidebar"] .stRadio > div {
    gap: 4px !important;
}
[data-testid="stSidebar"] .stRadio > div > label {
    background: rgba(255,255,255,0.05) !important;
    border-radius: 10px !important;
    padding: 10px 16px !important;
    cursor: pointer !important;
    transition: all 0.2s !important;
    border: 1px solid transparent !important;
}
[data-testid="stSidebar"] .stRadio > div > label:has(input:checked) {
    background: rgba(200,16,46,0.2) !important;
    border-color: #C8102E !important;
    color: white !important;
}

/* ── Alert boxes */
.alert-info {
    background: #EFF6FF; border-left: 4px solid #3B82F6;
    padding: 12px 16px; border-radius: 0 10px 10px 0; margin: 8px 0;
}
.alert-warn {
    background: #FFFBEB; border-left: 4px solid #F59E0B;
    padding: 12px 16px; border-radius: 0 10px 10px 0; margin: 8px 0;
}
.alert-danger {
    background: #FEF2F2; border-left: 4px solid #C8102E;
    padding: 12px 16px; border-radius: 0 10px 10px 0; margin: 8px 0;
}
</style>
""", unsafe_allow_html=True)

# ─── Plotly theme ─────────────────────────────────────────────────────────────
COLORS = {
    "rouge": "#C8102E", "rouge_clair": "#E8415B", "rouge_pale": "#F5A0AE",
    "bleu": "#2980B9", "vert": "#27AE60", "orange": "#E67E22",
    "gris_fonce": "#1B2A3B", "gris": "#64748B",
}
PALETTE = [COLORS["rouge"], COLORS["bleu"], COLORS["vert"], COLORS["orange"],
           "#9B59B6", "#1ABC9C", "#E8415B", "#F5A0AE"]

PLOT_LAYOUT = dict(
    font_family="DM Sans, sans-serif",
    plot_bgcolor="white", paper_bgcolor="white",
    font_color="#1B2A3B", title_font_size=14,
    title_font_color="#C8102E",
    margin=dict(t=50, b=30, l=30, r=20),
    legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor="#E2E8F0", borderwidth=1),
)


def _apply_plot(fig, height=380):
    fig.update_layout(**PLOT_LAYOUT, height=height)
    return fig


# ─── Session State ────────────────────────────────────────────────────────────
for k in ["inv_data", "df_imp"]:
    if k not in st.session_state:
        st.session_state[k] = None


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:20px 16px 24px; text-align:center;">
        <div style="font-size:36px; margin-bottom:8px;">🏢</div>
        <div style="font-size:16px; font-weight:700; color:white; letter-spacing:.5px;">AMENDIS</div>
        <div style="font-size:10px; color:#64748B; letter-spacing:2px; margin-top:2px;">DIRECTION DES SYSTÈMES D'INFORMATION</div>
        <div style="margin-top:10px; padding:4px 14px; background:rgba(200,16,46,0.2); border-radius:20px; display:inline-block;">
            <span style="font-size:11px; color:#F5A0AE; font-weight:600;">SITE TÉTOUAN · 2026</span>
        </div>
    </div>
    <hr style="border-color:#2D3E50; margin:0 0 20px;">
    """, unsafe_allow_html=True)

    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">📁 FICHIERS INVENTAIRE</p>', unsafe_allow_html=True)

    inv_file = st.file_uploader("Inventaire PC (.xlsx)", type=["xlsx","xls"], key="inv_upload",
                                 help="Fichier inventaire PC Tétouan 2026")
    imp_file = st.file_uploader("Imprimantes (.xlsx)",   type=["xlsx","xls"], key="imp_upload",
                                 help="Fichier inventaire imprimantes Tétouan")

    if st.button("⚡  Charger & Analyser", use_container_width=True):
        if inv_file:
            with st.spinner("Chargement inventaire PC..."):
                try:
                    data = load_inventaire(inv_file)
                    st.session_state.inv_data = data
                    pc = data["pc"]
                    chrom = data["chromebooks"]
                    mouvt = data["mouvements"]
                    msg = f"✅ PC: {len(pc)} enregistrements"
                    if chrom is not None: msg += f" | 💻 Chromebooks: {len(chrom)}"
                    if mouvt is not None: msg += f" | 🔄 Mouvements: {len(mouvt)}"
                    st.success(msg)
                except Exception as e:
                    st.error(f"Erreur PC: {e}")
        if imp_file:
            with st.spinner("Chargement imprimantes..."):
                try:
                    df_i = load_imprimantes(imp_file)
                    st.session_state.df_imp = df_i
                    st.success(f"✅ Imprimantes: {len(df_i)} enregistrements")
                except Exception as e:
                    st.error(f"Erreur imprimantes: {e}")

    st.markdown('<hr style="border-color:#2D3E50; margin:20px 0;">', unsafe_allow_html=True)
    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">🔍 NAVIGATION</p>', unsafe_allow_html=True)

    page = st.radio("", [
        "🏠 Tableau de bord",
        "🖥️ Rapport PC",
        "💻 Chromebooks",
        "🔄 Mouvements",
        "🖨️ Rapport Imprimantes",
        "📋 Rapport Général",
        "👤 Profil Utilisateur",
    ], label_visibility="collapsed")

    # Filtres
    st.markdown('<hr style="border-color:#2D3E50; margin:20px 0;">', unsafe_allow_html=True)
    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">🎛️ FILTRES</p>', unsafe_allow_html=True)

    inv_data = st.session_state.inv_data
    df_pc_raw = inv_data["pc"] if inv_data is not None else None

    filter_direction = []
    if df_pc_raw is not None and "direction" in df_pc_raw.columns:
        dirs = sorted(df_pc_raw["direction"].dropna().unique().tolist())
        filter_direction = st.multiselect("Direction(s)", dirs, default=[], key="f_dir")

    filter_site = []
    if df_pc_raw is not None and "site" in df_pc_raw.columns:
        sites = sorted(df_pc_raw["site"].dropna().unique().tolist())
        filter_site = st.multiselect("Site(s)", sites, default=[], key="f_site")

    filter_type = []
    if df_pc_raw is not None and "type_pc" in df_pc_raw.columns:
        types = sorted(df_pc_raw["type_pc"].dropna().unique().tolist())
        filter_type = st.multiselect("Type PC", types, default=[], key="f_type")

    st.markdown(f'<p style="font-size:10px; color:#334155; text-align:center; margin-top:20px;">v4.0 · {datetime.now().strftime("%d/%m/%Y")}</p>', unsafe_allow_html=True)


# ─── Helpers ──────────────────────────────────────────────────────────────────
def apply_filters(df):
    if df is None: return df
    d = df.copy()
    if filter_direction and "direction" in d.columns:
        d = d[d["direction"].isin(filter_direction)]
    if filter_site and "site" in d.columns:
        d = d[d["site"].isin(filter_site)]
    if filter_type and "type_pc" in d.columns:
        d = d[d["type_pc"].isin(filter_type)]
    return d


def kpi_card(col, icon, label, value, color="#C8102E", delta=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:{color}">
            <div class="kpi-icon">{icon}</div>
            <div class="kpi-value" style="color:{color}">{value}</div>
            <div class="kpi-label">{label}</div>
            {"<div class='kpi-delta'>"+delta+"</div>" if delta else ""}
        </div>
        """, unsafe_allow_html=True)


def section(title, icon=""):
    st.markdown(f'<div class="section-title">{icon} {title}</div>', unsafe_allow_html=True)


def val_str(v):
    if v is None: return "N/A"
    if isinstance(v, float) and np.isnan(v): return "N/A"
    s = str(v).strip()
    return s if s not in ("nan", "NaT", "None", "") else "N/A"


def export_excel(df: pd.DataFrame, sheet_name="Données") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book; ws = writer.sheets[sheet_name]
        hdr_fmt = wb.add_format({"bg_color":"#1B2A3B","font_color":"#FFFFFF","bold":True,"border":1,"align":"center"})
        for ci, col in enumerate(df.columns):
            ws.write(0, ci, col, hdr_fmt)
            ws.set_column(ci, ci, max(18, len(str(col))+4))
    buf.seek(0); return buf.read()


def pdf_download_row(label, desc, pdf_bytes, fname):
    c1, c2 = st.columns([4, 1])
    with c1:
        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:14px 18px; border:1px solid #E2E8F0;">
            <div style="font-weight:700; color:#1B2A3B; font-size:14px;">📄 {label}</div>
            <div style="font-size:12px; color:#94A3B8; margin-top:2px;">{desc}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button("⬇️ PDF", data=pdf_bytes,
                           file_name=fname, mime="application/pdf",
                           use_container_width=True)


# ─── TOP BAR ──────────────────────────────────────────────────────────────────
page_clean = page.split(" ", 1)[1] if " " in page else page
st.markdown(f"""
<div class="top-bar">
    <div>
        <div class="top-bar-title">DSI Amendis — Inventaire Informatique Tétouan 2026</div>
        <div class="top-bar-sub">Direction des Systèmes d'Information · {page_clean}</div>
    </div>
    <div style="font-size:12px; color:#94A3B8;">{datetime.now().strftime("%A %d %B %Y")}</div>
</div>
""", unsafe_allow_html=True)

# ─── Data ─────────────────────────────────────────────────────────────────────
inv_data  = st.session_state.inv_data
df_pc_raw = inv_data["pc"]              if inv_data is not None else None
df_mouvt  = inv_data["mouvements"]     if inv_data is not None else None
df_chrom  = inv_data["chromebooks"]    if inv_data is not None else None
df_imp    = st.session_state.df_imp

df_pc_f   = apply_filters(df_pc_raw)
df_imp_f  = apply_filters(df_imp) if df_imp is not None else None
dir_label = ", ".join(filter_direction) if filter_direction else "Toutes directions"
dir_single = filter_direction[0] if len(filter_direction) == 1 else None

if df_pc_raw is None and df_imp is None:
    st.markdown("""
    <div style="text-align:center; padding:80px 20px;">
        <div style="font-size:90px; margin-bottom:16px;">📂</div>
        <h2 style="color:#1B2A3B; font-weight:700;">Chargez vos fichiers inventaire</h2>
        <p style="color:#94A3B8; max-width:500px; margin:auto; font-size:15px;">
            Uploadez vos fichiers Excel via la barre latérale puis cliquez sur<br>
            <strong>⚡ Charger & Analyser</strong>
        </p>
        <div style="margin-top:24px; display:flex; gap:16px; justify-content:center; flex-wrap:wrap;">
            <div style="background:white; border-radius:12px; padding:16px 24px; border:1px solid #E2E8F0; font-size:13px; color:#64748B;">
                🖥️ <strong>Inventaire PC</strong><br><small>Feuilles: Ecart INV, mouvement, Chrombook</small>
            </div>
            <div style="background:white; border-radius:12px; padding:16px 24px; border:1px solid #E2E8F0; font-size:13px; color:#64748B;">
                🖨️ <strong>Inventaire Imprimantes</strong><br><small>Feuilles: DSI, DRH, DCF, DEA...</small>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Tableau de bord":
    df = df_pc_f

    # KPIs
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    total_pc = len(df) if df is not None else 0
    bureau   = df[df["type_pc"].str.upper()=="BUREAU"].shape[0]   if df is not None and "type_pc" in df.columns else 0
    portable = df[df["type_pc"].str.upper()=="PORTABLE"].shape[0] if df is not None and "type_pc" in df.columns else 0
    chrom_n  = df[df["type_pc"].str.upper()=="CHROMEBOOK"].shape[0] if df is not None and "type_pc" in df.columns else 0
    total_imp= len(df_imp_f) if df_imp_f is not None else 0
    urgent   = df[df["priorite"]=="🔴 Urgent"].shape[0] if df is not None and "priorite" in df.columns else 0

    kpi_card(c1, "🖥️", "Total PC",         total_pc,  "#C8102E")
    kpi_card(c2, "🖥️", "PC Bureau",        bureau,    "#2980B9")
    kpi_card(c3, "💻", "PC Portable",      portable,  "#1B4F72")
    kpi_card(c4, "📱", "Chromebooks",      chrom_n,   "#27AE60")
    kpi_card(c5, "🖨️", "Imprimantes",      total_imp, "#9B59B6")
    kpi_card(c6, "🔴", "Urgents",          urgent,    "#C8102E")

    st.markdown("<br>", unsafe_allow_html=True)

    if df is not None:
        # Charts row 1
        c1, c2 = st.columns([3, 2])
        with c1:
            if "direction" in df.columns:
                cnt = df.groupby("direction").size().reset_index(name="count").sort_values("count", ascending=True).tail(12)
                fig = px.bar(cnt, x="count", y="direction", orientation="h",
                              title="Parc PC par Direction",
                              color="count",
                              color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]])
                fig.update_coloraxes(showscale=False)
                st.plotly_chart(_apply_plot(fig, 420), use_container_width=True)
        with c2:
            if "type_pc" in df.columns:
                cnt2 = df["type_pc"].value_counts().reset_index()
                cnt2.columns = ["type","count"]
                fig2 = px.pie(cnt2, values="count", names="type",
                               title="Répartition par Type",
                               color_discrete_sequence=PALETTE, hole=0.5)
                fig2.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(_apply_plot(fig2, 420), use_container_width=True)

        # Charts row 2
        c1, c2, c3 = st.columns(3)
        with c1:
            if "categorie_age" in df.columns:
                order = ["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
                cnt3 = df["categorie_age"].value_counts().reindex(order, fill_value=0).reset_index()
                cnt3.columns = ["cat","count"]
                fig3 = px.bar(cnt3, x="cat", y="count", title="Ancienneté du Parc",
                               color="cat",
                               color_discrete_map={"Récent (≤3 ans)":"#27AE60","3-5 ans":"#2980B9",
                                                   "5-10 ans":"#E67E22","Plus de 10 ans":"#C8102E","Inconnu":"#94A3B8"})
                fig3.update_layout(showlegend=False)
                st.plotly_chart(_apply_plot(fig3, 340), use_container_width=True)
        with c2:
            if "ram_go" in df.columns:
                r = df["ram_go"].dropna().value_counts().sort_index().reset_index()
                r.columns = ["ram","count"]
                r["label"] = r["ram"].apply(lambda x: f"{int(x)} GO")
                fig4 = px.bar(r, x="label", y="count", title="Distribution RAM",
                               color="count",
                               color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]])
                fig4.update_coloraxes(showscale=False)
                st.plotly_chart(_apply_plot(fig4, 340), use_container_width=True)
        with c3:
            if "priorite" in df.columns:
                cnt4 = df["priorite"].value_counts().reset_index()
                cnt4.columns = ["prio","count"]
                colors_map = {"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22",
                               "🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"}
                fig5 = px.bar(cnt4, x="prio", y="count", title="Recommandations",
                               color="prio", color_discrete_map=colors_map)
                fig5.update_layout(showlegend=False)
                st.plotly_chart(_apply_plot(fig5, 340), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT PC
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🖥️ Rapport PC":
    df = df_pc_f
    if df is None:
        st.warning("Chargez d'abord le fichier inventaire PC."); st.stop()

    section("Rapport PC", "🖥️")
    st.caption(f"Filtre actif: **{dir_label}**  |  {len(df)} PC")

    # Export PDF
    col_p1, col_p2, col_p3 = st.columns([3,1,1])
    with col_p1:
        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:14px 18px; border:1px solid #E2E8F0;">
            <b>📄 Rapport PC — {dir_label}</b>
            <div style="font-size:12px; color:#94A3B8;">Statistiques complètes + recommandations + liste des PC urgents</div>
        </div>
        """, unsafe_allow_html=True)
    with col_p2:
        try:
            pdf = export_rapport_pc_pdf(df, direction=dir_single)
            st.download_button("⬇️ PDF Rapport",
                               data=pdf,
                               file_name=f"rapport_pc_{datetime.now().strftime('%Y%m%d')}.pdf",
                               mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"PDF: {e}")
    with col_p3:
        show_cols = [c for c in ["utilisateur","direction","site","type_pc","n_inventaire",
                                   "annee_acquisition","ram","processeur","hdd","priorite","observation"] if c in df.columns]
        xlsx = export_excel(df[show_cols], "Inventaire PC")
        st.download_button("⬇️ Excel",
                           data=xlsx,
                           file_name=f"inventaire_pc_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    tabs = st.tabs(["📊 Statistiques", "🔴 Recommandations", "📋 Données complètes"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            # Par direction
            if "direction" in df.columns:
                cnt = df.groupby("direction").size().reset_index(name="count").sort_values("count", ascending=True)
                fig = px.bar(cnt, x="count", y="direction", orientation="h",
                              title="PC par Direction", color="count",
                              color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]])
                fig.update_coloraxes(showscale=False)
                st.plotly_chart(_apply_plot(fig), use_container_width=True)
        with c2:
            if "categorie_age" in df.columns:
                order = ["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
                cnt3 = df["categorie_age"].value_counts().reindex(order, fill_value=0).reset_index()
                cnt3.columns = ["cat","count"]
                fig3 = px.bar(cnt3, x="cat", y="count", title="Ancienneté",
                               color="cat",
                               color_discrete_map={"Récent (≤3 ans)":"#27AE60","3-5 ans":"#2980B9",
                                                   "5-10 ans":"#E67E22","Plus de 10 ans":"#C8102E","Inconnu":"#94A3B8"})
                fig3.update_layout(showlegend=False)
                st.plotly_chart(_apply_plot(fig3), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            if "ram_go" in df.columns:
                r = df["ram_go"].dropna().value_counts().sort_index().reset_index()
                r.columns = ["ram","count"]; r["label"] = r["ram"].apply(lambda x: f"{int(x)} GO")
                fig4 = px.bar(r, x="label", y="count", title="Distribution RAM",
                               color="count", color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]])
                fig4.update_coloraxes(showscale=False)
                st.plotly_chart(_apply_plot(fig4), use_container_width=True)
        with c2:
            if "processeur" in df.columns:
                proc_map = {}
                for v in df["processeur"].dropna():
                    s = str(v).upper()
                    if "I7" in s: proc_map[v] = "Core i7"
                    elif "I5" in s: proc_map[v] = "Core i5"
                    elif "I3" in s: proc_map[v] = "Core i3"
                    elif "ULTRA" in s: proc_map[v] = "Intel Ultra"
                    elif "CELERON" in s: proc_map[v] = "Celeron"
                    elif "DUAL CORE" in s or "AMD" in s: proc_map[v] = "Dual Core/AMD"
                    elif "XEON" in s: proc_map[v] = "Xeon"
                    else: proc_map[v] = "Autre"
                df2 = df.copy(); df2["proc_cat"] = df2["processeur"].map(proc_map).fillna("Inconnu")
                cnt5 = df2["proc_cat"].value_counts().reset_index(); cnt5.columns = ["proc","count"]
                fig5 = px.pie(cnt5, values="count", names="proc", title="Processeurs",
                               color_discrete_sequence=PALETTE, hole=0.4)
                fig5.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(_apply_plot(fig5), use_container_width=True)

        # Timeline
        if "annee_acquisition" in df.columns:
            cnt6 = df["annee_acquisition"].dropna().value_counts().sort_index().reset_index()
            cnt6.columns = ["annee","count"]; cnt6["annee"] = cnt6["annee"].astype(int)
            fig6 = px.area(cnt6, x="annee", y="count", title="Acquisitions par Année",
                            color_discrete_sequence=["#C8102E"])
            fig6.update_traces(fill="tozeroy", fillcolor="rgba(200,16,46,0.15)")
            st.plotly_chart(_apply_plot(fig6, 280), use_container_width=True)

    with tabs[1]:
        if "priorite" not in df.columns:
            st.info("Recommandations non calculées.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                cnt4 = df["priorite"].value_counts().reset_index()
                cnt4.columns = ["prio","count"]
                colors_map = {"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22",
                               "🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"}
                fig5 = px.bar(cnt4, x="prio", y="count", title="Priorités de Remplacement",
                               color="prio", color_discrete_map=colors_map)
                fig5.update_layout(showlegend=False)
                st.plotly_chart(_apply_plot(fig5), use_container_width=True)
            with c2:
                fig_pie = px.pie(cnt4, values="count", names="prio",
                                  title="Répartition Priorités",
                                  color="prio", color_discrete_map=colors_map, hole=0.4)
                st.plotly_chart(_apply_plot(fig_pie), use_container_width=True)

            for prio, color, bg in [
                ("🔴 Urgent","#C8102E","#FEF2F2"),
                ("🟠 Prioritaire","#E67E22","#FFF7ED"),
                ("🟡 À surveiller","#D97706","#FFFBEB"),
                ("🟢 Opérationnel","#27AE60","#F0FDF4"),
            ]:
                sub = df[df["priorite"]==prio]
                if len(sub) == 0: continue
                with st.expander(f"{prio} — {len(sub)} PC", expanded=(prio=="🔴 Urgent")):
                    sc = [c for c in ["utilisateur","direction","site","type_pc",
                                       "annee_acquisition","ram","processeur","raisons","observation"] if c in sub.columns]
                    st.dataframe(sub[sc], use_container_width=True)
                    try:
                        pdf_urg = export_rapport_pc_pdf(sub, direction=f"{prio}")
                        st.download_button(f"⬇️ PDF — {prio}", data=pdf_urg,
                                           file_name=f"pc_urgent_{datetime.now().strftime('%Y%m%d')}.pdf",
                                           mime="application/pdf")
                    except: pass

    with tabs[2]:
        search = st.text_input("🔍 Rechercher (nom, direction, N° inventaire...)", key="search_pc")
        disp = df.copy()
        if search:
            mask = pd.Series(False, index=disp.index)
            for col in ["utilisateur","direction","n_inventaire","descriptions","site"]:
                if col in disp.columns:
                    mask |= disp[col].astype(str).str.contains(search, case=False, na=False)
            disp = disp[mask]

        sc = [c for c in ["utilisateur","direction","site","type_pc","n_inventaire",
                            "annee_acquisition","anciennete","ram","processeur","hdd",
                            "priorite","recommandation","observation"] if c in disp.columns]
        st.dataframe(disp[sc], use_container_width=True, height=450)
        st.caption(f"{len(disp)} PC affichés")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: CHROMEBOOKS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "💻 Chromebooks":
    section("Chromebooks", "💻")
    if df_chrom is None or len(df_chrom) == 0:
        st.info("Aucun Chromebook trouvé. Vérifiez que le fichier contient une feuille 'Chrombook'.")
    else:
        total = len(df_chrom)
        c1, c2, c3 = st.columns(3)
        kpi_card(c1, "💻", "Total Chromebooks", total, "#27AE60")
        if "direction" in df_chrom.columns:
            kpi_card(c2, "🏢", "Directions",        df_chrom["direction"].nunique(), "#2980B9")
        if "site" in df_chrom.columns:
            kpi_card(c3, "📍", "Sites",             df_chrom["site"].nunique(), "#9B59B6")

        st.markdown("<br>", unsafe_allow_html=True)

        # Export
        xlsx_c = export_excel(df_chrom, "Chromebooks")
        st.download_button("⬇️ Export Excel Chromebooks",
                           data=xlsx_c,
                           file_name=f"chromebooks_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        sc = [c for c in ["utilisateur","direction","site","n_inventaire",
                            "descriptions","annee_acquisition","ram","processeur","hdd","os"] if c in df_chrom.columns]
        st.dataframe(df_chrom[sc], use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MOUVEMENTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔄 Mouvements":
    section("Mouvements PC", "🔄")
    if df_mouvt is None or len(df_mouvt) == 0:
        st.info("Aucun mouvement trouvé. Vérifiez que le fichier contient une feuille 'mouvement 2025'.")
    else:
        st.metric("Total mouvements", len(df_mouvt))

        search_m = st.text_input("🔍 Rechercher", key="search_mouvt")
        disp_m = df_mouvt.copy()
        if search_m:
            mask = pd.Series(False, index=disp_m.index)
            for col in disp_m.columns:
                mask |= disp_m[col].astype(str).str.contains(search_m, case=False, na=False)
            disp_m = disp_m[mask]

        st.dataframe(disp_m, use_container_width=True, height=500)
        st.caption(f"{len(disp_m)} mouvement(s)")

        xlsx_m = export_excel(disp_m, "Mouvements")
        st.download_button("⬇️ Export Excel Mouvements",
                           data=xlsx_m,
                           file_name=f"mouvements_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT IMPRIMANTES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🖨️ Rapport Imprimantes":
    df_imp2 = df_imp_f
    if df_imp2 is None:
        st.warning("Chargez d'abord le fichier inventaire imprimantes."); st.stop()

    section("Rapport Imprimantes", "🖨️")

    c1, c2, c3 = st.columns([3,1,1])
    with c1:
        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:14px 18px; border:1px solid #E2E8F0;">
            <b>📄 Rapport Imprimantes — {dir_label}</b>
            <div style="font-size:12px; color:#94A3B8;">{len(df_imp2)} imprimantes · Toutes directions</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        try:
            pdf_i = export_rapport_imp_pdf(df_imp2, direction=dir_single)
            st.download_button("⬇️ PDF", data=pdf_i,
                               file_name=f"rapport_imprimantes_{datetime.now().strftime('%Y%m%d')}.pdf",
                               mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"PDF: {e}")
    with c3:
        xlsx_i = export_excel(df_imp2, "Imprimantes")
        st.download_button("⬇️ Excel", data=xlsx_i,
                           file_name=f"imprimantes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    c1, c2, c3 = st.columns(3)
    kpi_card(c1, "🖨️", "Total Imprimantes", len(df_imp2), "#9B59B6")
    kpi_card(c2, "🏢", "Directions",        df_imp2["direction"].nunique() if "direction" in df_imp2.columns else 0, "#2980B9")
    kpi_card(c3, "📍", "Sites",             df_imp2["site"].nunique() if "site" in df_imp2.columns else 0, "#27AE60")

    st.markdown("<br>", unsafe_allow_html=True)
    tabs_i = st.tabs(["📊 Statistiques", "📋 Données"])

    with tabs_i[0]:
        c1, c2 = st.columns(2)
        with c1:
            if "direction" in df_imp2.columns:
                cnt = df_imp2.groupby("direction").size().reset_index(name="count").sort_values("count", ascending=True)
                fig = px.bar(cnt, x="count", y="direction", orientation="h",
                              title="Imprimantes par Direction", color="count",
                              color_continuous_scale=[[0,"#E9D5FF"],[1,"#9B59B6"]])
                fig.update_coloraxes(showscale=False)
                st.plotly_chart(_apply_plot(fig, 420), use_container_width=True)
        with c2:
            if "type_imp" in df_imp2.columns:
                cnt2 = df_imp2["type_imp"].value_counts().reset_index()
                cnt2.columns = ["type","count"]
                fig2 = px.pie(cnt2, values="count", names="type", title="Type Réseau / Monoposte",
                               color_discrete_sequence=[COLORS["rouge"], COLORS["bleu"], COLORS["vert"]], hole=0.4)
                fig2.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(_apply_plot(fig2, 420), use_container_width=True)

        # Top modèles
        if "modele" in df_imp2.columns:
            cnt3 = df_imp2["modele"].value_counts().head(10).reset_index()
            cnt3.columns = ["modele","count"]; cnt3 = cnt3.sort_values("count", ascending=True)
            fig3 = px.bar(cnt3, x="count", y="modele", orientation="h", title="Top 10 Modèles",
                           color="count", color_continuous_scale=[[0,"#E9D5FF"],[1,"#9B59B6"]])
            fig3.update_coloraxes(showscale=False)
            st.plotly_chart(_apply_plot(fig3, 400), use_container_width=True)

        # Rapport PDF par direction
        if "direction" in df_imp2.columns:
            st.markdown("---")
            section("Rapports PDF par Direction", "📄")
            dirs_imp = sorted(df_imp2["direction"].dropna().unique().tolist())
            cols_dir = st.columns(min(4, len(dirs_imp)))
            for i, d in enumerate(dirs_imp):
                with cols_dir[i % 4]:
                    sub_d = df_imp2[df_imp2["direction"] == d]
                    try:
                        pdf_d = export_rapport_imp_pdf(sub_d, direction=d)
                        st.download_button(
                            f"📄 {d} ({len(sub_d)})",
                            data=pdf_d,
                            file_name=f"imp_{d}_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"pdf_imp_{d}"
                        )
                    except: pass

    with tabs_i[1]:
        search_i = st.text_input("🔍 Rechercher", key="search_imp")
        disp_i = df_imp2.copy()
        if search_i:
            mask = pd.Series(False, index=disp_i.index)
            for col in ["utilisateur","direction","modele","n_inventaire"]:
                if col in disp_i.columns:
                    mask |= disp_i[col].astype(str).str.contains(search_i, case=False, na=False)
            disp_i = disp_i[mask]
        sc = [c for c in ["utilisateur","direction","site","modele","type_imp","n_inventaire","n_serie","annee_acquisition","etat"] if c in disp_i.columns]
        st.dataframe(disp_i[sc], use_container_width=True, height=450)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT GÉNÉRAL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Rapport Général":
    section("Rapport Général", "📋")

    c1, c2 = st.columns([3,1])
    with c1:
        st.markdown(f"""
        <div style="background:white; border-radius:12px; padding:14px 18px; border:1px solid #E2E8F0;">
            <b>📄 Rapport Général — {dir_label}</b>
            <div style="font-size:12px; color:#94A3B8;">Synthèse complète PC + Imprimantes par direction</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        try:
            pdf_g = export_rapport_general_pdf(df_pc_f, df_imp_f, direction=dir_single)
            st.download_button("⬇️ PDF Global",
                               data=pdf_g,
                               file_name=f"rapport_general_{datetime.now().strftime('%Y%m%d')}.pdf",
                               mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"PDF: {e}")

    # Exports par direction
    if df_pc_f is not None and "direction" in df_pc_f.columns:
        section("Rapports PDF par Direction (PC)", "📁")
        dirs_pc = sorted(df_pc_f["direction"].dropna().unique().tolist())
        cols_d = st.columns(min(4, len(dirs_pc)))
        for i, d in enumerate(dirs_pc):
            with cols_d[i % 4]:
                sub_d = df_pc_f[df_pc_f["direction"] == d]
                try:
                    pdf_d2 = export_rapport_pc_pdf(sub_d, direction=d)
                    st.download_button(
                        f"🖥️ {d} ({len(sub_d)})",
                        data=pdf_d2,
                        file_name=f"pc_{d}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"pdf_dir_{d}"
                    )
                except: pass

    # Tableau synthèse
    if df_pc_f is not None and "direction" in df_pc_f.columns:
        st.markdown("<br>", unsafe_allow_html=True)
        section("Synthèse par Direction", "📊")
        agg = df_pc_f.groupby("direction").agg(Total_PC=("utilisateur","count")).reset_index()
        if "type_pc" in df_pc_f.columns:
            for tp, col in [("BUREAU","Bureau"),("PORTABLE","Portable"),("CHROMEBOOK","Chromebook")]:
                s2 = df_pc_f[df_pc_f["type_pc"]==tp].groupby("direction").size().reset_index(name=col)
                agg = agg.merge(s2, on="direction", how="left")
            agg = agg.fillna(0)
        if "priorite" in df_pc_f.columns:
            urg = df_pc_f[df_pc_f["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)
        if df_imp_f is not None and "direction" in df_imp_f.columns:
            ia = df_imp_f.groupby("direction").size().reset_index(name="Imprimantes")
            agg = agg.merge(ia, on="direction", how="left").fillna(0)
            agg["Imprimantes"] = agg["Imprimantes"].astype(int)
        agg = agg.sort_values("Total_PC", ascending=False)
        st.dataframe(agg, use_container_width=True)
        xlsx_g = export_excel(agg, "Synthèse Générale")
        st.download_button("⬇️ Export Excel",
                           data=xlsx_g,
                           file_name=f"synthese_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PROFIL UTILISATEUR
# ══════════════════════════════════════════════════════════════════════════════
elif page == "👤 Profil Utilisateur":
    section("Recherche Utilisateur", "👤")

    if df_pc_raw is None:
        st.warning("Chargez d'abord le fichier inventaire PC."); st.stop()

    c1, c2 = st.columns([4,1])
    with c1:
        search_name = st.text_input("",
                                     placeholder="🔍 Saisissez le nom d'un utilisateur...",
                                     key="profile_search",
                                     label_visibility="collapsed")
    with c2:
        st.button("Rechercher", use_container_width=True)

    if not search_name:
        # Liste des utilisateurs
        if "utilisateur" in df_pc_raw.columns:
            users_list = sorted(df_pc_raw["utilisateur"].dropna().unique().tolist())
            st.markdown(f"""
            <div style="background:white; border-radius:12px; padding:16px 20px; border:1px solid #E2E8F0; margin-top:12px;">
                <div style="font-weight:700; color:#1B2A3B; margin-bottom:12px;">
                    👥 {len(users_list)} utilisateurs dans l'inventaire
                </div>
            """, unsafe_allow_html=True)
            n_cols = 4
            rows_u = [users_list[i:i+n_cols] for i in range(0, min(len(users_list), 100), n_cols)]
            for row_u in rows_u:
                cols_u = st.columns(n_cols)
                for i, u in enumerate(row_u):
                    with cols_u[i]:
                        st.markdown(f"<span style='font-size:12px; color:#64748B;'>→ {u}</span>", unsafe_allow_html=True)
            if len(users_list) > 100:
                st.markdown(f"<i style='font-size:12px; color:#94A3B8;'>... et {len(users_list)-100} autres</i>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    if "utilisateur" not in df_pc_raw.columns:
        st.error("Colonne 'utilisateur' non trouvée."); st.stop()

    mask = df_pc_raw["utilisateur"].astype(str).str.contains(search_name, case=False, na=False)
    results = df_pc_raw[mask]

    if len(results) == 0:
        st.error(f"Aucun résultat pour : **{search_name}**")
        st.info("💡 Essayez avec juste le nom de famille."); st.stop()

    if len(results) > 1:
        user_options = results["utilisateur"].unique().tolist()
        selected_user = st.selectbox("Plusieurs correspondances — sélectionnez:", user_options)
        user_data = results[results["utilisateur"] == selected_user]
    else:
        user_data = results
        selected_user = results.iloc[0]["utilisateur"]

    row = user_data.iloc[0]

    nom      = val_str(row.get("utilisateur"))
    dir_     = val_str(row.get("direction"))
    site     = val_str(row.get("site"))
    type_pc  = val_str(row.get("type_pc"))
    prio     = val_str(row.get("priorite"), "🟢 Opérationnel")
    rec      = val_str(row.get("recommandation"), "Aucune action")
    raisons  = val_str(row.get("raisons"), "Équipement conforme")
    sensible = val_str(row.get("poste_sensible"), "NON").upper()
    waterp   = val_str(row.get("waterp"), "NON").upper()

    prio_col = {"Urgent":"#C8102E","Prioritaire":"#E67E22","surveiller":"#D97706","Opérationnel":"#27AE60","Operationnel":"#27AE60"}
    pc = next((v for k, v in prio_col.items() if k.lower() in prio.lower()), "#94A3B8")

    # Hero card
    badges = f"""
        <span class="profile-badge">{dir_}</span>
        <span class="profile-badge">{site}</span>
        <span class="profile-badge">{type_pc}</span>
        {"<span class='profile-badge badge-sensible'>⚠️ Poste Sensible</span>" if "OUI" in sensible else ""}
        {"<span class='profile-badge badge-waterp'>💧 WATERP</span>" if "OUI" in waterp else ""}
    """
    st.markdown(f"""
    <div class="profile-hero">
        <div class="profile-name">{nom}</div>
        <div class="profile-sub">{dir_} · {site}</div>
        <div style="margin-top:14px;">{badges}</div>
        <div style="margin-top:16px; background:rgba(255,255,255,0.12); border-radius:10px; padding:12px 16px; backdrop-filter:blur(10px);">
            <span style="font-size:13px; font-weight:600; color:rgba(255,255,255,0.9);">Statut DSI:</span>
            <span style="font-size:13px; color:white; margin:0 8px;">{prio}</span>
            <span style="font-size:13px; color:rgba(255,255,255,0.7);">→ {rec}</span>
            {f'<div style="font-size:11px; color:rgba(255,255,255,0.6); margin-top:4px;">Raisons: {raisons}</div>' if raisons and raisons != "Équipement conforme" else ""}
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Fiche technique
    st.markdown("#### 🖥️ Fiche Technique PC")
    pc_info = {
        "N° Inventaire":     val_str(row.get("n_inventaire")),
        "Modèle":            val_str(row.get("descriptions")),
        "N° Série":          val_str(row.get("n_serie")),
        "Année acquisition": val_str(row.get("annee_acquisition")),
        "Ancienneté":        f"{val_str(row.get('anciennete'))} ans",
        "Catégorie âge":     val_str(row.get("categorie_age")),
        "RAM":               val_str(row.get("ram")),
        "Processeur":        val_str(row.get("processeur")),
        "Disque dur":        val_str(row.get("hdd")),
    }
    cols3 = st.columns(3)
    for i, (lbl, val) in enumerate(pc_info.items()):
        with cols3[i % 3]:
            st.markdown(f"""
            <div class="info-tile">
                <div class="info-tile-label">{lbl}</div>
                <div class="info-tile-value">{val}</div>
            </div>
            """, unsafe_allow_html=True)

    # Observation
    obs = val_str(row.get("observation"))
    if obs and obs != "N/A":
        st.markdown(f"""
        <div class="alert-warn" style="margin-top:12px;">
            <b>📝 Observation:</b> {obs}
        </div>
        """, unsafe_allow_html=True)

    # PC recommandé
    st.markdown("#### 🛒 PC Recommandé par la DSI")
    is_sens = "OUI" in sensible
    is_port = "PORTABLE" in type_pc.upper()
    if is_sens:
        pc_rec = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_r = "Intel Core i7-1355U · 16 GO DDR5 · 512 GO SSD NVMe · Windows 11 Pro"
        justif = "Poste sensible — sécurité renforcée"
    elif is_port:
        pc_rec = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_r = "Intel Core i5-1335U · 8 GO DDR4 · 256 GO SSD · Windows 11"
        justif = "Utilisateur nomade"
    else:
        pc_rec = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_r = "Intel Core i5-12500 · 8 GO DDR4 · 256 GO SSD · Windows 11"
        justif = "Poste fixe standard"

    is_urg = any(x in prio for x in ["Urgent","Prioritaire"])
    rec_style = "alert-danger" if is_urg else "alert-info"
    st.markdown(f"""
    <div class="{rec_style}">
        <div style="font-weight:700; margin-bottom:4px;">🖥️ {pc_rec}</div>
        <div style="font-size:13px;">{specs_r}</div>
        <div style="font-size:12px; color:#64748B; margin-top:4px;">Justification: {justif}</div>
    </div>
    """, unsafe_allow_html=True)

    # Imprimantes direction
    imp_dir_data = None
    if df_imp is not None and "direction" in df_imp.columns:
        imp_dir_data = df_imp[df_imp["direction"].str.upper() == dir_.upper()]
        if len(imp_dir_data) > 0:
            st.markdown(f"#### 🖨️ Imprimantes — Direction {dir_}")
            sc_imp = [c for c in ["modele","type_imp","site","n_inventaire","n_serie","annee_acquisition"] if c in imp_dir_data.columns]
            st.dataframe(imp_dir_data[sc_imp], use_container_width=True)

    # Exports
    st.markdown("#### 📥 Exporter la Fiche")
    c_e1, c_e2 = st.columns(2)

    with c_e1:
        try:
            # Chercher ancien PC dans les mouvements
            ancien_pc = None
            if df_mouvt is not None and "utilisateur" in df_mouvt.columns:
                mask_m = df_mouvt["utilisateur"].astype(str).str.contains(selected_user, case=False, na=False)
                if mask_m.any():
                    ancien_pc = df_mouvt[mask_m].iloc[0]

            pdf_f = export_fiche_utilisateur_pdf(
                row, imp_dir_data if (imp_dir_data is not None and len(imp_dir_data)>0) else None,
                ancien_pc=ancien_pc
            )
            safe = selected_user.replace(" ","_").replace("/","-")
            st.download_button(
                f"📄 Fiche PDF — {nom}",
                data=pdf_f,
                file_name=f"fiche_{safe}_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"PDF: {e}")

    with c_e2:
        fiche_df = pd.DataFrame({
            "Champ":  list(pc_info.keys()) + ["Priorité","Recommandation","Raisons","PC Recommandé","Spécifications","Observation"],
            "Valeur": list(pc_info.values()) + [prio, rec, raisons, pc_rec, specs_r, obs],
        })
        safe = selected_user.replace(" ","_").replace("/","-")
        st.download_button(
            f"📊 Fiche Excel — {nom}",
            data=export_excel(fiche_df, "Profil"),
            file_name=f"profil_{safe}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
