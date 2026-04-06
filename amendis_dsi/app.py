"""
app.py — DSI Amendis Tétouan v5
Design rapport professionnel + fiche utilisateur spec-sheet
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

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&display=swap');
*, *::before, *::after { box-sizing: border-box; font-family: 'DM Sans', sans-serif; }
.stApp { background: #F0F4F8; }
.main .block-container { padding: 1.2rem 1.8rem 3rem; max-width: 1600px; }

[data-testid="stSidebar"] { background: #1B2A3B !important; border-right: none !important; box-shadow: 4px 0 20px rgba(0,0,0,0.15); }
[data-testid="stSidebar"] * { color: #CBD5E1 !important; font-family: 'DM Sans', sans-serif !important; }
[data-testid="stSidebar"] .stButton button { background: linear-gradient(135deg, #C8102E, #E8415B) !important; color: white !important; border: none !important; border-radius: 10px !important; font-weight: 700 !important; width: 100%; padding: 12px !important; box-shadow: 0 4px 15px rgba(200,16,46,0.4) !important; }
[data-testid="stSidebar"] .stRadio > div > label { background: rgba(255,255,255,0.05) !important; border-radius: 10px !important; padding: 10px 16px !important; cursor: pointer !important; border: 1px solid transparent !important; transition: all 0.2s !important; }
[data-testid="stSidebar"] .stRadio > div > label:has(input:checked) { background: rgba(200,16,46,0.2) !important; border-color: #C8102E !important; }
[data-testid="stSidebar"] hr { border-color: #2D3E50 !important; }

.top-bar { background: linear-gradient(135deg, #1B2A3B 0%, #2D3E50 100%); border-radius: 16px; padding: 20px 28px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 20px rgba(27,42,59,0.3); margin-bottom: 20px; }
.top-bar-title { font-size: 20px; font-weight: 800; color: white; margin: 0; }
.top-bar-sub { font-size: 12px; color: #94A3B8; margin: 3px 0 0; }
.top-bar-badge { background: rgba(200,16,46,0.25); border: 1px solid rgba(200,16,46,0.5); color: #F5A0AE !important; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; }

.kpi-card { background: white; border-radius: 14px; padding: 20px 22px; box-shadow: 0 2px 12px rgba(0,0,0,0.06); border-left: 5px solid #C8102E; position: relative; overflow: hidden; transition: all 0.3s; }
.kpi-card:hover { transform: translateY(-3px); box-shadow: 0 8px 25px rgba(0,0,0,0.1); }
.kpi-icon { font-size: 26px; margin-bottom: 6px; }
.kpi-value { font-size: 34px; font-weight: 800; line-height: 1; }
.kpi-label { font-size: 11px; color: #64748B; font-weight: 600; margin-top: 5px; text-transform: uppercase; letter-spacing: 0.5px; }
.kpi-sub { font-size: 11px; color: #94A3B8; margin-top: 3px; }

.rpt-header { background: linear-gradient(135deg, #1B2A3B, #2D3E50); color: white; border-radius: 10px; padding: 10px 20px; font-weight: 700; font-size: 14px; margin: 20px 0 12px; letter-spacing: 0.5px; }
.rpt-header-red { background: linear-gradient(135deg, #C8102E, #E8415B); color: white; border-radius: 10px; padding: 10px 20px; font-weight: 700; font-size: 14px; margin: 20px 0 12px; }

.stTabs [data-baseweb="tab-list"] { background: white; border-radius: 12px 12px 0 0; padding: 4px 4px 0; gap: 2px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
.stTabs [data-baseweb="tab"] { background: transparent !important; border-radius: 10px 10px 0 0 !important; font-weight: 600 !important; color: #64748B !important; padding: 10px 22px !important; font-size: 13px !important; border: none !important; }
.stTabs [aria-selected="true"] { background: #C8102E !important; color: white !important; }

.stButton > button { background: linear-gradient(135deg, #C8102E, #E8415B); color: white; border: none; border-radius: 10px; font-weight: 700; }
.stDownloadButton > button { background: linear-gradient(135deg, #1B4F72, #2980B9) !important; color: white !important; border: none !important; border-radius: 10px !important; font-weight: 600 !important; }

/* SPEC SHEET FICHE UTILISATEUR */
.spec-hero { background: linear-gradient(135deg, #1B2A3B 0%, #C8102E 100%); border-radius: 18px; padding: 28px 32px; color: white; margin-bottom: 24px; position: relative; overflow: hidden; }
.spec-name { font-size: 28px; font-weight: 800; margin: 0 0 4px; }
.spec-sub { font-size: 13px; opacity: 0.75; margin-bottom: 16px; }
.spec-badge { display: inline-block; background: rgba(255,255,255,0.18); backdrop-filter: blur(8px); color: white !important; padding: 5px 14px; border-radius: 20px; font-size: 12px; font-weight: 700; margin: 3px 4px 0 0; border: 1px solid rgba(255,255,255,0.3); }
.spec-badge-red { background: rgba(200,16,46,0.75) !important; }
.spec-badge-green { background: rgba(39,174,96,0.75) !important; }
.spec-status { background: rgba(255,255,255,0.12); border-radius: 12px; padding: 14px 18px; margin-top: 18px; backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.15); }

.spec-table { background: white; border-radius: 14px; overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,0.06); margin-bottom: 20px; }
.spec-table-header { background: #1B2A3B; color: white; padding: 12px 20px; font-weight: 700; font-size: 13px; letter-spacing: 0.5px; }
.spec-row { display: flex; border-bottom: 1px solid #F1F5F9; }
.spec-row:hover { background: #FFF5F7; }
.spec-row:last-child { border-bottom: none; }
.spec-key { width: 40%; padding: 12px 20px; font-size: 12px; font-weight: 600; color: #64748B; text-transform: uppercase; letter-spacing: 0.3px; background: #FAFBFC; border-right: 1px solid #F1F5F9; }
.spec-val { width: 60%; padding: 12px 20px; font-size: 14px; font-weight: 600; color: #1B2A3B; }

.alert-info { background:#EFF6FF; border-left:4px solid #3B82F6; padding:12px 16px; border-radius:0 10px 10px 0; margin:8px 0; }
.alert-warn { background:#FFFBEB; border-left:4px solid #F59E0B; padding:12px 16px; border-radius:0 10px 10px 0; margin:8px 0; }
.alert-danger { background:#FEF2F2; border-left:4px solid #C8102E; padding:12px 16px; border-radius:0 10px 10px 0; margin:8px 0; }
.alert-ok { background:#F0FDF4; border-left:4px solid #27AE60; padding:12px 16px; border-radius:0 10px 10px 0; margin:8px 0; }
</style>
""", unsafe_allow_html=True)

COLORS = {"rouge":"#C8102E","rouge_clair":"#E8415B","rouge_pale":"#F5A0AE","bleu":"#2980B9","vert":"#27AE60","orange":"#E67E22","gris_fonce":"#1B2A3B","gris":"#64748B","violet":"#9B59B6"}
PALETTE = ["#C8102E","#2980B9","#27AE60","#E67E22","#9B59B6","#1ABC9C","#E8415B","#F39C12"]
PLOT_LAYOUT = dict(font_family="DM Sans, sans-serif", plot_bgcolor="white", paper_bgcolor="white", font_color="#1B2A3B", title_font_size=13, title_font_color="#C8102E", margin=dict(t=45,b=25,l=25,r=15), legend=dict(bgcolor="rgba(255,255,255,0.9)",bordercolor="#E2E8F0",borderwidth=1))
def _fig(fig, h=360): fig.update_layout(**PLOT_LAYOUT, height=h); return fig

for k in ["inv_data","df_imp"]:
    if k not in st.session_state: st.session_state[k] = None

# ─── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""<div style="padding:20px 16px 24px; text-align:center;"><div style="font-size:40px; margin-bottom:8px;">🏢</div><div style="font-size:16px; font-weight:800; color:white; letter-spacing:1px;">AMENDIS</div><div style="font-size:10px; color:#64748B; letter-spacing:2px; margin-top:2px;">DIRECTION DES SYSTÈMES D'INFORMATION</div><div style="margin-top:10px; padding:4px 14px; background:rgba(200,16,46,0.2); border-radius:20px; display:inline-block;"><span style="font-size:11px; color:#F5A0AE; font-weight:600;">SITE TÉTOUAN · 2026</span></div></div><hr style="border-color:#2D3E50; margin:0 0 20px;">""", unsafe_allow_html=True)
    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">📁 FICHIERS INVENTAIRE</p>', unsafe_allow_html=True)
    inv_file = st.file_uploader("Inventaire PC (.xlsx)", type=["xlsx","xls"], key="inv_upload")
    imp_file = st.file_uploader("Imprimantes (.xlsx)",   type=["xlsx","xls"], key="imp_upload")
    if st.button("⚡  Charger & Analyser", use_container_width=True):
        if inv_file:
            with st.spinner("Chargement PC..."):
                try:
                    data = load_inventaire(inv_file); st.session_state.inv_data = data
                    pc=data["pc"]; chrom=data["chromebooks"]; mouvt=data["mouvements"]
                    msg = f"✅ PC: {len(pc)}"
                    if chrom is not None: msg += f" | 💻 {len(chrom)}"
                    if mouvt is not None: msg += f" | 🔄 {len(mouvt)}"
                    st.success(msg)
                except Exception as e: st.error(f"Erreur PC: {e}")
        if imp_file:
            with st.spinner("Chargement imprimantes..."):
                try:
                    df_i = load_imprimantes(imp_file); st.session_state.df_imp = df_i; st.success(f"✅ Imprimantes: {len(df_i)}")
                except Exception as e: st.error(f"Erreur: {e}")
    st.markdown('<hr style="border-color:#2D3E50; margin:20px 0;">', unsafe_allow_html=True)
    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">🔍 NAVIGATION</p>', unsafe_allow_html=True)
    page = st.radio("", ["🏠 Tableau de bord","🖥️ Rapport PC","💻 Chromebooks","🔄 Mouvements","🖨️ Rapport Imprimantes","📋 Rapport Général","👤 Profil Utilisateur"], label_visibility="collapsed")
    st.markdown('<hr style="border-color:#2D3E50; margin:20px 0;">', unsafe_allow_html=True)
    st.markdown('<p style="font-size:11px; color:#64748B; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">🎛️ FILTRES</p>', unsafe_allow_html=True)
    inv_data = st.session_state.inv_data
    df_pc_raw0 = inv_data["pc"] if inv_data is not None else None
    filter_direction, filter_site, filter_type = [], [], []
    if df_pc_raw0 is not None and "direction" in df_pc_raw0.columns:
        filter_direction = st.multiselect("Direction(s)", sorted(df_pc_raw0["direction"].dropna().unique().tolist()), default=[], key="f_dir")
    if df_pc_raw0 is not None and "site" in df_pc_raw0.columns:
        filter_site = st.multiselect("Site(s)", sorted(df_pc_raw0["site"].dropna().unique().tolist()), default=[], key="f_site")
    if df_pc_raw0 is not None and "type_pc" in df_pc_raw0.columns:
        filter_type = st.multiselect("Type PC", sorted(df_pc_raw0["type_pc"].dropna().unique().tolist()), default=[], key="f_type")
    st.markdown(f'<p style="font-size:10px; color:#475569; text-align:center; margin-top:20px;">v5.0 · {datetime.now().strftime("%d/%m/%Y")}</p>', unsafe_allow_html=True)

# ─── Helpers ───────────────────────────────────────────────────────────────────
def apply_filters(df):
    if df is None: return df
    d = df.copy()
    if filter_direction and "direction" in d.columns: d = d[d["direction"].isin(filter_direction)]
    if filter_site      and "site"      in d.columns: d = d[d["site"].isin(filter_site)]
    if filter_type      and "type_pc"   in d.columns: d = d[d["type_pc"].isin(filter_type)]
    return d

def val_str(v, default="N/A"):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return s if s not in ("nan","NaT","None","") else default

def export_excel(df, sheet_name="Données"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book; ws = writer.sheets[sheet_name]
        hdr = wb.add_format({"bg_color":"#1B2A3B","font_color":"#FFFFFF","bold":True,"border":1,"align":"center"})
        for ci, col in enumerate(df.columns):
            ws.write(0, ci, col, hdr); ws.set_column(ci, ci, max(16, len(str(col))+4))
    buf.seek(0); return buf.read()

def kpi(col, icon, label, value, color="#C8102E", sub=""):
    with col:
        st.markdown(f"""<div class="kpi-card" style="border-left-color:{color}"><div class="kpi-icon">{icon}</div><div class="kpi-value" style="color:{color}">{value}</div><div class="kpi-label">{label}</div>{"<div class='kpi-sub'>"+sub+"</div>" if sub else ""}</div>""", unsafe_allow_html=True)

def rpt_hdr(title, red=False):
    cls = "rpt-header-red" if red else "rpt-header"
    st.markdown(f'<div class="{cls}">{title}</div>', unsafe_allow_html=True)

def prio_color(prio):
    m = {"Urgent":"#C8102E","Prioritaire":"#E67E22","surveiller":"#D97706","Opérationnel":"#27AE60"}
    return next((v for k,v in m.items() if k.lower() in str(prio).lower()), "#94A3B8")

def cat_proc(v):
    s = str(v).upper()
    if "I7" in s or "ULTRA 7" in s: return "Core i7/Ultra7"
    if "I5" in s or "ULTRA 5" in s: return "Core i5/Ultra5"
    if "I3" in s: return "Core i3"
    if "CELERON" in s or "DUAL" in s: return "Celeron/Dual"
    if "XEON" in s: return "Xeon"
    return "Autre"

# ─── Data ──────────────────────────────────────────────────────────────────────
inv_data  = st.session_state.inv_data
df_pc_raw = inv_data["pc"]          if inv_data is not None else None
df_mouvt  = inv_data["mouvements"]  if inv_data is not None else None
df_chrom  = inv_data["chromebooks"] if inv_data is not None else None
df_imp    = st.session_state.df_imp
df_pc_f   = apply_filters(df_pc_raw)
df_imp_f  = apply_filters(df_imp) if df_imp is not None else None
dir_label  = ", ".join(filter_direction) if filter_direction else "Toutes directions"
dir_single = filter_direction[0] if len(filter_direction)==1 else None

page_clean = page.split(" ",1)[1] if " " in page else page
st.markdown(f"""<div class="top-bar"><div><div class="top-bar-title">DSI Amendis — Inventaire Informatique Tétouan 2026</div><div class="top-bar-sub">Direction des Systèmes d'Information · {page_clean}</div></div><div style="display:flex; align-items:center; gap:12px;"><span style="font-size:12px; color:#94A3B8;">{datetime.now().strftime("%A %d %B %Y")}</span><span class="top-bar-badge">Site Tétouan</span></div></div>""", unsafe_allow_html=True)

if df_pc_raw is None and df_imp is None:
    st.markdown("""<div style="text-align:center; padding:80px 20px;"><div style="font-size:90px; margin-bottom:16px;">📂</div><h2 style="color:#1B2A3B; font-weight:800;">Chargez vos fichiers inventaire</h2><p style="color:#94A3B8; max-width:500px; margin:auto; font-size:15px;">Uploadez vos fichiers Excel via la barre latérale puis cliquez sur <strong>⚡ Charger & Analyser</strong></p></div>""", unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Tableau de bord":
    df = df_pc_f
    total_pc = len(df) if df is not None else 0
    bureau   = df[df["type_pc"].str.upper()=="BUREAU"].shape[0]   if df is not None and "type_pc" in df.columns else 0
    portable = df[df["type_pc"].str.upper()=="PORTABLE"].shape[0] if df is not None and "type_pc" in df.columns else 0
    chrom_n  = len(df_chrom) if df_chrom is not None else 0
    total_imp= len(df_imp_f) if df_imp_f is not None else 0
    urgent   = df[df["priorite"]=="🔴 Urgent"].shape[0] if df is not None and "priorite" in df.columns else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    kpi(c1,"🖥️","Total PC",    total_pc,  "#C8102E")
    kpi(c2,"🖥️","Bureau",     bureau,    "#2980B9",  f"{round(bureau/total_pc*100) if total_pc else 0}%")
    kpi(c3,"💻","Portable",   portable,  "#1B4F72",  f"{round(portable/total_pc*100) if total_pc else 0}%")
    kpi(c4,"📱","Chromebooks",chrom_n,   "#27AE60")
    kpi(c5,"🖨️","Imprimantes",total_imp, "#9B59B6")
    kpi(c6,"🔴","Urgents",    urgent,    "#C8102E",  f"{round(urgent/total_pc*100) if total_pc else 0}%")
    st.markdown("<br>", unsafe_allow_html=True)

    if df is not None:
        rpt_hdr("📊 Répartition du Parc PC")
        c1,c2 = st.columns([3,2])
        with c1:
            if "direction" in df.columns:
                cnt = df.groupby("direction").size().reset_index(name="count").sort_values("count")
                fig = px.bar(cnt, x="count", y="direction", orientation="h", title="PC par Direction", color="count", color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]], text="count")
                fig.update_coloraxes(showscale=False); fig.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig,400), use_container_width=True)
        with c2:
            if "type_pc" in df.columns:
                cnt2 = df["type_pc"].value_counts().reset_index(); cnt2.columns=["type","count"]
                fig2 = px.pie(cnt2, values="count", names="type", title="Répartition par Type", color_discrete_sequence=PALETTE, hole=0.55)
                fig2.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(_fig(fig2,400), use_container_width=True)

        rpt_hdr("📈 Ancienneté, RAM & État du Parc")
        c1,c2,c3 = st.columns(3)
        with c1:
            if "categorie_age" in df.columns:
                order=["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
                clr={"Récent (≤3 ans)":"#27AE60","3-5 ans":"#2980B9","5-10 ans":"#E67E22","Plus de 10 ans":"#C8102E","Inconnu":"#94A3B8"}
                cnt3=df["categorie_age"].value_counts().reindex(order,fill_value=0).reset_index(); cnt3.columns=["cat","count"]
                fig3=px.bar(cnt3,x="cat",y="count",title="Ancienneté du Parc",color="cat",color_discrete_map=clr,text="count")
                fig3.update_layout(showlegend=False); fig3.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig3,340), use_container_width=True)
        with c2:
            if "ram_go" in df.columns:
                r=df["ram_go"].dropna().value_counts().sort_index().reset_index(); r.columns=["ram","count"]; r["label"]=r["ram"].apply(lambda x:f"{int(x)} GO")
                fig4=px.bar(r,x="label",y="count",title="Distribution RAM",color="count",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="count")
                fig4.update_coloraxes(showscale=False); fig4.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig4,340), use_container_width=True)
        with c3:
            if "priorite" in df.columns:
                cnt4=df["priorite"].value_counts().reset_index(); cnt4.columns=["prio","count"]
                clr2={"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22","🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"}
                fig5=px.pie(cnt4,values="count",names="prio",title="État du Parc",color="prio",color_discrete_map=clr2,hole=0.5)
                fig5.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig5,340), use_container_width=True)

        rpt_hdr("📅 Timeline & Processeurs")
        c1,c2 = st.columns(2)
        with c1:
            if "annee_acquisition" in df.columns:
                cnt6=df["annee_acquisition"].dropna().value_counts().sort_index().reset_index(); cnt6.columns=["annee","count"]; cnt6["annee"]=cnt6["annee"].astype(int)
                fig6=px.area(cnt6,x="annee",y="count",title="Acquisitions par Année",color_discrete_sequence=["#C8102E"])
                fig6.update_traces(fill="tozeroy",fillcolor="rgba(200,16,46,0.15)",line_width=2.5)
                fig6.add_vline(x=CURRENT_YEAR-5,line_dash="dash",line_color="#E67E22",annotation_text="Seuil 5 ans")
                fig6.add_vline(x=CURRENT_YEAR-10,line_dash="dash",line_color="#C8102E",annotation_text="Seuil 10 ans")
                st.plotly_chart(_fig(fig6,300), use_container_width=True)
        with c2:
            if "processeur" in df.columns:
                df2=df.copy(); df2["proc_cat"]=df2["processeur"].apply(cat_proc)
                cnt5=df2["proc_cat"].value_counts().reset_index(); cnt5.columns=["proc","count"]
                fig5b=px.bar(cnt5,x="proc",y="count",title="Répartition Processeurs",color="proc",color_discrete_sequence=PALETTE,text="count")
                fig5b.update_layout(showlegend=False); fig5b.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig5b,300), use_container_width=True)

        if "priorite" in df.columns and "direction" in df.columns:
            rpt_hdr("🔴 Urgences par Direction", red=True)
            urg_dir=df[df["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents").sort_values("Urgents",ascending=False)
            if len(urg_dir)>0:
                fig_u=px.bar(urg_dir,x="direction",y="Urgents",title="PC Urgents à Remplacer",color="Urgents",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="Urgents")
                fig_u.update_coloraxes(showscale=False); fig_u.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig_u,300), use_container_width=True)

        if "categorie_age" in df.columns and "direction" in df.columns:
            rpt_hdr("🗺️ Carte de Chaleur — Direction × Ancienneté")
            order=["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
            heat=df.groupby(["direction","categorie_age"]).size().unstack(fill_value=0)
            heat=heat.reindex(columns=[c for c in order if c in heat.columns])
            fig_h=px.imshow(heat,color_continuous_scale=[[0,"#F0FDF4"],[0.5,"#FFF7ED"],[1,"#C8102E"]],title="PC par Direction et Ancienneté",text_auto=True,aspect="auto")
            fig_h.update_layout(**PLOT_LAYOUT,height=360)
            st.plotly_chart(fig_h, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT PC
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🖥️ Rapport PC":
    df = df_pc_f
    if df is None: st.warning("Chargez d'abord le fichier inventaire PC."); st.stop()

    total_pc=len(df); urgent=df[df["priorite"]=="🔴 Urgent"].shape[0] if "priorite" in df.columns else 0
    prio_n=df[df["priorite"]=="🟠 Prioritaire"].shape[0] if "priorite" in df.columns else 0
    ok_n=df[df["priorite"]=="🟢 Opérationnel"].shape[0] if "priorite" in df.columns else 0
    nb_dir=df["direction"].nunique() if "direction" in df.columns else 0
    avg_age=df["anciennete"].mean() if "anciennete" in df.columns and df["anciennete"].notna().any() else 0

    st.markdown(f"""<div style="background:linear-gradient(135deg,#1B2A3B,#2D3E50); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;">
        <div style="font-size:22px; font-weight:800; margin-bottom:6px;">📊 Rapport PC — {dir_label}</div>
        <div style="font-size:13px; color:#94A3B8;">{total_pc} ordinateurs · {nb_dir} directions · Âge moyen: {avg_age:.1f} ans · {datetime.now().strftime("%d/%m/%Y")}</div>
        <div style="display:flex; gap:16px; margin-top:16px; flex-wrap:wrap;">
            <div style="background:rgba(200,16,46,0.2); border-radius:10px; padding:10px 18px; border:1px solid rgba(200,16,46,0.4);"><div style="font-size:24px; font-weight:800; color:#F5A0AE;">{urgent}</div><div style="font-size:11px; color:#94A3B8;">🔴 Urgents</div></div>
            <div style="background:rgba(230,126,34,0.2); border-radius:10px; padding:10px 18px; border:1px solid rgba(230,126,34,0.4);"><div style="font-size:24px; font-weight:800; color:#F39C12;">{prio_n}</div><div style="font-size:11px; color:#94A3B8;">🟠 Prioritaires</div></div>
            <div style="background:rgba(39,174,96,0.2); border-radius:10px; padding:10px 18px; border:1px solid rgba(39,174,96,0.4);"><div style="font-size:24px; font-weight:800; color:#27AE60;">{ok_n}</div><div style="font-size:11px; color:#94A3B8;">🟢 Opérationnels</div></div>
        </div></div>""", unsafe_allow_html=True)

    c1,c2,c3=st.columns([4,1,1])
    with c2:
        try:
            pdf=export_rapport_pc_pdf(df,direction=dir_single)
            st.download_button("⬇️ PDF",data=pdf,file_name=f"rapport_pc_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True)
        except Exception as e: st.error(f"PDF: {e}")
    with c3:
        sc=[c for c in ["utilisateur","direction","site","type_pc","n_inventaire","annee_acquisition","ram","processeur","hdd","priorite","observation"] if c in df.columns]
        st.download_button("⬇️ Excel",data=export_excel(df[sc],"Inventaire PC"),file_name=f"pc_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    tabs=st.tabs(["📊 Statistiques","🔴 Recommandations","📋 Données"])

    with tabs[0]:
        rpt_hdr("📊 Répartition & Ancienneté")
        c1,c2=st.columns(2)
        with c1:
            if "direction" in df.columns:
                cnt=df.groupby("direction").size().reset_index(name="count").sort_values("count")
                fig=px.bar(cnt,x="count",y="direction",orientation="h",title="PC par Direction",color="count",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="count")
                fig.update_coloraxes(showscale=False); fig.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig,400),use_container_width=True)
        with c2:
            if "categorie_age" in df.columns:
                order=["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
                clr={"Récent (≤3 ans)":"#27AE60","3-5 ans":"#2980B9","5-10 ans":"#E67E22","Plus de 10 ans":"#C8102E","Inconnu":"#94A3B8"}
                cnt3=df["categorie_age"].value_counts().reindex(order,fill_value=0).reset_index(); cnt3.columns=["cat","count"]
                fig3=px.bar(cnt3,x="cat",y="count",title="Ancienneté",color="cat",color_discrete_map=clr,text="count")
                fig3.update_layout(showlegend=False); fig3.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig3,400),use_container_width=True)

        rpt_hdr("💾 RAM & Processeurs")
        c1,c2=st.columns(2)
        with c1:
            if "ram_go" in df.columns:
                r=df["ram_go"].dropna().value_counts().sort_index().reset_index(); r.columns=["ram","count"]; r["label"]=r["ram"].apply(lambda x:f"{int(x)} GO")
                fig4=px.bar(r,x="label",y="count",title="Distribution RAM",color="count",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="count")
                fig4.update_coloraxes(showscale=False); fig4.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig4,320),use_container_width=True)
        with c2:
            if "processeur" in df.columns:
                df2=df.copy(); df2["proc_cat"]=df2["processeur"].apply(cat_proc)
                cnt5=df2["proc_cat"].value_counts().reset_index(); cnt5.columns=["proc","count"]
                fig5=px.pie(cnt5,values="count",names="proc",title="Processeurs",color_discrete_sequence=PALETTE,hole=0.45)
                fig5.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig5,320),use_container_width=True)

        rpt_hdr("📅 Timeline Acquisitions")
        if "annee_acquisition" in df.columns:
            cnt6=df["annee_acquisition"].dropna().value_counts().sort_index().reset_index(); cnt6.columns=["annee","count"]; cnt6["annee"]=cnt6["annee"].astype(int)
            fig6=px.area(cnt6,x="annee",y="count",title="Acquisitions par Année",color_discrete_sequence=["#C8102E"])
            fig6.update_traces(fill="tozeroy",fillcolor="rgba(200,16,46,0.15)",line_width=2.5)
            fig6.add_vline(x=CURRENT_YEAR-5,line_dash="dash",line_color="#E67E22",annotation_text="5 ans")
            fig6.add_vline(x=CURRENT_YEAR-10,line_dash="dash",line_color="#C8102E",annotation_text="10 ans")
            st.plotly_chart(_fig(fig6,260),use_container_width=True)

        if "categorie_age" in df.columns and "direction" in df.columns:
            rpt_hdr("🗺️ Heatmap Direction × Ancienneté")
            order=["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
            heat=df.groupby(["direction","categorie_age"]).size().unstack(fill_value=0).reindex(columns=[c for c in order if c in df["categorie_age"].unique()])
            fig_h=px.imshow(heat,color_continuous_scale=[[0,"#F0FDF4"],[0.5,"#FFF7ED"],[1,"#C8102E"]],title="PC par Direction et Ancienneté",text_auto=True,aspect="auto")
            fig_h.update_layout(**PLOT_LAYOUT,height=360)
            st.plotly_chart(fig_h,use_container_width=True)

    with tabs[1]:
        if "priorite" not in df.columns: st.info("Recommandations non calculées.")
        else:
            c1,c2=st.columns(2)
            with c1:
                cnt4=df["priorite"].value_counts().reset_index(); cnt4.columns=["prio","count"]
                clr2={"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22","🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"}
                fig5=px.bar(cnt4,x="prio",y="count",title="Priorités",color="prio",color_discrete_map=clr2,text="count")
                fig5.update_layout(showlegend=False); fig5.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig5,300),use_container_width=True)
            with c2:
                fig_pie=px.pie(cnt4,values="count",names="prio",title="Répartition",color="prio",color_discrete_map=clr2,hole=0.5)
                fig_pie.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig_pie,300),use_container_width=True)

            if "direction" in df.columns:
                rpt_hdr("🔴 Urgents + Prioritaires par Direction", red=True)
                urg_dir=df[df["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
                prio_dir=df[df["priorite"]=="🟠 Prioritaire"].groupby("direction").size().reset_index(name="Prioritaires")
                merged=urg_dir.merge(prio_dir,on="direction",how="outer").fillna(0)
                fig_stack=go.Figure()
                fig_stack.add_trace(go.Bar(name="🔴 Urgents",x=merged["direction"],y=merged["Urgents"],marker_color="#C8102E",text=merged["Urgents"].astype(int),textposition="auto"))
                fig_stack.add_trace(go.Bar(name="🟠 Prioritaires",x=merged["direction"],y=merged["Prioritaires"],marker_color="#E67E22",text=merged["Prioritaires"].astype(int),textposition="auto"))
                fig_stack.update_layout(**PLOT_LAYOUT,barmode="stack",height=300,title="Urgents & Prioritaires par Direction")
                st.plotly_chart(fig_stack,use_container_width=True)

            for prio,css in [("🔴 Urgent","alert-danger"),("🟠 Prioritaire","alert-warn"),("🟡 À surveiller","alert-info"),("🟢 Opérationnel","alert-ok")]:
                sub=df[df["priorite"]==prio]
                if len(sub)==0: continue
                with st.expander(f"{prio} — {len(sub)} PC", expanded=(prio=="🔴 Urgent")):
                    sc=[c for c in ["utilisateur","direction","site","type_pc","annee_acquisition","ram","processeur","raisons"] if c in sub.columns]
                    st.dataframe(sub[sc],use_container_width=True)
                    try:
                        pdf_urg=export_rapport_pc_pdf(sub,direction=f"{prio}")
                        st.download_button(f"⬇️ PDF {prio}",data=pdf_urg,file_name=f"pc_urgent_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf")
                    except: pass

    with tabs[2]:
        search=st.text_input("🔍 Rechercher",key="search_pc")
        disp=df.copy()
        if search:
            mask=pd.Series(False,index=disp.index)
            for col in ["utilisateur","direction","n_inventaire","descriptions","site"]:
                if col in disp.columns: mask|=disp[col].astype(str).str.contains(search,case=False,na=False)
            disp=disp[mask]
        sc=[c for c in ["utilisateur","direction","site","type_pc","n_inventaire","annee_acquisition","anciennete","ram","processeur","hdd","priorite","recommandation"] if c in disp.columns]
        st.dataframe(disp[sc],use_container_width=True,height=450)
        st.caption(f"{len(disp)} PC affichés")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: CHROMEBOOKS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "💻 Chromebooks":
    if df_chrom is None or len(df_chrom)==0: st.info("Aucun Chromebook trouvé."); st.stop()
    total=len(df_chrom); nb_dir=df_chrom["direction"].nunique() if "direction" in df_chrom.columns else 0; nb_site=df_chrom["site"].nunique() if "site" in df_chrom.columns else 0
    st.markdown(f"""<div style="background:linear-gradient(135deg,#27AE60,#1ABC9C); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;"><div style="font-size:22px; font-weight:800; margin-bottom:6px;">💻 Rapport Chromebooks</div><div style="font-size:13px; color:rgba(255,255,255,0.75);">{total} Chromebooks · {nb_dir} directions · {nb_site} sites</div></div>""", unsafe_allow_html=True)
    c1,c2,c3=st.columns(3); kpi(c1,"💻","Total",total,"#27AE60"); kpi(c2,"🏢","Directions",nb_dir,"#2980B9"); kpi(c3,"📍","Sites",nb_site,"#9B59B6")
    st.markdown("<br>",unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        if "direction" in df_chrom.columns:
            rpt_hdr("📊 Chromebooks par Direction")
            cnt=df_chrom.groupby("direction").size().reset_index(name="count").sort_values("count")
            fig=px.bar(cnt,x="count",y="direction",orientation="h",title="Par Direction",color="count",color_continuous_scale=[[0,"#A8F0C6"],[1,"#27AE60"]],text="count")
            fig.update_coloraxes(showscale=False); fig.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig,340),use_container_width=True)
    with c2:
        if "site" in df_chrom.columns:
            rpt_hdr("📍 Chromebooks par Site")
            cnt2=df_chrom.groupby("site").size().reset_index(name="count")
            fig2=px.pie(cnt2,values="count",names="site",title="Par Site",color_discrete_sequence=["#27AE60","#1ABC9C","#2ECC71","#A8F0C6"],hole=0.45)
            fig2.update_traces(textposition="inside",textinfo="percent+label")
            st.plotly_chart(_fig(fig2,340),use_container_width=True)
    st.download_button("⬇️ Export Excel",data=export_excel(df_chrom,"Chromebooks"),file_name=f"chromebooks_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    sc=[c for c in ["utilisateur","direction","site","n_inventaire","descriptions","annee_acquisition","ram","processeur","hdd","os"] if c in df_chrom.columns]
    st.dataframe(df_chrom[sc],use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MOUVEMENTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔄 Mouvements":
    if df_mouvt is None or len(df_mouvt)==0: st.info("Aucun mouvement trouvé."); st.stop()
    total_m=len(df_mouvt)
    st.markdown(f"""<div style="background:linear-gradient(135deg,#2980B9,#1B4F72); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;"><div style="font-size:22px; font-weight:800; margin-bottom:6px;">🔄 Mouvements PC</div><div style="font-size:13px; color:rgba(255,255,255,0.75);">{total_m} mouvements enregistrés</div></div>""", unsafe_allow_html=True)
    nb_dir_m=df_mouvt["direction"].nunique() if "direction" in df_mouvt.columns else 0
    c1,c2=st.columns(2); kpi(c1,"🔄","Total Mouvements",total_m,"#2980B9"); kpi(c2,"🏢","Directions",nb_dir_m,"#1B4F72")
    if "direction" in df_mouvt.columns:
        rpt_hdr("📊 Analyse des Mouvements")
        c1,c2=st.columns(2)
        with c1:
            cnt_d=df_mouvt.groupby("direction").size().reset_index(name="count").sort_values("count")
            fig_d=px.bar(cnt_d,x="count",y="direction",orientation="h",title="Mouvements par Direction",color="count",color_continuous_scale=[[0,"#AED6F1"],[1,"#2980B9"]],text="count")
            fig_d.update_coloraxes(showscale=False); fig_d.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig_d,360),use_container_width=True)
        with c2:
            if "type_pc" in df_mouvt.columns:
                cnt_t=df_mouvt["type_pc"].value_counts().reset_index(); cnt_t.columns=["type","count"]
                fig_t=px.pie(cnt_t,values="count",names="type",title="Type Matériel",color_discrete_sequence=PALETTE,hole=0.45)
                fig_t.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig_t,360),use_container_width=True)
    search_m=st.text_input("🔍 Rechercher",key="search_mouvt")
    disp_m=df_mouvt.copy()
    if search_m:
        mask=pd.Series(False,index=disp_m.index)
        for col in disp_m.columns: mask|=disp_m[col].astype(str).str.contains(search_m,case=False,na=False)
        disp_m=disp_m[mask]
    st.dataframe(disp_m,use_container_width=True,height=420)
    st.caption(f"{len(disp_m)} mouvement(s)")
    st.download_button("⬇️ Export Excel",data=export_excel(disp_m,"Mouvements"),file_name=f"mouvements_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT IMPRIMANTES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🖨️ Rapport Imprimantes":
    df_i=df_imp_f
    if df_i is None: st.warning("Chargez d'abord le fichier inventaire imprimantes."); st.stop()
    total_i=len(df_i); nb_dir_i=df_i["direction"].nunique() if "direction" in df_i.columns else 0; nb_site_i=df_i["site"].nunique() if "site" in df_i.columns else 0
    reseau=df_i[df_i["type_imp"].str.upper().str.contains("RÉSEAU|RESEAU",na=False)].shape[0] if "type_imp" in df_i.columns else 0
    st.markdown(f"""<div style="background:linear-gradient(135deg,#7D3C98,#9B59B6); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;"><div style="font-size:22px; font-weight:800; margin-bottom:6px;">🖨️ Rapport Imprimantes — {dir_label}</div><div style="font-size:13px; color:rgba(255,255,255,0.75);">{total_i} imprimantes · {nb_dir_i} directions · {nb_site_i} sites</div></div>""", unsafe_allow_html=True)
    c1,c2,c3,c4=st.columns(4); kpi(c1,"🖨️","Total",total_i,"#9B59B6"); kpi(c2,"🏢","Directions",nb_dir_i,"#7D3C98"); kpi(c3,"📍","Sites",nb_site_i,"#2980B9"); kpi(c4,"🌐","Réseau",reseau,"#27AE60")
    c1,c2=st.columns(2)
    with c1:
        try:
            pdf_i=export_rapport_imp_pdf(df_i,direction=dir_single)
            st.download_button("⬇️ PDF Rapport",data=pdf_i,file_name=f"rapport_imp_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True)
        except Exception as e: st.error(f"PDF: {e}")
    with c2:
        st.download_button("⬇️ Excel",data=export_excel(df_i,"Imprimantes"),file_name=f"imprimantes_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    tabs_i=st.tabs(["📊 Statistiques","📋 Données"])
    with tabs_i[0]:
        rpt_hdr("📊 Répartition des Imprimantes")
        c1,c2=st.columns(2)
        with c1:
            if "direction" in df_i.columns:
                cnt=df_i.groupby("direction").size().reset_index(name="count").sort_values("count")
                fig=px.bar(cnt,x="count",y="direction",orientation="h",title="Par Direction",color="count",color_continuous_scale=[[0,"#E9D5FF"],[1,"#9B59B6"]],text="count")
                fig.update_coloraxes(showscale=False); fig.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig,400),use_container_width=True)
        with c2:
            if "type_imp" in df_i.columns:
                cnt2=df_i["type_imp"].value_counts().reset_index(); cnt2.columns=["type","count"]
                fig2=px.pie(cnt2,values="count",names="type",title="Réseau vs Monoposte",color_discrete_sequence=["#C8102E","#9B59B6","#2980B9"],hole=0.5)
                fig2.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig2,400),use_container_width=True)

        # FIX: Top modèles - filtrer les numéros d'inventaire
        rpt_hdr("🏆 Top Modèles d'Imprimantes")
        if "modele" in df_i.columns:
            df_mod=df_i[df_i["modele"].notna()].copy()
            df_mod["modele_clean"]=df_mod["modele"].astype(str).str.strip()
            df_mod=df_mod[~df_mod["modele_clean"].str.match(r'^\d+\.?\d*$',na=False)]
            df_mod=df_mod[df_mod["modele_clean"].str.len()>3]
            cnt3=df_mod["modele_clean"].value_counts().head(12).reset_index(); cnt3.columns=["modele","count"]; cnt3=cnt3.sort_values("count")
            fig3=px.bar(cnt3,x="count",y="modele",orientation="h",title="Top 12 Modèles",color="count",color_continuous_scale=[[0,"#E9D5FF"],[1,"#7D3C98"]],text="count")
            fig3.update_coloraxes(showscale=False); fig3.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig3,460),use_container_width=True)

        if "annee_acquisition" in df_i.columns:
            rpt_hdr("📅 Acquisitions par Année")
            cnt_yr=df_i["annee_acquisition"].dropna().astype(int).value_counts().sort_index().reset_index(); cnt_yr.columns=["annee","count"]
            fig_yr=px.bar(cnt_yr,x="annee",y="count",title="Imprimantes par année d'acquisition",color="count",color_continuous_scale=[[0,"#E9D5FF"],[1,"#9B59B6"]],text="count")
            fig_yr.update_coloraxes(showscale=False); fig_yr.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig_yr,280),use_container_width=True)

        if "direction" in df_i.columns:
            rpt_hdr("📄 Rapports PDF par Direction")
            dirs_imp=sorted(df_i["direction"].dropna().unique().tolist())
            cols_dir=st.columns(min(4,len(dirs_imp)))
            for i,d in enumerate(dirs_imp):
                with cols_dir[i%4]:
                    sub_d=df_i[df_i["direction"]==d]
                    try:
                        pdf_d=export_rapport_imp_pdf(sub_d,direction=d)
                        st.download_button(f"📄 {d} ({len(sub_d)})",data=pdf_d,file_name=f"imp_{d}_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True,key=f"pdf_imp_{d}")
                    except: pass

    with tabs_i[1]:
        search_i=st.text_input("🔍 Rechercher",key="search_imp")
        disp_i=df_i.copy()
        if search_i:
            mask=pd.Series(False,index=disp_i.index)
            for col in ["utilisateur","direction","modele","n_inventaire"]:
                if col in disp_i.columns: mask|=disp_i[col].astype(str).str.contains(search_i,case=False,na=False)
            disp_i=disp_i[mask]
        sc=[c for c in ["utilisateur","direction","site","modele","type_imp","n_inventaire","n_serie","annee_acquisition","etat"] if c in disp_i.columns]
        st.dataframe(disp_i[sc],use_container_width=True,height=450)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: RAPPORT GÉNÉRAL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Rapport Général":
    total_pc_g=len(df_pc_f) if df_pc_f is not None else 0; total_imp_g=len(df_imp_f) if df_imp_f is not None else 0
    urgent_g=df_pc_f[df_pc_f["priorite"]=="🔴 Urgent"].shape[0] if df_pc_f is not None and "priorite" in df_pc_f.columns else 0
    nb_dir_g=df_pc_f["direction"].nunique() if df_pc_f is not None and "direction" in df_pc_f.columns else 0
    st.markdown(f"""<div style="background:linear-gradient(135deg,#1B2A3B,#2D3E50); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;"><div style="font-size:22px; font-weight:800; margin-bottom:6px;">📋 Rapport Général — {dir_label}</div><div style="font-size:13px; color:#94A3B8;">Synthèse complète PC + Imprimantes · {datetime.now().strftime("%d/%m/%Y")}</div></div>""", unsafe_allow_html=True)
    c1,c2,c3,c4=st.columns(4); kpi(c1,"🖥️","Total PC",total_pc_g,"#1B2A3B"); kpi(c2,"🖨️","Imprimantes",total_imp_g,"#9B59B6"); kpi(c3,"🔴","Urgents",urgent_g,"#C8102E"); kpi(c4,"🏢","Directions",nb_dir_g,"#2980B9")
    c1,c2=st.columns([3,1])
    with c2:
        try:
            pdf_g=export_rapport_general_pdf(df_pc_f,df_imp_f,direction=dir_single)
            st.download_button("⬇️ PDF Global",data=pdf_g,file_name=f"rapport_general_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True)
        except Exception as e: st.error(f"PDF: {e}")

    if df_pc_f is not None and df_imp_f is not None and "direction" in df_pc_f.columns and "direction" in df_imp_f.columns:
        rpt_hdr("📊 PC vs Imprimantes par Direction")
        agg_pc=df_pc_f.groupby("direction").size().reset_index(name="PC")
        agg_imp=df_imp_f.groupby("direction").size().reset_index(name="Imprimantes")
        merged=agg_pc.merge(agg_imp,on="direction",how="outer").fillna(0).sort_values("PC",ascending=False)
        fig_cmp=go.Figure()
        fig_cmp.add_trace(go.Bar(name="PC",x=merged["direction"],y=merged["PC"],marker_color="#C8102E",text=merged["PC"].astype(int),textposition="auto"))
        fig_cmp.add_trace(go.Bar(name="Imprimantes",x=merged["direction"],y=merged["Imprimantes"],marker_color="#9B59B6",text=merged["Imprimantes"].astype(int),textposition="auto"))
        fig_cmp.update_layout(**PLOT_LAYOUT,barmode="group",height=360,title="Parc PC & Imprimantes par Direction")
        st.plotly_chart(fig_cmp,use_container_width=True)

    if df_pc_f is not None and "priorite" in df_pc_f.columns and "direction" in df_pc_f.columns:
        rpt_hdr("🔴 État du Parc PC par Direction", red=True)
        agg=df_pc_f.groupby(["direction","priorite"]).size().reset_index(name="count")
        fig_stack=px.bar(agg,x="direction",y="count",color="priorite",title="Priorités par Direction",color_discrete_map={"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22","🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"},barmode="stack",text="count")
        fig_stack.update_traces(textposition="inside",textfont_size=11)
        st.plotly_chart(_fig(fig_stack,380),use_container_width=True)

    rpt_hdr("📊 Tableau de Synthèse par Direction")
    if df_pc_f is not None and "direction" in df_pc_f.columns:
        agg=df_pc_f.groupby("direction").agg(Total_PC=("utilisateur","count")).reset_index()
        if "type_pc" in df_pc_f.columns:
            for tp,col in [("BUREAU","Bureau"),("PORTABLE","Portable"),("CHROMEBOOK","Chromebook")]:
                s2=df_pc_f[df_pc_f["type_pc"]==tp].groupby("direction").size().reset_index(name=col); agg=agg.merge(s2,on="direction",how="left")
        agg=agg.fillna(0)
        if "priorite" in df_pc_f.columns:
            urg=df_pc_f[df_pc_f["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="🔴 Urgents")
            agg=agg.merge(urg,on="direction",how="left").fillna(0); agg["🔴 Urgents"]=agg["🔴 Urgents"].astype(int)
        if df_imp_f is not None and "direction" in df_imp_f.columns:
            ia=df_imp_f.groupby("direction").size().reset_index(name="Imprimantes"); agg=agg.merge(ia,on="direction",how="left").fillna(0); agg["Imprimantes"]=agg["Imprimantes"].astype(int)
        agg=agg.sort_values("Total_PC",ascending=False)
        st.dataframe(agg,use_container_width=True)
        st.download_button("⬇️ Export Excel",data=export_excel(agg,"Synthèse"),file_name=f"synthese_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if df_pc_f is not None and "direction" in df_pc_f.columns:
        rpt_hdr("📁 Rapports PDF par Direction (PC)")
        dirs_pc=sorted(df_pc_f["direction"].dropna().unique().tolist()); cols_d=st.columns(min(4,len(dirs_pc)))
        for i,d in enumerate(dirs_pc):
            with cols_d[i%4]:
                sub_d=df_pc_f[df_pc_f["direction"]==d]
                try:
                    pdf_d2=export_rapport_pc_pdf(sub_d,direction=d)
                    st.download_button(f"🖥️ {d} ({len(sub_d)})",data=pdf_d2,file_name=f"pc_{d}_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True,key=f"pdf_dir_{d}")
                except: pass

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PROFIL UTILISATEUR — Style spec-sheet HP
# ══════════════════════════════════════════════════════════════════════════════
elif page == "👤 Profil Utilisateur":
    if df_pc_raw is None: st.warning("Chargez d'abord le fichier inventaire PC."); st.stop()
    c1,c2=st.columns([4,1])
    with c1:
        search_name=st.text_input("",placeholder="🔍 Saisissez le nom d'un utilisateur...",key="profile_search",label_visibility="collapsed")
    with c2:
        st.button("Rechercher",use_container_width=True)

    if not search_name:
        if "utilisateur" in df_pc_raw.columns:
            users_list=sorted(df_pc_raw["utilisateur"].dropna().unique().tolist())
            st.markdown(f"""<div style="background:white; border-radius:14px; padding:20px; border:1px solid #E2E8F0; margin-top:12px;"><div style="font-weight:800; color:#1B2A3B; font-size:16px; margin-bottom:14px;">👥 {len(users_list)} utilisateurs dans l'inventaire</div>""", unsafe_allow_html=True)
            rows_u=[users_list[i:i+4] for i in range(0,min(len(users_list),120),4)]
            for row_u in rows_u:
                cols_u=st.columns(4)
                for i,u in enumerate(row_u):
                    with cols_u[i]: st.markdown(f"<span style='font-size:12px; color:#64748B;'>→ {u}</span>",unsafe_allow_html=True)
            if len(users_list)>120: st.markdown(f"<i style='font-size:12px; color:#94A3B8;'>... et {len(users_list)-120} autres</i>",unsafe_allow_html=True)
            st.markdown("</div>",unsafe_allow_html=True)
        st.stop()

    mask=df_pc_raw["utilisateur"].astype(str).str.contains(search_name,case=False,na=False)
    results=df_pc_raw[mask]
    if len(results)==0: st.error(f"Aucun résultat pour : **{search_name}**"); st.info("💡 Essayez avec juste le nom de famille."); st.stop()
    if len(results)>1:
        user_options=results["utilisateur"].unique().tolist()
        selected_user=st.selectbox("Plusieurs correspondances:",user_options); user_data=results[results["utilisateur"]==selected_user]
    else:
        user_data=results; selected_user=results.iloc[0]["utilisateur"]

    row=user_data.iloc[0]
    nom=val_str(row.get("utilisateur")); dir_=val_str(row.get("direction")); site=val_str(row.get("site")); type_pc=val_str(row.get("type_pc"))
    prio=val_str(row.get("priorite"),"🟢 Opérationnel"); rec=val_str(row.get("recommandation"),"Aucune action"); raisons=val_str(row.get("raisons"),"Équipement conforme")
    sensible=val_str(row.get("poste_sensible"),"NON").upper(); waterp_val=val_str(row.get("waterp"),"NON").upper()
    pc_col=prio_color(prio)

    # Badges — FIX: no HTML rendered as text, pure f-string
    s_badge = "<span class='spec-badge spec-badge-red'>⚠️ Poste Sensible</span>" if "OUI" in sensible else ""
    w_badge  = "<span class='spec-badge spec-badge-green'>💧 WATERP</span>"       if "OUI" in waterp_val   else ""
    raisons_html = f"<div style='font-size:11px; color:rgba(255,255,255,0.55); margin-top:6px;'>Raisons : {raisons}</div>" if raisons and raisons != "Équipement conforme" else ""

    st.markdown(f"""
    <div class="spec-hero">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <div class="spec-name">{nom}</div>
                <div class="spec-sub">{dir_} · {site} · {type_pc}</div>
                <div style="margin-top:12px;">
                    <span class="spec-badge">{dir_}</span>
                    <span class="spec-badge">{site}</span>
                    <span class="spec-badge">{type_pc}</span>
                    {s_badge}{w_badge}
                </div>
            </div>
            <div style="opacity:0.4; font-size:56px; line-height:1;">👤</div>
        </div>
        <div class="spec-status">
            <div style="font-size:11px; color:rgba(255,255,255,0.5); margin-bottom:4px; text-transform:uppercase; letter-spacing:1px;">Statut DSI</div>
            <div style="font-size:17px; font-weight:800; color:white;">{prio}</div>
            <div style="font-size:13px; color:rgba(255,255,255,0.8); margin-top:2px;">→ {rec}</div>
            {raisons_html}
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Spec table style HP datasheet
    pc_info = {"N° Inventaire":val_str(row.get("n_inventaire")),"Modèle / Description":val_str(row.get("descriptions")),"N° Série":val_str(row.get("n_serie")),"Année d'acquisition":val_str(row.get("annee_acquisition")),"Ancienneté":f"{val_str(row.get('anciennete'))} ans","Catégorie âge":val_str(row.get("categorie_age")),"Mémoire RAM":val_str(row.get("ram")),"Processeur":val_str(row.get("processeur")),"Disque dur / SSD":val_str(row.get("hdd")),"Direction":dir_,"Site":site,"Type de poste":type_pc,"Poste sensible":sensible,"WATERP":waterp_val}
    rows_html="".join(f'<div class="spec-row"><div class="spec-key">{k}</div><div class="spec-val">{v}</div></div>' for k,v in pc_info.items())
    st.markdown(f'<div class="spec-table"><div class="spec-table-header">🖥️ Fiche Technique PC</div>{rows_html}</div>', unsafe_allow_html=True)

    obs=val_str(row.get("observation"))
    if obs and obs!="N/A": st.markdown(f'<div class="alert-warn"><b>📝 Observation :</b> {obs}</div>',unsafe_allow_html=True)

    is_sens="OUI" in sensible; is_port="PORTABLE" in type_pc.upper()
    if is_sens: pc_rec="HP EliteBook 840 G10 / Dell Latitude 5540"; specs_r="Intel Core i7-1355U · 16 GO DDR5 · 512 GO SSD NVMe · Windows 11 Pro"; justif="Poste sensible — sécurité renforcée"
    elif is_port: pc_rec="HP ProBook 450 G10 / Dell Latitude 3540"; specs_r="Intel Core i5-1335U · 8 GO DDR4 · 256 GO SSD · Windows 11"; justif="Utilisateur nomade"
    else: pc_rec="HP ProDesk 400 G9 MT / Dell OptiPlex 3000"; specs_r="Intel Core i5-12500 · 8 GO DDR4 · 256 GO SSD · Windows 11"; justif="Poste fixe standard"

    rec_rows=f'<div class="spec-row"><div class="spec-key">Modèle recommandé</div><div class="spec-val" style="color:#C8102E; font-weight:800;">{pc_rec}</div></div><div class="spec-row"><div class="spec-key">Spécifications</div><div class="spec-val">{specs_r}</div></div><div class="spec-row"><div class="spec-key">Justification</div><div class="spec-val">{justif}</div></div><div class="spec-row"><div class="spec-key">Urgence</div><div class="spec-val"><span style="color:{pc_col}; font-weight:800;">{prio}</span></div></div>'
    st.markdown(f'<div class="spec-table" style="margin-top:20px;"><div class="spec-table-header">🛒 PC Recommandé par la DSI</div>{rec_rows}</div>', unsafe_allow_html=True)

    imp_dir_data=None
    if df_imp is not None and "direction" in df_imp.columns:
        imp_dir_data=df_imp[df_imp["direction"].str.upper()==dir_.upper()]
        if len(imp_dir_data)>0:
            st.markdown(f'<div class="spec-table" style="margin-top:20px;"><div class="spec-table-header">🖨️ Imprimantes — Direction {dir_} ({len(imp_dir_data)})</div></div>', unsafe_allow_html=True)
            sc_imp=[c for c in ["modele","type_imp","site","n_inventaire","n_serie","annee_acquisition"] if c in imp_dir_data.columns]
            st.dataframe(imp_dir_data[sc_imp],use_container_width=True)

    st.markdown('<div class="spec-table" style="margin-top:20px;"><div class="spec-table-header">📥 Exporter la Fiche</div></div>', unsafe_allow_html=True)
    c_e1,c_e2=st.columns(2)
    with c_e1:
        try:
            ancien_pc=None
            if df_mouvt is not None and "utilisateur" in df_mouvt.columns:
                mask_m=df_mouvt["utilisateur"].astype(str).str.contains(selected_user,case=False,na=False)
                if mask_m.any(): ancien_pc=df_mouvt[mask_m].iloc[0]
            pdf_f=export_fiche_utilisateur_pdf(row,imp_dir_data if (imp_dir_data is not None and len(imp_dir_data)>0) else None,ancien_pc=ancien_pc)
            safe=selected_user.replace(" ","_").replace("/","-")
            st.download_button(f"📄 Fiche PDF — {nom}",data=pdf_f,file_name=f"fiche_{safe}_{datetime.now().strftime('%Y%m%d')}.pdf",mime="application/pdf",use_container_width=True)
        except Exception as e: st.error(f"PDF: {e}")
    with c_e2:
        fiche_df=pd.DataFrame({"Champ":list(pc_info.keys())+["Priorité","Recommandation","Raisons","PC Recommandé","Spécifications","Observation"],"Valeur":list(pc_info.values())+[prio,rec,raisons,pc_rec,specs_r,obs]})
        safe=selected_user.replace(" ","_").replace("/","-")
        st.download_button(f"📊 Fiche Excel — {nom}",data=export_excel(fiche_df,"Profil"),file_name=f"profil_{safe}_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
