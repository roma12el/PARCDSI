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
    page = st.radio("", ["🏠 Tableau de bord","🖥️ Rapport PC","💻 Chromebooks","🔄 Mouvements","🖨️ Rapport Imprimantes","📋 Rapport Général","👤 Profil Utilisateur","👥 Segmentation Utilisateurs","📊 Dashboard HTML"], label_visibility="collapsed")
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
                cnt3=cnt3[cnt3["count"]>0]  # FIX: supprimer les categories vides
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
                cnt3=cnt3[cnt3["count"]>0]  # FIX: supprimer les categories vides (0)
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
            fig.update_coloraxes(showscale=False)
            fig.update_traces(textposition="outside")
            fig.update_layout(yaxis=dict(tickmode="linear", automargin=True), xaxis=dict(showgrid=True, gridcolor="#F2F2F2"))
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

    # FIX: nettoyer type_pc — garder seulement BUREAU, PORTABLE, STATION, CHROMEBOOK
    VALID_TYPES = {"BUREAU","PORTABLE","STATION","CHROMEBOOK","LAPTOP","FIXE"}
    # FIX: supprimer les fausses directions ("Oui","OUI/NON","DIRECTION","SITE"...)
    TRASH_DIRS = {"OUI","NON","OUI/NON","DIRECTION","SITE","TYPE","MODELE",
                  "UTILISATEUR","NOM","PRENOM","NAN","","TOTAL","OBSERVATIONS"}
    mouvt_clean = df_mouvt.copy()
    if "direction" in mouvt_clean.columns:
        mouvt_clean = mouvt_clean[
            ~mouvt_clean["direction"].astype(str).str.upper().str.strip().isin(TRASH_DIRS)
        ]
        mouvt_clean = mouvt_clean[mouvt_clean["direction"].notna()]
    if "type_pc" in mouvt_clean.columns:
        def _norm_type(v):
            s = str(v).upper().strip()
            if s in ("BUREAU","PC BUREAU","FIXE","DESKTOP","PC FIXE"): return "BUREAU"
            if s in ("PORTABLE","LAPTOP","NOTEBOOK","PC PORTABLE"):    return "PORTABLE"
            if s in ("STATION","WORKSTATION"):                          return "STATION"
            if s in ("CHROMEBOOK","CHROME"):                            return "CHROMEBOOK"
            return "BUREAU"
        mouvt_clean["type_pc_clean"] = mouvt_clean["type_pc"].apply(_norm_type)
    else:
        mouvt_clean["type_pc_clean"] = "BUREAU"

    # Raisons de mouvement depuis observation
    OBS_CATS = {
        "Hors Service":  ["HORS SERVICE","H/S","HS","PANNE"],
        "Réaffecté":     ["REAFFECT","RÉAFFECT","AFFECTÉ","AFFECTE"],
        "Stock DSI":     ["STOCK","STOCKAGE"],
        "Restitué":      ["RESTITUÉ","RESTITUE","RENDU"],
        "Remplacé":      ["REMPLACÉ","REMPLACE","REMPLACEMENT","NOUVEAU"],
        "Vol/Perte":     ["VOL","PERDU","PERTE","DISPARU"],
        "En cours":      ["EN COURS","COURS","TRANSIT"],
    }
    def cat_obs(v):
        s = str(v).upper()
        for cat, kws in OBS_CATS.items():
            if any(k in s for k in kws): return cat
        return "Autre" if str(v).strip() not in ("nan","None","—","") else "Non renseigné"
    if "observation" in mouvt_clean.columns:
        mouvt_clean["raison_mouvement"] = mouvt_clean["observation"].apply(cat_obs)
    else:
        mouvt_clean["raison_mouvement"] = "Non renseigné"

    total_m = len(mouvt_clean)
    nb_dir_m = mouvt_clean["direction"].nunique() if "direction" in mouvt_clean.columns else 0

    st.markdown(f"""<div style="background:linear-gradient(135deg,#2980B9,#1B4F72); border-radius:16px; padding:24px 28px; margin-bottom:20px; color:white;">
        <div style="font-size:22px; font-weight:800; margin-bottom:6px;">🔄 Mouvements PC</div>
        <div style="font-size:13px; color:rgba(255,255,255,0.75);">{total_m} mouvements · {nb_dir_m} directions</div>
    </div>""", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    hs_n   = (mouvt_clean["raison_mouvement"]=="Hors Service").sum()
    reaf_n = (mouvt_clean["raison_mouvement"]=="Réaffecté").sum()
    kpi(c1,"🔄","Total Mouvements", total_m,   "#2980B9")
    kpi(c2,"🏢","Directions",       nb_dir_m,  "#1B4F72")
    kpi(c3,"💀","Hors Service",     hs_n,      "#C8102E")
    kpi(c4,"♻️","Réaffectés",       reaf_n,    "#27AE60")
    st.markdown("<br>", unsafe_allow_html=True)

    # Visualisations
    rpt_hdr("📊 Analyse des Mouvements")
    c1,c2 = st.columns(2)
    with c1:
        if "direction" in mouvt_clean.columns:
            cnt_d = mouvt_clean.groupby("direction").size().reset_index(name="count").sort_values("count")
            fig_d = px.bar(cnt_d, x="count", y="direction", orientation="h",
                           title="Mouvements par Direction", color="count",
                           color_continuous_scale=[[0,"#AED6F1"],[1,"#2980B9"]], text="count")
            fig_d.update_coloraxes(showscale=False); fig_d.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig_d,380), use_container_width=True)
    with c2:
        # FIX: utiliser type_pc_clean au lieu de type_pc brut
        cnt_t = mouvt_clean["type_pc_clean"].value_counts().reset_index(); cnt_t.columns=["type","count"]
        fig_t = px.pie(cnt_t, values="count", names="type", title="Type de Matériel",
                       color_discrete_sequence=["#1B4F72","#2980B9","#5DADE2","#AED6F1"], hole=0.50)
        fig_t.update_traces(textposition="inside", textinfo="percent+label", textfont_size=13)
        st.plotly_chart(_fig(fig_t,380), use_container_width=True)

    # Raisons de mouvement
    rpt_hdr("📋 Raisons des Mouvements")
    c1, c2 = st.columns(2)
    with c1:
        cnt_r = mouvt_clean["raison_mouvement"].value_counts().reset_index(); cnt_r.columns=["raison","count"]
        clr_r = {"Hors Service":"#C8102E","Réaffecté":"#27AE60","Stock DSI":"#2980B9",
                 "Restitué":"#9B59B6","Remplacé":"#27AE60","Vol/Perte":"#E67E22",
                 "En cours":"#F39C12","Autre":"#94A3B8","Non renseigné":"#CBD5E1"}
        fig_r = px.bar(cnt_r, x="raison", y="count", title="Raisons des Mouvements",
                       color="raison", color_discrete_map=clr_r, text="count")
        fig_r.update_layout(showlegend=False); fig_r.update_traces(textposition="outside")
        st.plotly_chart(_fig(fig_r,320), use_container_width=True)
    with c2:
        fig_r2 = px.pie(cnt_r, values="count", names="raison", title="Répartition des Raisons",
                        color="raison", color_discrete_map=clr_r, hole=0.45)
        fig_r2.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(_fig(fig_r2,320), use_container_width=True)

    # === RECHERCHE PAR NOM ===
    rpt_hdr("👤 Recherche par Utilisateur")
    col_s, col_b = st.columns([4,1])
    with col_s:
        search_user_m = st.text_input("", placeholder="🔍 Entrez le nom d'un utilisateur pour voir ses mouvements...",
                                       key="search_user_mouvt", label_visibility="collapsed")
    with col_b:
        st.button("Rechercher", key="btn_search_mouvt", use_container_width=True)

    if search_user_m:
        if "utilisateur" in mouvt_clean.columns:
            mask_u = mouvt_clean["utilisateur"].astype(str).str.contains(search_user_m, case=False, na=False)
            user_movs = mouvt_clean[mask_u]
        else:
            user_movs = pd.DataFrame()

        if len(user_movs) == 0:
            st.markdown(f'<div class="alert-warn">⚠️ Aucun mouvement trouvé pour <b>{search_user_m}</b></div>', unsafe_allow_html=True)
        else:
            # Sélection si plusieurs utilisateurs
            users_found = user_movs["utilisateur"].unique().tolist() if "utilisateur" in user_movs.columns else []
            if len(users_found) > 1:
                sel_u = st.selectbox("Plusieurs correspondances :", users_found, key="sel_u_mouvt")
                user_movs = user_movs[user_movs["utilisateur"] == sel_u]
                nom_u = sel_u
            else:
                nom_u = users_found[0] if users_found else search_user_m

            dir_u  = val_str(user_movs.iloc[0].get("direction")) if len(user_movs)>0 else "—"
            site_u = val_str(user_movs.iloc[0].get("site"))      if len(user_movs)>0 else "—"

            # Hero card utilisateur
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#1B4F72,#2980B9); border-radius:14px; padding:20px 24px; margin:12px 0; color:white;">
                <div style="font-size:20px; font-weight:800;">{nom_u}</div>
                <div style="font-size:12px; color:rgba(255,255,255,0.7); margin-top:4px;">{dir_u} · {site_u}</div>
                <div style="display:flex; gap:12px; margin-top:14px; flex-wrap:wrap;">
                    <div style="background:rgba(255,255,255,0.15); border-radius:8px; padding:8px 16px;">
                        <div style="font-size:20px; font-weight:800;">{len(user_movs)}</div>
                        <div style="font-size:10px; color:rgba(255,255,255,0.6);">Mouvements</div>
                    </div>
                    <div style="background:rgba(255,255,255,0.15); border-radius:8px; padding:8px 16px;">
                        <div style="font-size:20px; font-weight:800;">{user_movs["raison_mouvement"].value_counts().index[0] if len(user_movs)>0 else "—"}</div>
                        <div style="font-size:10px; color:rgba(255,255,255,0.6);">Raison principale</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Timeline des mouvements
            rpt_hdr("📜 Historique des Mouvements")
            sc_m = [c for c in ["utilisateur","direction","site","type_pc_clean","n_inventaire","descriptions",
                                  "ram","processeur","hdd","observation","raison_mouvement",
                                  "date_acquisition","Année d'acquisition"] if c in user_movs.columns]
            # Rename pour affichage
            disp_u = user_movs[sc_m].copy()
            rename_map = {"type_pc_clean":"Type PC","raison_mouvement":"Raison","Année d'acquisition":"Année"}
            disp_u = disp_u.rename(columns={k:v for k,v in rename_map.items() if k in disp_u.columns})
            st.dataframe(disp_u, use_container_width=True)

            # Analyse détaillée raisons
            if len(user_movs) > 1:
                rpt_hdr("📊 Analyse Détaillée")
                cnt_u = user_movs["raison_mouvement"].value_counts().reset_index(); cnt_u.columns=["raison","count"]
                fig_u = px.bar(cnt_u, x="raison", y="count", title=f"Raisons des mouvements — {nom_u}",
                               color="raison", color_discrete_map=clr_r, text="count")
                fig_u.update_layout(showlegend=False); fig_u.update_traces(textposition="outside")
                st.plotly_chart(_fig(fig_u, 260), use_container_width=True)

            # Observations détaillées
            if "observation" in user_movs.columns:
                obs_list = user_movs["observation"].dropna().unique().tolist()
                obs_list = [o for o in obs_list if str(o).strip() not in ("nan","None","")]
                if obs_list:
                    rpt_hdr("📝 Observations Enregistrées")
                    for i, obs_val in enumerate(obs_list):
                        cat = cat_obs(obs_val)
                        col_obs = {"Hors Service":"alert-danger","Réaffecté":"alert-ok","Stock DSI":"alert-info",
                                   "Vol/Perte":"alert-danger","Restitué":"alert-info"}.get(cat,"alert-warn")
                        st.markdown(f'<div class="{col_obs}"><b>#{i+1}</b> {obs_val} <span style="float:right;font-size:11px;font-weight:700;">{cat}</span></div>', unsafe_allow_html=True)

    # Tableau global filtrable
    st.markdown("<br>", unsafe_allow_html=True)
    rpt_hdr("📋 Tableau Complet des Mouvements")
    search_m = st.text_input("🔍 Filtrer le tableau (direction, modèle, N° inventaire...)", key="search_mouvt")
    disp_m = mouvt_clean.copy()
    # Supprimer colonnes Unnamed
    disp_m = disp_m[[c for c in disp_m.columns if not str(c).startswith("Unnamed")]]
    if search_m:
        mask = pd.Series(False, index=disp_m.index)
        for col in disp_m.columns:
            mask |= disp_m[col].astype(str).str.contains(search_m, case=False, na=False)
        disp_m = disp_m[mask]
    # Colonnes utiles pour affichage
    sc_all = [c for c in ["utilisateur","direction","site","type_pc_clean","n_inventaire",
                           "descriptions","ram","processeur","hdd","observation","raison_mouvement",
                           "date_acquisition","Année d'acquisition"] if c in disp_m.columns]
    disp_show = disp_m[sc_all].rename(columns={"type_pc_clean":"Type PC","raison_mouvement":"Raison","Année d'acquisition":"Année"})
    st.dataframe(disp_show, use_container_width=True, height=400)
    st.caption(f"{len(disp_m)} mouvement(s) affichés")
    st.download_button("⬇️ Export Excel", data=export_excel(disp_show,"Mouvements"),
                       file_name=f"mouvements_{datetime.now().strftime('%Y%m%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                # FIX: garder seulement Monoposte, Réseau, Traceur — ignorer N° série et modèles mal classés
                def _cat_type_imp(v):
                    s = str(v).upper().strip()
                    if any(k in s for k in ["RÉSEAU","RESEAU","NETWORK","RESEAUX"]): return "Réseau"
                    if any(k in s for k in ["TRACEUR","PLOTTER","TRACT"]): return "Traceur"
                    # Si c'est un numéro de série ou un modèle (contient des chiffres seuls, ou un nom de modèle HP/EPSON...)
                    # → c'est en fait un monoposte mal saisi
                    return "Monoposte"
                df_i2 = df_i.copy()
                df_i2["type_clean"] = df_i2["type_imp"].apply(_cat_type_imp)
                cnt2 = df_i2["type_clean"].value_counts().reset_index(); cnt2.columns=["type","count"]
                clr_pie = {"Monoposte":"#C8102E","Réseau":"#2980B9","Traceur":"#27AE60"}
                fig2=px.pie(cnt2,values="count",names="type",title="Réseau vs Monoposte",
                            color="type",color_discrete_map=clr_pie,hole=0.5)
                fig2.update_traces(textposition="inside",textinfo="percent+label")
                st.plotly_chart(_fig(fig2,400),use_container_width=True)

        # FIX 3: Top modèles avec normalisation pour regrouper les doublons de saisie
        rpt_hdr("🏆 Top Modèles d'Imprimantes")
        if "modele" in df_i.columns:
            import re as _re
            df_mod = df_i[df_i["modele"].notna()].copy()
            df_mod["modele_clean"] = df_mod["modele"].astype(str).str.strip()
            # Filtrer numéros d'inventaire purs
            df_mod = df_mod[~df_mod["modele_clean"].str.match(r'^\d+\.?\d*$', na=False)]
            df_mod = df_mod[df_mod["modele_clean"].str.len() > 3]
            # Normalisation : supprimer espaces multiples, mettre en majuscule, supprimer tirets/points
            def _norm_modele(s):
                s = s.upper().strip()
                s = _re.sub(r'[\s\-\.]+', ' ', s)   # espaces multiples, tirets, points → espace unique
                s = _re.sub(r'\s+', ' ', s).strip()
                return s
            df_mod["modele_norm"] = df_mod["modele_clean"].apply(_norm_modele)
            cnt3 = df_mod["modele_norm"].value_counts().head(12).reset_index()
            cnt3.columns = ["modele","count"]; cnt3 = cnt3.sort_values("count")
            fig3=px.bar(cnt3,x="count",y="modele",orientation="h",title="Top 12 Modèles",
                        color="count",color_continuous_scale=[[0,"#E9D5FF"],[1,"#7D3C98"]],text="count")
            fig3.update_coloraxes(showscale=False); fig3.update_traces(textposition="outside")
            fig3.update_layout(yaxis=dict(tickmode="linear", automargin=True))
            st.plotly_chart(_fig(fig3,480),use_container_width=True)

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
            users_list = sorted(df_pc_raw["utilisateur"].dropna().unique().tolist())
            # Build ONE complete HTML block — never split open/close tags across st.markdown calls
            items_html = ""
            for u in users_list[:200]:
                safe_u = str(u).replace("<","&lt;").replace(">","&gt;")
                items_html += f"<span style='display:inline-block; background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:5px 12px; margin:3px; font-size:12px; color:#1B2A3B; font-weight:500;'>→ {safe_u}</span>"
            extra = f"<div style='margin-top:8px; font-size:12px; color:#94A3B8; font-style:italic;'>... et {len(users_list)-200} autres utilisateurs</div>" if len(users_list) > 200 else ""
            st.markdown(f"""
            <div style="background:white; border-radius:14px; padding:24px; border:1px solid #E2E8F0; margin-top:12px;">
              <div style="font-weight:800; color:#1B2A3B; font-size:16px; margin-bottom:16px;">
                👥 {len(users_list)} utilisateurs dans l'inventaire — saisissez un nom pour rechercher
              </div>
              <div style="line-height:2.2;">{items_html}</div>
              {extra}
            </div>""", unsafe_allow_html=True)
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

    # Build badges safely — all in one HTML string, never split across calls
    badges_parts = [
        f"<span class='spec-badge'>{dir_}</span>",
        f"<span class='spec-badge'>{site}</span>",
        f"<span class='spec-badge'>{type_pc}</span>",
    ]
    if "OUI" in sensible:   badges_parts.append("<span class='spec-badge spec-badge-red'>⚠️ Poste Sensible</span>")
    if "OUI" in waterp_val: badges_parts.append("<span class='spec-badge spec-badge-green'>💧 WATERP</span>")
    badges_html = "".join(badges_parts)
    raisons_html = f"<div style='font-size:11px;color:rgba(255,255,255,0.55);margin-top:6px;'>Raisons : {raisons}</div>" if raisons and raisons not in ("Équipement conforme","N/A","—") else ""

    # ONE complete st.markdown call — never open a div and close it in a separate call
    st.markdown(f"""
    <div class="spec-hero">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;">
        <div>
          <div class="spec-name">{nom}</div>
          <div class="spec-sub">{dir_} · {site} · {type_pc}</div>
          <div style="margin-top:12px;">{badges_html}</div>
        </div>
        <div style="opacity:0.3;font-size:52px;line-height:1;">👤</div>
      </div>
      <div class="spec-status">
        <div style="font-size:11px;color:rgba(255,255,255,0.5);margin-bottom:4px;text-transform:uppercase;letter-spacing:1px;">Statut DSI</div>
        <div style="font-size:17px;font-weight:800;color:white;">{prio}</div>
        <div style="font-size:13px;color:rgba(255,255,255,0.8);margin-top:2px;">→ {rec}</div>
        {raisons_html}
      </div>
    </div>""", unsafe_allow_html=True)

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

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SEGMENTATION UTILISATEURS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "👥 Segmentation Utilisateurs":
    if df_pc_raw is None:
        st.warning("Chargez d'abord le fichier inventaire PC."); st.stop()

    df_seg = df_pc_f.copy() if df_pc_f is not None else df_pc_raw.copy()

    st.markdown("""<div style="background:linear-gradient(135deg,#1B2A3B,#2D3E50);border-radius:16px;padding:24px 28px;margin-bottom:20px;color:white;">
        <div style="font-size:22px;font-weight:800;margin-bottom:6px;">👥 Segmentation des Utilisateurs</div>
        <div style="font-size:13px;color:#94A3B8;">Filtrez et analysez les utilisateurs par profil PC · direction · ancienneté · priorité</div>
    </div>""", unsafe_allow_html=True)

    rpt_hdr("🎛️ Critères de Segmentation")
    cf1,cf2,cf3,cf4 = st.columns(4)
    with cf1:
        dirs_all = ["Toutes"]+sorted(df_seg["direction"].dropna().unique().tolist()) if "direction" in df_seg.columns else ["Toutes"]
        seg_dir  = st.selectbox("Direction", dirs_all, key="seg_dir")
    with cf2:
        prios_all = ["Toutes"]+sorted(df_seg["priorite"].dropna().unique().tolist()) if "priorite" in df_seg.columns else ["Toutes"]
        seg_prio  = st.selectbox("Priorité DSI", prios_all, key="seg_prio")
    with cf3:
        ages_all = ["Toutes","Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
        seg_age  = st.selectbox("Ancienneté PC", ages_all, key="seg_age")
    with cf4:
        types_all = ["Tous"]+sorted(df_seg["type_pc"].dropna().unique().tolist()) if "type_pc" in df_seg.columns else ["Tous"]
        seg_type  = st.selectbox("Type PC", types_all, key="seg_type")

    df_f2 = df_seg.copy()
    if seg_dir  != "Toutes" and "direction"     in df_f2.columns: df_f2 = df_f2[df_f2["direction"]     == seg_dir]
    if seg_prio != "Toutes" and "priorite"      in df_f2.columns: df_f2 = df_f2[df_f2["priorite"]      == seg_prio]
    if seg_age  != "Toutes" and "categorie_age" in df_f2.columns: df_f2 = df_f2[df_f2["categorie_age"] == seg_age]
    if seg_type != "Tous"   and "type_pc"        in df_f2.columns: df_f2 = df_f2[df_f2["type_pc"]       == seg_type]

    n_seg = len(df_f2); n_tot = len(df_seg); pct_seg = round(n_seg/n_tot*100) if n_tot>0 else 0
    st.markdown("<br>",unsafe_allow_html=True)
    if n_seg == 0:
        st.info("Aucun utilisateur ne correspond à ces critères."); st.stop()

    urg_s = df_f2[df_f2["priorite"]=="🔴 Urgent"].shape[0]      if "priorite" in df_f2.columns else 0
    ok_s  = df_f2[df_f2["priorite"]=="🟢 Opérationnel"].shape[0] if "priorite" in df_f2.columns else 0
    avg_r = df_f2["ram_go"].mean()      if "ram_go"    in df_f2.columns and df_f2["ram_go"].notna().any()    else 0
    avg_a = df_f2["anciennete"].mean()  if "anciennete" in df_f2.columns and df_f2["anciennete"].notna().any() else 0
    c1,c2,c3,c4,c5 = st.columns(5)
    kpi(c1,"👥","Utilisateurs",  n_seg,  "#C8102E", f"{pct_seg}% du parc")
    kpi(c2,"🔴","Urgents",       urg_s,  "#E67E22",  f"{round(urg_s/n_seg*100) if n_seg else 0}%")
    kpi(c3,"🟢","Opérationnels", ok_s,   "#27AE60",  f"{round(ok_s/n_seg*100)  if n_seg else 0}%")
    kpi(c4,"💾","RAM Moy.",      f"{avg_r:.0f} GO", "#2980B9")
    kpi(c5,"📅","Âge Moy.",      f"{avg_a:.1f} ans","#1B2A3B")

    st.markdown("<br>",unsafe_allow_html=True)
    rpt_hdr("📊 Profil du Segment")
    c1,c2 = st.columns(2)
    with c1:
        if "direction" in df_f2.columns:
            cnt_d2=df_f2.groupby("direction").size().reset_index(name="count").sort_values("count")
            fig_d2=px.bar(cnt_d2,x="count",y="direction",orientation="h",title="Par Direction",color="count",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="count")
            fig_d2.update_coloraxes(showscale=False); fig_d2.update_traces(textposition="outside")
            fig_d2.update_layout(yaxis=dict(tickmode="linear",automargin=True))
            st.plotly_chart(_fig(fig_d2,380),use_container_width=True)
    with c2:
        if "priorite" in df_f2.columns:
            cnt_p2=df_f2["priorite"].value_counts().reset_index(); cnt_p2.columns=["prio","count"]
            clr_p2={"🔴 Urgent":"#C8102E","🟠 Prioritaire":"#E67E22","🟡 À surveiller":"#F1C40F","🟢 Opérationnel":"#27AE60"}
            fig_p2=px.pie(cnt_p2,values="count",names="prio",title="Répartition Priorités",color="prio",color_discrete_map=clr_p2,hole=0.52)
            fig_p2.update_traces(textposition="inside",textinfo="percent+label")
            st.plotly_chart(_fig(fig_p2,380),use_container_width=True)

    c1,c2,c3=st.columns(3)
    with c1:
        if "categorie_age" in df_f2.columns:
            order_a=["Récent (≤3 ans)","3-5 ans","5-10 ans","Plus de 10 ans","Inconnu"]
            clr_a2={"Récent (≤3 ans)":"#27AE60","3-5 ans":"#2980B9","5-10 ans":"#E67E22","Plus de 10 ans":"#C8102E","Inconnu":"#94A3B8"}
            cnt_a2=df_f2["categorie_age"].value_counts().reindex(order_a,fill_value=0).reset_index(); cnt_a2.columns=["cat","count"]
            fig_a2=px.bar(cnt_a2,x="cat",y="count",title="Ancienneté",color="cat",color_discrete_map=clr_a2,text="count")
            fig_a2.update_layout(showlegend=False); fig_a2.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig_a2,300),use_container_width=True)
    with c2:
        if "ram_go" in df_f2.columns:
            r2=df_f2["ram_go"].dropna().value_counts().sort_index().reset_index(); r2.columns=["ram","count"]; r2["label"]=r2["ram"].apply(lambda x:f"{int(x)} GO")
            fig_r2=px.bar(r2,x="label",y="count",title="Distribution RAM",color="count",color_continuous_scale=[[0,"#F5A0AE"],[1,"#C8102E"]],text="count")
            fig_r2.update_coloraxes(showscale=False); fig_r2.update_traces(textposition="outside")
            st.plotly_chart(_fig(fig_r2,300),use_container_width=True)
    with c3:
        if "type_pc" in df_f2.columns:
            cnt_t2=df_f2["type_pc"].value_counts().reset_index(); cnt_t2.columns=["type","count"]
            fig_t2=px.pie(cnt_t2,values="count",names="type",title="Type de PC",color_discrete_sequence=["#C8102E","#2980B9","#27AE60","#E67E22"],hole=0.48)
            fig_t2.update_traces(textposition="inside",textinfo="percent+label")
            st.plotly_chart(_fig(fig_t2,300),use_container_width=True)

    rpt_hdr("📋 Liste des Utilisateurs du Segment")
    sc_u=[c for c in ["utilisateur","direction","site","type_pc","annee_acquisition","anciennete","ram","processeur","priorite","recommandation"] if c in df_f2.columns]
    s_seg=st.text_input("🔍 Rechercher dans ce segment",key="search_seg")
    df_show=df_f2[sc_u].copy()
    if s_seg:
        m_s=pd.Series(False,index=df_show.index)
        for col in df_show.columns: m_s|=df_show[col].astype(str).str.contains(s_seg,case=False,na=False)
        df_show=df_show[m_s]
    st.dataframe(df_show,use_container_width=True,height=420)
    st.caption(f"👥 {len(df_show)} utilisateur(s)")
    st.download_button("⬇️ Exporter Segment Excel",data=export_excel(df_show,"Segment"),file_name=f"segment_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD HTML INTERACTIF
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📊 Dashboard HTML":
    if df_pc_raw is None and df_imp is None:
        st.warning("Chargez d'abord vos fichiers inventaire."); st.stop()

    df_d = df_pc_f if df_pc_f is not None else pd.DataFrame()
    total_pc   = len(df_d)
    urgent_n   = int(df_d[df_d["priorite"]=="🔴 Urgent"].shape[0])          if "priorite" in df_d.columns else 0
    prio_n     = int(df_d[df_d["priorite"]=="🟠 Prioritaire"].shape[0])      if "priorite" in df_d.columns else 0
    ok_n2      = int(df_d[df_d["priorite"]=="🟢 Opérationnel"].shape[0])     if "priorite" in df_d.columns else 0
    bureau_n   = int(df_d[df_d["type_pc"].str.upper()=="BUREAU"].shape[0])   if "type_pc"  in df_d.columns else 0
    portable_n = int(df_d[df_d["type_pc"].str.upper()=="PORTABLE"].shape[0]) if "type_pc"  in df_d.columns else 0
    chrom_nn   = len(df_chrom) if df_chrom is not None else 0
    imp_n2     = len(df_imp_f) if df_imp_f is not None else 0
    avg_age_d  = round(df_d["anciennete"].mean(),1) if "anciennete" in df_d.columns and df_d["anciennete"].notna().any() else 0
    old_n2     = int(df_d[df_d["categorie_age"]=="Plus de 10 ans"].shape[0]) if "categorie_age" in df_d.columns else 0
    nb_dir_d   = df_d["direction"].nunique() if "direction" in df_d.columns else 0

    dir_rows_html=""
    if "direction" in df_d.columns and "priorite" in df_d.columns:
        agg_d=df_d.groupby("direction").agg(Total=("utilisateur","count")).reset_index()
        urg_d=df_d[df_d["priorite"]=="🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
        agg_d=agg_d.merge(urg_d,on="direction",how="left").fillna(0); agg_d["Urgents"]=agg_d["Urgents"].astype(int)
        for _,r2 in agg_d.sort_values("Total",ascending=False).head(14).iterrows():
            pct=round(r2["Total"]/total_pc*100) if total_pc else 0
            uc="color:#C8102E;font-weight:700;" if r2["Urgents"]>0 else "color:#27AE60;"
            dir_rows_html+=f"<tr><td style='padding:9px 12px;font-weight:600;color:#1B2A3B;'>{r2['direction']}</td><td style='padding:9px 12px;text-align:center;'>{int(r2['Total'])}</td><td style='padding:9px 12px;'><div style='background:#F2F2F2;border-radius:4px;height:8px;'><div style='background:#C8102E;width:{pct}%;height:100%;border-radius:4px;'></div></div><span style='font-size:11px;color:#64748B;'>{pct}%</span></td><td style='padding:9px 12px;text-align:center;{uc}'>{int(r2['Urgents'])}</td></tr>"

    proc_html=""
    if "processeur" in df_d.columns:
        df_pr=df_d.copy(); df_pr["pc"]=df_pr["processeur"].apply(cat_proc)
        cnt_pr=df_pr["pc"].value_counts().head(6); max_p=cnt_pr.max() if len(cnt_pr)>0 else 1
        clr_pr={"Core i7/Ultra7":"#C8102E","Core i5/Ultra5":"#2980B9","Core i3":"#27AE60","Celeron/Dual":"#E67E22","Xeon":"#9B59B6","Autre":"#94A3B8"}
        for pn,pc2 in cnt_pr.items():
            pct_p=round(pc2/max_p*100); col_p=clr_pr.get(pn,"#94A3B8")
            proc_html+=f"<div style='margin-bottom:10px;'><div style='display:flex;justify-content:space-between;margin-bottom:3px;'><span style='font-size:13px;font-weight:600;color:#1B2A3B;'>{pn}</span><span style='font-size:13px;font-weight:700;color:{col_p};'>{pc2}</span></div><div style='background:#F2F2F2;border-radius:5px;height:9px;'><div style='background:{col_p};width:{pct_p}%;height:100%;border-radius:5px;'></div></div></div>"

    ram_html=""
    if "ram_go" in df_d.columns:
        rc2=df_d["ram_go"].dropna().value_counts().sort_index(); tr2=rc2.sum()
        clr_r2={4:"#E8415B",8:"#2980B9",16:"#27AE60",32:"#9B59B6"}
        for rv,rc in rc2.items():
            pct_r=round(rc/tr2*100) if tr2 else 0; col_r=clr_r2.get(int(rv),"#94A3B8")
            ram_html+=f"<div style='display:flex;align-items:center;gap:10px;margin-bottom:9px;'><div style='width:48px;font-size:13px;font-weight:700;color:{col_r};'>{int(rv)} GO</div><div style='flex:1;background:#F2F2F2;border-radius:5px;height:11px;'><div style='background:{col_r};width:{pct_r}%;height:100%;border-radius:5px;'></div></div><div style='width:60px;text-align:right;font-size:12px;font-weight:600;color:#64748B;'>{rc} ({pct_r}%)</div></div>"

    urg_html=""
    if urgent_n>0 and "priorite" in df_d.columns:
        for _,ur in df_d[df_d["priorite"]=="🔴 Urgent"].head(8).iterrows():
            un=val_str(ur.get("utilisateur")); ud=val_str(ur.get("direction")); ua=val_str(ur.get("anciennete")); um=val_str(ur.get("descriptions",""))[:30]
            urg_html+=f"<div style='display:flex;align-items:center;gap:12px;padding:9px 0;border-bottom:1px solid #F2F2F2;'><div style='width:34px;height:34px;border-radius:50%;background:#FEF2F2;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;'>🔴</div><div style='flex:1;min-width:0;'><div style='font-weight:700;font-size:13px;color:#1B2A3B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>{un}</div><div style='font-size:11px;color:#64748B;'>{ud} · {um}</div></div><div style='font-size:12px;font-weight:700;color:#C8102E;white-space:nowrap;'>{ua} ans</div></div>"

    html = f"""<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Segoe UI',system-ui,sans-serif;}}
body{{background:#F0F4F8;padding:0;}}
.grid{{display:grid;gap:14px;}}
.g2{{grid-template-columns:1fr 1fr;}}
.g3{{grid-template-columns:1fr 1fr 1fr;}}
.g4{{grid-template-columns:repeat(4,1fr);}}
.g23{{grid-template-columns:2fr 1.2fr;}}
.card{{background:white;border-radius:14px;padding:18px 20px;box-shadow:0 2px 10px rgba(0,0,0,.06);}}
.card-title{{font-size:12px;font-weight:700;color:#1B2A3B;text-transform:uppercase;letter-spacing:.5px;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #F2F2F2;}}
.kpi-card{{background:white;border-radius:12px;padding:16px 18px;box-shadow:0 2px 10px rgba(0,0,0,.06);border-left:5px solid #C8102E;}}
.kpi-big{{font-size:34px;font-weight:900;line-height:1;}}
.kpi-lbl{{font-size:10px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:.5px;margin-top:4px;}}
.kpi-sub{{font-size:11px;color:#94A3B8;margin-top:2px;}}
.hdr{{background:linear-gradient(135deg,#1B2A3B,#2D3E50);border-radius:14px;padding:20px 24px;color:white;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;}}
.hdr-t{{font-size:20px;font-weight:800;}}
.hdr-s{{font-size:11px;color:#94A3B8;margin-top:3px;}}
.pill{{display:inline-block;background:rgba(200,16,46,.15);border:1px solid rgba(200,16,46,.3);color:#F5A0AE;padding:3px 10px;border-radius:20px;font-size:10px;font-weight:600;margin-top:10px;margin-right:4px;}}
table{{width:100%;border-collapse:collapse;}}
th{{background:#1B2A3B;color:white;padding:8px 12px;font-size:10px;font-weight:700;text-align:left;text-transform:uppercase;letter-spacing:.5px;}}
tr:nth-child(even){{background:#FAFBFC;}}
tr:hover{{background:#FFF5F7;}}
.stat-box{{display:flex;justify-content:space-between;align-items:center;padding:9px 12px;background:#F8FAFC;border-radius:8px;border-left:4px solid #C8102E;margin-bottom:8px;}}
.foot{{background:#1B2A3B;border-radius:10px;padding:12px 20px;display:flex;justify-content:space-between;align-items:center;margin-top:14px;}}
@media(max-width:900px){{.g4,.g3,.g2,.g23{{grid-template-columns:1fr;}}}}
</style></head><body>

<div class="hdr">
  <div>
    <div class="hdr-t">DSI Amendis — Inventaire Informatique Tetouan 2026</div>
    <div class="hdr-s">Direction des Systemes d'Information · Tableau de bord</div>
    <div><span class="pill">📅 {datetime.now().strftime("%d/%m/%Y")}</span><span class="pill">🏙️ Tetouan</span><span class="pill">🏢 {nb_dir_d} Directions</span><span class="pill">🖥️ {total_pc} PC</span></div>
  </div>
  <div style="font-size:38px;opacity:.25;">🖥️</div>
</div>

<div class="grid g4" style="margin-bottom:14px;">
  <div class="kpi-card" style="border-left-color:#C8102E;"><div class="kpi-big" style="color:#C8102E;">{total_pc}</div><div class="kpi-lbl">Total PC</div><div class="kpi-sub">Bureau {bureau_n} · Portable {portable_n}</div></div>
  <div class="kpi-card" style="border-left-color:#E67E22;"><div class="kpi-big" style="color:#E67E22;">{urgent_n}</div><div class="kpi-lbl">Urgents a Remplacer</div><div class="kpi-sub">{round(urgent_n/total_pc*100) if total_pc else 0}% du parc</div></div>
  <div class="kpi-card" style="border-left-color:#27AE60;"><div class="kpi-big" style="color:#27AE60;">{ok_n2}</div><div class="kpi-lbl">Operationnels</div><div class="kpi-sub">{round(ok_n2/total_pc*100) if total_pc else 0}% du parc</div></div>
  <div class="kpi-card" style="border-left-color:#9B59B6;"><div class="kpi-big" style="color:#9B59B6;">{imp_n2}</div><div class="kpi-lbl">Imprimantes</div><div class="kpi-sub">Chromebooks: {chrom_nn}</div></div>
</div>

<div class="grid g23" style="margin-bottom:14px;">
  <div class="card">
    <div class="card-title">Etat du Parc PC</div>
    <div style="display:flex;gap:12px;justify-content:space-around;margin-bottom:16px;">
      <div style="text-align:center;"><div style="width:60px;height:60px;border-radius:50%;background:#C8102E;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;color:white;margin:0 auto 4px;">{urgent_n}</div><div style="font-size:10px;color:#64748B;font-weight:600;">URGENTS</div></div>
      <div style="text-align:center;"><div style="width:60px;height:60px;border-radius:50%;background:#E67E22;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;color:white;margin:0 auto 4px;">{prio_n}</div><div style="font-size:10px;color:#64748B;font-weight:600;">PRIORITAIRES</div></div>
      <div style="text-align:center;"><div style="width:60px;height:60px;border-radius:50%;background:#27AE60;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800;color:white;margin:0 auto 4px;">{ok_n2}</div><div style="font-size:10px;color:#64748B;font-weight:600;">OPERATIONNELS</div></div>
    </div>
    <div style="display:flex;justify-content:space-between;font-size:12px;color:#64748B;margin-bottom:5px;"><span>Taux urgents</span><span style="font-weight:700;color:#C8102E;">{round(urgent_n/total_pc*100) if total_pc else 0}%</span></div>
    <div style="background:#F2F2F2;border-radius:5px;height:8px;margin-bottom:10px;"><div style="background:#C8102E;width:{round(urgent_n/total_pc*100) if total_pc else 0}%;height:100%;border-radius:5px;"></div></div>
    <div style="display:flex;justify-content:space-between;font-size:12px;color:#64748B;margin-bottom:5px;"><span>Taux operationnels</span><span style="font-weight:700;color:#27AE60;">{round(ok_n2/total_pc*100) if total_pc else 0}%</span></div>
    <div style="background:#F2F2F2;border-radius:5px;height:8px;"><div style="background:#27AE60;width:{round(ok_n2/total_pc*100) if total_pc else 0}%;height:100%;border-radius:5px;"></div></div>
  </div>
  <div class="card">
    <div class="card-title">Chiffres Cles</div>
    <div class="stat-box" style="border-left-color:#C8102E;"><span style="font-size:13px;color:#64748B;">Age moyen</span><span style="font-size:18px;font-weight:800;color:#C8102E;">{avg_age_d} ans</span></div>
    <div class="stat-box" style="border-left-color:#E67E22;"><span style="font-size:13px;color:#64748B;">PC sup 10 ans</span><span style="font-size:18px;font-weight:800;color:#E67E22;">{old_n2} ({round(old_n2/total_pc*100) if total_pc else 0}%)</span></div>
    <div class="stat-box" style="border-left-color:#2980B9;"><span style="font-size:13px;color:#64748B;">Directions</span><span style="font-size:18px;font-weight:800;color:#2980B9;">{nb_dir_d}</span></div>
    <div class="stat-box" style="border-left-color:#27AE60;"><span style="font-size:13px;color:#64748B;">Chromebooks</span><span style="font-size:18px;font-weight:800;color:#27AE60;">{chrom_nn}</span></div>
  </div>
</div>

<div class="grid g3" style="margin-bottom:14px;">
  <div class="card"><div class="card-title">PC par Direction</div><div style="overflow-y:auto;max-height:320px;"><table><thead><tr><th>Direction</th><th style="text-align:center;">PC</th><th>Part</th><th style="text-align:center;">Urgents</th></tr></thead><tbody>{dir_rows_html}</tbody></table></div></div>
  <div class="card"><div class="card-title">Processeurs</div>{proc_html if proc_html else '<p style="color:#94A3B8;font-size:13px;">Donnees non disponibles</p>'}</div>
  <div class="card"><div class="card-title">Distribution RAM</div>{ram_html if ram_html else '<p style="color:#94A3B8;font-size:13px;">Donnees non disponibles</p>'}</div>
</div>

{"<div class='card' style='margin-bottom:14px;'><div class='card-title'>PC Urgents a Remplacer (Top "+str(min(urgent_n,8))+")</div>"+urg_html+"</div>" if urgent_n>0 else ""}

<div class="foot"><span style="font-size:11px;color:#64748B;">DSI — Amendis Tetouan · Inventaire 2026</span><span style="font-size:11px;color:#C8102E;font-weight:600;">Genere le {datetime.now().strftime("%d/%m/%Y a %H:%M")}</span></div>
</body></html>"""

    st.components.v1.html(html, height=1900, scrolling=True)
