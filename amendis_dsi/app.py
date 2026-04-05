"""
app.py — Application DSI Amendis Tétouan
Tableau de bord d'inventaire informatique (PC + Imprimantes)
Thème: Blanc & Rouge Amendis
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
from datetime import datetime

from data_loader import load_inventaire, load_imprimantes
from charts import (
    chart_pc_par_direction, chart_type_pc, chart_age_distribution,
    chart_ram_distribution, chart_processeur, chart_priorite_remplacement,
    chart_heatmap_direction_type, chart_acquisitions_timeline,
    chart_imp_par_direction, chart_imp_type, chart_imp_modele,
    COLORS,
)

# ─── Config Page ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DSI Amendis — Inventaire Tétouan",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS Amendis (Rouge & Blanc) ──────────────────────────────────────────────
st.markdown("""
<style>
/* ── Fonts & Reset ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

* { font-family: 'Inter', 'Segoe UI', Arial, sans-serif; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1A1A1A 0%, #2D0A10 100%);
    border-right: 3px solid #C8102E;
}
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stTextInput label { color: #F5A0AE !important; font-weight: 600; }
[data-testid="stSidebar"] .stButton button {
    background: #C8102E !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    width: 100%;
    padding: 10px !important;
    transition: all 0.2s;
}
[data-testid="stSidebar"] .stButton button:hover {
    background: #E8415B !important;
    transform: translateY(-1px);
}

/* ── Main area ── */
.main .block-container { padding: 1.5rem 2rem; max-width: 1400px; }

/* ── Header ── */
.amendis-header {
    background: linear-gradient(135deg, #C8102E 0%, #8B0A1E 100%);
    color: white;
    padding: 24px 32px;
    border-radius: 16px;
    margin-bottom: 24px;
    box-shadow: 0 4px 20px rgba(200,16,46,0.3);
}
.amendis-header h1 { margin: 0; font-size: 28px; font-weight: 700; }
.amendis-header p { margin: 4px 0 0; opacity: 0.85; font-size: 14px; }

/* ── KPI Cards ── */
.kpi-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    border-left: 5px solid #C8102E;
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.12); }
.kpi-number { font-size: 36px; font-weight: 700; color: #C8102E; margin: 0; }
.kpi-label { font-size: 13px; color: #6B6B6B; font-weight: 500; margin: 4px 0 0; }

/* ── Section headers ── */
.section-title {
    font-size: 20px;
    font-weight: 700;
    color: #C8102E;
    border-bottom: 2px solid #C8102E;
    padding-bottom: 8px;
    margin: 32px 0 20px;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] {
    background: #F2F2F2;
    border-radius: 8px 8px 0 0;
    font-weight: 600;
    color: #2D2D2D;
    padding: 10px 20px;
}
.stTabs [aria-selected="true"] {
    background: #C8102E !important;
    color: white !important;
}

/* ── DataFrames ── */
.dataframe thead tr th {
    background-color: #C8102E !important;
    color: white !important;
    font-weight: 700;
}

/* ── Buttons ── */
.stButton > button {
    background: #C8102E;
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    padding: 8px 20px;
    transition: all 0.2s;
}
.stButton > button:hover {
    background: #E8415B;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(200,16,46,0.4);
}

/* ── Profile card ── */
.profile-card {
    background: linear-gradient(135deg, #FFF8F9 0%, #FFECEF 100%);
    border: 2px solid #C8102E;
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 16px;
}
.profile-name { font-size: 22px; font-weight: 700; color: #C8102E; }
.profile-detail { font-size: 14px; color: #2D2D2D; margin: 4px 0; }
.profile-badge {
    display: inline-block;
    background: #C8102E;
    color: white;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    margin: 4px 4px 0 0;
}
.profile-badge-green {
    display: inline-block;
    background: #2ECC71;
    color: white;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
    margin: 4px 4px 0 0;
}

/* ── Alert boxes ── */
.alert-urgent {
    background: #FFF0F2;
    border-left: 4px solid #C8102E;
    padding: 12px 16px;
    border-radius: 0 8px 8px 0;
    margin: 8px 0;
}
.alert-ok {
    background: #F0FFF4;
    border-left: 4px solid #2ECC71;
    padding: 12px 16px;
    border-radius: 0 8px 8px 0;
    margin: 8px 0;
}

/* ── Info boxes ── */
div[data-testid="stInfo"] { border-radius: 10px; }
div[data-testid="stSuccess"] { border-radius: 10px; }
div[data-testid="stWarning"] { border-radius: 10px; }
div[data-testid="stError"] { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)


# ─── State ────────────────────────────────────────────────────────────────────
if "df_inv" not in st.session_state:
    st.session_state.df_inv = None
if "df_imp" not in st.session_state:
    st.session_state.df_imp = None


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding: 16px 0 24px;">
        <div style="font-size:40px;">🏢</div>
        <div style="font-size:18px; font-weight:700; color:white;">AMENDIS</div>
        <div style="font-size:11px; color:#F5A0AE; letter-spacing:2px;">DIRECTION DES SYSTÈMES D'INFORMATION</div>
        <div style="font-size:10px; color:#888; margin-top:4px;">SITE TÉTOUAN — 2026</div>
    </div>
    <hr style="border-color:#C8102E; margin:0 0 20px;">
    """, unsafe_allow_html=True)

    st.markdown("### 📁 Fichiers d'inventaire")

    inv_file = st.file_uploader(
        "Inventaire PC (xlsx)",
        type=["xlsx", "xls"],
        key="inv_upload",
        help="Fichier inventaire PC & PC Portable Tétouan 2026",
    )
    imp_file = st.file_uploader(
        "Inventaire Imprimantes (xlsx)",
        type=["xlsx", "xls"],
        key="imp_upload",
        help="Fichier inventaire imprimantes Tétouan",
    )

    if st.button("⚡ Charger & Analyser", use_container_width=True):
        if inv_file:
            with st.spinner("Chargement inventaire PC..."):
                try:
                    st.session_state.df_inv = load_inventaire(inv_file)
                    st.success(f"✅ PC: {len(st.session_state.df_inv)} enregistrements")
                except Exception as e:
                    st.error(f"Erreur PC: {e}")
        if imp_file:
            with st.spinner("Chargement imprimantes..."):
                try:
                    st.session_state.df_imp = load_imprimantes(imp_file)
                    st.success(f"✅ Imprimantes: {len(st.session_state.df_imp)} enregistrements")
                except Exception as e:
                    st.error(f"Erreur imprimantes: {e}")

    st.markdown("---")
    st.markdown("### 🔍 Navigation")
    page = st.radio(
        "",
        ["Tableau de bord", "Rapport PC", "Rapport Imprimantes",
         "Rapport General", "Profil Utilisateur"],
        label_visibility="collapsed",
    )

    # Filtres dynamiques si données chargées
    st.markdown("---")
    st.markdown("### 🎛️ Filtres")
    df_inv = st.session_state.df_inv
    df_imp = st.session_state.df_imp

    filter_direction = None
    if df_inv is not None and "direction" in df_inv.columns:
        dirs = sorted(df_inv["direction"].dropna().unique().tolist())
        filter_direction = st.multiselect("Direction(s)", dirs, default=[], key="f_dir")

    filter_site = None
    if df_inv is not None and "site" in df_inv.columns:
        sites = sorted(df_inv["site"].dropna().unique().tolist())
        filter_site = st.multiselect("Site(s)", sites, default=[], key="f_site")

    filter_type = None
    if df_inv is not None and "type_pc" in df_inv.columns:
        types = sorted(df_inv["type_pc"].dropna().unique().tolist())
        filter_type = st.multiselect("Type PC", types, default=[], key="f_type")

    st.markdown("---")
    st.markdown(f"<div style='font-size:10px;color:#666;text-align:center;'>v2.0 • {datetime.now().strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)


# ─── Apply filters ────────────────────────────────────────────────────────────
def apply_filters(df):
    if df is None:
        return df
    d = df.copy()
    if filter_direction:
        d = d[d.get("direction", pd.Series()).isin(filter_direction)]
    if filter_site:
        d = d[d.get("site", pd.Series()).isin(filter_site)]
    if filter_type and "type_pc" in d.columns:
        d = d[d["type_pc"].isin(filter_type)]
    return d


# ─── Helper: KPI metric ───────────────────────────────────────────────────────
def kpi(label, value, icon=""):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-number">{icon} {value}</div>
        <div class="kpi-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)


def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


# ─── Export Excel helper ──────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame, sheet_name="Données") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        hdr_fmt = wb.add_format({
            "bg_color": "#C8102E", "font_color": "#FFFFFF",
            "bold": True, "border": 1, "align": "center",
        })
        for ci, col in enumerate(df.columns):
            ws.write(0, ci, col, hdr_fmt)
            ws.set_column(ci, ci, max(18, len(str(col)) + 4))
    buf.seek(0)
    return buf.read()


# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="amendis-header">
    <h1>🖥️ DSI Amendis — Inventaire Informatique Tétouan 2026</h1>
    <p>Direction des Systèmes d'Information · Tableau de bord analytique · Site Tétouan</p>
</div>
""", unsafe_allow_html=True)

# Check data
df_inv_f = apply_filters(st.session_state.df_inv)
df_imp_f = apply_filters(st.session_state.df_imp) if st.session_state.df_imp is not None else None

if st.session_state.df_inv is None and st.session_state.df_imp is None:
    st.info("👈 Veuillez uploader vos fichiers Excel dans la barre latérale puis cliquer sur **⚡ Charger & Analyser**")
    st.markdown("""
    <div style="text-align:center; padding:60px 20px;">
        <div style="font-size:80px;">📂</div>
        <h3 style="color:#C8102E;">Chargez vos fichiers inventaire</h3>
        <p style="color:#6B6B6B; max-width:500px; margin:auto;">
            Uploadez le fichier <strong>Inventaire PC</strong> et/ou <strong>Imprimantes</strong> 
            pour accéder aux rapports statistiques, analyses et profils utilisateurs.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Tableau de bord
# ═══════════════════════════════════════════════════════════════════════════════
if page == "Tableau de bord":
    section("📌 Indicateurs Clés")

    cols = st.columns(5)
    df_inv = df_inv_f
    df_imp = df_imp_f

    with cols[0]:
        kpi("Total PC", len(df_inv) if df_inv is not None else "—", "")
    with cols[1]:
        if df_inv is not None and "type_pc" in df_inv.columns:
            bureau = df_inv[df_inv["type_pc"].str.upper().str.strip() == "BUREAU"].shape[0]
            kpi("PC Bureau", bureau, "")
        else:
            kpi("PC Bureau", "—")
    with cols[2]:
        if df_inv is not None and "type_pc" in df_inv.columns:
            port = df_inv[df_inv["type_pc"].str.upper().str.strip() == "PORTABLE"].shape[0]
            kpi("PC Portables", port, "")
        else:
            kpi("PC Portables", "—")
    with cols[3]:
        kpi("Total Imprimantes", len(df_imp) if df_imp is not None else "—", "")
    with cols[4]:
        if df_inv is not None and "priorite" in df_inv.columns:
            urgent = df_inv[df_inv["priorite"] == "🔴 Urgent"].shape[0]
            kpi("PC à remplacer", urgent, "")
        else:
            kpi("Directions", df_inv["direction"].nunique() if df_inv is not None and "direction" in df_inv.columns else "—", "")

    st.markdown("<br>", unsafe_allow_html=True)

    # Row 1
    if df_inv is not None:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = chart_pc_par_direction(df_inv)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_type_pc(df_inv)
            st.plotly_chart(fig, use_container_width=True)

        # Row 2
        c1, c2, c3 = st.columns(3)
        with c1:
            fig = chart_age_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_ram_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c3:
            fig = chart_processeur(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

    if df_imp is not None:
        c1, c2 = st.columns(2)
        with c1:
            fig = chart_imp_par_direction(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_imp_type(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport PC
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport PC":
    df_inv = df_inv_f
    if df_inv is None:
        st.warning("Veuillez charger le fichier inventaire PC.")
        st.stop()

    section("Rapport Statistique — Parc PC")

    # Sub-report selector
    report_type = st.radio(
        "Type de rapport",
        ["Par Direction", "Par Site", "Rapport Global PC"],
        horizontal=True,
        label_visibility="collapsed",
    )

    tabs = st.tabs(["Statistiques", "Recommandations", "Donnees detaillees"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            if report_type == "Par Direction":
                fig = chart_pc_par_direction(df_inv)
            elif report_type == "Par Site":
                if "site" in df_inv.columns:
                    cnt = df_inv.groupby("site").size().reset_index(name="count").sort_values("count", ascending=True)
                    fig = px.bar(cnt, x="count", y="site", orientation="h",
                                  title="📍 Parc PC par Site",
                                  labels={"count": "Nombre", "site": "Site"},
                                  color_discrete_sequence=["#C8102E"])
                    fig.update_layout(height=400, plot_bgcolor="white")
                else:
                    fig = chart_pc_par_direction(df_inv)
            else:
                fig = chart_pc_par_direction(df_inv)
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig = chart_age_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig = chart_ram_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_processeur(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        fig = chart_acquisitions_timeline(df_inv)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        fig = chart_heatmap_direction_type(df_inv)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    with tabs[1]:
        section("Recommandations pour la Prise de Decision")
        fig = chart_priorite_remplacement(df_inv)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        if "priorite" in df_inv.columns:
            for prio, color, bg in [
                ("🔴 Urgent", "#C8102E", "#FFF0F2"),
                ("🟠 Prioritaire", "#E67E22", "#FFF8F0"),
                ("🟡 À surveiller", "#F1C40F", "#FEFEF0"),
                ("🟢 Opérationnel", "#2ECC71", "#F0FFF4"),
            ]:
                sub = df_inv[df_inv["priorite"] == prio]
                if len(sub) > 0:
                    with st.expander(f"{prio} — {len(sub)} PC", expanded=(prio == "🔴 Urgent")):
                        show_cols = [c for c in ["utilisateur", "direction", "site", "type_pc",
                                                   "descriptions", "annee_acquisition", "ram",
                                                   "processeur", "raisons", "recommandation"] if c in sub.columns]
                        st.dataframe(sub[show_cols], use_container_width=True)

    with tabs[2]:
        section("Donnees Inventaire PC")
        show_cols = [c for c in ["utilisateur", "direction", "site", "type_pc", "n_inventaire",
                                   "descriptions", "annee_acquisition", "anciennete",
                                   "ram", "processeur", "hdd", "priorite"] if c in df_inv.columns]
        search = st.text_input("🔍 Rechercher un utilisateur...", key="search_inv")
        df_show = df_inv.copy()
        if search:
            mask = df_show.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
            df_show = df_show[mask]
        st.dataframe(df_show[show_cols], use_container_width=True, height=500)

        col1, col2 = st.columns(2)
        with col1:
            xlsx_bytes = export_excel(df_inv[show_cols], "Inventaire PC")
            st.download_button(
                "⬇️ Télécharger Excel",
                data=xlsx_bytes,
                file_name=f"inventaire_pc_tetouan_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col2:
            csv = df_inv[show_cols].to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Télécharger CSV", data=csv,
                               file_name=f"inventaire_pc_{datetime.now().strftime('%Y%m%d')}.csv")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport Imprimantes
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport Imprimantes":
    df_imp = df_imp_f
    if df_imp is None:
        st.warning("Veuillez charger le fichier inventaire imprimantes.")
        st.stop()

    section("Rapport Statistique — Parc Imprimantes")

    report_type = st.radio(
        "Type de rapport",
        ["Par Direction", "Par Site", "Rapport Global Imprimantes"],
        horizontal=True,
        label_visibility="collapsed",
    )

    tabs = st.tabs(["Statistiques", "Donnees detaillees"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            if report_type == "Par Direction" and "direction" in df_imp.columns:
                fig = chart_imp_par_direction(df_imp)
            elif report_type == "Par Site" and "site" in df_imp.columns:
                cnt = df_imp.groupby("site").size().reset_index(name="count").sort_values("count", ascending=True)
                fig = px.bar(cnt, x="count", y="site", orientation="h",
                              title="📍 Imprimantes par Site",
                              labels={"count": "Nombre", "site": "Site"},
                              color_discrete_sequence=["#C8102E"])
                fig.update_layout(height=420, plot_bgcolor="white")
            else:
                fig = chart_imp_par_direction(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig = chart_imp_type(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        fig = chart_imp_modele(df_imp)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        # Acquisitions timeline
        if "annee_acquisition" in df_imp.columns:
            cnt = df_imp["annee_acquisition"].dropna().value_counts().sort_index().reset_index()
            cnt.columns = ["annee", "count"]
            cnt["annee"] = cnt["annee"].astype(int)
            fig = px.bar(cnt, x="annee", y="count",
                          title="📅 Acquisitions Imprimantes par Année",
                          labels={"annee": "Année", "count": "Nombre"},
                          color_discrete_sequence=["#C8102E"])
            fig.update_layout(plot_bgcolor="white", height=360)
            st.plotly_chart(fig, use_container_width=True)

    with tabs[1]:
        show_cols = [c for c in ["utilisateur", "direction", "site", "n_inventaire",
                                   "modele", "type_imp", "annee_acquisition", "n_serie"] if c in df_imp.columns]
        search = st.text_input("🔍 Rechercher...", key="search_imp")
        df_show = df_imp.copy()
        if search:
            mask = df_show.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
            df_show = df_show[mask]
        st.dataframe(df_show[show_cols], use_container_width=True, height=500)

        xlsx_bytes = export_excel(df_imp[show_cols], "Imprimantes")
        st.download_button(
            "⬇️ Télécharger Excel",
            data=xlsx_bytes,
            file_name=f"inventaire_imprimantes_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport Général Global
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport General":
    section("Rapport General Global — Site Tetouan 2026")

    df_inv = df_inv_f
    df_imp = df_imp_f

    # ── Vue d'ensemble ─────────────────────────────────────────────────────
    st.markdown("#### Vue d'ensemble du Parc Informatique")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    metrics = []
    if df_inv is not None:
        metrics += [
            ("Total PC", len(df_inv), ""),
            ("PC Bureau", len(df_inv[df_inv.get("type_pc", pd.Series(dtype=str)).str.strip() == "BUREAU"]) if "type_pc" in df_inv.columns else 0, ""),
            ("PC Portables", len(df_inv[df_inv.get("type_pc", pd.Series(dtype=str)).str.strip() == "PORTABLE"]) if "type_pc" in df_inv.columns else 0, ""),
        ]
    if df_imp is not None:
        metrics += [("Total Imprimantes", len(df_imp), "")]
    if df_inv is not None and "priorite" in df_inv.columns:
        urgent = len(df_inv[df_inv["priorite"] == "🔴 Urgent"])
        ok = len(df_inv[df_inv["priorite"] == "🟢 Opérationnel"])
        metrics += [("PC Urgents", urgent, ""), ("PC Opérationnels", ok, "")]

    cols = st.columns(len(metrics)) if metrics else []
    for i, (label, val, icon) in enumerate(metrics):
        with cols[i]:
            kpi(label, val, icon)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Analyses globales ─────────────────────────────────────────────────
    if df_inv is not None:
        st.markdown("#### Analyse PC")
        c1, c2, c3 = st.columns(3)
        with c1:
            fig = chart_pc_par_direction(df_inv)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_type_pc(df_inv)
            st.plotly_chart(fig, use_container_width=True)
        with c3:
            fig = chart_age_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig = chart_ram_distribution(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_processeur(df_inv)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        fig = chart_priorite_remplacement(df_inv)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        fig = chart_acquisitions_timeline(df_inv)
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    if df_imp is not None:
        st.markdown("#### Analyse Imprimantes")
        c1, c2, c3 = st.columns(3)
        with c1:
            fig = chart_imp_par_direction(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_imp_type(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
        with c3:
            fig = chart_imp_modele(df_imp)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

    # ── Synthèse par direction ────────────────────────────────────────────
    if df_inv is not None and "direction" in df_inv.columns:
        section("Synthese par Direction")
        summary_rows = []
        for dir_name, grp in df_inv.groupby("direction"):
            row = {"Direction": dir_name, "Total PC": len(grp)}
            if "type_pc" in grp.columns:
                row["Bureau"] = grp[grp["type_pc"].str.strip() == "BUREAU"].shape[0]
                row["Portable"] = grp[grp["type_pc"].str.strip() == "PORTABLE"].shape[0]
            if "priorite" in grp.columns:
                row["🔴 Urgent"] = grp[grp["priorite"] == "🔴 Urgent"].shape[0]
                row["🟢 OK"] = grp[grp["priorite"] == "🟢 Opérationnel"].shape[0]
            if "ram_go" in grp.columns:
                row["RAM Moy. (GO)"] = round(grp["ram_go"].mean(), 1)
            if df_imp is not None and "direction" in df_imp.columns:
                row["Imprimantes"] = df_imp[df_imp["direction"] == dir_name].shape[0]
            summary_rows.append(row)

        df_summary = pd.DataFrame(summary_rows).sort_values("Total PC", ascending=False)
        st.dataframe(df_summary, use_container_width=True)

        xlsx_bytes = export_excel(df_summary, "Synthèse Globale")
        st.download_button(
            "⬇️ Exporter Synthèse Excel",
            data=xlsx_bytes,
            file_name=f"rapport_global_tetouan_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Profil Utilisateur
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Profil Utilisateur":
    section("Profil Technique Utilisateur — Fiche AMENDIS")

    df_inv = st.session_state.df_inv
    df_imp = st.session_state.df_imp

    if df_inv is None:
        st.warning("Veuillez charger le fichier inventaire PC.")
        st.stop()

    st.markdown(
        '<div style="background:#FFF8F9; border:1px solid #F5A0AE; border-radius:10px;'
        ' padding:16px; margin-bottom:20px;">'
        '<b>Mode Profil</b> — Saisissez le nom d\'un utilisateur pour générer sa fiche technique.'
        ' Cette fiche inclut son PC, ses caractéristiques techniques, les recommandations DSI'
        ' et ses imprimantes associées.</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([3, 1])
    with col1:
        search_name = st.text_input(
            "Nom de l'utilisateur",
            placeholder="Ex: Mohamed Benali, Fatima El Idrissi...",
            key="profile_search",
        )
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("Rechercher", use_container_width=True)

    if not search_name:
        if "utilisateur" in df_inv.columns:
            users_list = sorted(df_inv["utilisateur"].dropna().unique().tolist())
            st.markdown("**Utilisateurs disponibles dans l'inventaire:**")
            n_cols = 4
            rows_u = [users_list[i:i+n_cols] for i in range(0, min(len(users_list), 80), n_cols)]
            for row_u in rows_u:
                cols_u = st.columns(n_cols)
                for i, u in enumerate(row_u):
                    with cols_u[i]:
                        st.markdown(f"<span style='font-size:13px;'>- {u}</span>", unsafe_allow_html=True)
            if len(users_list) > 80:
                st.caption(f"... et {len(users_list) - 80} autres utilisateurs")
        st.stop()

    if "utilisateur" not in df_inv.columns:
        st.error("Colonne Utilisateur non trouvee.")
        st.stop()

    mask = df_inv["utilisateur"].str.contains(search_name, case=False, na=False)
    results = df_inv[mask]

    if len(results) == 0:
        st.error(f"Aucun utilisateur trouve pour: {search_name}")
        st.info("Astuce: Essayez avec juste le nom de famille.")
        st.stop()

    if len(results) > 1:
        st.info(f"{len(results)} correspondance(s) trouvee(s). Selectionnez l'utilisateur:")
        user_options = results["utilisateur"].unique().tolist()
        selected_user = st.selectbox("Utilisateur", user_options)
        user_data = results[results["utilisateur"] == selected_user]
    else:
        user_data = results
        selected_user = results.iloc[0]["utilisateur"]

    row = user_data.iloc[0]

    # ── Valeurs nettoyées ─────────────────────────────────────────────────
    prio        = str(row.get("priorite", "Operationnel"))
    raisons     = str(row.get("raisons", ""))
    rec         = str(row.get("recommandation", ""))
    waterp      = str(row.get("waterp", "")).upper()
    sensible    = str(row.get("poste_sensible", "")).upper()
    direction   = str(row.get("direction", "N/A"))
    site        = str(row.get("site", "N/A"))
    type_pc     = str(row.get("type_pc", "N/A"))

    prio_color = "#C8102E"
    if "Prioritaire" in prio:
        prio_color = "#E67E22"
    elif "surveiller" in prio:
        prio_color = "#F1C40F"
    elif "Operationnel" in prio or "Opérationnel" in prio:
        prio_color = "#2ECC71"

    # ── Badges HTML (construits séparément) ──────────────────────────────
    badges_html = (
        f'<span class="profile-badge">{direction}</span>'
        f'<span class="profile-badge">{site}</span>'
        f'<span class="profile-badge">{type_pc}</span>'
    )
    if sensible == "OUI":
        badges_html += '<span class="profile-badge">Poste Sensible</span>'
    if waterp == "OUI":
        badges_html += '<span class="profile-badge">WATERP</span>'

    # ── Statut DSI ───────────────────────────────────────────────────────
    raisons_html = ""
    if raisons and raisons not in ("nan", ""):
        raisons_html = f'<br><span style="font-size:12px;color:#6B6B6B;">Raisons: {raisons}</span>'

    statut_html = (
        f'<div style="margin-top:16px;padding:12px;background:white;border-radius:10px;'
        f'border-left:4px solid {prio_color};">'
        f'<b>Statut DSI:</b> {prio} &nbsp;|&nbsp; <b>Action:</b> {rec}'
        f'{raisons_html}</div>'
    )

    profile_html = (
        '<div class="profile-card">'
        f'<div class="profile-name">{selected_user}</div>'
        f'<div style="margin:12px 0 8px;">{badges_html}</div>'
        f'{statut_html}'
        '</div>'
    )
    st.markdown(profile_html, unsafe_allow_html=True)

    # ── Fiche Technique PC ────────────────────────────────────────────────
    st.markdown("#### Fiche Technique PC")
    c1, c2, c3 = st.columns(3)

    def val_str(v):
        return str(v) if pd.notna(v) and str(v) not in ("nan", "NaT", "None") else "N/A"

    pc_info = {
        "N Inventaire":      val_str(row.get("n_inventaire")),
        "Modele":            val_str(row.get("descriptions")),
        "N Serie":           val_str(row.get("n_serie")),
        "Annee acquisition": val_str(row.get("annee_acquisition")),
        "Anciennete":        val_str(row.get("anciennete")) + " ans",
        "Categorie age":     val_str(row.get("categorie_age")),
        "RAM":               val_str(row.get("ram")),
        "Processeur":        val_str(row.get("processeur")),
        "Disque dur":        val_str(row.get("hdd")),
    }

    items = list(pc_info.items())
    cols3 = [c1, c2, c3]
    for i, (label, valeur) in enumerate(items):
        card_html = (
            '<div style="background:#F9F9F9;border:1px solid #E0E0E0;border-radius:8px;'
            'padding:14px;margin:6px 0;">'
            f'<div style="font-size:12px;color:#888;font-weight:500;">{label}</div>'
            f'<div style="font-size:17px;font-weight:700;color:#C8102E;margin-top:4px;">{valeur}</div>'
            '</div>'
        )
        with cols3[i % 3]:
            st.markdown(card_html, unsafe_allow_html=True)

    # ── Recommandation PC ─────────────────────────────────────────────────
    st.markdown("#### Recommandation DSI")

    if sensible == "OUI":
        pc_rec    = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_rec = "Intel Core i7, 16 GO RAM, 512 GO SSD, Windows 11 Pro"
        reason_rec = "Poste sensible — securite renforcee requise"
    elif type_pc.strip().upper() == "PORTABLE":
        pc_rec    = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_rec = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        reason_rec = "Utilisateur nomade — PC portable recommande"
    else:
        pc_rec    = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_rec = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        reason_rec = "Poste fixe standard — remplacement planifie"

    is_urgent = any(x in prio for x in ["Urgent", "Prioritaire"])
    rec_bg    = "#FFF0F2" if is_urgent else "#F0FFF4"
    rec_border = "#C8102E" if is_urgent else "#2ECC71"

    rec_html = (
        f'<div style="background:{rec_bg};border-left:4px solid {rec_border};'
        f'border-radius:0 8px 8px 0;padding:16px;margin:12px 0;">'
        f'<div style="font-size:14px;margin-bottom:6px;"><b>PC Recommande:</b> {pc_rec}</div>'
        f'<div style="font-size:14px;margin-bottom:6px;"><b>Specifications:</b> {specs_rec}</div>'
        f'<div style="font-size:14px;"><b>Justification:</b> {reason_rec}</div>'
        '</div>'
    )
    st.markdown(rec_html, unsafe_allow_html=True)

    # ── Observation ───────────────────────────────────────────────────────
    obs = val_str(row.get("observation", ""))
    if obs and obs != "N/A":
        obs_html = (
            '<div style="background:#FFF8DC;border:1px solid #F1C40F;border-radius:8px;'
            f'padding:12px;margin:12px 0;"><b>Observation:</b> {obs}</div>'
        )
        st.markdown(obs_html, unsafe_allow_html=True)

    # ── Imprimantes de la direction ───────────────────────────────────────
    if df_imp is not None and "direction" in df_imp.columns:
        user_dir = direction.upper()
        imp_data = df_imp[df_imp["direction"].str.upper() == user_dir]
        if len(imp_data) > 0:
            st.markdown(f"#### Imprimantes — Direction {user_dir}")
            show_cols = [c for c in ["modele", "type_imp", "site", "n_inventaire", "n_serie", "annee_acquisition"] if c in imp_data.columns]
            st.dataframe(imp_data[show_cols], use_container_width=True)
            st.caption(f"{len(imp_data)} imprimante(s) dans la direction {user_dir}")

    # ── Export Fiche Excel ────────────────────────────────────────────────
    st.markdown("#### Exporter la Fiche")
    fiche_data = {
        "Champ":  list(pc_info.keys()) + ["Priorite", "Recommandation", "Raisons"],
        "Valeur": list(pc_info.values()) + [prio, rec, raisons],
    }
    df_fiche = pd.DataFrame(fiche_data)
    xlsx_bytes = export_excel(df_fiche, "Profil Utilisateur")
    safe_name = selected_user.replace(" ", "_").replace("/", "-")
    st.download_button(
        label=f"Telecharger fiche — {selected_user}",
        data=xlsx_bytes,
        file_name=f"profil_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
