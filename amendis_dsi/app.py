"""
app.py — Application DSI Amendis Tétouan
Tableau de bord d'inventaire informatique (PC + Imprimantes)
Version 3.0 — Lecture robuste des données + Export PDF professionnel
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
from pdf_export import (
    export_fiche_utilisateur_pdf,
    export_rapport_pc_pdf,
    export_rapport_imp_pdf,
    export_rapport_general_pdf,
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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
* { font-family: 'Inter', 'Segoe UI', Arial, sans-serif; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1A1A1A 0%, #2D0A10 100%);
    border-right: 3px solid #C8102E;
}
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label { color: #F5A0AE !important; font-weight: 600; }
[data-testid="stSidebar"] .stButton button {
    background: #C8102E !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 700 !important; width: 100%; padding: 10px !important;
}
[data-testid="stSidebar"] .stButton button:hover { background: #E8415B !important; }

.main .block-container { padding: 1.5rem 2rem; max-width: 1400px; }

.amendis-header {
    background: linear-gradient(135deg, #C8102E 0%, #8B0A1E 100%);
    color: white; padding: 24px 32px; border-radius: 16px;
    margin-bottom: 24px; box-shadow: 0 4px 20px rgba(200,16,46,0.3);
}
.amendis-header h1 { margin: 0; font-size: 28px; font-weight: 700; }
.amendis-header p { margin: 4px 0 0; opacity: 0.85; font-size: 14px; }

.kpi-card {
    background: white; border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08); border-left: 5px solid #C8102E;
}
.kpi-number { font-size: 36px; font-weight: 700; color: #C8102E; margin: 0; }
.kpi-label { font-size: 13px; color: #6B6B6B; font-weight: 500; margin: 4px 0 0; }

.section-title {
    font-size: 20px; font-weight: 700; color: #C8102E;
    border-bottom: 2px solid #C8102E; padding-bottom: 8px; margin: 32px 0 20px;
}

.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] {
    background: #F2F2F2; border-radius: 8px 8px 0 0;
    font-weight: 600; color: #2D2D2D; padding: 10px 20px;
}
.stTabs [aria-selected="true"] { background: #C8102E !important; color: white !important; }

.dataframe thead tr th { background-color: #C8102E !important; color: white !important; font-weight: 700; }

.stButton > button {
    background: #C8102E; color: white; border: none;
    border-radius: 8px; font-weight: 600; padding: 8px 20px; transition: all 0.2s;
}
.stButton > button:hover { background: #E8415B; }

.profile-card {
    background: linear-gradient(135deg, #FFF8F9 0%, #FFECEF 100%);
    border: 2px solid #C8102E; border-radius: 16px; padding: 24px; margin-bottom: 16px;
}
.profile-name { font-size: 22px; font-weight: 700; color: #C8102E; }
.profile-badge {
    display: inline-block; background: #C8102E; color: white;
    padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; margin: 4px 4px 0 0;
}

.pdf-btn-container {
    background: #F9F0F2; border: 1px solid #F5A0AE; border-radius: 10px;
    padding: 16px; margin: 12px 0;
}
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
         "Rapport Général", "Profil Utilisateur"],
        label_visibility="collapsed",
    )

    # Filtres dynamiques
    st.markdown("---")
    st.markdown("### 🎛️ Filtres")
    df_inv_raw = st.session_state.df_inv
    df_imp_raw = st.session_state.df_imp

    filter_direction = None
    if df_inv_raw is not None and "direction" in df_inv_raw.columns:
        dirs = sorted(df_inv_raw["direction"].dropna().unique().tolist())
        filter_direction = st.multiselect("Direction(s)", dirs, default=[], key="f_dir")

    filter_site = None
    if df_inv_raw is not None and "site" in df_inv_raw.columns:
        sites = sorted(df_inv_raw["site"].dropna().unique().tolist())
        filter_site = st.multiselect("Site(s)", sites, default=[], key="f_site")

    filter_type = None
    if df_inv_raw is not None and "type_pc" in df_inv_raw.columns:
        types = sorted(df_inv_raw["type_pc"].dropna().unique().tolist())
        filter_type = st.multiselect("Type PC", types, default=[], key="f_type")

    st.markdown("---")
    st.markdown(f"<div style='font-size:10px;color:#666;text-align:center;'>v3.0 • {datetime.now().strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)


# ─── Apply filters ────────────────────────────────────────────────────────────
def apply_filters(df):
    if df is None:
        return df
    d = df.copy()
    if filter_direction and "direction" in d.columns:
        d = d[d["direction"].isin(filter_direction)]
    if filter_site and "site" in d.columns:
        d = d[d["site"].isin(filter_site)]
    if filter_type and "type_pc" in d.columns:
        d = d[d["type_pc"].isin(filter_type)]
    return d


def kpi(label, value, icon=""):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-number">{icon} {value}</div>
        <div class="kpi-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)


def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


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


def val_str(v):
    return str(v) if pd.notna(v) and str(v) not in ("nan", "NaT", "None") else "N/A"


# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="amendis-header">
    <h1>🖥️ DSI Amendis — Inventaire Informatique Tétouan 2026</h1>
    <p>Direction des Systèmes d'Information · Tableau de bord analytique · Site Tétouan</p>
</div>
""", unsafe_allow_html=True)

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
    df_inv = df_inv_f
    df_imp = df_imp_f

    cols = st.columns(5)
    with cols[0]:
        kpi("Total PC", len(df_inv) if df_inv is not None else "—", "🖥️")
    with cols[1]:
        if df_inv is not None and "type_pc" in df_inv.columns:
            bureau = df_inv[df_inv["type_pc"].str.upper().str.strip() == "BUREAU"].shape[0]
            kpi("PC Bureau", bureau, "🖥️")
        else:
            kpi("PC Bureau", "—")
    with cols[2]:
        if df_inv is not None and "type_pc" in df_inv.columns:
            port = df_inv[df_inv["type_pc"].str.upper().str.strip() == "PORTABLE"].shape[0]
            kpi("PC Portables", port, "💻")
        else:
            kpi("PC Portables", "—")
    with cols[3]:
        kpi("Total Imprimantes", len(df_imp) if df_imp is not None else "—", "🖨️")
    with cols[4]:
        if df_inv is not None and "priorite" in df_inv.columns:
            urgent = df_inv[df_inv["priorite"] == "🔴 Urgent"].shape[0]
            kpi("PC à remplacer", urgent, "🔴")
        elif df_inv is not None and "direction" in df_inv.columns:
            kpi("Directions", df_inv["direction"].nunique(), "🏢")

    st.markdown("<br>", unsafe_allow_html=True)

    if df_inv is not None:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = chart_pc_par_direction(df_inv)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_type_pc(df_inv)
            st.plotly_chart(fig, use_container_width=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            fig = chart_age_distribution(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_ram_distribution(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c3:
            fig = chart_processeur(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)

    if df_imp is not None:
        c1, c2 = st.columns(2)
        with c1:
            fig = chart_imp_par_direction(df_imp)
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_imp_type(df_imp)
            if fig: st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport PC
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport PC":
    df_inv = df_inv_f
    if df_inv is None:
        st.warning("Veuillez charger le fichier inventaire PC.")
        st.stop()

    section("Rapport Statistique — Parc PC")

    # ── Export PDF
    st.markdown('<div class="pdf-btn-container">', unsafe_allow_html=True)
    col_pdf1, col_pdf2 = st.columns([3, 1])
    with col_pdf1:
        st.markdown("📄 **Exporter le rapport PC en PDF**")
        st.caption("Rapport professionnel avec statistiques, recommandations et détail des PC urgents")
    with col_pdf2:
        try:
            pdf_bytes = export_rapport_pc_pdf(df_inv)
            st.download_button(
                "⬇️ Télécharger PDF",
                data=pdf_bytes,
                file_name=f"rapport_pc_amendis_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur PDF: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    report_type = st.radio(
        "Affichage",
        ["Par Direction", "Par Site", "Vue Globale"],
        horizontal=True,
        label_visibility="collapsed",
    )

    tabs = st.tabs(["📊 Statistiques", "🔴 Recommandations", "📋 Données détaillées"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            if report_type == "Par Site" and "site" in df_inv.columns:
                cnt = df_inv.groupby("site").size().reset_index(name="count").sort_values("count", ascending=True)
                fig = px.bar(cnt, x="count", y="site", orientation="h",
                              title="📍 Parc PC par Site",
                              labels={"count": "Nombre", "site": "Site"},
                              color_discrete_sequence=["#C8102E"])
                fig.update_layout(height=400, plot_bgcolor="white")
            else:
                fig = chart_pc_par_direction(df_inv)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_age_distribution(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig = chart_ram_distribution(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_processeur(df_inv)
            if fig: st.plotly_chart(fig, use_container_width=True)

        fig = chart_acquisitions_timeline(df_inv)
        if fig: st.plotly_chart(fig, use_container_width=True)

        fig = chart_heatmap_direction_type(df_inv)
        if fig: st.plotly_chart(fig, use_container_width=True)

    with tabs[1]:
        section("Recommandations pour la Prise de Décision")
        fig = chart_priorite_remplacement(df_inv)
        if fig: st.plotly_chart(fig, use_container_width=True)

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
                                                   "annee_acquisition", "ram", "processeur", "raisons"] if c in sub.columns]
                        st.dataframe(sub[show_cols], use_container_width=True)
                        xlsx = export_excel(sub[show_cols], prio.replace("🔴 ", "").replace("🟠 ", "").replace("🟡 ", "").replace("🟢 ", ""))
                        st.download_button(
                            f"⬇️ Export Excel — {prio}",
                            data=xlsx,
                            file_name=f"pc_{prio.replace(' ', '_').replace('🔴', '').replace('🟠', '').replace('🟡', '').replace('🟢', '').strip()}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

    with tabs[2]:
        section("Données Détaillées")
        show_cols = [c for c in ["utilisateur", "direction", "site", "type_pc",
                                   "n_inventaire", "annee_acquisition", "ram", "processeur",
                                   "hdd", "priorite", "recommandation"] if c in df_inv.columns]
        search = st.text_input("Rechercher un utilisateur ou une direction:", key="search_pc")
        disp = df_inv
        if search:
            mask = pd.Series(False, index=disp.index)
            for col in ["utilisateur", "direction", "n_inventaire"]:
                if col in disp.columns:
                    mask |= disp[col].astype(str).str.contains(search, case=False, na=False)
            disp = disp[mask]
        st.dataframe(disp[show_cols], use_container_width=True)
        st.caption(f"{len(disp)} enregistrement(s)")
        xlsx = export_excel(disp[show_cols], "Inventaire PC")
        st.download_button("⬇️ Export Excel complet", data=xlsx,
                           file_name=f"inventaire_pc_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport Imprimantes
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport Imprimantes":
    df_imp = df_imp_f
    if df_imp is None:
        st.warning("Veuillez charger le fichier inventaire imprimantes.")
        st.stop()

    section("Rapport — Parc Imprimantes")

    # ── Export PDF
    st.markdown('<div class="pdf-btn-container">', unsafe_allow_html=True)
    col_pdf1, col_pdf2 = st.columns([3, 1])
    with col_pdf1:
        st.markdown("📄 **Exporter le rapport Imprimantes en PDF**")
    with col_pdf2:
        try:
            pdf_bytes = export_rapport_imp_pdf(df_imp)
            st.download_button(
                "⬇️ Télécharger PDF",
                data=pdf_bytes,
                file_name=f"rapport_imprimantes_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur PDF: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    tabs = st.tabs(["📊 Statistiques", "📋 Données"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            fig = chart_imp_par_direction(df_imp)
            if fig: st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = chart_imp_type(df_imp)
            if fig: st.plotly_chart(fig, use_container_width=True)

        fig = chart_imp_modele(df_imp)
        if fig: st.plotly_chart(fig, use_container_width=True)

    with tabs[1]:
        show_cols = [c for c in ["utilisateur", "direction", "site", "modele",
                                   "type_imp", "n_inventaire", "n_serie", "annee_acquisition"] if c in df_imp.columns]
        search = st.text_input("Rechercher:", key="search_imp")
        disp = df_imp
        if search:
            mask = pd.Series(False, index=disp.index)
            for col in ["utilisateur", "direction", "modele"]:
                if col in disp.columns:
                    mask |= disp[col].astype(str).str.contains(search, case=False, na=False)
            disp = disp[mask]
        st.dataframe(disp[show_cols], use_container_width=True)
        st.caption(f"{len(disp)} enregistrement(s)")
        xlsx = export_excel(disp[show_cols], "Imprimantes")
        st.download_button("⬇️ Export Excel", data=xlsx,
                           file_name=f"inventaire_imprimantes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: Rapport Général
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Rapport Général":
    df_inv = df_inv_f
    df_imp = df_imp_f
    section("Rapport Général — Parc Informatique Complet")

    # ── Export PDF Général
    st.markdown('<div class="pdf-btn-container">', unsafe_allow_html=True)
    col_pdf1, col_pdf2 = st.columns([3, 1])
    with col_pdf1:
        st.markdown("📄 **Exporter le rapport général en PDF**")
        st.caption("Synthèse complète PC + Imprimantes par direction")
    with col_pdf2:
        try:
            pdf_bytes = export_rapport_general_pdf(df_inv, df_imp)
            st.download_button(
                "⬇️ Télécharger PDF",
                data=pdf_bytes,
                file_name=f"rapport_general_amendis_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur PDF: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    # KPIs globaux
    cols = st.columns(4)
    with cols[0]:
        kpi("Total PC", len(df_inv) if df_inv is not None else "—", "🖥️")
    with cols[1]:
        kpi("Total Imprimantes", len(df_imp) if df_imp is not None else "—", "🖨️")
    with cols[2]:
        total_eq = (len(df_inv) if df_inv is not None else 0) + (len(df_imp) if df_imp is not None else 0)
        kpi("Total Équipements", total_eq, "💡")
    with cols[3]:
        if df_inv is not None and "priorite" in df_inv.columns:
            urgent = df_inv[df_inv["priorite"] == "🔴 Urgent"].shape[0]
            kpi("PC Urgents", urgent, "🔴")

    st.markdown("<br>", unsafe_allow_html=True)

    # Synthèse par direction
    if df_inv is not None and "direction" in df_inv.columns:
        section("Synthèse par Direction")
        agg = df_inv.groupby("direction").agg(
            Total_PC=("utilisateur", "count")
        ).reset_index()

        if "type_pc" in df_inv.columns:
            for tp in ["BUREAU", "PORTABLE"]:
                sub = df_inv[df_inv["type_pc"] == tp].groupby("direction").size().reset_index(name=tp.capitalize())
                agg = agg.merge(sub, on="direction", how="left")
            agg = agg.fillna(0)
            agg["Bureau"] = agg["Bureau"].astype(int)
            agg["Portable"] = agg["Portable"].astype(int)

        if "priorite" in df_inv.columns:
            urg = df_inv[df_inv["priorite"] == "🔴 Urgent"].groupby("direction").size().reset_index(name="Urgents")
            agg = agg.merge(urg, on="direction", how="left").fillna(0)
            agg["Urgents"] = agg["Urgents"].astype(int)

        if df_imp is not None and "direction" in df_imp.columns:
            imp_agg = df_imp.groupby("direction").size().reset_index(name="Imprimantes")
            agg = agg.merge(imp_agg, on="direction", how="left").fillna(0)
            agg["Imprimantes"] = agg["Imprimantes"].astype(int)

        agg = agg.sort_values("Total_PC", ascending=False)
        st.dataframe(agg, use_container_width=True)
        xlsx = export_excel(agg, "Synthèse")
        st.download_button("⬇️ Export Excel synthèse", data=xlsx,
                           file_name=f"synthese_generale_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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
        '<b>Mode Profil</b> — Saisissez le nom d\'un utilisateur pour générer sa fiche technique complète '
        'avec export PDF et Excel professionnel.</div>',
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
        st.error("Colonne Utilisateur non trouvée.")
        st.stop()

    mask = df_inv["utilisateur"].str.contains(search_name, case=False, na=False)
    results = df_inv[mask]

    if len(results) == 0:
        st.error(f"Aucun utilisateur trouvé pour: {search_name}")
        st.info("Astuce: Essayez avec juste le nom de famille.")
        st.stop()

    if len(results) > 1:
        st.info(f"{len(results)} correspondance(s) trouvée(s). Sélectionnez l'utilisateur:")
        user_options = results["utilisateur"].unique().tolist()
        selected_user = st.selectbox("Utilisateur", user_options)
        user_data = results[results["utilisateur"] == selected_user]
    else:
        user_data = results
        selected_user = results.iloc[0]["utilisateur"]

    row = user_data.iloc[0]

    prio        = val_str(row.get("priorite"))
    raisons     = val_str(row.get("raisons"))
    rec         = val_str(row.get("recommandation"))
    waterp      = val_str(row.get("waterp", "NON")).upper()
    sensible    = val_str(row.get("poste_sensible", "NON")).upper()
    direction   = val_str(row.get("direction"))
    site        = val_str(row.get("site"))
    type_pc     = val_str(row.get("type_pc"))

    prio_color = "#C8102E"
    if "Prioritaire" in prio: prio_color = "#E67E22"
    elif "surveiller" in prio: prio_color = "#F1C40F"
    elif "Opérationnel" in prio or "Operationnel" in prio: prio_color = "#2ECC71"

    badges_html = (
        f'<span class="profile-badge">{direction}</span>'
        f'<span class="profile-badge">{site}</span>'
        f'<span class="profile-badge">{type_pc}</span>'
    )
    if "OUI" in sensible:
        badges_html += '<span class="profile-badge">Poste Sensible</span>'
    if "OUI" in waterp:
        badges_html += '<span class="profile-badge">WATERP</span>'

    raisons_html = ""
    if raisons and raisons not in ("nan", "", "N/A"):
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

    # Fiche Technique PC
    st.markdown("#### Fiche Technique PC")
    c1, c2, c3 = st.columns(3)

    pc_info = {
        "N° Inventaire":     val_str(row.get("n_inventaire")),
        "Modèle":            val_str(row.get("descriptions")),
        "N° Série":          val_str(row.get("n_serie")),
        "Année acquisition": val_str(row.get("annee_acquisition")),
        "Ancienneté":        val_str(row.get("anciennete")) + " ans",
        "Catégorie âge":     val_str(row.get("categorie_age")),
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

    # Recommandation PC
    st.markdown("#### Recommandation DSI")
    is_sensible = "OUI" in sensible
    is_portable = "PORTABLE" in type_pc.upper()

    if is_sensible:
        pc_rec    = "HP EliteBook 840 G10 / Dell Latitude 5540"
        specs_rec = "Intel Core i7, 16 GO RAM, 512 GO SSD, Windows 11 Pro"
        reason_rec = "Poste sensible — sécurité renforcée requise"
    elif is_portable:
        pc_rec    = "HP ProBook 450 G10 / Dell Latitude 3540"
        specs_rec = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        reason_rec = "Utilisateur nomade — PC portable recommandé"
    else:
        pc_rec    = "HP ProDesk 400 G9 MT / Dell OptiPlex 3000"
        specs_rec = "Intel Core i5, 8 GO RAM, 256 GO SSD, Windows 11"
        reason_rec = "Poste fixe standard — remplacement planifié"

    is_urgent = any(x in prio for x in ["Urgent", "Prioritaire"])
    rec_bg    = "#FFF0F2" if is_urgent else "#F0FFF4"
    rec_border = "#C8102E" if is_urgent else "#2ECC71"

    rec_html = (
        f'<div style="background:{rec_bg};border-left:4px solid {rec_border};'
        f'border-radius:0 8px 8px 0;padding:16px;margin:12px 0;">'
        f'<div style="font-size:14px;margin-bottom:6px;"><b>PC Recommandé:</b> {pc_rec}</div>'
        f'<div style="font-size:14px;margin-bottom:6px;"><b>Spécifications:</b> {specs_rec}</div>'
        f'<div style="font-size:14px;"><b>Justification:</b> {reason_rec}</div>'
        '</div>'
    )
    st.markdown(rec_html, unsafe_allow_html=True)

    # Observation
    obs = val_str(row.get("observation", ""))
    if obs and obs != "N/A":
        obs_html = (
            '<div style="background:#FFF8DC;border:1px solid #F1C40F;border-radius:8px;'
            f'padding:12px;margin:12px 0;"><b>Observation:</b> {obs}</div>'
        )
        st.markdown(obs_html, unsafe_allow_html=True)

    # Imprimantes direction
    imp_dir_data = None
    if df_imp is not None and "direction" in df_imp.columns:
        user_dir = direction.upper()
        imp_dir_data = df_imp[df_imp["direction"].str.upper() == user_dir]
        if len(imp_dir_data) > 0:
            st.markdown(f"#### Imprimantes — Direction {user_dir}")
            show_cols = [c for c in ["modele", "type_imp", "site", "n_inventaire", "n_serie", "annee_acquisition"] if c in imp_dir_data.columns]
            st.dataframe(imp_dir_data[show_cols], use_container_width=True)
            st.caption(f"{len(imp_dir_data)} imprimante(s) dans la direction {user_dir}")

    # ── EXPORTS
    st.markdown("#### Exporter la Fiche")
    col_e1, col_e2 = st.columns(2)

    with col_e1:
        # Export PDF Fiche
        try:
            pdf_fiche = export_fiche_utilisateur_pdf(row, imp_dir_data if (imp_dir_data is not None and len(imp_dir_data) > 0) else None)
            safe_name = selected_user.replace(" ", "_").replace("/", "-")
            st.download_button(
                label=f"📄 Fiche PDF — {selected_user}",
                data=pdf_fiche,
                file_name=f"fiche_{safe_name}_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur PDF: {e}")

    with col_e2:
        # Export Excel Fiche
        fiche_data = {
            "Champ":  list(pc_info.keys()) + ["Priorité", "Recommandation", "Raisons", "PC Recommandé", "Spécifications"],
            "Valeur": list(pc_info.values()) + [prio, rec, raisons, pc_rec, specs_rec],
        }
        df_fiche = pd.DataFrame(fiche_data)
        xlsx_bytes = export_excel(df_fiche, "Profil Utilisateur")
        safe_name = selected_user.replace(" ", "_").replace("/", "-")
        st.download_button(
            label=f"📊 Fiche Excel — {selected_user}",
            data=xlsx_bytes,
            file_name=f"profil_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
