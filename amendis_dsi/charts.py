"""
charts.py — Visualisations Plotly pour l'application Amendis DSI
"""
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np

# ─── Palette Amendis (blanc / rouge) ──────────────────────────────────────────
COLORS = {
    "rouge":        "#C8102E",
    "rouge_clair":  "#E8415B",
    "rouge_pale":   "#F5A0AE",
    "gris_fonce":   "#2D2D2D",
    "gris_moyen":   "#6B6B6B",
    "gris_clair":   "#F2F2F2",
    "blanc":        "#FFFFFF",
    "vert":         "#2ECC71",
    "orange":       "#E67E22",
    "bleu":         "#2980B9",
}

PALETTE = [
    COLORS["rouge"],
    COLORS["rouge_clair"],
    COLORS["rouge_pale"],
    COLORS["gris_fonce"],
    COLORS["gris_moyen"],
    COLORS["bleu"],
    COLORS["vert"],
    COLORS["orange"],
    "#9B59B6", "#1ABC9C",
]

LAYOUT = dict(
    font_family="Segoe UI, Arial, sans-serif",
    plot_bgcolor=COLORS["blanc"],
    paper_bgcolor=COLORS["blanc"],
    font_color=COLORS["gris_fonce"],
    title_font_size=16,
    title_font_color=COLORS["rouge"],
    margin=dict(t=60, b=40, l=40, r=20),
    legend=dict(
        bgcolor="rgba(255,255,255,0.9)",
        bordercolor=COLORS["rouge"],
        borderwidth=1,
    ),
)


def _apply(fig, height=400):
    fig.update_layout(**LAYOUT, height=height)
    return fig


# ─── PC Charts ────────────────────────────────────────────────────────────────
def chart_pc_par_direction(df: pd.DataFrame):
    cnt = df.groupby("direction").size().reset_index(name="count").sort_values("count", ascending=True)
    fig = px.bar(
        cnt, x="count", y="direction", orientation="h",
        title="📊 Parc PC par Direction",
        labels={"count": "Nombre de PC", "direction": "Direction"},
        color="count",
        color_continuous_scale=[[0, COLORS["rouge_pale"]], [1, COLORS["rouge"]]],
    )
    fig.update_coloraxes(showscale=False)
    return _apply(fig, 450)


def chart_type_pc(df: pd.DataFrame):
    cnt = df["type_pc"].value_counts().reset_index()
    cnt.columns = ["type", "count"]
    fig = px.pie(
        cnt, values="count", names="type",
        title="💻 Répartition Bureau / Portable",
        color_discrete_sequence=[COLORS["rouge"], COLORS["gris_fonce"], COLORS["rouge_pale"]],
        hole=0.4,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return _apply(fig, 380)


def chart_age_distribution(df: pd.DataFrame):
    if "categorie_age" not in df.columns:
        return None
    order = ["Récent (≤3 ans)", "3-5 ans", "5-10 ans", "Plus de 10 ans", "Inconnu"]
    cnt = df["categorie_age"].value_counts().reindex(order, fill_value=0).reset_index()
    cnt.columns = ["categorie", "count"]
    color_map = {
        "Récent (≤3 ans)": COLORS["vert"],
        "3-5 ans": COLORS["bleu"],
        "5-10 ans": COLORS["orange"],
        "Plus de 10 ans": COLORS["rouge"],
        "Inconnu": COLORS["gris_moyen"],
    }
    fig = px.bar(
        cnt, x="categorie", y="count",
        title="📅 Ancienneté du Parc PC",
        labels={"categorie": "Catégorie d'âge", "count": "Nombre de PC"},
        color="categorie",
        color_discrete_map=color_map,
    )
    fig.update_layout(showlegend=False)
    return _apply(fig, 380)


def chart_ram_distribution(df: pd.DataFrame):
    if "ram_go" not in df.columns:
        return None
    r = df["ram_go"].dropna()
    cnt = r.value_counts().sort_index().reset_index()
    cnt.columns = ["ram", "count"]
    cnt["label"] = cnt["ram"].apply(lambda x: f"{int(x)} GO")
    fig = px.bar(
        cnt, x="label", y="count",
        title="🧠 Distribution RAM",
        labels={"label": "RAM", "count": "Nombre"},
        color="count",
        color_continuous_scale=[[0, COLORS["rouge_pale"]], [1, COLORS["rouge"]]],
    )
    fig.update_coloraxes(showscale=False)
    return _apply(fig, 360)


def chart_processeur(df: pd.DataFrame):
    if "processeur" not in df.columns:
        return None
    proc_map = {}
    for v in df["processeur"].dropna():
        s = str(v).upper().strip()
        if "I7" in s:
            proc_map[v] = "Intel Core i7"
        elif "I5" in s:
            proc_map[v] = "Intel Core i5"
        elif "I3" in s:
            proc_map[v] = "Intel Core i3"
        elif "CELERON" in s:
            proc_map[v] = "Intel Celeron"
        elif "DUAL" in s:
            proc_map[v] = "Dual Core"
        elif "XEON" in s:
            proc_map[v] = "Intel Xeon"
        else:
            proc_map[v] = "Autre"
    df2 = df.copy()
    df2["proc_cat"] = df2["processeur"].map(proc_map).fillna("Inconnu")
    cnt = df2["proc_cat"].value_counts().reset_index()
    cnt.columns = ["processeur", "count"]
    fig = px.pie(
        cnt, values="count", names="processeur",
        title="⚙️ Répartition Processeurs",
        color_discrete_sequence=PALETTE,
        hole=0.35,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return _apply(fig, 400)


def chart_priorite_remplacement(df: pd.DataFrame):
    if "priorite" not in df.columns:
        return None
    cnt = df["priorite"].value_counts().reset_index()
    cnt.columns = ["priorite", "count"]
    color_map = {
        "🔴 Urgent": COLORS["rouge"],
        "🟠 Prioritaire": COLORS["orange"],
        "🟡 À surveiller": "#F1C40F",
        "🟢 Opérationnel": COLORS["vert"],
    }
    fig = px.bar(
        cnt, x="priorite", y="count",
        title="⚠️ Recommandations de Remplacement PC",
        labels={"priorite": "Priorité", "count": "Nombre de PC"},
        color="priorite",
        color_discrete_map=color_map,
    )
    fig.update_layout(showlegend=False)
    return _apply(fig, 380)


def chart_heatmap_direction_type(df: pd.DataFrame):
    if "direction" not in df.columns or "type_pc" not in df.columns:
        return None
    pivot = pd.crosstab(df["direction"], df["type_pc"])
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=list(pivot.columns),
        y=list(pivot.index),
        colorscale=[[0, "#FFF0F2"], [0.5, COLORS["rouge_pale"]], [1, COLORS["rouge"]]],
        text=pivot.values,
        texttemplate="%{text}",
        showscale=True,
    ))
    fig.update_layout(
        title="🗺️ Heatmap Direction × Type PC",
        xaxis_title="Type PC",
        yaxis_title="Direction",
    )
    return _apply(fig, 450)


def chart_acquisitions_timeline(df: pd.DataFrame):
    if "annee_acquisition" not in df.columns:
        return None
    cnt = df["annee_acquisition"].dropna().value_counts().sort_index().reset_index()
    cnt.columns = ["annee", "count"]
    cnt["annee"] = cnt["annee"].astype(int)
    fig = px.area(
        cnt, x="annee", y="count",
        title="📈 Évolution des Acquisitions PC par Année",
        labels={"annee": "Année", "count": "Nombre d'acquisitions"},
        color_discrete_sequence=[COLORS["rouge"]],
    )
    fig.update_traces(fill="tozeroy", fillcolor=COLORS["rouge_pale"])
    return _apply(fig, 360)


# ─── Imprimante Charts ────────────────────────────────────────────────────────
def chart_imp_par_direction(df: pd.DataFrame):
    if "direction" not in df.columns:
        return None
    cnt = df.groupby("direction").size().reset_index(name="count").sort_values("count", ascending=True)
    fig = px.bar(
        cnt, x="count", y="direction", orientation="h",
        title="🖨️ Parc Imprimantes par Direction",
        labels={"count": "Nombre", "direction": "Direction"},
        color="count",
        color_continuous_scale=[[0, COLORS["rouge_pale"]], [1, COLORS["rouge"]]],
    )
    fig.update_coloraxes(showscale=False)
    return _apply(fig, 420)


def chart_imp_type(df: pd.DataFrame):
    if "type_imp" not in df.columns:
        return None
    cnt = df["type_imp"].value_counts().reset_index()
    cnt.columns = ["type", "count"]
    fig = px.pie(
        cnt, values="count", names="type",
        title="🔌 Réseau vs Monoposte",
        color_discrete_sequence=[COLORS["rouge"], COLORS["bleu"], COLORS["rouge_pale"]],
        hole=0.4,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return _apply(fig, 360)


def chart_imp_modele(df: pd.DataFrame):
    if "modele" not in df.columns:
        return None
    cnt = df["modele"].value_counts().head(10).reset_index()
    cnt.columns = ["modele", "count"]
    cnt = cnt.sort_values("count", ascending=True)
    fig = px.bar(
        cnt, x="count", y="modele", orientation="h",
        title="🏆 Top 10 Modèles d'Imprimantes",
        labels={"count": "Nombre", "modele": "Modèle"},
        color="count",
        color_continuous_scale=[[0, COLORS["rouge_pale"]], [1, COLORS["rouge"]]],
    )
    fig.update_coloraxes(showscale=False)
    return _apply(fig, 420)
