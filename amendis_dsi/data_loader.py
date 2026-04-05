"""
data_loader.py — Chargement et nettoyage des fichiers Excel Amendis DSI Tétouan
"""
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO


# ─── Colonnes attendues ───────────────────────────────────────────────────────
INV_COLS = {
    "Utilisateur": "utilisateur",
    "TRAITE DEMANDE CLIENT ": "traite_demande",
    "WATERP OUI/NON": "waterp",
    "Poste sensible": "poste_sensible",
    "DIRECTION": "direction",
    "SITE": "site",
    "TYPE": "type_pc",
    "N° Inventaire": "n_inventaire",
    "INVENTAIRE FIN ANNEE": "inventaire_fin",
    "Date d'acquisition": "date_acquisition",
    "Année d'acquisition": "annee_acquisition",
    "Ancienneté": "anciennete",
    "Période d'ancienneté": "periode_anciennete",
    "N° SERIE ": "n_serie",
    "DESCRIPTIONS ": "descriptions",
    "RAM": "ram",
    "PROCESSEUR": "processeur",
    "HDD": "hdd",
    "OBSERVATION ": "observation",
}

IMP_COLS_NON_LOC = {
    0: "utilisateur",
    1: "direction",
    2: "site",
    3: "n_inventaire",
    4: "date_acquisition",
    5: "annee_acquisition",
    6: "modele",
    7: "type_imp",
    8: "n_serie",
    9: "etat",
}


# ─── Helpers ──────────────────────────────────────────────────────────────────
def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from all string columns."""
    for col in df.columns:
        try:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({"nan": np.nan, "NaT": np.nan, "None": np.nan})
        except Exception:
            pass
    return df


def _clean_ram(val):
    """Normalise la RAM vers GO numérique."""
    if pd.isna(val):
        return np.nan
    s = str(val).upper().strip()
    if "GO" in s or "GB" in s:
        try:
            return int("".join(filter(str.isdigit, s.split("G")[0])))
        except Exception:
            return np.nan
    return np.nan


def _age_category(annee):
    """Catégorie d'ancienneté basée sur l'année d'acquisition."""
    try:
        age = 2026 - int(annee)
        if age <= 3:
            return "Récent (≤3 ans)"
        elif age <= 5:
            return "3-5 ans"
        elif age <= 10:
            return "5-10 ans"
        else:
            return "Plus de 10 ans"
    except Exception:
        return "Inconnu"


def _recommend_pc(row):
    """Recommandation de remplacement PC."""
    age = 2026 - int(row["annee_acquisition"]) if pd.notna(row["annee_acquisition"]) else 0
    ram = row.get("ram_go", 0) or 0
    proc = str(row.get("processeur", "")).upper()
    type_pc = str(row.get("type_pc", "")).upper()

    score = 0
    reasons = []

    if age > 10:
        score += 3
        reasons.append("Ancienneté >10 ans")
    elif age > 7:
        score += 2
        reasons.append("Ancienneté >7 ans")

    if ram < 4:
        score += 2
        reasons.append(f"RAM insuffisante ({ram}GO)")

    if "DUAL CORE" in proc or "CELERON" in proc:
        score += 2
        reasons.append("Processeur obsolète")

    if score >= 5:
        priority = "🔴 Urgent"
        rec = "Remplacement immédiat requis"
    elif score >= 3:
        priority = "🟠 Prioritaire"
        rec = "Remplacement dans l'année"
    elif score >= 1:
        priority = "🟡 À surveiller"
        rec = "Planifier le renouvellement"
    else:
        priority = "🟢 Opérationnel"
        rec = "Aucune action immédiate"

    return pd.Series({
        "priorite": priority,
        "recommandation": rec,
        "raisons": ", ".join(reasons) if reasons else "Équipement conforme",
        "score": score,
    })


# ─── Chargement inventaire PC ─────────────────────────────────────────────────
def load_inventaire(file) -> pd.DataFrame:
    """Charge et nettoie le fichier inventaire PC (Tétouan 2026)."""
    xl = pd.ExcelFile(file)

    # Feuille principale avec données propres
    df = None
    for sheet in ["Ecart INV FIN 2025", "mouvement 2025"]:
        if sheet in xl.sheet_names:
            if sheet == "Ecart INV FIN 2025":
                raw = pd.read_excel(file, sheet_name=sheet)
                # Vérifier que les colonnes sont correctes
                if "Utilisateur" in raw.columns:
                    df = raw
                    break
            elif sheet == "mouvement 2025":
                raw = pd.read_excel(file, sheet_name=sheet, header=3)
                if "Utilisateur" in raw.columns:
                    df = raw
                    break

    if df is None:
        raise ValueError("Impossible de trouver la feuille d'inventaire principale.")

    # Garder seulement les colonnes connues
    rename_map = {}
    for original, clean in INV_COLS.items():
        for col in df.columns:
            if str(col).strip() == original.strip():
                rename_map[col] = clean
                break

    df = df.rename(columns=rename_map)

    # Garder les colonnes qui existent
    keep = [c for c in INV_COLS.values() if c in df.columns]
    df = df[keep].copy()
    df = _normalize(df)

    # Filtrer les lignes vides ou parasites
    df = df[df["utilisateur"].notna()]
    df = df[~df["utilisateur"].str.upper().isin(["UTILISATEUR", "PC BUREAU", "PC PORTABLE", "NAN"])]
    df = df[~df.get("direction", pd.Series(dtype=str)).str.upper().isin(["DIRECTION", "NAN"])]

    # Nettoyage colonnes clés
    if "type_pc" in df.columns:
        df["type_pc"] = df["type_pc"].str.upper().str.strip()
        df = df[df["type_pc"].isin(["BUREAU", "BUREAU ", "PORTABLE", "PORTABLE ", "STATION", "CHROMBOOK"])]
        df["type_pc"] = df["type_pc"].str.strip()

    if "annee_acquisition" in df.columns:
        df["annee_acquisition"] = pd.to_numeric(df["annee_acquisition"], errors="coerce")

    if "anciennete" in df.columns:
        df["anciennete"] = pd.to_numeric(df["anciennete"], errors="coerce")

    if "ram" in df.columns:
        df["ram_go"] = df["ram"].apply(_clean_ram)

    if "direction" in df.columns:
        df["direction"] = df["direction"].str.upper().str.strip()

    if "annee_acquisition" in df.columns:
        df["categorie_age"] = df["annee_acquisition"].apply(_age_category)

    # Recommandations
    cols_rec = ["annee_acquisition", "ram_go", "processeur", "type_pc"]
    if all(c in df.columns for c in cols_rec):
        recs = df.apply(_recommend_pc, axis=1)
        df = pd.concat([df, recs], axis=1)

    df = df.reset_index(drop=True)
    return df


# ─── Chargement inventaire Imprimantes ────────────────────────────────────────
def load_imprimantes(file) -> pd.DataFrame:
    """Charge et nettoie le fichier inventaire imprimantes."""
    xl = pd.ExcelFile(file)
    all_dfs = []

    # Feuilles par direction avec données d'imprimantes
    direction_sheets = [
        "DSI", "DSO+DSC+DCOM", "DRH", "DCF", "DAAI",
        "DCJA", "DEA", "DEL", "DET", "DCL", "DQSE", "DOP",
        "DRH MAJ", "NON LOCALISE ",
    ]

    for sheet in direction_sheets:
        if sheet not in xl.sheet_names:
            continue
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None)

            # Trouver la ligne de headers
            header_row = None
            for i, row in raw.iterrows():
                row_str = " ".join(str(v) for v in row.values).upper()
                if "NOM" in row_str or "UTILISATEUR" in row_str:
                    header_row = i
                    break

            if header_row is None:
                continue

            df_sheet = pd.read_excel(file, sheet_name=sheet, header=header_row)
            df_sheet = df_sheet.dropna(how="all")

            # Mapper les colonnes
            col_map = {}
            for col in df_sheet.columns:
                cs = str(col).strip().upper()
                if any(k in cs for k in ["NOM", "UTILISATEUR", "PRENOM"]):
                    col_map[col] = "utilisateur"
                elif cs in ["DIRECTION", "DIR"]:
                    col_map[col] = "direction"
                elif cs in ["SITE", "LOCALITE", "LOCALITÉ"]:
                    col_map[col] = "site"
                elif "INV" in cs or "N°" in cs:
                    col_map[col] = "n_inventaire"
                elif "DATE" in cs and "ACQ" in cs:
                    col_map[col] = "date_acquisition"
                elif "ANNÉE" in cs or "ANNEE" in cs:
                    col_map[col] = "annee_acquisition"
                elif "MARQUE" in cs or "MODEL" in cs or "DESCRI" in cs or "IMPRIM" in cs:
                    col_map[col] = "modele"
                elif "RESEAU" in cs or "MONO" in cs or "TYPE" in cs:
                    col_map[col] = "type_imp"
                elif "SERIE" in cs or "SN" in cs:
                    col_map[col] = "n_serie"

            if not col_map:
                continue

            df_sheet = df_sheet.rename(columns=col_map)

            # Déduire la direction depuis le nom de la feuille
            if "direction" not in df_sheet.columns:
                dir_name = sheet.split("+")[0].strip() if "+" in sheet else sheet.strip()
                dir_name = dir_name.replace("MAJ", "").replace(" ", "")
                df_sheet["direction"] = dir_name if dir_name != "NON LOCALISE" else np.nan

            keep = [c for c in ["utilisateur", "direction", "site", "n_inventaire",
                                  "date_acquisition", "annee_acquisition", "modele",
                                  "type_imp", "n_serie"] if c in df_sheet.columns]
            df_sheet = df_sheet[keep].copy()
            # Drop duplicate columns
            df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]
            all_dfs.append(df_sheet)

        except Exception:
            continue

    if not all_dfs:
        raise ValueError("Impossible de charger les données d'imprimantes.")

    df = pd.concat(all_dfs, ignore_index=True)
    df = _normalize(df)

    if "utilisateur" in df.columns:
        df = df[df["utilisateur"].notna()]
        df = df[~df["utilisateur"].str.upper().isin(["NAN", "NOM", "UTILISATEUR", "NOM PRÉNOM"])]

    if "annee_acquisition" in df.columns:
        df["annee_acquisition"] = pd.to_numeric(df["annee_acquisition"], errors="coerce")

    if "direction" in df.columns:
        df["direction"] = df["direction"].str.upper().str.strip()

    df = df.drop_duplicates(subset=["utilisateur", "n_inventaire"] if "n_inventaire" in df.columns else None)
    df = df.reset_index(drop=True)
    return df
