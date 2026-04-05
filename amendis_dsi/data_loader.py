"""
data_loader.py — Chargement et nettoyage robuste des fichiers Excel Amendis DSI Tétouan
"""
import pandas as pd
import numpy as np
from io import BytesIO

CURRENT_YEAR = 2026

# ─── Colonnes inventaire PC ───────────────────────────────────────────────────
INV_COLS = {
    "Utilisateur": "utilisateur",
    "TRAITE DEMANDE CLIENT ": "traite_demande",
    "TRAITE DEMANDE CLIENT": "traite_demande",
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
    "N° SERIE": "n_serie",
    "DESCRIPTIONS ": "descriptions",
    "DESCRIPTIONS": "descriptions",
    "RAM": "ram",
    "PROCESSEUR": "processeur",
    "HDD": "hdd",
    "OBSERVATION ": "observation",
    "OBSERVATION": "observation",
}


# ─── Helpers ──────────────────────────────────────────────────────────────────
def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace de toutes les colonnes string."""
    for col in df.columns:
        try:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({"nan": np.nan, "NaT": np.nan, "None": np.nan, "": np.nan})
        except Exception:
            pass
    return df


def _clean_ram(val):
    """Normalise la RAM vers GO numérique."""
    if pd.isna(val):
        return np.nan
    s = str(val).upper().strip()
    # Cas: "8 GO", "8GO", "8 GB", "8GB", "8"
    import re
    m = re.search(r'(\d+)', s)
    if m:
        v = int(m.group(1))
        # Valeurs réalistes: 1, 2, 4, 8, 16, 32, 64
        if 0 < v <= 128:
            return v
    return np.nan


def _age_category(annee):
    """Catégorie d'ancienneté basée sur l'année d'acquisition."""
    try:
        age = CURRENT_YEAR - int(float(annee))
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
    """Recommandation de remplacement PC basée sur l'âge, RAM, processeur."""
    try:
        age = CURRENT_YEAR - int(float(row.get("annee_acquisition", CURRENT_YEAR)))
    except Exception:
        age = 0

    ram = row.get("ram_go", 0)
    try:
        ram = float(ram) if pd.notna(ram) else 0
    except Exception:
        ram = 0

    proc = str(row.get("processeur", "")).upper()
    score = 0
    reasons = []

    if age > 10:
        score += 3
        reasons.append(f"Ancienneté critique: {age} ans")
    elif age > 7:
        score += 2
        reasons.append(f"Ancienneté: {age} ans")
    elif age > 5:
        score += 1
        reasons.append(f"Ancienneté: {age} ans")

    if ram > 0 and ram < 4:
        score += 2
        reasons.append(f"RAM insuffisante ({int(ram)} GO)")
    elif ram > 0 and ram < 8:
        score += 1
        reasons.append(f"RAM faible ({int(ram)} GO)")

    if any(k in proc for k in ["DUAL CORE", "CELERON", "PENTIUM", "ATOM"]):
        score += 2
        reasons.append("Processeur obsolète")
    elif "I3" in proc and age > 5:
        score += 1
        reasons.append("Processeur i3 vieillissant")

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
        "age_annees": age,
    })


# ─── Recherche flexible de colonnes ──────────────────────────────────────────
def _find_col(df_columns, keywords):
    """Trouve une colonne par mots-clés (insensible à la casse)."""
    for col in df_columns:
        col_up = str(col).upper().strip()
        for kw in keywords:
            if kw.upper() in col_up:
                return col
    return None


def _auto_rename(df: pd.DataFrame) -> pd.DataFrame:
    """Renommage automatique des colonnes PC par détection de mots-clés."""
    rename = {}
    cols = list(df.columns)

    mappings = [
        (["UTILISATEUR", "NOM PRENOM", "NOM ET PRENOM", "EMPLOYE"], "utilisateur"),
        (["DIRECTION", " DIR "], "direction"),
        (["SITE", "LOCALITE", "LOCALITÉ"], "site"),
        (["N° INV", "INVENTAIRE", "N°INV", "NINV"], "n_inventaire"),
        (["DATE ACQ", "DATE D'ACQ"], "date_acquisition"),
        (["ANNÉE ACQ", "ANNEE ACQ", "ANNÉE D'ACQ", "ANNEE D'ACQ"], "annee_acquisition"),
        (["ANCIENNETÉ", "ANCIENNETE"], "anciennete"),
        (["TYPE"], "type_pc"),
        (["N° SERIE", "N°SERIE", "SERIE"], "n_serie"),
        (["DESCRIPTION", "MODELE", "MARQUE"], "descriptions"),
        (["RAM"], "ram"),
        (["PROCESSEUR", "CPU", "PROCESSOR"], "processeur"),
        (["HDD", "DISQUE", "STOCKAGE"], "hdd"),
        (["OBSERVATION"], "observation"),
        (["WATERP"], "waterp"),
        (["SENSIBLE"], "poste_sensible"),
        (["TRAITE", "DEMANDE"], "traite_demande"),
        (["PERIODE", "PÉRIODE"], "periode_anciennete"),
    ]

    used_targets = set()
    for keywords, target in mappings:
        if target in used_targets:
            continue
        found = _find_col(cols, keywords)
        if found and found not in rename:
            rename[found] = target
            used_targets.add(target)

    return df.rename(columns=rename)


# ─── Chargement inventaire PC ─────────────────────────────────────────────────
def load_inventaire(file) -> pd.DataFrame:
    """Charge et nettoie le fichier inventaire PC Tétouan."""
    xl = pd.ExcelFile(file)
    df = None

    # Essayer toutes les feuilles dans l'ordre de priorité
    priority_sheets = ["Ecart INV FIN 2025", "mouvement 2025", "INVENTAIRE", "Inventaire", "PC", "Sheet1"]
    sheets_to_try = priority_sheets + [s for s in xl.sheet_names if s not in priority_sheets]

    for sheet in sheets_to_try:
        if sheet not in xl.sheet_names:
            continue
        try:
            # Essayer d'abord sans header spécial
            for header_row in [0, 1, 2, 3, 4]:
                try:
                    raw = pd.read_excel(file, sheet_name=sheet, header=header_row)
                    raw = raw.dropna(how="all")
                    if len(raw) < 3:
                        continue
                    # Vérifier qu'on a au moins une colonne reconnaissable
                    cols_up = [str(c).upper().strip() for c in raw.columns]
                    has_user = any("UTILISATEUR" in c or "NOM" in c for c in cols_up)
                    has_dir = any("DIRECTION" in c or "DIR" in c for c in cols_up)
                    if has_user or (has_dir and len(raw) > 10):
                        df = raw
                        break
                except Exception:
                    continue
            if df is not None:
                break
        except Exception:
            continue

    if df is None:
        # Dernier recours: première feuille
        try:
            df = pd.read_excel(file, sheet_name=0)
        except Exception:
            raise ValueError("Impossible de lire le fichier Excel. Vérifiez le format.")

    # Renommer les colonnes
    # D'abord essayer le mapping exact
    rename_map = {}
    for original, clean in INV_COLS.items():
        for col in df.columns:
            if str(col).strip() == original.strip():
                rename_map[col] = clean
                break

    if len(rename_map) < 3:
        # Fallback: renommage automatique par mots-clés
        df = _auto_rename(df)
    else:
        df = df.rename(columns=rename_map)

    df = _normalize(df)

    # Garder les colonnes utiles
    target_cols = list(set(INV_COLS.values()))
    keep = [c for c in target_cols if c in df.columns]
    df = df[keep].copy()

    # Filtrer les lignes vides/parasites
    if "utilisateur" in df.columns:
        df = df[df["utilisateur"].notna()]
        garbage = ["UTILISATEUR", "PC BUREAU", "PC PORTABLE", "NAN", "NOM", "NOM PRÉNOM",
                   "NOM ET PRENOM", "TOTAL", "SOUS-TOTAL", ""]
        df = df[~df["utilisateur"].str.upper().isin(garbage)]

    # Nettoyage type_pc
    if "type_pc" in df.columns:
        df["type_pc"] = df["type_pc"].astype(str).str.upper().str.strip()
        valid_types = ["BUREAU", "PORTABLE", "STATION", "CHROMBOOK", "LAPTOP", "DESKTOP"]
        # Garder seulement les valeurs valides (ou NaN)
        mask = df["type_pc"].isin(valid_types) | df["type_pc"].isna()
        df = df[mask]
        # Normalisation
        df["type_pc"] = df["type_pc"].replace({
            "LAPTOP": "PORTABLE", "DESKTOP": "BUREAU",
            "BUREAU ": "BUREAU", "PORTABLE ": "PORTABLE"
        })

    # Conversions numériques
    for col in ["annee_acquisition", "anciennete"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Filtrer les années aberrantes
    if "annee_acquisition" in df.columns:
        df = df[df["annee_acquisition"].isna() | ((df["annee_acquisition"] >= 1990) & (df["annee_acquisition"] <= CURRENT_YEAR))]

    # RAM nettoyée
    if "ram" in df.columns:
        df["ram_go"] = df["ram"].apply(_clean_ram)

    # Normalisation direction
    if "direction" in df.columns:
        df["direction"] = df["direction"].astype(str).str.upper().str.strip()
        df = df[df["direction"].notna() & (df["direction"] != "NAN") & (df["direction"] != "DIRECTION")]

    # Catégorie âge
    if "annee_acquisition" in df.columns:
        df["categorie_age"] = df["annee_acquisition"].apply(_age_category)

    # Recommandations
    needed_cols = ["annee_acquisition"]
    if all(c in df.columns for c in needed_cols):
        recs = df.apply(_recommend_pc, axis=1)
        df = pd.concat([df, recs], axis=1)

    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    return df


# ─── Chargement inventaire Imprimantes ────────────────────────────────────────
def load_imprimantes(file) -> pd.DataFrame:
    """Charge et nettoie le fichier inventaire imprimantes."""
    xl = pd.ExcelFile(file)
    all_dfs = []

    direction_sheets = [
        "DSI", "DSO+DSC+DCOM", "DRH", "DCF", "DAAI", "DCJA",
        "DEA", "DEL", "DET", "DCL", "DQSE", "DOP",
        "DRH MAJ", "NON LOCALISE ", "NON LOCALISE",
    ]

    sheets_to_process = [s for s in xl.sheet_names if s not in ["Sommaire", "SOMMAIRE", "Total"]]

    for sheet in sheets_to_process:
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None)
            raw = raw.dropna(how="all")

            # Trouver la ligne de headers
            header_row = None
            for i, row in raw.iterrows():
                row_str = " ".join(str(v).upper() for v in row.values if pd.notna(v))
                if any(k in row_str for k in ["NOM", "UTILISATEUR", "MODELE", "MARQUE", "N° INV"]):
                    header_row = i
                    break

            if header_row is None:
                # Essayer la première ligne comme header
                header_row = raw.index[0] if len(raw) > 0 else 0

            df_sheet = pd.read_excel(file, sheet_name=sheet, header=header_row)
            df_sheet = df_sheet.dropna(how="all").dropna(axis=1, how="all")

            if len(df_sheet) < 2:
                continue

            # Mapper les colonnes automatiquement
            col_map = {}
            used = set()
            for col in df_sheet.columns:
                cs = str(col).strip().upper()
                if any(k in cs for k in ["NOM", "UTILISATEUR", "PRENOM"]) and "utilisateur" not in used:
                    col_map[col] = "utilisateur"; used.add("utilisateur")
                elif cs in ["DIRECTION", "DIR"] and "direction" not in used:
                    col_map[col] = "direction"; used.add("direction")
                elif any(k in cs for k in ["SITE", "LOCALITE", "LOCALITÉ"]) and "site" not in used:
                    col_map[col] = "site"; used.add("site")
                elif any(k in cs for k in ["N° INV", "NINV", "INVENTAIRE"]) and "n_inventaire" not in used:
                    col_map[col] = "n_inventaire"; used.add("n_inventaire")
                elif "DATE" in cs and "ACQ" in cs and "date_acquisition" not in used:
                    col_map[col] = "date_acquisition"; used.add("date_acquisition")
                elif any(k in cs for k in ["ANNÉE", "ANNEE"]) and "annee_acquisition" not in used:
                    col_map[col] = "annee_acquisition"; used.add("annee_acquisition")
                elif any(k in cs for k in ["MARQUE", "MODEL", "DESCRI", "IMPRIM"]) and "modele" not in used:
                    col_map[col] = "modele"; used.add("modele")
                elif any(k in cs for k in ["RESEAU", "MONO", "TYPE IMP", "TYPE"]) and "type_imp" not in used:
                    col_map[col] = "type_imp"; used.add("type_imp")
                elif any(k in cs for k in ["SERIE", "SN", "N° SERIE"]) and "n_serie" not in used:
                    col_map[col] = "n_serie"; used.add("n_serie")
                elif any(k in cs for k in ["ETAT", "ÉTAT", "STATUT"]) and "etat" not in used:
                    col_map[col] = "etat"; used.add("etat")

            if not col_map:
                continue

            df_sheet = df_sheet.rename(columns=col_map)

            # Inférer la direction depuis le nom de la feuille
            if "direction" not in df_sheet.columns:
                dir_name = sheet.split("+")[0].strip().replace("MAJ", "").strip()
                df_sheet["direction"] = dir_name if "LOCALISE" not in dir_name.upper() else np.nan

            keep = [c for c in ["utilisateur", "direction", "site", "n_inventaire",
                                  "date_acquisition", "annee_acquisition", "modele",
                                  "type_imp", "n_serie", "etat"] if c in df_sheet.columns]
            df_sheet = df_sheet[keep].copy()
            df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]
            all_dfs.append(df_sheet)

        except Exception:
            continue

    if not all_dfs:
        raise ValueError("Impossible de charger les données d'imprimantes. Vérifiez le format du fichier.")

    df = pd.concat(all_dfs, ignore_index=True)
    df = _normalize(df)

    if "utilisateur" in df.columns:
        df = df[df["utilisateur"].notna()]
        df = df[~df["utilisateur"].str.upper().isin(["NAN", "NOM", "UTILISATEUR", "NOM PRÉNOM", "NOM ET PRÉNOM"])]

    if "annee_acquisition" in df.columns:
        df["annee_acquisition"] = pd.to_numeric(df["annee_acquisition"], errors="coerce")

    if "direction" in df.columns:
        df["direction"] = df["direction"].astype(str).str.upper().str.strip()

    # Dédoublonnage
    dedup_cols = ["utilisateur", "n_inventaire"] if "n_inventaire" in df.columns else ["utilisateur"]
    df = df.drop_duplicates(subset=dedup_cols)
    df = df.reset_index(drop=True)
    return df
