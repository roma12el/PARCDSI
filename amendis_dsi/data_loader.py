"""
data_loader.py — Chargement robuste des fichiers Excel Amendis DSI Tétouan v4
- Garde TOUS les utilisateurs même avec valeurs manquantes
- Extrait les observations
- Lit la feuille "mouvement 2025"
- Lit la feuille "Chrombook" avec specs auto-complétées
- Lit toutes les directions d'imprimantes
"""
import pandas as pd
import numpy as np
import re

CURRENT_YEAR = 2026

# ─── Specs Chromebook par marque ──────────────────────────────────────────────
CHROMEBOOK_SPECS = {
    "default": {"ram": "4 GO", "ram_go": 4, "processeur": "Intel Celeron N4020", "hdd": "32 GO eMMC", "os": "ChromeOS"},
    "HP":      {"ram": "4 GO", "ram_go": 4, "processeur": "Intel Celeron N4020", "hdd": "32 GO eMMC", "os": "ChromeOS"},
    "LENOVO":  {"ram": "4 GO", "ram_go": 4, "processeur": "MediaTek MT8183",     "hdd": "32 GO eMMC", "os": "ChromeOS"},
    "ACER":    {"ram": "4 GO", "ram_go": 4, "processeur": "Intel Celeron N4020", "hdd": "32 GO eMMC", "os": "ChromeOS"},
}

# ─── Base specs PC par modèle ─────────────────────────────────────────────────
PC_SPECS_DB = {
    "HP COMPAQ 6000 PRO":       {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core 2 Duo E8400",  "hdd": "250 GO"},
    "HP COMPAQ 6005 PRO":       {"ram": "4 GO", "ram_go": 4,  "processeur": "AMD Athlon II X2 B24",    "hdd": "250 GO"},
    "HP COMPAQ 6305":           {"ram": "4 GO", "ram_go": 4,  "processeur": "AMD A4-5300B",            "hdd": "250 GO"},
    "HP PRO 3010":              {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core 2 Duo E7500",  "hdd": "250 GO"},
    "HP PRO 3120":              {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core 2 Duo",        "hdd": "250 GO"},
    "HP PRODESK 400 G3":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-6500",      "hdd": "500 GO"},
    "HP PRO DESK 400 G4":       {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7500",      "hdd": "500 GO"},
    "HP PRODESK 400 G4":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7500",      "hdd": "500 GO"},
    "HP PRODESK 400 G5":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-8500",      "hdd": "500 GO"},
    "HP PRO 400 G4":            {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7500",      "hdd": "500 GO"},
    "HP PRO 290 G9":            {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-12500",     "hdd": "512 GO SSD"},
    "HP PRO 300 G6":            {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-10400",     "hdd": "500 GO"},
    "HP PRO SFF 400 G9":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-12500",     "hdd": "256 GO SSD"},
    "HP COMPAQ 600":            {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-4570T",     "hdd": "500 GO"},
    "HP 600 G2":                {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-6500",      "hdd": "500 GO"},
    "HP Z 620":                 {"ram": "16 GO","ram_go": 16, "processeur": "Intel Xeon E5-1620",      "hdd": "1 TO"},
    "HP Z 640":                 {"ram": "16 GO","ram_go": 16, "processeur": "Intel Xeon E5-2630",      "hdd": "1 TO"},
    "HP Z6 G4":                 {"ram": "32 GO","ram_go": 32, "processeur": "Intel Xeon Silver 4110",  "hdd": "512 GO SSD"},
    "HP Z8 G4":                 {"ram": "32 GO","ram_go": 32, "processeur": "Intel Xeon Gold 5118",    "hdd": "512 GO SSD"},
    "HP Z 600":                 {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Xeon X5650",        "hdd": "500 GO"},
    "DELL OPTI PLEX 3020":      {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-4590",      "hdd": "500 GO"},
    "DELL OPTIPLEX 3020":       {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-4590",      "hdd": "500 GO"},
    "DELL OPTIPLEX FORM 7020":  {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-14500",     "hdd": "256 GO SSD"},
    "DELL OPTIPLEX 9010":       {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-3470",      "hdd": "500 GO"},
    "DELL VOSTRO 3670":         {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-8400",      "hdd": "1 TO"},
    "LENOVO NEO 50T":           {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-12400",     "hdd": "512 GO SSD"},
    "LENOVO THINK CENTRE M72":  {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Celeron G1610",     "hdd": "250 GO"},
    "LENOVO M72":               {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Celeron G1610",     "hdd": "250 GO"},
    "HP PROBOOK 640 G1":        {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core i5-4200M",     "hdd": "320 GO"},
    "HP PROBOOK 6470":          {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core i5-3210M",     "hdd": "300 GO"},
    "HP PROBOOK 6460":          {"ram": "4 GO", "ram_go": 4,  "processeur": "Intel Core i5-2520M",     "hdd": "250 GO"},
    "HP 250 G6":                {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7200U",     "hdd": "500 GO"},
    "HP PROBOOK 450 G4":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7200U",     "hdd": "500 GO"},
    "HP PROBOOK 450 G5":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-8250U",     "hdd": "500 GO"},
    "HP PROBOOK 450 G8":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-1135G7",    "hdd": "256 GO SSD"},
    "HP PROBOOK 250 G8":        {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-1135G7",    "hdd": "256 GO SSD"},
    "HP PROBOOK 440 G11":       {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-1334U",     "hdd": "256 GO SSD"},
    "HP 450 G4":                {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Core i5-7200U",     "hdd": "500 GO"},
    "HP OMEN":                  {"ram": "16 GO","ram_go": 16, "processeur": "Intel Core i7-11800H",    "hdd": "1 TO SSD"},
    "HP ELITEBOOK 830":         {"ram": "16 GO","ram_go": 16, "processeur": "Intel Core Ultra 7",      "hdd": "512 GO SSD"},
    "HP PROBOOK 4 G1I":         {"ram": "8 GO", "ram_go": 8,  "processeur": "Intel Ultra 5",           "hdd": "512 GO SSD"},
}

INV_COL_MAP = {
    "UTILISATEURS": "utilisateur", "UTILISATEUR": "utilisateur",
    "NOM PRENOM": "utilisateur", "NOM ET PRENOM": "utilisateur",
    "TRAITE DEMANDE CLIENT": "traite_demande", "TRAITE DEMANDE CLIENT ": "traite_demande",
    "WATERP OUI/NON": "waterp", "POSTES SENSIBLE": "poste_sensible", "POSTE SENSIBLE": "poste_sensible",
    "DIRECTION": "direction", "SITE": "site", "TYPE": "type_pc",
    "N° INVENTAIRE": "n_inventaire", "N°INVENTAIRE": "n_inventaire",
    "INVENTAIRE FIN ANNEE": "inventaire_fin",
    "DATE D'ACQUISITION": "date_acquisition", "ANNEE D'ACQUISITION": "annee_acquisition",
    "ANCIENNETE": "anciennete", "ANCIENNETÉ": "anciennete",
    "PERIODE  D'ANCIENNETE": "periode_anciennete", "PERIODE D'ANCIENNETE": "periode_anciennete",
    "N° SERIE ": "n_serie", "N° SERIE": "n_serie", "N°SERIE": "n_serie",
    "MODELES": "descriptions", "MODÈLES": "descriptions",
    "DESCRIPTIONS": "descriptions", "DESCRIPTIONS ": "descriptions",
    "RAM": "ram", "PROCESSEUR": "processeur", "HDD": "hdd",
    "OBSERVATION ": "observation", "OBSERVATION": "observation",
}

TRASH_USERS = {
    "UTILISATEUR", "UTILISATEURS", "NOM", "NOM PRENOM", "NOM ET PRENOM",
    "PC BUREAU", "PC PORTABLE", "NAN", "", "TOTAL", "SOUS-TOTAL",
    "SITE TETOUAN", "D S I", "DSI", "DIRECTION"
}


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        try:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({"nan": np.nan, "NaT": np.nan, "None": np.nan, "": np.nan})
        except Exception:
            pass
    return df


def _clean_ram(val):
    if pd.isna(val): return np.nan
    s = str(val).upper().strip()
    m = re.search(r'(\d+)', s)
    if m:
        v = int(m.group(1))
        if 0 < v <= 128: return v
    return np.nan


def _age_category(annee):
    try:
        age = CURRENT_YEAR - int(float(annee))
        if age <= 3:    return "Récent (≤3 ans)"
        elif age <= 5:  return "3-5 ans"
        elif age <= 10: return "5-10 ans"
        else:            return "Plus de 10 ans"
    except Exception:
        return "Inconnu"


def _recommend_pc(row):
    try:    age = CURRENT_YEAR - int(float(row.get("annee_acquisition", CURRENT_YEAR)))
    except: age = 0
    try:    ram = float(row.get("ram_go", 0) or 0)
    except: ram = 0
    proc = str(row.get("processeur", "")).upper()

    score, reasons = 0, []
    if age > 10:   score += 3; reasons.append(f"Ancienneté critique: {age} ans")
    elif age > 7:  score += 2; reasons.append(f"Ancienneté: {age} ans")
    elif age > 5:  score += 1; reasons.append(f"Ancienneté: {age} ans")
    if 0 < ram < 4:   score += 2; reasons.append(f"RAM insuffisante ({int(ram)} GO)")
    elif 0 < ram < 8: score += 1; reasons.append(f"RAM faible ({int(ram)} GO)")
    if any(k in proc for k in ["DUAL CORE", "CELERON", "PENTIUM", "ATOM"]):
        score += 2; reasons.append("Processeur obsolète")
    elif "I3" in proc and age > 5:
        score += 1; reasons.append("Processeur i3 vieillissant")

    if   score >= 5: prio, rec = "🔴 Urgent",        "Remplacement immédiat requis"
    elif score >= 3: prio, rec = "🟠 Prioritaire",   "Remplacement dans l'année"
    elif score >= 1: prio, rec = "🟡 À surveiller",  "Planifier le renouvellement"
    else:             prio, rec = "🟢 Opérationnel", "Aucune action immédiate"

    return pd.Series({"priorite": prio, "recommandation": rec,
                      "raisons": ", ".join(reasons) if reasons else "Équipement conforme",
                      "score": score, "age_annees": age})


def _rename_cols(df: pd.DataFrame) -> pd.DataFrame:
    rename = {}
    for col in df.columns:
        key = str(col).strip().upper().rstrip()
        if key in INV_COL_MAP:
            rename[col] = INV_COL_MAP[key]
        else:
            for k, v in INV_COL_MAP.items():
                if k in key or (len(key) > 3 and key in k):
                    rename[col] = v; break
    return df.rename(columns=rename)


def _fill_specs(df: pd.DataFrame) -> pd.DataFrame:
    """Complète specs manquantes depuis la base de données."""
    def fill_row(row):
        desc = str(row.get("descriptions", "")).upper()
        specs = None
        for model_key, model_specs in PC_SPECS_DB.items():
            if model_key.upper() in desc:
                specs = model_specs; break
        if specs is None: return row
        if pd.isna(row.get("ram"))        or str(row.get("ram", "")).strip()        in ("", "nan"): row["ram"]        = specs.get("ram", row.get("ram"))
        if pd.isna(row.get("ram_go"))     or (row.get("ram_go") or 0) == 0:                          row["ram_go"]     = specs.get("ram_go", row.get("ram_go"))
        if pd.isna(row.get("processeur")) or str(row.get("processeur", "")).strip() in ("", "nan"): row["processeur"] = specs.get("processeur", row.get("processeur"))
        if pd.isna(row.get("hdd"))        or str(row.get("hdd", "")).strip()        in ("", "nan"): row["hdd"]        = specs.get("hdd", row.get("hdd"))
        return row
    return df.apply(fill_row, axis=1)


def _find_header_row(raw: pd.DataFrame, keywords=None):
    if keywords is None:
        keywords = ["UTILISATEUR", "UTILISATEURS", "NOM", "DIRECTION", "TYPE", "N° INV"]
    for i, row in raw.iterrows():
        row_str = " ".join(str(v).upper() for v in row.values if pd.notna(v))
        if any(k in row_str for k in keywords):
            return i
    return 0


def _clean_pc_df(df: pd.DataFrame) -> pd.DataFrame:
    df = _normalize(df)
    if "type_pc" in df.columns:
        df["type_pc"] = df["type_pc"].astype(str).str.upper().str.strip()
        df["type_pc"] = df["type_pc"].replace({
            "BUREAU ": "BUREAU", "PORTABLE ": "PORTABLE",
            "CHROMBOOK": "CHROMEBOOK", "CHROMBOK": "CHROMEBOOK", "STATION ": "STATION",
        })
    if "direction" in df.columns:
        df["direction"] = df["direction"].astype(str).str.upper().str.strip()
    for col in ["annee_acquisition", "anciennete"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "ram" in df.columns:
        df["ram_go"] = df["ram"].apply(_clean_ram)
    if "descriptions" in df.columns:
        df = _fill_specs(df)
    if "annee_acquisition" in df.columns:
        df["categorie_age"] = df["annee_acquisition"].apply(_age_category)
        recs = df.apply(_recommend_pc, axis=1)
        df = pd.concat([df, recs], axis=1)
    return df.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT INVENTAIRE PC
# ═══════════════════════════════════════════════════════════════════════════════
def load_inventaire(file) -> dict:
    """
    Retourne: {
      'pc': DataFrame principal (TOUS les utilisateurs),
      'mouvements': DataFrame mouvements (None si absent),
      'chromebooks': DataFrame Chrombooks (None si absent),
    }
    """
    xl = pd.ExcelFile(file)
    sheets = xl.sheet_names
    result = {"pc": None, "mouvements": None, "chromebooks": None}

    # ── 1. Feuille principale PC ──────────────────────────────────────────────
    pc_priority = ["Ecart INV FIN 2025", "INVENTAIRE", "Inventaire", "PC", "Sheet1", "Feuil1"]
    for sheet in pc_priority + [s for s in sheets if s not in pc_priority]:
        if sheet not in sheets:
            continue
        # Sauter les feuilles mouvement/chrom
        if any(k in sheet.lower() for k in ["mouvement", "mouvt", "chrom"]):
            continue
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None).dropna(how="all").reset_index(drop=True)
            if len(raw) < 5: continue
            hrow = _find_header_row(raw)
            df = pd.read_excel(file, sheet_name=sheet, header=hrow).dropna(how="all").reset_index(drop=True)
            df = _rename_cols(df)
            if "utilisateur" not in df.columns and "direction" not in df.columns:
                continue
            # GARDER TOUS — même si valeurs manquantes — juste supprimer lignes sans utilisateur du tout
            if "utilisateur" in df.columns:
                df = df[df["utilisateur"].notna()]
                df = df[~df["utilisateur"].astype(str).str.upper().str.strip().isin(TRASH_USERS)]
            if len(df) > 5:
                df = _clean_pc_df(df)
                result["pc"] = df
                break
        except Exception as e:
            continue

    # ── 2. Feuille "mouvement" ────────────────────────────────────────────────
    for sheet in sheets:
        if "mouvement" in sheet.lower() or "mouvt" in sheet.lower():
            try:
                raw = pd.read_excel(file, sheet_name=sheet, header=None).dropna(how="all").reset_index(drop=True)
                hrow = _find_header_row(raw)
                df_m = pd.read_excel(file, sheet_name=sheet, header=hrow).dropna(how="all").reset_index(drop=True)
                df_m = _rename_cols(df_m)
                if "utilisateur" in df_m.columns:
                    df_m = df_m[df_m["utilisateur"].notna()]
                    df_m = df_m[~df_m["utilisateur"].astype(str).str.upper().str.strip().isin(TRASH_USERS)]
                if len(df_m) > 0:
                    df_m = _normalize(df_m)
                    result["mouvements"] = df_m.reset_index(drop=True)
            except Exception:
                pass

    # ── 3. Feuille "Chrombook" ────────────────────────────────────────────────
    for sheet in sheets:
        if "chrom" in sheet.lower():
            try:
                raw = pd.read_excel(file, sheet_name=sheet, header=None).dropna(how="all").reset_index(drop=True)
                hrow = _find_header_row(raw)
                df_c = pd.read_excel(file, sheet_name=sheet, header=hrow).dropna(how="all").reset_index(drop=True)
                df_c = _rename_cols(df_c)
                if "utilisateur" in df_c.columns:
                    df_c = df_c[df_c["utilisateur"].notna()]
                    df_c = df_c[~df_c["utilisateur"].astype(str).str.upper().str.strip().isin(TRASH_USERS)]
                df_c["type_pc"] = "CHROMEBOOK"
                # Compléter specs Chromebook
                for idx, row in df_c.iterrows():
                    desc = str(row.get("descriptions", "")).upper()
                    marque = next((m for m in ["HP", "LENOVO", "ACER"] if m in desc), "default")
                    sp = CHROMEBOOK_SPECS.get(marque, CHROMEBOOK_SPECS["default"])
                    for field, val in [("ram", sp["ram"]), ("ram_go", sp["ram_go"]),
                                       ("processeur", sp["processeur"]), ("hdd", sp["hdd"])]:
                        if pd.isna(row.get(field)) or str(row.get(field, "")).strip() in ("", "nan", "0"):
                            df_c.at[idx, field] = val
                    df_c.at[idx, "os"] = sp["os"]
                df_c = _normalize(df_c)
                # Ajouter catégories et recommandations
                if "annee_acquisition" in df_c.columns:
                    df_c["annee_acquisition"] = pd.to_numeric(df_c["annee_acquisition"], errors="coerce")
                    df_c["categorie_age"] = df_c["annee_acquisition"].apply(_age_category)
                    recs = df_c.apply(_recommend_pc, axis=1)
                    df_c = pd.concat([df_c, recs], axis=1)
                result["chromebooks"] = df_c.reset_index(drop=True)
            except Exception:
                pass

    if result["pc"] is None:
        raise ValueError("Impossible de trouver les données PC. Vérifiez le fichier Excel.")

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT IMPRIMANTES — TOUTES DIRECTIONS
# ═══════════════════════════════════════════════════════════════════════════════
def load_imprimantes(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    sheets = xl.sheet_names
    skip = {"SOMMAIRE", "TOTAL", "RECAP", "RESUME", "BILAN", "SYNTHESE", "SYNTHÈSE"}
    all_dfs = []

    for sheet in sheets:
        if sheet.upper().strip() in skip:
            continue
        try:
            raw = pd.read_excel(file, sheet_name=sheet, header=None).dropna(how="all").reset_index(drop=True)
            if len(raw) < 2: continue

            hrow = _find_header_row(raw, ["NOM", "UTILISATEUR", "MODELE", "MARQUE", "N° INV", "DIRECTION", "SERIE"])
            df_s = pd.read_excel(file, sheet_name=sheet, header=hrow).dropna(how="all").reset_index(drop=True)
            df_s = df_s.loc[:, ~df_s.columns.duplicated()]

            col_map, used = {}, set()
            for col in df_s.columns:
                cs = str(col).strip().upper()
                def try_map(keywords, target):
                    if target not in used and any(k in cs for k in keywords):
                        col_map[col] = target; used.add(target)
                try_map(["NOM", "UTILISATEUR", "PRENOM"], "utilisateur")
                try_map(["DIRECTION", " DIR "], "direction")
                try_map(["SITE", "LOCALITE", "LOCALITÉ"], "site")
                try_map(["N° INV", "NINV", "INVENTAIRE"], "n_inventaire")
                try_map(["DATE ACQ", "DATE D'ACQ"], "date_acquisition")
                try_map(["ANNÉE ACQ", "ANNEE ACQ", "ANNÉE D'ACQ", "ANNEE D'ACQ"], "annee_acquisition")
                try_map(["MARQUE", "MODEL", "DESCRI", "IMPRIM", "DESIGN"], "modele")
                try_map(["RESEAU", "MONO", "TYPE IMP", "TYPE"], "type_imp")
                try_map(["SERIE", "SN", "N° SERIE"], "n_serie")
                try_map(["ETAT", "ÉTAT", "STATUT", "OBS"], "etat")

            if not col_map: continue
            df_s = df_s.rename(columns=col_map)

            if "direction" not in df_s.columns or df_s.get("direction", pd.Series()).isna().all():
                dir_name = re.sub(r'\s*(MAJ|2025|2026)\s*', '', sheet.upper().split("+")[0]).strip()
                df_s["direction"] = dir_name if "LOCALISE" not in dir_name else np.nan

            keep = [c for c in ["utilisateur","direction","site","n_inventaire",
                                  "date_acquisition","annee_acquisition","modele",
                                  "type_imp","n_serie","etat"] if c in df_s.columns]
            if not keep: continue
            df_s = df_s[keep].copy()

            if "utilisateur" in df_s.columns:
                df_s = df_s[df_s["utilisateur"].notna()]
                trash = {"NAN","NOM","UTILISATEUR","NOM PRÉNOM","NOM ET PRÉNOM",""}
                df_s = df_s[~df_s["utilisateur"].astype(str).str.upper().str.strip().isin(trash)]

            if len(df_s) > 0:
                all_dfs.append(df_s)
        except Exception:
            continue

    if not all_dfs:
        raise ValueError("Impossible de charger les données d'imprimantes.")

    df = pd.concat(all_dfs, ignore_index=True)
    df = _normalize(df)
    if "annee_acquisition" in df.columns:
        df["annee_acquisition"] = pd.to_numeric(df["annee_acquisition"], errors="coerce")
    if "direction" in df.columns:
        df["direction"] = df["direction"].astype(str).str.upper().str.strip()
    df = df.drop_duplicates().reset_index(drop=True)
    return df
