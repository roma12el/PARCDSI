# 🏢 DSI Amendis — Tableau de Bord Inventaire Tétouan v6.0

## 🔧 Corrections v6

| # | Problème | Fix |
|---|----------|-----|
| 1 | PDF couleurs (teal/turquoise hors charte) | Alignées avec rouge Amendis #C8102E + navy #1B2A3B |
| 2 | Chromebook bar chart désaligné | tickmode=linear + automargin |
| 3 | "Oui" / "DIRECTION" apparaissent comme direction | Filtre TRASH_DIRS dans Mouvements |
| 4 | "BUREAU" apparaît 2× dans pie type matériel | Normalisation _norm_type() |
| 5 | Nouvelle page Segmentation Utilisateurs | Filtres direction/priorité/âge/type + graphiques + export |
| 6 | Nouvelle page Dashboard HTML | Dashboard interactif intégré dans Streamlit |

## 📁 Structure
```
amendis_dsi_v6/
├── app.py              ← Application principale (v6)
├── data_loader.py      ← Chargement Excel
├── charts.py           ← Couleurs
├── pdf_export.py       ← Export PDF (couleurs corrigées)
├── requirements.txt
├── .streamlit/
│   └── config.toml
└── README.md
```

## 🚀 Déploiement
Main file: `app.py` sur share.streamlit.io
