# 🏢 DSI Amendis — Tableau de Bord Inventaire Tétouan

Application Streamlit avancée pour la gestion de l'inventaire informatique du site Tétouan — Direction des Systèmes d'Information.

## 🚀 Déploiement sur Streamlit Cloud (share.streamlit.io)

### Étapes:

1. **Upload sur GitHub**
   - Créez un nouveau repository GitHub (ex: `amendis-dsi-inventaire`)
   - Uploadez tous les fichiers de ce ZIP dans le repository

2. **Déployer sur Streamlit Cloud**
   - Allez sur [share.streamlit.io](https://share.streamlit.io)
   - Cliquez sur **"New app"**
   - Sélectionnez votre repository GitHub
   - **Main file path:** `app.py`
   - Cliquez **"Deploy!"**

3. **Utilisation**
   - L'app sera disponible via une URL publique
   - Uploadez vos fichiers Excel via la sidebar

## 📁 Structure des fichiers

```
amendis_dsi/
├── app.py              ← Application principale Streamlit
├── data_loader.py      ← Chargement & nettoyage des données Excel
├── charts.py           ← Toutes les visualisations Plotly
├── requirements.txt    ← Dépendances Python
├── .streamlit/
│   └── config.toml    ← Thème Rouge & Blanc Amendis
└── README.md           ← Ce fichier
```

## 🎯 Fonctionnalités

| Page | Description |
|------|-------------|
| 🏠 Tableau de bord | KPIs globaux + graphiques d'aperçu |
| 📊 Rapport PC | Statistiques par direction/site, recommandations, données |
| 🖨️ Rapport Imprimantes | Analyse parc imprimantes |
| 📋 Rapport Général | Vue globale complète avec synthèse par direction |
| 👤 Profil Utilisateur | Fiche technique personnalisée + recommandation PC |

## 📊 Fichiers Excel supportés

- **Inventaire PC**: `_Inventaire_Tétouan_2026_.xlsx`
  - Feuilles: `Ecart INV FIN 2025`, `mouvement 2025`
- **Imprimantes**: `Inventaire_des_Imprimantes_Tétouan_.xlsx`
  - Feuilles: DSI, DRH, DEA, DEL, DAAI, DCL, etc.

## 🎨 Thème

- **Couleurs principales**: Rouge Amendis (#C8102E) + Blanc
- **Police**: Inter / Segoe UI
- **Visualisations**: Plotly Express (interactif)
