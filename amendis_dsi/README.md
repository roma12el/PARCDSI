# 🏢 DSI Amendis — Tableau de Bord Inventaire Tétouan v3.0

Application Streamlit avancée pour la gestion de l'inventaire informatique du site Tétouan — Direction des Systèmes d'Information.

## 🚀 Déploiement sur Streamlit Cloud

1. **Upload sur GitHub** — Créez un repository et uploadez tous les fichiers
2. **Déployer** sur [share.streamlit.io](https://share.streamlit.io)  
   - Main file path: `app.py`

## 📁 Structure

```
amendis_dsi/
├── app.py              ← Application principale Streamlit
├── data_loader.py      ← Chargement robuste des données Excel (v3)
├── charts.py           ← Visualisations Plotly
├── pdf_export.py       ← Génération PDF professionnelle (NOUVEAU)
├── requirements.txt
├── .streamlit/
│   └── config.toml
└── README.md
```

## 🎯 Fonctionnalités v3.0

| Page | Description |
|------|-------------|
| 🏠 Tableau de bord | KPIs globaux + graphiques interactifs |
| 📊 Rapport PC | Statistiques + recommandations + **export PDF** |
| 🖨️ Rapport Imprimantes | Analyse parc + **export PDF** |
| 📋 Rapport Général | Vue globale synthèse + **export PDF** |
| 👤 Profil Utilisateur | Fiche technique + **export PDF & Excel** |

## 📄 Exports PDF disponibles
- **Fiche utilisateur** : profil complet avec recommandation DSI
- **Rapport PC** : statistiques, âge, RAM, processeurs, urgents
- **Rapport Imprimantes** : répartition et inventaire complet
- **Rapport Général** : synthèse PC + Imprimantes par direction

## 🔧 Amélioration du chargement des données (v3)
- Détection automatique des colonnes par mots-clés
- Lecture robuste sur toutes les feuilles Excel
- Nettoyage intelligent RAM (formats variés)
- Filtrage des lignes parasites amélioré
- Calcul d'ancienneté et recommandations plus précis

## 📊 Fichiers Excel supportés
- **PC** : `_Inventaire_Tétouan_2026_.xlsx` (feuilles: Ecart INV FIN 2025, mouvement 2025...)
- **Imprimantes** : `Inventaire_des_Imprimantes_Tétouan_.xlsx` (feuilles DSI, DRH, DEA...)
