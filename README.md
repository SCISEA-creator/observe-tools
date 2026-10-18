# Observe-Tools

**Version 3.6.5** · © 2026 SCISEA — Tous droits réservés / All rights reserved

[Français](#français) · [English](#english)

---

## Français

Observe-Tools est une application de bureau destinée aux **observateurs scientifiques embarqués sur les palangriers**. À partir des données saisies dans **ObServe**, elle vérifie la qualité des données, produit les figures et tableaux de la marée, puis génère le rapport de mission au format Word.

Développé par **SCISEA** pour l'**IRD (UMR MARBEC, OB7)**, dans le cadre du programme européen **DCF** (*Data Collection Framework*).

### Fonctionnalités

| Étape | Rôle |
|---|---|
| **Accueil** | Import des données ObServe, résumé de la marée, choix de la zone de pêche, création du dossier de marée |
| **Phase 1 — Correction des données** | Contrôle automatique des 9 tables (règles de cohérence, plages de valeurs, tailles par espèce…), annotation des anomalies (*Validé* / *Justifié*), export d'un rapport de vérification Excel protégé |
| **Phase 2 — Figures et tableaux** | 19 figures (captures, tailles, état, devenir, déprédation, CPUE, profils TDR, carte des filages) et un classeur Excel de tableaux |
| **Phase 3 — Rapport** | Rapport de mission Word pré-rempli ; les éléments à compléter par l'observateur apparaissent en orange |

### Données acceptées

- **ObServe 10** : les 9 tables de la marée exportées en CSV (`export-llObservation…csv`), sélectionnées ensemble dans la fenêtre d'import (Ctrl+A ou Ctrl+clic), ou réunies dans un fichier `.zip`.
- **ObServe 9** : l'ancien export `.zip` (tableaux Excel).

L'import contrôle la sélection : fichier non reconnu, table sélectionnée deux fois, fichiers de plusieurs marées, exports réalisés à des moments éloignés, tables manquantes. Chaque cas est expliqué à l'utilisateur.

### Installation (Windows 10/11)

> **Recommandé :** utiliser l'archive de distribution `Observe-Tools-V3.6.5.zip` fournie par SCISEA. Elle contient l'application, le référentiel ObServe et le manuel d'utilisation (Word et PDF), et fonctionne hors ligne dès l'installation.
>
> Ce dépôt contient l'application et ses fonds de carte, mais **pas le référentiel ObServe** : l'application propose de le télécharger depuis le serveur de l'IRD au premier lancement (connexion internet nécessaire).

1. Décompresser `Observe-Tools-V3.6.5.zip` (ou télécharger ce dépôt) dans un dossier de votre choix.
2. Double-cliquer sur **`INSTALLER.bat`** (une seule fois, connexion internet nécessaire : installe Python et les bibliothèques via [uv](https://docs.astral.sh/uv/)).
3. Lancer l'application avec **`LANCER.vbs`** (fonctionne ensuite **hors ligne**).

Le manuel d'utilisation détaillé (`Manuel_Observe-Tools_V3.6.5`, Word et PDF) est fourni dans l'archive de distribution.

### Utilisation

1. Exporter les 9 tables de la marée en CSV depuis ObServe, puis les importer depuis l'onglet **Accueil**.
2. Sélectionner la zone de pêche.
3. Vérifier les données dans l'onglet **Correction des données**.
4. Corriger les erreurs dans ObServe si nécessaire, puis réimporter.
5. Exporter le rapport de vérification.
6. Générer les figures, puis le rapport de marée.

### Structure du dépôt

```
app.py                  Application principale (onglets Accueil / Correction / Figures / Rapport)
utils.py                Fonctions communes, version, lecture des exports
import_csv.py           Import des CSV ObServe 10 (traduction vers le format de l'application)
phase1_detection/       Règles de vérification, référentiel, traductions, règles de correction (Excel)
phase2_graphiques/      Calculs et figures
phase3_rapport/         Génération du rapport Word (modèle anonymisé inclus)
INSTALLER.bat, LANCER.bat, LANCER.vbs, METTRE_A_JOUR.bat
pyproject.toml, uv.lock Dépendances Python
```

Interface en français.

### Licence

© 2026 SCISEA — Tous droits réservés. Le code est rendu consultable pour information et évaluation ; toute reproduction, modification, distribution ou utilisation nécessite l'autorisation écrite préalable de SCISEA. Voir [LICENSE](LICENSE).

### Données tierces

- Côtes : [Natural Earth](https://www.naturalearthdata.com/) (domaine public).
- Zones économiques exclusives : Flanders Marine Institute (2023). *Maritime Boundaries Geodatabase, version 12*. [Marine Regions](https://www.marineregions.org/) — licence CC BY 4.0.
- Référentiel ObServe : IRD, Observatoire des Écosystèmes Pélagiques Tropicaux Exploités (OB7).

### Contact

Nicolas Guillon — SCISEA — nicolas.guillon@scisea.fr

---

## English

Observe-Tools is a desktop application for **scientific observers on board longliners**. From the data recorded in **ObServe**, it checks data quality, produces the trip's figures and tables, and generates the trip report as a Word document.

Developed by **SCISEA** for **IRD (UMR MARBEC, OB7)**, within the European **DCF** programme (*Data Collection Framework*).

### Features

| Step | Purpose |
|---|---|
| **Home** | Import of ObServe data, trip summary, fishing area selection, trip folder creation |
| **Phase 1 — Data checking** | Automatic checks of the 9 tables (consistency rules, value ranges, size limits per species…), annotation of anomalies (*Validated* / *Justified*), export of a protected Excel verification report |
| **Phase 2 — Figures and tables** | 19 figures (catches, sizes, condition, fate, depredation, CPUE, TDR profiles, set map) and an Excel workbook of tables |
| **Phase 3 — Report** | Pre-filled Word trip report; items to be completed by the observer are shown in orange |

### Supported data

- **ObServe 10**: the trip's 9 tables exported as CSV (`export-llObservation…csv`), selected together in the import dialog (Ctrl+A or Ctrl+click), or gathered in a `.zip` file.
- **ObServe 9**: the former `.zip` export (Excel tables).

The import checks the selection (unrecognised file, table selected twice, files from several trips, exports made far apart in time, missing tables) and explains each case to the user.

### Installation (Windows 10/11)

> **Recommended:** use the distribution archive `Observe-Tools-V3.6.5.zip` provided by SCISEA. It contains the application, the ObServe reference data and the user manual (Word and PDF), and works offline right after installation.
>
> This repository contains the application and its base maps, but **not the ObServe reference data**: the application offers to download it from the IRD server on first launch (internet connection required).

1. Unzip `Observe-Tools-V3.6.5.zip` (or download this repository) into a folder of your choice.
2. Double-click **`INSTALLER.bat`** (once, internet connection required: installs Python and libraries via [uv](https://docs.astral.sh/uv/)).
3. Start the application with **`LANCER.vbs`** (then works **offline**).

The detailed user manual (`Manuel_Observe-Tools_V3.6.5`, Word and PDF, in French) is included in the distribution archive.

### Usage

1. Export the trip's 9 tables as CSV from ObServe, then import them from the **Home** tab.
2. Select the fishing area.
3. Check the data in the **Data checking** tab.
4. Fix errors in ObServe if needed, then re-import.
5. Export the verification report.
6. Generate the figures, then the trip report.

### Repository structure

```
app.py                  Main application (Home / Checking / Figures / Report tabs)
utils.py                Shared functions, version, export reading
import_csv.py           ObServe 10 CSV import (conversion to the application format)
phase1_detection/       Checking rules, reference data, translations, correction rules (Excel)
phase2_graphiques/      Calculations and figures
phase3_rapport/         Word report generation (anonymised template included)
INSTALLER.bat, LANCER.bat, LANCER.vbs, METTRE_A_JOUR.bat
pyproject.toml, uv.lock Python dependencies
```

User interface in French.

### License

© 2026 SCISEA — All rights reserved. The source code is made available for information and evaluation purposes only; any reproduction, modification, distribution or use requires the prior written permission of SCISEA. See [LICENSE](LICENSE).

### Third-party data

- Coastlines: [Natural Earth](https://www.naturalearthdata.com/) (public domain).
- Exclusive economic zones: Flanders Marine Institute (2023). *Maritime Boundaries Geodatabase, version 12*. [Marine Regions](https://www.marineregions.org/) — CC BY 4.0 licence.
- ObServe reference data: IRD, Observatory of Exploited Tropical Pelagic Ecosystems (OB7).

### Contact

Nicolas Guillon — SCISEA — nicolas.guillon@scisea.fr
