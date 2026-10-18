#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Phase 2 : Génération des figures et tableaux
IRD OB7 - SCISEA

Ce programme génère automatiquement les figures et tableaux statistiques
à partir d'un fichier export ObServe (.zip ou .xlsx).
Les sorties (PNG + Excel) sont sauvegardées dans un dossier choisi par l'utilisateur.

Lancer : python phase2_graphiques/main.py
"""

# ─────────────────────────────────────────────────────────────────────────────
# INSTALLATION AUTOMATIQUE DE MATPLOTLIB SI ABSENT
# ─────────────────────────────────────────────────────────────────────────────

import subprocess
import sys

def _installer_si_absent(nom_paquet, nom_import=None):
    """Installe un paquet via pip s'il n'est pas déjà disponible."""
    if nom_import is None:
        nom_import = nom_paquet
    try:
        __import__(nom_import)
    except ImportError:
        print(f"Installation de {nom_paquet}...")
        subprocess.check_call(
            [sys.executable, '-m', 'pip', 'install', nom_paquet, '--quiet'],
            stdout=subprocess.DEVNULL
        )

_installer_si_absent('matplotlib')
_installer_si_absent('openpyxl')
_installer_si_absent('Pillow', 'PIL')

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────

import os
import shutil
import zipfile
import tempfile
import threading

# utils.py est à la racine du projet — on s'assure qu'elle est dans sys.path
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from utils import lire_zip, lire_fichier, COULEURS, POLICE, appliquer_hover, ajouter_logos_bandeau, ajouter_titre_bandeau

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from PIL import Image, ImageTk

# Modules locaux Phase 2
import calculs
import graphiques as gfx

# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# ORDRE D'AFFICHAGE — correspond à l'ordre des sections du rapport IRD
# Chaque entrée : (type, label_affiche, reference)
#   type 'png'   → reference = nom du fichier PNG généré
#   type 'excel' → reference = nom de la feuille dans le fichier Excel
# ─────────────────────────────────────────────────────────────────────────────

ORDRE_RAPPORT = [
    # § 2.2 — Description de l'engin de pêche
    ('excel', "§ 2.2  Engin de pêche",             '2.2 Engin'),
    ('excel', "§ 2.2  Orins — schéma palangre",    '2.2 Orins'),
    ('excel', "§ 2.2  Hameçons",                   '2.2 Hameçons'),
    ('excel', "§ 2.2  Appâts et attractants",      '2.2 Appâts'),
    # § 2.4 — Carte de la route et des zones de pêche
    ('png',   "§ 2.4  Carte des opérations",       'fig_19_carte_filages.png'),
    # § 2.5 — Caractéristiques des opérations de filage
    ('excel', "§ 2.5  Opérations de filage",       '2.5 Filage'),
    # § 2.6 — Caractéristiques des opérations de virage
    ('excel', "§ 2.6  Opérations de virage",       '2.6 Virage'),
    # § 3.1 — Conditions de pêche de la ligne (profil TDR)
    ('png',   "§ 3.1  Profil de profondeur TDR",   'fig_18_profil_tdr.png'),
    # § 3.2 — Profils des profondeurs TDR (données)
    ('excel', "§ 3.2  Données TDR",                '3.2 TDR'),
    # § 4.1 — Récapitulatif des captures
    ('excel', "§ 4.1  Tableau récapitulatif",       '4.1 Captures'),
    ('png',   "§ 4.1  Captures — tous groupes",    'fig_01_captures_groupes.png'),
    ('png',   "§ 4.1  Captures — porte-épées",     'fig_02_porte_epees.png'),
    ('png',   "§ 4.1  Captures — thonidés",        'fig_03_thonides.png'),
    ('png',   "§ 4.1  Captures — requins",         'fig_04_requins.png'),
    ('png',   "§ 4.1  Captures — poissons osseux", 'fig_05_autres_poissons.png'),
    ('png',   "§ 4.1  Captures — raies",           'fig_06_raies.png'),
    ('png',   "§ 4.1  Captures — tortues",         'fig_07_tortues.png'),
    ('png',   "§ 4.1  Captures — cétacés",         'fig_08_cetaces.png'),
    # § 4.2 — Classes de tailles
    ('png',   "§ 4.2  Tailles — espadon (SWO)",    'fig_09_taille_swo.png'),
    ('png',   "§ 4.2  Tailles — thonidés",         'fig_10_taille_thons.png'),
    # § 4.3 — État à la capture
    ('png',   "§ 4.3  État à la capture",          'fig_11_etat_capture.png'),
    # § 4.4 — Survie des requins
    ('png',   "§ 4.4  Survie des requins",         'fig_12_survie_requins.png'),
    # § 4.5 — Devenir des captures
    ('png',   "§ 4.5  Devenir — global",           'fig_13_devenir_global.png'),
    ('png',   "§ 4.5  Devenir — par espèce",       'fig_14_devenir_especes.png'),
    # § 4.6 — Déprédation
    ('png',   "§ 4.6  Déprédation — totale",       'fig_15_depredation_totale.png'),
    ('png',   "§ 4.6  Déprédation — par espèce",   'fig_16_depredation_especes.png'),
    # § 4.7 — CPUE
    ('png',   "§ 4.7  CPUE par filage",            'fig_17_cpue.png'),
    ('excel', "§ 4.7  Tableau CPUE",               '4.7 CPUE'),
]

# ─────────────────────────────────────────────────────────────────────────────
# LECTURE DU FICHIER OBSERVÉ (copie locale de la fonction Phase 1)
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT EXCEL : TABLEAUX FORMATÉS
# ─────────────────────────────────────────────────────────────────────────────

def _bord_fin():
    """Bordure fine uniforme."""
    s = Side(style='thin', color='FFB0B0B0')
    return Border(left=s, right=s, top=s, bottom=s)


def _bord_separateur_groupe():
    """Bordure avec trait épais noir en bas — séparateur entre groupes d'espèces."""
    s_fin   = Side(style='thin',   color='FFB0B0B0')
    s_epais = Side(style='medium', color='FF000000')  # noir
    return Border(left=s_fin, right=s_fin, top=s_fin, bottom=s_epais)


def _ecrire_tableau_excel(ws, df, titre, col_groupe=None):
    """
    Écrit un DataFrame dans une feuille openpyxl avec mise en forme :
    - Titre fusionné en tête
    - En-têtes sur fond bleu foncé
    - Alternance de couleurs par ligne
    - Bordure fine entre chaque ligne
    - Bordure épaisse (bleu IRD) entre les groupes d'espèces si col_groupe est fourni
    """
    fill_titre   = PatternFill('solid', fgColor='FF1F4E79')
    fill_entete  = PatternFill('solid', fgColor='FF555555')  # gris foncé
    fill_total   = PatternFill('solid', fgColor='FFD0D0D0')  # gris moyen
    fill_pair    = PatternFill('solid', fgColor='FFF0F0F0')  # gris clair
    fill_impair  = PatternFill('solid', fgColor='FFFFFFFF')  # blanc

    font_blanc  = Font(bold=True, color='FFFFFFFF', name='Calibri', size=10)
    font_titre  = Font(bold=True, color='FFFFFFFF', name='Calibri', size=12)
    font_total  = Font(bold=True, name='Calibri', size=10)
    font_normal = Font(name='Calibri', size=9)

    align_c = Alignment(horizontal='center', vertical='center', wrap_text=False)
    align_g = Alignment(horizontal='left',   vertical='center', wrap_text=False)

    # Colonnes centrées : Code FAO et toutes les colonnes numériques (Nb...)
    COLONNES_CENTREES = {
        'Code FAO', 'Nb captures', 'Nb conservés', 'Nb rejetés', 'Nb échappés',
        'Nb dépredatés', 'Nb mesurés', 'Nb vivants capture', 'Nb morts capture',
        'Nb hameçons',
    }

    nb_cols = len(df.columns)

    # Indices (0-based) des lignes qui sont la DERNIÈRE de leur groupe
    fins_groupes = set()
    if col_groupe and col_groupe in df.columns:
        valeurs_groupe = df[col_groupe].astype(str).tolist()
        for i in range(len(valeurs_groupe) - 1):
            g_actuel  = valeurs_groupe[i].strip().lower()
            g_suivant = valeurs_groupe[i + 1].strip().lower()
            if g_actuel not in ('', 'total') and g_actuel != g_suivant:
                fins_groupes.add(i)

    # ── Ligne titre ───────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1,   end_column=nb_cols)
    cell_titre = ws.cell(row=1, column=1, value=titre)
    cell_titre.fill      = fill_titre
    cell_titre.font      = font_titre
    cell_titre.alignment = align_c
    ws.row_dimensions[1].height = 22

    # ── Ligne en-têtes ────────────────────────────────────────────────────────
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=2, column=j, value=col)
        c.fill      = fill_entete
        c.font      = font_blanc
        c.alignment = align_c
        c.border    = _bord_fin()
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = 'A3'

    # ── Lignes de données ─────────────────────────────────────────────────────
    for i, (_, row) in enumerate(df.iterrows()):
        excel_row = i + 3
        est_total = (str(row.iloc[0]).strip().lower() == 'total')

        fill = fill_total if est_total else (fill_pair if i % 2 == 0 else fill_impair)
        font = font_total if est_total else font_normal
        bord = _bord_separateur_groupe() if i in fins_groupes else _bord_fin()

        for j, val in enumerate(row, start=1):
            if hasattr(val, 'item'):
                val = val.item()
            if pd.isna(val):
                val = ''

            c = ws.cell(row=excel_row, column=j, value=val)
            c.fill   = fill
            c.font   = font
            c.border = bord
            nom_col  = df.columns[j - 1]
            c.alignment = align_c if nom_col in COLONNES_CENTREES else align_g

    # ── Largeurs automatiques — toutes les valeurs visibles en entier ─────────
    for j, col in enumerate(df.columns, start=1):
        valeurs_col = [str(col)] + [str(df.iloc[i, j-1]) for i in range(len(df))]
        largeur = max(len(v) for v in valeurs_col) * 1.25 + 2
        ws.column_dimensions[get_column_letter(j)].width = max(largeur, 10)


def exporter_excel(trip_id, tabl_captures, dossier_sortie,
                   tableaux_engin=None, tableaux_rapport=None):
    """
    Crée le fichier Excel complet avec tous les tableaux du rapport.
    Mise en page uniforme sur toutes les feuilles via _ecrire_tableau_excel.
    Retourne le chemin du fichier créé.

    tableaux_engin   : dict {'engin', 'orins', 'hamecons', 'appats'}
    tableaux_rapport : dict {'filage', 'virage', 'tdr', 'cpue'}
    """
    chemin = os.path.join(dossier_sortie, f'{trip_id}_tableaux.xlsx')
    wb = openpyxl.Workbook()

    # Toutes les feuilles sont définies ici dans l'ordre logique du rapport.
    # Chaque entrée : (DataFrame, nom_feuille, titre_tableau, col_groupe)
    # col_groupe = colonne servant de séparateur de groupe (None si pas de groupe)
    # Les feuilles sont ajoutées dans l'ordre du rapport (§ 2.2, 2.5, 2.6, 3.2, 4.1, 4.7)
    feuilles = []

    # ── Description de l'engin (§ 2.2) ──────────────────────────────────────
    te = tableaux_engin or {}
    configs_engin = [
        ('engin',    '2.2 Engin',    "Caractéristiques de l'engin de pêche"),
        ('orins',    '2.2 Orins',    'Schéma de palangre — orins'),
        ('hamecons', '2.2 Hameçons', 'Hameçons'),
        ('appats',   '2.2 Appâts',   'Appâts et attractants'),
    ]
    for cle, nom_feuille, titre in configs_engin:
        df_e = te.get(cle)
        if df_e is not None and not df_e.empty:
            feuilles.append((df_e, nom_feuille, f'{titre} — {trip_id}', None))

    # ── Opérations de pêche (§ 2.5 / 2.6) ───────────────────────────────────
    tr = tableaux_rapport or {}
    if tr.get('filage') is not None and not tr['filage'].empty:
        feuilles.append((tr['filage'], '2.5 Filage',
                         f'Opérations de filage — {trip_id}', None))
    if tr.get('virage') is not None and not tr['virage'].empty:
        feuilles.append((tr['virage'], '2.6 Virage',
                         f'Opérations de virage — {trip_id}', None))

    # ── TDR (§ 3.2) ───────────────────────────────────────────────────────────
    if tr.get('tdr') is not None and not tr['tdr'].empty:
        feuilles.append((tr['tdr'], '3.2 TDR',
                         f'Données TDR — {trip_id}', None))

    # ── Captures (§ 4.1) ──────────────────────────────────────────────────────
    if tabl_captures is not None and not tabl_captures.empty:
        feuilles.append((tabl_captures, '4.1 Captures',
                         f'Récapitulatif des captures — {trip_id}', 'Groupe'))

    # ── CPUE (§ 4.7) ──────────────────────────────────────────────────────────
    if tr.get('cpue') is not None and not tr['cpue'].empty:
        feuilles.append((tr['cpue'], '4.7 CPUE',
                         f'CPUE par filage — {trip_id}', None))

    # ── Écriture dans le classeur ─────────────────────────────────────────────
    for i, (df, nom, titre, col_groupe) in enumerate(feuilles):
        ws = wb.active if i == 0 else wb.create_sheet(nom)
        ws.title = nom
        _ecrire_tableau_excel(ws, df, titre, col_groupe=col_groupe)

    wb.save(chemin)
    return chemin


# ─────────────────────────────────────────────────────────────────────────────
# GÉNÉRATION COMPLÈTE : figures + tableaux
# ─────────────────────────────────────────────────────────────────────────────

def _construire_mapping_couleurs(catch_df):
    """
    Construit un dict {nom_commun: couleur_hex} cohérent par groupe d'espèces.

    Au sein de chaque groupe, les espèces sont triées par nombre de captures
    décroissant : la plus capturée reçoit la couleur de base (la plus foncée),
    les suivantes reçoivent des dégradés progressifs vers le clair.

    Ce mapping est utilisé dans les graphiques où des espèces de groupes
    différents apparaissent ensemble (état à la capture, CPUE...).
    """
    if catch_df.empty or 'species_group' not in catch_df.columns:
        return {}

    # Nombre de captures par code FAO + groupe
    comptage = (
        catch_df.groupby(['fao_code', 'species_group'])
        .size()
        .reset_index(name='nb')
    )
    comptage['groupe_fr'] = comptage['species_group'].map(calculs.GROUPES_FR).fillna(
        comptage['species_group']
    )

    mapping = {}
    for groupe_en, grp in comptage.groupby('species_group'):
        groupe_fr  = calculs.GROUPES_FR.get(groupe_en, groupe_en)
        ordre_fixe = calculs.ORDRE_ESPECES_GROUPE.get(groupe_en, [])

        # Codes présents dans les données pour ce groupe
        codes_presents = grp.set_index('fao_code')['nb'].to_dict()

        # 1. Codes dans l'ordre fixe (uniquement s'ils apparaissent dans les données)
        codes_ordonnes = [c for c in ordre_fixe if c in codes_presents]
        # 2. Codes hors ordre fixe → triés par nb de captures décroissant
        codes_restants = sorted(
            [c for c in codes_presents if c not in ordre_fixe],
            key=lambda c: -codes_presents[c]
        )
        codes_ordonnes.extend(codes_restants)

        palette = gfx._palette_groupe(groupe_fr, len(codes_ordonnes))
        for code, couleur in zip(codes_ordonnes, palette):
            if pd.notna(code):
                mapping[code] = couleur

    return mapping


def generer_tout(tables, dossier_sortie, callback_statut=None):
    """
    Orchestre le calcul de toutes les statistiques et la génération
    de toutes les figures et tableaux.

    tables         : dict {nom_table: DataFrame} (depuis lire_fichier)
    dossier_sortie : chemin du dossier de sortie
    callback_statut: fonction(message, avancement_0_100) pour mettre à jour l'UI

    Retourne (nb_figures_generees, nb_erreurs, chemin_excel).
    """
    def prog(msg, pct=None):
        if callback_statut:
            callback_statut(msg, pct)

    # Vérifier que le fichier est au format ObServe nouveau (tables numérotées 1_trip, 2_set...)
    # L'ancien format utilise des noms comme Tab1_parametres, Tab2_operations, etc.
    TABLES_ATTENDUES = {'1_trip', '2_set', '8_catch'}
    tables_presentes = set(tables.keys())
    if not TABLES_ATTENDUES.intersection(tables_presentes):
        prog("⚠ Format de fichier non reconnu — ce fichier semble être un ancien export ObServe "
             "(format Tab1/Tab2). Seules les données compatibles seront générées.", 0)

    catch_df  = tables.get('8_catch',  pd.DataFrame())
    sample_df = tables.get('9_sample', pd.DataFrame())
    set_df    = tables.get('2_set',    pd.DataFrame())
    trip_df   = tables.get('1_trip',   pd.DataFrame())
    tdr_df    = tables.get('7_tdr',    pd.DataFrame())

    # Récupérer l'ID de la marée pour nommer les fichiers
    trip_id = 'MAREE'
    if not catch_df.empty and 'trip_id' in catch_df.columns:
        trip_id = str(catch_df['trip_id'].iloc[0])

    # Mapping couleurs construit une seule fois : {nom_commun: couleur_hex}
    couleurs_especes = _construire_mapping_couleurs(catch_df)

    nb_figures = 0
    nb_erreurs = 0

    def sauver_fig(fig, nom_fichier):
        """Sauvegarde une figure PNG dans le dossier de sortie."""
        nonlocal nb_figures
        chemin = os.path.join(dossier_sortie, nom_fichier)
        fig.savefig(chemin, dpi=150, bbox_inches='tight')
        gfx._fermer(fig)
        nb_figures += 1

    def tenter(label, fn, *args):
        """
        Essaie d'exécuter fn(*args). Si une erreur survient, l'enregistre
        sans bloquer la suite. Retourne le résultat ou None.
        """
        nonlocal nb_erreurs
        try:
            return fn(*args)
        except Exception as e:
            nb_erreurs += 1
            prog(f"⚠ Erreur sur {label} : {e}")
            return None

    # ── Figure 1 : captures par groupe ────────────────────────────────────────
    prog("Figure 1 : captures par groupe...", 5)
    data = tenter("calc_captures_par_groupe", calculs.calc_captures_par_groupe, catch_df)
    if data is not None and not data.empty:
        fig = tenter("fig_captures_par_groupe", gfx.fig_captures_par_groupe, data)
        if fig:
            sauver_fig(fig, 'fig_01_captures_groupes.png')

    # ── Figures 2-5 : camembert par groupe d'espèces ─────────────────────────
    # Une figure par groupe est TOUJOURS générée.
    # Si le groupe n'a aucune capture, une figure "aucune capture" est produite.
    # groupe_en     = nom anglais ObServe dans species_group
    # groupe_fr_pal = nom français pour la palette de couleurs (COULEURS_BASE)
    # nom_fichier   = nom du PNG de sortie
    # Numérotation des figures :
    # 01 = captures par groupe
    # 02-08 = camemberts par groupe (1 par groupe, toujours généré)
    # 09-12 = histogrammes de taille
    # 13 = état à la capture
    # 14 = survie requins
    # 15 = devenir global
    # 16 = devenir par espèce
    # 15 = déprédation totale
    # 16 = déprédation par espèce
    # 17 = CPUE
    # 18 = profil TDR
    # 19 = carte des positions de filage et virage

    groupes_config = [
        ('Billfishes',        'Porte-épées',    'fig_02_porte_epees.png'),
        ('Tunas nei',         'Thons',          'fig_03_thonides.png'),
        ('Sharks',            'Requins',        'fig_04_requins.png'),
        ('Other bony fishes', 'Poissons osseux','fig_05_autres_poissons.png'),
        ('Rays',              'Raies',          'fig_06_raies.png'),
        ('Turtles',           'Tortues',        'fig_07_tortues.png'),
        ('Cetaceans',         'Cétacés',        'fig_08_cetaces.png'),
    ]
    for i, (groupe_en, groupe_fr, nom_fichier) in enumerate(groupes_config):
        prog(f"Figures 2-8 : {groupe_fr.lower()}...", 10 + i * 3)

        data = tenter(f"captures_{groupe_fr}",
                      calculs.calc_captures_par_espece_groupe, catch_df, groupe_en)

        if data is not None and len(data) > 0:
            titre = (f"Proportion de chaque espèce\n"
                     f"sur l'ensemble des captures de {groupe_fr.lower()}")
            fig = tenter(f"fig_{groupe_fr}",
                         gfx.fig_captures_groupe_detail, data, titre, groupe_fr, couleurs_especes)
        else:
            fig = tenter(f"fig_vide_{groupe_fr}",
                         gfx.fig_aucune_capture, groupe_fr)

        if fig:
            sauver_fig(fig, nom_fichier)

    # ── Figure 09 : histogramme de taille SWO (espadon) ─────────────────────
    prog("Figure 09 : tailles SWO...", 40)
    if not sample_df.empty:
        data = tenter("taille_SWO", calculs.calc_classes_taille, sample_df, 'SWO')
        if data is not None:
            fig = tenter("fig_taille_SWO", gfx.fig_classe_taille,
                         data, 'SWO', couleurs_especes.get('SWO'))
            if fig:
                sauver_fig(fig, 'fig_09_taille_swo.png')

    # ── Figure 10 : histogrammes de taille thons (ALB, BET, YFT) combinés ───
    prog("Figure 10 : tailles thons...", 44)
    if not sample_df.empty:
        data_thons = {}
        for code in calculs.ESPECES_TAILLE_THONS:
            d = tenter(f"taille_{code}", calculs.calc_classes_taille, sample_df, code)
            if d is not None:
                data_thons[code] = d
        if data_thons:
            fig = tenter("fig_taille_thons",
                         gfx.fig_classes_taille_thons, data_thons, couleurs_especes)
            if fig:
                sauver_fig(fig, 'fig_10_taille_thons.png')

    # ── Figure 11 : état à la capture ────────────────────────────────────────
    prog("Figure 11 : état à la capture...", 50)
    data = tenter("etat_capture", calculs.calc_etat_capture, catch_df)
    if data is not None:
        fig = tenter("fig_etat_capture", gfx.fig_etat_capture, data)
        if fig:
            sauver_fig(fig, 'fig_11_etat_capture.png')

    # ── Figure 12 : survie requins ────────────────────────────────────────────
    prog("Figure 12 : survie requins...", 55)
    data = tenter("survie_requins", calculs.calc_survie_requins, catch_df)
    if data is not None:
        fig = tenter("fig_survie_requins", gfx.fig_survie_requins, data)
        if fig:
            sauver_fig(fig, 'fig_12_survie_requins.png')

    # ── Figure 13 : devenir global ────────────────────────────────────────────
    prog("Figure 13 : devenir global...", 60)
    data = tenter("devenir_global", calculs.calc_devenir_global, catch_df)
    if data is not None and not data.empty:
        fig = tenter("fig_devenir_global", gfx.fig_devenir_global, data)
        if fig:
            sauver_fig(fig, 'fig_13_devenir_global.png')

    # ── Figure 14 : devenir par espèce ───────────────────────────────────────
    prog("Figure 14 : devenir par espèce...", 65)
    data = tenter("devenir_espece", calculs.calc_devenir_par_espece, catch_df)
    if data is not None and not data.empty:
        fig = tenter("fig_devenir_espece", gfx.fig_devenir_par_espece, data)
        if fig:
            sauver_fig(fig, 'fig_14_devenir_especes.png')

    # ── Figure 15 : déprédation totale ───────────────────────────────────────
    prog("Figure 15 : déprédation totale...", 70)
    data = tenter("depred_totale", calculs.calc_depredation_totale, catch_df)
    if data is not None and not data.empty:
        fig = tenter("fig_depred_totale", gfx.fig_depredation_totale, data)
        if fig:
            sauver_fig(fig, 'fig_15_depredation_totale.png')

    # ── Figure 16 : déprédation par espèce ───────────────────────────────────
    prog("Figure 16 : déprédation par espèce...", 75)
    data = tenter("depred_espece", calculs.calc_depredation_par_espece, catch_df)
    if data is not None and not data.empty:
        fig = tenter("fig_depred_espece", gfx.fig_depredation_par_espece, data)
        if fig:
            sauver_fig(fig, 'fig_16_depredation_especes.png')

    # ── Figure 17 : CPUE ─────────────────────────────────────────────────────
    prog("Figure 17 : CPUE...", 80)
    data = tenter("cpue", calculs.calc_cpue, catch_df, set_df)
    if data is not None and not data.empty:
        fig = tenter("fig_cpue", gfx.fig_cpue, data, couleurs_especes)
        if fig:
            sauver_fig(fig, 'fig_17_cpue.png')

    # ── Figure 18 : profil TDR ───────────────────────────────────────────────
    prog("Figure 18 : profil TDR...", 85)
    data = tenter("profil_tdr", calculs.calc_profil_tdr, tdr_df)
    if data is not None:
        fig = tenter("fig_profil_tdr", gfx.fig_profil_tdr, data)
        if fig:
            sauver_fig(fig, 'fig_18_profil_tdr.png')

    # ── Figure 19 : carte des positions de filage et virage ──────────────────
    prog("Figure 19 : carte des positions...", 88)
    data = tenter("carte_filages", calculs.calc_carte_filages, set_df, trip_df)
    if data is not None:
        fig = tenter("fig_carte_filages", gfx.fig_carte_filages, data)
        if fig:
            sauver_fig(fig, 'fig_19_carte_filages.png')

    # ── Données tableaux — calcul pour l'Excel (pas de PNG) ──────────────────
    prog("Calcul des tableaux...", 90)
    tableaux_engin  = tenter("tableaux_engin", calculs.calc_tableaux_engin, tables)
    df_tdr_rapport  = tenter("tabl_tdr",  calculs.calc_tableau_tdr,        tdr_df, set_df)
    df_cpue_rapport = tenter("tabl_cpue", calculs.calc_tableau_cpue_rapport, catch_df, set_df)

    df_filage_rapp, df_virage_rapp = None, None
    res_filvir = tenter("filage_virage", calculs.calc_tableaux_filage_virage, set_df)
    if res_filvir is not None:
        df_filage_rapp, df_virage_rapp = res_filvir

    # ── Tableaux Excel — tous les tableaux dans un seul fichier ──────────────
    prog("Tableaux Excel...", 95)
    tabl_captures = tenter("tabl_captures", calculs.calc_tableau_captures, catch_df, sample_df)

    tableaux_rapport = {
        'filage': df_filage_rapp,
        'virage': df_virage_rapp,
        'tdr':    df_tdr_rapport,
        'cpue':   df_cpue_rapport,
    }

    chemin_excel = None
    if tabl_captures is not None:
        chemin_excel = tenter(
            "export_excel", exporter_excel,
            trip_id, tabl_captures, dossier_sortie, tableaux_engin, tableaux_rapport
        )

    prog("Terminé !", 100)
    return nb_figures, nb_erreurs, chemin_excel


# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE GRAPHIQUE
# ─────────────────────────────────────────────────────────────────────────────

class AppPhase2(tk.Frame):
    """
    Interface principale de la Phase 2 (figures et tableaux).
    Hérite de tk.Frame pour pouvoir être embarquée dans l'application unifiée
    (app.py) ou lancée seule via lancer_seul().
    """

    def __init__(self, parent, mode='standalone'):
        super().__init__(parent, bg=COULEURS['fond'])
        # mode='standalone' : fenêtre indépendante avec bouton import
        # mode='embed'      : embarqué dans app.py, import géré par la page Accueil
        self._mode = mode

        self.tables            = {}   # tables chargées depuis le fichier ObServe
        self.trip_info         = {}   # informations de marée extraites
        self._fichiers_generes = {}   # {nom_fichier: chemin_complet} après génération
        self._tk_image         = None # référence PIL (évite le garbage collection)
        self._chemin_image_courante = None

        # Dossier temporaire pour les fichiers générés (supprimé à la fermeture)
        self._dossier_temp = tempfile.mkdtemp(prefix='observe_tools_')
        # Le protocole WM_DELETE_WINDOW est enregistré par le parent (lancer_seul ou app.py)

        self._construire_interface()

    # ── Construction de l'interface ──────────────────────────────────────────

    def _construire_interface(self):

        # ── En-tête ───────────────────────────────────────────────────────────
        en_tete = tk.Frame(self, bg=COULEURS['en_tete'], height=52)
        en_tete.pack(fill='x')
        en_tete.pack_propagate(False)

        self._logo_titre = ajouter_titre_bandeau(
            en_tete, COULEURS['en_tete'], "Observe-Tools — Figures et tableaux"
        )

        self._logos = ajouter_logos_bandeau(en_tete, COULEURS['en_tete'])

        # ── Barre de boutons ──────────────────────────────────────────────────
        barre = tk.Frame(self, bg=COULEURS['fond'])
        barre.pack(fill='x', padx=12, pady=8)

        # Bouton Import — uniquement en mode standalone
        # En mode embed, l'import est géré par la page Accueil dans app.py
        if self._mode == 'standalone':
            self.btn_importer = tk.Button(
                barre, text="Importer fichier ObServe",
                command=self._importer_fichier,
                bg=COULEURS['accent'], fg='white',
                font=(POLICE, 10, 'bold'), relief='flat', padx=12,
                cursor='hand2'
            )
            self.btn_importer.pack(side='left', padx=4)
        else:
            self.btn_importer = None

        self.btn_exporter = ttk.Button(
            barre, text="Exporter Figures et tableaux",
            command=self._exporter,
            style='Amber.TButton',
            state='disabled',
            cursor='hand2'
        )
        self.btn_exporter.pack(side='left', padx=4)

        # ── Zone principale ───────────────────────────────────────────────────
        corps = tk.Frame(self, bg=COULEURS['fond'])
        corps.pack(fill='both', expand=True, padx=12, pady=4)

        # Panneau gauche — informations sur la marée
        self.frame_info = tk.LabelFrame(
            corps, text="Informations de la marée",
            bg=COULEURS['fond'], font=(POLICE, 10, 'bold'), padx=8, pady=8
        )
        self.frame_info.pack(side='left', fill='y', padx=(0, 8), pady=4)

        self.lbl_aucun = tk.Label(
            self.frame_info,
            text="Aucun fichier chargé.\nImportez un export ObServe (.zip ou .xlsx).",
            bg=COULEURS['fond'], font=(POLICE, 9), justify='left', fg='#555555'
        )
        self.lbl_aucun.pack(anchor='w', pady=4)

        self.frame_details = tk.Frame(self.frame_info, bg=COULEURS['fond'])

        # Panneau droit — liste + prévisualisation séparées par une barre glissable
        panneau_droite = tk.PanedWindow(
            corps, orient='horizontal',
            bg=COULEURS['fond'], sashwidth=5, sashrelief='flat'
        )
        panneau_droite.pack(side='left', fill='both', expand=True, pady=4)

        # ── Sous-panneau liste ────────────────────────────────────────────────
        cadre_liste_outer = tk.Frame(panneau_droite, bg=COULEURS['blanc'],
                                     relief='groove', bd=1)
        panneau_droite.add(cadre_liste_outer, minsize=180, width=240)

        tk.Label(
            cadre_liste_outer, text="Figures générées",
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 10, 'bold')
        ).pack(fill='x')

        cadre_liste = tk.Frame(cadre_liste_outer, bg=COULEURS['blanc'])
        cadre_liste.pack(fill='both', expand=True, padx=4, pady=4)

        scrollbar = tk.Scrollbar(cadre_liste)
        scrollbar.pack(side='right', fill='y')

        self.liste_figures = tk.Listbox(
            cadre_liste, bg=COULEURS['blanc'], font=(POLICE, 9),
            yscrollcommand=scrollbar.set, activestyle='dotbox',
            selectmode='browse',
            selectbackground=COULEURS['accent'], selectforeground='white',
        )
        self.liste_figures.pack(fill='both', expand=True)
        scrollbar.config(command=self.liste_figures.yview)
        self.liste_figures.bind('<<ListboxSelect>>', self._sur_selection)

        # ── Sous-panneau prévisualisation ─────────────────────────────────────
        cadre_prev_outer = tk.Frame(panneau_droite, bg=COULEURS['blanc'],
                                    relief='groove', bd=1)
        panneau_droite.add(cadre_prev_outer, minsize=300)

        tk.Label(
            cadre_prev_outer, text="Prévisualisation",
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 10, 'bold')
        ).pack(fill='x')

        self._cadre_preview = tk.Frame(cadre_prev_outer, bg=COULEURS['blanc'])
        self._cadre_preview.pack(fill='both', expand=True)

        # Message d'invite (affiché par défaut)
        self._lbl_invite = tk.Label(
            self._cadre_preview,
            text="Cliquez sur un fichier\npour le prévisualiser",
            bg=COULEURS['blanc'], fg='#999999', font=(POLICE, 10)
        )
        self._lbl_invite.pack(expand=True)

        # Canvas pour les images — s'adapte automatiquement à la taille du panneau
        self._canvas_img = tk.Canvas(
            self._cadre_preview, bg=COULEURS['blanc'], highlightthickness=0
        )
        # Redessine l'image quand le panneau est redimensionné
        self._canvas_img.bind('<Configure>', self._redimensionner_image)

        # Frame tableau (caché par défaut)
        self._frame_tableau = tk.Frame(self._cadre_preview, bg=COULEURS['blanc'])

        # Chemin de l'image actuellement affichée (pour le redimensionnement)
        self._chemin_image_courante = None

        # ── Barre de statut ───────────────────────────────────────────────────
        bas = tk.Frame(self, bg=COULEURS['en_tete'], height=26)
        bas.pack(fill='x', side='bottom')
        bas.pack_propagate(False)

        self.var_statut = tk.StringVar(value="Importer un fichier ObServe pour commencer.")
        self._lbl_dot = tk.Label(bas, text='●', bg=COULEURS['en_tete'], fg='#94A3B8', font=(POLICE, 9))
        self._lbl_dot.pack(side='left', padx=(10, 2))
        tk.Label(
            bas, textvariable=self.var_statut,
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 9), anchor='w'
        ).pack(side='left', padx=(0, 10), fill='y')

        def _maj_dot(*args):
            msg = self.var_statut.get()
            m = msg.lower()
            if '✗' in msg or 'erreur' in m or 'impossible' in m:
                fg = '#F87171'
            elif '⚠' in msg or 'avertissement' in m:
                fg = '#FBB76A'
            elif 'cours' in m or 'génération' in m:
                fg = '#93C5FD'
            elif '✓' in msg or 'chargé' in m or 'exporté' in m or 'prêt' in m or 'terminé' in m:
                fg = '#4ADE80'
            else:
                fg = '#94A3B8'
            self._lbl_dot.config(fg=fg)
        self.var_statut.trace_add('write', _maj_dot)

        # Dossier de sortie (mémorisé)
        self._dossier_sortie = None

    # ── Gestion des actions ───────────────────────────────────────────────────

    def _importer_fichier(self):
        """Ouvre une boîte de dialogue et charge le fichier sélectionné."""
        chemin = filedialog.askopenfilename(
            title='Choisir un export ObServe',
            filetypes=[
                ('Fichiers ObServe', ('*.zip', '*.xlsx', '*.xls')),
                ('ZIP', '*.zip'),
                ('Excel', '*.xlsx *.xls'),
                ('Tous', '*.*'),
            ]
        )
        if not chemin:
            return

        try:
            tables = lire_fichier(chemin)
        except Exception as e:
            messagebox.showerror(
                "Erreur de lecture",
                f"Impossible de lire ce fichier :\n\n{e}"
            )
            return
        if not tables:
            messagebox.showerror("Erreur", "Impossible de lire ce fichier.")
            return

        self.tables = tables
        self._afficher_resume()
        self.var_statut.set(f"Fichier chargé : {os.path.basename(chemin)}")
        # Lancer la génération automatiquement dans le dossier temporaire
        self._lancer_generation()

    def _recevoir_tables(self, tables, nom_fichier):
        """
        Appelé par la page Accueil (app.py) quand un fichier est importé.
        Équivalent de _importer_fichier mais sans boîte de dialogue.
        """
        self.tables = tables
        self._afficher_resume()
        self.var_statut.set(f"Fichier chargé : {nom_fichier}")
        self._lancer_generation()

    def _afficher_resume(self):
        """Affiche les informations clés de la marée dans le panneau gauche."""
        # Vider les widgets précédents
        for w in self.frame_details.winfo_children():
            w.destroy()
        self.lbl_aucun.pack_forget()

        catch_df  = self.tables.get('8_catch',  pd.DataFrame())
        sample_df = self.tables.get('9_sample', pd.DataFrame())
        set_df    = self.tables.get('2_set',    pd.DataFrame())
        trip_df   = self.tables.get('1_trip',   pd.DataFrame())

        def ligne(label, valeur):
            f = tk.Frame(self.frame_details, bg=COULEURS['fond'])
            f.pack(fill='x', pady=1)
            tk.Label(f, text=label + ' :', width=14, anchor='w',
                     bg=COULEURS['fond'], font=(POLICE, 9, 'bold')).pack(side='left')
            tk.Label(f, text=str(valeur), anchor='w',
                     bg=COULEURS['fond'], font=(POLICE, 9)).pack(side='left')

        # Infos marée (depuis 1_trip si disponible, sinon depuis 8_catch)
        def _fmt_date(val):
            """Retourne la date seule (sans heure) depuis une valeur datetime ou string."""
            if val is None or str(val) in ('?', 'NaT', 'nan'):
                return '?'
            try:
                return pd.Timestamp(val).strftime('%d/%m/%Y')
            except Exception:
                return str(val).split(' ')[0]  # fallback : couper après l'espace

        # Zones de pêche — calculées ici pour pouvoir les afficher au bon endroit
        zones = calculs.calc_zones_peche(set_df)

        if not trip_df.empty:
            t = trip_df.iloc[0]
            ligne("Marée",       t.get('trip_id', '?'))
            ligne("Navire",      t.get('vessel', '?'))
            ligne("Capitaine",   f"{t.get('captain_firstname', '')} {t.get('captain_lastname', '')}".strip() or '?')
            ligne("Observateur", f"{t.get('observer_firstname', '')} {t.get('observer_lastname', '')}")
            # ZEE entre Observateur et Dates
            if zones:
                for i, z in enumerate(zones):
                    ligne("Zone pêche" if i == 0 else "", z)
            else:
                ligne("Zone pêche", "Eaux internationales")
            ligne("Dates",       f"{_fmt_date(t.get('trip_start_date'))} → {_fmt_date(t.get('trip_end_date'))}")
        elif not catch_df.empty:
            ligne("Marée", catch_df['trip_id'].iloc[0] if 'trip_id' in catch_df.columns else '?')
            if zones:
                for i, z in enumerate(zones):
                    ligne("Zone pêche" if i == 0 else "", z)
            else:
                ligne("Zone pêche", "Eaux internationales")

        ligne("Filages",  len(set_df) if not set_df.empty else '?')
        ligne("Captures", len(catch_df) if not catch_df.empty else 0)
        ligne("Mesures",  len(sample_df) if not sample_df.empty else 0)

        if not catch_df.empty and 'species_group' in catch_df.columns:
            n_groupes = catch_df['species_group'].nunique()
            n_especes = catch_df['fao_code'].nunique()
            ligne("Groupes",  n_groupes)
            ligne("Espèces",  n_especes)

        self.frame_details.pack(anchor='w')

        # Mettre à jour la liste des figures attendues
        self.liste_figures.delete(0, 'end')
        for typ, label, _ in ORDRE_RAPPORT:
            icone = '🖼' if typ == 'png' else '📊'
            self.liste_figures.insert('end', f"  {icone}  {label}")

    def _lancer_generation(self):
        """Lance la génération dans le dossier temporaire (appelé automatiquement à l'import)."""
        if not self.tables:
            return

        # Vider le dossier temporaire (en cas de re-import)
        for f in os.listdir(self._dossier_temp):
            chemin_f = os.path.join(self._dossier_temp, f)
            if os.path.isfile(chemin_f):
                os.remove(chemin_f)

        self._dossier_sortie = self._dossier_temp
        if self.btn_importer:
            self.btn_importer.config(state='disabled')
        self.btn_exporter.config(state='disabled')
        self.var_statut.set("Génération en cours...")

        # Effacer la liste et la marquer "en cours"
        self.liste_figures.delete(0, 'end')
        self.liste_figures.insert('end', "  ⏳ Génération en cours...")

        # Snapshot des tables AVANT le lancement du thread : si l'utilisateur importe
        # une nouvelle marée pendant la génération, le thread garde les données d'origine.
        tables_snapshot  = self.tables
        dossier_snapshot = self._dossier_sortie

        # Lancer dans un thread pour ne pas bloquer l'interface
        thread = threading.Thread(
            target=self._thread_generation,
            args=(tables_snapshot, dossier_snapshot),
            daemon=True
        )
        thread.start()

    def _thread_generation(self, tables, dossier_sortie):
        """Exécuté en arrière-plan : génère toutes les figures et tableaux."""

        def callback(message, pct=None):
            # Toutes les mises à jour de l'UI doivent passer par after() depuis un thread
            self.after(0, lambda m=message: self.var_statut.set(m))

        try:
            nb_figs, nb_err, chemin_excel = generer_tout(
                tables, dossier_sortie, callback
            )
            # Mise à jour finale depuis le thread principal
            self.after(0, lambda: self._generation_terminee(nb_figs, nb_err, chemin_excel))
        except Exception as e:
            self.after(0, lambda msg=str(e): self._generation_erreur(msg))

    def _generation_terminee(self, nb_figs, nb_err, chemin_excel):
        """Appelé dans le thread principal quand la génération est finie."""
        if self.btn_importer:
            self.btn_importer.config(state='normal')

        if nb_err == 0:
            self.var_statut.set(f"✓ {nb_figs} figure(s) prête(s) — cliquez pour prévisualiser")
        else:
            self.var_statut.set(
                f"✓ {nb_figs} figure(s) prête(s) ({nb_err} avertissement(s))"
            )

        # Mettre à jour la liste dans l'ordre du rapport
        self.liste_figures.delete(0, 'end')
        self._items_rapport = []  # [{type, label, ref, path, sheet}]

        # Trouver le fichier Excel généré
        chemin_excel_genere = None
        if self._dossier_sortie and os.path.isdir(self._dossier_sortie):
            for f in os.listdir(self._dossier_sortie):
                if f.endswith('.xlsx'):
                    chemin_excel_genere = os.path.join(self._dossier_sortie, f)
                    break

        # Lire les feuilles disponibles dans l'Excel
        feuilles_excel = []
        if chemin_excel_genere:
            try:
                wb = openpyxl.load_workbook(chemin_excel_genere, read_only=True)
                feuilles_excel = wb.sheetnames
                wb.close()
            except Exception:
                pass

        for typ, label, ref in ORDRE_RAPPORT:
            item = {'type': typ, 'label': label, 'ref': ref,
                    'path': None, 'sheet': None}
            disponible = False

            if typ == 'png' and self._dossier_sortie:
                chemin_png = os.path.join(self._dossier_sortie, ref)
                if os.path.isfile(chemin_png):
                    item['path'] = chemin_png
                    disponible = True
            elif typ == 'excel' and chemin_excel_genere and ref in feuilles_excel:
                item['path']  = chemin_excel_genere
                item['sheet'] = ref
                disponible = True

            self._items_rapport.append(item)
            icone = '✓' if disponible else '○'
            self.liste_figures.insert('end', f"  {icone}  {label}")

        if nb_figs > 0:
            self.btn_exporter.config(state='normal')

        if nb_err > 0:
            messagebox.showwarning(
                "Génération terminée avec avertissements",
                f"{nb_figs} figure(s) générée(s).\n"
                f"{nb_err} erreur(s) non bloquante(s) — certaines figures peuvent être absentes."
            )

    def _generation_erreur(self, message):
        """Appelé dans le thread principal si generer_tout() lève une exception."""
        if self.btn_importer:
            self.btn_importer.config(state='normal')
        self.var_statut.set("✗ Erreur lors de la génération.")
        self.liste_figures.delete(0, 'end')
        messagebox.showerror("Erreur de génération", message)

    # ── Prévisualisation ─────────────────────────────────────────────────────

    def _sur_selection(self, event):
        """Appelé quand l'utilisateur clique sur un élément de la liste."""
        sel = self.liste_figures.curselection()
        if not sel:
            return

        idx = sel[0]
        if not hasattr(self, '_items_rapport') or idx >= len(self._items_rapport):
            return

        item = self._items_rapport[idx]
        if not item.get('path'):
            return  # élément non généré (○)

        self.var_statut.set(f"Prévisualisation : {item['label']}")

        try:
            if item['type'] == 'png':
                self._previsualiser_image(item['path'])
            elif item['type'] == 'excel':
                self._previsualiser_tableau(item['path'], sheet_name=item.get('sheet'))
        except Exception as e:
            self.var_statut.set(f"Erreur : {e}")
            messagebox.showerror("Erreur de prévisualisation", str(e))

    def _masquer_preview(self):
        """Cache tous les widgets de prévisualisation."""
        self._lbl_invite.pack_forget()
        self._canvas_img.pack_forget()
        self._frame_tableau.pack_forget()

    def _previsualiser_image(self, chemin):
        """Affiche un PNG dans le canvas — le redimensionnement est géré par <Configure>."""
        self._masquer_preview()
        self._chemin_image_courante = chemin
        # Afficher le canvas avant de dessiner (il doit être rendu pour avoir une taille)
        self._canvas_img.pack(fill='both', expand=True)
        # Laisser tkinter rendre le canvas, puis dessiner l'image
        self.after(30, lambda: self._redimensionner_image(None))

    def _redimensionner_image(self, event):
        """Redessine l'image courante à la taille du canvas (appelé au resize aussi)."""
        if not self._chemin_image_courante:
            return
        w = self._canvas_img.winfo_width()
        h = self._canvas_img.winfo_height()
        if w <= 1 or h <= 1:
            # Canvas pas encore rendu — réessayer dans 50 ms
            self.after(50, lambda: self._redimensionner_image(None))
            return
        img = Image.open(self._chemin_image_courante)
        img.thumbnail((w - 4, h - 4), Image.LANCZOS)
        self._tk_image = ImageTk.PhotoImage(img)  # conserver la référence (sinon GC l'efface)
        self._canvas_img.delete('all')
        self._canvas_img.create_image(w // 2, h // 2, anchor='center', image=self._tk_image)

    def _previsualiser_tableau(self, chemin, sheet_name=None):
        """Affiche un tableau Excel dans la zone de prévisualisation.
        sheet_name : feuille à afficher par défaut (None = première feuille).
        """
        self._masquer_preview()

        # Vider le frame tableau des widgets précédents
        for widget in self._frame_tableau.winfo_children():
            widget.destroy()

        # Lire les noms de feuilles sans charger les données
        wb = openpyxl.load_workbook(chemin, read_only=True)
        feuilles = wb.sheetnames
        wb.close()

        # Feuille initiale : celle demandée si elle existe, sinon la première
        feuille_initiale = sheet_name if sheet_name and sheet_name in feuilles else feuilles[0]

        # Sélecteur de feuille
        barre = tk.Frame(self._frame_tableau, bg=COULEURS['fond'])
        barre.pack(fill='x', padx=4, pady=4)
        tk.Label(barre, text="Feuille :", bg=COULEURS['fond'],
                 font=(POLICE, 9)).pack(side='left')
        var_feuille = tk.StringVar(value=feuille_initiale)
        combo = ttk.Combobox(barre, textvariable=var_feuille,
                             values=feuilles, state='readonly', width=22)
        combo.pack(side='left', padx=6)

        # Conteneur du Treeview
        cadre_tree = tk.Frame(self._frame_tableau, bg=COULEURS['blanc'])
        cadre_tree.pack(fill='both', expand=True, padx=4, pady=(0, 4))

        def _charger(nom_feuille):
            """Charge et affiche la feuille sélectionnée dans le Treeview."""
            for w in cadre_tree.winfo_children():
                w.destroy()

            # header=1 : saute la ligne de titre fusionnée (ligne 1)
            # et utilise la ligne des en-têtes réelles (ligne 2)
            df = pd.read_excel(chemin, sheet_name=nom_feuille, engine='openpyxl', header=1)
            cols = list(df.columns)

            COLS_CENTREES = {
                'Code FAO', 'Nb captures', 'Nb conservés', 'Nb rejetés', 'Nb échappés',
                'Nb dépredatés', 'Nb mesurés', 'Nb vivants capture', 'Nb morts capture',
                'Nb hameçons',
            }

            tree = ttk.Treeview(cadre_tree, columns=cols, show='headings')

            # En-têtes en gras, colonnes centrées ou alignées à gauche
            for col in cols:
                anchor = 'center' if col in COLS_CENTREES else 'w'
                # Largeur d'après le contenu réel de la colonne
                vals_col = [str(col)] + [
                    '' if pd.isna(df[col].iloc[i]) else str(df[col].iloc[i])
                    for i in range(len(df))
                ]
                largeur = max(len(v) for v in vals_col) * 8 + 10
                largeur = max(min(largeur, 300), 50)
                tree.heading(col, text=str(col))
                tree.column(col, width=largeur, minwidth=40, anchor=anchor)

            # Style en-têtes gras via ttk.Style (appliqué localement)
            s = ttk.Style()
            s.configure('Treeview.Heading', font=(POLICE, 9, 'bold'))

            # Tags de style des lignes — gris clair
            tree.tag_configure('pair',   background='#F0F0F0')
            tree.tag_configure('impair', background='#FFFFFF')
            tree.tag_configure('total',  background='#D0D0D0',
                               font=(POLICE, 9, 'bold'))
            tree.tag_configure('sep',    background='#555555')  # ligne séparatrice fine

            col_groupe  = 'Groupe' if 'Groupe' in cols else None
            groupe_prec = None
            pair        = True

            for _, row in df.iterrows():
                vals = ['' if pd.isna(v) else str(v) for v in row]

                if vals[0].strip().lower() == 'total':
                    tree.insert('', 'end', values=vals, tags=('total',))
                    continue

                if col_groupe:
                    groupe_actuel = vals[cols.index(col_groupe)].strip()
                    groupe_prec   = groupe_actuel

                tag = 'pair' if pair else 'impair'
                pair = not pair
                tree.insert('', 'end', values=vals, tags=(tag,))

            sb_y = ttk.Scrollbar(cadre_tree, orient='vertical',   command=tree.yview)
            sb_x = ttk.Scrollbar(cadre_tree, orient='horizontal', command=tree.xview)
            tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
            sb_y.pack(side='right',  fill='y')
            sb_x.pack(side='bottom', fill='x')
            tree.pack(fill='both', expand=True)

        combo.bind('<<ComboboxSelected>>', lambda e: _charger(var_feuille.get()))
        _charger(feuille_initiale)  # charge la feuille pré-sélectionnée

        self._frame_tableau.pack(fill='both', expand=True)

    def _exporter(self):
        """Copie tous les fichiers générés (PNG + Excel) dans un dossier choisi."""
        if not self._dossier_sortie or not os.path.isdir(self._dossier_sortie):
            messagebox.showwarning("", "Aucune figure à exporter.")
            return
        dossier = filedialog.askdirectory(title="Choisir le dossier d'export")
        if not dossier:
            return
        fichiers = [f for f in os.listdir(self._dossier_sortie)
                    if f.endswith('.png') or f.endswith('.xlsx')]
        try:
            for f in fichiers:
                shutil.copy2(os.path.join(self._dossier_sortie, f),
                             os.path.join(dossier, f))
        except PermissionError as e:
            messagebox.showerror(
                "Fichier verrouillé",
                f"Impossible de copier un fichier — il est peut-être ouvert dans Excel ou un autre programme.\n\n{e}"
            )
            return

        # Compter les figures manquantes (items attendus sans fichier généré)
        nb_png = sum(1 for f in fichiers if f.endswith('.png'))
        nb_attendus = sum(1 for it in self._items_rapport if it.get('type') == 'figure')
        nb_manquants = max(0, nb_attendus - nb_png)

        self.var_statut.set(f"✓ Export terminé → {dossier}")
        msg = f"{len(fichiers)} fichier(s) exportés dans :\n{dossier}"
        if nb_manquants > 0:
            msg += f"\n\n⚠ {nb_manquants} figure(s) manquante(s) (marquées ○ dans la liste)."
        if messagebox.askyesno("Export terminé", msg + "\n\nOuvrir le dossier ?"):
            os.startfile(dossier)

    def _fermer(self):
        """Supprime le dossier temporaire et ferme la fenêtre parente."""
        shutil.rmtree(self._dossier_temp, ignore_errors=True)
        self.winfo_toplevel().destroy()


# ─────────────────────────────────────────────────────────────────────────────
# POINT D'ENTRÉE
# ─────────────────────────────────────────────────────────────────────────────

def lancer_seul():
    """Lance la Phase 2 dans sa propre fenêtre (mode standalone)."""
    root = tk.Tk()
    root.title("Observe-Tools — Génération figures et tableaux")
    root.geometry("1200x700")
    root.resizable(True, True)
    app = AppPhase2(root)
    root.protocol("WM_DELETE_WINDOW", app._fermer)
    app.pack(fill='both', expand=True)
    root.mainloop()


if __name__ == '__main__':
    lancer_seul()
