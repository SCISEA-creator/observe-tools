#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Import des exports CSV d'ObServe (une table par fichier).
IRD OB7 - SCISEA

Depuis la nouvelle version d'ObServe, l'export zip complet n'existe plus :
chaque table est exportée séparément en CSV (séparateur ';', UTF-8), avec des
noms de colonnes EN FRANÇAIS ('Code FAO', 'Devenir'...).

Ce module transforme ces CSV pour qu'ils ressemblent exactement aux anciennes
tables Excel du zip : mêmes noms techniques de colonnes ('fao_code',
'catch_fate'...), mêmes types de valeurs (booléens, dates, nombres, groupes
d'espèces en anglais). Ainsi, toutes les règles de la Phase 1, les figures de
la Phase 2 et le rapport de la Phase 3 fonctionnent sans aucune modification.
"""

import os
import re
import glob
from datetime import datetime, timedelta
import pandas as pd


# Mot-clé présent dans le nom du fichier exporté → nom de table interne.
# Ex : 'export-llObservationCatches--2026-10-18--11-24-32.csv' → '8_catch'
FICHIERS_CSV = {
    'llObservationTrips':       '1_trip',
    'llObservationSets':        '2_set',
    'llObservationFloatlines':  '3_floatline',
    'llObservationBranchlines': '4_branchline',
    'llObservationHooks':       '5_hook',
    'llObservationBaits':       '6_bait',
    'llObservationTdrs':        '7_tdr',
    'llObservationCatches':     '8_catch',
    'llObservationSamples':     '9_sample',
}

# Correspondance libellé français (CSV) → nom technique (ancien zip), dans
# l'ordre exact des colonnes. Établie en comparant un export CSV récent et un
# ancien export zip d'ObServe : même nombre de colonnes, même ordre.
COLONNES_CSV = {
    '1_trip': [
        ('Id marée', 'trip_id'),
        ('Type', 'trip_type'),
        ("Méthode d'observation", 'observation_method'),
        ('Bateau', 'vessel'),
        ('Pavillon', 'vessel_flag'),
        ('Date de début', 'trip_start_date'),
        ('Date de fin', 'trip_end_date'),
        ('Port de départ', 'departure_port'),
        ('Port de débarquement', 'landing_port'),
        ("Nom de l'observateur", 'observer_lastname'),
        ("Prénom de l'observateur", 'observer_firstname'),
        ('Nom du capitaine', 'captain_lastname'),
        ('Prénom du capitaine', 'captain_firstname'),
        ('Commentaire', 'trip_comment'),
        ('Nom du programme', 'program'),
        ('Id interne marée', 'trip_topiaid'),
    ],
    '2_set': [
        ('Identifiant marée', 'trip_id'),
        ('Identifiant  coup', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date début de filage', 'setting_start_timestamp'),
        ('Latitude début de filage', 'setting_start_latitude'),
        ('Longitude début de filage', 'setting_start_longitude'),
        ("Date fin de mise à l'eau", 'setting_end_timestamp'),
        ('Latitude fin de filage', 'setting_end_latitude'),
        ('Longitude fin de filage', 'setting_end_longitude'),
        ('Vitesse de filage', 'setting_speed'),
        ('Lanceur utilisé', 'shooter_used'),
        ('Vitesse du shooter', 'shooter_speed'),
        ('Émerillon lesté', 'weighted_swivel'),
        ('Poids émerillon', 'swivel_weight'),
        ('Nombre de lightsticks par panier', 'lightsticks_per_basket_count'),
        ("Nombre d'avançons par panier", 'branchlines_per_basket_count'),
        ('Nombre de paniers par section', 'baskets_per_section_count'),
        ('Nombre de sections', 'total_sections_count'),
        ('Profondeur maximale visée', 'max_depth_targeted'),
        ('Direction de virage identique', 'hauling_direction_same_as_setting'),
        ('Date début de virage', 'hauling_start_timestamp'),
        ('Latitude début de virage', 'hauling_start_latitude'),
        ('Longitude début de virage', 'hauling_start_longitude'),
        ('Date fin de virage', 'hauling_end_timestamp'),
        ('Latitude fin de virage', 'hauling_end_latitude'),
        ('Longitude fin de virage', 'hauling_end_longitude'),
        ('Nombre de coupures', 'hauling_breaks'),
        ('Intervalle entre hameçons', 'time_between_hooks'),
        ('Snap lesté', 'weighted_snap'),
        ('Poids du snap', 'snap_weight'),
        ('Nombre de paniers', 'total_baskets_count'),
        ("Nombre d'hameçons", 'total_hooks_count'),
        ('Forme de la ligne', 'setting_shape'),
        ('Type de ligne mère', 'mainline_type'),
        ('Type de lightsticks', 'lightstick_type'),
        ('Couleur lightsticks', 'lightstick_color'),
        ('Ligne instrumentée', 'line_instrumented'),
        ('SST', 'sst'),
        ('Commentaire', 'set_comment'),
        ('Identifiant interne coup', 'set_topiaid'),
        ('Identifiant interne marée', 'trip_topiaid'),
    ],
    '3_floatline': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date de début de filage', 'setting_start_timestamp'),
        ('Proportion de ligne flottante', 'floatline_proportion'),
        ('Longueur de ligne flottante (m)', 'floatline_length_m'),
        ('Type de ligne', 'floatline_line_type'),
        ('Identifiant interne coup', 'set_topiaid'),
        ('Identifiant interne marée', 'trip_topiaid'),
    ],
    '4_branchline': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date de début de filage', 'setting_start_timestamp'),
        ("Proportion d'avançons", 'branchline_proportion'),
        ('Longueur des avançons (m)', 'branchline_length_m'),
        ('Type de ligne supérieure des avançons', 'branchline_topline_type'),
        ('Type de ligne terminale des avançons', 'branchline_traceline_type'),
        ('Identifiant technique calée', 'set_topiaid'),
        ('Identifiant technique marée', 'trip_topiaid'),
    ],
    '5_hook': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date début filage', 'setting_start_timestamp'),
        ('Proportion d’hameçons', 'hook_proportion'),
        ('Code d’hameçon', 'hook_code'),
        ('Type d’hameçon', 'hook_type'),
        ('Identifiant technique calée', 'set_topiaid'),
        ('Identifiant technique marée', 'trip_topiaid'),
    ],
    '6_bait': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date début filage', 'setting_start_timestamp'),
        ('Proportion appât', 'bait_proportion'),
        ('Code appât', 'bait_code'),
        ('Libellé appât', 'bait_label'),
        ('Statut filage appât', 'bait_setting_status'),
        ('Taille appât', 'bait_size'),
        ('Poids appât', 'bait_weight'),
        ('Identifiant technique calée', 'set_topiaid'),
        ('Identifiant technique marée', 'trip_topiaid'),
    ],
    '7_tdr': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Id TDR', 'tdr_id'),
        ('Numéro de série TDR', 'tdr_serial'),
        ('Marque capteur', 'sensor_brand'),
        ('Début déploiement', 'deployment_start'),
        ('Fin déploiement', 'deployment_end'),
        ('Début pêche', 'fishing_start'),
        ('Fin pêche', 'fishing_end'),
        ('Profondeur début pêche', 'fishing_start_depth'),
        ('Profondeur fin pêche', 'fishing_end_depth'),
        ('Profondeur moyenne déploiement', 'mean_deployment_depth'),
        ('Profondeur médiane déploiement', 'median_deployment_depth'),
        ('Profondeur minimale pêche', 'min_fishing_depth'),
        ('Profondeur maximale pêche', 'max_fishing_depth'),
        ('Profondeur moyenne pêche', 'mean_fishing_depth'),
        ('Profondeur médiane pêche', 'median_fishing_depth'),
        ('Position horizontale item', 'item_horizontal_position'),
        ('Position verticale item', 'item_vertical_position'),
        ('Longueur ligne flottante 1', 'floatline1_length'),
        ('Longueur ligne flottante 2', 'floatline2_length'),
        ('Id section au filage', 'section_setting_id'),
        ('Id section au virage', 'section_hauling_id'),
        ('Id filage panier', 'basket_setting_id'),
        ('Id virage panier', 'basket_hauling_id'),
        ('Id filage avançon', 'branchline_setting_id'),
        ('Id virage avançon', 'branchline_hauling_id'),
        ('Emplacement données', 'data_location'),
        ('Nom fichier données', 'data_file_name'),
        ('Code FAO', 'fao_code'),
        ('Identifiant technique TDR', 'tdr_topiaid'),
    ],
    '8_catch': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date capture', 'catch_date'),
        ('Numéro capture', 'catch_number'),
        ('Code FAO', 'fao_code'),
        ('Nom scientifique', 'scientific_name'),
        ('Nom commun', 'nom_commun'),
        ('Groupe espèce', 'species_group'),
        ('Nombre', 'count'),
        ('Poids total', 'total_weight'),
        ('Statut à la capture', 'status_at_capture'),
        ('Devenir', 'catch_fate'),
        ('Statut à la remise à l’eau', 'status_at_release'),
        ('Hameçon au rejet', 'hook_when_discarded'),
        ('Prédaté', 'depredated'),
        ('Espèce prédatrice', 'depredating_species'),
        ('Commentaire', 'catch_comment'),
        ('Id technique capture', 'catch_topiaid'),
        ('Id technique calée', 'set_topiaid'),
    ],
    '9_sample': [
        ('Id marée', 'trip_id'),
        ('Id calée', 'set_id'),
        ('Numéro de calée', 'set_number'),
        ('Date capture', 'catch_date'),
        ('Numéro capture', 'catch_number'),
        ('Code FAO', 'fao_code'),
        ('Nom scientifique', 'scientific_name'),
        ('Nom commun', 'nom_commun'),
        ('Groupe espèce', 'species_group'),
        ('Taille', 'size'),
        ('Code taille', 'size_code'),
        ('Poids', 'weight'),
        ('Code poids', 'weight_code'),
        ('Sexe', 'sex'),
        ('Maturité', 'maturity'),
        ('Commentaire', 'catch_comment'),
        ('Id technique capture', 'catch_topiaid'),
        ('Id technique calée', 'set_topiaid'),
    ],
}

# Valeurs booléennes reconnues : ObServe en français ou en anglais, ou fichier ré-enregistré par Excel
VRAI = {'Oui', 'Yes', 'True', 'true', 'TRUE', 'VRAI', 'Vrai'}
FAUX = {'Non', 'No', 'False', 'false', 'FALSE', 'FAUX', 'Faux'}

# Colonnes Oui/Non → booléen Python (comme dans l'ancien zip Excel)
COLONNES_BOOLEENNES = {
    'shooter_used', 'weighted_swivel', 'hauling_direction_same_as_setting',
    'weighted_snap', 'line_instrumented', 'depredated',
}

# hook_when_discarded était codé 1.0 / 0.0 dans l'ancien zip (et non True/False)
COLONNES_BOOL_FLOAT = {'hook_when_discarded'}

# Colonnes de dates : texte 'jj/mm/aaaa hh:mm' ou 'jj/mm/aaaa' → datetime
COLONNES_DATES = {
    'trip_start_date', 'trip_end_date',
    'setting_start_timestamp', 'setting_end_timestamp',
    'hauling_start_timestamp', 'hauling_end_timestamp',
    'deployment_start', 'deployment_end', 'fishing_start', 'fishing_end',
    'catch_date',
}

# Groupes d'espèces : libellé français (label2 du référentiel ObServe
# SpeciesGroup) → libellé anglais (label1), utilisé partout en Phase 2 et 3.
GROUPES_ESPECES_FR_EN = {
    'Poissons porte épée':    'Billfishes',
    'Oiseaux marins':         'Seabirds',
    'Raies':                  'Rays',
    'Requin-baleine':         'Whale shark',
    'Requins':                'Sharks',
    'Autres poissons osseux': 'Other bony fishes',
    'Tortues':                'Turtles',
    'Cétacés':                'Cetaceans',
    'Cephalopodes':           'Cephalopods',
    'Thonidés nca':           'Tunas nei',
    'Inconnu':                'Unknown',
}

# Valeurs qu'ObServe écrit pour « pas de valeur »
VALEURS_VIDES = {'', '@Null@'}


def _normaliser(libelle):
    """Compare les libellés sans se soucier des espaces doubles ni des apostrophes ’ / '."""
    return ' '.join(str(libelle).replace('’', "'").split()).lower()


def _renommer_colonnes(df, nom_table):
    """
    Remplace les libellés français par les noms techniques.
    1) par le libellé (robuste si ObServe change l'ordre des colonnes) ;
    2) à défaut, par la position si le nombre de colonnes est inchangé
       (robuste si ObServe retouche un libellé).
    Une colonne inconnue garde son libellé : l'app continue de fonctionner,
    seules les règles qui la concernent sont ignorées.
    """
    attendu = COLONNES_CSV[nom_table]
    par_libelle = {_normaliser(fr): tech for fr, tech in attendu}
    meme_nb = len(df.columns) == len(attendu)

    noms = []
    for i, col in enumerate(df.columns):
        tech = par_libelle.get(_normaliser(col))
        if tech is None and meme_nb:
            tech = attendu[i][1]
        noms.append(tech or col)
    df.columns = noms
    return df


def _convertir_dates(serie):
    """Texte 'jj/mm/aaaa hh:mm' (avec ou sans secondes, ou sans heure) → datetime ; illisible → NaT."""
    resultat = pd.to_datetime(serie, format='%d/%m/%Y %H:%M', errors='coerce')
    for fmt in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y'):
        resultat = resultat.fillna(pd.to_datetime(serie, format=fmt, errors='coerce'))
    return resultat


def _convertir_valeurs(df):
    """Ramène chaque colonne au type qu'elle avait dans l'ancien zip Excel."""
    df = df.apply(lambda s: s.where(~s.isin(VALEURS_VIDES)))   # '' et @Null@ → NaN

    for col in df.columns:
        s = df[col]
        if col in COLONNES_BOOLEENNES:
            df[col] = s.map({**{v: True for v in VRAI}, **{v: False for v in FAUX}})
        elif col in COLONNES_BOOL_FLOAT:
            df[col] = s.map({**{v: 1.0 for v in VRAI}, **{v: 0.0 for v in FAUX}})
        elif col in COLONNES_DATES:
            df[col] = _convertir_dates(s)
        elif col == 'species_group':
            df[col] = s.replace(GROUPES_ESPECES_FR_EN)
        else:
            # Colonne entièrement numérique → nombre (comme pandas.read_excel) ;
            # sinon on garde le texte (trip_id, codes, commentaires...).
            # Regex stricte : '5E27' (n° de série TDR) ne doit pas devenir 5×10^27.
            # La virgule décimale (fichier ré-enregistré par Excel FR) est acceptée.
            remplies = s.dropna()
            if not remplies.empty and remplies.str.fullmatch(r'-?\d+([.,]\d+)?').all():
                df[col] = pd.to_numeric(s.str.replace(',', '.', regex=False))
    return df


# Nom lisible de chaque table, pour les messages destinés à l'utilisateur
NOMS_TABLES = {
    '1_trip': 'Marée', '2_set': 'Filages', '3_floatline': 'Orins', '4_branchline': 'Avançons',
    '5_hook': 'Hameçons', '6_bait': 'Appâts', '7_tdr': 'TDR', '8_catch': 'Captures',
    '9_sample': 'Mensurations',
}

# Écart maximal toléré entre les exports des différentes tables d'une même marée
ECART_MAX_EXPORTS = timedelta(minutes=30)


def table_du_fichier(nom_fichier):
    """'export-llObservationCatches--….csv' → '8_catch' ; None si ce n'est pas une table ObServe."""
    base = os.path.basename(nom_fichier)
    return next((t for mot, t in FICHIERS_CSV.items() if mot in base), None)


def horodatage_du_fichier(nom_fichier):
    """
    Date/heure d'export écrite par ObServe dans le nom du fichier.
    ObServe n'ajoute pas de zéro devant les heures ('--8-38-40') : on lit les nombres.
    Retourne None si le nom ne contient pas d'horodatage (fichier renommé).
    """
    m = re.search(r'--(\d{4})-(\d{1,2})-(\d{1,2})--(\d{1,2})-(\d{1,2})-(\d{1,2})',
                  os.path.basename(nom_fichier))
    if not m:
        return None
    try:
        return datetime(*(int(x) for x in m.groups()))
    except ValueError:
        return None


def verifier_horodatages(noms_fichiers):
    """
    Les tables d'une même marée sont exportées à la suite : leurs horodatages doivent être
    proches. Retourne None si tout va bien, sinon un message d'avertissement à afficher
    (l'utilisateur peut choisir de continuer : il a pu réexporter volontairement une table).
    """
    dates = sorted((h, table_du_fichier(n)) for n in noms_fichiers
                   if (h := horodatage_du_fichier(n)) and table_du_fichier(n))
    if len(dates) < 2 or dates[-1][0] - dates[0][0] <= ECART_MAX_EXPORTS:
        return None
    lignes = '\n'.join(f"  - {NOMS_TABLES[t]} : {h.strftime('%d/%m/%Y à %Hh%M')}" for h, t in dates)
    return ("Les fichiers sélectionnés n'ont pas été exportés en même temps :\n\n"
            f"{lignes}\n\n"
            "Vous avez peut-être mélangé deux exports différents de la même marée "
            "(par exemple un ancien export des captures avec un nouvel export des filages).\n\n"
            "Continuer quand même ?")


def _lire_un_csv(chemin):
    nom_fichier = os.path.basename(chemin)
    try:
        try:
            return pd.read_csv(chemin, sep=';', encoding='utf-8-sig', dtype=str, keep_default_na=False)
        except UnicodeDecodeError:
            # Fichier ré-enregistré par Excel sous Windows : encodage cp1252
            return pd.read_csv(chemin, sep=';', encoding='cp1252', dtype=str, keep_default_na=False)
    except PermissionError:
        raise ValueError(f"Impossible d'ouvrir « {nom_fichier} ».\n\n"
                         "Le fichier est probablement ouvert dans Excel : fermez-le, puis recommencez.")
    except Exception as e:
        raise ValueError(f"Impossible de lire « {nom_fichier} » :\n{e}\n\n"
                         "Vérifiez qu'il s'agit bien d'un export CSV d'ObServe, non modifié.")


def lire_fichiers_csv(chemins):
    """
    Lit les CSV ObServe choisis par l'utilisateur et retourne un dict
    { 'nom_table': DataFrame } identique à celui de lire_zip().

    L'app ne devine rien : elle lit exactement les fichiers sélectionnés.
    Lève ValueError avec un message qui explique l'erreur et comment la corriger :
      - fichier qui n'est pas une table ObServe ;
      - même table sélectionnée plusieurs fois ;
      - fichiers venant de marées différentes.
    """
    # ── 1. Chaque fichier doit être une table ObServe reconnue ────────────────
    inconnus = [os.path.basename(c) for c in chemins if table_du_fichier(c) is None]
    if inconnus:
        raise ValueError(
            "Fichier(s) non reconnu(s) :\n  - " + "\n  - ".join(inconnus) + "\n\n"
            "Ce ne sont pas des tables exportées par ObServe : leur nom doit commencer par "
            "« export-llObservation… » (par exemple export-llObservationCatches--….csv).\n"
            "Retirez-les de la sélection, et ne renommez pas les fichiers exportés.")

    # ── 2. Une seule fois chaque table ────────────────────────────────────────
    par_table = {}
    for c in chemins:
        par_table.setdefault(table_du_fichier(c), []).append(os.path.basename(c))
    doublons = {t: f for t, f in par_table.items() if len(f) > 1}
    if doublons:
        details = '\n\n'.join(f"Table « {NOMS_TABLES[t]} » sélectionnée {len(f)} fois :\n  - "
                              + "\n  - ".join(sorted(f)) for t, f in sorted(doublons.items()))
        raise ValueError(f"{details}\n\n"
                         "Gardez un seul fichier par table (normalement le plus récent), puis recommencez.")

    # ── 3. Lecture et conversion au format de l'ancien zip ────────────────────
    tables = {}
    for c in chemins:
        nom_table = table_du_fichier(c)
        tables[nom_table] = _convertir_valeurs(_renommer_colonnes(_lire_un_csv(c), nom_table))

    # ── 4. Toutes les tables doivent venir de la même marée ───────────────────
    marees = {}
    for nom_table, df in tables.items():
        if 'trip_id' in df.columns:
            for tid in df['trip_id'].dropna().astype(str).str.strip().unique():
                marees.setdefault(tid, []).append(NOMS_TABLES[nom_table])
    if len(marees) > 1:
        details = '\n'.join(f"  - {tid} : {', '.join(sorted(set(t)))}" for tid, t in sorted(marees.items()))
        raise ValueError("Les fichiers sélectionnés viennent de plusieurs marées :\n\n"
                         f"{details}\n\n"
                         "Sélectionnez uniquement les fichiers d'une seule marée, puis recommencez.")
    return tables


def lire_dossier_csv(dossier):
    """Lit tous les CSV ObServe d'un dossier (mêmes contrôles que pour une sélection de fichiers)."""
    return lire_fichiers_csv([c for c in sorted(glob.glob(os.path.join(dossier, '*.csv')))
                              if table_du_fichier(c)])
