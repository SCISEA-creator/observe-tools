#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Phase 2 : Calculs statistiques
IRD OB7 - SCISEA

Ce module contient toutes les fonctions de calcul statistique pour les
tableaux et graphiques du rapport. Chaque fonction prend des DataFrames
pandas en entrée et retourne un DataFrame ou une Series.
Aucune fonction d'affichage ici — seulement des calculs purs.
"""

import os
import json

import pandas as pd
import numpy as np


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES DE RÉFÉRENCE
# ─────────────────────────────────────────────────────────────────────────────

# Noms courts IRD pour les groupes principaux.
# Ces noms sont volontairement raccourcis par rapport au référentiel ObServe
# (ex: "Thons" au lieu de "Thonidés nca") pour correspondre à l'usage IRD
# dans les rapports d'observateurs.
_NOMS_COURTS_IRD = {
    'Billfishes':        'Porte-épées',
    'Tunas nei':         'Thons',
    'Sharks':            'Requins',
    'Rays':              'Raies',
    'Other bony fishes': 'Poissons osseux',
    'Cetaceans':         'Cétacés',
    'Turtles':           'Tortues',
}


def _charger_groupes_fr():
    """
    Construit le dictionnaire {nom_anglais_ObServe: nom_français} des groupes d'espèces.

    Priorité :
    1. Noms courts IRD (_NOMS_COURTS_IRD) pour les groupes principaux.
    2. Nom français du référentiel ObServe pour tout groupe inconnu (groupes rares
       ajoutés dans une future version d'ObServe).
    3. Nom anglais tel quel si le référentiel n'est pas trouvé.

    Le référentiel est lu depuis phase1_detection/data/referentiel_observe.json.
    Si le fichier est absent ou illisible, seuls les noms courts IRD sont utilisés.
    """
    mapping = dict(_NOMS_COURTS_IRD)

    # Chemin du référentiel : remonte d'un niveau depuis ce fichier (phase2_graphiques/)
    # puis entre dans phase1_detection/data/
    racine     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    chemin_ref = os.path.join(racine, 'phase1_detection', 'data', 'referentiel_observe.json')

    if not os.path.isfile(chemin_ref):
        return mapping

    try:
        with open(chemin_ref, encoding='utf-8') as f:
            ref = json.load(f)
        cle    = 'fr.ird.observe.entities.referential.common.SpeciesGroup'
        groupes = ref.get('referentiels', {}).get(cle, [])
        for g in groupes:
            nom_en = (g.get('label1') or '').strip()
            nom_fr = (g.get('label2') or '').strip()
            # N'ajouter que si le groupe n'a pas déjà un nom court IRD
            if nom_en and nom_en not in mapping:
                mapping[nom_en] = nom_fr or nom_en
    except Exception:
        # En cas d'erreur de lecture, on continue avec les noms courts seulement
        pass

    return mapping


# Dictionnaire chargé une fois au démarrage.
# Contient tous les groupes connus : noms courts IRD + groupes rares du référentiel.
GROUPES_FR = _charger_groupes_fr()

# Traduction des codes d'espèces déprédatrices
DEPREDATEURS_FR = {
    'ISB': 'Squalelet féroce',
    'SQU': 'Calmar',
    'SKH': 'Requin',
    'ODN': 'Odontocète',
    'CEP': 'Céphalopode',
    'FAW': 'Fausse orque',
    'GLO': 'Globicéphale',
    'XXX*': 'Prédateur inconnu',   # code ObServe quand le prédateur n'a pas pu être identifié
}

# Ordre d'affichage des groupes principaux dans les tableaux.
# Les groupes rares (Oiseaux marins, Requin-baleine, etc.) non listés ici
# apparaîtront en dernier (ordre 99 dans calc_tableau_captures).
ORDRE_GROUPES = ['Porte-épées', 'Thons', 'Requins', 'Raies', 'Poissons osseux', 'Tortues', 'Cétacés']

# Ordre prioritaire des espèces dans les camemberts et figures par groupe.
# La première espèce reçoit la couleur de base (la plus foncée) du groupe.
# Les espèces présentes dans les données mais absentes de cette liste sont
# ajoutées après, triées par nombre de captures décroissant, dans le même dégradé.
ORDRE_ESPECES_GROUPE = {
    'Billfishes': ['SWO', 'BUM', 'BLM', 'MLS', 'SFA', 'SSP', 'BIL'],
    'Tunas nei':  ['YFT', 'BET', 'ALB', 'SKJ', 'TUN'],
}

# Espèces cibles retenues pour la CPUE individuelle
ESPECES_CPUE = ['SWO', 'YFT', 'BET', 'ALB']

# Espèces commerciales pour l'état à la capture et la déprédation
ESPECES_COMMERCIALES = [
    'SWO', 'BUM', 'BLM', 'MLS', 'SSP', 'SFA',
    'YFT', 'BET', 'ALB', 'SKJ',
    'WAH', 'DOL', 'LEC', 'ALX',
]

# Espèces dont on génère les histogrammes de taille (si assez de mesures)
ESPECES_TAILLE              = ['SWO', 'ALB', 'BET', 'YFT']
ESPECES_TAILLE_INDIVIDUELLES = ['SWO']           # une figure séparée par espèce
ESPECES_TAILLE_THONS         = ['ALB', 'BET', 'YFT']  # regroupées en une seule figure
TAILLE_STEP                  = 10  # largeur des classes en cm
NB_MESURES_MIN               = 3   # nombre minimal de mesures pour tracer un histogramme


# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES INTERNES
# ─────────────────────────────────────────────────────────────────────────────

def _groupe_fr(en):
    """Retourne le nom de groupe en français à partir du nom anglais ObServe."""
    return GROUPES_FR.get(en, en)


def _label_espece(row):
    """Retourne le code FAO de la capture pour l'affichage dans les figures."""
    return str(row.get('fao_code', '')).strip()


def _est_vivant_capture(status):
    """
    True si le statut à la capture indique un animal vivant (A ou AW = blessé vivant).
    """
    return status in ('A', 'AW')



# ─────────────────────────────────────────────────────────────────────────────
# 1 — TABLEAU RÉCAPITULATIF DES CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

def calc_tableau_captures(catch_df, sample_df):
    """
    Construit le tableau récapitulatif des captures par espèce.

    Colonnes : Groupe, Nom scientifique, Nom commun, Code FAO,
               Nb captures, Nb conservés, Nb rejetés, Nb échappés,
               Nb dépredatés, Nb mesurés, Nb capturés vivants, Nb capturés morts

    Retourne un DataFrame avec une ligne TOTAL en tête, puis une ligne par espèce.
    """
    if catch_df.empty:
        return pd.DataFrame()

    # Nombre de mensurations par espèce (depuis 9_sample)
    mensurations = sample_df.groupby('fao_code').size().rename('nb_mesures') \
        if not sample_df.empty else pd.Series(dtype=int, name='nb_mesures')

    grp = catch_df.groupby('fao_code')

    resultat = pd.DataFrame({
        'Groupe':             grp['species_group'].first().map(_groupe_fr),
        'Nom scientifique':   grp['scientific_name'].first(),
        'Nom commun':         grp['nom_commun'].first(),
        'Code FAO':           grp['fao_code'].first(),
        'Nb captures':        grp.size(),
        'Nb conservés':       grp.apply(lambda d: (d['catch_fate'] == 'RET').sum()),
        'Nb rejetés':         grp.apply(lambda d: (d['catch_fate'] == 'DIS').sum()),
        'Nb échappés':        grp.apply(lambda d: (d['catch_fate'] == 'ESC').sum()),
        # Nb dépredatés = captures ayant subi au moins un événement de déprédation
        # (le champ depredating_species est renseigné dans ObServe)
        'Nb dépredatés':      grp.apply(lambda d: d['depredating_species'].notna().sum()),
        'Nb vivants capture': grp.apply(lambda d: d['status_at_capture'].apply(_est_vivant_capture).sum()),
        'Nb morts capture':   grp.apply(lambda d: (d['status_at_capture'] == 'D').sum()),
    })

    # Joindre les mensurations
    resultat = resultat.join(mensurations, how='left')
    resultat['Nb mesurés'] = resultat['nb_mesures'].fillna(0).astype(int)
    resultat = resultat.drop(columns=['nb_mesures'])

    # Trier par groupe (ordre logique) puis par nb captures décroissant
    ordre = {g: i for i, g in enumerate(ORDRE_GROUPES)}
    resultat['_ordre'] = resultat['Groupe'].map(ordre).fillna(99)
    resultat = resultat.sort_values(['_ordre', 'Nb captures'], ascending=[True, False])
    resultat = resultat.drop(columns=['_ordre'])

    # Ligne total en fin de tableau
    total = {
        'Groupe': 'Total', 'Nom scientifique': '', 'Nom commun': '', 'Code FAO': '',
        'Nb captures':        resultat['Nb captures'].sum(),
        'Nb conservés':       resultat['Nb conservés'].sum(),
        'Nb rejetés':         resultat['Nb rejetés'].sum(),
        'Nb échappés':        resultat['Nb échappés'].sum(),
        'Nb dépredatés':      resultat['Nb dépredatés'].sum(),
        'Nb mesurés':         resultat['Nb mesurés'].sum(),
        'Nb vivants capture': resultat['Nb vivants capture'].sum(),
        'Nb morts capture':   resultat['Nb morts capture'].sum(),
    }
    return pd.concat([resultat, pd.DataFrame([total])], ignore_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# 2a — TABLEAUX TDR ET CPUE (PNG rapport)
# ─────────────────────────────────────────────────────────────────────────────

def calc_tableau_tdr(tdr_df, set_df):
    """
    Construit le tableau des données TDR pour le rapport.

    Une ligne par (SET, TDR) — si un TDR est associé à plusieurs captures,
    les codes FAO sont concaténés dans la colonne 'Espèce(s) cap.' (ex: 'LKV / SWO').

    Colonnes : SET | TDR ID | N° série | Orin 1 (m) | Orin 2 (m) |
               Prof. min (m) | Prof. max (m) | Prof. méd. (m) | Espèce(s) cap.
    """
    cols_requises_tdr = {'set_number', 'tdr_id'}
    cols_requises_set = {'set_number', 'set_id'}
    if (tdr_df.empty or set_df.empty
            or not cols_requises_tdr.issubset(tdr_df.columns)
            or not cols_requises_set.issubset(set_df.columns)):
        return pd.DataFrame()

    set_id_map = set_df.set_index('set_number')['set_id'].to_dict()

    cols_fixes = [
        'set_number', 'tdr_id', 'tdr_serial',
        'floatline1_length', 'floatline2_length',
        'branchline_setting_id', 'branchline_hauling_id',
        'min_fishing_depth', 'max_fishing_depth', 'median_fishing_depth',
    ]
    # Garder uniquement les colonnes présentes
    cols_fixes = [c for c in cols_fixes if c in tdr_df.columns]

    lignes = []
    for (sn, tid), g in tdr_df.sort_values(['set_number', 'tdr_id']).groupby(
            ['set_number', 'tdr_id'], sort=False):
        row = {col: g[col].iloc[0] for col in cols_fixes}
        # Concaténer les espèces capturées (plusieurs possibles si TDR associé à 2+ captures)
        especes = sorted(g['fao_code'].dropna().unique().tolist()) if 'fao_code' in g.columns else []
        row['fao_code'] = ' / '.join(especes) if especes else '—'
        row['set_id']   = set_id_map.get(sn, str(sn))
        lignes.append(row)

    df = pd.DataFrame(lignes)

    # Arrondir les profondeurs et longueurs à 1 décimale
    for col in ['floatline1_length', 'floatline2_length',
                'min_fishing_depth', 'max_fishing_depth', 'median_fishing_depth']:
        if col in df.columns:
            df[col] = df[col].round(1)

    # Sélectionner et renommer les colonnes pour l'affichage
    cols_out = ['set_id', 'tdr_id', 'tdr_serial',
                'floatline1_length', 'floatline2_length',
                'branchline_setting_id', 'branchline_hauling_id',
                'min_fishing_depth', 'max_fishing_depth', 'median_fishing_depth', 'fao_code']
    cols_out = [c for c in cols_out if c in df.columns]
    df = df[cols_out]

    noms = {
        'set_id':               'SET',
        'tdr_id':               'TDR ID',
        'tdr_serial':           'N° série',
        'floatline1_length':    'Orin 1 (m)',
        'floatline2_length':    'Orin 2 (m)',
        'branchline_setting_id':  'Pos. pose',
        'branchline_hauling_id':  'Pos. virage',
        'min_fishing_depth':      'Prof. min (m)',
        'max_fishing_depth':      'Prof. max (m)',
        'median_fishing_depth':   'Prof. méd. (m)',
        'fao_code':               'Espèce(s) cap.',
    }
    return df.rename(columns=noms)


def calc_tableau_cpue_rapport(catch_df, set_df):
    """
    Construit le tableau CPUE pour le rapport (une ligne par filage).

    Colonnes : SET | Nb hameçons | [espèce RET 1] | ... | [espèce RET N] | CPUE totale
    - Espèces : uniquement celles conservées (catch_fate == 'RET'), triées par code FAO
    - CPUE espèce = captures RET / (hameçons / 1000), arrondi à 2 décimales
    - CPUE totale = toutes captures (RET + DIS + ESC) / (hameçons / 1000)
    """
    cols_requises = {'set_number', 'set_id', 'total_hooks_count'}
    if catch_df.empty or set_df.empty or not cols_requises.issubset(set_df.columns):
        return pd.DataFrame()

    set_id_map = set_df.set_index('set_number')['set_id'].to_dict()
    hooks      = set_df.set_index('set_number')['total_hooks_count']

    ret         = catch_df[catch_df['catch_fate'] == 'RET']
    especes_ret = sorted(ret['fao_code'].dropna().unique().tolist())

    lignes = []
    for sn in sorted(set_df['set_number'].unique()):
        # .get() sur une Series indexée peut retourner une Series si l'index n'est pas unique
        # On force la conversion en scalaire Python pour éviter l'ambiguïté booléenne
        nb_h_raw = hooks.get(sn, 0)
        raw_val  = nb_h_raw.iloc[0] if hasattr(nb_h_raw, 'iloc') else nb_h_raw
        # NaN si la cellule total_hooks_count est vide ou 0 — CPUE laissée à None
        # (Phase 1 doit déjà signaler cette erreur avant qu'on arrive ici)
        nb_h = int(raw_val) if pd.notna(raw_val) and raw_val > 0 else 0
        nb_h_mil = nb_h / 1000 if nb_h > 0 else None
        row = {
            'SET':          set_id_map.get(sn, str(sn)),
            'Nb hameçons':  nb_h if nb_h > 0 else '—',
        }
        for esp in especes_ret:
            n = len(ret[(ret['set_number'] == sn) & (ret['fao_code'] == esp)])
            row[esp] = round(n / nb_h_mil, 2) if nb_h_mil else None
        n_total = len(catch_df[catch_df['set_number'] == sn])
        row['CPUE totale'] = round(n_total / nb_h_mil, 2) if nb_h_mil else None
        lignes.append(row)

    return pd.DataFrame(lignes)


# ─────────────────────────────────────────────────────────────────────────────
# 2 — TABLEAU DES OPÉRATIONS DE FILAGE
# ─────────────────────────────────────────────────────────────────────────────

def calc_tableau_filage(set_df):
    """
    Construit le tableau des opérations de filage à partir de 2_set.
    """
    if set_df.empty:
        return pd.DataFrame()

    cols_sources = {
        'set_id':                    'ID filage',
        'setting_start_timestamp':   'Début filage',
        'setting_start_latitude':    'Lat. début filage',
        'setting_start_longitude':   'Lon. début filage',
        'setting_end_timestamp':     'Fin filage',
        'setting_end_latitude':      'Lat. fin filage',
        'setting_end_longitude':     'Lon. fin filage',
        'total_hooks_count':         'Nb hameçons',
        'hauling_start_timestamp':   'Début virage',
        'hauling_end_timestamp':     'Fin virage',
        'sst':                       'Temp. surface (°C)',
    }
    # Garder seulement les colonnes qui existent dans set_df
    cols_presentes = {k: v for k, v in cols_sources.items() if k in set_df.columns}
    df = set_df[list(cols_presentes.keys())].copy()
    df = df.rename(columns=cols_presentes)
    return df.sort_values('ID filage') if 'ID filage' in df.columns else df


# ─────────────────────────────────────────────────────────────────────────────
# 2b — TABLEAUX FILAGE ET VIRAGE (PNG rapport)
# ─────────────────────────────────────────────────────────────────────────────

def _dd_vers_ddmm(val, is_lat):
    """
    Convertit des degrés décimaux en format DD°MM,DD' avec hémisphère.
    Exemples : -18.949 (lat) → 18°56,94'S   |   59.782 (lon) → 59°46,93'E
    Convention française : virgule comme séparateur décimal des minutes.
    """
    if pd.isna(val):
        return '—'
    hemi = ('N' if val >= 0 else 'S') if is_lat else ('E' if val >= 0 else 'W')
    val  = abs(float(val))
    deg  = int(val)
    minutes = (val - deg) * 60
    # Format : degrés sans zéro superflu, minutes sur 5 caractères (MM,DD)
    return f"{deg}°{minutes:05.2f}'{hemi}".replace('.', ',')


def _fmt_datetime(val):
    """Formate un timestamp en 'JJ/MM/AAAA HH:MM' (sans secondes)."""
    if pd.isna(val):
        return '—'
    try:
        return pd.Timestamp(val).strftime('%d/%m/%Y %H:%M')
    except Exception:
        return str(val)


def calc_tableaux_filage_virage(set_df):
    """
    Construit les 2 tableaux des opérations de filage et de virage.

    Colonnes filage : SET | Date/Heure début | Lat. | Lon. | Date/Heure fin | Lat. | Lon. | Vitesse (nd) | Nb hameçons
    Colonnes virage : SET | Date/Heure début | Lat. | Lon. | Date/Heure fin | Lat. | Lon. | Coupures

    SET = identifiant complet (ex: SEALORXXX-1).
    Coordonnées en DD°MM,DD'. Datetimes sans secondes.
    Retourne un tuple (df_filage, df_virage).
    """
    if set_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = set_df.sort_values('set_number').copy()

    # Mapping set_number → set_id complet (ex: 1 → 'SEALORXXX-1')
    set_id_map = df.set_index('set_number')['set_id'].to_dict() if 'set_id' in df.columns else {}

    def _set_id(row):
        sn = row.get('set_number')
        return set_id_map.get(sn, str(sn)) if sn is not None else '?'

    # ── Tableau filage ────────────────────────────────────────────────────────
    lignes_filage = []
    for _, row in df.iterrows():
        lignes_filage.append({
            'SET':                 _set_id(row),
            'Date/Heure début':    _fmt_datetime(row.get('setting_start_timestamp')),
            'Lat. début':          _dd_vers_ddmm(row.get('setting_start_latitude'),  is_lat=True),
            'Lon. début':          _dd_vers_ddmm(row.get('setting_start_longitude'), is_lat=False),
            'Date/Heure fin':      _fmt_datetime(row.get('setting_end_timestamp')),
            'Lat. fin':            _dd_vers_ddmm(row.get('setting_end_latitude'),    is_lat=True),
            'Lon. fin':            _dd_vers_ddmm(row.get('setting_end_longitude'),   is_lat=False),
            'Vitesse (nd)':        (f"{row['setting_speed']:.1f}"
                                    if 'setting_speed' in df.columns
                                    and pd.notna(row.get('setting_speed')) else '—'),
            'Nb hameçons':         (int(row['total_hooks_count'])
                                    if 'total_hooks_count' in df.columns
                                    and pd.notna(row.get('total_hooks_count')) else '—'),
        })
    df_filage = pd.DataFrame(lignes_filage)

    # ── Tableau virage ────────────────────────────────────────────────────────
    lignes_virage = []
    for _, row in df.iterrows():
        lignes_virage.append({
            'SET':                 _set_id(row),
            'Date/Heure début':    _fmt_datetime(row.get('hauling_start_timestamp')),
            'Lat. début':          _dd_vers_ddmm(row.get('hauling_start_latitude'),  is_lat=True),
            'Lon. début':          _dd_vers_ddmm(row.get('hauling_start_longitude'), is_lat=False),
            'Date/Heure fin':      _fmt_datetime(row.get('hauling_end_timestamp')),
            'Lat. fin':            _dd_vers_ddmm(row.get('hauling_end_latitude'),    is_lat=True),
            'Lon. fin':            _dd_vers_ddmm(row.get('hauling_end_longitude'),   is_lat=False),
            'Coupures':            (int(row['hauling_breaks'])
                                    if 'hauling_breaks' in df.columns
                                    and pd.notna(row.get('hauling_breaks')) else '—'),
        })
    df_virage = pd.DataFrame(lignes_virage)

    return df_filage, df_virage


# ─────────────────────────────────────────────────────────────────────────────
# 3 — DONNÉES POUR LES CAMEMBERTS CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

def calc_captures_par_groupe(catch_df):
    """
    Camembert Figure 1 : répartition des captures par groupe d'espèces.
    Retourne une Series {groupe_fr: nb_captures}, triée décroissant.

    Inclut tous les groupes présents dans les données, même les groupes rares
    (Oiseaux marins, Requin-baleine, Céphalopodes, Inconnu...).
    Le groupby final fusionne les éventuels doublons de noms français.
    """
    counts = catch_df.groupby('species_group').size().rename(index=GROUPES_FR)
    # Fusionner si deux codes anglais ont le même nom français (précaution)
    counts = counts.groupby(level=0).sum()
    return counts.sort_values(ascending=False)


def calc_captures_par_espece_groupe(catch_df, groupe_en):
    """
    Camemberts Figures 2-5 : captures par espèce pour un groupe donné.
    groupe_en : nom en anglais ObServe ('Billfishes', 'Sharks', etc.)
    Retourne une Series {nom_commun_ou_code: nb_captures}.
    """
    subset = catch_df[catch_df['species_group'] == groupe_en].copy()
    if subset.empty:
        return None

    # Utiliser nom_commun si disponible, sinon fao_code
    subset['label'] = subset.apply(_label_espece, axis=1)
    result = subset.groupby('label').size().sort_values(ascending=False)
    return result if len(result) > 0 else None


# ─────────────────────────────────────────────────────────────────────────────
# 4 — HISTOGRAMMES DE TAILLE
# ─────────────────────────────────────────────────────────────────────────────

def calc_classes_taille(sample_df, fao_code):
    """
    Histogramme Figure 6-9 : classes de taille (10 cm) pour une espèce.
    Retourne une Series {classe: nb_individus} ou None si pas assez de données.
    """
    subset = sample_df[sample_df['fao_code'] == fao_code]['size'].dropna()
    if len(subset) < NB_MESURES_MIN:
        return None

    # Bornes des classes : multiples de TAILLE_STEP
    t_min = int(subset.min() // TAILLE_STEP) * TAILLE_STEP
    t_max = int(subset.max() // TAILLE_STEP + 1) * TAILLE_STEP
    bins  = list(range(t_min, t_max + TAILLE_STEP, TAILLE_STEP))

    counts, edges = np.histogram(subset, bins=bins)
    labels = [f'{int(edges[i])}-{int(edges[i+1])}' for i in range(len(edges)-1)]
    return pd.Series(counts, index=labels, name=fao_code)


# ─────────────────────────────────────────────────────────────────────────────
# 5 — ÉTAT À LA CAPTURE DES ESPÈCES COMMERCIALES
# ─────────────────────────────────────────────────────────────────────────────

def calc_etat_capture(catch_df):
    """
    Figure 11 : nb vivants / nb morts à la capture, pour les espèces commerciales
    CONSERVÉES EN CALE uniquement (catch_fate == 'RET').

    Correspond au titre du rapport : "État à la capture des espèces conservées en cale".
    Retourne un DataFrame colonnes ['Vivant', 'Mort'] indexé par nom espèce.
    """
    # On ne regarde que les captures conservées (RET) des espèces commerciales
    subset = catch_df[
        (catch_df['fao_code'].isin(ESPECES_COMMERCIALES)) &
        (catch_df['catch_fate'] == 'RET')
    ].copy()
    if subset.empty:
        return None

    subset['label']  = subset.apply(_label_espece, axis=1)
    subset['vivant'] = subset['status_at_capture'].apply(_est_vivant_capture)
    subset['mort']   = subset['status_at_capture'] == 'D'

    result = subset.groupby('label').agg(
        Vivant=('vivant', 'sum'),
        Mort  =('mort',   'sum')
    )
    result = result[result.sum(axis=1) > 0]
    # Trier par total (Vivant + Mort) décroissant
    result = result.loc[result.sum(axis=1).sort_values(ascending=False).index]
    return result if not result.empty else None


# ─────────────────────────────────────────────────────────────────────────────
# 6 — SURVIE DES REQUINS
# ─────────────────────────────────────────────────────────────────────────────

def calc_survie_requins(catch_df):
    """
    Figure 12 : vivant / mort au relâcher pour les requins relâchés (DIS ou ESC)
    ayant un statut au relâcher renseigné.

    La référence IRD inclut DIS et ESC — un requin échappé peut aussi avoir
    un statut au relâcher noté par l'observateur.
    Retourne une Series {statut: nb} ou None si pas de données.
    """
    requins_relaches = catch_df[
        (catch_df['species_group'] == 'Sharks') &
        (catch_df['catch_fate'].isin(['DIS', 'ESC']))
    ].copy()

    if len(requins_relaches) == 0:
        return None

    def _categorie(row):
        s = row['status_at_release']
        if pd.isna(s):
            # Un requin échappé (ESC) sans statut noté était vivant par définition
            # (il a activement échappé) → Vivant
            return 'Vivant'
        if s == 'D':
            return 'Mort'
        # A = vivant, AW = blessé vivant, AUSC = vivant sous stress (compté Vivant IRD)
        if s in ('A', 'AW', 'AUSC'):
            return 'Vivant'
        return 'Incertain'

    return requins_relaches.apply(_categorie, axis=1).value_counts()


# ─────────────────────────────────────────────────────────────────────────────
# 7 — DEVENIR DES CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

def calc_devenir_global(catch_df):
    """
    Figure 13 : camembert global conservés / rejetés / échappés.
    """
    mapping = {'RET': 'Conservés', 'DIS': 'Rejetés', 'ESC': 'Échappés'}
    return catch_df['catch_fate'].map(mapping).value_counts()


def calc_devenir_par_espece(catch_df):
    """
    Figure 14 : nb conservés, rejetés, échappés par espèce.
    Retourne un DataFrame colonnes ['Conservés','Rejetés','Échappés'].
    """
    mapping = {'RET': 'Conservés', 'DIS': 'Rejetés', 'ESC': 'Échappés'}
    df = catch_df.copy()
    df['devenir'] = df['catch_fate'].map(mapping)
    df['label']   = df.apply(_label_espece, axis=1)

    pivot = df.groupby(['label', 'devenir']).size().unstack(fill_value=0)
    for col in ['Conservés', 'Rejetés', 'Échappés']:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['_total'] = pivot.sum(axis=1)
    pivot = pivot.sort_values('_total', ascending=False).drop(columns='_total')
    return pivot[['Conservés', 'Rejetés', 'Échappés']]


# ─────────────────────────────────────────────────────────────────────────────
# 8 — DÉPRÉDATION
# ─────────────────────────────────────────────────────────────────────────────

def calc_depredation_totale(catch_df):
    """
    Figure 15 : déprédation totale par type de déprédateur.

    Les codes composites (ex: 'ISB;SQU') sont décomposés : une capture attaquée
    par deux déprédateurs compte 1 pour chacun. Chaque valeur de {nom_depredateur}
    représente le nombre de CAPTURES où ce déprédateur est présent.

    Note : data.sum() peut dépasser le nombre de captures déprédatées si des
    composites existent. Ne pas utiliser data.sum() comme dénominateur de %.
    Voir construire_depredation_par1 pour la gestion correcte.

    Retourne une Series {nom_depredateur: nb_captures} ou None si aucune déprédation.
    """
    subset = catch_df[catch_df['depredating_species'].notna()]
    if subset.empty:
        return None

    compteur = {}
    for val in subset['depredating_species']:
        # Un animal peut être attaqué par plusieurs déprédateurs simultanément.
        # Ex : 'ISB;SQU' = attaqué par squalelet féroce ET calmar → compte 1 pour chacun.
        for code in str(val).split(';'):
            code = code.strip()
            if code:
                nom = DEPREDATEURS_FR.get(code, code)
                compteur[nom] = compteur.get(nom, 0) + 1

    return pd.Series(compteur).sort_values(ascending=False) if compteur else None


def calc_depredation_par_espece(catch_df):
    """
    Figure 16 : déprédation par espèce commercialisable et type de déprédateur.
    Retourne un DataFrame ou None si aucune déprédation sur espèces commerciales.
    """
    subset = catch_df[
        catch_df['depredating_species'].notna() &
        catch_df['fao_code'].isin(ESPECES_COMMERCIALES)
    ]
    if subset.empty:
        return None

    lignes = []
    for _, row in subset.iterrows():
        # Même logique que calc_depredation_totale : on décompose les codes composites.
        # Un animal attaqué par ISB et SQU compte une fois pour chaque déprédateur.
        for code in str(row['depredating_species']).split(';'):
            code = code.strip()
            if code:
                lignes.append({
                    'espece':     _label_espece(row),
                    'depredateur': DEPREDATEURS_FR.get(code, code),
                })

    if not lignes:
        return None

    df_dep = pd.DataFrame(lignes)
    pivot  = df_dep.groupby(['espece', 'depredateur']).size().unstack(fill_value=0)
    pivot['_total'] = pivot.sum(axis=1)
    pivot = pivot[pivot['_total'] > 0].sort_values('_total', ascending=False).drop(columns='_total')
    return pivot if not pivot.empty else None


# ─────────────────────────────────────────────────────────────────────────────
# 9 — CPUE (captures par 1000 hameçons)
# ─────────────────────────────────────────────────────────────────────────────

def calc_cpue(catch_df, set_df):
    """
    Figure 17 : CPUE = nb captures / (nb hameçons / 1000) par calée.

    Colonnes : une par espèce cible (ESPECES_CPUE présentes dans les données)
               + une colonne 'CPUE totale' (toutes espèces confondues).
    Retourne un DataFrame : index = numéro de calée.
    """
    if set_df.empty:
        return None

    especes_presentes = [sp for sp in ESPECES_CPUE
                         if (catch_df['fao_code'] == sp).any()]

    # Nb hameçons par calée — agréger si set_number dupliqué (données avec erreurs de saisie)
    hooks = set_df.groupby('set_number')['total_hooks_count'].first()

    # CPUE totale : toutes captures confondues
    total_par_set = catch_df.groupby('set_number').size()
    cpue_totale   = total_par_set.div(hooks / 1000).reindex(hooks.index, fill_value=0)

    if not especes_presentes:
        # Pas d'espèces cibles : retourner uniquement la CPUE totale
        return pd.DataFrame({'CPUE totale': cpue_totale}).sort_index()

    # Nb captures par (set_number, fao_code) pour les espèces cibles
    counts = catch_df.groupby(['set_number', 'fao_code']).size().unstack(fill_value=0)
    counts = counts.reindex(columns=especes_presentes, fill_value=0)

    cpue = counts.div(hooks / 1000, axis=0)

    # Renommer les colonnes avec le code FAO
    noms = {}
    for sp in especes_presentes:
        row = catch_df[catch_df['fao_code'] == sp].iloc[0]
        noms[sp] = _label_espece(row)
    cpue = cpue.rename(columns=noms)

    # Ajouter la CPUE totale en dernière colonne
    cpue['CPUE totale'] = cpue_totale.reindex(cpue.index, fill_value=0)

    return cpue.sort_index()


# ─────────────────────────────────────────────────────────────────────────────
# 10 — PROFIL DE PROFONDEUR TDR
# ─────────────────────────────────────────────────────────────────────────────

def calc_profil_tdr(tdr_df):
    """
    Figure 18 : comportement médian de la ligne en profondeur.

    Pour chaque SET, collecte les profondeurs médianes de pêche de tous les TDR
    déployés lors de ce filage. Chaque boîte représente la variation de profondeur
    sur la ligne pour un SET donné.
    Retourne un dict {set_number: array_profondeurs} ou None si données insuffisantes.
    """
    if tdr_df.empty or 'median_fishing_depth' not in tdr_df.columns:
        return None

    col_set = 'set_number' if 'set_number' in tdr_df.columns else None
    if col_set is None:
        return None

    set_nums = sorted(tdr_df[col_set].dropna().unique())
    donnees = {}
    for sn in set_nums:
        profondeurs = tdr_df[tdr_df[col_set] == sn]['median_fishing_depth'].dropna()
        if len(profondeurs) > 0:
            donnees[int(sn)] = profondeurs.values

    return donnees if donnees else None


# ─────────────────────────────────────────────────────────────────────────────
# 11 — CARTE DES POSITIONS DE FILAGE ET VIRAGE
# ─────────────────────────────────────────────────────────────────────────────

# Coordonnées (lat, lon) des ports principaux de l'Océan Indien occidental.
# Correspondance établie avec le référentiel ObServe (table Harbour) par nom.
PORTS_COORDS = {
    'POINTE DES GALETS':           (-20.927,  55.292),
    'REUNION':                     (-21.000,  55.500),
    'PORT VICTORIA':               ( -4.619,  55.452),
    'PORT LOUIS':                  (-20.161,  57.499),
    'PORT MATHURIN':               (-19.670,  63.416),
    'ANTSIRANANA (DIEGO SUAREZ)':  (-12.350,  49.296),
    'DIEGO SUAREZ':                (-12.350,  49.296),
    'ANTSIRANANA':                 (-12.350,  49.296),
    'LONGONI (MAYOTTE)':           (-12.715,  45.203),
    'LONGONI':                     (-12.715,  45.203),
    'MAMOUDZOU (MAYOTTE)':         (-12.782,  45.228),
    'DZAOUDZI (MAYOTTE)':          (-12.788,  45.226),
    'MAJUNGA (MAHAJANGA)':         (-15.716,  46.316),
    'MAHAJANGA':                   (-15.716,  46.316),
    'MAJUNGA':                     (-15.716,  46.316),
    'NOSY BE':                     (-13.328,  48.273),
    'TAMATAVE':                    (-18.148,  49.396),
    'TOAMASINA':                   (-18.148,  49.396),
    'FORT DAUPHIN':                (-25.033,  47.000),
    'TOLAGNARO':                   (-25.033,  47.000),
    'TULEAR':                      (-23.352,  43.669),
    'TOLIARA':                     (-23.352,  43.669),
    'DURBAN':                      (-29.858,  31.029),
    'CAP TOWN':                    (-33.928,  18.423),
    'DAR ES SALAAM':               ( -6.813,  39.289),
    'MOMBASA':                     ( -4.053,  39.659),
}


def calc_carte_filages(set_df, trip_df):
    """
    Extrait les données de position pour la carte des filages et virages.

    Retourne un dict avec :
    - 'sets'         : liste des filages (positions de filage et virage, date, numéro)
    - 'port_depart'  : dict {nom, lat, lon} ou None
    - 'port_arrivee' : dict {nom, lat, lon} ou None
    """
    if set_df.empty:
        return None

    cols_req = [
        'setting_start_latitude', 'setting_start_longitude',
        'setting_end_latitude',   'setting_end_longitude',
        'hauling_start_latitude', 'hauling_start_longitude',
        'hauling_end_latitude',   'hauling_end_longitude',
    ]
    if not all(c in set_df.columns for c in cols_req):
        return None

    sets = []
    for _, row in set_df.sort_values('set_number').iterrows():
        # Date du début de filage (format JJ/MM)
        date_str = ''
        if pd.notna(row.get('setting_start_timestamp')):
            try:
                date_str = pd.Timestamp(row['setting_start_timestamp']).strftime('%d/%m')
            except Exception:
                pass

        # Ignorer les lignes sans coordonnées valides
        coords = [row['setting_start_latitude'], row['setting_start_longitude'],
                  row['setting_end_latitude'],   row['setting_end_longitude'],
                  row['hauling_start_latitude'], row['hauling_start_longitude'],
                  row['hauling_end_latitude'],   row['hauling_end_longitude']]
        if any(pd.isna(c) for c in coords):
            continue

        sets.append({
            'num':           int(row.get('set_number', 0)),
            'date':          date_str,
            'setting_start': (float(row['setting_start_latitude']),
                              float(row['setting_start_longitude'])),
            'setting_end':   (float(row['setting_end_latitude']),
                              float(row['setting_end_longitude'])),
            'hauling_start': (float(row['hauling_start_latitude']),
                              float(row['hauling_start_longitude'])),
            'hauling_end':   (float(row['hauling_end_latitude']),
                              float(row['hauling_end_longitude'])),
        })

    if not sets:
        return None

    def _trouver_port(nom):
        """Cherche les coordonnées d'un port par son nom (correspondance partielle)."""
        if not nom or pd.isna(nom):
            return None
        nom_up = str(nom).strip().upper()
        if nom_up in PORTS_COORDS:
            lat, lon = PORTS_COORDS[nom_up]
            return {'nom': str(nom).strip().title(), 'lat': lat, 'lon': lon}
        for cle, (lat, lon) in PORTS_COORDS.items():
            if cle in nom_up or nom_up in cle:
                return {'nom': str(nom).strip().title(), 'lat': lat, 'lon': lon}
        return None

    port_dep = port_arr = None
    if not trip_df.empty:
        t = trip_df.iloc[0]
        port_dep = _trouver_port(t.get('departure_port'))
        port_arr = _trouver_port(t.get('landing_port'))

    return {'sets': sets, 'port_depart': port_dep, 'port_arrivee': port_arr}


# ─────────────────────────────────────────────────────────────────────────────
# DÉTECTION DES ZEE DE PÊCHE PAR FILAGE
# ─────────────────────────────────────────────────────────────────────────────

# Correspondance nom Marine Regions → nom court en français
_NOM_ZEE_FR = {
    'Madagascan Exclusive Economic Zone':                          'Madagascar',
    'Mauritian Exclusive Economic Zone':                           'Maurice',
    'Mauritian Exclusive Economic Zone (Chagos Archipelago)':     'Maurice (Chagos)',
    'Seychellois Exclusive Economic Zone':                         'Seychelles',
    'Mozambican Exclusive Economic Zone':                          'Mozambique',
    'Tanzanian Exclusive Economic Zone':                           'Tanzanie',
    'Kenyan Exclusive Economic Zone':                              'Kenya',
    'Somali Exclusive Economic Zone':                              'Somalie',
    'South African Exclusive Economic Zone':                       'Afrique du Sud',
    'South African Exclusive Economic Zone (Prince Edward Islands)': 'Afrique du Sud (Pr. Edward)',
    'Namibian Exclusive Economic Zone':                            'Namibie',
    'Angolan Exclusive Economic Zone':                             'Angola',
    'Comoran Exclusive Economic Zone':                             'Comores',
    'Indian Exclusive Economic Zone':                              'Inde',
    'Sri Lankan Exclusive Economic Zone':                          'Sri Lanka',
    'Maldivian Exclusive Economic Zone':                           'Maldives',
    'French Exclusive Economic Zone (Réunion)':                   'France (Réunion)',
    'French Exclusive Economic Zone (R\xe9union)':                 'France (Réunion)',
    'French Exclusive Economic Zone (Europa Island)':              'France (Europa)',
    'French Exclusive Economic Zone (Juan de Nova Island)':        'France (Juan de Nova)',
    'French Exclusive Economic Zone (Bassas da India)':            'France (Bassas da India)',
    'French Exclusive Economic Zone (Amsterdam and Saint Paul Islands)': 'France (Amsterdam/St-Paul)',
    'French Exclusive Economic Zone (Crozet Islands)':             'France (Crozet)',
    'French Exclusive Economic Zone (Kerguel\xe9n)':               'France (Kerguélen)',
    'Gabonese Exclusive Economic Zone':                            'Gabon',
    'Congolese (Republic of) Exclusive Economic Zone':             'Congo',
    'Congolese (Democratic Republic of) Exclusive Economic Zone':  'RDC',
    'Cameroonian Exclusive Economic Zone':                         'Cameroun',
    'Nigerian Exclusive Economic Zone':                            'Nigeria',
    'Equatorial Guinean Exclusive Economic Zone':                  'Guinée Équatoriale',
    'S\xe3o Tom\xe9an Exclusive Economic Zone':                    'São Tomé',
}


def _nom_zee_court(nom_complet):
    """Retourne le nom court français d'une ZEE à partir de son nom Marine Regions."""
    if nom_complet in _NOM_ZEE_FR:
        return _NOM_ZEE_FR[nom_complet]
    # Fallback : retirer " Exclusive Economic Zone" et garder le reste
    court = nom_complet.replace(' Exclusive Economic Zone', '').strip()
    return court if court else nom_complet


def calc_zones_peche(set_df):
    """
    Détermine les noms des ZEE dans lesquelles ont eu lieu les filages.

    Utilise un test point-dans-polygone (ray casting) sur les positions GPS
    de début de filage, à partir du fichier ocean_indien_zee_polygones.geojson.

    Retourne une liste triée de noms de ZEE (chaînes de caractères).
    Retourne une liste vide si le fichier n'existe pas ou si les données GPS manquent.
    """
    import os, json

    if set_df.empty:
        return []

    geojson_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'data', 'ocean_indien_zee_polygones.geojson'
    )
    if not os.path.exists(geojson_path):
        return []

    with open(geojson_path, 'r', encoding='utf-8') as f:
        gj = json.load(f)

    def _ray_cast(lon, lat, ring):
        """Test point-dans-polygone par lancer de rayon."""
        n, inside, j = len(ring), False, len(ring) - 1
        for i in range(n):
            xi, yi = ring[i][0], ring[i][1]
            xj, yj = ring[j][0], ring[j][1]
            if ((yi > lat) != (yj > lat)) and \
               (lon < (xj - xi) * (lat - yi) / (yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    # Positions GPS de début de filage
    positions = []
    for _, row in set_df.iterrows():
        try:
            lat = float(row['setting_start_latitude'])
            lon = float(row['setting_start_longitude'])
            if pd.notna(lat) and pd.notna(lon):
                positions.append((lon, lat))
        except Exception:
            pass

    zones = set()
    for feature in gj['features']:
        nom  = feature['properties']['nom']
        # Ignorer les zones de chevauchement
        if 'overlap' in nom.lower() or 'claim' in nom.lower():
            continue
        ring = feature['geometry']['coordinates'][0]
        for lon, lat in positions:
            if _ray_cast(lon, lat, ring):
                zones.add(nom)
                break  # au moins 1 filage dans cette ZEE suffit

    return sorted(f"ZEE {_nom_zee_court(z)}" for z in zones)


# ─────────────────────────────────────────────────────────────────────────────
# 12 — DESCRIPTION DE L'ENGIN DE PÊCHE (tables 2 à 6)
# ─────────────────────────────────────────────────────────────────────────────

def calc_tableaux_engin(tables):
    """
    Construit les 4 tableaux de la section "Description de l'engin de pêche"
    à partir des tables ObServe 2_set, 3_floatline, 4_branchline, 5_hook, 6_bait.

    Toutes les traductions viennent du référentiel ObServe local (referentiel_observe.json).
    Si le référentiel est absent, les codes bruts sont affichés.

    Retourne un dict {clé: DataFrame} :
      'engin'    : Tableau 1 — Caractéristiques de l'engin de pêche (caract. + compo.)
      'orins'    : Tableau 2 — Schéma de palangre (orins)
      'hamecons' : Tableau 3 — Hameçons
      'appats'   : Tableau 4 — Appâts et attractants
    """
    set_df        = tables.get('2_set',        pd.DataFrame())
    floatline_df  = tables.get('3_floatline',  pd.DataFrame())
    branchline_df = tables.get('4_branchline', pd.DataFrame())
    hook_df       = tables.get('5_hook',       pd.DataFrame())
    bait_df       = tables.get('6_bait',       pd.DataFrame())

    n_sets = max(len(set_df), 1)  # dénominateur pour les moyennes sur tous les filages

    # ── Chargement du référentiel ─────────────────────────────────────────────

    ref_brut = {}
    try:
        racine     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        chemin_ref = os.path.join(racine, 'phase1_detection', 'data', 'referentiel_observe.json')
        with open(chemin_ref, encoding='utf-8') as f:
            ref_brut = json.load(f).get('referentiels', {})
    except Exception:
        pass  # référentiel absent → codes bruts affichés

    def _entrees_ref(nom_court):
        """Retourne la liste d'entrées d'une table du référentiel (recherche par nom court)."""
        for cle in ref_brut:
            if cle.endswith('.' + nom_court):
                return ref_brut[cle]
        return []

    def _dict_code_fr(nom_court):
        """Construit {code: label_fr} depuis une table du référentiel."""
        return {
            str(e.get('code', '')): (e.get('label2') or e.get('label1') or '')
            for e in _entrees_ref(nom_court) if e.get('code')
        }

    def _dict_en_fr(nom_court):
        """Construit {label_en: label_fr} pour les colonnes qui stockent le label anglais."""
        return {
            str(e.get('label1', '')): (e.get('label2') or e.get('label1') or '')
            for e in _entrees_ref(nom_court) if e.get('label1')
        }

    # Dictionnaires de traduction depuis le référentiel
    ref_ligne          = _dict_code_fr('LineType')          # MON → Nylon monofilament
    ref_forme          = _dict_code_fr('SettingShape')      # D → Rectiligne, U → En U
    ref_etat_appat     = _dict_code_fr('BaitSettingStatus') # DEF → Décongelé
    ref_type_cyal_code = _dict_code_fr('LightsticksType')  # CH → Chimique
    ref_type_cyal_en   = _dict_en_fr('LightsticksType')    # Chemical → Chimique
    ref_coul_cyal_code = _dict_code_fr('LightsticksColor') # Y → Jaune
    ref_coul_cyal_en   = _dict_en_fr('LightsticksColor')   # Yellow → Jaune

    # Fonctions de traduction : essaient le label EN d'abord, puis le code, sinon brut
    def _trad_ligne(code):
        return ref_ligne.get(str(code).strip().upper(), str(code))

    def _trad_forme(code):
        return ref_forme.get(str(code).strip().upper(), str(code))

    def _trad_type_cyal(val):
        v = str(val).strip()
        return ref_type_cyal_en.get(v) or ref_type_cyal_code.get(v.upper(), v)

    def _trad_coul_cyal(val):
        v = str(val).strip()
        return ref_coul_cyal_en.get(v) or ref_coul_cyal_code.get(v.upper(), v)

    # ── Fonctions utilitaires internes ────────────────────────────────────────

    def _fmt(v):
        """Supprime les '.0' inutiles sur les nombres entiers stockés comme float."""
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)

    def _val_num(df, col):
        """
        Retourne la valeur typique d'une colonne numérique.
        Si toutes identiques → valeur seule. Si variation → 'dominant (min–max)'.
        """
        if col not in df.columns:
            return '—'
        serie = df[col].dropna()
        if serie.empty:
            return '—'
        dominant = serie.mode().iloc[0]
        if serie.nunique() == 1:
            return _fmt(dominant)
        return f'{_fmt(dominant)} ({_fmt(serie.min())}–{_fmt(serie.max())})'

    def _val_code(df, col, traduction):
        """Retourne le label traduit de la valeur dominante d'une colonne code."""
        if col not in df.columns:
            return '—'
        serie = df[col].dropna()
        if serie.empty:
            return '—'
        return traduction(serie.mode().iloc[0])

    def _bool_val(df, col):
        """Traduit une colonne booléenne/0-1 en 'Oui' / 'Non'."""
        if col not in df.columns:
            return '—'
        serie = df[col].dropna()
        if serie.empty:
            return '—'
        dominant = serie.mode().iloc[0]
        if isinstance(dominant, bool):
            return 'Oui' if dominant else 'Non'
        if isinstance(dominant, (int, float)):
            return 'Oui' if dominant == 1 else 'Non'
        return 'Oui' if str(dominant).strip().upper() in {'TRUE', 'YES', '1', 'VRAI', 'OUI'} else 'Non'

    def _formes_str(df):
        """Retourne 'Forme1 : N filages / Forme2 : M filages' depuis setting_shape."""
        if 'setting_shape' not in df.columns:
            return '—'
        vc = df['setting_shape'].value_counts()
        return ' / '.join(
            f"{_trad_forme(code)} : {n} filage{'s' if n > 1 else ''}"
            for code, n in vc.items()
        )

    def _hooks_str(df):
        """
        Nombre moyen d'hameçons par filage, en excluant les filages partiels
        (< 80 % du maximum). Les sets exclus sont indiqués.
        """
        if 'total_hooks_count' not in df.columns:
            return '—'
        hooks    = df['total_hooks_count']
        seuil    = hooks.max() * 0.80
        normaux  = df[hooks >= seuil]
        partiels = df[hooks <  seuil]
        moy = int(normaux['total_hooks_count'].mean())
        if len(partiels) > 0:
            sets_p = (partiels['set_number'].tolist()
                      if 'set_number' in df.columns else ['?'])
            nb_p   = int(partiels['total_hooks_count'].iloc[0])
            return f'{moy} (set(s) {sets_p} exclus : {nb_p} hameçons)'
        return str(moy)

    def _proportion_str(vc, n_total, traduction=None):
        """
        Depuis une value_counts(), calcule les proportions sur n_total filages
        et retourne 'Label1 : X % / Label2 : Y %'.
        traduction : fonction optionnelle pour traduire les valeurs.
        """
        parties = []
        for val, n in vc.items():
            label = traduction(val) if traduction else str(val)
            pct   = round(n / n_total * 100)
            parties.append(f'{label} : {pct} %')
        return ' / '.join(parties)

    # ── TABLEAU 1 — Caractéristiques de l'engin de pêche ─────────────────────

    lignes = []
    if not set_df.empty:
        lignes.append(('Type de ligne principale',
            _val_code(set_df, 'mainline_type', _trad_ligne)))
        lignes.append(('Forme du filage',
            _formes_str(set_df)))
        lignes.append(('Vitesse de filage (nœuds)',
            _val_num(set_df, 'setting_speed')))
        lignes.append(('Shooter utilisé',
            _bool_val(set_df, 'shooter_used')))
        lignes.append(('Profondeur ciblée (m)',
            _val_num(set_df, 'max_depth_targeted')))
        lignes.append(('Nombre de sections par filage',
            _val_num(set_df, 'total_sections_count')))
        lignes.append(('Nombre de baskets par section',
            _val_num(set_df, 'baskets_per_section_count')))
        lignes.append(("Nombre d'avançons par basket",
            _val_num(set_df, 'branchlines_per_basket_count')))
        # Longueur du basket = nb avançons × 50 m (espacement standard entre avançons)
        nb_avancons = _val_num(set_df, 'branchlines_per_basket_count')
        try:
            longueur_basket = f'{int(float(nb_avancons)) * 50} m'
        except (ValueError, TypeError):
            longueur_basket = '—'
        lignes.append(('Longueur du basket (m)', longueur_basket))
        lignes.append(("Nb moyen d'hameçons par filage",
            _hooks_str(set_df)))

    if not branchline_df.empty:
        if 'branchline_length_m' in branchline_df.columns:
            # Stocké en mètres avec 4 décimales (multiples de brasse) → arrondi à 1
            lignes.append(('Longueur des avançons (m)',
                f"{branchline_df['branchline_length_m'].mode()[0]:.1f}"))
        lignes.append(('Type haut de ligne avançon',
            _val_code(branchline_df, 'branchline_topline_type', _trad_ligne)))
        lignes.append(('Type bas de ligne avançon',
            _val_code(branchline_df, 'branchline_traceline_type', _trad_ligne)))

    if not set_df.empty:
        lignes.append(('Snap lestés',
            _bool_val(set_df, 'weighted_snap')))
        lignes.append(('Émerillons lestés',
            _bool_val(set_df, 'weighted_swivel')))
        lignes.append(('Poids des lests (kg)',
            _val_num(set_df, 'snap_weight')))

    df_engin = pd.DataFrame(lignes, columns=['Paramètre', 'Valeur'])

    # ── TABLEAU 2 — Schéma de palangre (orins) ────────────────────────────────
    # Proportion = somme sur les filages présents / n_sets total (0 si filage absent)

    df_orins = pd.DataFrame()
    if not floatline_df.empty:
        cols_req = {'floatline_line_type', 'floatline_length_m', 'floatline_proportion'}
        if cols_req.issubset(floatline_df.columns):
            result = []
            for (lt, lm), g in floatline_df.groupby(
                    ['floatline_line_type', 'floatline_length_m'], sort=False):
                prop = g['floatline_proportion'].sum() / n_sets
                result.append((_trad_ligne(lt), lm, prop))
            result.sort(key=lambda x: (x[0], float(x[1])))  # tri par type puis longueur
            df_orins = pd.DataFrame(result, columns=[
                'Type de ligne', "Longueur de l'orin (m)", 'Proportion (%)',
            ])
            df_orins["Longueur de l'orin (m)"] = df_orins[
                "Longueur de l'orin (m)"].apply(lambda v: f'{v:.1f}')
            df_orins['Proportion (%)'] = df_orins['Proportion (%)'].apply(
                lambda v: f'{v:.1f} %')

    # ── TABLEAU 3 — Hameçons ──────────────────────────────────────────────────
    # Proportion = somme / n_sets (0 si filage absent)

    df_hamecons = pd.DataFrame()
    if not hook_df.empty and {'hook_type', 'hook_proportion'}.issubset(hook_df.columns):
        result = []
        for ht, g in hook_df.groupby('hook_type'):
            prop = g['hook_proportion'].sum() / n_sets
            result.append((ht, prop))
        result.sort(key=lambda x: -x[1])
        df_hamecons = pd.DataFrame(result, columns=["Type d'hameçon", 'Proportion (%)'])
        df_hamecons['Proportion (%)'] = df_hamecons['Proportion (%)'].apply(
            lambda v: f'{v:.1f} %')

    # ── TABLEAU 4 — Appâts et attractants ─────────────────────────────────────

    lignes_appat = []
    if not bait_df.empty and 'bait_label' in bait_df.columns:
        # Appâts : proportion = somme / n_sets, tailles non moyennées
        appat_prop    = {}
        appat_tailles = {}
        for label, g in bait_df.groupby('bait_label'):
            if 'bait_proportion' in g.columns:
                appat_prop[label] = g['bait_proportion'].sum() / n_sets
            if 'bait_size' in g.columns:
                tailles = sorted(g['bait_size'].dropna().unique().tolist())
                appat_tailles[label] = ', '.join(str(int(t)) for t in tailles)

        tri = sorted(appat_prop.items(), key=lambda x: -x[1])
        lignes_appat.append(('Appâts (proportion)',
            ' / '.join(f'{b} : {p:.0f} %' for b, p in tri)))
        lignes_appat.append(('Taille des appâts (cm)',
            ' / '.join(f'{b} : {appat_tailles.get(b, "?")} cm' for b, _ in tri)))

        if 'bait_setting_status' in bait_df.columns:
            code_etat = str(bait_df['bait_setting_status'].mode()[0]).strip()
            lignes_appat.append(("État de l'appât",
                ref_etat_appat.get(code_etat, code_etat)))

    if not set_df.empty:
        lignes_appat.append(('Cyalumes par basket',
            _val_num(set_df, 'lightsticks_per_basket_count')))
        if 'lightstick_type' in set_df.columns:
            lignes_appat.append(('Type de cyalumes',
                _proportion_str(set_df['lightstick_type'].value_counts(),
                                n_sets, _trad_type_cyal)))
        if 'lightstick_color' in set_df.columns:
            lignes_appat.append(('Couleur des cyalumes',
                _proportion_str(set_df['lightstick_color'].value_counts(),
                                n_sets, _trad_coul_cyal)))

    df_appat = pd.DataFrame(lignes_appat, columns=['Paramètre', 'Valeur'])

    return {
        'engin':    df_engin,
        'orins':    df_orins,
        'hamecons': df_hamecons,
        'appats':   df_appat,
    }
