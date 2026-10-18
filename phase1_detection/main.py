#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Phase 1 : Détection des erreurs dans les exports ObServe
IRD OB7 - SCISEA

Ce programme ouvre un fichier export d'ObServe (zip ou xlsx),
lit les tables de données, détecte les erreurs de saisie courantes,
les affiche dans une interface graphique, et permet d'exporter
un rapport des anomalies.

Lancer : python phase1_detection/main.py
"""

# ─────────────────────────────────────────────────────────────────────────────
# INSTALLATION AUTOMATIQUE DES LIBRAIRIES
# Ce bloc vérifie que tksheet est installé. S'il est absent, il l'installe
# automatiquement via pip (nécessite internet uniquement lors du 1er lancement).
# Ensuite l'application fonctionne entièrement hors ligne.
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
        print(f"Installation de {nom_paquet} (première utilisation)...")
        subprocess.check_call(
            [sys.executable, '-m', 'pip', 'install', nom_paquet, '--quiet'],
            stdout=subprocess.DEVNULL
        )
        print(f"{nom_paquet} installé avec succès.")

_installer_si_absent('tksheet')

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS — bibliothèques nécessaires au fonctionnement
# ─────────────────────────────────────────────────────────────────────────────

import os           # pour naviguer dans les dossiers
import json         # lecture/écriture des fichiers d'annotations
import re           # expressions régulières (sanitize du trip_id)
import hashlib      # hash MD5 pour construire une clé d'annotation unique
import zipfile      # pour ouvrir les fichiers .zip
import tempfile     # pour créer un dossier temporaire lors de l'extraction du zip

# utils.py est à la racine du projet — on s'assure qu'elle est dans sys.path
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
from utils import lire_zip, lire_fichier, VERSION, COULEURS, POLICE, appliquer_hover, ajouter_logos_bandeau, ajouter_titre_bandeau   # fonctions communes + constantes
import tkinter as tk                      # interface graphique (inclus dans Python)
from tkinter import ttk, filedialog, messagebox  # composants de l'interface
import pandas as pd  # lecture et manipulation des tableaux Excel
import tksheet       # tableau avec surlignage de cellules individuelles
import threading     # pour ne pas bloquer l'interface pendant les téléchargements
import referentiel   # gestion du référentiel ObServe local
import lang          # gestion de la langue (fr/en)
from lang import T   # fonction de traduction : T('clé') retourne le texte traduit

# ─────────────────────────────────────────────────────────────────────────────
# DONNÉES DE RÉFÉRENCE
# Intégrées directement dans le script pour fonctionner entièrement hors ligne.
# ─────────────────────────────────────────────────────────────────────────────

# Codes FAO attendus dans 8_catch (source : règles de correction SCISEA)
CODES_FAO_8CATCH = {
    'YFT', 'BET', 'ALB', 'SKJ', 'SWO', 'BUM', 'BLM', 'MLS', 'SSP', 'SFA',
    'FAL', 'OCS', 'TIG', 'SMA', 'BSH', 'BTH', 'ISB', 'WAH', 'RZV', 'MOP',
    'DOL', 'LEC', 'OIL', 'GES', 'ALX', 'GBA', 'LGH', 'PLS', 'TUG', 'LKV',
    'TTH', 'TTL', 'DKK',
}

# Espèces autorisées à être retenues à bord (catch_fate = 'RET')
ESPECES_RETENUES_OK = {
    'YFT', 'BET', 'ALB', 'SWO', 'BUM', 'BLM', 'MLS', 'SSP', 'SFA', 'WAH', 'DOL',
}

# Espèces cibles commercialisables — utilisées dans la règle "rejet suspect"
# Ce sont les mêmes que ESPECES_RETENUES_OK (poissons que l'on est autorisé à conserver).
CODES_FAO_CIBLES = ESPECES_RETENUES_OK

# Espèces "hors gabarit" : trop grandes pour être remontées à bord et mesurées.
# Inclut : requins, raies (sauf PLS qui est petite et doit être mesurée),
#          cétacés, requin baleine.
# Exclut : tortues, oiseaux marins (assez petits pour être mesurés par l'observateur).
# Source : référentiel ObServe — groupes Sharks(2), Rays(11)-PLS, Cetaceans(5), Whale shark(12).
CODES_HORS_GABARIT = {
    # Sharks (groupe 2) — ISB exclu : Isistius brasiliensis (requin cigare, ~14–56 cm, mesurable)
    'ALS', 'ALV', 'AML', 'ASK', 'BLR', 'BRO', 'BSH', 'BSK', 'BTH',
    'CCA', 'CCB', 'CCE', 'CCF', 'CCL', 'CCP', 'CVX', 'CWZ', 'DUS',
    'FAL', 'FOD', 'FRH', 'HDQ', 'HXW', 'LMA', 'LMP', 'LMZ',
    'MAK', 'MSK', 'NGA', 'OCS', 'OCX', 'POR', 'PSK', 'PTH', 'PWS',
    'REX', 'RHA', 'RSK', 'SHL', 'SHX', 'SKH', 'SKX', 'SMA',
    'SPK', 'SPL', 'SPN', 'SPY', 'SPZ', 'SSQ', 'THR', 'TIG', 'TRK', 'WSH',
    # Rays (groupe 11) — sauf PLS (Pelagic stingray, 12–96 cm, mesurable)
    'DHF', 'EAG', 'JDP', 'MAE', 'MAN', 'MNT', 'MRM', 'MYL', 'NZX',
    'RAJ', 'RDB', 'RDC', 'RDX', 'RDZ', 'RHI', 'RMA', 'RMB', 'RMC',
    'RMH', 'RMJ', 'RMM', 'RMO', 'RMT', 'RMV', 'RRY', 'SRX', 'STT', 'TOD', 'TTO',
    # Cetaceans (groupe 5)
    'BBW', 'BCW', 'BGW', 'BLW', 'BRW', 'DBO', 'DCL', 'DCO', 'DLP',
    'DPN', 'DRR', 'DSA', 'DSI', 'DST', 'DWW', 'FAW', 'FIW', 'FRD',
    'GLO', 'HUW', 'KIW', 'KPW', 'MAM', 'MEW', 'MIW', 'MYS', 'ODN',
    'PIW', 'PYW', 'RTD', 'SHW', 'SIW', 'SPW',
    # Whale shark (groupe 12)
    'RHN',
}

# Référence tailles et codes de mesure par espèce (source : Capture measurement.xlsx)
# size_min / size_max : limites en centimètres
# size_codes : ensemble des codes de mesure valides pour cette espèce
CAPTURE_MEASUREMENT = {
    'YFT': {'size_min': 40,  'size_max': 180, 'size_codes': {'FL'}},
    'BET': {'size_min': 40,  'size_max': 180, 'size_codes': {'FL'}},
    'ALB': {'size_min': 80,  'size_max': 120, 'size_codes': {'FL'}},
    'SKJ': {'size_min': 40,  'size_max': 100, 'size_codes': {'FL'}},
    'SWO': {'size_min': 40,  'size_max': 250, 'size_codes': {'LJFL'}},
    'BUM': {'size_min': 80,  'size_max': 300, 'size_codes': {'LJFL'}},
    'BLM': {'size_min': 80,  'size_max': 300, 'size_codes': {'LJFL'}},
    'MLS': {'size_min': 80,  'size_max': 200, 'size_codes': {'LJFL'}},
    'SSP': {'size_min': 70,  'size_max': 200, 'size_codes': {'LJFL'}},
    'SFA': {'size_min': 70,  'size_max': 250, 'size_codes': {'LJFL'}},
    'ISB': {'size_min': 20,  'size_max': 60,  'size_codes': {'FL'}},
    'WAH': {'size_min': 50,  'size_max': 165, 'size_codes': {'FL'}},
    'RZV': {'size_min': 40,  'size_max': 100, 'size_codes': {'TL'}},
    'DOL': {'size_min': 40,  'size_max': 150, 'size_codes': {'FL'}},
    'LEC': {'size_min': 30,  'size_max': 150, 'size_codes': {'FL'}},
    'OIL': {'size_min': 30,  'size_max': 100, 'size_codes': {'FL'}},
    'GES': {'size_min': 40,  'size_max': 150, 'size_codes': {'FL'}},
    'ALX': {'size_min': 40,  'size_max': 150, 'size_codes': {'FL'}},
    'GBA': {'size_min': 40,  'size_max': 150, 'size_codes': {'FL'}},
    'LGH': {'size_min': 30,  'size_max': 60,  'size_codes': {'TL'}},
    'PLS': {'size_min': 40,  'size_max': 65,  'size_codes': {'DW'}},
    'TUG': {'size_min': 30,  'size_max': 80,  'size_codes': {'SCL', 'SCW'}},
    'LKV': {'size_min': 30,  'size_max': 80,  'size_codes': {'SCL', 'SCW'}},  # olivâtre — corrigé (250 était une erreur d'espèce)
    'TTH': {'size_min': 30,  'size_max': 80,  'size_codes': {'SCL', 'SCW'}},
    'TTL': {'size_min': 30,  'size_max': 80,  'size_codes': {'SCL', 'SCW'}},
    'DKK': {'size_min': 30,  'size_max': 200, 'size_codes': {'SCL', 'SCW'}},  # luth — ajouté
}

# Colonnes à surveiller pour les fautes de frappe dans 2_set
# Ce sont des colonnes qui devraient avoir la même valeur pour toutes les calées
COLS_ERREUR_FRAPPE_2SET = [
    'trip_id', 'shooter_used', 'weighted_swivel', 'swivel_weight',
    'lightsticks_per_basket_count', 'branchlines_per_basket_count',
    'baskets_per_section_count', 'total_sections_count',
    'hauling_direction_same_as_setting', 'hauling_breaks', 'time_between_hooks',
    'weighted_snap', 'snap_weight', 'mainline_type', 'lightstick_type',
    'lightstick_color', 'line_instrumented',
]

# Colonnes à surveiller pour les fautes de frappe dans 4_branchline
COLS_ERREUR_FRAPPE_4BRANCHLINE = [
    'branchline_proportion', 'branchline_length_m',
    'branchline_topline_type', 'branchline_traceline_type',
]


# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS DE CHARGEMENT DES DONNÉES
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS UTILITAIRES POUR LES VÉRIFICATIONS
# Ces fonctions sont partagées par toutes les règles de toutes les tables.
# ─────────────────────────────────────────────────────────────────────────────

def _norm_key(val):
    """
    Normalise un set_number ou catch_number en chaîne entière stable.
    Évite le bug pandas float : quand une colonne contient des NaN, pandas
    la charge en float64, donc astype(str) donne '1.0' au lieu de '1',
    ce qui fait échouer les jointures silencieusement.
    '1.0' → '1', '2.0' → '2', etc. Les valeurs non numériques sont
    converties en str brut (fallback).
    """
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val).strip()


def _ref_ligne(df, idx):
    """
    Génère une référence lisible pour identifier une ligne dans les messages.
    Utilise set_number et catch_number si disponibles, sinon le numéro de ligne.
    """
    if 'set_number' in df.columns and 'catch_number' in df.columns:
        set_n   = df.at[idx, 'set_number']   if pd.notna(df.at[idx, 'set_number'])   else '?'
        catch_n = df.at[idx, 'catch_number'] if pd.notna(df.at[idx, 'catch_number']) else '?'
        return T('verif.ref_set_catch', set_n=set_n, catch_n=catch_n)
    elif 'set_number' in df.columns:
        set_n = df.at[idx, 'set_number'] if pd.notna(df.at[idx, 'set_number']) else '?'
        return T('verif.ref_set', set_n=set_n)
    return T('verif.ref_ligne', n=idx + 1)


def _verif_champs_vides(df, nom_table, exclues=None):
    """
    Vérifie que tous les champs obligatoires ne sont pas vides.
    Les colonnes 'topiaid' (identifiants internes) sont toujours ignorées.
    exclues : ensemble de noms de colonnes à ne pas vérifier (champs optionnels)
    """
    anomalies = []
    # On exclut les colonnes passées en paramètre + toutes les colonnes 'topiaid'
    colonnes_ignorees = (exclues or set()) | {c for c in df.columns if 'topiaid' in c}

    for col in df.columns:
        if col in colonnes_ignorees:
            continue
        for idx, val in df[col].items():
            if pd.isna(val) or str(val).strip() == '':
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   nom_table,
                    'ligne':   idx,
                    'colonne': col,
                    'message': T('verif.champ_vide', ref=ref, col=col),
                })
    return anomalies


def _verif_erreur_frappe(df, nom_table, colonnes):
    """
    Détecte les valeurs non-dominantes dans des colonnes qui devraient être homogènes.
    Logique : si une valeur est clairement dominante (> 50% des lignes), toute autre
    valeur présente est signalée comme probable faute de frappe — qu'elle apparaisse
    une ou plusieurs fois.
    """
    anomalies = []

    for col in colonnes:
        if col not in df.columns:
            continue
        valeurs = df[col].dropna()
        # Pas assez de lignes pour détecter une anomalie
        if len(valeurs) < 3:
            continue

        compte = valeurs.value_counts()
        valeur_dominante       = compte.index[0]
        valeur_dominante_count = compte.iloc[0]

        # On n'analyse que si une valeur est clairement dominante (> 50% des lignes)
        if valeur_dominante_count > len(valeurs) / 2:
            for val, count in compte.items():
                if val == valeur_dominante:
                    continue
                # Toutes les lignes avec une valeur non-dominante sont signalées
                for idx in df[df[col] == val].index:
                    ref = _ref_ligne(df, idx)
                    anomalies.append({
                        'niveau':  'Avertissement',
                        'table':   nom_table,
                        'ligne':   idx,
                        'colonne': col,
                        'message': T('verif.erreur_frappe', ref=ref, col=col,
                                     val=val, attendue=valeur_dominante),
                    })
    return anomalies


def _verif_sequence_entiere(df, nom_table, colonne):
    """
    Vérifie que les entiers d'une colonne se suivent sans trou ni doublon (1, 2, 3...).
    Utilisé pour set_number.
    """
    anomalies = []
    if colonne not in df.columns:
        return anomalies
    try:
        valeurs = sorted(df[colonne].dropna().astype(int).tolist())
    except (ValueError, TypeError):
        return anomalies

    if not valeurs:
        return anomalies

    attendu = list(range(1, len(valeurs) + 1))
    if valeurs != attendu:
        manquants = sorted(set(attendu) - set(valeurs))
        doublons  = sorted({v for v in valeurs if valeurs.count(v) > 1})
        parties = []
        if manquants:
            parties.append(T('verif.seq_manquants', manquants=manquants))
        if doublons:
            parties.append(T('verif.seq_doublons', doublons=doublons))
        anomalies.append({
            'niveau':  'Erreur',
            'table':   nom_table,
            'ligne':   None,
            'colonne': colonne,
            'message': T('verif.sequence', col=colonne, details=', '.join(parties)),
        })
    return anomalies


def _verif_sequence_set_id(df, nom_table):
    """
    Vérifie que les set_id se suivent dans l'ordre.
    Le set_id a la forme 'NOMMAREE-N' (ex: SEALORXXX-1, SEALORXXX-2...).
    On extrait le N et on vérifie qu'il est consécutif.
    """
    anomalies = []
    if 'set_id' not in df.columns:
        return anomalies
    try:
        numeros = sorted(int(str(v).split('-')[-1]) for v in df['set_id'].dropna())
    except (ValueError, TypeError):
        return anomalies

    if not numeros:
        return anomalies

    attendu = list(range(1, len(numeros) + 1))
    if numeros != attendu:
        manquants = sorted(set(attendu) - set(numeros))
        doublons  = sorted({v for v in numeros if numeros.count(v) > 1})
        parties = []
        if manquants:
            parties.append(T('verif.seq_manquants', manquants=manquants))
        if doublons:
            parties.append(T('verif.seq_doublons', doublons=doublons))
        anomalies.append({
            'niveau':  'Erreur',
            'table':   nom_table,
            'ligne':   None,
            'colonne': 'set_id',
            'message': T('verif.sequence', col='set_id', details=', '.join(parties)),
        })
    return anomalies


def _verif_plage(df, nom_table, colonne, v_min, v_max, niveau='Erreur'):
    """
    Vérifie que les valeurs numériques d'une colonne sont dans une plage donnée.
    Les cellules vides sont ignorées (elles doivent être gérées par _verif_champs_vides).
    """
    anomalies = []
    if colonne not in df.columns:
        return anomalies
    for idx, ligne in df.iterrows():
        val = ligne.get(colonne)
        if pd.isna(val):
            continue
        try:
            val_num = float(val)
        except (ValueError, TypeError):
            continue
        if val_num < v_min or val_num > v_max:
            ref = _ref_ligne(df, idx)
            anomalies.append({
                'niveau':  niveau,
                'table':   nom_table,
                'ligne':   idx,
                'colonne': colonne,
                'message': T('verif.hors_plage', ref=ref, col=colonne,
                             val=val_num, vmin=v_min, vmax=v_max),
            })
    return anomalies


def _est_vrai(val):
    """
    Retourne True si la valeur représente un booléen vrai (True, 1, "Yes", etc.).
    Utilisé pour les colonnes de type oui/non (weighted_swivel, weighted_snap...).
    """
    if pd.isna(val):
        return False
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val == 1
    return str(val).strip().upper() in {'TRUE', 'YES', 'OUI', '1', 'VRAI'}


BRASSE_M = 1.8288  # 1 brasse = 1,8288 mètres

# Colonne booléenne quasi-numérique (0.0/1.0) — gérée par nom de colonne
# pour éviter de traduire n'importe quel 0/1 numérique.
_COLONNES_BOOL_FLOAT = {'hook_when_discarded'}


def _afficher_val(col, val):
    """
    Convertit une valeur pour l'affichage dans le tableau (uniquement visuel).
    Ne modifie JAMAIS les données brutes dans self.tables.

    Règles appliquées :
    - Valeur vide/NaN      → chaîne vide
    - bool Python True/False → 'Oui'/'Non' (ou 'Yes'/'No' selon la langue)
      (openpyxl lit les booléens Excel comme bool Python, pas comme 0/1 int)
    - hook_when_discarded 1.0/0.0 → idem (cas quasi-booléen connu par son nom)
    - Tout le reste        → str() standard
    """
    if pd.isna(val):
        return ''
    if isinstance(val, bool):
        return T('booleens.oui') if val else T('booleens.non')
    if col in _COLONNES_BOOL_FLOAT and isinstance(val, (int, float)) and val in (0, 1):
        return T('booleens.oui') if int(val) == 1 else T('booleens.non')
    return str(val)


def _en_tete(col, multiline=True):
    """
    Retourne le texte d'en-tête d'une colonne.

    - Si une traduction existe dans colonnes.<col>, l'utilise.
    - Sinon, retombe sur le nom technique avec '_' → '\\n' (fallback propre).
    - Ne plante jamais : T() retourne la clé si absente, on détecte ce cas.
    - Si multiline=True et que le nom traduit dépasse 14 caractères,
      coupe au premier espace pour obtenir un en-tête sur deux lignes.
    """
    traduit = T('colonnes.' + col)
    # T() retourne la clé brute si absente → fallback sur nom technique
    if traduit == 'colonnes.' + col:
        return col.replace('_', '\n') if multiline else col
    if multiline and len(traduit) > 14 and ' ' in traduit:
        idx = traduit.index(' ')
        return traduit[:idx] + '\n' + traduit[idx + 1:]
    return traduit


def _verif_multiple_brasse(df, nom_table, colonne):
    """
    Vérifie que les valeurs d'une colonne sont un multiple de 1,8288 m (1 brasse).
    Les valeurs sont saisies en brasses dans ObServe et converties en mètres à
    l'export. Une valeur non-multiple indique une saisie incorrecte.
    Tolérance de 0,001 m pour les arrondis flottants.
    """
    anomalies = []
    if colonne not in df.columns:
        return anomalies
    for idx, ligne in df.iterrows():
        val = ligne.get(colonne)
        if pd.isna(val):
            continue
        try:
            val_num = float(val)
            if val_num <= 0:
                continue
            ratio = val_num / BRASSE_M
            # Si le ratio n'est pas proche d'un entier → pas un multiple de brasse
            if abs(ratio - round(ratio)) > 0.001:
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   nom_table,
                    'ligne':   idx,
                    'colonne': colonne,
                    'message': T('verif.pas_multiple_brasse', ref=ref,
                                 col=colonne, val=val_num),
                })
        except (ValueError, TypeError):
            continue
    return anomalies


# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS DE VÉRIFICATION PAR TABLE
# Chaque fonction reçoit le dict complet des tables et retourne une liste
# d'anomalies pour la table concernée.
# ─────────────────────────────────────────────────────────────────────────────

def verifier_1_trip(tables):
    """Règles pour 1_trip : champs vides sauf trip_comment."""
    if '1_trip' not in tables:
        return []
    df = tables['1_trip']
    return _verif_champs_vides(df, '1_trip', exclues={'trip_comment'})


def verifier_2_set(tables):
    """
    Règles pour 2_set :
    - Champs vides
    - Fautes de frappe dans les colonnes homogènes
    - set_id et set_number doivent être consécutifs
    - Plages numériques : setting_speed, snap_weight, total_baskets_count,
                          total_hooks_count, sst
    """
    if '2_set' not in tables:
        return []
    df = tables['2_set']
    anomalies = []

    # Champs optionnels (peuvent être remplis selon les missions)
    exclues_2set = {
        'shooter_speed', 'swivel_weight', 'snap_weight',
        'max_depth_targeted', 'time_between_hooks', 'set_comment',
    }
    anomalies += _verif_champs_vides(df, '2_set', exclues_2set)
    anomalies += _verif_erreur_frappe(df, '2_set', COLS_ERREUR_FRAPPE_2SET)
    anomalies += _verif_sequence_set_id(df, '2_set')
    anomalies += _verif_sequence_entiere(df, '2_set', 'set_number')
    if 'weighted_swivel' in df.columns and 'swivel_weight' in df.columns:
        for idx, ligne in df.iterrows():
            if _est_vrai(ligne.get('weighted_swivel')):
                if pd.isna(ligne.get('swivel_weight')) or str(ligne.get('swivel_weight', '')).strip() == '':
                    anomalies.append({
                        'niveau':  'Erreur',
                        'table':   '2_set',
                        'ligne':   idx,
                        'colonne': 'swivel_weight',
                        'message': T('verif.swivel_manquant', ref=_ref_ligne(df, idx)),
                    })

    if 'weighted_snap' in df.columns and 'snap_weight' in df.columns:
        for idx, ligne in df.iterrows():
            if _est_vrai(ligne.get('weighted_snap')):
                if pd.isna(ligne.get('snap_weight')) or str(ligne.get('snap_weight', '')).strip() == '':
                    anomalies.append({
                        'niveau':  'Erreur',
                        'table':   '2_set',
                        'ligne':   idx,
                        'colonne': 'snap_weight',
                        'message': T('verif.snap_manquant', ref=_ref_ligne(df, idx)),
                    })

    anomalies += _verif_plage(df, '2_set', 'setting_speed',        5.5,  8.5,  'Avertissement')
    anomalies += _verif_plage(df, '2_set', 'snap_weight',          0.05, 0.12, 'Erreur')
    anomalies += _verif_plage(df, '2_set', 'total_baskets_count',  1,    300,  'Erreur')
    anomalies += _verif_plage(df, '2_set', 'total_hooks_count',    1,    2500, 'Erreur')
    anomalies += _verif_plage(df, '2_set', 'sst',                  20,   35,   'Avertissement')
    anomalies += _verif_plage(df, '2_set', 'branchline_length_m',       10,  20,  'Avertissement')
    anomalies += _verif_plage(df, '2_set', 'branchlines_per_basket_count', 6,   25,  'Avertissement')
    anomalies += _verif_plage(df, '2_set', 'hauling_speed',                5.5,  8,  'Avertissement')

    # Cohérence total_hooks_count vs branchlines_per_basket_count × total_baskets_count.
    # Les derniers baskets peuvent être partiels : total_hooks ≤ max théorique est normal.
    # Erreur si total_hooks > max (impossible physiquement).
    # Avertissement si l'écart dépasse 2 baskets complets (saisie probablement erronée).
    if all(c in df.columns for c in ['total_hooks_count', 'branchlines_per_basket_count', 'total_baskets_count']):
        for idx, ligne in df.iterrows():
            try:
                hooks      = float(ligne['total_hooks_count'])
                branchlines = float(ligne['branchlines_per_basket_count'])
                baskets    = float(ligne['total_baskets_count'])
            except (TypeError, ValueError):
                continue
            if any(pd.isna(v) for v in [hooks, branchlines, baskets]) or branchlines <= 0:
                continue
            max_theorique = branchlines * baskets
            ref = _ref_ligne(df, idx)
            if hooks > max_theorique:
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   '2_set',
                    'ligne':   idx,
                    'colonne': 'total_hooks_count',
                    'message': T('verif.hooks_impossible', ref=ref,
                                 hooks=int(hooks), branchlines=int(branchlines),
                                 baskets=int(baskets), max=int(max_theorique)),
                })
            elif (max_theorique - hooks) > 2 * branchlines:
                anomalies.append({
                    'niveau':  'Avertissement',
                    'table':   '2_set',
                    'ligne':   idx,
                    'colonne': 'total_hooks_count',
                    'message': T('verif.hooks_ecart', ref=ref,
                                 hooks=int(hooks), branchlines=int(branchlines),
                                 baskets=int(baskets), max=int(max_theorique)),
                })

    # Format et cohérence du set_id : doit être [trip_id]-[set_number] sans zéro devant.
    # Ex: SEALORXXX-3 pour le SET 3 de la marée SEALORXXX. SEALORXXX-03 est incorrect.
    if '1_trip' in tables and 'set_id' in df.columns and 'set_number' in df.columns:
        trip_df = tables['1_trip']
        if 'trip_id' in trip_df.columns and not trip_df.empty:
            trip_id = str(trip_df.iloc[0]['trip_id']).strip()
            prefixe = trip_id + '-'
            for idx, ligne in df.iterrows():
                sid  = str(ligne.get('set_id', '')).strip()
                snum = ligne.get('set_number')
                if not sid or sid == 'nan':
                    continue
                ref = _ref_ligne(df, idx)
                # Vérification 1 : le set_id commence bien par trip_id + '-'
                if not sid.startswith(prefixe):
                    anomalies.append({
                        'niveau':  'Erreur',
                        'table':   '2_set',
                        'ligne':   idx,
                        'colonne': 'set_id',
                        'message': T('verif.set_id_format', ref=ref,
                                     sid=sid, trip_id=trip_id),
                    })
                    continue
                # Vérification 2 : le suffixe est un entier sans zéro devant
                suffixe = sid[len(prefixe):]
                try:
                    num_id = int(suffixe)
                except ValueError:
                    anomalies.append({
                        'niveau':  'Erreur',
                        'table':   '2_set',
                        'ligne':   idx,
                        'colonne': 'set_id',
                        'message': T('verif.set_id_format', ref=ref,
                                     sid=sid, trip_id=trip_id),
                    })
                    continue
                # Vérification 3 : la partie numérique correspond à set_number
                # et il n'y a pas de zéro devant (str(num_id) = suffixe exact)
                try:
                    num_set = int(float(snum))
                except (TypeError, ValueError):
                    continue
                if num_id != num_set or suffixe != str(num_id):
                    anomalies.append({
                        'niveau':  'Erreur',
                        'table':   '2_set',
                        'ligne':   idx,
                        'colonne': 'set_id',
                        'message': T('verif.set_id_incoherent', ref=ref,
                                     sid=sid, set_number=num_set),
                    })

    return anomalies


def verifier_3_floatline(tables):
    """
    Règles pour 3_floatline :
    - Champs vides
    - floatline_length_m : doit être un multiple de 1,8288 m (1 brasse)
    """
    if '3_floatline' not in tables:
        return []
    df = tables['3_floatline']
    anomalies = _verif_champs_vides(df, '3_floatline')
    anomalies += _verif_multiple_brasse(df, '3_floatline', 'floatline_length_m')
    return anomalies


def verifier_4_branchline(tables):
    """
    Règles pour 4_branchline :
    - Champs vides
    - Fautes de frappe dans les colonnes homogènes
    - branchline_length_m : doit être un multiple de 1,8288 m (1 brasse)
    """
    if '4_branchline' not in tables:
        return []
    df = tables['4_branchline']
    anomalies = _verif_champs_vides(df, '4_branchline')
    anomalies += _verif_erreur_frappe(df, '4_branchline', COLS_ERREUR_FRAPPE_4BRANCHLINE)
    anomalies += _verif_multiple_brasse(df, '4_branchline', 'branchline_length_m')
    return anomalies


def verifier_5_hook(tables):
    """Règles pour 5_hook : champs vides."""
    if '5_hook' not in tables:
        return []
    df = tables['5_hook']
    return _verif_champs_vides(df, '5_hook')


def verifier_6_bait(tables):
    """
    Règles pour 6_bait :
    - Champs vides
    - bait_code : valeur inattendue (attendu SQU ou MAZ)
    - bait_setting_status : valeur inattendue (attendu DEF)
    - bait_size : entre 10 et 35
    """
    if '6_bait' not in tables:
        return []
    df = tables['6_bait']
    exclues_6bait = {'bait_weight'}  # optionnel selon les missions
    anomalies = _verif_champs_vides(df, '6_bait', exclues_6bait)

    # bait_code : on attend SQU ou MAZ ; tout autre code est inhabituel
    if 'bait_code' in df.columns:
        for idx, ligne in df.iterrows():
            code = str(ligne.get('bait_code', '')).strip().upper()
            if code and code != 'NAN' and code not in {'SQU', 'MAZ'}:
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Avertissement',
                    'table':   '6_bait',
                    'ligne':   idx,
                    'colonne': 'bait_code',
                    'message': T('verif.bait_code', ref=ref, code=code),
                })

    # bait_setting_status : on attend DEF ; toute autre valeur est inhabituelle
    if 'bait_setting_status' in df.columns:
        for idx, ligne in df.iterrows():
            statut = str(ligne.get('bait_setting_status', '')).strip().upper()
            if statut and statut != 'NAN' and statut != 'DEF':
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Avertissement',
                    'table':   '6_bait',
                    'ligne':   idx,
                    'colonne': 'bait_setting_status',
                    'message': T('verif.bait_status', ref=ref, statut=statut),
                })

    anomalies += _verif_plage(df, '6_bait', 'bait_size', 10, 35, 'Erreur')

    return anomalies


def verifier_7_tdr(tables):
    """
    Règles pour 7_tdr :
    - Champs vides
    - set_number : couvre toute la séquence des calées (pas de trou)
    - set_number : même nombre de TDR pour chaque calée
    - tdr_id : les mêmes TDR doivent être présents pour chaque calée
    - tdr_serial : un tdr_id donné doit toujours avoir le même numéro de série
    - Profondeurs : entre 0 et 400
    - item_horizontal_position : doit être 'Basket'
    - item_vertical_position : doit être 'Hook' ou 'Basse'
    """
    if '7_tdr' not in tables:
        return []
    df = tables['7_tdr']
    # Champs jamais remplis par les observateurs → optionnels
    exclues_7tdr = {
        'section_setting_id', 'section_hauling_id',
        'basket_setting_id',  'basket_hauling_id',
        'data_location',      'data_file_name',
        'fao_code',
    }
    anomalies = _verif_champs_vides(df, '7_tdr', exclues_7tdr)

    # set_number : on vérifie qu'il n'y a pas de calée manquante
    if 'set_number' in df.columns:
        try:
            set_numbers_presents = sorted(df['set_number'].dropna().astype(int).unique().tolist())
            if set_numbers_presents:
                attendu = list(range(1, max(set_numbers_presents) + 1))
                manquants = sorted(set(attendu) - set(set_numbers_presents))
                if manquants:
                    anomalies.append({
                        'niveau':  'Avertissement',
                        'table':   '7_tdr',
                        'ligne':   None,
                        'colonne': 'set_number',
                        'message': T('verif.tdr_manquant', calees=manquants),
                    })
        except (ValueError, TypeError):
            pass

    # set_number : même nombre de TDR pour chaque calée
    if 'set_number' in df.columns:
        compte_par_set = df['set_number'].value_counts()
        if compte_par_set.nunique() > 1:
            # On liste les calées avec un nombre différent du maximum
            nb_max = compte_par_set.max()
            calees_diff = sorted(k for k, v in compte_par_set.items() if v != nb_max)
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '7_tdr',
                'ligne':   None,
                'colonne': 'set_number',
                'message': T('verif.tdr_nb_diff', calees_diff=calees_diff, nb_max=nb_max),
            })

    # tdr_id vide : le champ est obligatoire — l'observateur doit toujours renseigner l'ID TDR
    if 'tdr_id' in df.columns:
        for idx, ligne in df.iterrows():
            val = ligne.get('tdr_id')
            if pd.isna(val) or str(val).strip() == '' or str(val).strip().lower() == 'nan':
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Avertissement',
                    'table':   '7_tdr',
                    'ligne':   idx,
                    'colonne': 'tdr_id',
                    'message': T('verif.tdr_id_vide', ref=ref),
                })

    # tdr_id : les mêmes TDR doivent être présents pour chaque calée.
    # On filtre les NaN avant le frozenset pour éviter les faux positifs
    # (deux NaN ne sont jamais égaux en Python, nan != nan).
    if 'set_number' in df.columns and 'tdr_id' in df.columns:
        tdr_ids_par_set = df.groupby('set_number')['tdr_id'].apply(
            lambda x: frozenset(v for v in x if pd.notna(v))
        )
        sets_uniques = tdr_ids_par_set.unique()
        if len(sets_uniques) > 1:
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '7_tdr',
                'ligne':   None,
                'colonne': 'tdr_id',
                'message': T('verif.tdr_ids_diff'),
            })

    # tdr_serial : un même tdr_id doit toujours avoir le même numéro de série
    if 'tdr_id' in df.columns and 'tdr_serial' in df.columns:
        for tdr_id_val, groupe in df.groupby('tdr_id'):
            serials_uniques = groupe['tdr_serial'].dropna().unique()
            if len(serials_uniques) > 1:
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   '7_tdr',
                    'ligne':   groupe.index[0],
                    'colonne': 'tdr_serial',
                    'message': T('verif.tdr_serial', id=tdr_id_val,
                                 serials=list(serials_uniques)),
                })

    # Profondeurs : toutes les colonnes de profondeur doivent être entre 0 et 400
    cols_profondeur = [
        'fishing_start_depth', 'fishing_end_depth',
        'mean_deployment_depth', 'median_deployment_depth',
        'min_fishing_depth', 'max_fishing_depth',
        'mean_fishing_depth', 'median_fishing_depth',
    ]
    for col in cols_profondeur:
        anomalies += _verif_plage(df, '7_tdr', col, 0, 400, 'Erreur')

    # item_horizontal_position : doit être 'Basket'
    if 'item_horizontal_position' in df.columns:
        for idx, ligne in df.iterrows():
            val = str(ligne.get('item_horizontal_position', '')).strip()
            if val and val.lower() != 'nan' and val != 'Basket':
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   '7_tdr',
                    'ligne':   idx,
                    'colonne': 'item_horizontal_position',
                    'message': T('verif.item_h', ref=ref, val=val),
                })

    # item_vertical_position : 'Hook' (anciens exports zip) ou 'Basse' (exports CSV ObServe récents)
    if 'item_vertical_position' in df.columns:
        for idx, ligne in df.iterrows():
            val = str(ligne.get('item_vertical_position', '')).strip()
            if val and val.lower() != 'nan' and val not in ('Hook', 'Basse'):
                ref = _ref_ligne(df, idx)
                anomalies.append({
                    'niveau':  'Erreur',
                    'table':   '7_tdr',
                    'ligne':   idx,
                    'colonne': 'item_vertical_position',
                    'message': T('verif.item_v', ref=ref, val=val),
                })

    return anomalies


def verifier_8_catch(tables):
    """
    Règles pour 8_catch :
    - Champs vides (sauf status_at_release, hook_when_discarded,
                    depredating_species, catch_comment)
    - fao_code : doit être dans la liste de référence
    - catch_fate RET : espèce doit être autorisée à être retenue
    - catch_fate DIS : status_at_release doit être renseigné
    - status_at_release : si capturé mort (D), ne peut pas avoir un
                          status_at_release différent de D
    - hook_when_discarded : obligatoire si catch_fate = DIS
    """
    if '8_catch' not in tables:
        return []
    df = tables['8_catch']
    anomalies = []

    # Champs obligatoires (les colonnes optionnelles sont listées dans exclues)
    exclues_8catch = {
        'total_weight',           # jamais rempli par les observateurs
        'status_at_release',      # optionnel (sauf si DIS, vérifié plus bas)
        'hook_when_discarded',    # optionnel (sauf si DIS, vérifié plus bas)
        'depredating_species',
        'catch_comment',
    }
    anomalies += _verif_champs_vides(df, '8_catch', exclues_8catch)

    for idx, ligne in df.iterrows():
        code    = str(ligne.get('fao_code', '')).strip().upper()
        fate    = str(ligne.get('catch_fate', '')).strip().upper()
        statut  = str(ligne.get('status_at_capture', '')).strip().upper()
        relacher = ligne.get('status_at_release')
        hook    = ligne.get('hook_when_discarded')
        ref     = f"Calée {ligne.get('set_number', '?')} capture #{ligne.get('catch_number', '?')}"

        # fao_code : valeur inattendue
        if code and code != 'NAN' and code not in CODES_FAO_8CATCH:
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'fao_code',
                'message': T('verif.fao_inattendu', ref=ref, code=code),
            })

        if fate == 'RET' and code and code != 'NAN' and code not in ESPECES_RETENUES_OK:
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'catch_fate',
                'message': T('verif.ret_interdit', ref=ref, code=code),
            })

        if fate == 'DIS' and (pd.isna(relacher) or str(relacher).strip() == ''):
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'status_at_release',
                'message': T('verif.dis_sans_release', ref=ref),
            })

        if (statut == 'D'
                and pd.notna(relacher)
                and str(relacher).strip() != ''
                and str(relacher).strip().upper() != 'D'):
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'status_at_release',
                'message': T('verif.mort_release', ref=ref, val=relacher),
            })

        if fate == 'DIS' and (pd.isna(hook) or str(hook).strip() == ''):
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'hook_when_discarded',
                'message': T('verif.dis_sans_hook', ref=ref),
            })

        # Un animal mort ne peut pas s'être échappé — combinaison biologiquement impossible
        if statut == 'D' and fate == 'ESC':
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'catch_fate',
                'message': T('verif.mort_escape', ref=ref),
            })

    # ── Rejet suspect : espèce commercialisable, sans déprédation, avec mensuration ──
    # Si une espèce cible est rejetée (DIS/ESC) sans déprédation ET qu'une mensuration
    # existe dans 9_sample à taille normale → probablement une erreur de saisie (DIS au lieu de RET).
    # Si aucune mensuration → poisson non remonté à bord → on ignore.
    sample_df = tables.get('9_sample')
    if sample_df is not None and not sample_df.empty:
        # Construire un ensemble des (set_number, catch_number) présents dans 9_sample
        cles_sample = set(
            zip(sample_df['set_number'].apply(_norm_key),
                sample_df['catch_number'].apply(_norm_key))
        )
        for idx, ligne in df.iterrows():
            fate = str(ligne.get('catch_fate', '')).strip().upper()
            code = str(ligne.get('fao_code',   '')).strip().upper()
            depred = ligne.get('depredating_species')

            # Ne concerne que les rejets d'espèces cibles commercialisables
            if fate not in ('DIS', 'ESC'):
                continue
            if code not in CODES_FAO_CIBLES:
                continue
            # Ignorer si déprédation présente
            if pd.notna(depred) and str(depred).strip() not in ('', 'NAN'):
                continue

            # Vérifier si une mensuration existe dans 9_sample pour cette capture
            set_n   = _norm_key(ligne.get('set_number',   ''))
            catch_n = _norm_key(ligne.get('catch_number', ''))
            if (set_n, catch_n) not in cles_sample:
                continue  # pas de mensuration → poisson non remonté → on ignore

            # Vérifier que la taille est dans la plage commerciale
            ref_espece = CAPTURE_MEASUREMENT.get(code)
            if ref_espece is None:
                continue  # espèce sans référence de taille → on ignore
            lignes_sample = sample_df[
                (sample_df['set_number'].apply(_norm_key)   == set_n) &
                (sample_df['catch_number'].apply(_norm_key) == catch_n)
            ]
            for _, s_ligne in lignes_sample.iterrows():
                try:
                    taille = float(s_ligne.get('size'))
                except (TypeError, ValueError):
                    continue
                if ref_espece['size_min'] <= taille <= ref_espece['size_max']:
                    ref = f"Calée {set_n} capture #{catch_n}"
                    anomalies.append({
                        'niveau':  'Avertissement',
                        'table':   '8_catch',
                        'ligne':   idx,
                        'colonne': 'catch_fate',
                        'message': (
                            f"{ref} [{code}] : rejeté ({fate}) sans déprédation "
                            f"avec une taille de {taille:.0f} cm (plage commerciale "
                            f"{ref_espece['size_min']}–{ref_espece['size_max']} cm). "
                            f"Vérifiez si le devenir est correct."
                        ),
                    })
                    break  # un seul avertissement par capture

    # ── Doublon de capture consécutive (même espèce, même statut, même taille) ──
    # Deux captures consécutives (catch_number N et N+1, même calée) avec exactement
    # les mêmes champs biologiques ET la même taille dans 9_sample → probable doublon de saisie.
    # Si l'une des deux n'a pas de mensuration → on ignore (on ne peut pas comparer).
    if sample_df is not None and not sample_df.empty:
        # Construire un dict (set_number, catch_number) → (size, size_code) depuis 9_sample
        tailles_sample = {}
        for _, s in sample_df.iterrows():
            # _norm_key évite le bug float64 : NaN dans la colonne → pandas type float
            # → astype(str) donne '1.0' au lieu de '1' → jointure silencieuse échoue
            cle = (_norm_key(s.get('set_number', '')), _norm_key(s.get('catch_number', '')))
            tailles_sample[cle] = (
                str(s.get('size',      '')).strip(),
                str(s.get('size_code', '')).strip().upper(),
            )

        # Trier 8_catch par calée puis par numéro de capture
        df_tri = df.sort_values(['set_number', 'catch_number']).reset_index()
        COLS_COMP = ['fao_code', 'catch_fate', 'status_at_capture',
                     'status_at_release', 'depredating_species']

        for i in range(len(df_tri) - 1):
            lig_a = df_tri.iloc[i]
            lig_b = df_tri.iloc[i + 1]

            # Les deux doivent être sur la même calée et avoir des catch_number consécutifs
            if _norm_key(lig_a.get('set_number', '')) != _norm_key(lig_b.get('set_number', '')):
                continue
            try:
                if int(lig_b['catch_number']) - int(lig_a['catch_number']) != 1:
                    continue
            except (ValueError, TypeError):
                continue

            # Comparer les champs biologiques
            def _val(lig, col):
                v = lig.get(col, '')
                return '' if pd.isna(v) else str(v).strip().upper()

            if any(_val(lig_a, c) != _val(lig_b, c) for c in COLS_COMP):
                continue

            # Comparer les tailles dans 9_sample
            set_n = _norm_key(lig_a.get('set_number', ''))
            cle_a = (set_n, _norm_key(lig_a['catch_number']))
            cle_b = (set_n, _norm_key(lig_b['catch_number']))
            if cle_a not in tailles_sample or cle_b not in tailles_sample:
                continue  # au moins une capture sans mensuration → on ignore
            if tailles_sample[cle_a] != tailles_sample[cle_b]:
                continue  # tailles différentes → captures distinctes

            # Toutes les conditions remplies → doublon probable
            size_val, size_cod = tailles_sample[cle_a]
            ref_a = f"Calée {set_n} capture #{lig_a['catch_number']}"
            ref_b = f"capture #{lig_b['catch_number']}"
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   int(lig_a['index']),
                'colonne': 'catch_number',
                'message': (
                    f"{ref_a} et {ref_b} [{_val(lig_a, 'fao_code')}] : "
                    f"captures consécutives identiques ({_val(lig_a, 'catch_fate')}, "
                    f"{_val(lig_a, 'status_at_capture')}) avec la même taille "
                    f"{size_val} {size_cod}. Probable doublon de saisie."
                ),
            })

    # ── Numéros de capture non consécutifs ──────────────────────────────────
    # Pour chaque calée, les catch_number doivent se suivre sans saut ni répétition.
    # Cas 1 — Saut : 1, 2, 4, 5 (le n°3 manque entre deux captures successives)
    # Cas 2 — Mauvais ordre ou doublon : 1, 2, 5, 3 ou 1, 2, 2, 3
    # On ne vérifie pas si la séquence commence à 1 (la première calée peut
    # commencer à n'importe quel numéro selon la configuration ObServe).
    for set_n, groupe in df.groupby('set_number', sort=False):
        # Convertir les catch_number en entiers (ignorer les valeurs non convertibles)
        nums_int = []
        for v in groupe['catch_number'].tolist():
            try:
                nums_int.append(int(v))
            except (ValueError, TypeError):
                continue

        if len(nums_int) < 2:
            continue  # une seule capture → pas de séquence à vérifier

        # Cas 2 : doublon de numéro (même catch_number utilisé deux fois dans la même calée)
        if len(set(nums_int)) != len(nums_int):
            nums_str = ', '.join(str(n) for n in nums_int)
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   int(groupe.index[0]),
                'colonne': 'catch_number',
                'message': (
                    f"Calée {set_n} : numéros de capture en doublon ({nums_str}). "
                    f"Vérifier la saisie dans ObServe."
                ),
            })
            continue  # doublon masquerait l'analyse de saut, on s'arrête là

        # Cas 1 : saut dans la séquence (ex: 1, 2, 4 → manque 3)
        nums_tries = sorted(nums_int)
        saut = None
        for i in range(len(nums_tries) - 1):
            if nums_tries[i + 1] - nums_tries[i] != 1:
                saut = (nums_tries[i], nums_tries[i + 1])
                break

        if saut:
            nums_str = ', '.join(str(n) for n in nums_tries)
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   int(groupe.index[0]),
                'colonne': 'catch_number',
                'message': (
                    f"Calée {set_n} : saut dans les numéros de capture entre "
                    f"#{saut[0]} et #{saut[1]} (séquence : {nums_str})."
                ),
            })
            continue

        # Cas 2 : mauvais ordre (tous les numéros sont là mais dans le désordre)
        if nums_int != sorted(nums_int):
            nums_str = ', '.join(str(n) for n in nums_int)
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   int(groupe.index[0]),
                'colonne': 'catch_number',
                'message': (
                    f"Calée {set_n} : numéros de capture dans le mauvais ordre "
                    f"({nums_str}). Vérifier la saisie dans ObServe."
                ),
            })

    # ── Mensuration manquante dans 9_sample ─────────────────────────────────
    # Pour chaque capture dont on attendrait une mensuration, vérifier qu'elle
    # a bien au moins une ligne dans 9_sample.
    # On exclut :
    #   - catch_fate = ESC : poisson non remonté à bord → impossible de mesurer
    #   - déprédation ODN ou SKH : poisson trop endommagé pour être mesuré
    #   - espèces "hors gabarit" (CODES_HORS_GABARIT) : trop grandes pour être
    #     remontées à bord (requins, raies sauf PLS, cétacés, requin baleine)
    sample_df_mes = tables.get('9_sample')

    # Si 9_sample est absent ou vide, on émet un seul avertissement global
    # et on saute la vérification individuelle des mensurations manquantes.
    if sample_df_mes is None or sample_df_mes.empty:
        anomalies.append({
            'niveau':  'Avertissement',
            'table':   '8_catch',
            'ligne':   None,
            'colonne': None,
            'message': (
                "La table 9_sample est absente ou vide — "
                "les mensurations n'ont pas pu être vérifiées."
            ),
        })
    else:
        # Construire l'ensemble des clés (set_number, catch_number) dans 9_sample.
        # _norm_key évite le bug float64 de pandas : NaN dans une colonne entière
        # force le type float, et astype(str) donne '1.0' au lieu de '1'.
        cles_mesures = set(
            zip(sample_df_mes['set_number'].apply(_norm_key),
                sample_df_mes['catch_number'].apply(_norm_key))
        )

        # Codes de déprédation qui justifient l'absence de mesure.
        # ODN = orques (Orcinus orca), SKH = requins divers, FAW = fausse orque
        # (Pseudorca crassidens, déprédateur #1 en océan Indien), GLO = globicéphale.
        DEPRED_EXCL_MES = {'ODN', 'SKH', 'FAW', 'GLO'}

        for idx, ligne in df.iterrows():
            fate = str(ligne.get('catch_fate', '')).strip().upper()
            if fate == 'ESC':
                continue  # poisson jamais remonté à bord → pas de mesure possible

            code = str(ligne.get('fao_code', '')).strip().upper()
            if code in CODES_HORS_GABARIT:
                continue  # espèce trop grande pour être mesurée à bord

            # Les exports CSV récents peuvent lister plusieurs prédateurs : 'ISB;ODN'
            depred = str(ligne.get('depredating_species', '')).strip().upper()
            codes_depred = {c.strip() for c in depred.split(';')}
            if codes_depred & DEPRED_EXCL_MES:
                continue  # poisson trop endommagé pour être mesuré

            set_n   = _norm_key(ligne.get('set_number',   ''))
            catch_n = _norm_key(ligne.get('catch_number', ''))
            if (set_n, catch_n) in cles_mesures:
                continue  # mensuration présente → OK

            ref = f"Calée {set_n} capture #{catch_n} [{code}]"
            anomalies.append({
                'niveau':  'Avertissement',
                'table':   '8_catch',
                'ligne':   idx,
                'colonne': 'catch_number',
                'message': (
                    f"{ref} ({fate}) : aucune mensuration dans 9_sample. "
                    f"L'observateur a peut-être oublié de saisir la taille."
                ),
            })

    return anomalies


def verifier_9_sample(tables):
    """
    Règles pour 9_sample :
    - Champs vides (sauf weight, weight_code, sex, maturity, catch_comment)
    - size : doit être dans les limites de la table Capture measurement
    - size_code : doit correspondre au code attendu pour l'espèce
    """
    if '9_sample' not in tables:
        return []
    df = tables['9_sample']
    anomalies = []

    exclues_9sample = {'weight', 'weight_code', 'sex', 'maturity', 'catch_comment'}
    anomalies += _verif_champs_vides(df, '9_sample', exclues_9sample)

    for idx, ligne in df.iterrows():
        code      = str(ligne.get('fao_code', '')).strip().upper()
        ref_espece = CAPTURE_MEASUREMENT.get(code)

        # Pas de données de référence pour cette espèce : on ne peut pas vérifier
        if not ref_espece:
            continue

        ref = f"Calée {ligne.get('set_number', '?')} capture #{ligne.get('catch_number', '?')} [{code}]"

        # Vérification de la taille (en cm)
        taille = ligne.get('size')
        if pd.notna(taille):
            try:
                taille_num = float(taille)
                if taille_num < ref_espece['size_min']:
                    anomalies.append({
                        'niveau':  'Avertissement',
                        'table':   '9_sample',
                        'ligne':   idx,
                        'colonne': 'size',
                        'message': T('verif.taille_petite', ref=ref, val=taille_num,
                                     min_val=ref_espece['size_min']),
                    })
                elif taille_num > ref_espece['size_max']:
                    anomalies.append({
                        'niveau':  'Avertissement',
                        'table':   '9_sample',
                        'ligne':   idx,
                        'colonne': 'size',
                        'message': T('verif.taille_grande', ref=ref, val=taille_num,
                                     max_val=ref_espece['size_max']),
                    })
            except (ValueError, TypeError):
                pass

        # Vérification du code de mesure
        size_code_saisi = str(ligne.get('size_code', '')).strip().upper()
        codes_valides   = {c.upper() for c in ref_espece['size_codes']}
        if size_code_saisi and size_code_saisi != 'NAN' and size_code_saisi not in codes_valides:
            anomalies.append({
                'niveau':  'Erreur',
                'table':   '9_sample',
                'ligne':   idx,
                'colonne': 'size_code',
                'message': T('verif.size_code', ref=ref, code=size_code_saisi,
                             attendus=', '.join(sorted(ref_espece['size_codes']))),
            })

    # ── Doublon de taille consécutive ────────────────────────────────────────
    # Si deux mensurations consécutives (catch_number N et N+1, même calée) ont
    # exactement la même size ET le même size_code → probable copie par erreur de
    # la taille du poisson précédent (comportement courant dans l'interface ObServe).
    df_tri = df.sort_values(['set_number', 'catch_number']).reset_index()
    for i in range(len(df_tri) - 1):
        lig_a = df_tri.iloc[i]
        lig_b = df_tri.iloc[i + 1]

        # Les deux doivent être sur la même calée
        if _norm_key(lig_a.get('set_number', '')) != _norm_key(lig_b.get('set_number', '')):
            continue
        # Les catch_number doivent être consécutifs
        try:
            if int(lig_b['catch_number']) - int(lig_a['catch_number']) != 1:
                continue
        except (ValueError, TypeError):
            continue

        # Comparer size et size_code
        size_a = str(lig_a.get('size',      '')).strip()
        size_b = str(lig_b.get('size',      '')).strip()
        cod_a  = str(lig_a.get('size_code', '')).strip().upper()
        cod_b  = str(lig_b.get('size_code', '')).strip().upper()

        if not size_a or size_a == 'NAN' or not size_b or size_b == 'NAN':
            continue  # taille absente → on ne peut pas comparer
        if size_a != size_b or cod_a != cod_b:
            continue  # tailles différentes → captures distinctes

        set_n = _norm_key(lig_a.get('set_number', ''))
        ref_a = f"Calée {set_n} capture #{lig_a['catch_number']}"
        ref_b = f"capture #{lig_b['catch_number']}"
        anomalies.append({
            'niveau':  'Avertissement',
            'table':   '9_sample',
            'ligne':   int(lig_a['index']) if 'index' in lig_a else i,
            'colonne': 'size',
            'message': (
                f"{ref_a} et {ref_b} : taille identique consécutive "
                f"{size_a} {cod_a}. "
                f"Probable copie de la taille du poisson précédent."
            ),
        })

    return anomalies


# ─────────────────────────────────────────────────────────────────────────────
# TABLE DE DISPATCH — associe chaque table à sa fonction de vérification
# Pour ajouter une règle : modifier la fonction correspondante ci-dessus.
# Pour ajouter une nouvelle table : créer la fonction et l'ajouter ici.
# ─────────────────────────────────────────────────────────────────────────────

VERIFICATEURS = {
    '1_trip':      verifier_1_trip,
    '2_set':       verifier_2_set,
    '3_floatline': verifier_3_floatline,
    '4_branchline': verifier_4_branchline,
    '5_hook':      verifier_5_hook,
    '6_bait':      verifier_6_bait,
    '7_tdr':       verifier_7_tdr,
    '8_catch':     verifier_8_catch,
    '9_sample':    verifier_9_sample,
}


# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE GRAPHIQUE
# Construite avec tkinter, disponible sans installation supplémentaire.
# ─────────────────────────────────────────────────────────────────────────────

class AppObserveTools(tk.Frame):
    """
    Interface principale de la Phase 1 (correction des données ObServe).
    Hérite de tk.Frame pour pouvoir être embarquée dans l'application unifiée
    (app.py) ou lancée seule via lancer_seul().
    """

    def __init__(self, parent, mode='standalone'):
        super().__init__(parent, bg=COULEURS['fond'])
        # mode='standalone' : fenêtre indépendante avec tous les boutons (import, référentiel, langue)
        # mode='embed'      : embarqué dans app.py — ces éléments sont gérés par la page Accueil
        self._mode = mode

        # Variables internes
        self.tables           = {}   # dict { nom_table: DataFrame }
        self.anomalies        = []   # anomalies filtrées pour la table affichée
        self.anomalies_global = []   # toutes les anomalies toutes tables confondues
        self._anomalies_filtrees = []  # anomalies visibles dans la liste (après filtre niveau/table)
        self._noms_tables     = []   # liste ordonnée des noms de tables chargées
        self._chemin_fichier        = None # chemin complet du zip importé (pour localiser le .json annotations)
        self.annotations            = {}   # dict { clé_anomalie: {statut, commentaire} }
        self._zone_peche            = ''   # zone(s) de pêche — renseignée par app.py via set_zone_peche()
        self._zone_peche_getter     = None # getter optionnel (callable) pour lire la zone en temps réel
        self._rapport_exporte       = True # False si des annotations existent mais le rapport n'a pas été exporté
        self._anomalie_selectionnee = None # anomalie actuellement sélectionnée dans la liste
        self.nom_maree = tk.StringVar(value=T('app.aucune_maree'))
        self.statut    = tk.StringVar(value=T('app.demarrage'))
        # Correspondance index DataFrame → numéro de ligne dans tksheet (base 0)
        self._idx_map        = {}
        # Cellule actuellement surlignée : (df_index, nom_colonne) ou None
        self._cellule_marquee = None

        # Construction de l'interface
        self._construire_interface()

        # Vérification du référentiel au démarrage — seulement en mode standalone
        # En mode embed, c'est la page Accueil (app.py) qui gère le référentiel
        if self._mode == 'standalone':
            self.after(1500, self._lancer_verification_referentiel)

    def _construire_interface(self):
        """Construit tous les éléments visuels de la fenêtre."""

        # ── En-tête ──────────────────────────────────────────────────────────
        en_tete = tk.Frame(self, bg=COULEURS['en_tete'], height=52)
        en_tete.pack(fill='x')
        en_tete.pack_propagate(False)

        self._logo_titre = ajouter_titre_bandeau(
            en_tete, COULEURS['en_tete'], T('app.titre')
        )

        self._logos = ajouter_logos_bandeau(en_tete, COULEURS['en_tete'])

        # Sélecteur de langue (FR / EN) — uniquement en mode standalone
        # En mode embed, le sélecteur est dans la page Accueil
        if self._mode == 'standalone':
            self.var_langue = tk.StringVar(value=lang.langue().upper())
            menu_langue = ttk.Combobox(
                en_tete, textvariable=self.var_langue,
                values=['FR', 'EN'], state='readonly', width=4,
                font=(POLICE, 9)
            )
            menu_langue.pack(side='right', padx=(0, 8))
            menu_langue.bind('<<ComboboxSelected>>', self._changer_langue)
            tk.Label(
                en_tete, text=T('boutons.langue') + ' :',
                bg=COULEURS['en_tete'], fg='#C7D2FE', font=(POLICE, 9)
            ).pack(side='right')

        # ── Barre de boutons ─────────────────────────────────────────────────
        barre = tk.Frame(self, bg=COULEURS['fond'])
        barre.pack(fill='x', padx=10, pady=8)

        # Bouton Import — uniquement en mode standalone
        if self._mode == 'standalone':
            tk.Button(
                barre, text=T('boutons.importer'),
                command=self._importer_fichier,
                bg=COULEURS['accent'], fg='white',
                font=(POLICE, 10, 'bold'), relief='flat', padx=12,
                cursor='hand2'
            ).pack(side='left', padx=4)

        self._btn_verifier_tout = ttk.Button(
            barre, text=T('boutons.verifier_tout'),
            command=self._lancer_verification_tout,
            style='Violet.TButton',
            state='disabled',
            cursor='hand2'
        )
        self._btn_verifier_tout.pack(side='left', padx=4)

        # Bouton export rapport de vérification annoté — Excel professionnel verrouillé
        self._btn_rapport = ttk.Button(
            barre, text=T('boutons.rapport_verifcation'),
            command=self._exporter_rapport_annote,
            style='Amber.TButton',
            state='disabled',
            cursor='hand2'
        )
        self._btn_rapport.pack(side='left', padx=4)

        # Bouton Référentiel — uniquement en mode standalone
        if self._mode == 'standalone':
            tk.Button(
                barre, text=T('boutons.referentiel'),
                command=self._ouvrir_referentiel,
                bg=COULEURS['blanc'], fg=COULEURS['texte'],
                font=(POLICE, 10), relief='groove', padx=12,
                cursor='hand2'
            ).pack(side='left', padx=4)

        # ── Zone centrale : tableau à gauche, anomalies à droite ─────────────
        # PanedWindow horizontal — l'utilisateur peut glisser la barre de séparation
        corps = ttk.PanedWindow(self, orient='horizontal')
        corps.pack(fill='both', expand=True, padx=10, pady=4)

        # Panneau droit — liste des anomalies (déclaré en premier pour les refs internes)
        droite = tk.Frame(corps, bg=COULEURS['blanc'], relief='groove', bd=1)

        tk.Label(
            droite, text=T('anomalies.titre'),
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 10, 'bold')
        ).pack(fill='x', pady=0)

        # Filtre par niveau
        filtre_frame = tk.Frame(droite, bg=COULEURS['blanc'])
        filtre_frame.pack(fill='x', padx=6, pady=4)
        tk.Label(filtre_frame, text=T('anomalies.filtre_label'),
                 bg=COULEURS['blanc'], font=(POLICE, 9)).pack(side='left')
        # Les valeurs du filtre correspondent aux niveaux canoniques internes
        self._filtres_canon  = ['Tout',          'Erreur',                 'Avertissement',        'Info']
        self._filtres_affiches = [T('anomalies.filtre_tout'), T('anomalies.filtre_erreur'),
                                   T('anomalies.filtre_avert'), T('anomalies.filtre_info')]
        self.var_filtre = tk.StringVar(value=self._filtres_affiches[0])
        menu_filtre = ttk.Combobox(
            filtre_frame, textvariable=self.var_filtre,
            values=self._filtres_affiches,
            state='readonly', width=14
        )
        menu_filtre.pack(side='left', padx=4)
        menu_filtre.bind('<<ComboboxSelected>>', lambda e: self._afficher_anomalies())

        # Liste des anomalies + scrollbar
        cadre_liste = tk.Frame(droite, bg=COULEURS['blanc'])
        cadre_liste.pack(fill='both', expand=True, padx=4, pady=4)
        scroll_anomalies = tk.Scrollbar(cadre_liste)
        scroll_anomalies.pack(side='right', fill='y')
        self.liste_anomalies = tk.Listbox(
            cadre_liste, bg=COULEURS['blanc'], activestyle='dotbox',
            font=(POLICE, 9), selectmode='single',
            selectbackground=COULEURS['accent'],
            selectforeground='white',
            yscrollcommand=scroll_anomalies.set,
        )
        self.liste_anomalies.pack(side='left', fill='both', expand=True)
        scroll_anomalies.config(command=self.liste_anomalies.yview)
        self.liste_anomalies.bind('<<ListboxSelect>>', self._sur_selection_anomalie)
        # Flèches ↑/↓ : naviguer sans souris (utile quand le clavier est occupé à saisir un commentaire)
        self.liste_anomalies.bind('<Up>',   lambda e: self._naviguer_anomalie(-1))
        self.liste_anomalies.bind('<Down>',  lambda e: self._naviguer_anomalie(+1))
        # Raccourcis annotation : C = correct (+ suivant), V = justifié (+ focus commentaire)
        # Raccourcis = initiales des statuts : V = Validé, J = Justifié
        self.liste_anomalies.bind('<Key-v>', lambda e: self._raccourci_correct())
        self.liste_anomalies.bind('<Key-V>', lambda e: self._raccourci_correct())
        self.liste_anomalies.bind('<Key-j>', lambda e: self._raccourci_justifier())
        self.liste_anomalies.bind('<Key-J>', lambda e: self._raccourci_justifier())

        # Détail de l'anomalie sélectionnée
        self.detail_anomalie = tk.Label(
            droite, text=T('anomalies.cliquer_detail'),
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9), anchor='w', justify='left', wraplength=340
        )
        self.detail_anomalie.pack(fill='x', padx=6, pady=4)

        # ── Panneau annotations ───────────────────────────────────────────────
        # Toujours visible, boutons grisés tant qu'aucune anomalie n'est sélectionnée
        cadre_annot = tk.Frame(droite, bg=COULEURS['blanc'])
        cadre_annot.pack(fill='x', padx=4, pady=(0, 4))

        # Titre de section
        tk.Label(
            cadre_annot, text=T('annotations.titre'),
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9, 'bold'), anchor='w'
        ).pack(fill='x', padx=4, pady=(4, 0))
        ttk.Separator(cadre_annot, orient='horizontal').pack(fill='x', padx=4, pady=(2, 4))

        # Ligne de boutons
        ligne_btn = tk.Frame(cadre_annot, bg=COULEURS['blanc'])
        ligne_btn.pack(fill='x', padx=4, pady=4)

        self.btn_valide = ttk.Button(
            ligne_btn, text=T('annotations.btn_valide'),
            command=self._annoter_valide,
            style='OK.TButton',
            state='disabled',
            cursor='hand2'
        )
        self.btn_valide.pack(side='left', padx=(0, 4))

        self.btn_comment = ttk.Button(
            ligne_btn, text=T('annotations.btn_commentaire'),
            command=self._annoter_commentaire,
            style='Accent.TButton',
            state='disabled',
            cursor='hand2'
        )
        self.btn_comment.pack(side='left', padx=4)

        # Bouton effacer — affiché seulement quand une annotation existe (géré dans _mettre_a_jour_panneau_annotation)
        self.btn_effacer = tk.Button(
            ligne_btn, text='✕',
            command=self._effacer_annotation,
            bg=COULEURS['erreur'], fg='white',
            font=(POLICE, 9, 'bold'), relief='flat', padx=6,
            state='disabled', disabledforeground='#AAAAAA',
            cursor='hand2'
        )
        # btn_effacer n'est pas pack() ici — affiché dynamiquement selon l'état de l'annotation

        # Zone de texte commentaire — cachée par défaut, visible quand statut = 'justifié'
        self.frame_comment = tk.Frame(cadre_annot, bg=COULEURS['blanc'])
        # frame_comment n'est pas pack() ici — affiché dynamiquement

        self.txt_commentaire = tk.Text(
            self.frame_comment, height=3,
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9), relief='solid', borderwidth=1,
            wrap='word', undo=True
        )
        self.txt_commentaire.pack(fill='x', padx=4, pady=(0, 4))
        # Sauvegarde automatique quand l'observateur clique ailleurs
        self.txt_commentaire.bind('<FocusOut>', self._sauvegarder_commentaire)
        # Échap : sauvegarder et remettre le focus sur la liste pour reprendre la navigation clavier
        self.txt_commentaire.bind('<Escape>', lambda e: self.liste_anomalies.focus_set())
        # Tab : sauvegarder le commentaire et passer à l'anomalie suivante
        self.txt_commentaire.bind('<Tab>', self._tab_vers_suivante)

        # Panneau gauche — tableau des données
        gauche = tk.Frame(corps, bg=COULEURS['fond'])

        # Ajout dans l'ordre gauche → droite, avec poids pour le ratio initial
        corps.add(gauche, weight=3)
        corps.add(droite, weight=1)

        # Sélecteur de table
        sel_frame = tk.Frame(gauche, bg=COULEURS['fond'])
        sel_frame.pack(fill='x', pady=(0, 4))
        tk.Label(sel_frame, text=T('tableau.table_affichee'), bg=COULEURS['fond'], font=(POLICE, 10)).pack(side='left')
        self.var_table = tk.StringVar()
        self.menu_table = ttk.Combobox(sel_frame, textvariable=self.var_table, state='readonly', width=30)
        self.menu_table.pack(side='left', padx=6)
        self.menu_table.bind('<<ComboboxSelected>>', self._changer_table)

        # Tableau tksheet — remplace Treeview pour permettre le surlignage de cellules
        cadre_tableau = tk.Frame(gauche, bg=COULEURS['blanc'], relief='groove', bd=1)
        cadre_tableau.pack(fill='both', expand=True)

        self.tableau = tksheet.Sheet(
            cadre_tableau,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            header_height=52,
            row_height=22,
            font=(POLICE, 9, 'normal'),
            header_font=(POLICE, 9, 'bold'),
            show_row_index=False,
        )
        self.tableau.enable_bindings(
            'single_select', 'column_width_resize',
            'arrowkeys', 'column_select', 'row_select',
        )
        self.tableau.pack(fill='both', expand=True)

        # ── Barre de statut ───────────────────────────────────────────────────
        barre_statut = tk.Frame(self, bg=COULEURS['en_tete'], height=26)
        barre_statut.pack(fill='x', side='bottom')
        barre_statut.pack_propagate(False)

        self._lbl_dot = tk.Label(barre_statut, text='●', bg=COULEURS['en_tete'], fg='#94A3B8', font=(POLICE, 9))
        self._lbl_dot.pack(side='left', padx=(10, 2))
        tk.Label(
            barre_statut, textvariable=self.statut,
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 9), anchor='w'
        ).pack(side='left', padx=(0, 10))

        def _maj_dot(*args):
            msg = self.statut.get()
            m = msg.lower()
            if '✗' in msg or 'erreur' in m or 'impossible' in m:
                fg = '#F87171'
            elif '⚠' in msg or 'avertissement' in m:
                fg = '#FBB76A'
            elif 'cours' in m or 'chargement' in m or 'vérification' in m:
                fg = '#93C5FD'
            elif '✓' in msg or 'chargé' in m or 'exporté' in m or 'prêt' in m or 'terminé' in m:
                fg = '#4ADE80'
            else:
                fg = '#94A3B8'
            self._lbl_dot.config(fg=fg)
        self.statut.trace_add('write', _maj_dot)

        self.lbl_erreurs = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#FF9999')
        self.lbl_erreurs.pack(side='right', padx=(4, 10))
        self.lbl_avert = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#FFD080')
        self.lbl_avert.pack(side='right', padx=(4, 4))
        # Compteurs d'annotations : à traiter / justifié / correct (packés droite→gauche)
        self.lbl_a_traiter = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#FFFFFF',
                                      font=(POLICE, 9))
        self.lbl_a_traiter.pack(side='right', padx=(4, 12))
        self.lbl_justifie = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#93C5FD',
                                     font=(POLICE, 9))
        self.lbl_justifie.pack(side='right', padx=(4, 2))
        self.lbl_correct = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#4ADE80',
                                    font=(POLICE, 9))
        self.lbl_correct.pack(side='right', padx=(4, 2))
        # Indicateur annotations non exportées
        self.lbl_non_exporte = tk.Label(barre_statut, text='', bg=COULEURS['en_tete'], fg='#FFD080',
                                        font=(POLICE, 9))
        self.lbl_non_exporte.pack(side='right', padx=(4, 4))

    # ── Méthodes de gestion des actions ──────────────────────────────────────

    # ── Gestion du référentiel ObServe ───────────────────────────────────────

    def _lancer_verification_referentiel(self):
        """Lance la vérification du référentiel dans un thread pour ne pas bloquer l'interface."""
        thread = threading.Thread(target=self._thread_referentiel, daemon=True)
        thread.start()

    def _thread_referentiel(self):
        """
        Exécuté en arrière-plan : vérifie la connexion internet et l'état
        du référentiel local, puis planifie l'affichage de la popup dans
        le thread principal.
        """
        date_locale = referentiel.date_referentiel_local()

        if not referentiel.verifier_internet():
            if date_locale is None:
                self.after(0, lambda: messagebox.showwarning(
                    T('ref_popup.titre_hors_ligne'),
                    T('ref_popup.message_hors_ligne')
                ))
            return

        # Internet disponible
        if date_locale is None:
            # Premier lancement : référentiel absent, téléchargement nécessaire
            self.after(0, self._popup_telechargement_obligatoire)
        else:
            # Référentiel présent : proposer la mise à jour
            self.after(0, lambda d=date_locale: self._popup_proposer_mise_a_jour(d))

    def _popup_telechargement_obligatoire(self):
        reponse = messagebox.askokcancel(
            T('ref_popup.titre_manquant'),
            T('ref_popup.message_manquant')
        )
        if reponse:
            self._telecharger_avec_progression()

    def _popup_proposer_mise_a_jour(self, date_locale):
        reponse = messagebox.askyesno(
            T('ref_popup.titre_maj'),
            T('ref_popup.message_maj', date=date_locale)
        )
        if reponse:
            self._telecharger_avec_progression()

    def _telecharger_avec_progression(self):
        """
        Affiche une fenêtre de progression pendant le téléchargement
        du référentiel. Le téléchargement tourne dans un thread séparé.
        """
        # Fenêtre modale de progression
        fenetre = tk.Toplevel(self)
        fenetre.title("Mise à jour du référentiel")
        fenetre.geometry("420x130")
        fenetre.resizable(False, False)
        fenetre.grab_set()  # bloque la fenêtre principale pendant le téléchargement
        fenetre.configure(bg=COULEURS['fond'])

        tk.Label(
            fenetre, text=T('ref_popup.label_progression'),
            font=(POLICE, 11, 'bold'), bg=COULEURS['fond']
        ).pack(pady=(18, 6))

        var_msg = tk.StringVar(value=T('ref_popup.init'))
        tk.Label(
            fenetre, textvariable=var_msg,
            font=(POLICE, 9), bg=COULEURS['fond'], fg=COULEURS['texte']
        ).pack()

        def maj_message(msg):
            # Appelé depuis le thread de téléchargement → passage par after()
            fenetre.after(0, lambda m=msg: var_msg.set(m))

        def thread_download():
            succes = referentiel.telecharger_referentiel(callback_progression=maj_message)
            if succes:
                self.after(0, lambda: (
                    fenetre.destroy(),
                    messagebox.showinfo(T('ref_popup.titre_succes'),
                                        T('ref_popup.message_succes'))
                ))
            else:
                self.after(0, lambda: (
                    fenetre.destroy(),
                    messagebox.showerror(T('ref_popup.titre_erreur'),
                                         T('ref_popup.message_erreur'))
                ))

        t = threading.Thread(target=thread_download, daemon=True)
        t.start()

    def _importer_fichier(self):
        """Ouvre une boîte de dialogue pour choisir un fichier, puis le charge."""
        chemin = filedialog.askopenfilename(
            title='Choisir un export ObServe',
            filetypes=[
                ('Fichiers ObServe', ('*.zip', '*.xlsx', '*.xls')),
                ('ZIP', '*.zip'),
                ('Excel', '*.xlsx *.xls'),
                ('Tous les fichiers', '*.*'),
            ]
        )

        if not chemin:
            return

        try:
            self.tables = lire_fichier(chemin)
        except ValueError as e:
            messagebox.showerror(T('dialogs.erreur_titre'), str(e))
            return
        except Exception as e:
            messagebox.showerror(T('dialogs.erreur_titre'),
                                 f"Erreur inattendue lors de la lecture :\n{e}")
            return

        if not self.tables:
            messagebox.showerror(T('dialogs.erreur_titre'), T('dialogs.erreur_lecture'))
            return

        self._chemin_fichier = chemin  # sauvegardé pour localiser le fichier annotations
        nom_fichier = os.path.basename(chemin)
        self.nom_maree.set(f"Fichier : {nom_fichier}")
        self._noms_tables = sorted(self.tables.keys())
        self._reinitialiser_anomalies()
        self.annotations = self._charger_annotations()

        self.menu_table['values'] = self._noms_tables
        if self._noms_tables:
            self.var_table.set(self._noms_tables[0])
            self._afficher_table()

        self.statut.set(T('statut.fichier_charge', n=len(self.tables)))
        # Activer "Vérifier tout" — le bouton rapport reste désactivé jusqu'à la vérification
        self._btn_verifier_tout.config(state='normal')

    def _recevoir_tables(self, tables, nom_fichier, chemin_fichier=None):
        """
        Appelé par la page Accueil (app.py) quand un fichier est importé.
        Équivalent de _importer_fichier mais sans boîte de dialogue.
        chemin_fichier : chemin complet du zip, passé par app.py pour localiser le .json annotations
        """
        self.tables = tables
        self._chemin_fichier = chemin_fichier  # peut être None si app.py ne le transmet pas encore
        self.nom_maree.set(f"Fichier : {nom_fichier}")
        self._noms_tables = sorted(self.tables.keys())
        self._reinitialiser_anomalies()
        self.annotations = self._charger_annotations()
        self.menu_table['values'] = self._noms_tables
        if self._noms_tables:
            self.var_table.set(self._noms_tables[0])
            self._afficher_table()
        self.statut.set(T('statut.fichier_charge', n=len(self.tables)))
        # Activer "Vérifier tout" — le bouton rapport reste désactivé jusqu'à la vérification
        self._btn_verifier_tout.config(state='normal')

    # ── Gestion des annotations (Validé / Justifié + commentaire) ───────────────

    def _trip_id_courant(self):
        """Retourne le trip_id de la marée chargée, ou None si indisponible."""
        try:
            return str(self.tables['1_trip']['trip_id'].iloc[0]).strip()
        except Exception:
            return None

    def _chemin_annotations(self):
        """
        Retourne le chemin complet du fichier .json d'annotations pour la marée courante.
        Le fichier est stocké dans un dossier 'annotations/' à la racine de l'app
        (Option B : emplacement fixe, indépendant du dossier du zip importé).
        Retourne None si on ne peut pas déterminer le trip_id.
        """
        trip_id = self._trip_id_courant()
        if not trip_id:
            return None

        # Sanitize : on garde uniquement les caractères alphanumériques, tirets, underscores
        # pour éviter tout problème de nom de fichier sous Windows
        trip_id_safe = re.sub(r'[^\w\-]', '_', trip_id)

        # Dossier annotations/ à côté de ce fichier main.py (racine de l'app Phase 1)
        dossier = os.path.join(os.path.dirname(__file__), 'annotations')
        try:
            os.makedirs(dossier, exist_ok=True)
        except Exception:
            # Impossible de créer le dossier (permissions, disque réseau, etc.)
            # On retourne None : _sauvegarder_annotations() sautera la sauvegarde
            return None

        return os.path.join(dossier, f"{trip_id_safe}_annotations.json")

    def _charger_annotations(self):
        """
        Charge le fichier .json d'annotations pour la marée courante.
        Retourne un dict vide si le fichier est absent, illisible ou corrompu.
        En cas de corruption, l'app continue normalement sans annotations.
        """
        chemin = self._chemin_annotations()
        if not chemin or not os.path.exists(chemin):
            return {}
        try:
            with open(chemin, 'r', encoding='utf-8') as f:
                donnees = json.load(f)
            # Vérification minimale : on attend un dict
            if not isinstance(donnees, dict):
                return {}
            return donnees
        except Exception:
            # Fichier corrompu ou illisible → on repart de zéro sans planter
            return {}

    def _sauvegarder_annotations(self):
        """
        Sauvegarde self.annotations dans le fichier .json de la marée courante.
        Utilise une écriture atomique : on écrit d'abord dans un .tmp, puis on renomme.
        Ainsi, une coupure de courant pendant l'écriture ne corrompt pas le fichier existant.
        """
        chemin = self._chemin_annotations()
        if not chemin:
            return  # pas de trip_id → on ne peut pas sauvegarder

        chemin_tmp = chemin + '.tmp'
        try:
            with open(chemin_tmp, 'w', encoding='utf-8') as f:
                json.dump(self.annotations, f, ensure_ascii=False, indent=2)
            # Renommage atomique : remplace l'ancien fichier d'un seul coup
            os.replace(chemin_tmp, chemin)
            # Marquer que les annotations ont changé et ne sont plus exportées
            self._rapport_exporte = False
            self._mettre_a_jour_indicateur_export()
        except Exception:
            # Échec silencieux — on ne plante pas l'app pour un problème de sauvegarde.
            # On nettoie le .tmp pour ne pas laisser un fichier orphelin sur le disque.
            try:
                os.remove(chemin_tmp)
            except OSError:
                pass

    @staticmethod
    def _cle_annotation(anomalie):
        """
        Construit une clé unique pour identifier une anomalie dans le dict d'annotations.
        Format : table__set_norm__catch_norm__colonne__msghash
        - _norm_key() normalise set/catch_number (évite les problèmes int/float)
        - msghash (6 car.) distingue deux anomalies sur la même cellule mais de règles différentes
        """
        table     = anomalie.get('table', '') or ''
        set_n     = _norm_key(anomalie.get('set_number'))   if anomalie.get('set_number')   is not None else '_'
        catch_n   = _norm_key(anomalie.get('catch_number')) if anomalie.get('catch_number') is not None else '_'
        colonne   = anomalie.get('colonne', '') or '_'
        msg       = anomalie.get('message', '') or ''
        msghash   = hashlib.md5(msg.encode('utf-8')).hexdigest()[:6]
        return f"{table}__{set_n}__{catch_n}__{colonne}__{msghash}"

    def _mettre_a_jour_panneau_annotation(self, anomalie):
        """
        Met à jour le panneau annotations selon l'anomalie sélectionnée :
        active les boutons et affiche l'état actuel (validé / commentaire / rien).
        """
        if anomalie is None:
            # Aucune anomalie sélectionnée — griser tous les boutons
            self.btn_valide.config(state='disabled')
            self.btn_comment.config(state='disabled')
            self.btn_effacer.pack_forget()
            self.frame_comment.pack_forget()
            return

        cle = self._cle_annotation(anomalie)
        annot = self.annotations.get(cle, {})
        statut = annot.get('statut', '')

        # Activer les boutons
        self.btn_valide.config(state='normal')
        self.btn_comment.config(state='normal')

        if statut == 'correct':
            self.frame_comment.pack_forget()
            self.btn_effacer.config(state='normal')
            self.btn_effacer.pack(side='right', padx=(4, 0))

        elif statut == 'justifie':
            self.frame_comment.pack(fill='x', padx=4, pady=(0, 4))
            # Remplir le champ avec le commentaire existant
            self.txt_commentaire.delete('1.0', 'end')
            self.txt_commentaire.insert('1.0', annot.get('commentaire', ''))
            self.btn_effacer.config(state='normal')
            self.btn_effacer.pack(side='right', padx=(4, 0))

        else:
            # Aucune annotation — pas de zone texte, pas de bouton effacer
            self.frame_comment.pack_forget()
            self.btn_effacer.pack_forget()

    def _annoter_valide(self):
        """Marque l'anomalie sélectionnée comme Validée et sauvegarde."""
        if self._anomalie_selectionnee is None:
            return
        cle = self._cle_annotation(self._anomalie_selectionnee)
        # Cliquer sur un bouton déjà actif ne fait rien
        if self.annotations.get(cle, {}).get('statut') == 'correct':
            return
        self.annotations[cle] = {'statut': 'correct', 'commentaire': ''}
        self._sauvegarder_annotations()
        self._mettre_a_jour_panneau_annotation(self._anomalie_selectionnee)
        self._mettre_a_jour_compteur_annotations()
        self._mettre_a_jour_menu_tables()
        self._rafraichir_liste_apres_annotation()

    def _annoter_commentaire(self):
        """Active le mode commentaire pour l'anomalie sélectionnée."""
        if self._anomalie_selectionnee is None:
            return
        cle = self._cle_annotation(self._anomalie_selectionnee)
        # Cliquer sur un bouton déjà actif ne fait rien
        if self.annotations.get(cle, {}).get('statut') == 'justifie':
            return
        # Initialiser avec un commentaire vide si pas encore de justification
        commentaire_existant = self.annotations.get(cle, {}).get('commentaire', '')
        self.annotations[cle] = {'statut': 'justifie', 'commentaire': commentaire_existant}
        self._sauvegarder_annotations()
        self._mettre_a_jour_panneau_annotation(self._anomalie_selectionnee)
        self._mettre_a_jour_compteur_annotations()
        self._mettre_a_jour_menu_tables()
        self._rafraichir_liste_apres_annotation()
        # Focus direct dans le champ texte pour que l'utilisateur puisse taper immédiatement
        self.txt_commentaire.focus_set()

    def _sauvegarder_commentaire(self, event=None):
        """Sauvegarde le texte du commentaire sur perte de focus du champ texte."""
        if self._anomalie_selectionnee is None:
            return
        cle = self._cle_annotation(self._anomalie_selectionnee)
        if self.annotations.get(cle, {}).get('statut') != 'justifie':
            return
        texte = self.txt_commentaire.get('1.0', 'end-1c').strip()
        self.annotations[cle]['commentaire'] = texte
        self._sauvegarder_annotations()

    def _effacer_annotation(self):
        """Supprime l'annotation de l'anomalie sélectionnée après confirmation."""
        if self._anomalie_selectionnee is None:
            return
        if not messagebox.askyesno(
            T('dialogs.confirmer_effacer_titre'),
            T('dialogs.confirmer_effacer_msg'),
        ):
            return
        cle = self._cle_annotation(self._anomalie_selectionnee)
        if cle in self.annotations:
            del self.annotations[cle]
            self._sauvegarder_annotations()
        self._mettre_a_jour_panneau_annotation(self._anomalie_selectionnee)
        self._mettre_a_jour_compteur_annotations()
        self._mettre_a_jour_menu_tables()
        self._rafraichir_liste_apres_annotation()

    def _mettre_a_jour_indicateur_export(self):
        """Affiche un avertissement si des annotations existent mais le rapport n'a pas été exporté."""
        if self.annotations and not self._rapport_exporte:
            self.lbl_non_exporte.config(text='⚠ Annotations non exportées')
        else:
            self.lbl_non_exporte.config(text='')

    def _mettre_a_jour_compteur_annotations(self):
        """
        Met à jour les 3 compteurs d'annotation dans la barre de statut.
        Comptage sur self.anomalies (table courante uniquement).
        """
        if not self.anomalies:
            self.lbl_correct.config(text='')
            self.lbl_justifie.config(text='')
            self.lbl_a_traiter.config(text='')
            return
        total     = len(self.anomalies)
        n_valide  = sum(1 for a in self.anomalies
                        if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'correct')
        n_comment = sum(1 for a in self.anomalies
                        if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'justifie')
        n_traiter = total - n_valide - n_comment
        self.lbl_correct.config(text=f'✓{n_valide} validé(s)' if n_valide else '')
        self.lbl_justifie.config(text=f'💬{n_comment} justifié(s)' if n_comment else '')
        self.lbl_a_traiter.config(text=f'{n_traiter} à traiter' if n_traiter else '')

    def _changer_table(self, event=None):
        """
        Appelé quand l'utilisateur sélectionne une table dans le menu déroulant.
        Si "Vérifier tout" a déjà été lancé, on filtre les anomalies de cette table
        depuis anomalies_global. Sinon on réinitialise.
        """
        # Supprimer le surlignage du texte dans le champ du combobox après sélection
        self.menu_table.selection_clear()
        # Le libellé dans le combobox peut être "2_set ❌2 ⚠3" — on prend seulement le nom brut
        nom = self.var_table.get().split()[0]

        if self.anomalies_global:
            # Filtrer les anomalies déjà calculées pour cette table
            self.anomalies = [a for a in self.anomalies_global if a['table'] == nom]
            self._cellule_marquee = None
            # Sauvegarder le commentaire en cours avant de changer de table
            self._sauvegarder_commentaire()
            self._anomalie_selectionnee = None
            self._mettre_a_jour_panneau_annotation(None)  # effacer le panneau de l'ancienne table
            self._afficher_anomalies()
            self._afficher_table()
            n_err   = sum(1 for a in self.anomalies if a['niveau'] == 'Erreur')
            n_avert = sum(1 for a in self.anomalies if a['niveau'] == 'Avertissement')
            self.lbl_erreurs.config(text=T('anomalies.n_erreurs', n=n_err))
            self.lbl_avert.config(text=T('anomalies.n_avert', n=n_avert))
            self._mettre_a_jour_compteur_annotations()
            self.statut.set(T('statut.verif_terminee', table=nom, n_err=n_err, n_avert=n_avert))
        else:
            self._reinitialiser_anomalies()
            self._afficher_table()
            if nom in VERIFICATEURS:
                self.statut.set(T('statut.cliquer_verifier', table=nom))
            else:
                self.statut.set(T('statut.aucune_regle', table=nom))

    def _reinitialiser_anomalies(self):
        """Efface toutes les anomalies (globales et filtrées) et remet les compteurs à zéro."""
        self.anomalies              = []
        self.anomalies_global       = []
        self._cellule_marquee       = None
        self._anomalie_selectionnee = None
        self._rapport_exporte       = True   # réinitialiser l'indicateur export à chaque nouvelle marée
        self._btn_rapport.config(state='disabled')  # désactiver jusqu'à la prochaine vérification
        self.liste_anomalies.delete(0, 'end')
        self.detail_anomalie.config(text=T('anomalies.cliquer_detail'))
        self.lbl_erreurs.config(text='')
        self.lbl_avert.config(text='')
        self.lbl_correct.config(text='')
        self.lbl_justifie.config(text='')
        self.lbl_a_traiter.config(text='')
        self._mettre_a_jour_panneau_annotation(None)
        self._mettre_a_jour_indicateur_export()   # efface visuellement le label "non exporté"
        # Remettre les libellés du menu sans compteurs
        if self._noms_tables:
            self.menu_table['values'] = self._noms_tables

    def _lancer_verification_tout(self):
        """
        Lance les vérifications sur TOUTES les tables connues en une seule passe.
        Remplit self.anomalies_global avec l'ensemble des anomalies, puis filtre
        self.anomalies pour n'afficher que la table actuellement visible.
        Met à jour le combobox pour montrer les compteurs par table.
        """
        if not self.tables:
            messagebox.showwarning("", T('dialogs.aucun_fichier'))
            return

        ordre = {'Erreur': 0, 'Avertissement': 1, 'Info': 2}
        toutes = []

        # Lancer chaque vérificateur disponible pour les tables chargées
        for nom_table, func in VERIFICATEURS.items():
            if nom_table in self.tables:
                anomalies_table = func(self.tables)
                toutes.extend(anomalies_table)

        # Tri global : Erreur → Avertissement → ligne
        toutes.sort(key=lambda a: (
            ordre.get(a['niveau'], 9),
            a['ligne'] if a['ligne'] is not None else -1,
        ))

        self.anomalies_global = toutes

        # Mettre à jour le combobox avec les compteurs par table
        self._mettre_a_jour_menu_tables()

        # Afficher les anomalies de la table courante
        nom_courant = self.var_table.get().split()[0]
        self.anomalies = [a for a in self.anomalies_global if a['table'] == nom_courant]
        self._cellule_marquee = None
        self._afficher_anomalies()
        self._afficher_table()

        n_err_total   = sum(1 for a in self.anomalies_global if a['niveau'] == 'Erreur')
        n_avert_total = sum(1 for a in self.anomalies_global if a['niveau'] == 'Avertissement')
        n_err_table   = sum(1 for a in self.anomalies if a['niveau'] == 'Erreur')
        n_avert_table = sum(1 for a in self.anomalies if a['niveau'] == 'Avertissement')

        self.lbl_erreurs.config(text=T('anomalies.n_erreurs', n=n_err_table))
        self.lbl_avert.config(text=T('anomalies.n_avert', n=n_avert_table))
        self.statut.set(T('statut.verif_tout_terminee', n_err=n_err_total, n_avert=n_avert_total))

        # Nettoyer les annotations orphelines : si une anomalie n'existe plus après re-vérification
        # (l'observateur a corrigé le problème dans ObServe), son annotation doit être supprimée
        cles_valides = {self._cle_annotation(a) for a in self.anomalies_global}
        cles_orphelines = [c for c in self.annotations if c not in cles_valides]
        if cles_orphelines:
            for c in cles_orphelines:
                del self.annotations[c]
            self._sauvegarder_annotations()

        self._mettre_a_jour_compteur_annotations()
        # Activer le bouton rapport uniquement après une vérification (sinon le rapport serait vide)
        self._btn_rapport.config(state='normal')
        # Reprise automatique : si des annotations existent déjà (réimport), positionner
        # sur la 1ère anomalie non encore traitée pour reprendre sans chercher manuellement
        self._reprendre_position()

    def _reprendre_position(self):
        """
        Si des annotations existent déjà, sélectionne et affiche la 1ère anomalie
        de la liste courante qui n'a pas encore été annotée.
        Utile après un réimport : l'observateur reprend là où il en était.
        Ne fait rien si aucune annotation ou si toutes sont annotées.
        """
        if not self.annotations:
            return
        for i, anomalie in enumerate(self._anomalies_filtrees):
            if self._cle_annotation(anomalie) not in self.annotations:
                self.liste_anomalies.selection_set(i)
                self.liste_anomalies.see(i)
                self._anomalie_selectionnee = anomalie
                self._mettre_a_jour_panneau_annotation(anomalie)
                return

    def _mettre_a_jour_menu_tables(self):
        """
        Met à jour les valeurs du combobox pour afficher les compteurs d'anomalies
        et le nombre d'anomalies traitées (annotées) par table.
        Format : "2_set ❌2 ⚠3 · 1/5" ou "1_trip ✓"
        Conserve la sélection courante.
        """
        if not self._noms_tables:
            return

        nom_courant = self.var_table.get().split()[0]
        nouveaux_libelles = []

        for nom in self._noms_tables:
            anoms = [a for a in self.anomalies_global if a['table'] == nom]
            n_err   = sum(1 for a in anoms if a['niveau'] == 'Erreur')
            n_avert = sum(1 for a in anoms if a['niveau'] == 'Avertissement')
            n_total = len(anoms)

            # Nombre d'anomalies annotées par type
            n_correct  = sum(1 for a in anoms
                             if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'correct')
            n_justifie = sum(1 for a in anoms
                             if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'justifie')

            if n_err or n_avert:
                label = nom
                if n_err:
                    label += f' ❌{n_err}'
                if n_avert:
                    label += f' ⚠{n_avert}'
                # Afficher le détail correct/justifié si au moins une annotation
                if n_correct or n_justifie:
                    label += f' · ✓{n_correct} 💬{n_justifie}'
            else:
                # Table vérifiée sans anomalie
                label = f'{nom} ✓' if nom in VERIFICATEURS else nom

            nouveaux_libelles.append(label)

        self.menu_table['values'] = nouveaux_libelles

        # Restaurer la sélection sur la table courante (avec son nouveau libellé)
        for libelle in nouveaux_libelles:
            if libelle.split()[0] == nom_courant:
                self.var_table.set(libelle)
                break

    def _marquer_cellule(self, df_idx, nom_colonne):
        """
        Surligne en rouge gras la cellule exacte concernée par l'anomalie.
        Retire le surlignage de la cellule précédente avant d'en appliquer un nouveau.
        """
        nom = self.var_table.get().split()[0]
        df  = self.tables.get(nom)
        if df is None:
            return
        colonnes = [c for c in df.columns if 'topiaid' not in c]

        # Retire le surlignage de la cellule précédente
        if self._cellule_marquee:
            prev_df_idx, prev_col = self._cellule_marquee
            prev_sheet_row = self._idx_map.get(prev_df_idx)
            if prev_sheet_row is not None and prev_col in colonnes:
                prev_col_idx = colonnes.index(prev_col)
                try:
                    self.tableau.dehighlight_cells(
                        row=prev_sheet_row, column=prev_col_idx, redraw=False
                    )
                    # Restaure la couleur de ligne d'origine (erreur ou avertissement)
                    niveau_orig = next(
                        (a['niveau'] for a in self.anomalies if a['ligne'] == prev_df_idx),
                        None
                    )
                    if niveau_orig:
                        bg_orig = '#FFE0E0' if niveau_orig == 'Erreur' else '#FFF3CD'
                        self.tableau.highlight_rows(
                            rows=[prev_sheet_row], bg=bg_orig, redraw=False
                        )
                except Exception:
                    pass
            self._cellule_marquee = None

        # Applique le surlignage rouge gras sur la nouvelle cellule
        sheet_row = self._idx_map.get(df_idx)
        if sheet_row is not None and nom_colonne in colonnes:
            col_idx = colonnes.index(nom_colonne)
            try:
                self.tableau.highlight_cells(
                    row=sheet_row, column=col_idx,
                    bg='#CC0000', fg='white',
                    redraw=True
                )
                self._cellule_marquee = (df_idx, nom_colonne)
            except Exception:
                pass

    def _afficher_table(self):
        """Affiche le contenu de la table sélectionnée dans le tableau tksheet."""
        self._cellule_marquee = None
        self._idx_map         = {}
        # Supprime tous les surlignages avant de reconstruire le tableau
        try:
            self.tableau.dehighlight_all()
        except Exception:
            pass

        # Le libellé combobox peut être "2_set ❌2 ⚠3" — on prend seulement le nom brut
        nom = self.var_table.get().split()[0]
        if not nom or nom not in self.tables:
            return

        df       = self.tables[nom]
        colonnes = [c for c in df.columns if 'topiaid' not in c]

        # Construction des données et de la correspondance df_index → sheet_row
        donnees = []
        for sheet_row, (df_idx, ligne) in enumerate(df.iterrows()):
            self._idx_map[df_idx] = sheet_row
            # _afficher_val traduit les booléens et les vides — display uniquement
            donnees.append([_afficher_val(c, ligne[c]) for c in colonnes])

        # En-têtes traduits dans la langue active, avec fallback au nom technique
        entetes = [_en_tete(col) for col in colonnes]

        # Chargement dans tksheet
        self.tableau.set_sheet_data(donnees, redraw=False)
        self.tableau.headers(entetes)

        # Largeurs des colonnes : on prend le maximum entre la largeur de l'en-tête
        # et la largeur du contenu le plus long dans les données, pour que tout soit
        # lisible sans avoir à agrandir les colonnes manuellement.
        for i, col in enumerate(colonnes):
            # Largeur nécessaire pour l'en-tête traduit (ligne la plus longue)
            lignes_entete  = _en_tete(col).split('\n')
            largeur_entete = max(len(l) for l in lignes_entete) * 9

            # Largeur nécessaire pour les données affichées (valeur la plus longue)
            # Limité aux 200 premières lignes pour éviter un freeze sur les grandes tables (9_sample)
            valeurs = [_afficher_val(col, v) for v in df[col].iloc[:200]]
            largeur_data = max((len(v) for v in valeurs), default=0) * 8

            largeur = max(55, largeur_entete, largeur_data) + 10  # +10 px de marge
            self.tableau.column_width(column=i, width=largeur, redraw=False)

        # Surlignage des lignes en anomalie (rose = erreur, jaune = avertissement)
        lignes_anomalie = {
            a['ligne']: a['niveau']
            for a in self.anomalies
            if a['table'] == nom and a['ligne'] is not None
        }
        for df_idx, niveau in lignes_anomalie.items():
            sheet_row = self._idx_map.get(df_idx)
            if sheet_row is not None:
                bg = '#FFE0E0' if niveau == 'Erreur' else '#FFF3CD'
                self.tableau.highlight_rows(rows=[sheet_row], bg=bg, redraw=False)

        self.tableau.redraw()

    def _afficher_anomalies(self):
        """Remplit la liste des anomalies selon le filtre sélectionné."""
        self.liste_anomalies.delete(0, 'end')

        # On récupère le niveau canonique correspondant à la valeur affichée
        idx_filtre = self._filtres_affiches.index(self.var_filtre.get()) \
                     if self.var_filtre.get() in self._filtres_affiches else 0
        filtre_canon = self._filtres_canon[idx_filtre]

        self._anomalies_filtrees = []

        icones  = {'Erreur': '❌', 'Avertissement': '⚠ ', 'Info': 'ℹ '}
        couleurs = {
            'Erreur':        COULEURS['erreur'],
            'Avertissement': COULEURS['avertissement'],
            'Info':          COULEURS['accent'],
        }

        for anomalie in self.anomalies:
            if filtre_canon != 'Tout' and anomalie['niveau'] != filtre_canon:
                continue
            self._anomalies_filtrees.append(anomalie)
            # Préfixe et couleur selon annotation (priorité sur le niveau de sévérité)
            statut_annot = self.annotations.get(self._cle_annotation(anomalie), {}).get('statut')
            if statut_annot == 'correct':
                prefixe = '✓ '
                fg = '#27AE60'   # vert
            elif statut_annot == 'justifie':
                prefixe = '💬 '
                fg = '#0097A7'   # vert-bleu
            else:
                prefixe = ''
                fg = couleurs.get(anomalie['niveau'], '#333')
            texte = f"{prefixe}{icones.get(anomalie['niveau'], '?')}  [{anomalie['table']}]  {anomalie['message']}"
            self.liste_anomalies.insert('end', texte)
            self.liste_anomalies.itemconfig('end', fg=fg)

        # Si l'anomalie sélectionnée a disparu du filtre, vider le panneau
        if self._anomalie_selectionnee is not None and \
                self._anomalie_selectionnee not in self._anomalies_filtrees:
            self._anomalie_selectionnee = None
            self._mettre_a_jour_panneau_annotation(None)

    def _rafraichir_liste_apres_annotation(self):
        """
        Recharge la liste des anomalies (pour mettre à jour les icônes/couleurs d'annotation)
        et restaure automatiquement la sélection sur l'anomalie courante.
        """
        self._afficher_anomalies()
        if self._anomalie_selectionnee in self._anomalies_filtrees:
            idx = self._anomalies_filtrees.index(self._anomalie_selectionnee)
            self.liste_anomalies.selection_set(idx)
            self.liste_anomalies.see(idx)

    def _tab_vers_suivante(self, event=None):
        """Tab depuis le champ commentaire : sauvegarder et passer à l'anomalie suivante."""
        self._sauvegarder_commentaire()
        self._naviguer_anomalie(+1)
        self.liste_anomalies.focus_set()  # retour focus liste pour enchaîner C/V sans souris
        return 'break'  # empêche l'insertion du caractère tab dans le texte

    def _raccourci_correct(self):
        """Touche C depuis la liste : marquer correct puis passer à l'anomalie suivante."""
        if not self._anomalie_selectionnee:
            return
        self._annoter_valide()
        self._naviguer_anomalie(+1)
        self.liste_anomalies.focus_set()

    def _raccourci_justifier(self):
        """Touche V depuis la liste : marquer justifié puis focus dans le champ commentaire."""
        if not self._anomalie_selectionnee:
            return
        self._annoter_commentaire()
        # _annoter_commentaire() fait déjà focus_set() sur txt_commentaire

    def _naviguer_anomalie(self, delta):
        """Déplace la sélection de delta (+1 ou -1) dans la liste des anomalies."""
        if not self._anomalies_filtrees:
            return
        selection = self.liste_anomalies.curselection()
        idx_actuel = selection[0] if selection else -1
        idx_nouveau = max(0, min(len(self._anomalies_filtrees) - 1, idx_actuel + delta))
        self.liste_anomalies.selection_clear(0, 'end')
        self.liste_anomalies.selection_set(idx_nouveau)
        self.liste_anomalies.see(idx_nouveau)
        self.liste_anomalies.event_generate('<<ListboxSelect>>')

    def _sur_selection_anomalie(self, event):
        """
        Quand l'utilisateur clique sur une anomalie, affiche le détail
        et navigue vers la ligne concernée dans le tableau (si applicable).
        """
        selection = self.liste_anomalies.curselection()
        if not selection:
            return

        idx_liste = selection[0]
        if idx_liste >= len(self._anomalies_filtrees):
            return

        anomalie = self._anomalies_filtrees[idx_liste]
        # Sauvegarder le commentaire en cours AVANT de changer _anomalie_selectionnee.
        # Sur Windows, pack_forget() sur frame_comment peut déclencher FocusOut sur
        # txt_commentaire — si _anomalie_selectionnee est déjà la nouvelle anomalie,
        # _sauvegarder_commentaire() chercherait la mauvaise clé et ignorerait le commentaire.
        self._sauvegarder_commentaire()
        self._anomalie_selectionnee = anomalie
        self._mettre_a_jour_panneau_annotation(anomalie)

        # Affichage du détail
        # Pour les anomalies globales (ligne = None), on indique "table entière"
        ligne_txt   = T('detail.table_entiere') if anomalie['ligne']   is None else str(anomalie['ligne'])
        colonne_txt = T('detail.table_entiere') if anomalie['colonne'] is None else str(anomalie['colonne'])
        self.detail_anomalie.config(
            text=(
                T('detail.table', table=anomalie['table']) + '\n' +
                T('detail.ligne', ligne=ligne_txt) + '\n' +
                T('detail.colonne', colonne=colonne_txt) + '\n\n' +
                anomalie['message']
            ),
            fg={
                'Erreur':        COULEURS['erreur'],
                'Avertissement': COULEURS['avertissement'],
            }.get(anomalie['niveau'], COULEURS['texte'])
        )

        # Navigation vers la cellule dans le tableau (seulement si ligne précise)
        if anomalie['ligne'] is not None and anomalie['table'] in self.tables:
            # Trouver le libellé décoré dans le combobox (ex: "2_set ❌2 ⚠3") si présent
            libelle_cible = anomalie['table']
            for val in self.menu_table['values']:
                if val.split()[0] == anomalie['table']:
                    libelle_cible = val
                    break
            self.var_table.set(libelle_cible)
            self._afficher_table()
            try:
                df       = self.tables[anomalie['table']]
                colonnes = [c for c in df.columns if 'topiaid' not in c]
                sheet_row = self._idx_map.get(anomalie['ligne'])
                col_idx   = colonnes.index(anomalie['colonne']) if anomalie['colonne'] in colonnes else 0

                if sheet_row is not None:
                    # Défilement et sélection de la cellule
                    self.tableau.see(row=sheet_row, column=col_idx)
                    self.tableau.select_cell(row=sheet_row, column=col_idx)
                    # Surlignage rouge gras de la cellule concernée
                    self._marquer_cellule(anomalie['ligne'], anomalie['colonne'])
            except Exception:
                pass

    def _changer_langue(self, event=None):
        """Sauvegarde la langue choisie et demande un redémarrage."""
        code = self.var_langue.get().lower()
        lang.sauvegarder_langue(code)
        nom_langue = 'Français' if code == 'fr' else 'English'
        messagebox.showinfo(
            T('langue.titre_popup'),
            T('langue.message_popup', langue=nom_langue)
        )

    def _ouvrir_referentiel(self):
        """Ouvre la fenêtre de visualisation du référentiel ObServe."""
        FenetreReferentiel(self)

    def set_zone_peche(self, zone):
        """
        Permet à app.py de transmettre la zone de pêche sélectionnée dans l'Accueil.
        Appelé après _recevoir_tables() depuis le callback d'import.
        """
        self._zone_peche = zone or ''

    def set_zone_peche_getter(self, getter):
        """
        Permet à app.py de passer un callable retournant la zone courante.
        Prioritaire sur _zone_peche : la zone est lue au moment de l'export,
        donc une sélection faite APRÈS l'import est bien prise en compte.
        """
        self._zone_peche_getter = getter

    def _exporter_rapport_annote(self):
        """
        Exporte le rapport annoté au format Excel professionnel (Point 5 Vision B).
        - Format validé visuellement par l'utilisateur avant codage.
        - Feuille protégée (cellules verrouillées, filtres actifs).
        - Logos IRD / OB7 / MARBEC / SCISEA, infos marée, tableau de validation,
          tableau des anomalies coloré par statut d'annotation.
        """
        if not self.anomalies_global:
            # 0 anomalie après vérification = fichier conforme → demander confirmation
            if not messagebox.askyesno(
                "Rapport sans anomalie",
                "Aucune anomalie n'a été détectée — le fichier est conforme.\n\n"
                "Le rapport mentionnera explicitement l'absence d'anomalie.\n\n"
                "Continuer ?"
            ):
                return

        from datetime import datetime
        now = datetime.now()

        # ── Extraction des infos marée depuis 1_trip ──────────────────────────
        def _val(col, defaut='—'):
            """Lit une colonne de 1_trip de façon sécurisée."""
            try:
                v = self.tables['1_trip'][col].iloc[0]
                return str(v).strip() if pd.notna(v) else defaut
            except Exception:
                return defaut

        def _date(col):
            """Formate une colonne date de 1_trip en jj/mm/aaaa, ou '—'."""
            try:
                v = self.tables['1_trip'][col].iloc[0]
                if pd.isna(v):
                    return '—'
                return pd.to_datetime(v).strftime('%d/%m/%Y')
            except Exception:
                return '—'

        def _date_iso(col):
            """Formate une colonne date en YYYYMMDD pour le nom de fichier, ou 'XXXXXXXX'."""
            try:
                v = self.tables['1_trip'][col].iloc[0]
                if pd.isna(v):
                    return 'XXXXXXXX'
                return pd.to_datetime(v).strftime('%Y%m%d')
            except Exception:
                return 'XXXXXXXX'

        trip_id     = _val('trip_id')
        navire      = _val('vessel')
        capitaine   = (f"{_val('captain_firstname', '')} {_val('captain_lastname', '')}").strip() or '—'
        observateur = (f"{_val('observer_firstname', '')} {_val('observer_lastname', '')}").strip() or '—'
        dep_date    = _date('trip_start_date')
        arr_date    = _date('trip_end_date')
        dep_port    = _val('departure_harbour', _val('departure_port', '—'))
        arr_port    = _val('arrival_harbour',   _val('arrival_port',   '—'))
        n_sets      = len(self.tables.get('2_set',    pd.DataFrame()))
        n_catches   = len(self.tables.get('8_catch',  pd.DataFrame()))
        n_samples   = len(self.tables.get('9_sample', pd.DataFrame()))
        # Priorité au getter (valeur en temps réel depuis l'Accueil) sur la valeur statique
        if self._zone_peche_getter:
            raw_zee = self._zone_peche_getter()
        else:
            raw_zee = self._zone_peche
        zee = raw_zee if raw_zee and raw_zee != '[Zone de pêche non renseignée]' else '—'

        # ── Nom de fichier : même logique que le dossier de marée ────────────
        dep_iso     = _date_iso('trip_start_date')
        arr_iso     = _date_iso('trip_end_date')
        navire_safe = ''.join(c for c in navire.upper() if c.isalnum())
        obs_nom     = _val('observer_lastname', 'OBS').upper()
        date_rapport = now.strftime('%Y%m%d')
        nom_defaut  = (f"{trip_id}_{dep_iso}-{arr_iso}_{navire_safe}_{obs_nom}"
                       f"_RapportVérification_{date_rapport}.xlsx")

        # ── Dossier initial : _Rapport/ à côté du zip, sinon dossier du zip ─────
        initial_dir = None
        if self._chemin_fichier:
            dossier_zip = os.path.dirname(self._chemin_fichier)
            dossier_rapport = os.path.join(dossier_zip, '_Rapport')
            if os.path.isdir(dossier_rapport):
                initial_dir = dossier_rapport
            else:
                initial_dir = dossier_zip

        # ── Boîte de dialogue pour le fichier de sortie ───────────────────────
        chemin_sortie = filedialog.asksaveasfilename(
            title="Enregistrer le rapport annoté",
            defaultextension=".xlsx",
            filetypes=[('Excel', '*.xlsx')],
            initialfile=nom_defaut,
            initialdir=initial_dir,
        )
        if not chemin_sortie:
            return

        # ── Génération et sauvegarde ──────────────────────────────────────────
        try:
            self._generer_xlsx_annote(
                chemin_sortie, trip_id, navire, capitaine, observateur, zee,
                dep_date, dep_port, arr_date, arr_port,
                n_sets, n_catches, n_samples, now
            )
            messagebox.showinfo(
                "Export terminé",
                f"Rapport annoté enregistré :\n{chemin_sortie}"
            )
            self.statut.set(f"Rapport annoté exporté : {os.path.basename(chemin_sortie)}")
            self._rapport_exporte = True
            self._mettre_a_jour_indicateur_export()
        except PermissionError:
            messagebox.showerror(
                "Fichier ouvert",
                "Impossible d'enregistrer : le fichier est déjà ouvert dans Excel.\n"
                "Fermez-le puis réessayez."
            )
        except Exception as e:
            messagebox.showerror("Erreur export", f"Impossible de générer le rapport :\n{e}")

    def _generer_xlsx_annote(self, chemin, trip_id, navire, capitaine, observateur,
                              zee, dep_date, dep_port, arr_date, arr_port,
                              n_sets, n_catches, n_samples, now):
        """
        Génère le fichier Excel annoté (Vision B Point 5).
        Séparé de _exporter_rapport_annote() pour être testable indépendamment.
        Tout le code openpyxl est ici — pas dans la méthode UI.
        """
        # Imports openpyxl — en local pour ne pas alourdir le démarrage de l'app
        from openpyxl import Workbook
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
        from openpyxl.styles.protection import Protection
        from openpyxl.utils import get_column_letter

        # ── Chemins des logos ─────────────────────────────────────────────────
        # Les logos sont dans assets/logos/ à la racine du projet,
        # deux niveaux au-dessus de phase1_detection/main.py.
        RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        LOGOS_DIR = os.path.join(RACINE, 'assets', 'logos')
        LOGO_IRD    = os.path.join(LOGOS_DIR, 'logo_IRD_nobg.png')
        LOGO_OB7    = os.path.join(LOGOS_DIR, 'logoOb7_nobg.png')
        LOGO_MARBEC = os.path.join(LOGOS_DIR, 'logo_marbec_nobg.png')
        LOGO_SCISEA = os.path.join(LOGOS_DIR, 'logo_SCISEA_nobg.png')

        # ── Helpers de style ──────────────────────────────────────────────────
        def remplir(hex_color):
            return PatternFill(fill_type='solid', fgColor=hex_color)

        def police(bold=False, color='000000', size=10, italic=False):
            return Font(bold=bold, color=color, size=size,
                        italic=italic, name='Calibri')

        def aligner(h='left', v='center', wrap=False):
            return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

        def bordure_fine():
            s = Side(style='thin', color='CCCCCC')
            return Border(left=s, right=s, top=s, bottom=s)

        def bordure_droite_fine():
            """Bord droit seul — pour clore les cellules fusionnées."""
            s = Side(style='thin', color='CCCCCC')
            return Border(right=s, top=s, bottom=s)

        def ancrer_logo(ws, filepath, col_0, row_0, img_w, img_h, col_w_px, row_h_px):
            """Ajoute un logo centré dans sa cellule via ancre EMU."""
            PX = 9525   # 1 pixel = 9525 EMU à 96 DPI (standard OpenXML)
            img = XLImage(filepath)
            img.width  = img_w
            img.height = img_h
            x_off = max(0, (col_w_px - img_w) // 2) * PX
            y_off = max(0, (row_h_px - img_h) // 2) * PX
            marker = AnchorMarker(col=col_0, colOff=x_off, row=row_0, rowOff=y_off)
            size   = XDRPositiveSize2D(cx=img_w * PX, cy=img_h * PX)
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(img)

        # ── Création du classeur ──────────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = 'Rapport de vérification'

        # Largeurs des colonnes (en caractères)
        col_widths = {'A': 20, 'B': 14, 'C': 14, 'D': 18, 'E': 50, 'F': 12, 'G': 30}
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # ── LIGNE 1 : LOGOS ───────────────────────────────────────────────────
        ws.row_dimensions[1].height = 85
        ws.row_dimensions[2].height = 10   # séparateur fin

        ws.merge_cells('D1:F1')
        ws['D1'].value     = 'RAPPORT DE VÉRIFICATION\nObserve-Tools — IRD OB7 / SCISEA'
        ws['D1'].font      = police(bold=True, size=14, color='1F4E79')
        ws['D1'].alignment = aligner(h='center', v='center', wrap=True)

        for col in 'ABCDEFG':
            ws[f'{col}1'].fill = remplir('D0E8F5')   # fond bleu IRD

        ROW1_H_PX = int(85 * 4 / 3)   # 85pt → ~113px
        COL_PX = {k: v * 7 for k, v in col_widths.items()}   # 1 char ≈ 7px

        H_G = 45    # hauteur harmonisée logos gauche (px)
        H_S = 80    # hauteur logo SCISEA (px)

        if os.path.exists(LOGO_IRD):
            ancrer_logo(ws, LOGO_IRD,    0, 0, int(H_G*2.11), H_G, COL_PX['A'], ROW1_H_PX)
        if os.path.exists(LOGO_OB7):
            ancrer_logo(ws, LOGO_OB7,    1, 0, int(H_G*1.39), H_G, COL_PX['B'], ROW1_H_PX)
        if os.path.exists(LOGO_MARBEC):
            ancrer_logo(ws, LOGO_MARBEC, 2, 0, int(H_G*1.47), H_G, COL_PX['C'], ROW1_H_PX)
        if os.path.exists(LOGO_SCISEA):
            ancrer_logo(ws, LOGO_SCISEA, 6, 0, int(H_S*0.78), H_S, COL_PX['G'], ROW1_H_PX)

        # ── LIGNES 3-7 : INFOS MARÉE ──────────────────────────────────────────
        FILL_INFO  = remplir('EBF4FB')
        FILL_LABEL = remplir('D0E8F5')
        F_LABEL    = police(bold=True,  color='1F4E79', size=10)
        F_VALUE    = police(bold=False, color='000000', size=10)

        date_ref = referentiel.date_referentiel_local() or '—'

        infos_gauche = [
            ('Marée',        trip_id),
            ('Navire',       navire),
            ('Capitaine',    capitaine),
            ('Observateur',  observateur),
            ('ZEE',          zee),
            ('Référentiel',  date_ref),
        ]
        infos_droite = [
            ('Départ',       f"{dep_date}  —  {dep_port}"),
            ('Arrivée',      f"{arr_date}  —  {arr_port}"),
            ('Filages',      str(n_sets)),
            ('Captures',     str(n_catches)),
            ('Mensurations', str(n_samples)),
            ('Version app',  f"Observe-Tools v{VERSION}"),
        ]

        s_fin = Side(style='thin', color='CCCCCC')
        for i, ((lbl_g, val_g), (lbl_d, val_d)) in enumerate(
                zip(infos_gauche, infos_droite)):
            row = 3 + i
            ws.row_dimensions[row].height = 18

            ws[f'A{row}'].value     = lbl_g
            ws[f'A{row}'].font      = F_LABEL
            ws[f'A{row}'].fill      = FILL_LABEL
            ws[f'A{row}'].alignment = aligner(h='right')
            ws[f'A{row}'].border    = bordure_fine()

            ws.merge_cells(f'B{row}:C{row}')
            ws[f'B{row}'].value     = val_g
            ws[f'B{row}'].font      = F_VALUE
            ws[f'B{row}'].fill      = FILL_INFO
            ws[f'B{row}'].alignment = aligner(h='left')
            ws[f'B{row}'].border    = bordure_fine()

            ws[f'D{row}'].fill = remplir('FFFFFF')   # colonne séparatrice

            ws[f'E{row}'].value     = lbl_d
            ws[f'E{row}'].font      = F_LABEL
            ws[f'E{row}'].fill      = FILL_LABEL
            ws[f'E{row}'].alignment = aligner(h='right')
            ws[f'E{row}'].border    = bordure_fine()

            ws.merge_cells(f'F{row}:G{row}')
            ws[f'F{row}'].value     = val_d
            ws[f'F{row}'].font      = F_VALUE
            ws[f'F{row}'].fill      = FILL_INFO
            ws[f'F{row}'].alignment = aligner(h='left')
            ws[f'F{row}'].border    = bordure_fine()

            # Bords droits des cellules fusionnées (openpyxl ne les propage pas)
            ws[f'C{row}'].border = Border(right=s_fin, top=s_fin, bottom=s_fin)
            ws[f'G{row}'].border = Border(right=s_fin, top=s_fin, bottom=s_fin)

        # ── LIGNE 9 : BILAN ───────────────────────────────────────────────────
        # NB : la boucle infos ci-dessus remplit les lignes 3 à 8 (6 items).
        # Le bilan doit donc être en ligne 9 pour éviter le chevauchement de fusion.
        n_valide  = sum(1 for a in self.anomalies_global
                        if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'correct')
        n_comment = sum(1 for a in self.anomalies_global
                        if self.annotations.get(self._cle_annotation(a), {}).get('statut') == 'justifie')
        n_traiter = len(self.anomalies_global) - n_valide - n_comment
        total     = len(self.anomalies_global)
        now_str   = now.strftime('%d/%m/%Y à %H:%M')

        ws.row_dimensions[9].height = 16
        ws.merge_cells('A9:G9')
        ws['A9'].value     = (f"Généré le {now_str}  —  {total} anomalie(s)  •  "
                              f"{n_valide} validé(s)  •  {n_comment} justifié(s)  •  "
                              f"{n_traiter} à traiter")
        ws['A9'].font      = police(bold=False, color='555555', size=9, italic=True)
        ws['A9'].fill      = remplir('E8E8E8')
        ws['A9'].alignment = aligner(h='center')

        # ── LIGNES 10-14 : TABLEAU DE VALIDATION ─────────────────────────────
        # Ligne 10 : séparateur fin
        # Ligne 11 : en-tête validation
        # Lignes 12-13 : SCISEA et IRD
        # Ligne 14 : séparateur fin
        ws.row_dimensions[10].height = 6
        ws.row_dimensions[14].height = 6

        FILL_VH = remplir('4BACC6')   # bleu moyen en-tête validation
        FILL_VL = remplir('D6EEF8')   # bleu clair label validation
        FILL_VC = remplir('FFFFFF')   # blanc cellule à remplir
        F_VH    = police(bold=True,  color='FFFFFF', size=10)
        F_VL    = police(bold=True,  color='1F4E79', size=10)
        UNLOCK  = Protection(locked=False)

        # En-tête validation (ligne 11)
        ws.row_dimensions[11].height = 18
        ws.merge_cells('B11:C11')
        ws.merge_cells('E11:G11')
        for col, texte in (('A', 'Étape'), ('B', 'Validé par'),
                           ('D', 'Date'), ('E', 'Commentaire')):
            c = ws[f'{col}11']
            c.value     = texte
            c.font      = F_VH
            c.fill      = FILL_VH
            c.alignment = aligner(h='center')
            c.border    = bordure_fine()
        for col in ('C', 'F', 'G'):
            ws[f'{col}11'].fill   = FILL_VH
            ws[f'{col}11'].border = bordure_fine()

        # Lignes SCISEA (12) et IRD (13)
        for label, row in (('Responsable SCISEA', 12), ('Responsable IRD', 13)):
            ws.row_dimensions[row].height = 22
            ws[f'A{row}'].value     = label
            ws[f'A{row}'].font      = F_VL
            ws[f'A{row}'].fill      = FILL_VL
            ws[f'A{row}'].alignment = aligner(h='left')
            ws[f'A{row}'].border    = bordure_fine()

            ws.merge_cells(f'B{row}:C{row}')
            ws[f'B{row}'].fill       = FILL_VC
            ws[f'B{row}'].alignment  = aligner(h='left')
            ws[f'B{row}'].border     = bordure_fine()
            ws[f'B{row}'].protection = UNLOCK
            ws[f'C{row}'].border     = bordure_droite_fine()

            ws[f'D{row}'].fill       = FILL_VC
            ws[f'D{row}'].alignment  = aligner(h='center')
            ws[f'D{row}'].border     = bordure_fine()
            ws[f'D{row}'].protection = UNLOCK

            ws.merge_cells(f'E{row}:G{row}')
            ws[f'E{row}'].fill       = FILL_VC
            ws[f'E{row}'].alignment  = aligner(h='left', wrap=True)
            ws[f'E{row}'].border     = bordure_fine()
            ws[f'E{row}'].protection = UNLOCK
            for col in ('F', 'G'):
                ws[f'{col}{row}'].border = bordure_droite_fine()

        # ── LIGNES 15+ : TABLEAU ANOMALIES (ou message "aucune anomalie") ────────
        # Tri d'abord (nécessaire pour le test vide ci-dessous)
        # Tri : Table (ordre d'import) → Niveau (Erreur avant Avertissement) → Ligne
        ordre_tables = {t: i for i, t in enumerate(self._noms_tables)}
        ordre_niveau = {'Erreur': 0, 'Avertissement': 1, 'Info': 2}
        anomalies_triees = sorted(
            self.anomalies_global,
            key=lambda a: (
                ordre_tables.get(a.get('table', ''), 99),
                ordre_niveau.get(a.get('niveau', ''), 9),
                a['ligne'] if a.get('ligne') is not None else -1,
            )
        )

        if not anomalies_triees:
            # Aucune anomalie → message vert bien visible à la place du tableau
            ws.row_dimensions[15].height = 40
            ws.merge_cells('A15:G15')
            cell_ok = ws['A15']
            cell_ok.value     = "✓  Aucune erreur ou anomalie n'a été détectée — le fichier est conforme."
            cell_ok.font      = police(bold=True, color='FFFFFF', size=13)
            cell_ok.fill      = remplir('27AE60')   # vert
            cell_ok.alignment = aligner(h='center')
            cell_ok.border    = bordure_fine()
            derniere_ligne    = 15
        else:
            # ── LIGNE 15 : EN-TÊTE TABLEAU ANOMALIES ─────────────────────────
            ws.row_dimensions[15].height = 20
            FILL_HEADER = remplir('1F4E79')
            for col_idx, entete in enumerate(
                    ['Niveau', 'Table', 'Ligne', 'Colonne', 'Message', 'Statut', 'Commentaire'],
                    start=1):
                col  = get_column_letter(col_idx)
                cell = ws[f'{col}15']
                cell.value     = entete
                cell.font      = police(bold=True, color='FFFFFF', size=10)
                cell.fill      = FILL_HEADER
                cell.alignment = aligner(h='center')
                cell.border    = bordure_fine()

        ICONES = {'Erreur': '❌', 'Avertissement': '⚠'}
        COULEURS_STATUT = {
            'correct'  : ('E8F8EE', '27AE60'),   # fond vert clair, texte vert
            'justifie' : ('E4F4F7', '0097A7'),   # fond teal clair, texte teal
        }
        COULEURS_NIVEAU = {
            'Erreur'       : 'C0392B',
            'Avertissement': 'E67E22',
        }
        LIBELLES_STATUT = {
            'correct'  : 'Validé',
            'justifie' : 'Justifié',
        }
        FILL_SEP = remplir('2E4057')   # bleu-gris sombre pour les séparateurs de table

        # Hauteur adaptée au contenu : estime les lignes visuelles selon largeur colonne
        def _lignes_visuelles(texte, largeur_col):
            """Compte les lignes visuelles d'un texte dans une cellule wrap_text."""
            if not texte:
                return 1
            import math
            n = 0
            for segment in str(texte).splitlines() or ['']:
                n += max(1, math.ceil(len(segment) / largeur_col)) if segment else 1
            return max(1, n)

        row = 15          # incrémenté avant chaque écriture
        table_courante = None

        for anomalie in anomalies_triees:
            # Insérer une ligne séparatrice à chaque changement de table
            if anomalie.get('table') != table_courante:
                row += 1
                table_courante = anomalie.get('table', '')
                ws.row_dimensions[row].height = 16
                ws[f'A{row}'].value     = f'▶  {table_courante}'
                ws[f'A{row}'].font      = police(bold=True, color='FFFFFF', size=9)
                ws[f'A{row}'].fill      = FILL_SEP
                ws[f'A{row}'].alignment = aligner(h='left')
                for col_idx in range(1, 8):
                    c = ws[f'{get_column_letter(col_idx)}{row}']
                    c.fill   = FILL_SEP
                    c.border = bordure_fine()

            row += 1

            statut_annot = self.annotations.get(
                self._cle_annotation(anomalie), {}
            ).get('statut', '')
            commentaire_texte = self.annotations.get(
                self._cle_annotation(anomalie), {}
            ).get('commentaire', '')
            fill_bg, _ = COULEURS_STATUT.get(statut_annot, ('FFFFFF', '333333'))
            libelle_statut = LIBELLES_STATUT.get(statut_annot, 'À traiter')

            msg_texte = anomalie.get('message', '') or ''
            valeurs = [
                f"{ICONES.get(anomalie['niveau'], '')} {anomalie['niveau']}",
                anomalie.get('table', ''),
                anomalie.get('ligne', '') if anomalie.get('ligne') is not None else '',
                anomalie.get('colonne', ''),
                msg_texte,
                libelle_statut,
                commentaire_texte,
            ]

            n_lignes = max(
                _lignes_visuelles(msg_texte,         col_widths['E']),
                _lignes_visuelles(commentaire_texte, col_widths['G']),
            )
            ws.row_dimensions[row].height = max(16, n_lignes * 14)

            for col_idx, valeur in enumerate(valeurs, start=1):
                col  = get_column_letter(col_idx)
                cell = ws[f'{col}{row}']
                cell.value     = valeur
                cell.fill      = remplir(fill_bg)
                cell.alignment = aligner(h='left', wrap=(col_idx in (5, 7)))
                cell.border    = bordure_fine()

                if col_idx == 1:   # Niveau : couleur selon sévérité
                    cell.font = police(bold=True,
                                       color=COULEURS_NIVEAU.get(anomalie['niveau'], '333333'))
                elif col_idx == 6:  # Statut : couleur selon annotation
                    _, fg = COULEURS_STATUT.get(statut_annot, ('FFFFFF', '333333'))
                    cell.font      = police(bold=True, color=fg)
                    cell.alignment = aligner(h='center')
                else:
                    cell.font = police()

        # ── FILTRES, VOLETS, PROTECTION ───────────────────────────────────────
        derniere_ligne = row if anomalies_triees else 15
        ws.auto_filter.ref = f'A15:G{derniere_ligne}'

        ws.protection.sheet      = True
        ws.protection.password   = 'DOL2026'
        ws.protection.autoFilter = False   # filtres utilisables malgré la protection
        ws.protection.sort       = False   # tri utilisable malgré la protection

        wb.save(chemin)


# ─────────────────────────────────────────────────────────────────────────────
# FENÊTRE DE VISUALISATION DU RÉFÉRENTIEL
# ─────────────────────────────────────────────────────────────────────────────

class FenetreReferentiel(tk.Toplevel):
    """
    Fenêtre de consultation du référentiel ObServe local.
    Arborescence à gauche (3 groupes + tables), contenu à droite.
    """

    COLS_MASQUEES = {'topiaId', 'topiaCreateDate', 'uri', 'homeId', 'lastUpdateDate'}

    # Ordre d'affichage et libellés des 3 groupes
    GROUPES = [
        ('common',         'fr.ird.observe.entities.referential.common'),
        ('ll.common',      'fr.ird.observe.entities.referential.ll.common'),
        ('ll.observation', 'fr.ird.observe.entities.referential.ll.observation'),
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Référentiel ObServe")
        self.geometry("1000x620")
        self.configure(bg=COULEURS['fond'])
        self.resizable(True, True)

        data = referentiel.charger_referentiel_local()
        if not data:
            tk.Label(
                self,
                text=T('referentiel.non_disponible'),
                font=(POLICE, 11), bg=COULEURS['fond'], justify='center'
            ).pack(expand=True)
            return

        self.referentiels  = data.get('referentiels', {})
        date_maj           = data.get('date_telechargement', 'inconnue')
        self._item_vers_cle = {}   # id nœud arbre → clé référentiel complète
        self._donnees_table  = []
        self._colonnes_table = []

        # ── En-tête ───────────────────────────────────────────────────────────
        en_tete = tk.Frame(self, bg=COULEURS['en_tete'], height=44)
        en_tete.pack(fill='x')
        en_tete.pack_propagate(False)
        tk.Label(
            en_tete, text=T('referentiel.titre_entete'),
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 13, 'bold')
        ).pack(side='left', padx=16, pady=8)
        tk.Label(
            en_tete, text=T('referentiel.date_maj', date=date_maj),
            bg=COULEURS['en_tete'], fg='#C7D2FE', font=(POLICE, 9)
        ).pack(side='right', padx=16)

        # ── Corps ─────────────────────────────────────────────────────────────
        corps = tk.Frame(self, bg=COULEURS['fond'])
        corps.pack(fill='both', expand=True, padx=8, pady=8)

        # Panneau gauche — arborescence (déclaré AVANT le panneau droit)
        gauche = tk.Frame(corps, width=240, bg=COULEURS['blanc'], relief='groove', bd=1)
        gauche.pack(side='left', fill='y', padx=(0, 8))
        gauche.pack_propagate(False)

        tk.Label(
            gauche, text=T('referentiel.groupes_tables'),
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 9, 'bold')
        ).pack(fill='x')

        # Champ de recherche — filtre les nœuds de l'arbre en temps réel
        self.var_filtre = tk.StringVar()
        self.var_filtre.trace('w', lambda *a: self._filtrer_arbre())
        tk.Entry(gauche, textvariable=self.var_filtre, font=(POLICE, 9)).pack(
            fill='x', padx=4, pady=(4, 2)
        )
        tk.Label(
            gauche, text=T('referentiel.filtrer_nom'),
            font=(POLICE, 7), bg=COULEURS['blanc'], fg='#888'
        ).pack(anchor='w', padx=4)

        # Arborescence ttk.Treeview
        style = ttk.Style()
        style.configure('Ref.Treeview',      font=(POLICE, 9),      rowheight=22)
        style.configure('Ref.Treeview.Item', font=(POLICE, 9))

        cadre_arbre = tk.Frame(gauche, bg=COULEURS['blanc'])
        cadre_arbre.pack(fill='both', expand=True, pady=(4, 0))

        self.arbre = ttk.Treeview(cadre_arbre, show='tree', selectmode='browse',
                                   style='Ref.Treeview')
        scroll_arbre = ttk.Scrollbar(cadre_arbre, command=self.arbre.yview)
        self.arbre.configure(yscrollcommand=scroll_arbre.set)
        scroll_arbre.pack(side='right', fill='y')
        self.arbre.pack(fill='both', expand=True)
        self.arbre.bind('<<TreeviewSelect>>', self._sur_clic_arbre)

        # Panneau droit — contenu de la table sélectionnée
        droite = tk.Frame(corps, bg=COULEURS['fond'])
        droite.pack(side='left', fill='both', expand=True)

        barre_droite = tk.Frame(droite, bg=COULEURS['fond'])
        barre_droite.pack(fill='x', pady=(0, 4))

        self.lbl_table = tk.Label(
            barre_droite, text=T('referentiel.selectionner'),
            font=(POLICE, 10, 'bold'), bg=COULEURS['fond'], anchor='w'
        )
        self.lbl_table.pack(side='left')

        tk.Label(barre_droite, text=T('referentiel.rechercher'),
                 font=(POLICE, 9), bg=COULEURS['fond']).pack(side='right', padx=(8, 2))
        self.var_recherche_contenu = tk.StringVar()
        self.var_recherche_contenu.trace('w', lambda *a: self._filtrer_contenu())
        tk.Entry(barre_droite, textvariable=self.var_recherche_contenu,
                 font=(POLICE, 9), width=20).pack(side='right')

        self.tableau = tksheet.Sheet(
            droite,
            show_x_scrollbar=True, show_y_scrollbar=True,
            header_height=28, row_height=22,
            font=(POLICE, 9, 'normal'), header_font=(POLICE, 9, 'bold'),
            show_row_index=False,
        )
        self.tableau.enable_bindings('single_select', 'column_width_resize',
                                      'arrowkeys', 'copy')
        self.tableau.pack(fill='both', expand=True)

        # Construction initiale de l'arborescence
        self._construire_arbre(self.referentiels)

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _nom_court(cle):
        return cle.split('.')[-1]

    # ── Arborescence ──────────────────────────────────────────────────────────

    def _construire_arbre(self, referentiels_a_afficher):
        """
        Construit (ou reconstruit) l'arborescence à partir d'un sous-ensemble
        du référentiel. Les 3 groupes sont les nœuds parents, les tables sont
        les feuilles.
        """
        # Supprime tous les nœuds existants
        for item in self.arbre.get_children():
            self.arbre.delete(item)
        self._item_vers_cle.clear()

        for nom_groupe, prefixe in self.GROUPES:
            # Tables de ce groupe présentes dans referentiels_a_afficher
            tables_groupe = sorted(
                (cle for cle in referentiels_a_afficher if cle.startswith(prefixe)),
                key=self._nom_court
            )
            if not tables_groupe:
                continue  # groupe vide : on ne l'affiche pas

            # Nœud parent (groupe)
            id_groupe = self.arbre.insert(
                '', 'end',
                text=f'  {nom_groupe}  ({len(tables_groupe)})',
                open=True,
                tags=('groupe',)
            )

            # Feuilles (tables)
            for cle in tables_groupe:
                nb = len(referentiels_a_afficher.get(cle, []))
                id_table = self.arbre.insert(
                    id_groupe, 'end',
                    text=f'    {self._nom_court(cle)}  ({nb})',
                    tags=('table',)
                )
                self._item_vers_cle[id_table] = cle

        # Style visuel : groupes en gras, tables en normal
        self.arbre.tag_configure('groupe', font=(POLICE, 9, 'bold'))
        self.arbre.tag_configure('table',  font=(POLICE, 9))

    def _filtrer_arbre(self):
        """Reconstruit l'arbre en ne gardant que les tables correspondant au filtre."""
        terme = self.var_filtre.get().lower()
        if terme:
            filtres = {
                cle: val for cle, val in self.referentiels.items()
                if terme in self._nom_court(cle).lower()
            }
        else:
            filtres = self.referentiels
        self._construire_arbre(filtres)

    # ── Sélection et affichage ────────────────────────────────────────────────

    def _sur_clic_arbre(self, event=None):
        """Appelé quand l'utilisateur clique sur un nœud de l'arbre."""
        selection = self.arbre.selection()
        if not selection:
            return
        cle = self._item_vers_cle.get(selection[0])
        if cle:  # cle est None pour les nœuds groupes
            self._charger_table(cle)

    def _charger_table(self, cle):
        """Charge et affiche le contenu d'une table du référentiel."""
        entrees = self.referentiels.get(cle, [])
        groupe  = cle.split('.')[-2] if 'll' not in cle.split('.')[-2] else \
                  '.'.join(cle.split('.')[-3:-1])

        self.lbl_table.config(
            text=f"{self._nom_court(cle)}   ({len(entrees)} entrées)"
        )

        if not entrees:
            self.tableau.set_sheet_data([])
            self.tableau.headers([])
            self._donnees_table  = []
            self._colonnes_table = []
            return

        self._colonnes_table = [
            c for c in entrees[0].keys() if c not in self.COLS_MASQUEES
        ]
        self._donnees_table = [
            [str(e.get(c, '') or '') for c in self._colonnes_table]
            for e in entrees
        ]
        self.var_recherche_contenu.set('')
        self._afficher_donnees(self._donnees_table)

    def _afficher_donnees(self, donnees):
        """Charge les données dans tksheet et ajuste les largeurs."""
        self.tableau.set_sheet_data(donnees, redraw=False)
        self.tableau.headers(self._colonnes_table)
        for i, col in enumerate(self._colonnes_table):
            valeurs = [row[i] for row in donnees]
            largeur = max(
                55,
                max((len(v) for v in valeurs), default=0) * 8,
                len(col) * 9
            ) + 10
            self.tableau.column_width(column=i, width=largeur, redraw=False)
        self.tableau.redraw()

    def _filtrer_contenu(self):
        """Filtre les lignes du tableau selon la recherche dans le contenu."""
        terme = self.var_recherche_contenu.get().lower()
        if not terme:
            self._afficher_donnees(self._donnees_table)
            return
        filtrees = [l for l in self._donnees_table
                    if any(terme in c.lower() for c in l)]
        self._afficher_donnees(filtrees)


# ─────────────────────────────────────────────────────────────────────────────
# POINT D'ENTRÉE
# ─────────────────────────────────────────────────────────────────────────────

def lancer_seul():
    """Lance la Phase 1 dans sa propre fenêtre (mode standalone)."""
    lang.charger_langue(lang.langue_depuis_config())
    root = tk.Tk()
    root.title(T('app.titre'))
    root.geometry("1200x700")
    root.resizable(True, True)
    app = AppObserveTools(root)
    app.pack(fill='both', expand=True)
    root.mainloop()


if __name__ == '__main__':
    lancer_seul()
