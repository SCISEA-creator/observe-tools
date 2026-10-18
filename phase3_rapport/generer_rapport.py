#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phase 3 — Génération automatique du rapport Word.

Point d'entrée principal :
    generer_rapport(tables, zone_peche, output_path)
    - tables      : dict {cle: DataFrame} issu de la lecture du zip ObServe
    - zone_peche  : chaîne déjà construite, ex. "ZEE de Maurice et ZEE des Seychelles"
    - output_path : chemin complet du fichier .docx à créer

Lève une exception en cas d'erreur (à attraper dans main.py avec try/except).
"""
import sys
import os
import re
import tempfile
import types
import unicodedata

import matplotlib
matplotlib.use('Agg')   # pas de fenêtre graphique — obligatoire en mode app
import matplotlib.pyplot as plt
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Chemins clés
RACINE   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(RACINE, 'phase3_rapport', 'template_rapport.docx')

# Import des modules Phase 2 (calculs + génération figures)
sys.path.insert(0, os.path.join(RACINE, 'phase2_graphiques'))
import calculs as calc
# Réutilise l'instance déjà chargée par app.py si disponible
# (évite deux copies du module Phase 2 en mémoire)
if 'phase2_main' in sys.modules:
    ph2_main = sys.modules['phase2_main']
else:
    import main as ph2_main   # mode test standalone (_test_batch.py)

# Marqueur de position du tableau CPUE (remplacé en post-rendu)
MARQUEUR_CPUE = '__CPUE_TABLE__'

# ─────────────────────────────────────────────────────────────────────────────
# ZONES DE PÊCHE
# Correspondance noms courts (interface Accueil) → libellés complets pour le rapport.
# IMPORTANT : ce dict existe aussi dans app.py (liste déroulante de l'Accueil).
# Les deux doivent rester synchronisés si des zones sont ajoutées.
# ─────────────────────────────────────────────────────────────────────────────

ZONES_PECHE = {
    'Maurice':               'ZEE de Maurice',
    'Réunion':               'ZEE de La Réunion',
    'Madagascar':            'ZEE de Madagascar',
    'Europa':                "ZEE d'Europa",
    'Bassas da India':       'ZEE de Bassas da India',
    'Juan de Nova':          'ZEE de Juan de Nova',
    'Mayotte':               'ZEE de Mayotte',
    'Comores':               'ZEE des Comores',
    'Seychelles':            'ZEE des Seychelles',
    'Eaux Internationales':  'Eaux internationales',
    'Glorieuses':            'ZEE des Glorieuses',
    'Tromelin':              'ZEE de Tromelin',
    'Chagos':                'ZEE des Chagos',
    'Maldives':              'ZEE des Maldives',
}


def construire_zone_peche(zones_selectionnees):
    """
    Construit la chaine de zone(s) de peche a partir d'une liste de noms courts.
    Exemple : ['Maurice'] -> 'ZEE de Maurice'
    Exemple : ['Maurice', 'Seychelles'] -> 'ZEE de Maurice et ZEE des Seychelles'
    """
    libelles = [ZONES_PECHE.get(z, f'ZEE de {z}') for z in zones_selectionnees]
    if not libelles:
        return '[Zone de pêche non renseignée]'
    if len(libelles) == 1:
        return libelles[0]
    return ', '.join(libelles[:-1]) + ' et ' + libelles[-1]


# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES GÉNÉRAUX
# ─────────────────────────────────────────────────────────────────────────────

# Compteur d'avertissements de calcul — réinitialisé à chaque génération.
# Permet d'informer l'utilisateur si des calculs ont échoué silencieusement.
_avertissements_generation = []


def essayer(fn, *args, **kwargs):
    """Appelle fn(*args) en capturant les exceptions. Retourne None si erreur."""
    try:
        return fn(*args, **kwargs)
    except Exception as e:
        msg = f'{fn.__name__}: {e}'
        print(f'    Avertissement {msg}')
        _avertissements_generation.append(msg)
        return None


def get_val(df, col, defaut='—'):
    """Lit la premiere valeur non-nulle d'une colonne DataFrame."""
    if df is None or df.empty or col not in df.columns:
        return defaut
    v = df[col].dropna()
    return str(v.iloc[0]) if not v.empty else defaut


def fmt_date(df, col):
    """Formate une date DataFrame en JJ/MM/AAAA."""
    v = get_val(df, col)
    if v == '—':
        return v
    try:
        return pd.Timestamp(v).strftime('%d/%m/%Y')
    except Exception:
        return v


def duree_moy(df, c1, c2):
    """Calcule la duree moyenne entre deux colonnes timestamp. Retourne 'XHxx'."""
    if df is None or df.empty or c1 not in df.columns or c2 not in df.columns:
        return '—'
    try:
        d = pd.to_datetime(df[c2]) - pd.to_datetime(df[c1])
        m = d.mean()
        h, r = divmod(int(m.total_seconds()), 3600)
        return f'{h}H{r // 60:02d}'
    except Exception:
        return '—'


def _cle_colonne(nom):
    """'Date/Heure début' et 'Date/Heure debut' → même clé : sans accents, ni ponctuation, ni casse."""
    sans_accents = unicodedata.normalize('NFKD', str(nom)).encode('ascii', 'ignore').decode()
    return re.sub(r'[^a-z0-9]', '', sans_accents.lower())


def df_to_rows(df, mapping):
    """
    Convertit un DataFrame en liste de SimpleNamespace pour les tableaux docxtpl.
    Les colonnes sont retrouvées sans tenir compte des accents ni de la ponctuation :
    une différence d'accent entre calculs.py et ce fichier ne vide plus les tableaux.
    """
    if df is None or df.empty:
        return []
    colonnes = {_cle_colonne(c): c for c in df.columns}
    rows = []
    for _, ligne in df.iterrows():
        obj = types.SimpleNamespace()
        for attr, col in mapping.items():
            vraie = colonnes.get(_cle_colonne(col))
            v = ligne.get(vraie, '—') if vraie is not None else '—'
            setattr(obj, attr, '—' if str(v) in ('nan', 'None', 'NaT') else str(v))
        rows.append(obj)
    return rows


def df_param_valeur(df):
    """Convertit un DataFrame a 2 colonnes en liste (parametre, valeur) pour docxtpl."""
    if df is None or df.empty:
        return []
    cols = list(df.columns)
    return [types.SimpleNamespace(
        parametre=str(r[cols[0]]),
        valeur=str(r[cols[1]]) if len(cols) > 1 else '—'
    ) for _, r in df.iterrows()]


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.1 : GROUPES DE CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

# Accord du participe passe par groupe (feminin = 'capturees')
GROUPES_ACCORD_41 = {
    'Porte-épées':     'capturés',
    'Thons':           'capturés',
    'Requins':         'capturés',
    'Raies':           'capturées',
    'Poissons osseux': 'capturés',
    'Cétacés':         'capturés',
    'Tortues':         'capturées',
}

# Pluriels irreguliers : noms ObServe dont la pluralisation automatique est incorrecte.
# Ajouter ici les cas trouves sur de nouvelles marees.
NOMS_PLURIEL_OVERRIDE = {
    'Peau bleue':           'Requins peau bleue',
    'Thon obèse(=Patudo)': 'Thons obèses (patudos)',   # clé = nom commun ObServe exact, accent compris
}

# Codes FAO consideres commercialisables dans Other bony fishes
CODES_FAO_COMMERCIAUX = {'WAH', 'DOL'}


def _pluraliser_fr(nom):
    """
    Pluralise un nom commun francais simple (heuristique noms d'especes).
    Ajoute 's' au premier mot et aux adjectifs qui suivent.
    S'arrete aux prepositions pour ne pas pluraliser les complements du nom.
    """
    if not nom or ',' in nom:
        return nom
    mots = nom.split(' ')
    resultat = []
    PREP = {'a', 'de', 'du', "d'", 'des', 'en', 'et', 'ou', 'nca', 'nei', 'nep'}
    for i, mot in enumerate(mots):
        if i > 0 and mot.lower() in PREP:
            resultat.append(mot)
            resultat.extend(mots[i + 1:])
            break
        if not mot.endswith(('s', 'x', 'z')):
            mot = mot + 'x' if mot.endswith('eau') else mot + 's'
        resultat.append(mot)
    return ' '.join(resultat)


def _article_et_nombre(nom):
    """Retourne toujours le pluriel (Les ...) pour toutes les especes."""
    return "Les ", False


def _construire_especes_groupe(catch_df, groupe_en, nom_groupe_fr):
    """
    Construit la liste de phrases pour un groupe d'especes (generique).
    groupe_en     : nom anglais ObServe ('Billfishes', 'Sharks'...)
    nom_groupe_fr : libelle dans la phrase ('porte-epees', 'requins'...)
    """
    data = essayer(calc.calc_captures_par_espece_groupe, catch_df, groupe_en)
    if data is None or data.empty:
        return []

    subset = catch_df[catch_df['species_group'] == groupe_en]
    noms = {}
    if not subset.empty and 'nom_commun' in subset.columns:
        noms = subset.groupby('fao_code')['nom_commun'].first().to_dict()

    total = int(data.sum())
    if total == 0:
        return []
    nb_especes = len(data)
    especes = []

    for rang, (code_fao, nb) in enumerate(data.items(), 1):
        nb = int(nb)
        if nb == 0:
            continue
        pct = round(nb / total * 100, 1)
        nom_base = noms.get(code_fao, code_fao)
        if nom_base in NOMS_PLURIEL_OVERRIDE:
            nom_affiche = NOMS_PLURIEL_OVERRIDE[nom_base]
            article, singulier = "Les ", False
        else:
            article, singulier = _article_et_nombre(nom_base)
            nom_affiche = nom_base if singulier else _pluraliser_fr(nom_base)
        ind = 'individu' if nb == 1 else 'individus'

        if singulier:
            if nb_especes == 1:
                phrase = (f"{article}{nom_affiche} constitue la totalité des {nom_groupe_fr} "
                          f"({nb} {ind}, 100%).")
            elif rang == 1:
                if pct > 50:
                    phrase = (f"{article}{nom_affiche} représente la majorité des {nom_groupe_fr} "
                              f"avec {nb} {ind} ({pct}%).")
                else:
                    phrase = (f"{article}{nom_affiche} constitue le groupe dominant des {nom_groupe_fr} "
                              f"avec {nb} {ind} ({pct}%).")
            elif rang == 2:
                phrase = (f"{article}{nom_affiche} constitue le 2ème groupe le plus capturé "
                          f"avec {nb} {ind} ({pct}%).")
            else:
                phrase = (f"{article}{nom_affiche} a été capturé à hauteur de "
                          f"{nb} {ind} ({pct}%).")
        else:
            if nb_especes == 1:
                phrase = (f"{article}{nom_affiche} constituent la totalité des {nom_groupe_fr} "
                          f"({nb} {ind}, 100%).")
            elif rang == 1:
                if pct > 50:
                    phrase = (f"{article}{nom_affiche} représentent la majorité des {nom_groupe_fr} "
                              f"avec {nb} {ind} ({pct}%).")
                else:
                    phrase = (f"{article}{nom_affiche} constituent le groupe dominant des {nom_groupe_fr} "
                              f"avec {nb} {ind} ({pct}%).")
            elif rang == 2:
                phrase = (f"{article}{nom_affiche} constituent le 2ème groupe le plus capturé "
                          f"avec {nb} {ind} ({pct}%).")
            else:
                phrase = (f"{article}{nom_affiche} ont été capturés à hauteur de "
                          f"{nb} {ind} ({pct}%).")

        especes.append(types.SimpleNamespace(phrase=phrase))

    return especes


def construire_especes_epees(catch_df):
    return _construire_especes_groupe(catch_df, 'Billfishes', 'porte-épées')


def construire_especes_requins(catch_df):
    return _construire_especes_groupe(catch_df, 'Sharks', 'requins')


def construire_especes_thons(catch_df):
    return _construire_especes_groupe(catch_df, 'Tunas nei', 'thonidés')


def construire_especes_autres_comm(catch_df):
    """
    Especes commercialisables (WAH, DOL) dans Other bony fishes.
    Retourne [] si aucune -> la phrase intro disparait dans le template.
    """
    if catch_df is None or catch_df.empty:
        return []
    tmp = catch_df[
        (catch_df['species_group'] == 'Other bony fishes') &
        (catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX))
    ].copy()
    if tmp.empty:
        return []
    tmp['species_group'] = '__COMM__'
    return _construire_especes_groupe(tmp, '__COMM__', 'espèces commercialisables')


def construire_especes_autres_non_comm(catch_df):
    """
    Especes non-commercialisables dans Other bony fishes (toutes sauf WAH et DOL).
    Retourne [] si aucune -> la phrase intro disparait dans le template.
    """
    if catch_df is None or catch_df.empty:
        return []
    tmp = catch_df[
        (catch_df['species_group'] == 'Other bony fishes') &
        (~catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX))
    ].copy()
    if tmp.empty:
        return []
    tmp['species_group'] = '__NON_COMM__'
    return _construire_especes_groupe(tmp, '__NON_COMM__', 'espèces accessoires')


def construire_groupes_41(catch_df):
    """
    Phrases pour la section 4.1 global : liste triee par nombre de captures decroissant.
    Retourne une liste de SimpleNamespace(phrase=...) pour la boucle docxtpl.
    """
    data = essayer(calc.calc_captures_par_groupe, catch_df)
    if data is None or data.empty:
        return []

    total = int(data.sum())
    if total == 0:
        return []
    nb_groupes = len(data)
    groupes = []

    for rang, (nom, nb) in enumerate(data.items(), 1):
        nb = int(nb)
        if nb == 0:
            continue
        pct = round(nb / total * 100, 1)
        accord = GROUPES_ACCORD_41.get(nom, 'captures')
        ind = 'individu' if nb == 1 else 'individus'

        if nb_groupes == 1:
            phrase = (f"Les {nom} constituent la totalité des captures "
                      f"({nb} {ind}, 100%).")
        elif rang == 1:
            if pct > 50:
                phrase = (f"Les {nom} représentent la majorité des captures "
                          f"avec {nb} {ind} ({pct}%).")
            else:
                phrase = (f"Les {nom} constituent le groupe dominant des captures "
                          f"avec {nb} {ind} ({pct}%).")
        elif rang == 2:
            phrase = (f"Les {nom} constituent le 2ème groupe le plus capturé "
                      f"avec {nb} {ind} ({pct}%).")
        else:
            phrase = (f"Les {nom} ont été {accord} à hauteur de "
                      f"{nb} {ind} ({pct}%).")

        groupes.append(types.SimpleNamespace(phrase=phrase))

    return groupes


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.3 : ETAT A LA CAPTURE
# ─────────────────────────────────────────────────────────────────────────────

def _etat_majoritaire(catch_df, filtre):
    """
    Retourne 'morts' ou 'vivants' selon la majorite de status_at_capture
    pour les captures conservees (catch_fate='RET') correspondant au filtre.
    A/AW = vivant, D = mort. En cas d'egalite ou donnees absentes -> 'morts'.
    """
    if catch_df is None or catch_df.empty:
        return 'morts'
    ret = catch_df[(catch_df['catch_fate'] == 'RET') & filtre(catch_df)]
    if ret.empty:
        return 'morts'
    nb_vivants = ret[ret['status_at_capture'].isin(['A', 'AW'])]['count'].sum()
    nb_morts   = ret[ret['status_at_capture'] == 'D']['count'].sum()
    return 'vivants' if nb_vivants > nb_morts else 'morts'


def etat_espadons(catch_df):
    """'morts' ou 'vivants' selon la majorite des SWO conserves."""
    return _etat_majoritaire(catch_df, lambda df: df['fao_code'] == 'SWO')


def etat_thons_hors_germons(catch_df):
    """'morts' ou 'vivants' selon la majorite des thonides (sans ALB) conserves."""
    return _etat_majoritaire(
        catch_df,
        lambda df: (df['species_group'] == 'Tunas nei') & (df['fao_code'] != 'ALB')
    )


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.4 : SURVIE REQUINS
# ─────────────────────────────────────────────────────────────────────────────

def construire_survie_requins(catch_df):
    """
    Chiffres de survie requins pour le template §4.4.
    Retourne un dict avec 4 cles pretes pour le contexte docxtpl.
    """
    vide = {
        'requins_vivants_nb': 0, 'requins_vivants_pct': '0,0',
        'requins_morts_nb':   0, 'requins_morts_pct':   '0,0',
    }
    data = essayer(calc.calc_survie_requins, catch_df)
    if data is None or data.empty:
        return vide
    vivants = int(data.get('Vivant', 0))
    morts   = int(data.get('Mort',   0))
    total   = vivants + morts
    if total == 0:
        return vide
    return {
        'requins_vivants_nb':  vivants,
        'requins_vivants_pct': f'{100 * vivants / total:.1f}'.replace('.', ','),
        'requins_morts_nb':    morts,
        'requins_morts_pct':   f'{100 * morts / total:.1f}'.replace('.', ','),
    }


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.5 : DEVENIR DES CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

def construire_devenir(catch_df):
    """
    Chiffres devenir global pour le template §4.5.
    Retourne un dict avec 4 cles pretes pour le contexte docxtpl.
    """
    vide = {
        'devenir_ret_nb': 0, 'devenir_ret_pct': '0,0',
        'devenir_dis_nb': 0, 'devenir_dis_pct': '0,0',
    }
    data = essayer(calc.calc_devenir_global, catch_df)
    if data is None or data.empty:
        return vide
    ret   = int(data.get('Conservés', 0))
    dis   = int(data.get('Rejetés',   0))
    total = int(data.sum())
    if total == 0:
        return vide
    return {
        'devenir_ret_nb':  ret,
        'devenir_ret_pct': f'{100 * ret / total:.1f}'.replace('.', ','),
        'devenir_dis_nb':  dis,
        'devenir_dis_pct': f'{100 * dis / total:.1f}'.replace('.', ','),
    }


def construire_especes_conservees(catch_df):
    """Phrase listant toutes les especes conservees (RET), triees par nb decroissant."""
    if catch_df is None or catch_df.empty:
        return '[espèces conservées — données manquantes]'
    ret = catch_df[catch_df['catch_fate'] == 'RET'].copy()
    if ret.empty:
        return '[aucune espèce conservée]'
    grouped = (ret.groupby('fao_code')
               .agg(nb=('count', 'sum'), nom=('nom_commun', 'first'))
               .reset_index()
               .sort_values('nb', ascending=False))
    parties = []
    for _, row in grouped.iterrows():
        nb = int(row['nb'])
        nom_base = str(row['nom']) if pd.notna(row['nom']) and row['nom'] else row['fao_code']
        nom_pl = NOMS_PLURIEL_OVERRIDE.get(nom_base, _pluraliser_fr(nom_base))
        ind = 'individu' if nb == 1 else 'individus'
        parties.append(f'les {nom_pl.lower()} avec {nb} {ind}')
    if not parties:
        return '[aucune espèce conservée]'
    if len(parties) == 1:
        return f'Les espèces conservées pendant la marée sont {parties[0]}.'
    return ('Les espèces conservées pendant la marée sont '
            + ', '.join(parties[:-1]) + ' et ' + parties[-1] + '.')


def construire_especes_accessoires_rejetees(catch_df):
    """Phrase sur les especes accessoires rejetees/echappees (top 3)."""
    if catch_df is None or catch_df.empty:
        return '[espèces accessoires rejetées — données manquantes]'
    mask_commercial = (
        catch_df['species_group'].isin({'Billfishes', 'Tunas nei'}) |
        catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX)
    )
    rej = catch_df[catch_df['catch_fate'].isin(['DIS', 'ESC']) & ~mask_commercial].copy()
    if rej.empty:
        return '[aucune espece accessoire rejetee]'
    grouped = (rej.groupby('fao_code')
               .agg(nb=('count', 'sum'), nom=('nom_commun', 'first'))
               .reset_index()
               .sort_values('nb', ascending=False)
               .head(3))
    parties = []
    for _, row in grouped.iterrows():
        nb = int(row['nb'])
        nom_base = str(row['nom']) if pd.notna(row['nom']) and row['nom'] else row['fao_code']
        nom_pl = NOMS_PLURIEL_OVERRIDE.get(nom_base, _pluraliser_fr(nom_base))
        ind = 'individu' if nb == 1 else 'individus'
        parties.append(f'des {nom_pl.lower()} avec {nb} {ind}')
    if not parties:
        return '[aucune espèce accessoire rejetée]'
    liste = (', '.join(parties[:-1]) + ' et ' + parties[-1]) if len(parties) > 1 else parties[0]
    return (f'Des espèces accessoires ont été rejetées, notamment '
            f'{liste} relâchés sur l\'ensemble de la marée.')


def construire_especes_commerciales_rejetees(catch_df):
    """Phrase sur les especes commerciales rejetees/echappees."""
    if catch_df is None or catch_df.empty:
        return '[espèces commerciales rejetées — données manquantes]'
    mask_commercial = (
        catch_df['species_group'].isin({'Billfishes', 'Tunas nei'}) |
        catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX)
    )
    rej = catch_df[catch_df['catch_fate'].isin(['DIS', 'ESC']) & mask_commercial].copy()
    if rej.empty:
        return '[aucune espece commerciale rejetee]'
    grouped = (rej.groupby('fao_code')
               .agg(nb=('count', 'sum'), nom=('nom_commun', 'first'))
               .reset_index()
               .sort_values('nb', ascending=False))
    parties = []
    for _, row in grouped.iterrows():
        nb = int(row['nb'])
        nom_base = str(row['nom']) if pd.notna(row['nom']) and row['nom'] else row['fao_code']
        nom_pl = NOMS_PLURIEL_OVERRIDE.get(nom_base, _pluraliser_fr(nom_base))
        ind = 'individu' if nb == 1 else 'individus'
        parties.append(f'les {nom_pl.lower()} avec {nb} {ind}')
    if not parties:
        return '[aucune espèce commerciale rejetée]'
    liste = (', '.join(parties[:-1]) + ' et ' + parties[-1]) if len(parties) > 1 else parties[0]
    return f'Les espèces commerciales ont aussi été rejetées comme {liste}.'


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.6 : DEPREDATION
# ─────────────────────────────────────────────────────────────────────────────

def construire_depredation_par1(catch_df):
    """
    Para 1 §4.6 : depredateur principal + especes commercialisables touchees.
    Exclut ISB (Squalelet feroce).
    """
    EXCLURE = {calc.DEPREDATEURS_FR['ISB']}
    if catch_df is None or catch_df.empty:
        return "Aucune déprédation n'a été observée durant cette marée."
    mask_commercial = (
        catch_df['species_group'].isin({'Billfishes', 'Tunas nei'}) |
        catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX)
    )
    data = essayer(calc.calc_depredation_totale, catch_df)
    if data is None or data.empty:
        return "Aucune déprédation n'a été observée durant cette marée."
    # Filtrage ISB AVANT le calcul du total : le dénominateur exclut ISB
    # car la phrase porte sur les déprédateurs qui empêchent la commercialisation
    data = data[~data.index.isin(EXCLURE)]
    if data.empty:
        return "Aucune déprédation n'a été observée durant cette marée."
    dep1_nom = data.index[0]
    # Nombre brut de captures attaquées par dep1 (composites décomposés = 1 par prédateur)
    # On n'utilise PAS data.sum() comme dénominateur : il peut dépasser le nb de captures
    # si des composites existent → pourcentages sous-estimés et faux.
    dep1_count = int(data.iloc[0])
    dep1_capture_s = 'captures' if dep1_count > 1 else 'capture'
    dep1_codes = {code for code, nom in calc.DEPREDATEURS_FR.items() if nom == dep1_nom}
    mask_dep1 = catch_df['depredating_species'].fillna('').apply(
        lambda x: bool(dep1_codes & {c.strip() for c in str(x).split(';') if c.strip()})
    )
    subset = catch_df[mask_dep1 & mask_commercial]
    debut = (
        f"La majorité de la déprédation empêchant la commercialisation du poisson "
        f"a été faite par {_pluraliser_fr(dep1_nom).lower()} "
        f"({dep1_count}\xa0{dep1_capture_s}, Figure\xa015)"
    )
    if subset.empty:
        return debut + "."
    counts = subset.groupby('nom_commun')['count'].sum().sort_values(ascending=False)
    parties = []
    for esp, nb in counts.items():
        nb = int(nb)
        nom_pl = NOMS_PLURIEL_OVERRIDE.get(esp, _pluraliser_fr(esp)) if nb > 1 else esp
        parties.append(f"{nb} {nom_pl.lower()}")
    return f"{debut} avec {', '.join(parties)} (Figure\xa016)."


def construire_depredations_secondaires(catch_df):
    """
    Para 2+ §4.6 : depredateurs de rang 2+ (hors ISB).
    Retourne une liste de dict {'phrase': '...'} pour la boucle docxtpl.
    Liste vide si moins de 2 depredateurs.
    """
    EXCLURE = {calc.DEPREDATEURS_FR['ISB']}
    if catch_df is None or catch_df.empty:
        return []
    mask_commercial = (
        catch_df['species_group'].isin({'Billfishes', 'Tunas nei'}) |
        catch_df['fao_code'].isin(CODES_FAO_COMMERCIAUX)
    )
    data = essayer(calc.calc_depredation_totale, catch_df)
    if data is None or data.empty:
        return []
    data = data[~data.index.isin(EXCLURE)]
    result = []
    for i, (dep_nom, dep_count) in enumerate(data.items()):
        if i == 0:
            continue
        # Nombre brut de captures (même logique que construire_depredation_par1)
        nb = int(dep_count)
        capture_s = 'captures' if nb > 1 else 'capture'
        dep_codes = {code for code, nom in calc.DEPREDATEURS_FR.items() if nom == dep_nom}
        mask_dep = catch_df['depredating_species'].fillna('').apply(
            lambda x, dc=dep_codes: bool(dc & {c.strip() for c in str(x).split(';') if c.strip()})
        )
        subset = catch_df[mask_dep & mask_commercial]
        debut = (
            f"{_pluraliser_fr(dep_nom)} sont responsables de la déprédation "
            f"sur {nb}\xa0{capture_s} (Figure\xa015)"
        )
        if subset.empty:
            phrase = debut + "."
        else:
            counts = subset.groupby('nom_commun')['count'].sum().sort_values(ascending=False)
            parties = []
            for esp, nb in counts.items():
                nb = int(nb)
                nom_pl = NOMS_PLURIEL_OVERRIDE.get(esp, _pluraliser_fr(esp)) if nb > 1 else esp
                parties.append(f"{nb} {nom_pl.lower()}")
            phrase = f"{debut} avec {', '.join(parties)} (Figure\xa016)."
        result.append({'phrase': phrase})
    return result


def construire_depredation_isb(catch_df):
    """
    Phrase sur la depredation par les squalelets feroces (ISB).
    Retourne '' si pas d'ISB (paragraphe masque dans le template).
    """
    if catch_df is None or catch_df.empty:
        return ''
    if 'depredating_species' not in catch_df.columns:
        return ''
    mask_isb = catch_df['depredating_species'].fillna('').apply(
        lambda x: 'ISB' in {c.strip() for c in str(x).split(';')}
    )
    nb = int(catch_df[mask_isb]['count'].sum())
    if nb == 0:
        return ''
    ind = 'individu' if nb == 1 else 'individus'
    return (
        f"La déprédation par les squalelets féroces concerne {nb}\xa0{ind}. "
        f"Ce type de déprédation n'a pas vraiment d'influence sur la pêche, "
        f"car le prélèvement est minime et n'empêche pas la commercialisation du poisson."
    )


def construire_depredateurs_phrase(catch_df):
    """Phrase listant les depredateurs presents (hors ISB) pour §4.5."""
    EXCLURE = {calc.DEPREDATEURS_FR['ISB']}
    data = essayer(calc.calc_depredation_totale, catch_df)
    if data is None or data.empty:
        return '[depredateurs — donnees manquantes]'
    presents = [nom for nom, nb in data.items() if nb > 0 and nom not in EXCLURE]
    if not presents:
        return '[depredateurs — donnees manquantes]'
    parties = [f'les {_pluraliser_fr(nom).lower()}' for nom in presents]
    liste = (', '.join(parties[:-1]) + ' et ' + parties[-1]) if len(parties) > 1 else parties[0]
    return (f'Les rejets de poissons commercialisables sont dus à la déprédation par '
            f'{liste} qui rendent le poisson impropre à la consommation.')


# ─────────────────────────────────────────────────────────────────────────────
# TEXTE AUTOMATIQUE — SECTION 4.7 : CPUE
# ─────────────────────────────────────────────────────────────────────────────

def calc_cpue_texte_47(catch_df, set_df):
    """
    Max et min de CPUE totale par filage pour le texte §4.7.
    Retourne (cpue_max, cpue_min) formates en notation francaise (virgule decimale).
    """
    df = essayer(calc.calc_tableau_cpue_rapport, catch_df, set_df)
    if df is None or df.empty or 'CPUE totale' not in df.columns:
        return '—', '—'
    col = df['CPUE totale'].dropna()
    if col.empty:
        return '—', '—'

    def fmt(v):
        return f'{v:.2f}'.replace('.', ',')

    return fmt(col.max()), fmt(col.min())


# ─────────────────────────────────────────────────────────────────────────────
# TABLEAU CPUE (post-rendu, colonnes dynamiques)
# ─────────────────────────────────────────────────────────────────────────────

def _inserer_tableau_cpue(doc, catch_df, set_df):
    """
    Cherche le paragraphe marqueur MARQUEUR_CPUE dans le doc rendu et le remplace
    par un tableau CPUE stylise (colonnes dynamiques selon les especes de la maree).
    """
    df = essayer(calc.calc_tableau_cpue_rapport, catch_df, set_df)

    para_cible = None
    for para in doc.paragraphs:
        if MARQUEUR_CPUE in para.text:
            para_cible = para
            break

    if para_cible is None:
        print('  Avertissement : marqueur CPUE non trouve dans le doc rendu')
        return
    if df is None or df.empty:
        para_cible.clear()
        para_cible.add_run('[Pas de donnees CPUE disponibles]')
        return

    cols     = list(df.columns)
    n_cols   = len(cols)
    n_data   = len(df)
    ENTETE_HEX = 'C0C0C0'
    BORD_HEX   = 'D0D0D0'

    set_pct      = max(12, min(20, 200 // n_cols))
    autre_pct    = (100 - set_pct) / (n_cols - 1)
    largeurs_pct = [set_pct] + [autre_pct] * (n_cols - 1)

    table = doc.add_table(rows=1 + n_data, cols=n_cols)
    table.style = 'Table Grid'

    tblPr = table._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    tblBorders = OxmlElement('w:tblBorders')
    for cote in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{cote}')
        b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0');    b.set(qn('w:color'), BORD_HEX)
        tblBorders.append(b)
    tblPr.append(tblBorders)

    for old_grid in table._tbl.findall(qn('w:tblGrid')):
        table._tbl.remove(old_grid)
    tblGrid = OxmlElement('w:tblGrid')
    for pct in largeurs_pct:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(int(pct * 50)))
        tblGrid.append(gridCol)
    tblPr.addnext(tblGrid)

    def _shd(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for s in tcPr.findall(qn('w:shd')):
            tcPr.remove(s)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)

    def _valign(cell, valeur='center'):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn('w:vAlign')):
            tcPr.remove(old)
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), valeur)
        tcPr.append(vAlign)

    def _set_tcW(cell, pct):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn('w:tcW')):
            tcPr.remove(old)
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), str(int(pct * 50)))
        tcW.set(qn('w:type'), 'pct')
        tcPr.append(tcW)

    for j, col_name in enumerate(cols):
        cell = table.rows[0].cells[j]
        _shd(cell, ENTETE_HEX)
        _valign(cell, 'center')
        _set_tcW(cell, largeurs_pct[j])
        para = cell.paragraphs[0]
        para.clear()
        run = para.add_run(col_name)
        run.bold = True
        run.font.size = Pt(9)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (_, row_d) in enumerate(df.iterrows()):
        tr = table.rows[i + 1]
        trPr = tr._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            tr._tr.insert(0, trPr)
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '340'); trH.set(qn('w:hRule'), 'exact')
        trPr.append(trH)
        for j, (col_name, val) in enumerate(zip(cols, row_d)):
            cell = tr.cells[j]
            _shd(cell, 'FFFFFF')
            _valign(cell, 'center')
            _set_tcW(cell, largeurs_pct[j])
            para = cell.paragraphs[0]
            para.clear()
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after  = Pt(0)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if col_name == 'SET':
                txt = str(val)
            elif col_name == 'Nb hamecons':
                txt = str(int(val)) if pd.notna(val) else '0'
            else:
                txt = f'{float(val):.2f}' if pd.notna(val) else '0.00'
            run = para.add_run(txt)
            run.font.size = Pt(9)

    p_el   = para_cible._p
    parent = p_el.getparent()
    p_idx  = list(parent).index(p_el)
    tbl_el = table._tbl
    parent.remove(tbl_el)
    parent.remove(p_el)
    parent.insert(p_idx, tbl_el)
    print(f'  Tableau CPUE insere ({n_data} calees x {n_cols} colonnes)')


# ─────────────────────────────────────────────────────────────────────────────
# GÉNÉRATION DES FIGURES
# ─────────────────────────────────────────────────────────────────────────────

# Mapping noms de fichiers Phase 2 -> tags images Phase 3
_PH2_FICHIER_VERS_TAG = {
    'fig_01_captures_groupes.png':    'image_captures_groupes',
    'fig_02_porte_epees.png':         'image_porte_epees',
    'fig_03_thonides.png':            'image_thonides',
    'fig_04_requins.png':             'image_requins',
    'fig_05_autres_poissons.png':     'image_autres_poissons',
    'fig_06_raies.png':               'image_raies',
    'fig_07_tortues.png':             'image_tortues',
    'fig_08_cetaces.png':             'image_cetaces',
    'fig_09_taille_swo.png':          'image_taille_swo',
    'fig_10_taille_thons.png':        'image_taille_thons',
    'fig_11_etat_capture.png':        'image_etat_capture',
    'fig_12_survie_requins.png':      'image_survie_requins',
    'fig_13_devenir_global.png':      'image_devenir_global',
    'fig_14_devenir_especes.png':     'image_devenir_especes',
    'fig_15_depredation_totale.png':  'image_depred_totale',
    'fig_16_depredation_especes.png': 'image_depred_especes',
    'fig_17_cpue.png':                'image_cpue',
    'fig_18_profil_tdr.png':          'image_profil_tdr',
    'fig_19_carte_filages.png':       'image_carte',
}


def generer_figures(tables, tmp_dir):
    """
    Genere les figures via Phase 2 (generer_tout) — resultat identique a l'app.
    Retourne un dict {tag_phase3: chemin_png_ou_None}.
    """
    ph2_main.generer_tout(tables, tmp_dir)
    figs = {}
    for nom_fich, tag in _PH2_FICHIER_VERS_TAG.items():
        chemin = os.path.join(tmp_dir, nom_fich)
        figs[tag] = chemin if os.path.exists(chemin) else None
    # Photo TDR specifique : fournie manuellement par l'observateur, pas generee
    figs['image_tdr_specifique'] = None
    ok = sum(1 for v in figs.values() if v is not None)
    print(f'  Figures Phase 2 : {ok}/{len(figs) - 1} generees')
    return figs


# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES EXPORTÉS
# ─────────────────────────────────────────────────────────────────────────────

def calculer_nom_dossier(trip_df):
    """
    Construit le nom standard du dossier de maree a partir de la table 1_trip.
    Format : {trip_id}_{YYYYMMDD_debut}-{YYYYMMDD_fin}_{NAVIRE}_{NOM_OBS}
    Exemple : SEALORXXX_20260101-20260115_NAVIRE_NOM
    Retourne 'RAPPORT' si les donnees sont insuffisantes.
    """
    try:
        trip_id = get_val(trip_df, 'trip_id', 'INCONNU')
        try:
            d_debut = pd.Timestamp(get_val(trip_df, 'trip_start_date')).strftime('%Y%m%d')
        except Exception:
            d_debut = 'XXXXXXXX'
        try:
            d_fin = pd.Timestamp(get_val(trip_df, 'trip_end_date')).strftime('%Y%m%d')
        except Exception:
            d_fin = 'XXXXXXXX'
        navire_code = str(get_val(trip_df, 'vessel', 'NAVIRE')).strip().upper().replace(' ', '')
        nom_obs     = str(get_val(trip_df, 'observer_lastname', 'OBS')).strip().upper()
        return f'{trip_id}_{d_debut}-{d_fin}_{navire_code}_{nom_obs}'
    except Exception:
        return 'RAPPORT'


# ─────────────────────────────────────────────────────────────────────────────
# POINT D'ENTRÉE PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def generer_rapport(tables, zone_peche, output_path):
    """
    Genere le rapport Word complet a partir des tables ObServe.

    Parametres
    ----------
    tables      : dict {cle: DataFrame} — sortie de la lecture du zip ObServe
    zone_peche  : str — libelle complet, ex. "ZEE de Maurice et ZEE des Seychelles"
                  Construire avec construire_zone_peche() si besoin.
    output_path : str — chemin complet du .docx a creer

    Leve une exception en cas d'erreur.
    Retourne le nombre d'avertissements de calcul (0 = tout OK).
    """
    global _avertissements_generation
    _avertissements_generation = []   # reset pour cette génération

    # ── Validation de base ────────────────────────────────────────────────────
    if '1_trip' not in tables or tables['1_trip'].empty:
        raise ValueError("Table '1_trip' absente ou vide — impossible de generer le rapport.")
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"Template introuvable : {TEMPLATE}")

    # ── Extraction des DataFrames ─────────────────────────────────────────────
    trip_df   = tables.get('1_trip',   pd.DataFrame())
    set_df    = tables.get('2_set',    pd.DataFrame())
    catch_df  = tables.get('8_catch',  pd.DataFrame())
    # La colonne 'count' indique le nb d'individus par ligne (autoéchantillonnage).
    # Elle peut être absente selon la version d'ObServe — on met 1 par défaut.
    if not catch_df.empty and 'count' not in catch_df.columns:
        catch_df = catch_df.copy()
        catch_df['count'] = 1
    sample_df = tables.get('9_sample', pd.DataFrame())
    tdr_df    = tables.get('7_tdr',    pd.DataFrame())

    # ── Infos generales ───────────────────────────────────────────────────────
    trip_id     = get_val(trip_df, 'trip_id', 'INCONNU')
    navire      = get_val(trip_df, 'vessel', '—')
    observateur = (get_val(trip_df, 'observer_firstname', '') + ' ' +
                   get_val(trip_df, 'observer_lastname', '')).strip() or '—'
    capitaine   = (get_val(trip_df, 'captain_firstname', '') + ' ' +
                   get_val(trip_df, 'captain_lastname', '')).strip() or '—'
    indicatif   = get_val(trip_df, 'vessel_callsign', get_val(trip_df, 'vessel_flag', '—'))
    date_debut  = fmt_date(trip_df, 'trip_start_date')
    date_fin    = fmt_date(trip_df, 'trip_end_date')

    # Nom du dossier de maree — source unique : calculer_nom_dossier()
    nom_dossier = calculer_nom_dossier(trip_df)

    try:
        duree_maree = str((pd.Timestamp(get_val(trip_df, 'trip_end_date')) -
                           pd.Timestamp(get_val(trip_df, 'trip_start_date'))).days)
    except Exception:
        duree_maree = '—'

    nb_filages     = str(len(set_df)) if not set_df.empty else '—'
    total_hamecons = (str(int(set_df['total_hooks_count'].sum()))
                      if not set_df.empty and 'total_hooks_count' in set_df.columns else '—')
    duree_moy_filage = duree_moy(set_df, 'setting_start_timestamp', 'setting_end_timestamp')
    duree_moy_virage = duree_moy(set_df, 'hauling_start_timestamp', 'hauling_end_timestamp')

    try:
        ts = pd.to_datetime(set_df['setting_start_timestamp'].dropna())
        sec = ts.dt.hour * 3600 + ts.dt.minute * 60 + ts.dt.second
        moy_s = int(sec.mean())
        heure_moy_debut = f'{moy_s // 3600}H{(moy_s % 3600) // 60:02d}'
    except Exception:
        heure_moy_debut = '—'

    try:
        bl_df = tables.get('4_branchline', pd.DataFrame())
        longueur_avancon = int(bl_df['branchline_length_m'].mode()[0])
    except Exception:
        longueur_avancon = '?'

    try:
        tdr_pos_1 = int(tdr_df['branchline_setting_id'].mode()[0])
    except Exception:
        tdr_pos_1 = '?'
    try:
        tdr_pos_2 = int(tdr_df['branchline_hauling_id'].mode()[0])
    except Exception:
        tdr_pos_2 = '?'
    try:
        hamecons_par_basket = int(set_df['branchlines_per_basket_count'].mode()[0])
    except Exception:
        hamecons_par_basket = '?'

    try:
        depths = tdr_df['median_fishing_depth'].dropna()
        prof_min_mediane = int(round(depths.min() / 5) * 5)
        prof_max_mediane = int(round(depths.max() / 5) * 5)
    except Exception:
        prof_min_mediane = '?'
        prof_max_mediane = '?'

    print(f'  {trip_id} | {navire} | {observateur} | {date_debut} -> {date_fin} ({duree_maree}j)')

    # ── Tableaux ──────────────────────────────────────────────────────────────
    _engin = essayer(calc.calc_tableaux_engin, tables)
    tableaux_engin = _engin if _engin is not None else {}
    _fv = essayer(calc.calc_tableaux_filage_virage, set_df)
    df_filage, df_virage = _fv if _fv is not None else (pd.DataFrame(), pd.DataFrame())
    _tdr = essayer(calc.calc_tableau_tdr, tdr_df, set_df)
    df_tdr = _tdr if _tdr is not None else pd.DataFrame()
    _cap = essayer(calc.calc_tableau_captures, catch_df, sample_df)
    df_captures = _cap if _cap is not None else pd.DataFrame()

    engin_rows    = df_param_valeur(tableaux_engin.get('engin'))
    orins_rows    = df_to_rows(tableaux_engin.get('orins'), {
        'type_ligne': 'Type de ligne', 'longueur': "Longueur de l'orin (m)",
        'proportion': 'Proportion (%)'})
    hamecons_rows = df_to_rows(tableaux_engin.get('hamecons'), {
        'type_hamecon': "Type d'hamecon", 'proportion': 'Proportion (%)'})
    appats_rows   = df_param_valeur(tableaux_engin.get('appats'))
    filage_rows   = df_to_rows(df_filage, {
        'set_id': 'SET', 'dt_debut': 'Date/Heure debut',
        'lat_debut': 'Lat. debut', 'lon_debut': 'Lon. debut',
        'dt_fin': 'Date/Heure fin', 'lat_fin': 'Lat. fin', 'lon_fin': 'Lon. fin',
        'vitesse': 'Vitesse (nd)', 'nb_hamecons': 'Nb hamecons'})
    virage_rows   = df_to_rows(df_virage, {
        'set_id': 'SET', 'dt_debut': 'Date/Heure debut',
        'lat_debut': 'Lat. debut', 'lon_debut': 'Lon. debut',
        'dt_fin': 'Date/Heure fin', 'lat_fin': 'Lat. fin', 'lon_fin': 'Lon. fin',
        'coupures': 'Coupures'})
    tdr_rows      = df_to_rows(df_tdr, {
        'set_id': 'SET', 'tdr_id': 'TDR ID', 'tdr_serial': 'N serie',
        'orin1': 'Orin 1 (m)', 'orin2': 'Orin 2 (m)',
        'pos_pose': 'Pos. pose', 'pos_virage': 'Pos. virage',
        'prof_min': 'Prof. min (m)', 'prof_max': 'Prof. max (m)',
        'prof_med': 'Prof. med. (m)', 'especes': 'Espece(s) cap.'})
    captures_rows = df_to_rows(df_captures, {
        'groupe': 'Groupe', 'nom_scientifique': 'Nom scientifique',
        'nom_commun': 'Nom commun', 'code_fao': 'Code FAO',
        'nb_captures': 'Nb captures', 'nb_conserves': 'Nb conserves',
        'nb_rejetes': 'Nb rejetes', 'nb_echappes': 'Nb echappes',
        'nb_depredates': 'Nb depredates', 'nb_mesures': 'Nb mesures',
        'nb_vivants': 'Nb vivants capture', 'nb_morts': 'Nb morts capture'})

    print(f'  Filage={len(filage_rows)} | Virage={len(virage_rows)} | '
          f'TDR={len(tdr_rows)} | Captures={len(captures_rows)}')

    # ── Booleens de sections conditionnelles ──────────────────────────────────
    _depred_data = essayer(calc.calc_depredation_totale, catch_df)
    has_depredation = _depred_data is not None and not _depred_data.empty

    # has_requins = au moins 1 requin capturé (quelle que soit sa fate RET/DIS/ESC)
    # Indépendant de calc_survie_requins qui retourne None si aucun requin relâché
    has_requins = (
        'species_group' in catch_df.columns
        and (catch_df['species_group'] == 'Sharks').any()
    )

    # ── Figures ───────────────────────────────────────────────────────────────
    with tempfile.TemporaryDirectory() as tmp_dir:
        figs = generer_figures(tables, tmp_dir)

        # ── Rendu docxtpl ─────────────────────────────────────────────────────
        tpl = DocxTemplate(TEMPLATE)

        # make_img est definie ici pour capturer tpl (InlineImage en a besoin)
        # Si l'image est absente, retourne None -> le template affiche '[image]'
        def make_img(chemin, largeur_cm=17.8):
            if chemin and os.path.exists(chemin):
                return InlineImage(tpl, chemin, width=Cm(largeur_cm))
            return None

        contexte = {
            'trip_id':         trip_id,
            'navire':          navire,
            'observateur':     observateur,
            'capitaine':       capitaine,
            'zone_peche':      zone_peche,
            'nom_dossier':     nom_dossier,
            'has_tdr':         not tdr_df.empty,
            'has_depredation': has_depredation,
            'has_requins':     has_requins,
            'indicatif_radio': indicatif,
            'date_debut':      date_debut,
            'date_fin':        date_fin,
            'duree_maree':     duree_maree,
            'nb_filages':      nb_filages,
            'zone_texte':      zone_peche,
            'total_hamecons':  total_hamecons,
            'duree_moy_filage': duree_moy_filage,
            'heure_moy_debut':  heure_moy_debut,
            'duree_moy_virage': duree_moy_virage,
            'longueur_avancon': longueur_avancon,
            'tdr_pos_1':        tdr_pos_1,
            'tdr_pos_2':        tdr_pos_2,
            'hamecons_par_basket':  hamecons_par_basket,
            'prof_min_mediane': prof_min_mediane,
            'prof_max_mediane': prof_max_mediane,
            'groupes_41':               construire_groupes_41(catch_df),
            'especes_epees':            construire_especes_epees(catch_df),
            'especes_requins':          construire_especes_requins(catch_df),
            'especes_thons':            construire_especes_thons(catch_df),
            'especes_autres_comm':      construire_especes_autres_comm(catch_df),
            'especes_autres_non_comm':  construire_especes_autres_non_comm(catch_df),
            'especes_conservees':               construire_especes_conservees(catch_df),
            'especes_accessoires_rejetees':     construire_especes_accessoires_rejetees(catch_df),
            'especes_commerciales_rejetees':    construire_especes_commerciales_rejetees(catch_df),
            'depredation_par1':         construire_depredation_par1(catch_df),
            'depredations_secondaires': construire_depredations_secondaires(catch_df),
            'depredation_isb':          construire_depredation_isb(catch_df),
            'depredateurs_phrase':      construire_depredateurs_phrase(catch_df),
            'etat_espadons':            etat_espadons(catch_df),
            'etat_thons_hors_germons':  etat_thons_hors_germons(catch_df),
            **dict(zip(('cpue_max_47', 'cpue_min_47'),
                       calc_cpue_texte_47(catch_df, set_df))),
            **construire_survie_requins(catch_df),
            **construire_devenir(catch_df),
            'engin_rows':    engin_rows,
            'orins_rows':    orins_rows,
            'hamecons_rows': hamecons_rows,
            'appats_rows':   appats_rows,
            'filage_rows':   filage_rows,
            'virage_rows':   virage_rows,
            'tdr_rows':      tdr_rows,
            'captures_rows': captures_rows,
            'tableau_cpue_complet': MARQUEUR_CPUE,
            # Toutes les figures en pleine largeur de page (17,8 cm). La figure 15 était
            # réduite à 10 cm du temps où c'était un camembert ; c'est aujourd'hui un
            # diagramme en barres, illisible à cette taille.
            **{tag: make_img(chemin) for tag, chemin in figs.items()},
        }

        tpl.render(contexte)
        tpl.save(output_path)

        # ── Post-rendu : tableau CPUE + mise a jour automatique des tables ────
        doc_final = Document(output_path)
        _inserer_tableau_cpue(doc_final, catch_df, set_df)

        settings = doc_final.settings.element
        for old in settings.findall(qn('w:updateFields')):
            settings.remove(old)
        uf = OxmlElement('w:updateFields')
        uf.set(qn('w:val'), 'true')
        settings.append(uf)

        doc_final.save(output_path)

    nb_avert = len(_avertissements_generation)
    print(f'  Rapport genere : {os.path.basename(output_path)} ({nb_avert} avertissement(s))')
    return nb_avert
