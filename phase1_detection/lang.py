#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module de gestion de la langue de l'application.
IRD OB7 - SCISEA

Fournit la fonction T(clé) qui retourne le texte dans la langue active.
La langue est sauvegardée dans config.json et chargée au démarrage.

Langues disponibles : 'fr' (français), 'en' (anglais)
Label référentiel   : label1 = anglais, label2 = français
"""

import os
import json

# ─────────────────────────────────────────────────────────────────────────────
# CHEMINS
# ─────────────────────────────────────────────────────────────────────────────

_DOSSIER        = os.path.dirname(os.path.abspath(__file__))
CHEMIN_CONFIG   = os.path.join(_DOSSIER, 'config.json')
CHEMIN_TRANS    = os.path.join(_DOSSIER, 'translations')

# ─────────────────────────────────────────────────────────────────────────────
# ÉTAT INTERNE
# ─────────────────────────────────────────────────────────────────────────────

_LANGUE = 'fr'
_TEXTES = {}       # dict chargé depuis le fichier JSON de la langue active


# ─────────────────────────────────────────────────────────────────────────────
# CHARGEMENT
# ─────────────────────────────────────────────────────────────────────────────

def charger_langue(code='fr'):
    """Charge les textes depuis le fichier JSON de la langue donnée."""
    global _LANGUE, _TEXTES
    _LANGUE = code
    chemin = os.path.join(CHEMIN_TRANS, f'{code}.json')
    if os.path.exists(chemin):
        with open(chemin, encoding='utf-8') as f:
            _TEXTES = json.load(f)
    else:
        _TEXTES = {}


def langue_depuis_config():
    """Lit la langue sauvegardée dans config.json (défaut : 'fr')."""
    if os.path.exists(CHEMIN_CONFIG):
        try:
            with open(CHEMIN_CONFIG, encoding='utf-8') as f:
                return json.load(f).get('langue', 'fr')
        except Exception:
            pass
    return 'fr'


def sauvegarder_langue(code):
    """Sauvegarde le choix de langue dans config.json."""
    config = {}
    if os.path.exists(CHEMIN_CONFIG):
        try:
            with open(CHEMIN_CONFIG, encoding='utf-8') as f:
                config = json.load(f)
        except Exception:
            pass
    config['langue'] = code
    with open(CHEMIN_CONFIG, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def langue():
    """Retourne le code de la langue active ('fr' ou 'en')."""
    return _LANGUE


# ─────────────────────────────────────────────────────────────────────────────
# TRADUCTION
# ─────────────────────────────────────────────────────────────────────────────

def T(cle, **kwargs):
    """
    Retourne le texte traduit pour une clé donnée.
    Les clés peuvent être imbriquées avec des points : 'boutons.verifier'
    Les paramètres dynamiques sont passés en kwargs : T('msg', ref='Calée 1')

    Si la clé n'existe pas, retourne la clé elle-même (ne plante jamais).
    """
    valeur = _TEXTES
    for partie in cle.split('.'):
        if isinstance(valeur, dict):
            valeur = valeur.get(partie, None)
        else:
            valeur = None
            break

    if valeur is None:
        return cle  # clé introuvable : on retourne la clé brute

    if isinstance(valeur, str) and kwargs:
        try:
            return valeur.format(**kwargs)
        except (KeyError, ValueError):
            return valeur

    return valeur if isinstance(valeur, str) else cle


def label_ref(entree):
    """
    Retourne le libellé d'une entrée du référentiel dans la langue active.
    label1 = anglais, label2 = français.
    """
    if _LANGUE == 'en':
        return entree.get('label1') or entree.get('label2') or entree.get('code', '')
    return entree.get('label2') or entree.get('label1') or entree.get('code', '')
