#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module de gestion du référentiel ObServe
IRD OB7 - SCISEA

Gère le téléchargement, le stockage local et le chargement du référentiel
ObServe (liste des espèces, codes statuts, types d'engins, etc.).

Le référentiel est téléchargé depuis l'API publique ObServe, filtré pour
ne garder que les 3 groupes utiles à la palangre, et sauvegardé en JSON local.
"""

import os
import json
import socket
import requests  # gère SSL correctement via certifi (urllib échoue sur certaines machines Windows)
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

# URL d'authentification. Les identifiants (lecture seule, fournis par l'IRD) ne sont
# PAS dans le code : ils sont lus dans data/identifiants_ird.json (exclu de GitHub,
# inclus dans le ZIP de distribution).
URL_TOKEN = "https://observe.ob7.ird.fr/observeweb/api/public/init/open"
PARAMS_TOKEN = {'config.modelVersion': '9.5', 'config.databaseName': 'main'}

CHEMIN_IDENTIFIANTS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'data', 'identifiants_ird.json'
)

# Mis à True par telecharger_referentiel() quand l'échec vient des identifiants
# (absents ou refusés par le serveur) : l'app propose alors d'en saisir de nouveaux.
ECHEC_IDENTIFIANTS = False

# URL du référentiel complet — le token est injecté dynamiquement
URL_REFERENTIEL = (
    "https://observe.ob7.ird.fr/observeweb/api/public/referential"
    "?authenticationToken={token}"
    "&config.loadReferential=true"
    "&config.recursive=true"
    "&config.prettyPrint=false"
    "&config.serializeNulls=true"
)

# Les 3 préfixes utiles pour la palangre (on ignore senne et logbook palangre)
PREFIXES_UTILES = (
    'fr.ird.observe.entities.referential.common',
    'fr.ird.observe.entities.referential.ll.common',
    'fr.ird.observe.entities.referential.ll.observation',
)

# Chemin du fichier JSON local (dans le dossier data/ qui est gitignore)
CHEMIN_REFERENTIEL = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    'data',
    'referentiel_observe.json'
)


# ─────────────────────────────────────────────────────────────────────────────
# FONCTIONS RÉSEAU
# ─────────────────────────────────────────────────────────────────────────────

def verifier_internet():
    """
    Vérifie si le serveur IRD ObServe est joignable (port 443).
    On teste directement le serveur cible, pas juste Google :
    si l'IRD est bloqué par le firewall de l'utilisateur, on retourne False
    et l'app n'essaiera pas de télécharger le référentiel.
    """
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(5)  # timeout local à cette socket uniquement — pas d'effet global sur le processus
        s.connect(("observe.ob7.ird.fr", 443))
        s.close()
        return True
    except socket.error:
        return False


def charger_identifiants():
    """Retourne (login, mot_de_passe) depuis data/identifiants_ird.json, ou (None, None)."""
    try:
        with open(CHEMIN_IDENTIFIANTS, encoding='utf-8') as f:
            d = json.load(f)
        login, mdp = str(d.get('login', '')).strip(), str(d.get('password', ''))
        return (login, mdp) if login and mdp else (None, None)
    except Exception:
        return None, None


def sauvegarder_identifiants(login, mot_de_passe):
    """Enregistre les identifiants IRD dans le dossier de l'app (data/identifiants_ird.json)."""
    os.makedirs(os.path.dirname(CHEMIN_IDENTIFIANTS), exist_ok=True)
    with open(CHEMIN_IDENTIFIANTS, 'w', encoding='utf-8') as f:
        json.dump({'login': login.strip(), 'password': mot_de_passe}, f, ensure_ascii=False, indent=2)


def obtenir_token():
    """
    Demande un token d'authentification à l'API ObServe.
    Retourne (token, None, False) en cas de succès,
             (None, message_erreur, identifiants_en_cause) en cas d'échec.
    identifiants_en_cause = True si les identifiants sont absents, ou si le serveur a
    répondu mais refusé la connexion (le réseau, lui, fonctionne).
    """
    login, mdp = charger_identifiants()
    if not login:
        return None, "Identifiants du serveur IRD absents", True
    try:
        reponse = requests.get(URL_TOKEN, timeout=10,
                               params={**PARAMS_TOKEN, 'config.login': login, 'config.password': mdp})
    except Exception as e:
        # Problème réseau : ce ne sont pas les identifiants
        return None, f"{type(e).__name__} (serveur observe.ob7.ird.fr injoignable)", False
    if reponse.status_code != 200:
        return None, f"Connexion refusée par le serveur IRD (code {reponse.status_code})", True
    try:
        token = reponse.json().get('authenticationToken')
    except ValueError:
        token = None
    if not token:
        return None, "Réponse inattendue du serveur IRD (pas de jeton d'accès)", True
    return token, None, False


# ─────────────────────────────────────────────────────────────────────────────
# TÉLÉCHARGEMENT ET SAUVEGARDE
# ─────────────────────────────────────────────────────────────────────────────

def telecharger_referentiel(callback_progression=None):
    """
    Télécharge le référentiel depuis l'API ObServe, filtre les 3 groupes utiles
    et sauvegarde le résultat dans un fichier JSON local.

    callback_progression : fonction(message: str) appelée pour indiquer l'avancement.
                           Doit être thread-safe (utiliser after() si tkinter).
    Retourne True si succès, False sinon.
    """
    global ECHEC_IDENTIFIANTS
    ECHEC_IDENTIFIANTS = False

    def prog(msg):
        if callback_progression:
            callback_progression(msg)

    try:
        # Étape 1 : obtenir le token d'accès
        prog("Connexion au serveur ObServe...")
        token, err_token, identifiants_en_cause = obtenir_token()
        if not token:
            ECHEC_IDENTIFIANTS = identifiants_en_cause
            prog(f"Erreur : {err_token}")
            return False

        # Étape 2 : télécharger le référentiel complet (~8 Mo)
        prog("Téléchargement du référentiel (quelques secondes)...")
        url = URL_REFERENTIEL.format(token=token)
        reponse = requests.get(url, timeout=60)
        reponse.raise_for_status()
        contenu_brut = reponse.json()

        # Étape 3 : filtrer uniquement les 3 groupes utiles
        prog("Filtrage des données palangre...")
        content = contenu_brut.get('content', {})
        referentiels_filtres = {
            cle: valeur
            for cle, valeur in content.items()
            if any(cle.startswith(p) for p in PREFIXES_UTILES)
        }

        # Étape 4 : sauvegarder avec métadonnées
        prog("Sauvegarde locale...")
        os.makedirs(os.path.dirname(CHEMIN_REFERENTIEL), exist_ok=True)

        sauvegarde = {
            'date_telechargement': datetime.now().strftime('%d/%m/%Y %H:%M'),
            'version_observe':     '9.5',
            'nb_tables':           len(referentiels_filtres),
            'referentiels':        referentiels_filtres,
        }
        with open(CHEMIN_REFERENTIEL, 'w', encoding='utf-8') as f:
            json.dump(sauvegarde, f, ensure_ascii=False)

        nb = len(referentiels_filtres)
        prog(f"Référentiel mis à jour ({nb} tables chargées).")
        return True

    except Exception as erreur:
        prog(f"Erreur lors du téléchargement : {erreur}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE DU RÉFÉRENTIEL LOCAL
# ─────────────────────────────────────────────────────────────────────────────

def charger_referentiel_local():
    """
    Charge le référentiel depuis le fichier JSON local.
    Retourne le dict complet ou None si le fichier n'existe pas ou est corrompu.
    """
    if not os.path.exists(CHEMIN_REFERENTIEL):
        return None
    try:
        with open(CHEMIN_REFERENTIEL, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def date_referentiel_local():
    """
    Retourne la date du dernier téléchargement (str 'DD/MM/YYYY HH:MM')
    ou None si le référentiel local n'existe pas.
    """
    data = charger_referentiel_local()
    if data:
        return data.get('date_telechargement')
    return None


def get_codes(nom_table):
    """
    Retourne l'ensemble des codes valides d'une table du référentiel local.
    Exemple : get_codes('Species') → {'YFT', 'BET', 'ALB', ...}
    Retourne un ensemble vide si la table est introuvable.
    """
    data = charger_referentiel_local()
    if not data:
        return set()

    referentiels = data.get('referentiels', {})

    # Cherche la table dans les 3 préfixes
    for prefixe in PREFIXES_UTILES:
        cle_complete = f"{prefixe}.{nom_table}"
        if cle_complete in referentiels:
            entrees = referentiels[cle_complete]
            return {str(e.get('code', '')).strip() for e in entrees if e.get('code')}

    return set()
