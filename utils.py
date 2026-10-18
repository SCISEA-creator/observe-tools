#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Utilitaires partagés entre toutes les phases.
IRD OB7 - SCISEA

Ce module contient les fonctions communes utilisées par Phase 1, Phase 2 et
Phase 3 pour éviter la duplication de code.
"""

# Version unique de l'application — importée par app.py et les phases
VERSION = "3.6.5"

# Police système Windows 10/11 — plus moderne qu'Arial, garantie présente
POLICE = 'Segoe UI'

# Palette de couleurs commune à toute l'interface (app.py, Phase 1, Phase 2)
# 'erreur' et 'avertissement' sont utilisés par Phase 1 uniquement
COULEURS = {
    'fond':          '#F8FAFC',
    'en_tete':       '#1E1B4B',   # Indigo-950 (modernisé 2025)
    'accent':        '#4F46E5',   # Indigo-600 (modernisé 2025)
    'erreur':        '#DC2626',
    'avertissement': '#D97706',
    'ok':            '#059669',   # Emerald-600 (modernisé 2025)
    'blanc':         '#FFFFFF',
    'texte':         '#1E293B',
}

import os
import zipfile
import tempfile

# Correspondance couleur normale → couleur survol (légèrement plus sombre)
_HOVER_MAP = {
    '#4F46E5': '#4338CA',   # Indigo-600 → Indigo-700
    '#059669': '#047857',   # Emerald-600 → Emerald-700
    '#1E1B4B': '#16124A',   # Indigo-950 → plus sombre
    '#7C3AED': '#6D28D9',   # Violet-600 → Violet-700
    '#0EA5E9': '#0284C7',   # Sky-500 → Sky-600
    '#0F766E': '#0D6460',   # Teal-700 → Teal-800
    '#D97706': '#B45309',   # Amber-600 → Amber-700
    '#DC2626': '#B91C1C',   # erreur
    # Anciennes valeurs — rétrocompatibilité
    '#2563EB': '#1D4ED8',
    '#16A34A': '#15803D',
    '#1F4E79': '#163D61',
    '#E07B00': '#B45309',
}


def appliquer_hover(btn):
    """Ajoute un effet visuel de survol (clair → foncé) à un tk.Button."""
    normale = btn.cget('bg')
    survol = _HOVER_MAP.get(normale, '#163D61')
    btn.bind('<Enter>', lambda e: btn.config(bg=survol))
    btn.bind('<Leave>', lambda e: btn.config(bg=normale))


def ajouter_titre_bandeau(frame_parent, bg_couleur, texte, hauteur_logo=42):
    """
    Place logo Observe-Tools + texte titre à gauche de frame_parent (bandeau indigo).
    Retourne le PhotoImage — le CALLER doit le stocker dans self._logo_titre (évite le GC).
    """
    import tkinter as tk
    from PIL import Image, ImageTk

    LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'assets', 'logos',
                             'observe-tools_sans-texte_couleur_transparent.png')

    cadre = tk.Frame(frame_parent, bg=bg_couleur)
    cadre.pack(side='left', padx=16, pady=(7, 3))

    photo = None
    try:
        img = Image.open(LOGO_PATH).convert('RGBA')
        # Rogner les pixels transparents en bordure (espace mort dans le fichier PNG)
        bbox = img.getbbox()
        if bbox:
            img = img.crop(bbox)
        # Redimensionner en préservant les proportions (comme ajouter_logos_bandeau)
        ratio = hauteur_logo / img.height
        img = img.resize((max(1, int(img.width * ratio)), hauteur_logo), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        tk.Label(cadre, image=photo, bg=bg_couleur).pack(side='left', padx=(0, 8))
    except Exception:
        pass

    tk.Label(cadre, text=texte, bg=bg_couleur, fg='white',
             font=(POLICE, 13, 'bold')).pack(side='left')

    return photo  # stocker dans self._logo_titre pour éviter le garbage collection


def ajouter_logos_bandeau(frame_parent, bg_couleur, hauteur=34):
    """
    Place les logos IRD, OB7, SCISEA à droite de frame_parent (bandeau bleu).
    Retourne la liste des PhotoImage — le CALLER doit la stocker (évite le GC).
    """
    import tkinter as tk
    from PIL import Image, ImageTk

    DOSSIER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'logos')
    fichiers = ['logo_IRD_nobg.png', 'logoOb7_nobg.png', 'logo_SCISEA_nobg.png']

    images = []
    cadre = tk.Frame(frame_parent, bg=bg_couleur)
    cadre.pack(side='right', padx=12, pady=9)  # 9px top/bottom → centré dans 52px

    for nom in fichiers:
        chemin = os.path.join(DOSSIER, nom)
        try:
            img = Image.open(chemin).convert('RGBA')
            ratio = hauteur / img.height
            img = img.resize((max(1, int(img.width * ratio)), hauteur), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            images.append(photo)
            tk.Label(cadre, image=photo, bg=bg_couleur).pack(side='left', padx=4)
        except Exception:
            pass

    return images  # stocker dans self._logos pour éviter le garbage collection

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE DES FICHIERS OBSERV (ZIP / XLSX)
# ─────────────────────────────────────────────────────────────────────────────

def lire_zip(chemin_zip):
    """
    Ouvre un fichier .zip exporté par ObServe et retourne un dictionnaire
    de DataFrames (un par table trouvée).

    chemin_zip : chemin vers le fichier .zip
    retourne   : dict { 'nom_table': DataFrame, ... }
    lève       : ValueError avec message lisible en cas d'erreur
    """
    tables = {}

    # Vérification rapide du ZIP avant extraction — donne des messages d'erreur clairs
    try:
        zf_test = zipfile.ZipFile(chemin_zip, 'r')
        zf_test.close()
    except zipfile.BadZipFile:
        raise ValueError("Le fichier n'est pas un ZIP valide (corrompu ou tronqué).")
    except PermissionError:
        raise ValueError("Impossible d'ouvrir le fichier : il est peut-être ouvert dans un autre logiciel.")
    except FileNotFoundError:
        raise ValueError(f"Fichier introuvable : {chemin_zip}")

    # Extraction dans un dossier temporaire puis lecture des fichiers Excel
    try:
        with tempfile.TemporaryDirectory() as dossier_temp:
            with zipfile.ZipFile(chemin_zip, 'r') as zf:
                zf.extractall(dossier_temp)

            for racine, sous_dossiers, fichiers in os.walk(dossier_temp):
                for nom_fichier in sorted(fichiers):

                    # Ignorer les fichiers cachés macOS
                    if nom_fichier.startswith('._') or '__MACOSX' in racine:
                        continue

                    if not nom_fichier.endswith(('.xlsx', '.xls')):
                        continue

                    chemin_fichier = os.path.join(racine, nom_fichier)

                    # Nom de table = "N_nomtable" (ex: "1_trip", "8_catch")
                    parties = nom_fichier.replace('.xlsx', '').replace('.xls', '').split('_')
                    if len(parties) >= 2:
                        nom_table = parties[0] + '_' + parties[1]
                    else:
                        nom_table = parties[0]

                    try:
                        df = pd.read_excel(chemin_fichier, engine='openpyxl')
                        tables[nom_table] = df
                    except Exception as e:
                        raise ValueError(f"Impossible de lire '{nom_fichier}' : {e}")

            # Aucun Excel (ancien format ObServe 9) : zip de CSV ObServe 10, à la racine
            # ou dans un sous-dossier → même lecture et mêmes contrôles qu'une sélection de CSV
            if not tables:
                from import_csv import lire_fichiers_csv, table_du_fichier
                csv = [os.path.join(r, f) for r, _, fs in os.walk(dossier_temp) for f in fs
                       if f.lower().endswith('.csv') and not f.startswith('._')
                       and '__MACOSX' not in r and table_du_fichier(f)]
                if csv:
                    tables = lire_fichiers_csv(csv)

    except ValueError:
        raise   # laisser remonter nos propres messages d'erreur
    except Exception as e:
        raise ValueError(f"Erreur lors de la lecture du ZIP : {e}")

    return tables


def lire_fichier(chemin):
    """
    Ouvre un fichier (zip ou xlsx) et retourne les tables sous forme de dict.

    chemin   : fichier .zip / .xlsx / .xls, liste de fichiers CSV ObServe, ou dossier de CSV
    retourne : dict { 'nom_table': DataFrame, ... } ou None si format non reconnu
    lève     : ValueError avec message lisible en cas d'erreur
    """
    if isinstance(chemin, (list, tuple)):
        from import_csv import lire_fichiers_csv
        return lire_fichiers_csv(list(chemin))

    if os.path.isdir(chemin):
        from import_csv import lire_dossier_csv
        return lire_dossier_csv(chemin)

    if chemin.endswith('.zip'):
        return lire_zip(chemin)

    elif chemin.endswith(('.xlsx', '.xls')):
        nom_table = os.path.basename(chemin).replace('.xlsx', '').replace('.xls', '')
        parties = nom_table.split('_')
        if len(parties) >= 2:
            nom_table = parties[0] + '_' + parties[1]
        try:
            df = pd.read_excel(chemin, engine='openpyxl')
        except PermissionError:
            raise ValueError("Impossible d'ouvrir le fichier : il est peut-être ouvert dans Excel.")
        except Exception as e:
            raise ValueError(f"Impossible de lire le fichier Excel : {e}")
        return {nom_table: df}

    else:
        return None
