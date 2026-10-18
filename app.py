#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Application principale unifiée
IRD OB7 - SCISEA

Trois onglets :
  - Accueil         : import du fichier ObServe, référentiel, sélecteur de langue
  - Correction      : Phase 1 — vérification et correction des données
  - Figures         : Phase 2 — génération des figures et tableaux Excel

Lancer : python app.py
         ou double-cliquer sur LANCER.bat
"""

import os
import sys
import zipfile
import importlib.util
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from utils import VERSION, COULEURS, POLICE, appliquer_hover, ajouter_logos_bandeau, ajouter_titre_bandeau   # constantes communes — source unique dans utils.py

# ─────────────────────────────────────────────────────────────────────────────
# CHEMINS DE BASE
# os.path.abspath(__file__) donne le chemin complet de app.py, même si
# l'utilisateur a lancé le script depuis un autre dossier.
# ─────────────────────────────────────────────────────────────────────────────

_BASE       = os.path.dirname(os.path.abspath(__file__))
_PHASE1_DIR = os.path.join(_BASE, 'phase1_detection')
_PHASE2_DIR = os.path.join(_BASE, 'phase2_graphiques')

# ─────────────────────────────────────────────────────────────────────────────
# AJOUT DES DOSSIERS DES PHASES DANS LE CHEMIN PYTHON
# Nécessaire pour que chaque phase trouve ses propres modules locaux.
# ─────────────────────────────────────────────────────────────────────────────

for _dossier in (_PHASE1_DIR, _PHASE2_DIR):
    if _dossier not in sys.path:
        sys.path.insert(0, _dossier)

# ─────────────────────────────────────────────────────────────────────────────
# CHARGEMENT DE LA LANGUE ET DU RÉFÉRENTIEL (modules de Phase 1)
# Doivent être chargés avant les phases elles-mêmes.
# ─────────────────────────────────────────────────────────────────────────────

import lang
lang.charger_langue(lang.langue_depuis_config())

import referentiel   # accessible car _PHASE1_DIR est dans sys.path

# ─────────────────────────────────────────────────────────────────────────────
# IMPORT DES DEUX PHASES VIA LEURS CHEMINS EXACTS (importlib)
# Les deux phases ont un fichier main.py. On leur donne des noms uniques pour
# éviter les conflits dans sys.modules.
# ─────────────────────────────────────────────────────────────────────────────

def _charger_module(nom_unique, chemin_fichier):
    """Charge un module Python depuis son chemin exact avec un nom unique."""
    spec = importlib.util.spec_from_file_location(nom_unique, chemin_fichier)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules[nom_unique] = mod
    spec.loader.exec_module(mod)
    return mod

_phase1 = _charger_module('phase1_main', os.path.join(_PHASE1_DIR, 'main.py'))
_phase2 = _charger_module('phase2_main', os.path.join(_PHASE2_DIR, 'main.py'))
try:
    _phase3 = _charger_module('phase3_main', os.path.join(os.path.dirname(__file__), 'phase3_rapport', 'main.py'))
    _PHASE3_OK = True
except Exception as _phase3_err:
    # Si docxtpl ou un autre module Phase 3 est absent (installation incomplète),
    # l'app démarre quand même — l'onglet Rapport Word est désactivé avec un message clair.
    # En Python 3, les variables "except ... as x" sont supprimées à la fin du bloc —
    # on sauvegarde le message avant qu'il disparaisse.
    _phase3_err_msg = str(_phase3_err)
    _phase3 = None
    _PHASE3_OK = False

# Classes des interfaces et fenêtre référentiel (définie dans Phase 1)
AppObserveTools    = _phase1.AppObserveTools
AppPhase2          = _phase2.AppPhase2
AppPhase3          = _phase3.AppPhase3 if _PHASE3_OK else None
FenetreReferentiel = _phase1.FenetreReferentiel

# Zones de pêche — source unique dans generer_rapport.py, importée ici via _phase3.
# Les clés correspondent exactement aux valeurs label2 de la table FpaZone du référentiel.
# Les zones sans entrée dans le dict utilisent le format par défaut "ZEE de {label}".
ZONES_PECHE = _phase3.ZONES_PECHE if _PHASE3_OK else {}


# ─────────────────────────────────────────────────────────────────────────────
# PAGE D'ACCUEIL
# Centralise l'import du fichier, le référentiel et le sélecteur de langue.
# Quand un fichier est chargé, les tables sont transmises aux deux phases
# via le callback on_fichier_charge.
#
# Pour étendre à une Phase 3 (rapport Word) :
#   → ajouter app3._recevoir_tables(tables, nom_fichier) dans le callback
#      on_fichier_charge dans main() ci-dessous.
# ─────────────────────────────────────────────────────────────────────────────

class PageAccueil(tk.Frame):
    """
    Page d'accueil de l'application unifiée.
    Gère l'import du fichier ObServe, le référentiel et la langue.
    """

    def __init__(self, parent, on_fichier_charge):
        """
        on_fichier_charge : callback(tables, nom_fichier) appelé après un import réussi.
                            Transmet les tables chargées à toutes les phases.
        """
        super().__init__(parent, bg=COULEURS['fond'])
        self._on_fichier_charge = on_fichier_charge
        self._tables_chargees      = None  # stockées pour la création du dossier
        self._zones_selectionnees  = []   # liste des zones choisies dans le popup
        self._lbl_zones_choisies   = None  # label affiché à côté du bouton Choisir
        self._construire_interface()
        # Vérification du référentiel au démarrage (délai pour laisser l'UI s'afficher)
        self.after(1500, self._lancer_verification_referentiel)

    def _construire_interface(self):

        # ── En-tête ───────────────────────────────────────────────────────────
        en_tete = tk.Frame(self, bg=COULEURS['en_tete'], height=52)
        en_tete.pack(fill='x')
        en_tete.pack_propagate(False)

        self._logo_titre = ajouter_titre_bandeau(
            en_tete, COULEURS['en_tete'], "Observe-Tools — Accueil et import données ObServe"
        )

        self._logos = ajouter_logos_bandeau(en_tete, COULEURS['en_tete'])

        # Sélecteur de langue (FR / EN)
        self._var_langue = tk.StringVar(value=lang.langue().upper())
        menu_langue = ttk.Combobox(
            en_tete, textvariable=self._var_langue,
            values=['FR', 'EN'], state='readonly', width=4, font=(POLICE, 9)
        )
        menu_langue.pack(side='right', padx=(0, 8))
        menu_langue.bind('<<ComboboxSelected>>', self._changer_langue)
        tk.Label(
            en_tete, text="Langue :",
            bg=COULEURS['en_tete'], fg='#C7D2FE', font=(POLICE, 9)
        ).pack(side='right')

        # ── Zone centrale ─────────────────────────────────────────────────────
        centre = tk.Frame(self, bg=COULEURS['fond'])
        centre.pack(fill='both', expand=True, padx=60, pady=20)

        # ── Ligne supérieure : bouton import (gauche) + méthode (droite) ────────
        top_row = tk.Frame(centre, bg=COULEURS['fond'])
        top_row.pack(fill='x', pady=(20, 8))

        # Colonne gauche — méthode de travail
        cadre_methode = tk.Frame(top_row, bg=COULEURS['fond'])
        cadre_methode.pack(side='left', anchor='n', padx=(0, 40))

        # Colonne droite — bouton import centré dans l'espace restant
        cadre_import = tk.Frame(top_row, bg=COULEURS['fond'])
        cadre_import.pack(side='left', fill='both', expand=True)

        # Un seul bouton : les 9 CSV d'ObServe 10 (sélection multiple) ou un .zip
        btn_import = ttk.Button(
            cadre_import,
            text="  📂  Importer les données ObServe  ",
            command=self._importer_fichier,
            style='BigAccent.TButton',
            cursor='hand2',
        )
        btn_import.pack(anchor='center', pady=(20, 0))

        tk.Label(
            cadre_import,
            text="Sélectionnez les 9 fichiers CSV de la marée (Ctrl+A ou Ctrl+clic), ou un seul fichier .zip",
            bg=COULEURS['fond'], fg='#777777', font=(POLICE, 9)
        ).pack(pady=(6, 0))

        tk.Label(
            cadre_methode, text="Méthode de travail",
            bg=COULEURS['fond'], font=(POLICE, 11, 'bold'), fg=COULEURS['texte']
        ).pack(anchor='w')

        instructions = (
            "  1. Exportez les 9 tables de la marée en CSV depuis ObServe, puis importez-les\n"
            "     ci-dessus (sélectionnez les 9 fichiers, ou un .zip qui les contient)\n"
            "  2. Sélectionnez la zone de pêche dans la section ci-dessous\n"
            "  3. Vérifiez les données via l'onglet \"Correction des données\"\n"
            "  4. Corrigez les erreurs dans ObServe si nécessaire\n"
            "  5. Réimportez les données corrigées ici\n"
            "  6. Exportez le rapport de vérification via l'onglet \"Correction des données\"\n"
            "  7. Générez le rapport de marée via l'onglet \"Rapport\""
        )
        tk.Label(
            cadre_methode, text=instructions,
            bg=COULEURS['fond'], fg='#374151',
            font=(POLICE, 10), justify='left'
        ).pack(anchor='w', pady=(6, 0))

        # Séparateur
        ttk.Separator(centre, orient='horizontal').pack(fill='x', pady=16)

        # ── Carte infos marée — masquée jusqu'au 1er import ─────────────────
        carte_maree = tk.Frame(centre, bg=COULEURS['blanc'], relief='groove', bd=1)
        carte_maree.pack(fill='x', pady=(0, 8))
        inner_maree = tk.Frame(carte_maree, bg=COULEURS['blanc'])
        inner_maree.pack(fill='both', padx=12, pady=8)

        self._lbl_aucun = tk.Label(
            inner_maree,
            text="Aucun fichier chargé. Cliquez sur le bouton ci-dessus pour commencer.",
            bg=COULEURS['blanc'], fg='#888888', font=(POLICE, 10, 'italic')
        )
        self._lbl_aucun.pack(pady=4)

        self._frame_maree = tk.Frame(inner_maree, bg=COULEURS['blanc'])
        # (affiché après le 1er import, via _afficher_infos_maree)

        # ── Carte référentiel ─────────────────────────────────────────────────
        carte_ref_outer = tk.Frame(centre, bg=COULEURS['blanc'], relief='groove', bd=1)
        carte_ref_outer.pack(fill='x', pady=(0, 8))
        cadre_ref = tk.Frame(carte_ref_outer, bg=COULEURS['blanc'])
        cadre_ref.pack(fill='both', padx=12, pady=8)

        tk.Label(
            cadre_ref, text="Référentiel ObServe",
            bg=COULEURS['blanc'], font=(POLICE, 11, 'bold')
        ).pack(anchor='w')

        ligne_ref = tk.Frame(cadre_ref, bg=COULEURS['blanc'])
        ligne_ref.pack(fill='x', pady=6)

        self._lbl_date_ref = tk.Label(
            ligne_ref, text="...",
            bg=COULEURS['blanc'], font=(POLICE, 9), fg='#555555'
        )
        self._lbl_date_ref.pack(side='left')

        tk.Button(
            ligne_ref, text="Voir le référentiel",
            command=self._ouvrir_referentiel,
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9), relief='groove', padx=8,
            cursor='hand2'
        ).pack(side='left', padx=12)

        tk.Button(
            ligne_ref, text="Mettre à jour le référentiel",
            command=self._telecharger_avec_progression,
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9), relief='groove', padx=8,
            cursor='hand2'
        ).pack(side='left', padx=4)

        # ── Barre de statut ───────────────────────────────────────────────────
        barre_bas = tk.Frame(self, bg=COULEURS['en_tete'], height=26)
        barre_bas.pack(fill='x', side='bottom')
        barre_bas.pack_propagate(False)

        self._var_statut = tk.StringVar(value="Prêt. Importez un fichier ObServe pour commencer.")
        self._lbl_dot = tk.Label(barre_bas, text='●', bg=COULEURS['en_tete'], fg='#94A3B8', font=(POLICE, 9))
        self._lbl_dot.pack(side='left', padx=(10, 2))
        tk.Label(
            barre_bas, textvariable=self._var_statut,
            bg=COULEURS['en_tete'], fg='white', font=(POLICE, 9), anchor='w'
        ).pack(side='left', padx=(0, 10))

        def _maj_dot(*args):
            msg = self._var_statut.get()
            m = msg.lower()
            if '✗' in msg or 'erreur' in m or 'impossible' in m:
                fg = '#F87171'
            elif '⚠' in msg or 'avertissement' in m:
                fg = '#FBB76A'
            elif 'cours' in m or 'chargement' in m:
                fg = '#93C5FD'
            elif '✓' in msg or 'chargé' in m or 'créé' in m or 'corrigé' in m or 'prêt' in m:
                fg = '#4ADE80'
            else:
                fg = '#94A3B8'
            self._lbl_dot.config(fg=fg)
        self._var_statut.trace_add('write', _maj_dot)

    # ── Import du fichier ─────────────────────────────────────────────────────

    def _importer_fichier(self):
        """
        Ouvre la boîte de dialogue (sélection multiple) : les 9 CSV d'une marée
        exportés par ObServe 10, ou un seul .zip (CSV ou ancien export Excel ObServe 9).
        """
        chemins = filedialog.askopenfilenames(
            title="Choisir les données ObServe : les 9 fichiers CSV de la marée, ou un fichier .zip",
            filetypes=[
                ('Exports ObServe (CSV, ZIP)', ('*.csv', '*.zip', '*.xlsx', '*.xls')),
                ('CSV', '*.csv'),
                ('ZIP', '*.zip'),
                ('Excel', '*.xlsx *.xls'),
                ('Tous les fichiers', '*.*'),
            ]
        )
        if not chemins:
            return
        self._charger(list(chemins))

    def _charger(self, chemins):
        """
        Charge un export ObServe et notifie toutes les phases (callback on_fichier_charge).
        chemins : liste de fichiers CSV d'une marée, ou un seul .zip / .xlsx
                  (une chaîne seule est aussi acceptée).
        Chaque erreur affiche ce que l'utilisateur a fait de faux et comment le corriger.
        """
        from import_csv import FICHIERS_CSV, NOMS_TABLES, verifier_horodatages
        if isinstance(chemins, str):
            chemins = [chemins]
        extensions = [os.path.splitext(c)[1].lower() for c in chemins]
        nb_csv = extensions.count('.csv')

        # ── 1. La sélection doit être cohérente ──────────────────────────────
        if nb_csv and nb_csv != len(chemins):
            messagebox.showerror(
                "Sélection incorrecte",
                "Vous avez sélectionné à la fois des fichiers CSV et un autre type de fichier "
                "(.zip ou Excel).\n\n"
                "Sélectionnez soit les fichiers CSV d'une seule marée, soit un seul fichier .zip."
            )
            return
        if not nb_csv and len(chemins) > 1:
            messagebox.showerror(
                "Sélection incorrecte",
                f"Vous avez sélectionné {len(chemins)} fichiers .zip ou Excel.\n\n"
                "Sélectionnez un seul fichier .zip (une seule marée à la fois)."
            )
            return
        if not nb_csv and extensions[0] not in ('.zip', '.xlsx', '.xls'):
            messagebox.showerror(
                "Format non reconnu",
                f"Le fichier « {os.path.basename(chemins[0])} » n'est pas un export ObServe.\n\n"
                "Sélectionnez les fichiers CSV exportés par ObServe, ou un fichier .zip."
            )
            return

        # ── 2. Lecture et erreurs bloquantes (fichier non reconnu, table en double,
        #       plusieurs marées) — mêmes contrôles pour les CSV seuls et les zip de CSV ─
        try:
            tables = _phase1.lire_fichier(chemins if nb_csv else chemins[0])
        except ValueError as e:
            messagebox.showerror("Import impossible", str(e))
            return
        except Exception as e:
            messagebox.showerror(
                "Erreur d'import",
                f"Une erreur inattendue s'est produite lors de la lecture :\n\n{e}"
            )
            return

        if tables is None:
            messagebox.showerror(
                "Format non reconnu",
                "Ce format de fichier n'est pas supporté.\n\n"
                "Sélectionnez les fichiers CSV exportés par ObServe, ou un fichier .zip."
            )
            return
        if not tables:
            messagebox.showerror(
                "Aucune table ObServe",
                f"« {os.path.basename(chemins[0])} » ne contient aucune table ObServe "
                "(ni fichiers CSV d'ObServe 10, ni tableaux Excel d'ObServe 9).\n\n"
                "Vérifiez que vous avez sélectionné le bon fichier."
            )
            return

        # ── 3. Les tables doivent avoir été exportées au même moment (question,
        #       non bloquant : une table a pu être réexportée volontairement) ──
        if nb_csv:
            noms_csv = chemins
        elif extensions[0] == '.zip':
            with zipfile.ZipFile(chemins[0]) as zf:     # déjà lu sans erreur à l'étape 2
                noms_csv = [n for n in zf.namelist() if n.lower().endswith('.csv')]
        else:
            noms_csv = []
        avertissement = verifier_horodatages(noms_csv) if noms_csv else None
        if avertissement and not messagebox.askyesno(
                "Exports réalisés à des moments différents", avertissement, icon='warning'):
            return

        # ── 4. Tables CSV : chacune est exportée à la main → signaler un oubli ─
        if noms_csv:
            manquantes = [f"{NOMS_TABLES[t]}  (export-{mot}…)" for mot, t in FICHIERS_CSV.items()
                          if t not in tables]
            if manquantes:
                messagebox.showwarning(
                    "Tables manquantes",
                    "Ces tables manquent dans votre sélection :\n\n  - "
                    + "\n  - ".join(manquantes)
                    + "\n\nExportez-les depuis ObServe, puis réimportez en sélectionnant "
                    "les 9 fichiers de la marée.\nL'import continue avec les tables présentes."
                )

        # Chemin de référence pour retrouver _Rapport/ (Phase 1 et Phase 3) :
        # CSV → leur dossier (ex : MAREE/_Observe) ; zip/xlsx → le fichier lui-même
        if nb_csv:
            self._chemin_import = os.path.dirname(chemins[0])
            nom_fichier = f"{len(chemins)} fichiers CSV"
        else:
            self._chemin_import = chemins[0]
            nom_fichier = os.path.basename(chemins[0])

        # Afficher le résumé sur la page d'accueil
        self._afficher_infos_maree(tables, nom_fichier)

        # Transmettre les tables à toutes les phases
        self._on_fichier_charge(tables, nom_fichier)

        self._var_statut.set(f"Données chargées : {nom_fichier} ({len(tables)} tables)")

    def _afficher_infos_maree(self, tables, nom_fichier):
        """Affiche un résumé de la marée sur la page d'accueil après l'import."""
        self._tables_chargees = tables  # mémoriser pour la création du dossier

        # Masquer le message "aucun fichier" et vider les infos précédentes
        self._lbl_aucun.pack_forget()
        for w in self._frame_maree.winfo_children():
            w.destroy()

        trip_df   = tables.get('1_trip',   pd.DataFrame())
        catch_df  = tables.get('8_catch',  pd.DataFrame())
        set_df    = tables.get('2_set',    pd.DataFrame())
        sample_df = tables.get('9_sample', pd.DataFrame())

        def ligne(label, valeur, gras=False):
            """Crée une ligne label : valeur dans le cadre infos."""
            f = tk.Frame(self._frame_maree, bg=COULEURS['blanc'])
            f.pack(fill='x', pady=1)
            tk.Label(f, text=label + ' :', width=16, anchor='w',
                     bg=COULEURS['blanc'], font=(POLICE, 10, 'bold')).pack(side='left')
            font_val = (POLICE, 10, 'bold') if gras else (POLICE, 10)
            tk.Label(f, text=str(valeur), anchor='w',
                     bg=COULEURS['blanc'], font=font_val).pack(side='left')

        if not trip_df.empty:
            t = trip_df.iloc[0]
            ligne("Marée",       t.get('trip_id', '?'),                              gras=True)
            ligne("Navire",      t.get('vessel',  '?'))
            cap = f"{t.get('captain_firstname', '')} {t.get('captain_lastname', '')}".strip()
            ligne("Capitaine",   cap or '?')
            obs = f"{t.get('observer_firstname', '')} {t.get('observer_lastname', '')}".strip()
            ligne("Observateur", obs or '?')
        elif not catch_df.empty and 'trip_id' in catch_df.columns:
            ligne("Marée", catch_df['trip_id'].iloc[0], gras=True)

        # ── Zone(s) de pêche ──────────────────────────────────────────────────
        # Placée entre Observateur et Filages, alignée avec les autres lignes.
        # Bouton "Choisir..." → popup 53 zones FpaZone, Ctrl+clic multi-sélection.
        ligne_zee = tk.Frame(self._frame_maree, bg=COULEURS['blanc'])
        ligne_zee.pack(fill='x', pady=1)

        tk.Label(
            ligne_zee, text="Zone(s) de pêche :", width=16, anchor='w',
            bg=COULEURS['blanc'], font=(POLICE, 10, 'bold')
        ).pack(side='left')

        tk.Button(
            ligne_zee, text="Choisir...",
            command=self._ouvrir_popup_zones,
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 9), relief='groove', padx=8, pady=1,
            cursor='hand2'
        ).pack(side='left', padx=(0, 10))

        self._lbl_zones_choisies = tk.Label(
            ligne_zee,
            text="⚠ Sélectionnez au moins une zone",
            bg=COULEURS['blanc'], fg='#C55A11', font=(POLICE, 9, 'italic')
        )
        self._lbl_zones_choisies.pack(side='left')

        ligne("Filages",  len(set_df)    if not set_df.empty    else '?')
        ligne("Captures", len(catch_df)  if not catch_df.empty  else 0)
        ligne("Mesures",  len(sample_df) if not sample_df.empty else 0)

        tk.Label(
            self._frame_maree, text=f"Fichier : {nom_fichier}",
            bg=COULEURS['blanc'], font=(POLICE, 9), fg='#555555'
        ).pack(anchor='w', pady=(8, 0))

        # Boutons dossier marée
        cadre_boutons = tk.Frame(self._frame_maree, bg=COULEURS['blanc'])
        cadre_boutons.pack(anchor='w', pady=(12, 0))

        nom, date_fin_ok = self._construire_nom_dossier()

        btn_dossier = ttk.Button(
            cadre_boutons,
            text="  📁  Créer le dossier de la marée  ",
            command=self._creer_dossier_maree,
            style='Teal.TButton',
            cursor='hand2',
        )
        btn_dossier.pack(side='left')

        # Le bouton "Corriger" n'apparaît que si la date de fin est disponible
        # (utile quand le dossier a été créé en début de marée avec XXXXXXXX)
        if date_fin_ok:
            btn_corriger = ttk.Button(
                cadre_boutons,
                text="  ✏️  Corriger le nom du dossier  ",
                command=self._corriger_nom_dossier,
                style='Teal.TButton',
                cursor='hand2',
            )
            btn_corriger.pack(side='left', padx=(12, 0))

        # Avertissement si date de fin absente
        if not date_fin_ok:
            tk.Label(
                self._frame_maree,
                text="⚠ Date de fin absente — le dossier sera créé avec XXXXXXXX comme date de fin.",
                bg=COULEURS['blanc'], fg='#C55A11', font=(POLICE, 9, 'italic')
            ).pack(anchor='w', pady=(4, 0))

        self._frame_maree.pack(pady=4)

    def _ouvrir_popup_zones(self):
        """Ouvre une fenêtre popup pour sélectionner une ou plusieurs zones de pêche."""
        popup = tk.Toplevel(self)
        popup.title("Zones de pêche")
        popup.geometry("360x400")
        popup.resizable(True, True)
        popup.grab_set()  # bloque l'app principale pendant la sélection
        popup.configure(bg=COULEURS['fond'])

        tk.Label(
            popup, text="Sélectionnez la ou les zones de pêche :",
            bg=COULEURS['fond'], font=(POLICE, 10, 'bold')
        ).pack(padx=16, pady=(14, 2), anchor='w')

        tk.Label(
            popup, text="Ctrl+clic pour sélectionner plusieurs zones",
            bg=COULEURS['fond'], fg='#777777', font=(POLICE, 8, 'italic')
        ).pack(padx=16, anchor='w', pady=(0, 8))

        cadre_lb = tk.Frame(popup, bg=COULEURS['fond'])
        cadre_lb.pack(padx=16, fill='both', expand=True)

        scrollbar = ttk.Scrollbar(cadre_lb, orient='vertical')
        listbox = tk.Listbox(
            cadre_lb,
            selectmode='extended',
            yscrollcommand=scrollbar.set,
            font=(POLICE, 10),
            exportselection=False,
            bd=1, relief='solid',
            activestyle='none',
        )
        scrollbar.config(command=listbox.yview)

        zones = self._charger_zones_referentiel()
        for z in zones:
            listbox.insert('end', z)

        # Restaurer la sélection précédente si le popup est rouvert
        for i, z in enumerate(zones):
            if z in self._zones_selectionnees:
                listbox.selection_set(i)

        scrollbar.pack(side='right', fill='y')
        listbox.pack(side='left', fill='both', expand=True)

        cadre_boutons = tk.Frame(popup, bg=COULEURS['fond'])
        cadre_boutons.pack(fill='x', padx=16, pady=(10, 14))

        def valider():
            self._zones_selectionnees = [listbox.get(i) for i in listbox.curselection()]
            self._mettre_a_jour_label_zones()
            popup.destroy()

        tk.Button(
            cadre_boutons, text="OK",
            command=valider,
            bg=COULEURS['accent'], fg='white',
            font=(POLICE, 10, 'bold'), relief='flat', padx=20, pady=5,
            cursor='hand2'
        ).pack(side='right')

        tk.Button(
            cadre_boutons, text="Annuler",
            command=popup.destroy,
            bg=COULEURS['blanc'], fg=COULEURS['texte'],
            font=(POLICE, 10), relief='groove', padx=12, pady=5,
            cursor='hand2'
        ).pack(side='right', padx=(0, 8))

    def _mettre_a_jour_label_zones(self):
        """Met à jour le label affiché à côté du bouton Choisir après validation."""
        if not self._lbl_zones_choisies:
            return
        if self._zones_selectionnees:
            libelles = [ZONES_PECHE.get(z, f'ZEE de {z}') for z in self._zones_selectionnees]
            texte = ', '.join(libelles)
            self._lbl_zones_choisies.config(text=texte, fg=COULEURS['texte'], font=(POLICE, 9))
        else:
            self._lbl_zones_choisies.config(
                text="⚠ Sélectionnez au moins une zone",
                fg='#C55A11', font=(POLICE, 9, 'italic')
            )

    def _charger_zones_referentiel(self):
        """
        Charge la liste des zones depuis le référentiel local (table FpaZone, 53 entrées).
        Si le référentiel est absent, retourne les zones connues du dict ZONES_PECHE.
        """
        try:
            data = referentiel.charger_referentiel_local()
            if data:
                entrees = data.get('referentiels', {}).get(
                    'fr.ird.observe.entities.referential.common.FpaZone', [])
                noms = sorted(
                    {e['label2'] for e in entrees if e.get('label2') and e.get('status') == 'enabled'},
                    key=lambda x: x.lower()
                )
                if noms:
                    return noms
        except Exception:
            pass
        # Fallback sur les zones connues si le référentiel n'est pas disponible
        return sorted(ZONES_PECHE.keys(), key=lambda x: x.lower())

    def get_zone_peche(self):
        """
        Retourne la chaîne zone(s) de pêche à partir des zones sélectionnées dans le popup.
        Utilisée par la Phase 3 pour remplir {{ zone_peche }} dans le rapport.
        Exemple : ['Maurice', 'Seychelles'] → 'ZEE de Maurice et ZEE des Seychelles'
        Délègue à construire_zone_peche() de generer_rapport — source unique de logique.
        """
        if not self._zones_selectionnees:
            return '[Zone de pêche non renseignée]'
        if _PHASE3_OK:
            return _phase3.construire_zone_peche(self._zones_selectionnees)
        # Fallback si Phase 3 non disponible (docxtpl absent)
        libelles = [ZONES_PECHE.get(z, f'ZEE de {z}') for z in self._zones_selectionnees]
        if len(libelles) == 1:
            return libelles[0]
        return ', '.join(libelles[:-1]) + ' et ' + libelles[-1]

    def get_chemin_import(self):
        """Retourne le chemin complet du fichier ObServe importé, ou None si aucun."""
        return getattr(self, '_chemin_import', None)

    # ── Dossier de marée ──────────────────────────────────────────────────────

    def _construire_nom_dossier(self):
        """
        Construit le nom du dossier de marée à partir des données de la table 1_trip.
        Format : {trip_id}_{YYYYMMDD_debut}-{YYYYMMDD_fin}_{navire_sans_espaces}_{NOM_OBS}
        Exemple : SEALORXXX_20260101-20260115_NAVIRE_NOM

        Retourne (nom, date_fin_ok) :
          - nom         : le nom complet du dossier (XXXXXXXX si date de fin absente)
          - date_fin_ok : True si la date de fin est disponible, False sinon
        """
        if not self._tables_chargees:
            return None, False
        trip_df = self._tables_chargees.get('1_trip', pd.DataFrame())
        if trip_df.empty:
            return None, False

        t = trip_df.iloc[0]

        trip_id = str(t.get('trip_id', 'MAREE')).strip()

        # Date de début
        try:
            d_debut = pd.Timestamp(t.get('trip_start_date')).strftime('%Y%m%d')
        except Exception:
            d_debut = 'XXXXXXXX'

        # Date de fin — peut être absente en cours de marée
        date_fin_ok = False
        try:
            ts_fin = pd.Timestamp(t.get('trip_end_date'))
            if pd.isna(ts_fin):
                raise ValueError
            d_fin = ts_fin.strftime('%Y%m%d')
            date_fin_ok = True
        except Exception:
            d_fin = 'XXXXXXXX'

        # Navire : supprimer espaces, mettre en majuscules
        navire = str(t.get('vessel', 'NAVIRE')).strip().upper().replace(' ', '')

        # Nom de famille de l'observateur en majuscules
        nom_obs = str(t.get('observer_lastname', 'OBS')).strip().upper()

        return f"{trip_id}_{d_debut}-{d_fin}_{navire}_{nom_obs}", date_fin_ok

    def _creer_dossier_maree(self):
        """
        Demande à l'utilisateur où créer le dossier de la marée,
        puis crée le dossier principal et les 6 sous-dossiers standard.
        Si la date de fin est absente, XXXXXXXX est utilisé comme placeholder.
        """
        nom, date_fin_ok = self._construire_nom_dossier()
        if not nom:
            messagebox.showerror("Erreur", "Aucune donnée de marée chargée.")
            return

        parent = filedialog.askdirectory(
            title=f"Choisir où créer le dossier  «  {nom}  »"
        )
        if not parent:
            return

        chemin_principal = os.path.join(parent, nom)
        sous_dossiers = ['Formulaires', 'Observe', 'Photos', 'PhotosObsbio', 'Rapport', 'TDR']

        try:
            os.makedirs(chemin_principal, exist_ok=True)
            for sd in sous_dossiers:
                os.makedirs(os.path.join(chemin_principal, f"{nom}_{sd}"), exist_ok=True)

            msg = f"Structure créée dans :\n{chemin_principal}\n\n"
            msg += "\n".join(f"  • {nom}_{sd}" for sd in sous_dossiers)
            if not date_fin_ok:
                msg += "\n\n⚠ La date de fin est provisoire (XXXXXXXX).\nUtilisez « Corriger le nom du dossier » en fin de marée."
            messagebox.showinfo("Dossier créé", msg)
            self._var_statut.set(f"Dossier créé : {nom}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de créer le dossier :\n{e}")

    def _corriger_nom_dossier(self):
        """
        Renomme un dossier créé en début de marée (avec XXXXXXXX comme date de fin)
        en utilisant la vraie date de fin maintenant disponible dans ObServe.
        Renomme aussi les 6 sous-dossiers qui contiennent le nom de la marée.
        """
        nom_correct, date_fin_ok = self._construire_nom_dossier()
        if not nom_correct or not date_fin_ok:
            messagebox.showerror("Erreur", "La date de fin n'est pas disponible dans ce fichier.")
            return

        # Demander à l'utilisateur de sélectionner le dossier temporaire (XXXXXXXX)
        dossier_temp = filedialog.askdirectory(
            title="Sélectionner le dossier temporaire à corriger (celui avec XXXXXXXX)"
        )
        if not dossier_temp:
            return

        nom_temp = os.path.basename(dossier_temp)
        parent   = os.path.dirname(dossier_temp)

        # Vérifier que le dossier sélectionné est bien un dossier temporaire XXXXXXXX
        if 'XXXXXXXX' not in nom_temp:
            messagebox.showerror(
                "Dossier incorrect",
                f"Le dossier sélectionné ne contient pas 'XXXXXXXX' dans son nom :\n{nom_temp}\n\n"
                "Sélectionnez le bon dossier temporaire."
            )
            return

        # Vérifier que le trip_id correspond (pour éviter de renommer le mauvais dossier)
        trip_df = self._tables_chargees.get('1_trip', pd.DataFrame())
        trip_id = str(trip_df.iloc[0].get('trip_id', '')).strip() if not trip_df.empty else ''
        if trip_id and not nom_temp.startswith(trip_id):
            if not messagebox.askyesno(
                "Attention",
                f"Le dossier sélectionné ({nom_temp})\nne correspond pas à la marée en cours ({trip_id}).\n\nContinuer quand même ?"
            ):
                return

        chemin_correct = os.path.join(parent, nom_correct)
        sous_dossiers  = ['Formulaires', 'Observe', 'Photos', 'PhotosObsbio', 'Rapport', 'TDR']

        # Construire la liste complète des renommages dans l'ordre (sous-dossiers d'abord, principal en dernier)
        renames = []
        for sd in sous_dossiers:
            ancien_sd  = os.path.join(dossier_temp, f"{nom_temp}_{sd}")
            nouveau_sd = os.path.join(dossier_temp, f"{nom_correct}_{sd}")
            if os.path.exists(ancien_sd):
                renames.append((ancien_sd, nouveau_sd))
        renames.append((dossier_temp, chemin_correct))

        effectues = []
        try:
            for ancien, nouveau in renames:
                os.rename(ancien, nouveau)
                effectues.append((ancien, nouveau))

            messagebox.showinfo(
                "Dossier corrigé",
                f"Renommage effectué :\n\nAvant : {nom_temp}\nAprès : {nom_correct}"
            )
            self._var_statut.set(f"Dossier corrigé : {nom_correct}")

        except Exception as e:
            # Rollback : re-renommer dans l'ordre inverse ce qui a déjà été fait
            for ancien, nouveau in reversed(effectues):
                try:
                    os.rename(nouveau, ancien)
                except Exception:
                    pass
            messagebox.showerror(
                "Erreur — renommage annulé",
                f"Impossible de renommer un dossier :\n{e}\n\nLes modifications ont été annulées."
            )

    # ── Référentiel ───────────────────────────────────────────────────────────

    def _ouvrir_referentiel(self):
        """Ouvre la fenêtre de visualisation du référentiel ObServe."""
        FenetreReferentiel(self)

    def _mettre_a_jour_label_ref(self):
        """Met à jour l'étiquette de date du référentiel."""
        date = referentiel.date_referentiel_local()
        if date:
            self._lbl_date_ref.config(text=f"Mis à jour le : {date}")
        else:
            self._lbl_date_ref.config(text="Référentiel non téléchargé.")

    def _lancer_verification_referentiel(self):
        """Met à jour l'étiquette de date et lance la vérification en arrière-plan."""
        self._mettre_a_jour_label_ref()
        threading.Thread(target=self._thread_referentiel, daemon=True).start()

    def _thread_referentiel(self):
        """
        Exécuté en arrière-plan : vérifie la connexion et l'état du référentiel,
        puis propose le téléchargement si nécessaire.
        """
        date_locale = referentiel.date_referentiel_local()

        if not referentiel.verifier_internet():
            if date_locale is None:
                self.after(0, lambda: messagebox.showwarning(
                    "Référentiel absent",
                    "Aucune connexion internet détectée et le référentiel ObServe est absent.\n"
                    "Certaines vérifications de la Phase 1 ne pourront pas être effectuées."
                ))
            return

        # Internet disponible
        if date_locale is None:
            self.after(0, self._popup_telechargement_obligatoire)
        else:
            self.after(0, lambda d=date_locale: self._popup_proposer_mise_a_jour(d))

    def _popup_telechargement_obligatoire(self):
        if messagebox.askokcancel(
            "Référentiel manquant",
            "Le référentiel ObServe local est absent.\n"
            "Il est nécessaire pour les vérifications de la Phase 1.\n\n"
            "Télécharger maintenant ?"
        ):
            self._telecharger_avec_progression()

    def _popup_proposer_mise_a_jour(self, date_locale):
        if messagebox.askyesno(
            "Mise à jour du référentiel",
            f"Le référentiel local date du {date_locale}.\nVoulez-vous le mettre à jour ?"
        ):
            self._telecharger_avec_progression()

    def _telecharger_avec_progression(self):
        """Affiche une fenêtre de progression pendant le téléchargement du référentiel."""
        fenetre = tk.Toplevel(self)
        fenetre.title("Mise à jour du référentiel")
        fenetre.geometry("420x130")
        fenetre.resizable(False, False)
        fenetre.grab_set()
        fenetre.configure(bg=COULEURS['fond'])

        tk.Label(
            fenetre, text="Téléchargement du référentiel...",
            font=(POLICE, 11, 'bold'), bg=COULEURS['fond']
        ).pack(pady=(18, 6))

        var_msg = tk.StringVar(value="Initialisation...")
        tk.Label(
            fenetre, textvariable=var_msg,
            font=(POLICE, 9), bg=COULEURS['fond'], fg=COULEURS['texte']
        ).pack()

        # On capture tous les messages de progression pour récupérer l'erreur exacte
        messages = []
        def maj_msg(msg):
            messages.append(msg)
            fenetre.after(0, lambda m=msg: var_msg.set(m))

        def thread_dl():
            # Vérifier la connexion avant de lancer le téléchargement — dans le thread
            # pour ne pas geler l'interface pendant le timeout socket (max 5s)
            if not referentiel.verifier_internet():
                self.after(0, lambda: (
                    fenetre.destroy(),
                    messagebox.showerror(
                        "Pas de connexion",
                        "Aucune connexion au serveur IRD détectée "
                        "(observe.ob7.ird.fr inaccessible).\n\n"
                        "L'application fonctionne normalement avec le référentiel existant."
                    )
                ))
                return
            succes = referentiel.telecharger_referentiel(callback_progression=maj_msg)
            dernier_msg = messages[-1] if messages else "Erreur inconnue"
            if succes:
                self.after(0, lambda: (
                    fenetre.destroy(),
                    self._mettre_a_jour_label_ref(),
                    messagebox.showinfo("Succès", "Le référentiel a été mis à jour.")
                ))
            elif referentiel.ECHEC_IDENTIFIANTS:
                # Identifiants absents ou refusés par le serveur → proposer d'en saisir de nouveaux
                self.after(0, lambda dm=dernier_msg: (
                    fenetre.destroy(),
                    self._demander_identifiants(dm)
                ))
            else:
                self.after(0, lambda dm=dernier_msg: (
                    fenetre.destroy(),
                    messagebox.showwarning(
                        "Référentiel non mis à jour",
                        f"Détail de l'erreur :\n{dm}\n\n"
                        "L'application fonctionne normalement avec le référentiel existant."
                    )
                ))

        threading.Thread(target=thread_dl, daemon=True).start()

    def _demander_identifiants(self, detail):
        """
        Fenêtre ouverte uniquement quand le serveur IRD refuse les identifiants enregistrés
        (ou s'ils sont absents). Les nouveaux identifiants sont enregistrés dans le dossier
        de l'app (phase1_detection/data/identifiants_ird.json), puis le téléchargement est relancé.
        """
        fen = tk.Toplevel(self)
        fen.title("Identifiants du serveur IRD")
        fen.resizable(False, False)
        fen.configure(bg=COULEURS['fond'])
        fen.transient(self.winfo_toplevel())
        fen.grab_set()

        tk.Label(
            fen, bg=COULEURS['fond'], fg=COULEURS['texte'], font=(POLICE, 10), justify='left',
            wraplength=420,
            text=("Le serveur ObServe de l'IRD a refusé la connexion :\n"
                  f"{detail}\n\n"
                  "Les identifiants ont peut-être changé. Saisissez les nouveaux identifiants "
                  "fournis par l'IRD pour télécharger le référentiel."),
        ).grid(row=0, column=0, columnspan=2, padx=16, pady=(14, 10), sticky='w')

        login_actuel, _ = referentiel.charger_identifiants()
        var_login = tk.StringVar(value=login_actuel or '')
        var_mdp = tk.StringVar()
        for i, (texte, var, masque) in enumerate((("Identifiant :", var_login, ''),
                                                   ("Mot de passe :", var_mdp, '•')), start=1):
            tk.Label(fen, text=texte, bg=COULEURS['fond'], font=(POLICE, 10)).grid(
                row=i, column=0, padx=(16, 6), pady=4, sticky='e')
            champ = tk.Entry(fen, textvariable=var, show=masque, width=32, font=(POLICE, 10))
            champ.grid(row=i, column=1, padx=(0, 16), pady=4, sticky='w')
            if i == (2 if login_actuel else 1):
                champ.focus_set()

        def annuler():
            fen.destroy()
            messagebox.showwarning(
                "Référentiel non mis à jour",
                "Le référentiel n'a pas été mis à jour.\n\n"
                "L'application fonctionne normalement avec le référentiel existant."
            )

        def enregistrer():
            if not var_login.get().strip() or not var_mdp.get():
                messagebox.showerror("Identifiants incomplets",
                                     "Saisissez l'identifiant ET le mot de passe fournis par l'IRD.",
                                     parent=fen)
                return
            referentiel.sauvegarder_identifiants(var_login.get(), var_mdp.get())
            fen.destroy()
            self._telecharger_avec_progression()   # nouvel essai avec les nouveaux identifiants

        boutons = tk.Frame(fen, bg=COULEURS['fond'])
        boutons.grid(row=3, column=0, columnspan=2, pady=(10, 14))
        tk.Button(boutons, text="Enregistrer et réessayer", command=enregistrer,
                  bg=COULEURS['accent'], fg='white', font=(POLICE, 10, 'bold'),
                  relief='flat', padx=12, cursor='hand2').pack(side='left', padx=6)
        tk.Button(boutons, text="Annuler", command=annuler,
                  font=(POLICE, 10), relief='groove', padx=12, cursor='hand2').pack(side='left', padx=6)
        fen.bind('<Return>', lambda e: enregistrer())
        fen.bind('<Escape>', lambda e: annuler())
        fen.protocol("WM_DELETE_WINDOW", annuler)

    # ── Langue ────────────────────────────────────────────────────────────────

    def _changer_langue(self, event=None):
        """Sauvegarde la langue choisie et demande un redémarrage."""
        code = self._var_langue.get().lower()
        lang.sauvegarder_langue(code)
        nom = 'Français' if code == 'fr' else 'English'
        messagebox.showinfo(
            "Langue / Language",
            f"Langue changée en {nom}.\nRelancez l'application pour appliquer le changement."
        )


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION PRINCIPALE
# ─────────────────────────────────────────────────────────────────────────────

def main():

    # Doit être appelé AVANT tk.Tk() — sinon Windows ignore le changement d'icône taskbar
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('IRD.ObserveTools.App.1')
    except Exception:
        pass

    root = tk.Tk()
    root.title(f"Observe-Tools v{VERSION} — IRD OB7 SCISEA")
    root.geometry("1280x760")
    root.resizable(True, True)
    root.configure(bg=COULEURS['fond'])

    # Icône barre de titre + taskbar Windows
    _ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'assets', 'logos', 'observe-tools.ico')
    try:
        root.iconbitmap(_ico_path)
    except Exception:
        pass

    def _appliquer_icone_taskbar():
        """Force l'icône via Win32 après que la fenêtre soit visible."""
        try:
            import ctypes
            LR_LOADFROMFILE = 0x0010
            LR_DEFAULTSIZE  = 0x0040
            IMAGE_ICON      = 1
            WM_SETICON      = 0x0080
            # Handle de la fenêtre parente (pas le widget interne tkinter)
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            hicon = ctypes.windll.user32.LoadImageW(
                None, _ico_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE
            )
            if hicon:
                ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon)  # grande icône
                ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon)  # petite icône
        except Exception:
            pass

    root.after(200, _appliquer_icone_taskbar)  # après que la fenêtre soit dessinée

    # ── Style ttk global — thème 'clam' pour un look plus moderne ────────────
    style = ttk.Style()
    style.theme_use('clam')

    # Supprimer le pointillé de focus sur tous les ttk.Button
    style.configure('TButton', focuscolor='none')

    # Sélection dans les listes déroulantes des Combobox
    root.option_add('*TCombobox*Listbox.background',       '#FFFFFF',              'startupFile')
    root.option_add('*TCombobox*Listbox.selectBackground', COULEURS['en_tete'],    'startupFile')
    root.option_add('*TCombobox*Listbox.selectForeground', '#FFFFFF',              'startupFile')

    # Onglets — l'onglet actif est plus grand et plus haut que les autres
    style.configure('TNotebook', background=COULEURS['fond'], borderwidth=0)
    style.configure('TNotebook.Tab',
                    font=(POLICE, 9, 'bold'),
                    padding=[28, 5],
                    background='#DDE3F0',
                    foreground=COULEURS['texte'])
    style.map('TNotebook.Tab',
              background=[('disabled active', COULEURS['fond']),
                          ('selected', COULEURS['en_tete']),
                          ('active', '#C8D3E8'),
                          ('disabled', COULEURS['fond'])],
              foreground=[('selected', '#FFFFFF'),
                          ('disabled active', COULEURS['fond']),
                          ('disabled', COULEURS['fond'])],
              font=[('selected', (POLICE, 11, 'bold')), ('disabled', (POLICE, 6))],
              padding=[('selected', [28, 10]), ('disabled', [0, 5])])

    # ── Styles nommés pour les ttk.Button — palette moderne 2025 ─────────────
    # BigAccent.TButton — Indigo-600, grand bouton import (police 13, padding généreux)
    style.configure('BigAccent.TButton',
                    font=(POLICE, 13, 'bold'), padding=[20, 12],
                    background=COULEURS['accent'], foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('BigAccent.TButton',
              background=[('active', '#4338CA'), ('pressed', '#4338CA'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Accent.TButton — Sky-500 (annoter/commenter)
    style.configure('Accent.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background='#0EA5E9', foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('Accent.TButton',
              background=[('active', '#0284C7'), ('pressed', '#0284C7'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # OK.TButton — Emerald-600 (valider / ✓ Validé)
    style.configure('OK.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background=COULEURS['ok'], foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('OK.TButton',
              background=[('active', '#047857'), ('pressed', '#047857'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Dark.TButton — Indigo-950 (secondaire/nav)
    style.configure('Dark.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background=COULEURS['en_tete'], foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('Dark.TButton',
              background=[('active', '#16124A'), ('pressed', '#16124A'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Violet.TButton — Rouge-600 (vérifier tout)
    style.configure('Violet.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background='#DC2626', foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('Violet.TButton',
              background=[('active', '#B91C1C'), ('pressed', '#B91C1C'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Amber.TButton — Amber-600 (exporter / générer)
    style.configure('Amber.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background='#D97706', foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('Amber.TButton',
              background=[('active', '#B45309'), ('pressed', '#B45309'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # BigAmber.TButton — Amber-600 grand format (générer rapport de marée)
    style.configure('BigAmber.TButton',
                    font=(POLICE, 13, 'bold'), padding=[20, 12],
                    background='#D97706', foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('BigAmber.TButton',
              background=[('active', '#B45309'), ('pressed', '#B45309'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Teal.TButton — Teal-700 (gérer fichiers / créer dossier)
    style.configure('Teal.TButton',
                    font=(POLICE, 10, 'bold'), padding=[12, 5],
                    background='#0F766E', foreground='#FFFFFF',
                    borderwidth=0, relief='flat', anchor='center')
    style.map('Teal.TButton',
              background=[('active', '#0D6460'), ('pressed', '#0D6460'),
                          ('disabled', '#CBD5E1')],
              foreground=[('disabled', '#6B7280')],
              relief=[('pressed', 'flat')])

    # Combobox
    style.configure('TCombobox',
                    font=(POLICE, 10),
                    fieldbackground=COULEURS['blanc'],
                    background=COULEURS['fond'],
                    foreground=COULEURS['texte'],
                    bordercolor='#B0BDD4',
                    arrowcolor=COULEURS['accent'])
    style.map('TCombobox',
              fieldbackground=[('readonly', COULEURS['blanc']),
                               ('disabled', '#E5E7EB')],
              foreground=[('readonly', COULEURS['texte']),
                          ('disabled', '#9CA3AF')])

    # Scrollbar
    style.configure('TScrollbar',
                    background=COULEURS['fond'],
                    troughcolor='#E2E8F0',
                    arrowcolor=COULEURS['accent'],
                    borderwidth=0)

    # Progressbar (utilisée dans Phase 3)
    style.configure('TProgressbar',
                    troughcolor='#E2E8F0',
                    background=COULEURS['accent'],
                    borderwidth=0)

    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    # ── Création des frames des onglets ──────────────────────────────────────
    tab0 = tk.Frame(notebook, bg=COULEURS['fond'])   # Accueil
    tab1 = tk.Frame(notebook, bg=COULEURS['fond'])   # Phase 1
    tab2 = tk.Frame(notebook, bg=COULEURS['fond'])   # Phase 2
    tab3 = tk.Frame(notebook, bg=COULEURS['fond'])   # Phase 3

    notebook.add(tab0, text="  🏠  Accueil  ")
    notebook.add(tab1, text="  ✔  Correction des données  ")
    notebook.add(tab2, text="  📊  Figures et tableaux  ")
    notebook.add(tab3, text="  📄  Rapport  ")

    # ── Phases en mode embed — sans bouton import ni référentiel ni langue ─────
    # Ces éléments sont centralisés dans la page Accueil ci-dessous.
    app1 = AppObserveTools(tab1, mode='embed')
    app1.pack(fill='both', expand=True)

    app2 = AppPhase2(tab2, mode='embed')
    app2.pack(fill='both', expand=True)

    # ── Page d'accueil — créée avant app3 pour que get_zone_peche soit disponible ──
    # app3 est créé après accueil pour pouvoir passer accueil.get_zone_peche.
    # Le callback on_fichier_charge est défini après les deux.
    accueil = PageAccueil(tab0, on_fichier_charge=lambda t, n: None)  # callback temporaire
    accueil.pack(fill='both', expand=True)

    if _PHASE3_OK:
        app3 = AppPhase3(
            tab3,
            get_zone_peche=accueil.get_zone_peche,
            get_chemin_import=accueil.get_chemin_import,
        )
        app3.pack(fill='both', expand=True)
    else:
        # Phase 3 indisponible — afficher un message d'erreur dans l'onglet
        app3 = None
        msg = tk.Label(
            tab3,
            text=(
                "⚠  L'onglet Rapport Word est indisponible.\n\n"
                "Un module nécessaire n'a pas pu être chargé :\n"
                f"{_phase3_err_msg}\n\n"
                "Solution : fermez l'application et relancez INSTALLER.bat."
            ),
            bg=COULEURS['fond'], fg='#C00000',
            font=(POLICE, 11), justify='center', wraplength=500,
        )
        msg.pack(expand=True)

    # ── Callback réel : transmet les tables à toutes les phases après import ───
    def on_fichier_charge(tables, nom_fichier):
        chemin = accueil.get_chemin_import()  # chemin complet du zip (pour annotations Phase 1)
        app1._recevoir_tables(tables, nom_fichier, chemin_fichier=chemin)
        app2._recevoir_tables(tables, nom_fichier)
        if app3 is not None:
            app3._recevoir_tables(tables, nom_fichier)

    # Getter de zone : lu au moment de l'export (pas à l'import) pour être toujours à jour
    app1.set_zone_peche_getter(accueil.get_zone_peche)

    accueil._on_fichier_charge = on_fichier_charge  # remplace le callback temporaire

    # ── Fermeture propre ──────────────────────────────────────────────────────
    # Phase 2 crée un dossier temporaire — il faut le supprimer proprement.
    def _fermer():
        try:
            app2._fermer()
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", _fermer)

    root.mainloop()


if __name__ == '__main__':
    main()
