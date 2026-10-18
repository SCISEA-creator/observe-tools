#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phase 3 — Interface tkinter pour la génération du rapport Word.

Classe principale : AppPhase3(tk.Frame)
  - _recevoir_tables(tables, nom_fichier) : appelé par app.py après import
  - get_zone_peche et get_chemin_import    : fonctions passées depuis app.py
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RACINE, 'phase3_rapport'))
if RACINE not in sys.path:
    sys.path.insert(0, RACINE)

from utils import COULEURS, POLICE, VERSION, appliquer_hover, ajouter_logos_bandeau, ajouter_titre_bandeau
from generer_rapport import generer_rapport, get_val, ZONES_PECHE, calculer_nom_dossier, construire_zone_peche


class AppPhase3(tk.Frame):
    """
    Onglet Phase 3 — Génération du rapport Word.

    Paramètres
    ----------
    parent           : widget parent tkinter
    get_zone_peche   : callable() -> str  — lit la zone depuis la page Accueil
    get_chemin_import: callable() -> str|None — lit le chemin du zip importé
    """

    def __init__(self, parent, get_zone_peche=None, get_chemin_import=None):
        super().__init__(parent)
        self._tables          = None   # tables ObServe courantes
        self._nom_fichier     = None   # basename du zip importé
        self._get_zone_peche  = get_zone_peche  or (lambda: '[Zone non renseignée]')
        self._get_chemin      = get_chemin_import or (lambda: None)
        self._construction_ui()

    # ── Construction de l'interface ──────────────────────────────────────────

    def _construction_ui(self):
        """Construit les widgets de l'onglet Phase 3."""
        F = COULEURS['fond']
        self.configure(bg=F)

        # ── En-tête bleu (identique aux autres onglets) ───────────────────────
        en_tete = tk.Frame(self, bg=COULEURS['en_tete'], height=52)
        en_tete.pack(fill='x')
        en_tete.pack_propagate(False)

        self._logo_titre = ajouter_titre_bandeau(
            en_tete, COULEURS['en_tete'], "Observe-Tools — Rapport de marée"
        )

        self._logos = ajouter_logos_bandeau(en_tete, COULEURS['en_tete'])

        # ── Barre de statut bleue en bas (identique aux autres onglets) ───────
        self._var_statut = tk.StringVar(value="En attente d'une marée importée...")
        barre_statut = tk.Frame(self, bg=COULEURS['en_tete'], height=26)
        barre_statut.pack(fill='x', side='bottom')
        barre_statut.pack_propagate(False)
        self._lbl_dot = tk.Label(barre_statut, text='●', bg=COULEURS['en_tete'], fg='#94A3B8', font=(POLICE, 9))
        self._lbl_dot.pack(side='left', padx=(10, 2))
        tk.Label(
            barre_statut, bg=COULEURS['en_tete'], fg='white',
            textvariable=self._var_statut, font=(POLICE, 9), anchor='w'
        ).pack(side='left', padx=(0, 10))

        def _maj_dot(*args):
            msg = self._var_statut.get()
            m = msg.lower()
            if '✗' in msg or 'erreur' in m or 'impossible' in m:
                fg = '#F87171'
            elif '⚠' in msg or 'avertissement' in m:
                fg = '#FBB76A'
            elif 'cours' in m or 'génération' in m:
                fg = '#93C5FD'
            elif '✓' in msg or 'généré' in m or 'prêt' in m:
                fg = '#4ADE80'
            else:
                fg = '#94A3B8'
            self._lbl_dot.config(fg=fg)
        self._var_statut.trace_add('write', _maj_dot)

        # ── Contenu central ───────────────────────────────────────────────────
        cadre = tk.Frame(self, bg=F, padx=24, pady=20)
        cadre.pack(fill='both', expand=True)

        # ── Infos marée ───────────────────────────────────────────────────────
        tk.Label(cadre, bg=F, text="Marée :",
                 font=(POLICE, 10, 'bold'), fg=COULEURS['texte']).pack(anchor='w')
        self._var_maree_detail = tk.StringVar(value="Aucune marée importée")
        tk.Label(cadre, bg=F, textvariable=self._var_maree_detail,
                 font=(POLICE, 11, 'bold'), fg=COULEURS['texte']).pack(anchor='w', padx=(16, 0))

        tk.Frame(cadre, bg=F, height=8).pack()

        # ── Destination du fichier ────────────────────────────────────────────
        tk.Label(cadre, bg=F, text="Fichier de sortie :",
                 font=(POLICE, 10, 'bold'), fg=COULEURS['texte']).pack(anchor='w')
        self._var_destination = tk.StringVar(value="—")
        tk.Label(cadre, bg=F, textvariable=self._var_destination,
                 font=(POLICE, 9), fg='#6B7280',
                 wraplength=580, justify='left').pack(anchor='w', padx=(16, 0))

        tk.Frame(cadre, bg=F, height=16).pack()

        # ── Bouton générer ────────────────────────────────────────────────────
        self._btn = ttk.Button(
            cadre,
            text="  📄  Générer le rapport de marée  ",
            style='BigAmber.TButton',
            state='disabled',
            command=self._generer,
            cursor='hand2'
        )
        self._btn.pack(anchor='w')

        tk.Frame(cadre, bg=F, height=12).pack()

        # ── Barre de progression (cachée par défaut) ──────────────────────────
        self._barre = ttk.Progressbar(cadre, mode='indeterminate', length=400)
        # La barre n'est pas affichée ici — elle apparaît uniquement pendant la génération

    # ── Réception des données ────────────────────────────────────────────────

    def _recevoir_tables(self, tables, nom_fichier):
        """
        Appelé par app.py après chaque import de fichier ObServe.
        Stocke les données et active le bouton de génération.
        """
        self._tables      = tables
        self._nom_fichier = nom_fichier

        # Infos marée pour l'affichage
        trip_df = tables.get('1_trip')
        if trip_df is not None and not trip_df.empty:
            trip_id     = get_val(trip_df, 'trip_id', '—')
            nom_dossier = self._calculer_nom_dossier(trip_df)
            self._var_maree_detail.set(nom_dossier)
        else:
            self._var_maree_detail.set(nom_fichier)

        # Afficher la destination probable du fichier
        self._var_destination.set(self._previsualiser_destination(trip_df))

        # Activer le bouton
        self._btn.config(state='normal')
        self._var_statut.set("Prêt à générer le rapport.")

    # ── Génération ───────────────────────────────────────────────────────────

    def _generer(self):
        """Appelé au clic du bouton — orchestre la génération du rapport."""
        if self._tables is None:
            messagebox.showwarning("Attention", "Aucune marée importée.")
            return

        # Lire la zone de pêche au moment du clic
        zone_peche = self._get_zone_peche()
        if not zone_peche or zone_peche.startswith('['):
            messagebox.showerror(
                "Zone de pêche manquante",
                "Aucune zone de pêche n'est sélectionnée.\n\n"
                "Retournez dans l'onglet Accueil pour choisir une zone avant de générer le rapport."
            )
            return

        # Calculer le nom du fichier de sortie
        trip_df    = self._tables.get('1_trip')
        nom_dossier = self._calculer_nom_dossier(trip_df)
        nom_fichier_docx = f"{nom_dossier}_Rapport.docx"

        # Chercher le dossier _Rapport/ à côté du zip importé
        output_path = self._trouver_chemin_sortie(nom_fichier_docx)
        if output_path is None:
            return  # l'utilisateur a annulé la boîte de dialogue

        # Confirmation si le fichier existe déjà
        if os.path.exists(output_path):
            if not messagebox.askyesno(
                "Fichier existant",
                f"Le fichier suivant existe déjà :\n{output_path}\n\nL'écraser ?"
            ):
                return

        # Désactiver le bouton et afficher la barre de progression
        self._btn.config(state='disabled')
        self._var_statut.set("Génération en cours...")
        self._barre.pack(anchor='w', pady=(8, 0))
        self._barre.start(12)  # intervalle en ms entre chaque déplacement
        self.update_idletasks()

        # Snapshot des tables AVANT le lancement du thread : si l'utilisateur importe
        # une nouvelle marée pendant la génération, le thread garde les données d'origine.
        tables_snapshot = self._tables

        # Lancer la génération dans un thread pour ne pas bloquer l'interface
        threading.Thread(
            target=self._generer_thread,
            args=(tables_snapshot, zone_peche, output_path),
            daemon=True
        ).start()

    def _generer_thread(self, tables, zone_peche, output_path):
        """Exécuté dans un thread séparé — appelle generer_rapport()."""
        try:
            nb_avert = generer_rapport(tables, zone_peche, output_path)
            # Retour dans le thread principal pour mettre à jour l'interface
            self.after(0, self._generation_ok, output_path, nb_avert)
        except PermissionError:
            self.after(0, self._generation_erreur,
                       "Impossible d'écrire le fichier.\n\n"
                       "Si le rapport est déjà ouvert dans Word, ferme-le d'abord.")
        except Exception as e:
            self.after(0, self._generation_erreur, str(e))

    def _generation_ok(self, output_path, nb_avert=0):
        """Appelé dans le thread principal après succès."""
        self._barre.stop()
        self._barre.pack_forget()
        self._btn.config(state='normal')
        nom = os.path.basename(output_path)
        if nb_avert > 0:
            self._var_statut.set(
                f"✓ Rapport généré : {nom}\n"
                f"⚠ {nb_avert} avertissement(s) de calcul — vérifier les chiffres.\n"
                f"{output_path}"
            )
            messagebox.showwarning(
                "Rapport généré avec avertissements",
                f"{nb_avert} calcul(s) ont échoué et ont été remplacés par « — » dans le rapport.\n\n"
                "Vérifiez les valeurs marquées « — » avant de soumettre le rapport."
            )
        else:
            self._var_statut.set(f"✓ Rapport généré : {nom}\n{output_path}")

    def _generation_erreur(self, message):
        """Appelé dans le thread principal après échec."""
        self._barre.stop()
        self._barre.pack_forget()
        self._btn.config(state='normal')
        self._var_statut.set(f"✗ Erreur lors de la génération.")
        messagebox.showerror("Erreur de génération", message)

    # ── Calculs internes ─────────────────────────────────────────────────────

    def _calculer_nom_dossier(self, trip_df):
        """Délègue à calculer_nom_dossier() — source unique dans generer_rapport.py."""
        return calculer_nom_dossier(trip_df)

    def _trouver_chemin_sortie(self, nom_fichier_docx):
        """
        Cherche le dossier _Rapport/ à côté du zip importé.
        Sinon, ouvre une boîte de dialogue pour choisir le dossier.
        Retourne le chemin complet du .docx, ou None si l'utilisateur annule.
        """
        chemin_zip = self._get_chemin()
        if chemin_zip:
            dossier_zip = os.path.dirname(chemin_zip)
            # Chercher _Rapport/ dans le même dossier que le zip
            dossier_rapport = os.path.join(dossier_zip, '_Rapport')
            if os.path.isdir(dossier_rapport):
                return os.path.join(dossier_rapport, nom_fichier_docx)
            # Chercher _Rapport/ dans le dossier parent (si le zip est dans un sous-dossier)
            dossier_parent = os.path.dirname(dossier_zip)
            dossier_rapport_parent = os.path.join(dossier_parent, '_Rapport')
            if os.path.isdir(dossier_rapport_parent):
                return os.path.join(dossier_rapport_parent, nom_fichier_docx)

        # Aucun dossier _Rapport/ trouvé — boîte de dialogue
        dossier = filedialog.askdirectory(
            title="Choisir le dossier de sauvegarde du rapport"
        )
        if not dossier:
            return None
        return os.path.join(dossier, nom_fichier_docx)

    def _previsualiser_destination(self, trip_df):
        """
        Calcule et retourne le chemin probable du fichier de sortie, SANS dialog.
        Utilisé pour afficher la destination dans l'interface avant de cliquer.
        """
        nom_dossier = self._calculer_nom_dossier(trip_df)
        nom_fichier_docx = f"{nom_dossier}_Rapport.docx"

        chemin_zip = self._get_chemin()
        if chemin_zip:
            dossier_zip = os.path.dirname(chemin_zip)
            dossier_rapport = os.path.join(dossier_zip, '_Rapport')
            if os.path.isdir(dossier_rapport):
                return os.path.join(dossier_rapport, nom_fichier_docx)
            dossier_parent = os.path.dirname(dossier_zip)
            dossier_rapport_parent = os.path.join(dossier_parent, '_Rapport')
            if os.path.isdir(dossier_rapport_parent):
                return os.path.join(dossier_rapport_parent, nom_fichier_docx)
            # Zip trouvé mais pas de dossier _Rapport/ — signaler à l'utilisateur
            return f"Dossier _Rapport/ introuvable — une boîte de dialogue s'ouvrira au moment de générer"

        return "Aucun fichier importé — le dossier de sortie sera demandé à la génération"
