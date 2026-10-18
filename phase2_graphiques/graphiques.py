#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Phase 2 : Génération des graphiques
IRD OB7 - SCISEA

Ce module génère tous les graphiques matplotlib pour le rapport.
Chaque fonction prend les données calculées (depuis calculs.py) et
retourne un objet Figure matplotlib prêt à être sauvegardé.

Backend 'Agg' = génération hors-ligne, sans fenêtre (compatible tkinter).
"""

import matplotlib
matplotlib.use('Agg')  # doit être fait avant tout import de pyplot

import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# SYSTÈME DE COULEURS PAR GROUPE D'ESPÈCES
# ─────────────────────────────────────────────────────────────────────────────

# Couleur de base pour chaque groupe — appliquée à l'espèce la plus capturée
# et aux figures mono-espèce (camembert global, fig_aucune_capture, etc.).
COULEURS_BASE = {
    'Porte-épées':     '#1F4E79',  # bleu foncé
    'Thons':           '#C00000',  # rouge foncé
    'Requins':         '#00695C',  # bleu-vert (teal)
    'Raies':           '#553280',  # violet foncé désaturé
    'Poissons osseux': '#C47A00',  # jaune orangé foncé (ambre)
    'Cétacés':         '#1A1A1A',  # noir
    'Tortues':         '#6B8C2A',  # vert olive (tortue)
}


# Ordre d'affichage standard des groupes dans tous les graphiques
ORDRE_GROUPES = [
    'Porte-épées', 'Thons', 'Requins',
    'Raies', 'Poissons osseux', 'Tortues', 'Cétacés',
]

# Couleurs fonctionnelles pour le devenir (indépendantes du groupe)
COULEUR_CONSERVE = '#C03030'   # rouge foncé (conservés en cale)
COULEUR_REJETE   = '#4472C4'   # bleu
COULEUR_ECHAPPE  = '#A9D18E'   # vert clair

# Convention vivant / mort — cohérente dans TOUTES les figures
COULEUR_VIVANT    = '#70AD47'  # vert
COULEUR_MORT      = '#767171'  # gris foncé
COULEUR_INCERTAIN = '#ED7D31'  # orange (ni vivant ni mort confirmé)

# Couleurs des déprédateurs par nom français — cohérentes avec les groupes d'espèces
# Les déprédateurs non listés reçoivent une couleur neutre (#888888).
COULEURS_DEPREDATEURS = {
    'Calmar':          '#8B3020',  # rouge brun rouillé (couleur calmar)
    'Requin':          '#00695C',  # même couleur que le groupe Requins
    'Squalelet féroce':'#636363',  # gris neutre
    'Odontocète':      '#1A1A1A',  # même couleur que le groupe Cétacés
    'Céphalopode':     '#A0522D',  # brun
    'Oiseaux marins':  '#FFF0D0',  # blanc crème
}

# Tailles des figures (en pouces)
FIG_CARRE  = (7, 6)    # camemberts
FIG_LARGE  = (10, 6)   # histogrammes et barres
FIG_XLARGE = (12, 7)   # barres avec beaucoup d'espèces


def _palette_groupe(groupe_fr, n):
    """
    Génère n couleurs pour un groupe, de la plus foncée (base) à la plus claire.
    Interpolation linéaire depuis la couleur de base (COULEURS_BASE) vers le blanc.
    fade_max=0.80 : contraste bien visible, tout en restant dans la famille de couleurs
    du groupe (bleu foncé pour porte-épées, bleu-vert pour requins, rouge pour thons...).
    """
    base_hex = COULEURS_BASE.get(groupe_fr, '#888888')
    if n <= 0:
        return []
    if n == 1:
        return [base_hex]

    base_rgb = mcolors.to_rgb(base_hex)
    blanc    = (1.0, 1.0, 1.0)
    fade_max = 0.80  # 80% vers le blanc → bien différencié, reste dans la famille de couleurs

    palette = []
    for i in range(n):
        t = (i / (n - 1)) * fade_max
        r = base_rgb[0] + (blanc[0] - base_rgb[0]) * t
        g = base_rgb[1] + (blanc[1] - base_rgb[1]) * t
        b = base_rgb[2] + (blanc[2] - base_rgb[2]) * t
        palette.append(mcolors.to_hex((r, g, b)))
    return palette


def _eclaircir(couleur_hex, facteur=0.45):
    """
    Retourne une version plus claire d'une couleur.
    facteur = 0 → identique, facteur = 1 → blanc pur.
    Utilisé pour les barres "Vivant" dans l'état à la capture.
    """
    rgb = mcolors.to_rgb(couleur_hex)
    r = rgb[0] + (1.0 - rgb[0]) * facteur
    g = rgb[1] + (1.0 - rgb[1]) * facteur
    b = rgb[2] + (1.0 - rgb[2]) * facteur
    return mcolors.to_hex((r, g, b))


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION GLOBALE
# ─────────────────────────────────────────────────────────────────────────────

def _style_global():
    """
    Applique un style propre et cohérent.
    Appelé au début de chaque fonction de figure.
    """
    matplotlib.rcParams.update({
        'font.family':       'DejaVu Sans',  # supporte les accents français
        'font.size':         10,
        'axes.titlesize':    11,
        'axes.titleweight':  'bold',
        'axes.labelsize':    10,
        'figure.dpi':        150,
        'axes.spines.top':   False,
        'axes.spines.right': False,
        'axes.grid':         False,
    })


def _fermer(fig):
    """Ferme la figure pour libérer la mémoire (appeler après savefig)."""
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES INTERNES
# ─────────────────────────────────────────────────────────────────────────────

def _camembert(data, titre, ax, fig, couleurs, trier=True):
    """
    Trace un camembert sur l'axes donné.
    couleurs : liste de couleurs hex dans le même ordre que data.
    trier    : si True, trie par valeur décroissante (ordre visuel standard).
               Passer False pour conserver l'ordre fourni (ordre fixe par espèce).

    Affichage des pourcentages :
    - Part >= 7% : % en blanc à l'intérieur du secteur
    - Part  > 0% : % à l'extérieur avec une courte ligne de connexion (toutes les espèces)
    """
    if trier:
        data = data.sort_values(ascending=False)

    total = data.values.sum()

    # Tracer le camembert sans autopct (on gère les labels manuellement)
    wedges, _ = ax.pie(
        data.values,
        labels=None,
        colors=couleurs[:len(data)],
        startangle=90,
        wedgeprops={'linewidth': 0},
    )

    # ── Pourcentages avec anti-chevauchement angulaire ───────────────────────
    # Principe : le rayon est FIXE pour tout le monde (flèches toujours courtes).
    # Seul l'ANGLE du texte est ajusté pour espacer les étiquettes trop proches.
    # La flèche pointe toujours vers le centre naturel du secteur.

    items = []
    for wedge, val in zip(wedges, data.values):
        if total == 0 or val == 0:
            continue
        pct       = val / total * 100
        nat_angle = (wedge.theta2 + wedge.theta1) / 2  # angle naturel du secteur
        items.append({
            'pct'      : pct,
            'nat_angle': nat_angle,
            'txt_angle': nat_angle,  # sera ajusté ci-dessous
        })

    # Étaler angulièrement les étiquettes trop proches.
    # À r=1.18, 13° d'arc ≈ 0.27 unités — largement suffisant pour fontsize=7.
    MIN_DEG = 13.0
    for _ in range(80):
        moved = False
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                a, b = items[i], items[j]
                diff = (b['txt_angle'] - a['txt_angle']) % 360
                if diff > 180:
                    diff = 360 - diff  # distance angulaire la plus courte
                if diff < MIN_DEG:
                    shift = (MIN_DEG - diff) / 2 + 0.3
                    a['txt_angle'] -= shift
                    b['txt_angle'] += shift
                    moved = True
        if not moved:
            break

    # Dessiner chaque étiquette avec son angle ajusté et un rayon constant
    R_TEXTE = 1.18
    for item in items:
        nat_cos = np.cos(np.radians(item['nat_angle']))
        nat_sin = np.sin(np.radians(item['nat_angle']))
        txt_cos = np.cos(np.radians(item['txt_angle']))
        txt_sin = np.sin(np.radians(item['txt_angle']))
        ha = 'center' if abs(txt_cos) < 0.4 else ('left' if txt_cos > 0 else 'right')
        ax.annotate(
            f"{item['pct']:.1f}%",
            xy=(0.95 * nat_cos, 0.95 * nat_sin),           # pointe sur le bord du secteur
            xytext=(R_TEXTE * txt_cos, R_TEXTE * txt_sin), # texte à angle ajusté, même rayon
            ha=ha, va='center', fontsize=7, color='black',
            annotation_clip=False,
            arrowprops=dict(arrowstyle='-', color='#999999', lw=0.7),
            bbox=dict(boxstyle='round,pad=0.1', facecolor='white',
                      edgecolor='none', alpha=0.8),
        )

    legendes = [f'{nom} ({n})' for nom, n in zip(data.index, data.values)]
    ax.legend(legendes, loc='center left', bbox_to_anchor=(1, 0.5),
              fontsize=9, frameon=False)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE DE SUBSTITUTION — GROUPE SANS CAPTURE
# ─────────────────────────────────────────────────────────────────────────────

def fig_aucune_capture(groupe_fr):
    """
    Générée à la place d'un camembert quand un groupe n'a aucune capture.
    Affiche un message centré avec la couleur du groupe pour rester cohérent
    visuellement avec les autres figures du rapport.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_CARRE)

    couleur = COULEURS_BASE.get(groupe_fr, '#888888')

    # Message centré
    ax.text(
        0.5, 0.5,
        f'Aucune capture de {groupe_fr}\nlors de cette marée',
        ha='center', va='center',
        fontsize=13, fontweight='bold',
        color=couleur,
        transform=ax.transAxes,
    )

    ax.set_axis_off()

    # Cadre coloré pour identifier visuellement le groupe
    fig.patch.set_linewidth(2)
    fig.patch.set_edgecolor(couleur)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 1 — CAPTURES PAR GROUPE D'ESPÈCES
# ─────────────────────────────────────────────────────────────────────────────

def fig_captures_par_groupe(data):
    """
    Figure 1 : camembert des captures par groupe.
    data : Series {groupe_fr: nb_captures}
    Chaque groupe utilise sa couleur de base définie dans COULEURS_BASE.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_CARRE)

    data = data.sort_values(ascending=False)
    couleurs = [COULEURS_BASE.get(g, '#888888') for g in data.index]

    fig = _camembert(data, titre="Proportion de chaque groupe d'espèces\nsur l'ensemble des captures de la marée",
                     ax=ax, fig=fig, couleurs=couleurs)
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURES 2-8 — CAMEMBERTS PAR GROUPE D'ESPÈCES
# ─────────────────────────────────────────────────────────────────────────────

def fig_captures_groupe_detail(data, titre, groupe_fr, couleurs_especes=None):
    """
    Figures 2-8 : camembert détaillé pour un groupe.
    data            : Series {fao_code: nb_captures}
    groupe_fr       : nom du groupe en français (pour la palette de secours)
    couleurs_especes: dict {fao_code: couleur_hex} avec l'ordre fixe IRD.
                      Si fourni, les couleurs ET l'ordre des espèces sont fixes.
                      Sinon, dégradé trié par nombre de captures décroissant.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_CARRE)

    if couleurs_especes:
        # Ordonner les codes selon l'ordre fixe du mapping (codes présents en premier,
        # puis les espèces rares hors mapping triées par nb de captures)
        codes_fixes   = [c for c in couleurs_especes if c in data.index]
        codes_restants = sorted(
            [c for c in data.index if c not in couleurs_especes],
            key=lambda c: -data[c]
        )
        data     = data.reindex(codes_fixes + codes_restants)
        couleurs = [couleurs_especes.get(c, '#888888') for c in data.index]
    else:
        data     = data.sort_values(ascending=False)
        couleurs = _palette_groupe(groupe_fr, len(data))

    fig = _camembert(data, titre=titre, ax=ax, fig=fig, couleurs=couleurs, trier=False)
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURES 6-9 — HISTOGRAMMES DE TAILLE
# ─────────────────────────────────────────────────────────────────────────────

def fig_classe_taille(data, nom_espece, couleur=None):
    """
    Figures 6-9 : histogramme des classes de taille.
    couleur : couleur hex de l'espèce (depuis le mapping groupe).
              Si None, utilise le bleu de base.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_LARGE)

    couleur_barres = couleur or COULEURS_BASE.get('Porte-épées', '#1F4E79')

    ax.bar(range(len(data)), data.values, color=couleur_barres,
           edgecolor='white', linewidth=0.5, width=0.8)

    # Pourcentage de chaque classe au-dessus de chaque barre avec ligne de connexion
    total   = data.values.sum()
    max_val = data.values.max() if data.values.max() > 0 else 1
    if total > 0:
        ax.set_ylim(0, max_val * 1.22)  # espace pour les annotations
        for i, val in enumerate(data.values):
            if val > 0:
                pct = val / total * 100
                ax.annotate(
                    f'{pct:.1f}%',
                    xy=(i, val),
                    xytext=(i, val + max_val * 0.06),
                    ha='center', va='bottom', fontsize=7, color='#333333',
                    annotation_clip=False,
                    arrowprops=dict(arrowstyle='-', color='#BBBBBB', lw=0.6),
                )

    ax.set_xticks(range(len(data)))
    ax.set_xticklabels(data.index, rotation=45, ha='right', fontsize=9)
    ax.set_xlabel('Classe de taille (cm)')
    ax.set_ylabel("Nombre d'individus")
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


def fig_classes_taille_thons(data_dict, couleurs_especes=None):
    """
    Figure 10 : classes de taille des thons (ALB, BET, YFT) dans un seul graphique.
    Barres groupées par classe de taille, une couleur par espèce.
    Les pourcentages sont calculés sur le total de chaque espèce séparément.
    data_dict       : {code_fao: Series classes de taille}
    couleurs_especes: dict {code_fao: couleur_hex}
    """
    _style_global()
    codes = [c for c, d in data_dict.items() if d is not None and len(d) > 0]
    if not codes:
        return None

    # Union de toutes les classes de taille présentes, triées numériquement
    toutes_classes = sorted(
        set().union(*[set(data_dict[c].index) for c in codes]),
        key=lambda x: int(x.split('-')[0])
    )

    n_classes = len(toutes_classes)
    n_especes = len(codes)
    largeur   = 0.75 / n_especes  # largeur de chaque barre dans le groupe

    fig, ax = plt.subplots(figsize=FIG_LARGE)
    x = np.arange(n_classes)

    max_global = 0
    for i, code in enumerate(codes):
        data    = data_dict[code].reindex(toutes_classes, fill_value=0)
        couleur = (couleurs_especes or {}).get(code) or COULEURS_BASE.get('Thons', '#C00000')
        offset  = (i - (n_especes - 1) / 2) * largeur

        barres = ax.bar(x + offset, data.values, largeur * 0.92,
                        color=couleur, label=code,
                        edgecolor='white', linewidth=0.5)
        max_global = max(max_global, data.values.max())

        # Pourcentage au-dessus de chaque barre (% du total de l'espèce)
        total = data_dict[code].values.sum()
        for bar, val in zip(barres, data.values):
            if val > 0 and total > 0:
                pct = val / total * 100
                ax.text(bar.get_x() + bar.get_width() / 2, val + max_global * 0.01,
                        f'{pct:.0f}%', ha='center', va='bottom',
                        fontsize=6, color='#333333')

    ax.set_ylim(0, max_global * 1.20)
    ax.set_xticks(x)
    ax.set_xticklabels(toutes_classes, rotation=45, ha='right', fontsize=9)
    ax.set_xlabel('Classe de taille (cm)')
    ax.set_ylabel("Nombre d'individus")
    ax.legend(frameon=False, loc='upper right')
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 11 — ÉTAT À LA CAPTURE DES ESPÈCES COMMERCIALES
# ─────────────────────────────────────────────────────────────────────────────

def fig_etat_capture(data):
    """
    Figure 11 : barres Mort / Vivant à la capture pour les espèces conservées.
    Mort = gris, Vivant = vert (convention cohérente avec toutes les figures).
    """
    _style_global()
    n   = len(data)
    fig, ax = plt.subplots(figsize=(max(10, n * 0.7), 6))

    x     = np.arange(n)
    width = 0.38

    ax.bar(x - width / 2, data['Mort'].values,   width,
           color=COULEUR_MORT,   edgecolor='white', linewidth=0.5, label='Capturé mort')
    ax.bar(x + width / 2, data['Vivant'].values, width,
           color=COULEUR_VIVANT, edgecolor='white', linewidth=0.5, label='Capturé vivant')

    # Pourcentage de vivants au-dessus de chaque paire de barres avec ligne de connexion
    totaux  = (data['Vivant'] + data['Mort']).values
    max_tot = totaux.max() if totaux.max() > 0 else 1
    # Étendre l'axe Y pour laisser de la place aux annotations au-dessus des barres
    ax.set_ylim(0, max_tot * 1.22)
    for i, espece in enumerate(data.index):
        total = totaux[i]
        if total > 0:
            pct_viv = data.loc[espece, 'Vivant'] / total * 100
            ax.annotate(
                f'{pct_viv:.0f}% vivants',
                xy=(i, total),
                xytext=(i, total + max_tot * 0.08),
                ha='center', va='bottom', fontsize=8, color='#333333',
                annotation_clip=False,
                arrowprops=dict(arrowstyle='-', color='#BBBBBB', lw=0.6),
            )

    ax.set_xticks(x)
    ax.set_xticklabels(data.index, rotation=40, ha='right', fontsize=9)
    ax.set_ylabel("Nombre d'individus")
    ax.legend(frameon=False)
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 12 — SURVIE DES REQUINS
# ─────────────────────────────────────────────────────────────────────────────

def fig_survie_requins(data):
    """
    Figure 12 : camembert Vivant / Mort au relâcher pour les requins.
    Vivant = vert, Mort = gris, Incertain = orange (convention cohérente).
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_CARRE)

    couleurs_map = {
        'Vivant':    COULEUR_VIVANT,
        'Mort':      COULEUR_MORT,
        'Incertain': COULEUR_INCERTAIN,
    }
    couleurs = [couleurs_map.get(k, '#AAAAAA') for k in data.index]

    _camembert(data, titre='État au rejet des requins capturés',
               ax=ax, fig=fig, couleurs=couleurs)
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 13 — DEVENIR GLOBAL DES CAPTURES
# ─────────────────────────────────────────────────────────────────────────────

def fig_devenir_global(data):
    """
    Figure 13 : camembert conservés / rejetés / échappés.
    Couleurs fonctionnelles (pas de groupe) : jaune / bleu / vert.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_CARRE)

    couleurs_map = {
        'Conservés': COULEUR_CONSERVE,
        'Rejetés':   COULEUR_REJETE,
        'Échappés':  COULEUR_ECHAPPE,
    }
    couleurs = [couleurs_map.get(k, '#999999') for k in data.index]

    _camembert(data, titre='Devenir des individus capturés pendant la marée',
               ax=ax, fig=fig, couleurs=couleurs)
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 14 — DEVENIR PAR ESPÈCE (barres empilées)
# ─────────────────────────────────────────────────────────────────────────────

def fig_devenir_par_espece(data):
    """
    Figure 14 : barres empilées conservés / rejetés / échappés par espèce.
    Couleurs fonctionnelles (devenir) : jaune / bleu / vert.
    """
    _style_global()
    n   = len(data)
    fig, ax = plt.subplots(figsize=(max(10, n * 0.6), 6))

    couleurs_map = {
        'Conservés': COULEUR_CONSERVE,
        'Rejetés':   COULEUR_REJETE,
        'Échappés':  COULEUR_ECHAPPE,
    }
    bas = np.zeros(n)
    for col in ['Conservés', 'Rejetés', 'Échappés']:
        if col in data.columns:
            vals = data[col].values.astype(float)
            ax.bar(range(n), vals, bottom=bas, label=col,
                   color=couleurs_map[col], linewidth=0)
            bas += vals

    ax.set_xticks(range(n))
    ax.set_xticklabels(data.index, rotation=40, ha='right', fontsize=9)
    ax.set_ylabel("Nombre d'individus")
    ax.legend(frameon=False, loc='upper right')
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 15 — DÉPRÉDATION TOTALE
# ─────────────────────────────────────────────────────────────────────────────

def fig_depredation_totale(data):
    """
    Figure 15 : histogramme de la déprédation par type de déprédateur.
    Une barre par déprédateur avec pourcentage, trié par nb décroissant.
    Couleurs par nom (COULEURS_DEPREDATEURS).
    """
    _style_global()
    data = data.sort_values(ascending=False)
    n    = len(data)
    fig, ax = plt.subplots(figsize=(max(6, n * 1.4), 6))

    couleurs = [COULEURS_DEPREDATEURS.get(k, '#888888') for k in data.index]
    ax.bar(range(n), data.values, color=couleurs, linewidth=0)

    # Pourcentage au-dessus de chaque barre avec flèche
    total   = data.values.sum()
    max_val = data.values.max() if data.values.max() > 0 else 1
    ax.set_ylim(0, max_val * 1.22)
    for i, val in enumerate(data.values):
        if val > 0 and total > 0:
            pct = val / total * 100
            ax.annotate(
                f'{pct:.1f}%',
                xy=(i, val),
                xytext=(i, val + max_val * 0.06),
                ha='center', va='bottom', fontsize=8, color='black',
                annotation_clip=False,
                arrowprops=dict(arrowstyle='-', color='#BBBBBB', lw=0.6),
            )

    ax.set_xticks(range(n))
    ax.set_xticklabels(data.index, rotation=25, ha='right', fontsize=10)
    ax.set_ylabel("Nombre d'événements de déprédation")
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 16 — DÉPRÉDATION PAR ESPÈCE COMMERCIALISABLE
# ─────────────────────────────────────────────────────────────────────────────

def fig_depredation_par_espece(data):
    """
    Figure 16 : barres groupées (côte à côte) de déprédation par espèce et déprédateur.
    Une barre par déprédateur pour chaque espèce, couleurs par nom de déprédateur.
    """
    _style_global()
    n_esp   = len(data)
    n_dep   = len(data.columns)
    largeur = 0.75 / n_dep
    x       = np.arange(n_esp)

    fig, ax = plt.subplots(figsize=(max(9, n_esp * 0.9 + 2), 6))

    max_val = data.values.max() if data.values.max() > 0 else 1
    ax.set_ylim(0, max_val * 1.22)

    for i, col in enumerate(data.columns):
        vals    = data[col].values.astype(float)
        couleur = COULEURS_DEPREDATEURS.get(col, '#888888')
        offset  = (i - (n_dep - 1) / 2) * largeur
        barres  = ax.bar(x + offset, vals, largeur * 0.92,
                         color=couleur, label=col, linewidth=0)

        # Pourcentage au-dessus de chaque barre avec flèche
        for bar, val in zip(barres, vals):
            if val > 0:
                ax.annotate(
                    f'{int(val)}',
                    xy=(bar.get_x() + bar.get_width() / 2, val),
                    xytext=(bar.get_x() + bar.get_width() / 2, val + max_val * 0.05),
                    ha='center', va='bottom', fontsize=7, color='black',
                    annotation_clip=False,
                    arrowprops=dict(arrowstyle='-', color='#BBBBBB', lw=0.6),
                )

    ax.set_xticks(x)
    ax.set_xticklabels(data.index, rotation=30, ha='right', fontsize=9)
    ax.set_ylabel("Nombre d'événements de déprédation")
    ax.legend(frameon=False, loc='upper right')
    ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 17 — CPUE PAR ESPÈCE ET PAR CALÉE
# ─────────────────────────────────────────────────────────────────────────────

def fig_cpue(data, couleurs_especes=None):
    """
    Figure 17 : CPUE par espèce cible et par filage — points seulement (pas de traits).
    La colonne 'CPUE totale' est affichée en noir avec un marqueur carré distinct.
    """
    _style_global()
    fig, ax = plt.subplots(figsize=FIG_XLARGE)

    for col in data.columns:
        if col == 'CPUE totale':
            # CPUE totale : points carrés noirs, plus grands, bien visibles
            ax.scatter(data.index, data[col], marker='s', s=130,
                       color='#222222', label=col, zorder=5)
        else:
            couleur = (couleurs_especes or {}).get(col, '#4472C4')
            ax.scatter(data.index, data[col], marker='o', s=45,
                       color=couleur, label=col, zorder=4)

    ax.set_xlabel('Numéro de filage')
    ax.set_ylabel('CPUE (nb captures / 1000 hameçons)')
    ax.set_xticks(data.index)
    ax.legend(frameon=False, loc='upper right', fontsize=9,
              ncol=min(3, len(data.columns)))
    ax.yaxis.grid(True, alpha=0.3, linestyle='--')
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 18 — PROFIL DE PROFONDEUR TDR
# ─────────────────────────────────────────────────────────────────────────────

def fig_profil_tdr(data):
    """
    Figure 18 : boîte à moustache de la profondeur médiane de pêche par TDR.
    data = {tdr_id: array_profondeurs_medianes}
    L'axe Y est inversé : profondeur croissante vers le bas.
    """
    _style_global()
    tdr_ids  = list(data.keys())
    valeurs  = [data[tid] for tid in tdr_ids]
    labels_x = [f'SET {sid}' for sid in tdr_ids]

    couleur_tdr  = '#34495E'
    couleur_fill = '#A9BCC9'  # bleu-gris clair pour le remplissage des boîtes

    fig, ax = plt.subplots(figsize=FIG_LARGE)

    bp = ax.boxplot(
        valeurs,
        tick_labels=labels_x,  # 'labels' renommé en 'tick_labels' dans matplotlib 3.9+
        patch_artist=True,   # remplissage coloré
        notch=False,
        showfliers=True,
        flierprops=dict(marker='o', markerfacecolor=couleur_tdr,
                        markersize=4, linestyle='none', alpha=0.6),
    )

    for patch in bp['boxes']:
        patch.set_facecolor(couleur_fill)
        patch.set_edgecolor(couleur_tdr)
        patch.set_linewidth(1.2)
    for element in ['whiskers', 'caps']:
        for item in bp[element]:
            item.set_color(couleur_tdr)
            item.set_linewidth(1.2)
    for med in bp['medians']:
        med.set_color('#1A2E3A')
        med.set_linewidth(2)

    ax.set_ylabel('Profondeur de pêche (m)')
    ax.invert_yaxis()
    ax.yaxis.grid(True, alpha=0.3, linestyle='--')
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE 19 — CARTE DES POSITIONS DE FILAGE ET VIRAGE
# ─────────────────────────────────────────────────────────────────────────────

def fig_carte_filages(donnees):
    """
    Carte des positions GPS de filage et virage pour toute la marée.

    Fond de carte : ocean_indien_cotes.geojson (local, hors ligne, ~214 KB).
    ZEE           : tracées approximativement à la main (traits pointillés).
    Pas de cartopy — uniquement matplotlib + json.

    donnees : dict retourné par calc_carte_filages :
              { 'sets'         : liste de dicts (num, date, setting_start/end, hauling_start/end),
                'port_depart'  : {'nom', 'lat', 'lon'} ou None,
                'port_arrivee' : {'nom', 'lat', 'lon'} ou None }
    Coordonnées en (lat, lon) — tracées en (lon, lat) sur les axes.
    """
    import os
    import json
    import matplotlib.patches as mpatches
    from matplotlib.lines import Line2D

    _style_global()

    sets         = donnees['sets']
    port_depart  = donnees.get('port_depart')
    port_arrivee = donnees.get('port_arrivee')

    # ── 1. Étendue géographique — auto-zoom sur les données réelles ───────────
    all_lats, all_lons = [], []
    for s in sets:
        all_lats += [s['setting_start'][0], s['setting_end'][0],
                     s['hauling_start'][0],  s['hauling_end'][0]]
        all_lons += [s['setting_start'][1], s['setting_end'][1],
                     s['hauling_start'][1],  s['hauling_end'][1]]
    for port in (port_depart, port_arrivee):
        if port:
            all_lats.append(port['lat'])
            all_lons.append(port['lon'])

    lat_min = min(all_lats) if all_lats else -30
    lat_max = max(all_lats) if all_lats else   0
    lon_min = min(all_lons) if all_lons else  40
    lon_max = max(all_lons) if all_lons else  80
    marge   = max((lat_max - lat_min) * 0.20,
                  (lon_max - lon_min) * 0.20, 2.0)
    LON_MIN = lon_min - marge
    LON_MAX = lon_max + marge
    LAT_MIN = lat_min - marge
    LAT_MAX = lat_max + marge

    fig, ax = plt.subplots(figsize=(12, 10))
    ax.set_facecolor('#A8D5E8')   # bleu océan
    ax.set_xlim(LON_MIN, LON_MAX)
    ax.set_ylim(LAT_MIN, LAT_MAX)

    # ── 2. Fond de carte (côtes) depuis le GeoJSON local ─────────────────────
    geojson_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'data', 'ocean_indien_cotes.geojson'
    )
    if os.path.exists(geojson_path):
        with open(geojson_path, 'r', encoding='utf-8') as f:
            geojson = json.load(f)
        for feature in geojson['features']:
            geom = feature['geometry']
            polys = (geom['coordinates'] if geom['type'] == 'MultiPolygon'
                     else [geom['coordinates']])
            for polygon in polys:
                for ring in polygon:
                    xs = [p[0] for p in ring]
                    ys = [p[1] for p in ring]
                    if any(LON_MIN <= x <= LON_MAX and LAT_MIN <= y <= LAT_MAX
                           for x, y in zip(xs, ys)):
                        ax.fill(xs, ys, color='#EDE8DC', zorder=1)
                        ax.plot(xs, ys, color='#555555',
                                linewidth=0.4, zorder=2)

    # ── 3. ZEE officielles — GeoJSON local pré-découpé ───────────────────────
    # Fichier généré une fois depuis le shapefile Marine Regions World EEZ v12.
    # Contient uniquement les 33 ZEE de la zone Océan Indien.
    zee_trace = False
    geojson_zee = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'data', 'ocean_indien_zee.geojson'
    )
    if os.path.exists(geojson_zee):
        try:
            with open(geojson_zee, 'r', encoding='utf-8') as f:
                gj_zee = json.load(f)
            for feature in gj_zee['features']:
                geom = feature['geometry']
                if geom['type'] == 'MultiLineString':
                    lignes = geom['coordinates']
                elif geom['type'] == 'LineString':
                    lignes = [geom['coordinates']]
                else:
                    continue
                for ligne in lignes:
                    xs = [p[0] for p in ligne]
                    ys = [p[1] for p in ligne]
                    ax.plot(xs, ys, '--', color='#2E75B6',
                            linewidth=0.9, alpha=0.80, zorder=3)
            zee_trace = True
        except Exception:
            pass  # fichier absent ou corrompu → carte sans ZEE

    # ── 4. Polygone de la zone balayée + marqueurs par filage ─────────────────
    for s in sets:
        lat_sf, lon_sf = s['setting_start']   # début filage
        lat_se, lon_se = s['setting_end']     # fin filage
        lat_hf, lon_hf = s['hauling_start']  # début virage
        lat_he, lon_he = s['hauling_end']    # fin virage

        # Aire jaune orangé transparent — zone balayée par la palangre.
        # Les 4 points sont triés par angle depuis le centroïde pour éviter
        # un polygone croisé (auto-intersectant) qui ne se remplirait pas.
        pts_4 = [(lon_sf, lat_sf), (lon_se, lat_se),
                 (lon_hf, lat_hf), (lon_he, lat_he)]
        cx4 = sum(p[0] for p in pts_4) / 4
        cy4 = sum(p[1] for p in pts_4) / 4
        pts_4 = sorted(pts_4, key=lambda p: np.arctan2(p[1]-cy4, p[0]-cx4))
        ax.fill([p[0] for p in pts_4], [p[1] for p in pts_4],
                color='#F39C12', alpha=0.45, zorder=4)

        # Traits filage (rouge) et virage (orange)
        ax.plot([lon_sf, lon_se], [lat_sf, lat_se],
                '-', color='#C0392B', linewidth=1.5, zorder=5)
        ax.plot([lon_hf, lon_he], [lat_hf, lat_he],
                '-', color='#E67E22', linewidth=1.5, zorder=5)

        # Marqueurs filage : triangle début (▲), carré fin (■)
        ax.plot(lon_sf, lat_sf, '^', color='#C0392B', markersize=7,
                markeredgecolor='white', markeredgewidth=0.5, zorder=6)
        ax.plot(lon_se, lat_se, 's', color='#C0392B', markersize=6,
                markeredgecolor='white', markeredgewidth=0.5, zorder=6)

        # Marqueurs virage : croix début (✕), rond fin (●)
        ax.plot(lon_hf, lat_hf, 'x', color='#E67E22',
                markersize=7, markeredgewidth=1.8, zorder=6)
        ax.plot(lon_he, lat_he, 'o', color='#E67E22', markersize=6,
                markeredgecolor='white', markeredgewidth=0.5, zorder=6)

        # Numéro de set + date
        label = f"SET {s['num']}"
        if s.get('date'):
            label += f"\n{s['date']}"
        ax.text(lon_sf + 0.05, lat_sf + 0.05, label,
                fontsize=7, color='#C0392B', fontweight='bold',
                ha='left', va='bottom', zorder=10)

    # ── 5. Trajet navire : port départ → 1er filage  ET  dernier virage → port arrivée
    # (on ne trace PAS le trait entre le premier et le dernier set)
    if port_depart and sets:
        ax.plot([port_depart['lon'], sets[0]['setting_start'][1]],
                [port_depart['lat'], sets[0]['setting_start'][0]],
                '--', color='#888888', linewidth=1.2, alpha=0.55, zorder=4)
    if port_arrivee and sets:
        ax.plot([sets[-1]['hauling_end'][1], port_arrivee['lon']],
                [sets[-1]['hauling_end'][0], port_arrivee['lat']],
                '--', color='#888888', linewidth=1.2, alpha=0.55, zorder=4)

    # Port de départ : étoile noire pleine
    if port_depart:
        ax.plot(port_depart['lon'], port_depart['lat'], '*', color='black',
                markersize=13, markeredgecolor='white',
                markeredgewidth=0.5, zorder=8)
        ax.text(port_depart['lon'] + 0.10, port_depart['lat'] + 0.10,
                port_depart['nom'],
                fontsize=7.5, fontweight='bold', color='black',
                ha='left', va='bottom', zorder=9)

    # Port d'arrivée : cercle vide noir — si même port que le départ, encercle l'étoile
    if port_arrivee:
        ax.plot(port_arrivee['lon'], port_arrivee['lat'], 'o',
                color='none', markeredgecolor='black',
                markeredgewidth=1.8, markersize=13, zorder=8)
        # N'afficher le nom que si port d'arrivée différent du port de départ
        if not port_depart or (port_arrivee['nom'] != port_depart['nom']):
            ax.text(port_arrivee['lon'] + 0.10, port_arrivee['lat'] - 0.15,
                    port_arrivee['nom'],
                    fontsize=7.5, fontweight='bold', color='black',
                    ha='left', va='top', zorder=9)

    # ── 6. Grille et axes ─────────────────────────────────────────────────────
    ax.grid(True, alpha=0.30, linestyle='--', linewidth=0.5, zorder=2)
    ax.set_xlabel('Longitude', fontsize=9)
    ax.set_ylabel('Latitude', fontsize=9)
    ax.tick_params(labelsize=8)
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

    # ── 7. Rose des vents (coin supérieur gauche) ─────────────────────────────
    ax_rose = ax.inset_axes([0.89, 0.86, 0.09, 0.12])
    ax_rose.set_xlim(-1.6, 1.6)
    ax_rose.set_ylim(-1.6, 1.6)
    ax_rose.set_aspect('equal')
    ax_rose.axis('off')
    ax_rose.patch.set_facecolor('white')
    ax_rose.patch.set_alpha(0.85)
    for angle_deg, etiq in [(90,'N'), (270,'S'), (0,'E'), (180,'O')]:
        r = np.radians(angle_deg)
        dx, dy = np.cos(r), np.sin(r)
        ax_rose.annotate('', xy=(dx, dy), xytext=(0, 0),
                         arrowprops=dict(arrowstyle='->', color='#333333', lw=1.2))
        ax_rose.text(dx * 1.45, dy * 1.45, etiq,
                     ha='center', va='center', fontsize=8,
                     fontweight='bold', color='#333333')
    ax_rose.add_patch(plt.Circle((0, 0), 1.0, fill=False,
                                  color='#666666', linewidth=0.8))

    # ── 8. Échelle en milles nautiques (coin inférieur gauche) ────────────────
    # 1 degré de latitude = 60 milles nautiques exactement
    lat_centre = (LAT_MIN + LAT_MAX) / 2
    ecart_lon  = LON_MAX - LON_MIN
    echelle_nm = 100  # fixé à 100 mn
    echelle_deg_lon = (echelle_nm / 60.0) / np.cos(np.radians(lat_centre))

    x0 = LON_MIN + 0.03 * (LON_MAX - LON_MIN)
    y0 = LAT_MIN + 0.03 * (LAT_MAX - LAT_MIN)
    tick_h = 0.012 * (LAT_MAX - LAT_MIN)

    ax.plot([x0, x0 + echelle_deg_lon], [y0, y0],
            '-', color='#555555', linewidth=1.2, zorder=9,
            solid_capstyle='butt')
    for x_tick in (x0, x0 + echelle_deg_lon):
        ax.plot([x_tick, x_tick], [y0 - tick_h, y0 + tick_h],
                '-', color='#555555', linewidth=0.9, zorder=9)
    ax.text(x0 + echelle_deg_lon / 2, y0 + tick_h * 1.4,
            '100 mn',
            ha='center', va='bottom', fontsize=7, color='#555555', zorder=9)

    # ── 9. Légende ───────────────────────────────────────────────────────────
    elements_legende = [
        Line2D([0],[0], marker='^', color='w', markerfacecolor='#C0392B',
               markersize=8, label='Début filage'),
        Line2D([0],[0], marker='s', color='w', markerfacecolor='#C0392B',
               markersize=7, label='Fin filage'),
        Line2D([0],[0], color='#C0392B', linewidth=1.5, label='Ligne de filage'),
        Line2D([0],[0], marker='x', color='#E67E22', markersize=8,
               markeredgewidth=1.5, label='Début virage', linestyle='none'),
        Line2D([0],[0], marker='o', color='w', markerfacecolor='#E67E22',
               markersize=7, label='Fin virage'),
        Line2D([0],[0], color='#E67E22', linewidth=1.5, label='Ligne de virage'),
        mpatches.Patch(facecolor='#F39C12', alpha=0.45, label='Zone balayée'),
        Line2D([0],[0], color='black', linewidth=1.2, linestyle='--',
               label='Trajet navire'),
    ]
    if port_depart:
        elements_legende.append(
            Line2D([0],[0], marker='*', color='w',
                   markerfacecolor='black', markersize=10, label='Port de départ')
        )
    if port_arrivee:
        elements_legende.append(
            Line2D([0],[0], marker='o', color='none',
                   markerfacecolor='none', markeredgecolor='black',
                   markeredgewidth=1.8, markersize=9, label="Port d'arrivée")
        )
    if zee_trace:
        elements_legende.append(
            Line2D([0],[0], color='#2E75B6', linewidth=0.9, linestyle='--',
                   label='Limite ZEE')
        )
    ax.legend(handles=elements_legende, loc='lower right',
              fontsize=7, framealpha=0.92, ncol=2)

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# TABLEAUX ENGIN — DESCRIPTION DE L'ENGIN DE PÊCHE
# ─────────────────────────────────────────────────────────────────────────────

def fig_tableau_engin(df, titre, figsize=None):
    """
    Génère une figure PNG affichant un DataFrame sous forme de tableau formaté.
    Style rapport IRD : titre bleu foncé, en-têtes gris foncé, alternance blanc/gris clair.

    df      : DataFrame à afficher (nombre de colonnes quelconque)
    titre   : titre affiché au-dessus du tableau
    figsize : tuple (largeur, hauteur) en pouces — calculé automatiquement si None
    """
    _style_global()

    n_rows, n_cols = len(df), len(df.columns)

    # Largeurs des colonnes proportionnelles au contenu le plus long de chaque colonne
    col_lens = []
    for col in df.columns:
        max_len = len(str(col))
        for val in df[col]:
            max_len = max(max_len, len(str(val)))
        col_lens.append(max(max_len, 3))
    total_len = sum(col_lens)
    col_widths = [l / total_len for l in col_lens]

    # Taille de la figure : largeur adaptée au nombre de colonnes, hauteur aux lignes
    if figsize is None:
        fig_w = max(8, n_cols * 1.6)
        fig_h = max(2.5, n_rows * 0.50 + 1.6)
        figsize = (fig_w, fig_h)

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_axis_off()

    col_labels = list(df.columns)
    # Toutes les valeurs sont converties en chaîne pour l'affichage
    cell_data  = [[str(v) for v in row] for _, row in df.iterrows()]

    table = ax.table(
        cellText=cell_data,
        colLabels=col_labels,
        colWidths=col_widths,
        loc='center',
        cellLoc='left',
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)

    COULEUR_TITRE   = '#1F4E79'  # bleu IRD foncé
    COULEUR_HEADER  = '#555555'  # gris foncé pour les en-têtes de colonne
    COULEUR_PAIR    = '#F0F0F0'  # gris clair pour les lignes paires
    COULEUR_IMPAIR  = '#FFFFFF'  # blanc pour les lignes impaires
    COULEUR_BORD    = '#D0D0D0'  # bordure gris clair

    hauteur_ligne = 0.06  # hauteur de chaque ligne en fraction normalisée de l'axe

    # Style des en-têtes de colonnes (ligne 0 dans la table matplotlib)
    for j in range(n_cols):
        cell = table[0, j]
        cell.set_facecolor(COULEUR_HEADER)
        cell.set_text_props(color='white', fontweight='bold')
        cell.set_edgecolor(COULEUR_BORD)
        cell.set_linewidth(0.4)
        cell.set_height(hauteur_ligne)

    # Style des lignes de données (lignes 1..n_rows)
    for i in range(n_rows):
        bg = COULEUR_PAIR if i % 2 == 0 else COULEUR_IMPAIR
        for j in range(n_cols):
            cell = table[i + 1, j]
            cell.set_facecolor(bg)
            cell.set_edgecolor(COULEUR_BORD)
            cell.set_linewidth(0.4)
            cell.set_height(hauteur_ligne)
            # Première colonne en gras pour les tableaux Paramètre/Valeur
            if j == 0 and n_cols == 2:
                cell.set_text_props(fontweight='bold')

    # Titre au-dessus du tableau avec fond bleu IRD
    ax.set_title(
        titre,
        fontsize=10,
        fontweight='bold',
        color='white',
        loc='left',
        pad=8,
        bbox=dict(
            boxstyle='square,pad=0.5',
            facecolor=COULEUR_TITRE,
            edgecolor='none',
        ),
    )

    plt.tight_layout()
    return fig
