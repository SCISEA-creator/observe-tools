#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script utilitaire — Génération du fichier Excel des règles de correction mis à jour
IRD OB7 - SCISEA

Lancer : python phase1_detection/generer_regles_xlsx.py
"""

import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── Styles ───────────────────────────────────────────────────────────────────

fill_en_tete = PatternFill('solid', fgColor='FF1F4E79')
fill_erreur  = PatternFill('solid', fgColor='FFFFD7D7')
fill_avert   = PatternFill('solid', fgColor='FFFFF3CD')

font_en_tete = Font(bold=True, color='FFFFFFFF', name='Calibri', size=11)
font_table   = Font(bold=True, color='FF1F4E79', name='Calibri', size=10)
font_normal  = Font(name='Calibri', size=10)
font_bold    = Font(bold=True, name='Calibri', size=10)

align_wrap   = Alignment(wrap_text=True, vertical='top')
align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)


def bord():
    s = Side(style='thin', color='FFAAAAAA')
    return Border(left=s, right=s, top=s, bottom=s)


# ─── Données des règles ───────────────────────────────────────────────────────
# Format : (table, nom_regle, description, colonnes_champs, niveau, champs_optionnels)
# niveau : 1 = Erreur, 2 = Avertissement, None = ligne vide séparatrice

REGLES = [

    # ── 1_trip ──────────────────────────────────────────────────────────────
    ('1_trip',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid (identifiants internes ObServe) sont toujours ignorées.',
     'Toutes colonnes',
     1,
     'trip_comment'),
    (None, None, None, None, None, None),

    # ── 2_set ───────────────────────────────────────────────────────────────
    ('2_set',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid et les champs optionnels sont ignorés.',
     'Toutes colonnes',
     1,
     'shooter_speed; swivel_weight; snap_weight; max_depth_targeted; time_between_hooks; set_comment'),

    ('2_set',
     'swivel_weight conditionnel',
     'Si weighted_swivel = vrai, alors swivel_weight est obligatoire. '
     'Si weighted_swivel est faux ou vide, swivel_weight peut rester vide.',
     'swivel_weight; weighted_swivel',
     1,
     ''),

    ('2_set',
     'snap_weight conditionnel',
     'Si weighted_snap = vrai, alors snap_weight est obligatoire. '
     'Si weighted_snap est faux ou vide, snap_weight peut rester vide.',
     'snap_weight; weighted_snap',
     1,
     ''),

    ('2_set',
     'Erreur frappe',
     'Si une valeur est clairement dominante dans une colonne (plus de 50% des lignes), '
     'toute autre valeur présente est signalée comme suspecte — même si elle apparait '
     'plusieurs fois (cas de fautes de frappe répétées sur plusieurs calées). '
     'Le message indique la valeur suspecte ET la valeur attendue.',
     'trip_id; shooter_used; weighted_swivel; swivel_weight; '
     'lightsticks_per_basket_count; branchlines_per_basket_count; '
     'baskets_per_section_count; total_sections_count; '
     'hauling_direction_same_as_setting; hauling_breaks; time_between_hooks; '
     'weighted_snap; snap_weight; mainline_type; lightstick_type; '
     'lightstick_color; line_instrumented',
     2,
     ''),

    ('2_set',
     'set_id séquentiel',
     "Les numéros de calée (partie numérique après le '-') doivent se suivre "
     "sans trou ni doublon, en commençant à 1. "
     "Ex: SEALORXXX-1, SEALORXXX-2, SEALORXXX-3...",
     'set_id',
     1,
     ''),

    ('2_set',
     'set_number séquentiel',
     'Les numéros de calée doivent se suivre sans trou ni doublon, '
     'en commençant à 1. Ex: 1, 2, 3, 4...',
     'set_number',
     1,
     ''),

    ('2_set',
     'setting_speed',
     'La vitesse de filage doit être entre 5,5 et 8,5 nœuds. '
     'Une valeur hors de cette plage est inhabituelle mais pas nécessairement fausse.',
     'setting_speed',
     2,
     ''),

    ('2_set',
     'snap_weight plage',
     'Le poids du snap doit être entre 0,05 et 0,12 kg. '
     'Vérifié uniquement si la valeur est remplie.',
     'snap_weight',
     1,
     ''),

    ('2_set',
     'total_baskets_count',
     'Le nombre total de paniers doit être entre 1 et 300.',
     'total_baskets_count',
     1,
     ''),

    ('2_set',
     'total_hooks_count',
     'Le nombre total d\'hameçons doit être entre 1 et 2500.',
     'total_hooks_count',
     1,
     ''),

    ('2_set',
     'total_hooks_count cohérence',
     'total_hooks_count doit être cohérent avec branchlines_per_basket_count × total_baskets_count. '
     'Si total_hooks > max théorique : Erreur (impossible physiquement). '
     'Si l\'écart dépasse 2 baskets complets : Avertissement (les derniers baskets peuvent être partiels). '
     'Ex : 20 avançons × 15 baskets = 300 max ; total_hooks = 310 → Erreur ; total_hooks = 240 → OK.',
     'total_hooks_count; branchlines_per_basket_count; total_baskets_count',
     1,
     ''),

    ('2_set',
     'set_id format et cohérence',
     'Le set_id doit respecter le format [trip_id]-[set_number] sans zéro devant le numéro. '
     'Exemple correct : SEALORXXX-3 pour le filage 3 de la marée SEALORXXX. '
     'Exemple incorrect : SEALORXXX-03 (zéro en tête) ou SEALORYYY-3 (mauvaise marée). '
     'La partie numérique après le tiret doit correspondre exactement à set_number.',
     'set_id; set_number',
     1,
     ''),

    ('2_set',
     'sst',
     'La température de surface de la mer doit être entre 20 et 35 °C. '
     'Une valeur hors plage est inhabituelle mais pas nécessairement fausse.',
     'sst',
     2,
     ''),

    ('2_set',
     'branchline_length_m',
     'La longueur d\'avançon doit être entre 10 et 20 m. '
     'Une valeur hors de cette plage est inhabituelle et peut indiquer une faute de frappe.',
     'branchline_length_m',
     2,
     ''),

    ('2_set',
     'branchlines_per_basket_count',
     'Le nombre d\'avançons par basket doit être entre 6 et 25. '
     'Une valeur hors de cette plage est inhabituelle et peut indiquer une faute de frappe.',
     'branchlines_per_basket_count',
     2,
     ''),

    ('2_set',
     'hauling_speed',
     'La vitesse de virage doit être entre 5,5 et 8 nœuds. '
     'Une valeur hors de cette plage est inhabituelle mais pas nécessairement fausse.',
     'hauling_speed',
     2,
     ''),

    (None, None, None, None, None, None),

    # ── 3_floatline ─────────────────────────────────────────────────────────
    ('3_floatline',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid sont ignorées.',
     'Toutes colonnes',
     1,
     ''),

    ('3_floatline',
     'floatline_length_m (brasse)',
     'La longueur doit être un multiple de 1,8288 m (1 brasse). '
     'Les longueurs sont saisies en brasses dans ObServe et converties en mètres '
     'à l\'export. Tolérance : 0,001 m pour les arrondis. '
     'Valeurs valides : 1,8288 / 3,6576 / 5,4864 / 7,3152 / 9,144...',
     'floatline_length_m',
     1,
     ''),
    (None, None, None, None, None, None),

    # ── 4_branchline ────────────────────────────────────────────────────────
    ('4_branchline',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid sont ignorées.',
     'Toutes colonnes',
     1,
     ''),

    ('4_branchline',
     'Erreur frappe',
     'Même logique que 2_set : si une valeur est clairement dominante (>50%), '
     'toute autre valeur est signalée comme suspecte avec indication de la valeur attendue.',
     'branchline_proportion; branchline_length_m; '
     'branchline_topline_type; branchline_traceline_type',
     2,
     ''),

    ('4_branchline',
     'branchline_length_m (brasse)',
     'La longueur doit être un multiple de 1,8288 m (1 brasse). '
     'Même logique que floatline_length_m.',
     'branchline_length_m',
     1,
     ''),
    (None, None, None, None, None, None),

    # ── 5_hook ──────────────────────────────────────────────────────────────
    ('5_hook',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid sont ignorées.',
     'Toutes colonnes',
     1,
     ''),
    (None, None, None, None, None, None),

    # ── 6_bait ──────────────────────────────────────────────────────────────
    ('6_bait',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid et les champs optionnels sont ignorés.',
     'Toutes colonnes',
     1,
     'bait_weight'),

    ('6_bait',
     'bait_code',
     'Le code appât doit être SQU (calmar) ou MAZ (maquereau). '
     'Toute autre valeur est inhabituelle et signalée en avertissement.',
     'bait_code',
     2,
     ''),

    ('6_bait',
     'bait_setting_status',
     'Le statut de l\'appât au filage doit être DEF (par défaut). '
     'Toute autre valeur est inhabituelle et signalée en avertissement.',
     'bait_setting_status',
     2,
     ''),

    ('6_bait',
     'bait_size',
     'La taille de l\'appât doit être entre 10 et 35 cm.',
     'bait_size',
     1,
     ''),
    (None, None, None, None, None, None),

    # ── 7_tdr ───────────────────────────────────────────────────────────────
    ('7_tdr',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid et les champs jamais remplis par les observateurs sont ignorés.',
     'Toutes colonnes',
     1,
     'section_setting_id; section_hauling_id; basket_setting_id; '
     'basket_hauling_id; data_location; data_file_name; fao_code'),

    ('7_tdr',
     'TDR manquant par calée',
     'Chaque calée de la marée doit avoir au moins un TDR associé dans la table. '
     'Si un numéro de calée est absent de 7_tdr, une anomalie est signalée.',
     'set_number',
     2,
     ''),

    ('7_tdr',
     'Nombre de TDR par calée',
     'Le nombre de lignes TDR doit être identique pour toutes les calées. '
     'Si une calée a un nombre différent des autres, elle est signalée.',
     'set_number',
     2,
     ''),

    ('7_tdr',
     'tdr_id vide',
     "L'identifiant TDR (tdr_id) doit toujours être renseigné. "
     'Un tdr_id vide signale probablement un oubli de saisie à corriger dans ObServe.',
     'tdr_id',
     2,
     ''),

    ('7_tdr',
     'tdr_id homogène',
     'Les mêmes TDR (mêmes valeurs de tdr_id) doivent être déployés pour toutes les calées. '
     'Si une calée a un ensemble de tdr_id différent des autres, une anomalie est signalée. '
     'Note : une panne ou perte de TDR en cours de marée est possible — Avertissement seulement.',
     'tdr_id',
     2,
     ''),

    ('7_tdr',
     'tdr_serial cohérent',
     'Pour un même tdr_id, le numéro de série (tdr_serial) doit être identique '
     'sur toute la marée. Une incohérence signale une confusion de capteur ou une erreur de saisie.',
     'tdr_serial',
     1,
     ''),

    ('7_tdr',
     'Profondeur TDR',
     'Les colonnes de profondeur doivent être entre 0 et 400 m. '
     'Vérification effectuée uniquement si la valeur est remplie.',
     'fishing_start_depth; fishing_end_depth; mean_deployment_depth; '
     'median_deployment_depth; min_fishing_depth; max_fishing_depth; '
     'mean_fishing_depth; median_fishing_depth',
     1,
     ''),

    ('7_tdr',
     'item_horizontal_position',
     'Doit être "Basket". Toute autre valeur indique une erreur de saisie '
     'de la position horizontale du TDR sur la ligne.',
     'item_horizontal_position',
     1,
     ''),

    ('7_tdr',
     'item_vertical_position',
     'Doit être "Hook" (anciens exports) ou "Basse" (exports CSV récents). '
     'Toute autre valeur indique une erreur de saisie '
     'de la position verticale du TDR.',
     'item_vertical_position',
     1,
     ''),
    (None, None, None, None, None, None),

    # ── 8_catch ─────────────────────────────────────────────────────────────
    ('8_catch',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid et les champs optionnels sont ignorés.',
     'Toutes colonnes',
     1,
     'total_weight; status_at_release; hook_when_discarded; '
     'depredating_species; catch_comment'),

    ('8_catch',
     'fao_code',
     'Le code espèce doit être dans la liste de référence. '
     'Toute valeur hors liste est signalée en avertissement (pas nécessairement fausse).',
     'YFT; BET; ALB; SKJ; SWO; BUM; BLM; MLS; SSP; SFA; FAL; OCS; TIG; SMA; '
     'BSH; BTH; ISB; WAH; RZV; MOP; DOL; LEC; OIL; GES; ALX; GBA; LGH; PLS; '
     'TUG; LKV; TTH; TTL; DKK',
     2,
     ''),

    ('8_catch',
     'catch_fate RET autorisé',
     'Seules les espèces suivantes peuvent être retenues (RET). '
     'Pour toute autre espèce, un catch_fate = RET est une erreur.',
     'YFT; BET; ALB; SWO; BUM; BLM; MLS; SSP; SFA; WAH; DOL',
     1,
     ''),

    ('8_catch',
     'DIS + status_at_release',
     'Si catch_fate = DIS (rejeté), le champ status_at_release est obligatoire. '
     'Il indique le statut de l\'animal au moment du rejet (vivant, mort...).',
     'catch_fate; status_at_release',
     1,
     ''),

    ('8_catch',
     'Capture morte + status_at_release',
     'Si l\'animal est capturé mort (status_at_capture = D), '
     'son statut au relâcher (status_at_release) ne peut pas être différent de D. '
     'Un animal capturé mort ne peut pas être relâché vivant.',
     'status_at_capture; status_at_release',
     1,
     ''),

    ('8_catch',
     'DIS + hook_when_discarded',
     'Si catch_fate = DIS (rejeté), le champ hook_when_discarded est obligatoire. '
     'Il indique si l\'hameçon était encore dans l\'animal au moment du rejet.',
     'hook_when_discarded; catch_fate',
     1,
     ''),

    ('8_catch',
     'Mort + ESC impossible',
     'Si status_at_capture = D (mort) ET catch_fate = ESC (échappé), '
     'la combinaison est biologiquement impossible : un animal mort ne peut pas s\'échapper. '
     'Signale une erreur de saisie du devenir ou du statut à la capture.',
     'status_at_capture; catch_fate',
     1,
     ''),

    ('8_catch',
     'Doublon captures consécutives',
     'Si deux captures consécutives (catch_number N et N+1, même calée) ont exactement '
     'les mêmes valeurs pour fao_code, catch_fate, status_at_capture, status_at_release '
     'et depredating_species, ET la même taille mesurée dans 9_sample (même size + size_code), '
     'c\'est probablement un doublon de saisie. '
     'Si l\'une des deux captures n\'a pas de mensuration dans 9_sample → ignoré.',
     'catch_number; fao_code; catch_fate; status_at_capture; status_at_release; depredating_species',
     2,
     ''),

    ('8_catch',
     'Numéros de capture non consécutifs',
     'Pour chaque calée, les catch_number doivent se suivre sans saut ni répétition. '
     'Cas 1 — Saut : les numéros présents ne se suivent pas (ex: 1, 2, 4 → le n°3 manque). '
     'Cas 2 — Mauvais ordre : les numéros existent tous mais dans le désordre (ex: 1, 2, 5, 3). '
     'Cas 3 — Doublon : le même numéro de capture est utilisé deux fois dans la même calée (ex: 1, 2, 2, 3). '
     'La séquence n\'est pas vérifiée si elle commence à un numéro différent de 1.',
     'catch_number',
     2,
     ''),

    ('8_catch',
     'Mensuration manquante dans 9_sample',
     'Pour chaque capture remontée à bord (catch_fate ≠ ESC), une mensuration doit exister dans 9_sample. '
     'L\'absence de mensuration peut indiquer un oubli de saisie de la taille par l\'observateur. '
     'Exceptions — pas d\'avertissement si : '
     '(1) catch_fate = ESC (poisson non remonté à bord, impossible à mesurer) ; '
     '(2) déprédation ODN, SKH, FAW ou GLO, seule ou combinée à un autre prédateur '
     '(ex : ISB;ODN) — poisson trop endommagé pour être mesuré ; '
     '(3) espèce "hors gabarit" : requins, raies (sauf PLS), cétacés, requin baleine '
     '(trop grandes pour être remontées à bord et mesurées). '
     'PLS, tortues et oiseaux marins ne sont PAS exclus : ces espèces sont assez petites pour être mesurées.',
     'catch_number',
     2,
     ''),

    ('8_catch',
     'Rejet suspect espèce commercialisable',
     'Si une espèce cible commercialisable (YFT, BET, ALB, SWO, BUM, BLM, MLS, SSP, SFA, WAH, DOL) '
     'est rejetée (DIS ou ESC) SANS déprédation signalée ET qu\'une mensuration est présente dans '
     '9_sample à une taille dans la plage commerciale normale, c\'est suspect : '
     'le poisson aurait probablement dû être retenu (RET). '
     'Si aucune mensuration n\'existe → le poisson n\'a pas été remonté à bord → ignoré.',
     'catch_fate; depredating_species',
     2,
     ''),
    (None, None, None, None, None, None),

    # ── 9_sample ────────────────────────────────────────────────────────────
    ('9_sample',
     'Champs vides',
     'Toutes les colonnes obligatoires doivent être remplies. '
     'Les colonnes topiaid et les champs optionnels sont ignorés.',
     'Toutes colonnes',
     1,
     'weight; weight_code; sex; maturity; catch_comment'),

    ('9_sample',
     'size hors limites',
     'La taille saisie doit être dans les limites min/max définies pour l\'espèce '
     'dans la table Capture measurement. '
     'La vérification est faite selon le fao_code de chaque ligne. '
     'Si l\'espèce n\'est pas dans la table de référence, la vérification est ignorée.',
     'size',
     2,
     ''),

    ('9_sample',
     'size_code incorrect',
     'Le code de mesure doit correspondre au code attendu pour l\'espèce '
     'dans la table Capture measurement (ex: FL pour YFT, LJFL pour SWO, SCL ou SCW pour TUG). '
     'La vérification est faite selon le fao_code de chaque ligne.',
     'size_code',
     1,
     ''),

    ('9_sample',
     'Doublon taille consécutive',
     'Si deux mensurations consécutives (catch_number N et N+1, même calée) ont exactement '
     'la même valeur de size ET le même size_code, c\'est probablement une copie par erreur '
     'de la taille du poisson précédent (comportement courant dans l\'interface ObServe). '
     'Aucune vérification du fao_code : un YFT et un BET avec la même taille consécutive '
     'sont aussi suspects.',
     'size; size_code',
     2,
     ''),
]


# ─── Création du fichier Excel ────────────────────────────────────────────────

def creer_fichier(chemin_sortie):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Regles de correction'

    # En-tête
    headers = [
        'Table',
        'Nom règle',
        'Description (logique exacte implémentée)',
        'Colonnes / Champs concernés',
        'Sévérité',
        'Champs optionnels exclus du contrôle vide',
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill      = fill_en_tete
        cell.font      = font_en_tete
        cell.alignment = align_center
        cell.border    = bord()

    # Largeurs
    for i, l in enumerate([14, 28, 72, 55, 14, 52], 1):
        ws.column_dimensions[get_column_letter(i)].width = l
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = 'A2'

    # Lignes de données
    excel_row = 2
    for (table, nom, desc, cols, niv, optionnels) in REGLES:
        if table is None:
            excel_row += 1
            continue

        fill_row = fill_erreur if niv == 1 else fill_avert
        niv_txt  = '1 - Erreur' if niv == 1 else '2 - Avertissement'

        vals = [table, nom, desc, cols, niv_txt, optionnels]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=excel_row, column=col, value=val or '')
            cell.fill      = fill_row
            cell.alignment = align_wrap
            cell.border    = bord()
            cell.font      = font_table if col == 1 else font_normal
            if col == 5:
                cell.alignment = align_center

        excel_row += 1

    # Légende
    excel_row += 1
    ws.cell(row=excel_row, column=1, value='LÉGENDE').font = font_bold
    excel_row += 1

    c = ws.cell(row=excel_row, column=1, value='1 - Erreur')
    c.fill = fill_erreur
    c.font = font_normal
    ws.cell(row=excel_row, column=2,
            value='Anomalie certaine : la valeur est clairement incorrecte selon les règles.')
    excel_row += 1

    c = ws.cell(row=excel_row, column=1, value='2 - Avertissement')
    c.fill = fill_avert
    c.font = font_normal
    ws.cell(row=excel_row, column=2,
            value='Valeur inhabituelle ou suspecte : à vérifier, peut être volontaire.')

    # Fichier informatif : protégé contre les modifications accidentelles des observateurs
    ws.protection.sheet    = True
    ws.protection.password = 'DOL2026'

    wb.save(chemin_sortie)
    print('Fichier sauvegardé :', chemin_sortie)


# ─── Point d'entrée ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    dossier = os.path.dirname(os.path.abspath(__file__))
    chemin = os.path.join(
        dossier,
        'Règles de corrections',
        'Règle correction Observe-tool.xlsx'
    )
    creer_fichier(chemin)
