#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Observe-Tools — Anonymisation du template du rapport Word.
IRD OB7 - SCISEA

Le template est fabriqué (prepare_template.py) à partir d'un vrai rapport de
marée. Il garde donc des informations issues des données de cette marée :
caractéristiques du navire, équipage, exemples rédigés, photos, anciennes
figures, métadonnées. Ces données appartiennent au programme d'observation :
on les remplace par des repères neutres « [XX] » en orange, à compléter par
l'observateur.

Principe de sécurité : on ne modifie QUE le texte de morceaux (runs) déjà
existants et le contenu binaire des images. La structure du document, les
balises Jinja2 ({{ … }}, {% … %}) et les contrôles photo restent intacts.
Le script ne contient volontairement aucun nom réel : il repère les zones
par leur position (texte orange, chiffres, lignes de l'équipage).

Usage : python anonymiser_template.py <entree.docx> [<sortie.docx>]
"""

import io
import os
import re
import sys
import zipfile

import docx
from PIL import Image, ImageDraw, ImageFont

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ORANGE = 'FF8000'


# ── Outils sur les morceaux de texte (runs) ──────────────────────────────────

def _texte(el):
    return ''.join(t.text or '' for t in el.iter(W + 't'))


def _runs(p):
    """Runs directs du paragraphe qui portent du texte (hors champs SEQ imbriqués)."""
    return [r for r in p.iter(W + 'r') if r.find(W + 't') is not None]


def _ecrire_run(r, texte, orange=False):
    """Remplace le texte d'un run (garde sa mise en forme) ; option : le passer en orange."""
    ts = r.findall(W + 't')
    ts[0].text = texte
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for t in ts[1:]:
        r.remove(t)
    if orange:
        rpr = r.find(W + 'rPr')
        if rpr is None:
            rpr = r.makeelement(W + 'rPr', {})
            r.insert(0, rpr)
        col = rpr.find(W + 'color')
        if col is None:
            col = rpr.makeelement(W + 'color', {})
            rpr.append(col)
        col.set(W + 'val', ORANGE)


def _est_orange(r):
    c = r.find(f'{W}rPr/{W}color')
    return c is not None and c.get(W + 'val', '').upper() == ORANGE


# ── Anonymisation du texte ───────────────────────────────────────────────────

def anonymiser_textes(document):
    """Retourne la liste des modifications faites (pour le contrôle)."""
    journal = []
    paras = list(document.element.body.iter(W + 'p'))

    for i, p in enumerate(paras):
        t = _texte(p)

        # §1.1 — caractéristiques du navire : les 4 valeurs orange, dans l'ordre
        # (longueur, armement, modèle de treuil, capacité du treuil)
        if '{{ navire }}' in t and 'palangrier de' in t:
            oranges = [r for r in _runs(p) if _est_orange(r) and _texte(r).strip()]
            reperes = ['[XX] m', '[armement]', '[modèle de treuil]', '[XX]']
            if len(oranges) == len(reperes):
                for r, rep in zip(oranges, reperes):
                    _ecrire_run(r, rep)
                journal.append('§1.1 caractéristiques du navire')

        # §1.2 — nombre de marins, puis chaque ligne « Fonction : Nom (nationalité) »
        elif 'composé de' in t and 'marins' in t:
            for r in _runs(p):
                if _texte(r).strip().isdigit():
                    _ecrire_run(r, '[X]', orange=True)
            journal.append('§1.2 nombre de marins')
            nb_lignes = 0
            for q in paras[i + 1:]:
                tq = _texte(q)
                if not tq.strip() and nb_lignes == 0:
                    continue                                # ligne vide avant la liste
                if ':' not in tq or len(tq) > 120:
                    break                                   # fin de la liste de l'équipage
                nb_lignes += 1
                runs = _runs(q)
                k = next(j for j, r in enumerate(runs) if ':' in _texte(r))
                suite = runs[k + 1:]
                jinja = [j for j, r in enumerate(suite) if '{{' in _texte(r)]
                if jinja:                                   # capitaine : nom = {{ capitaine }}
                    suite = suite[jinja[-1] + 1:]
                    repere = ' ([nationalité])'
                else:
                    # pas de double espace si le « : » est déjà suivi d'un espace
                    espace = '' if _texte(runs[k])[-1:].isspace() else ' '
                    repere = espace + '[Nom Prénom] ([nationalité])'
                if suite:
                    _ecrire_run(suite[0], repere, orange=True)
                    for r in suite[1:]:
                        _ecrire_run(r, '')
                else:                                       # tout était dans le run du « : »
                    debut = _texte(runs[k]).split(':', 1)[0]
                    _ecrire_run(runs[k], debut + ':' + repere, orange=True)
                journal.append('§1.2 équipage : ' + _texte(q).split(':')[0].strip())

        # Exemple rédigé (espèce protégée) : phrase « Une … a été capturée lors du SETn… »
        elif t.startswith('Une ') and 'a été capturée' in t and '[À compléter' in t:
            runs = _runs(p)
            if runs and not _est_orange(runs[0]):
                _ecrire_run(runs[0], 'Une [espèce] a été capturée lors du SET [N]. '
                                     '[Décrire son état à la capture et au relâcher.] ')
                journal.append('exemple espèce protégée')

        # Légendes « Photo N : <espèce de l'exemple> » (et leurs copies dans la table des photos).
        # Le numéro N est un champ automatique : on ne touche qu'au run qui contient l'espèce.
        elif t.startswith('Photo') and 'Tortue olivâtre' in t:
            runs = _runs(p)
            for j, r in enumerate(runs):
                tr = _texte(r)
                if 'Tortue olivâtre' in tr:
                    _ecrire_run(r, tr.replace('Tortue olivâtre', '[espèce]'))
                elif tr.endswith('Tortue ') and j + 1 < len(runs) and _texte(runs[j + 1]).startswith('olivâtre'):
                    # mot coupé sur deux runs : « Tortue » | « olivâtre »
                    _ecrire_run(r, tr[:-len('Tortue ')] + '[espèce]')
                    _ecrire_run(runs[j + 1], _texte(runs[j + 1])[len('olivâtre'):])
            journal.append('légende photo exemple')

        # Légende du profil TDR spécifique (et sa copie dans la table des figures)
        elif 'Profil de profondeur du TDR' in t:
            runs = _runs(p)
            for j, r in enumerate(runs):
                tr = _texte(r)
                precedent = _texte(runs[j - 1]) if j else ''
                if tr.strip().isdigit() and precedent.endswith('TDR'):
                    _ecrire_run(r, ' [N]')
                elif tr.strip().isdigit() and precedent.endswith('SET '):
                    _ecrire_run(r, '[N]')
                elif 'TDR' in tr or 'SET' in tr or 'olivâtre' in tr:
                    nouveau = re.sub(r'TDR\s?\d+', 'TDR [N]', tr)
                    nouveau = re.sub(r'SET\s?\d+', 'SET [N]', nouveau)
                    nouveau = nouveau.replace('d’une tortue olivâtre', 'd’un(e) [espèce]')
                    if nouveau != tr:
                        _ecrire_run(r, nouveau)
            journal.append('légende profil TDR')
    return journal


# ── Images de remplacement ───────────────────────────────────────────────────

def _image_neutre(largeur, hauteur, fmt):
    """Image grise « Photo à insérer » aux mêmes proportions que l'originale."""
    echelle = min(1.0, 800 / max(largeur, hauteur))
    l, h = max(1, int(largeur * echelle)), max(1, int(hauteur * echelle))
    im = Image.new('RGB', (l, h), (229, 231, 235))
    dessin = ImageDraw.Draw(im)
    texte = 'Photo à insérer'
    try:
        police = ImageFont.truetype('arial.ttf', max(12, l // 14))
    except OSError:
        police = ImageFont.load_default()
    x0, y0, x1, y1 = dessin.textbbox((0, 0), texte, font=police)
    dessin.text(((l - (x1 - x0)) / 2, (h - (y1 - y0)) / 2), texte, fill=(107, 114, 128), font=police)
    sortie = io.BytesIO()
    im.save(sortie, 'JPEG' if fmt == 'jpeg' else 'PNG', quality=80)
    return sortie.getvalue()


def _medias_du_corps(z):
    """Images liées au corps du document (les logos de l'en-tête sont exclus)."""
    rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    return {'word/' + c for c in re.findall(r'Target="(media/[^"]+)"', rels)}


def remplacer_images_et_metadonnees(chemin_in, chemin_out):
    """Réécrit le paquet zip : images du corps neutralisées, métadonnées nettoyées."""
    with zipfile.ZipFile(chemin_in) as zin:
        a_remplacer = _medias_du_corps(zin)
        with zipfile.ZipFile(chemin_out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in a_remplacer:
                    fmt = 'jpeg' if item.filename.lower().endswith(('.jpeg', '.jpg')) else 'png'
                    with Image.open(io.BytesIO(data)) as im:
                        data = _image_neutre(im.width, im.height, fmt)
                elif item.filename == 'docProps/core.xml':
                    txt = data.decode('utf-8')
                    txt = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator></dc:creator>', txt)
                    data = txt.encode('utf-8')
                elif item.filename == 'word/_rels/settings.xml.rels':
                    # Lien vers le modèle .dotx d'origine (chemin d'un ancien poste + ancienne marée)
                    txt = data.decode('utf-8')
                    txt = re.sub(r'(Type="[^"]*attachedTemplate"\s+Target=")[^"]*(")',
                                 r'\1Modele%20Rapport%20SEALOR.dotx\2', txt)
                    data = txt.encode('utf-8')
                zout.writestr(item, data)
    return len(a_remplacer)


def anonymiser(chemin_in, chemin_out=None):
    chemin_out = chemin_out or chemin_in
    tmp = chemin_out + '.tmp.docx'
    document = docx.Document(chemin_in)
    journal = anonymiser_textes(document)
    document.save(tmp)
    nb_images = remplacer_images_et_metadonnees(tmp, chemin_out)
    os.remove(tmp)
    return journal, nb_images


if __name__ == '__main__':
    entree = sys.argv[1]
    sortie = sys.argv[2] if len(sys.argv) > 2 else entree
    journal, nb = anonymiser(entree, sortie)
    for ligne in journal:
        print('  -', ligne)
    print(f'  - {nb} images du corps remplacées par une image neutre')
    print('Template anonymisé :', sortie)
