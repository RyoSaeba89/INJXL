#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""INJXL - point d'entree.

Sans argument : ouvre l'interface graphique.
Avec --profil  : rejoue une configuration enregistree, en ligne de commande.

Exemples
  INJXL.py
  INJXL.py profils\\exemple.json
  INJXL.py --profil profils\\exemple.json --simuler
  INJXL.py --profil profils\\exemple.json --source A.xlsx --cible B.xlsx --oui
"""

from __future__ import annotations

import argparse
import os
import sys
import traceback


def _sortie_sure(texte):
    """Ecrit sur la console meme si elle ne supporte pas l'UTF-8."""
    try:
        print(texte)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "ascii"
        print(texte.encode(enc, "replace").decode(enc, "replace"))


def _verifier_dependances():
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        _sortie_sure("ERREUR : le module openpyxl est introuvable.\n"
                     "Installation :  python -m pip install openpyxl")
        return False
    return True


def executer_cli(args):
    from injxl.model import Profile
    from injxl.engine import Moteur
    from injxl import tables

    profil = Profile.load(args.profil)
    if args.source:
        profil.source.path = args.source
    if args.cible:
        profil.target.path = args.cible
    if args.sortie:
        profil.output.path = args.sortie
        profil.output.mode = "copie"
    if args.surplace:
        profil.output.mode = "surplace"

    for chemin, role in ((profil.source.path, "source"), (profil.target.path, "cible")):
        if not chemin or not os.path.isfile(chemin):
            _sortie_sure("ERREUR : fichier %s introuvable : %r" % (role, chemin))
            return 2

    moteur = Moteur(profil)
    _sortie_sure("Profil    : %s" % profil.name)
    _sortie_sure("Source    : %s" % profil.source.path)
    _sortie_sure("Cible     : %s" % profil.target.path)
    _sortie_sure("Analyse en cours...")

    def progression(phase, fait, total):
        if total and fait and fait % 5000 == 0:
            _sortie_sure("  %s : %d / %d" % (phase, fait, total))

    moteur.ouvrir()
    analyse = moteur.analyser(progression)
    texte = moteur.rapport_texte(None)
    _sortie_sure(texte)

    n = analyse.cellules_retenues()
    if args.simuler:
        _sortie_sure("\nMode simulation : aucune donnee ecrite. (%d cellule(s) auraient ete ecrites)" % n)
        if profil.output.report_txt:
            chemin = moteur.ecrire_rapport_txt(texte, profil.target.path)
            _sortie_sure("Rapport : %s" % chemin)
        return 0
    if not n:
        _sortie_sure("\nAucune cellule a modifier.")
        return 0
    if not args.oui:
        try:
            reponse = input("\nEcrire %d cellule(s) ? (oui/non) : " % n).strip().lower()
        except EOFError:
            reponse = "non"
        if reponse not in ("oui", "o", "yes", "y"):
            _sortie_sure("Abandon. Aucun fichier modifie.")
            return 0

    resultat = moteur.appliquer(progression)
    _sortie_sure("\n%d cellule(s) ecrite(s)." % resultat["cellules"])
    _sortie_sure("Fichier produit : %s" % resultat["destination"])
    if resultat.get("sauvegarde"):
        _sortie_sure("Sauvegarde      : %s" % resultat["sauvegarde"])
    if profil.output.report_txt:
        chemin = moteur.ecrire_rapport_txt(moteur.rapport_texte(resultat), resultat["destination"])
        _sortie_sure("Rapport         : %s" % chemin)
    if profil.output.report_xlsx:
        chemin = moteur.ecrire_rapport_xlsx(resultat["destination"])
        _sortie_sure("Rapport Excel   : %s" % chemin)
    moteur.fermer()
    return 0


def main(argv=None):
    dossier = os.path.dirname(os.path.abspath(__file__))
    if dossier not in sys.path:
        sys.path.insert(0, dossier)

    analyseur = argparse.ArgumentParser(
        prog="INJXL", add_help=True,
        description="Injection de colonnes entre deux tableaux, par cle d'association.")
    analyseur.add_argument("profil_positionnel", nargs="?", default=None,
                           help="profil .json a charger dans l'interface")
    analyseur.add_argument("--profil", help="profil .json a rejouer sans interface")
    analyseur.add_argument("--source", help="remplace le fichier source du profil")
    analyseur.add_argument("--cible", help="remplace le fichier cible du profil")
    analyseur.add_argument("--sortie", help="chemin du fichier produit")
    analyseur.add_argument("--surplace", action="store_true",
                           help="modifie le fichier cible au lieu d'en creer une copie")
    analyseur.add_argument("--simuler", action="store_true",
                           help="analyse seulement, n'ecrit aucune donnee")
    analyseur.add_argument("--oui", action="store_true",
                           help="ne demande pas de confirmation avant d'ecrire")
    args = analyseur.parse_args(argv)

    if not _verifier_dependances():
        if not args.profil:
            input("\nAppuyez sur Entree pour quitter...")
        return 1

    if args.profil:
        return executer_cli(args)

    from injxl.model import Profile
    from injxl import gui
    profil, chemin = None, ""
    if args.profil_positionnel and os.path.isfile(args.profil_positionnel):
        try:
            profil = Profile.load(args.profil_positionnel)
            chemin = args.profil_positionnel
        except Exception as exc:
            _sortie_sure("Profil illisible : %s" % exc)
    gui.lancer(profil, chemin)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
    except Exception:
        traceback.print_exc()
        try:
            input("\nAppuyez sur Entree pour quitter...")
        except EOFError:
            pass
        sys.exit(1)
