# -*- coding: utf-8 -*-
"""Test de bout en bout du moteur INJXL, sans interface graphique.

    python tests/test_moteur.py

Cree ses propres classeurs dans un dossier temporaire : aucun fichier reel
n'est necessaire, et rien n'est ecrit dans le depot. Sort en code 1 au
premier ecart, pour servir de garde-fou en CI.
"""
import os
import sys
import shutil
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from injxl.model import (Profile, FileSpec, ColumnRef, KeyPair, Injection,
                         MODE_VIDES, MODE_TOUT, DUP_IGNORER)
from injxl.engine import Moteur
from injxl import tables
from injxl.model import ST_A_ECRIRE, ST_CONFLIT, ST_ORPHELIN, ST_IDENTIQUE, ST_AMBIGU

D = tempfile.mkdtemp(prefix="injxl_test_")
print("dossier:", D)

# --- fichier source ---------------------------------------------------------
src = os.path.join(D, "source.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Code Famille", "Annee", "Nom", "Quotient", "Tarif"])
lignes = [
    ("AB12", 2026, "Dupont", 850, "T1"),
    ("ab12 ", 2025, "Dupont", 700, "T2"),       # meme code, autre annee
    ("CD34", 2026, "Martin", 1200, "T3"),
    ("EF56", 2026, "Durand", 430, "T4"),
    ("GH78", 2026, "Petit", 999, "T5"),         # doublon en conflit
    ("GH78", 2026, "Petit", 111, "T5"),
    ("IJ90", 2026, "Leroy", None, "T6"),        # sans valeur
    (None, 2026, "Sans cle", 42, "T7"),
    ("KL11", 2026, "Absent-de-cible", 5, "T8"),
]
for l in lignes:
    ws.append(list(l))
wb.save(src)

# --- fichier cible ----------------------------------------------------------
cib = os.path.join(D, "cible.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Code Famille", "Annee", "Libelle", "Quotient", "Categorie"])
for l in [
    ("AB12", 2026, "Dupont", None, None),       # -> a ecrire 850
    ("AB12", 2025, "Dupont", None, None),       # -> a ecrire 700 (cle composite)
    ("CD34", 2026, "Martin", 1200, None),       # -> identique
    ("EF56", 2026, "Durand", 77, None),         # -> conflit (mode vides)
    ("GH78", 2026, "Petit", None, None),        # -> ambigu
    ("ZZ99", 2026, "Inconnu", None, None),      # -> orphelin
    ("IJ90", 2026, "Leroy", None, None),        # -> source vide => sans effet
]:
    ws.append(list(l))
wb.save(cib)

def profil_simple(cle_composite=False, mode=MODE_VIDES, nouvelle=False):
    p = Profile(name="test")
    p.source = FileSpec(src, "Sheet", 1)
    p.target = FileSpec(cib, "Sheet", 1)
    p.keys = [KeyPair(ColumnRef(1, "Code Famille"), ColumnRef(1, "Code Famille"))]
    if cle_composite:
        p.keys.append(KeyPair(ColumnRef(2, "Annee"), ColumnRef(2, "Annee")))
    cible = None if nouvelle else ColumnRef(4, "Quotient")
    p.injections = [Injection(ColumnRef(4, "Quotient"), cible, "Quotient importe", mode)]
    p.output.mode = "copie"
    p.output.path = os.path.join(D, "resultat.xlsx")
    return p

def compte(a):
    from collections import Counter
    return Counter(l.statut for l in a.lignes)

echecs = []
def verifier(nom, condition, detail=""):
    ligne = ("  OK   " if condition else "  ECHEC") + " | " + nom + ("  " + detail if detail else "")
    print(ligne.encode("ascii", "replace").decode("ascii"))
    if not condition:
        echecs.append(nom)

# --- test 1 : cle simple ----------------------------------------------------
print("\n[1] Cle simple, mode 'vides'")
m = Moteur(profil_simple())
a = m.analyser()
c = compte(a)
print("   ", dict(c))
verifier("AB12 ambigu en cle simple (2 quotients) -> rien a ecrire",
         c[ST_A_ECRIRE] == 0, str(c[ST_A_ECRIRE]))
verifier("CD34 identique", c[ST_IDENTIQUE] == 1)
verifier("EF56 conflit conserve", c[ST_CONFLIT] == 1)
verifier("3 lignes ambigues (2x AB12 + GH78)", c[ST_AMBIGU] == 3, str(c[ST_AMBIGU]))
verifier("ZZ99 orphelin", c[ST_ORPHELIN] == 1)

# --- test 2 : cle composite -------------------------------------------------
print("\n[2] Cle composite Code + Annee")
m2 = Moteur(profil_simple(cle_composite=True))
a2 = m2.analyser()
plans = {l.row: l for l in a2.lignes}
verifier("ligne 2 (AB12/2026) -> 850", plans[2].cells[0].nouveau == 850, str(plans[2].cells[0].nouveau))
verifier("ligne 3 (AB12/2025) -> 700", plans[3].cells[0].nouveau == 700, str(plans[3].cells[0].nouveau))

# --- test 3 : ecriture ------------------------------------------------------
print("\n[3] Ecriture reelle (copie)")
res = m2.appliquer()
verifier("fichier produit", os.path.isfile(res["destination"]), res["destination"])
wbv = openpyxl.load_workbook(res["destination"])
wsv = wbv.active
verifier("D2 == 850", wsv["D2"].value == 850, str(wsv["D2"].value))
verifier("D3 == 700", wsv["D3"].value == 700, str(wsv["D3"].value))
verifier("D5 inchange (conflit)", wsv["D5"].value == 77, str(wsv["D5"].value))
verifier("surlignage applique", wsv["D2"].fill.start_color.rgb in ("FFC6EFCE", "00C6EFCE"),
         str(wsv["D2"].fill.start_color.rgb))
txt = m2.rapport_texte(res)
verifier("rapport non vide", len(txt) > 500)
chemin_txt = m2.ecrire_rapport_txt(txt, res["destination"])
verifier("rapport txt ecrit", os.path.isfile(chemin_txt))
chemin_xlsx = m2.ecrire_rapport_xlsx(res["destination"])
verifier("rapport xlsx ecrit", os.path.isfile(chemin_xlsx))
m2.fermer()

# --- test 4 : nouvelle colonne ---------------------------------------------
print("\n[4] Injection dans une nouvelle colonne")
p4 = profil_simple(cle_composite=True, nouvelle=True)
p4.output.path = os.path.join(D, "resultat_nouvelle.xlsx")
m4 = Moteur(p4)
a4 = m4.analyser()
r4 = m4.appliquer()
wsv = openpyxl.load_workbook(r4["destination"]).active
verifier("nouvel en-tete en F1", wsv["F1"].value == "Quotient importe", str(wsv["F1"].value))
verifier("F2 == 850", wsv["F2"].value == 850, str(wsv["F2"].value))
m4.fermer()

# --- test 5 : mode ecrasement ----------------------------------------------
print("\n[5] Mode 'ecraser systematiquement'")
p5 = profil_simple(cle_composite=True, mode=MODE_TOUT)
p5.output.path = os.path.join(D, "resultat_ecrase.xlsx")
m5 = Moteur(p5)
a5 = m5.analyser()
r5 = m5.appliquer()
wsv = openpyxl.load_workbook(r5["destination"]).active
verifier("D5 ecrase par 430", wsv["D5"].value == 430, str(wsv["D5"].value))
m5.fermer()

# --- test 6 : exclusion de lignes -------------------------------------------
print("\n[6] Exclusion manuelle de lignes")
p6 = profil_simple(cle_composite=True)
p6.output.path = os.path.join(D, "resultat_exclu.xlsx")
m6 = Moteur(p6)
a6 = m6.analyser()
for l in a6.lignes:
    if l.row == 2:
        l.inclure = False
avant = a6.cellules_retenues()
r6 = m6.appliquer()
wsv = openpyxl.load_workbook(r6["destination"]).active
verifier("ligne exclue non ecrite", wsv["D2"].value is None, str(wsv["D2"].value))
verifier("ligne 3 ecrite", wsv["D3"].value == 700, str(wsv["D3"].value))
m6.fermer()

# --- test 7 : CSV en cible ---------------------------------------------------
print("\n[7] CSV en cible")
csv_cib = os.path.join(D, "cible.csv")
with open(csv_cib, "w", encoding="utf-8-sig", newline="") as f:
    f.write("Code Famille;Annee;Libelle;Quotient\n")
    f.write("AB12;2026;Dupont;\n")
    f.write("CD34;2026;Martin;\n")
p7 = Profile(name="csv")
p7.source = FileSpec(src, "Sheet", 1)
p7.target = FileSpec(csv_cib, "", 1)
p7.keys = [KeyPair(ColumnRef(1, "Code Famille"), ColumnRef(1, "Code Famille")),
           KeyPair(ColumnRef(2, "Annee"), ColumnRef(2, "Annee"))]
p7.injections = [Injection(ColumnRef(4, "Quotient"), ColumnRef(4, "Quotient"), "", MODE_VIDES)]
p7.output.path = os.path.join(D, "resultat.csv")
m7 = Moteur(p7)
a7 = m7.analyser()
r7 = m7.appliquer()
contenu = open(r7["destination"], encoding="utf-8-sig").read()
print("    ->", contenu.replace("\n", " / "))
verifier("CSV : AB12 renseigne", "AB12;2026;Dupont;850" in contenu)
verifier("CSV : CD34 renseigne", "CD34;2026;Martin;1200" in contenu)
m7.fermer()

# --- test 8 : ecriture sur place avec sauvegarde -----------------------------
print("\n[8] Ecriture sur place + sauvegarde")
cib8 = os.path.join(D, "surplace.xlsx")
shutil.copy2(cib, cib8)
p8 = profil_simple(cle_composite=True)
p8.target = FileSpec(cib8, "Sheet", 1)
p8.output.mode = "surplace"
p8.output.backup = True
m8 = Moteur(p8)
m8.analyser()
r8 = m8.appliquer()
verifier("sauvegarde creee", bool(r8["sauvegarde"]) and os.path.isfile(r8["sauvegarde"]))
verifier("original modifie", openpyxl.load_workbook(cib8).active["D2"].value == 850)
verifier("sauvegarde intacte", openpyxl.load_workbook(r8["sauvegarde"]).active["D2"].value is None)
m8.fermer()

# --- test 9 : profil json aller-retour ---------------------------------------
print("\n[9] Profil JSON aller-retour")
p9 = profil_simple(cle_composite=True)
chemin9 = os.path.join(D, "profil.json")
p9.save(chemin9)
p9b = Profile.load(chemin9)
verifier("2 cles rechargees", len(p9b.keys) == 2)
verifier("1 injection rechargee", len(p9b.injections) == 1)
verifier("chemin source conserve", p9b.source.path == src)

# --- test 10 : resolution par en-tete deplace --------------------------------
print("\n[10] Colonne deplacee : retrouvee par en-tete")
p10 = profil_simple(cle_composite=True)
p10.injections = [Injection(ColumnRef(9, "Quotient"), ColumnRef(9, "Quotient"), "", MODE_VIDES)]
p10.output.path = os.path.join(D, "resultat10.xlsx")
m10 = Moteur(p10)
a10 = m10.analyser()
verifier("colonne source retrouvee en 4", a10.cols_injection[0] == 4, str(a10.cols_injection))
verifier("avertissement emis", any("retrouv" in x for x in a10.avertissements),
         str(a10.avertissements[:2]))
m10.fermer()

print("\n" + "=" * 60)
if echecs:
    print("ECHECS (%d) : %s" % (len(echecs), echecs))
    sys.exit(1)
print("TOUS LES TESTS PASSENT")
