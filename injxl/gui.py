# -*- coding: utf-8 -*-
"""Interface graphique d'INJXL (Tkinter / ttk)."""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import traceback
import webbrowser

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from . import APP_NAME, __version__, tables
from .engine import Moteur, Analyse, Annulation, afficher
from .model import (
    Profile, FileSpec, ColumnRef, KeyPair, Injection,
    MODES_LIBELLES, MODE_VIDES, DUP_LIBELLES, DUP_IGNORER,
    ST_A_ECRIRE, ST_IDENTIQUE, ST_CONFLIT, ST_ORPHELIN, ST_AMBIGU, ST_SANS_CLE, ST_RIEN,
    STATUTS,
)

# --- apparence ---------------------------------------------------------------
FOND = "#f4f6f9"
ENCRE = "#1f2933"
ACCENT = "#1f6feb"
DOUX = "#6b7684"
VERT = "#e6f4ea"
ROUGE = "#fdecea"
JAUNE = "#fff8e1"
GRIS = "#f0f1f3"

COULEURS_STATUT = {
    ST_A_ECRIRE: VERT,
    ST_CONFLIT: ROUGE,
    ST_ORPHELIN: JAUNE,
    ST_AMBIGU: JAUNE,
    ST_SANS_CLE: GRIS,
    ST_IDENTIQUE: "#ffffff",
    ST_RIEN: "#ffffff",
}

COCHE, VIDE = "☑", "☐"
MAX_AFFICHAGE = 3000          # lignes affichees dans la table de resultats
FILTRES = ["Tout"] + STATUTS

DOSSIER_CONFIG = os.path.join(os.path.expanduser("~"), ".injxl")
FICHIER_RECENTS = os.path.join(DOSSIER_CONFIG, "recents.json")

TYPES_TABLEAU = [
    ("Tableaux pris en charge", "*.xlsx *.xlsm *.xls *.csv *.txt *.tsv"),
    ("Classeurs Excel", "*.xlsx *.xlsm"),
    ("Fichiers texte", "*.csv *.txt *.tsv"),
    ("Tous les fichiers", "*.*"),
]
TYPES_PROFIL = [("Profil INJXL", "*.json"), ("Tous les fichiers", "*.*")]


def charger_recents():
    try:
        with open(FICHIER_RECENTS, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def enregistrer_recents(d):
    try:
        os.makedirs(DOSSIER_CONFIG, exist_ok=True)
        with open(FICHIER_RECENTS, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def ouvrir_dans_explorateur(chemin):
    try:
        if sys.platform.startswith("win"):
            os.startfile(os.path.normpath(chemin))  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", chemin])
        else:
            subprocess.Popen(["xdg-open", chemin])
    except Exception:
        webbrowser.open("file:///" + chemin.replace("\\", "/"))


# =============================================================================
#  Panneau de description d'un tableau (fichier + feuille + en-tete + apercu)
# =============================================================================

class PanneauFichier(ttk.Frame):

    def __init__(self, parent, titre, aide, app, role):
        ttk.Frame.__init__(self, parent, padding=(12, 8))
        self.app = app
        self.role = role
        self.table = None
        self.entetes = []

        cadre = ttk.LabelFrame(self, text="  %s  " % titre, padding=10)
        cadre.pack(fill="both", expand=True)
        cadre.columnconfigure(1, weight=1)

        ttk.Label(cadre, text=aide, style="Aide.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        ttk.Label(cadre, text="Fichier").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.var_path = tk.StringVar()
        e = ttk.Entry(cadre, textvariable=self.var_path)
        e.grid(row=1, column=1, sticky="ew")
        ttk.Button(cadre, text="Parcourir...", command=self.parcourir, width=14).grid(
            row=1, column=2, padx=6)
        ttk.Button(cadre, text="Recharger", command=self.recharger, width=12).grid(row=1, column=3)

        barre = ttk.Frame(cadre)
        barre.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 6))
        ttk.Label(barre, text="Feuille").pack(side="left", padx=(0, 6))
        self.var_sheet = tk.StringVar()
        self.cb_sheet = ttk.Combobox(barre, textvariable=self.var_sheet, state="readonly", width=32)
        self.cb_sheet.pack(side="left")
        self.cb_sheet.bind("<<ComboboxSelected>>", lambda e: self.recharger())
        ttk.Label(barre, text="Ligne d'en-tete").pack(side="left", padx=(18, 6))
        self.var_header = tk.IntVar(value=1)
        sp = ttk.Spinbox(barre, from_=1, to=200, width=6, textvariable=self.var_header,
                         command=self.recharger)
        sp.pack(side="left")
        sp.bind("<Return>", lambda e: self.recharger())
        self.lbl_info = ttk.Label(barre, text="", style="Aide.TLabel")
        self.lbl_info.pack(side="left", padx=18)

        conteneur = ttk.Frame(cadre)
        conteneur.grid(row=3, column=0, columnspan=4, sticky="nsew")
        cadre.rowconfigure(3, weight=1)
        self.tree = ttk.Treeview(conteneur, show="headings", height=8)
        vs = ttk.Scrollbar(conteneur, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(conteneur, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        conteneur.rowconfigure(0, weight=1)
        conteneur.columnconfigure(0, weight=1)

    # --- actions ---
    def parcourir(self):
        initial = self.app.recents.get("dernier_dossier", "")
        chemin = filedialog.askopenfilename(
            title="Choisir le tableau %s" % self.role, filetypes=TYPES_TABLEAU,
            initialdir=initial if os.path.isdir(initial) else None)
        if not chemin:
            return
        self.app.recents["dernier_dossier"] = os.path.dirname(chemin)
        self.definir(chemin)

    def definir(self, chemin, sheet="", header_row=None):
        self.var_path.set(chemin)
        try:
            feuilles = tables.lister_feuilles(chemin)
        except ImportError:
            messagebox.showerror(APP_NAME, "Le module 'xlrd' est necessaire pour lire les .xls.\n"
                                           "Installez-le (pip install xlrd) ou convertissez en .xlsx.")
            return
        except Exception as exc:
            messagebox.showerror(APP_NAME, "Lecture impossible :\n%s" % exc)
            return
        self.cb_sheet["values"] = feuilles
        self.var_sheet.set(sheet if sheet in feuilles else (feuilles[0] if feuilles else ""))
        if header_row:
            self.var_header.set(int(header_row))
        self.recharger()

    def recharger(self):
        chemin = self.var_path.get().strip().strip('"')
        if not chemin or not os.path.isfile(chemin):
            return
        try:
            ligne = max(1, int(self.var_header.get() or 1))
        except Exception:
            ligne = 1
        self.app.statut("Lecture de %s..." % os.path.basename(chemin))
        try:
            t = tables.ouvrir(chemin, self.var_sheet.get(), ligne, writable=False)
        except Exception as exc:
            self.app.statut("")
            messagebox.showerror(APP_NAME, "Lecture impossible :\n%s" % exc)
            return
        self.table = t
        self.entetes = t.headers()
        self.lbl_info.configure(
            text="%d ligne(s) de donnees, %d colonne(s)" % (t.nb_lignes_donnees(), t.max_column))
        self._remplir_apercu(t)
        self.app.statut("")
        self.app.colonnes_modifiees()

    def _remplir_apercu(self, t):
        self.tree.delete(*self.tree.get_children())
        colonnes = ["c%d" % i for i in range(1, t.max_column + 1)]
        self.tree["columns"] = colonnes
        for i, cid in enumerate(colonnes, start=1):
            entete = self.entetes[i - 1] if i - 1 < len(self.entetes) else ""
            self.tree.heading(cid, text="%s  %s" % (tables.lettre_colonne(i), entete))
            largeur = min(240, max(70, 9 * (len(entete) + 6)))
            self.tree.column(cid, width=largeur, minwidth=50, stretch=False, anchor="w")
        for ligne in t.apercu(40):
            self.tree.insert("", "end", values=[afficher(v) for v in ligne])

    # --- liaison au profil ---
    def vers_spec(self):
        return FileSpec(self.var_path.get().strip().strip('"'),
                        self.var_sheet.get(), int(self.var_header.get() or 1))

    def libelles(self):
        return ["%s - %s" % (tables.lettre_colonne(i + 1), h) if h
                else tables.lettre_colonne(i + 1)
                for i, h in enumerate(self.entetes)]

    def ref(self, index):
        """ColumnRef pour la colonne 1-base donnee."""
        entete = self.entetes[index - 1] if 1 <= index <= len(self.entetes) else ""
        return ColumnRef(index, entete)


# =============================================================================
#  Application
# =============================================================================

class Application(tk.Tk):

    def __init__(self, profil=None, chemin_profil=""):
        tk.Tk.__init__(self)
        self.title("%s %s" % (APP_NAME, __version__))
        self.geometry("1180x800")
        self.minsize(1000, 680)
        self.configure(bg=FOND)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self.profil = profil or Profile()
        self.chemin_profil = chemin_profil
        self.recents = charger_recents()
        self.moteur = None
        self.analyse = None
        self.file_msg = queue.Queue()
        self.annulation = False
        self.travail_en_cours = False

        self._styles()
        self._menu()
        self._entete()
        self._corps()
        self._pied()
        self.protocol("WM_DELETE_WINDOW", self.quitter)
        if profil:
            self.appliquer_profil(self.profil)
        self.after(120, self._pomper)

    # --- habillage ---------------------------------------------------------

    def _styles(self):
        st = ttk.Style(self)
        for theme in ("vista", "winnative", "clam", "default"):
            if theme in st.theme_names():
                st.theme_use(theme)
                break
        police = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"
        self.option_add("*Font", (police, 10))
        st.configure(".", background=FOND, foreground=ENCRE)
        st.configure("TFrame", background=FOND)
        st.configure("TLabel", background=FOND, foreground=ENCRE)
        st.configure("TLabelframe", background=FOND)
        st.configure("TLabelframe.Label", background=FOND, foreground=ACCENT,
                     font=(police, 10, "bold"))
        st.configure("Titre.TLabel", font=(police, 20, "bold"), foreground=ACCENT, background="white")
        st.configure("SousTitre.TLabel", font=(police, 10), foreground=DOUX, background="white")
        st.configure("Aide.TLabel", foreground=DOUX)
        st.configure("Chiffre.TLabel", font=(police, 16, "bold"))
        st.configure("Entete.TFrame", background="white")
        st.configure("Accent.TButton", font=(police, 10, "bold"))
        st.configure("Treeview", rowheight=22, fieldbackground="white", background="white")
        st.configure("Treeview.Heading", font=(police, 9, "bold"))
        st.configure("TNotebook.Tab", padding=(16, 8), font=(police, 10))

    def _menu(self):
        barre = tk.Menu(self)
        m = tk.Menu(barre, tearoff=0)
        m.add_command(label="Nouveau profil", accelerator="Ctrl+N", command=self.profil_nouveau)
        m.add_command(label="Ouvrir un profil...", accelerator="Ctrl+O", command=self.profil_ouvrir)
        m.add_separator()
        m.add_command(label="Enregistrer", accelerator="Ctrl+S", command=self.profil_enregistrer)
        m.add_command(label="Enregistrer sous...", command=self.profil_enregistrer_sous)
        self.menu_recents = tk.Menu(m, tearoff=0)
        m.add_cascade(label="Profils recents", menu=self.menu_recents)
        m.add_separator()
        m.add_command(label="Quitter", command=self.quitter)
        barre.add_cascade(label="Profil", menu=m)

        o = tk.Menu(barre, tearoff=0)
        o.add_command(label="Lancer l'analyse", accelerator="F5", command=self.lancer_analyse)
        o.add_command(label="Deviner les cles communes", command=self.deviner_cles)
        o.add_command(label="Ouvrir le dossier du fichier cible", command=self.ouvrir_dossier_cible)
        barre.add_cascade(label="Outils", menu=o)

        a = tk.Menu(barre, tearoff=0)
        a.add_command(label="Mode d'emploi", accelerator="F1", command=self.aide)
        a.add_command(label="A propos", command=self.a_propos)
        barre.add_cascade(label="Aide", menu=a)
        self.configure(menu=barre)
        self._rafraichir_recents()

        self.bind("<Control-n>", lambda e: self.profil_nouveau())
        self.bind("<Control-o>", lambda e: self.profil_ouvrir())
        self.bind("<Control-s>", lambda e: self.profil_enregistrer())
        self.bind("<F5>", lambda e: self.lancer_analyse())
        self.bind("<F1>", lambda e: self.aide())

    def _entete(self):
        bandeau = ttk.Frame(self, style="Entete.TFrame", padding=(18, 12))
        bandeau.pack(fill="x")
        gauche = ttk.Frame(bandeau, style="Entete.TFrame")
        gauche.pack(side="left")
        ttk.Label(gauche, text=APP_NAME, style="Titre.TLabel").pack(anchor="w")
        ttk.Label(gauche, text="Injection de colonnes d'un tableau vers un autre, par cle d'association",
                  style="SousTitre.TLabel").pack(anchor="w")
        self.var_profil = tk.StringVar(value="Profil : Sans titre")
        ttk.Label(bandeau, textvariable=self.var_profil, style="SousTitre.TLabel").pack(
            side="right", anchor="e")
        ttk.Separator(self, orient="horizontal").pack(fill="x")

    def _corps(self):
        self.onglets = ttk.Notebook(self)
        self.onglets.pack(fill="both", expand=True, padx=12, pady=(10, 6))
        self.onglet_fichiers()
        self.onglet_association()
        self.onglet_injection()
        self.onglet_analyse()

    def _pied(self):
        pied = ttk.Frame(self, padding=(12, 6))
        pied.pack(fill="x")
        self.var_statut = tk.StringVar(value="Pret.")
        ttk.Label(pied, textvariable=self.var_statut, style="Aide.TLabel").pack(side="left")
        self.progression = ttk.Progressbar(pied, mode="determinate", length=220)
        self.progression.pack(side="right", padx=(8, 0))
        self.btn_suivant = ttk.Button(pied, text="Suivant  >", width=12, command=self.suivant)
        self.btn_suivant.pack(side="right")
        ttk.Button(pied, text="<  Precedent", width=12, command=self.precedent).pack(side="right", padx=6)

    # =========================================================================
    #  Onglet 1 : fichiers
    # =========================================================================

    def onglet_fichiers(self):
        page = ttk.Frame(self.onglets)
        self.onglets.add(page, text="  1 · Tableaux  ")
        volet = ttk.PanedWindow(page, orient="vertical")
        volet.pack(fill="both", expand=True)
        self.pan_source = PanneauFichier(
            volet, "TABLEAU SOURCE  (les donnees a rapatrier)",
            "Contient la colonne a injecter. Ce fichier n'est jamais modifie.", self, "source")
        self.pan_cible = PanneauFichier(
            volet, "TABLEAU CIBLE  (le tableau a completer)",
            "Recevra les donnees. Par defaut, une copie est produite : l'original reste intact.",
            self, "cible")
        volet.add(self.pan_source, weight=1)
        volet.add(self.pan_cible, weight=1)

    # =========================================================================
    #  Onglet 2 : association
    # =========================================================================

    def onglet_association(self):
        page = ttk.Frame(self.onglets, padding=12)
        self.onglets.add(page, text="  2 · Cle d'association  ")

        cadre = ttk.LabelFrame(page, text="  COLONNES QUI IDENTIFIENT UNE MEME LIGNE  ", padding=10)
        cadre.pack(fill="both", expand=True)
        ttk.Label(cadre, style="Aide.TLabel",
                  text="Une ligne de la source et une ligne de la cible sont appariees quand toutes "
                       "les colonnes ci-dessous ont la meme valeur.\n"
                       "Ajoutez plusieurs couples pour une cle composite (ex. Code + Annee).").pack(
            anchor="w", pady=(0, 10))

        ligne = ttk.Frame(cadre)
        ligne.pack(fill="x", pady=(0, 8))
        ttk.Label(ligne, text="Source").grid(row=0, column=0, padx=(0, 6))
        self.cb_cle_src = ttk.Combobox(ligne, state="readonly", width=38)
        self.cb_cle_src.grid(row=0, column=1)
        ttk.Label(ligne, text="↔").grid(row=0, column=2, padx=10)
        ttk.Label(ligne, text="Cible").grid(row=0, column=3, padx=(0, 6))
        self.cb_cle_cib = ttk.Combobox(ligne, state="readonly", width=38)
        self.cb_cle_cib.grid(row=0, column=4)
        ttk.Button(ligne, text="+ Ajouter le couple", command=self.ajouter_cle).grid(
            row=0, column=5, padx=12)
        ttk.Button(ligne, text="Deviner", command=self.deviner_cles).grid(row=0, column=6)

        corps = ttk.Frame(cadre)
        corps.pack(fill="both", expand=True)
        self.tv_cles = ttk.Treeview(corps, columns=("n", "src", "cib"), show="headings", height=6)
        for cid, texte, largeur in (("n", "#", 40), ("src", "Colonne source", 380),
                                    ("cib", "Colonne cible", 380)):
            self.tv_cles.heading(cid, text=texte)
            self.tv_cles.column(cid, width=largeur, anchor="w")
        self.tv_cles.pack(side="left", fill="both", expand=True)
        cotes = ttk.Frame(corps, padding=(10, 0))
        cotes.pack(side="left", fill="y")
        ttk.Button(cotes, text="Monter", width=14, command=lambda: self.deplacer_cle(-1)).pack(pady=2)
        ttk.Button(cotes, text="Descendre", width=14, command=lambda: self.deplacer_cle(1)).pack(pady=2)
        ttk.Button(cotes, text="Supprimer", width=14, command=self.supprimer_cle).pack(pady=2)
        ttk.Button(cotes, text="Tout effacer", width=14,
                   command=lambda: (self.profil.keys.clear(), self.rafraichir_cles())).pack(pady=2)

        opts = ttk.LabelFrame(page, text="  COMPARAISON DES VALEURS DE CLE  ", padding=10)
        opts.pack(fill="x", pady=(12, 0))
        self.var_opts = {}
        definitions = [
            ("ignore_case", "Ignorer la casse  (AB12 = ab12)"),
            ("trim", "Ignorer les espaces de debut/fin"),
            ("collapse_spaces", "Reduire les espaces multiples"),
            ("numeric_keys", "Traiter 1234,0 comme 1234"),
            ("ignore_accents", "Ignorer les accents  (CRECHE = CRECHE)"),
            ("ignore_separators", "Ignorer - _ . / et espaces"),
            ("ignore_leading_zeros", "Ignorer les zeros de tete  (007 = 7)"),
        ]
        for i, (cle, libelle) in enumerate(definitions):
            v = tk.BooleanVar(value=getattr(self.profil.match, cle))
            self.var_opts[cle] = v
            ttk.Checkbutton(opts, text=libelle, variable=v).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 30), pady=2)
        bas = ttk.Frame(opts)
        bas.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(bas, text="Si une cle apparait plusieurs fois dans la source "
                            "avec des valeurs differentes :").pack(side="left", padx=(0, 8))
        self.var_dup = tk.StringVar(value=DUP_LIBELLES[self.profil.match.on_duplicate])
        ttk.Combobox(bas, textvariable=self.var_dup, state="readonly", width=38,
                     values=list(DUP_LIBELLES.values())).pack(side="left")

    def ajouter_cle(self):
        i_src = self.cb_cle_src.current()
        i_cib = self.cb_cle_cib.current()
        if i_src < 0 or i_cib < 0:
            messagebox.showinfo(APP_NAME, "Choisissez une colonne de chaque cote.")
            return
        self.profil.keys.append(KeyPair(self.pan_source.ref(i_src + 1), self.pan_cible.ref(i_cib + 1)))
        self.rafraichir_cles()

    def supprimer_cle(self):
        sel = self.tv_cles.selection()
        for iid in sorted((int(i) for i in sel), reverse=True):
            if 0 <= iid < len(self.profil.keys):
                self.profil.keys.pop(iid)
        self.rafraichir_cles()

    def deplacer_cle(self, delta):
        sel = self.tv_cles.selection()
        if not sel:
            return
        i = int(sel[0])
        j = i + delta
        if 0 <= j < len(self.profil.keys):
            self.profil.keys[i], self.profil.keys[j] = self.profil.keys[j], self.profil.keys[i]
            self.rafraichir_cles()
            self.tv_cles.selection_set(str(j))

    def deviner_cles(self):
        """Propose les couples de colonnes qui portent le meme en-tete des deux cotes."""
        src, cib = self.pan_source.entetes, self.pan_cible.entetes
        if not src or not cib:
            messagebox.showinfo(APP_NAME, "Chargez d'abord les deux tableaux (onglet 1).")
            return
        index_cib = {}
        for i, h in enumerate(cib):
            cle = h.strip().lower()
            if cle:
                index_cib.setdefault(cle, i + 1)
        trouves = []
        for i, h in enumerate(src):
            cle = h.strip().lower()
            if cle and cle in index_cib:
                trouves.append((i + 1, index_cib[cle], h))
        if not trouves:
            messagebox.showinfo(APP_NAME, "Aucun en-tete commun aux deux tableaux.")
            return
        libelles = "\n".join("  • %s" % h for _, _, h in trouves[:15])
        if not messagebox.askyesno(
                APP_NAME, "%d colonne(s) portent le meme en-tete des deux cotes :\n\n%s\n\n"
                          "Utiliser la premiere comme cle d'association ?\n"
                          "(Repondez Non pour les ajouter toutes en cle composite)"
                          % (len(trouves), libelles)):
            for i_src, i_cib, _ in trouves:
                self.profil.keys.append(KeyPair(self.pan_source.ref(i_src), self.pan_cible.ref(i_cib)))
        else:
            i_src, i_cib, _ = trouves[0]
            self.profil.keys.append(KeyPair(self.pan_source.ref(i_src), self.pan_cible.ref(i_cib)))
        self.rafraichir_cles()

    def rafraichir_cles(self):
        self.tv_cles.delete(*self.tv_cles.get_children())
        for i, k in enumerate(self.profil.keys):
            self.tv_cles.insert("", "end", iid=str(i),
                                values=(i + 1, k.source.label(), k.target.label()))

    # =========================================================================
    #  Onglet 3 : injection
    # =========================================================================

    def onglet_injection(self):
        page = ttk.Frame(self.onglets, padding=12)
        self.onglets.add(page, text="  3 · Colonnes a injecter  ")

        cadre = ttk.LabelFrame(page, text="  QUOI RAPATRIER, ET OU L'ECRIRE  ", padding=10)
        cadre.pack(fill="both", expand=True)
        ttk.Label(cadre, style="Aide.TLabel",
                  text="Chaque ligne ci-dessous decrit une colonne de la source recopiee dans une "
                       "colonne de la cible.\nVous pouvez en definir autant que necessaire.").pack(
            anchor="w", pady=(0, 10))

        form = ttk.Frame(cadre)
        form.pack(fill="x", pady=(0, 8))
        ttk.Label(form, text="Colonne source").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.cb_inj_src = ttk.Combobox(form, state="readonly", width=34)
        self.cb_inj_src.grid(row=1, column=0, padx=(0, 10))
        ttk.Label(form, text="Destination dans la cible").grid(row=0, column=1, sticky="w", padx=(0, 6))
        self.cb_inj_cib = ttk.Combobox(form, state="readonly", width=34)
        self.cb_inj_cib.grid(row=1, column=1, padx=(0, 10))
        self.cb_inj_cib.bind("<<ComboboxSelected>>", self._maj_nouvelle_colonne)
        ttk.Label(form, text="En-tete si nouvelle colonne").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.var_nouveau = tk.StringVar()
        self.ent_nouveau = ttk.Entry(form, textvariable=self.var_nouveau, width=24, state="disabled")
        self.ent_nouveau.grid(row=1, column=2, padx=(0, 10))
        ttk.Label(form, text="Si la cellule cible est deja remplie").grid(
            row=0, column=3, sticky="w", padx=(0, 6))
        self.var_mode = tk.StringVar(value=MODES_LIBELLES[MODE_VIDES])
        ttk.Combobox(form, textvariable=self.var_mode, state="readonly", width=34,
                     values=list(MODES_LIBELLES.values())).grid(row=1, column=3)

        form2 = ttk.Frame(cadre)
        form2.pack(fill="x", pady=(0, 10))
        self.var_defaut_actif = tk.BooleanVar(value=False)
        ttk.Checkbutton(form2, text="Si la cle est introuvable dans la source, ecrire :",
                        variable=self.var_defaut_actif).pack(side="left")
        self.var_defaut = tk.StringVar()
        ttk.Entry(form2, textvariable=self.var_defaut, width=22).pack(side="left", padx=8)
        ttk.Button(form2, text="+ Ajouter l'injection", style="Accent.TButton",
                   command=self.ajouter_injection).pack(side="left", padx=20)
        ttk.Button(form2, text="Remplacer la selection", command=self.remplacer_injection).pack(side="left")

        corps = ttk.Frame(cadre)
        corps.pack(fill="both", expand=True)
        cols = ("n", "src", "cib", "mode", "def")
        self.tv_inj = ttk.Treeview(corps, columns=cols, show="headings", height=8)
        for cid, texte, largeur in (("n", "#", 40), ("src", "Colonne source", 250),
                                    ("cib", "Destination", 250),
                                    ("mode", "Cellule cible deja remplie", 280),
                                    ("def", "Valeur par defaut", 130)):
            self.tv_inj.heading(cid, text=texte)
            self.tv_inj.column(cid, width=largeur, anchor="w")
        self.tv_inj.pack(side="left", fill="both", expand=True)
        cotes = ttk.Frame(corps, padding=(10, 0))
        cotes.pack(side="left", fill="y")
        ttk.Button(cotes, text="Monter", width=14, command=lambda: self.deplacer_inj(-1)).pack(pady=2)
        ttk.Button(cotes, text="Descendre", width=14, command=lambda: self.deplacer_inj(1)).pack(pady=2)
        ttk.Button(cotes, text="Supprimer", width=14, command=self.supprimer_injection).pack(pady=2)
        ttk.Button(cotes, text="Tout effacer", width=14,
                   command=lambda: (self.profil.injections.clear(), self.rafraichir_inj())).pack(pady=2)

    def _maj_nouvelle_colonne(self, _evt=None):
        nouvelle = self.cb_inj_cib.current() == 0
        self.ent_nouveau.configure(state="normal" if nouvelle else "disabled")
        if nouvelle and not self.var_nouveau.get():
            i = self.cb_inj_src.current()
            if i >= 0 and i < len(self.pan_source.entetes):
                self.var_nouveau.set(self.pan_source.entetes[i])

    def _injection_du_formulaire(self):
        i_src = self.cb_inj_src.current()
        i_cib = self.cb_inj_cib.current()
        if i_src < 0 or i_cib < 0:
            messagebox.showinfo(APP_NAME, "Choisissez la colonne source et la destination.")
            return None
        if i_cib == 0:   # premiere entree = nouvelle colonne
            if not self.var_nouveau.get().strip():
                messagebox.showinfo(APP_NAME, "Donnez un en-tete a la nouvelle colonne.")
                return None
            cible = None
        else:
            cible = self.pan_cible.ref(i_cib)   # decalage de 1 du au choix "nouvelle colonne"
        mode = next((k for k, v in MODES_LIBELLES.items() if v == self.var_mode.get()), MODE_VIDES)
        return Injection(self.pan_source.ref(i_src + 1), cible, self.var_nouveau.get().strip(),
                         mode, self.var_defaut.get(), bool(self.var_defaut_actif.get()))

    def ajouter_injection(self):
        inj = self._injection_du_formulaire()
        if inj:
            self.profil.injections.append(inj)
            self.rafraichir_inj()

    def remplacer_injection(self):
        sel = self.tv_inj.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Selectionnez d'abord une injection dans la liste.")
            return
        inj = self._injection_du_formulaire()
        if inj:
            self.profil.injections[int(sel[0])] = inj
            self.rafraichir_inj()

    def supprimer_injection(self):
        for iid in sorted((int(i) for i in self.tv_inj.selection()), reverse=True):
            if 0 <= iid < len(self.profil.injections):
                self.profil.injections.pop(iid)
        self.rafraichir_inj()

    def deplacer_inj(self, delta):
        sel = self.tv_inj.selection()
        if not sel:
            return
        i = int(sel[0])
        j = i + delta
        if 0 <= j < len(self.profil.injections):
            self.profil.injections[i], self.profil.injections[j] = \
                self.profil.injections[j], self.profil.injections[i]
            self.rafraichir_inj()
            self.tv_inj.selection_set(str(j))

    def rafraichir_inj(self):
        self.tv_inj.delete(*self.tv_inj.get_children())
        for i, inj in enumerate(self.profil.injections):
            cible = inj.target.label() if inj.target else "[nouvelle] %s" % inj.new_header
            self.tv_inj.insert("", "end", iid=str(i), values=(
                i + 1, inj.source.label(), cible,
                MODES_LIBELLES.get(inj.mode, inj.mode),
                inj.default if inj.apply_default else ""))

    # =========================================================================
    #  Onglet 4 : analyse et ecriture
    # =========================================================================

    def onglet_analyse(self):
        page = ttk.Frame(self.onglets, padding=12)
        self.onglets.add(page, text="  4 · Verification et ecriture  ")

        haut = ttk.Frame(page)
        haut.pack(fill="x")
        self.btn_analyse = ttk.Button(haut, text="Analyser (simulation, rien n'est ecrit)",
                                      style="Accent.TButton", command=self.lancer_analyse)
        self.btn_analyse.pack(side="left")
        self.btn_stop = ttk.Button(haut, text="Interrompre", command=self.interrompre, state="disabled")
        self.btn_stop.pack(side="left", padx=8)
        self.lbl_resume = ttk.Label(haut, text="Aucune analyse effectuee.", style="Aide.TLabel")
        self.lbl_resume.pack(side="left", padx=16)

        self.cadre_chiffres = ttk.Frame(page, padding=(0, 10))
        self.cadre_chiffres.pack(fill="x")
        self.chiffres = {}
        for cle, libelle, couleur in (
                (ST_A_ECRIRE, "A ecrire", "#137333"), (ST_IDENTIQUE, "Identiques", DOUX),
                (ST_CONFLIT, "Conflits", "#c5221f"), (ST_ORPHELIN, "Orphelins", "#b06000"),
                (ST_AMBIGU, "Ambigus", "#b06000"), (ST_SANS_CLE, "Sans cle", DOUX)):
            case = ttk.Frame(self.cadre_chiffres, relief="solid", borderwidth=1, padding=(14, 8))
            case.pack(side="left", padx=(0, 10))
            v = tk.StringVar(value="-")
            ttk.Label(case, textvariable=v, style="Chiffre.TLabel", foreground=couleur).pack()
            ttk.Label(case, text=libelle, style="Aide.TLabel").pack()
            self.chiffres[cle] = v

        filtres = ttk.Frame(page)
        filtres.pack(fill="x", pady=(0, 6))
        ttk.Label(filtres, text="Filtrer").pack(side="left", padx=(0, 6))
        self.var_filtre = tk.StringVar(value="Tout")
        cb = ttk.Combobox(filtres, textvariable=self.var_filtre, values=FILTRES,
                          state="readonly", width=16)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self.rafraichir_resultats())
        ttk.Label(filtres, text="Rechercher").pack(side="left", padx=(16, 6))
        self.var_recherche = tk.StringVar()
        e = ttk.Entry(filtres, textvariable=self.var_recherche, width=26)
        e.pack(side="left")
        e.bind("<Return>", lambda ev: self.rafraichir_resultats())
        ttk.Button(filtres, text="Appliquer", command=self.rafraichir_resultats).pack(side="left", padx=6)
        ttk.Button(filtres, text="Tout cocher",
                   command=lambda: self.cocher_tout(True)).pack(side="left", padx=(20, 4))
        ttk.Button(filtres, text="Tout decocher",
                   command=lambda: self.cocher_tout(False)).pack(side="left")
        self.lbl_selection = ttk.Label(filtres, text="", style="Aide.TLabel")
        self.lbl_selection.pack(side="right")

        conteneur = ttk.Frame(page)
        conteneur.pack(fill="both", expand=True)
        self.tv_res = ttk.Treeview(conteneur, show="headings", height=12, selectmode="extended")
        vs = ttk.Scrollbar(conteneur, orient="vertical", command=self.tv_res.yview)
        hs = ttk.Scrollbar(conteneur, orient="horizontal", command=self.tv_res.xview)
        self.tv_res.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tv_res.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        conteneur.rowconfigure(0, weight=1)
        conteneur.columnconfigure(0, weight=1)
        self.tv_res.bind("<Button-1>", self.clic_resultat)
        self.tv_res.bind("<space>", lambda e: self.basculer_selection())
        for statut, couleur in COULEURS_STATUT.items():
            self.tv_res.tag_configure(statut, background=couleur)

        sortie = ttk.LabelFrame(page, text="  PRODUCTION DU RESULTAT  ", padding=10)
        sortie.pack(fill="x", pady=(10, 0))
        sortie.columnconfigure(1, weight=1)
        self.var_sortie = tk.StringVar(value=self.profil.output.mode)
        ttk.Radiobutton(sortie, text="Creer une copie (l'original reste intact)",
                        variable=self.var_sortie, value="copie",
                        command=self._maj_sortie).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(sortie, text="Modifier le fichier cible lui-meme",
                        variable=self.var_sortie, value="surplace",
                        command=self._maj_sortie).grid(row=0, column=1, sticky="w", padx=20)
        self.var_backup = tk.BooleanVar(value=self.profil.output.backup)
        ttk.Checkbutton(sortie, text="avec sauvegarde prealable", variable=self.var_backup).grid(
            row=0, column=2, sticky="w")

        ttk.Label(sortie, text="Fichier produit").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.var_chemin_sortie = tk.StringVar()
        self.ent_sortie = ttk.Entry(sortie, textvariable=self.var_chemin_sortie)
        self.ent_sortie.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(8, 0), padx=(0, 6))
        self.btn_sortie = ttk.Button(sortie, text="Parcourir...", command=self.choisir_sortie, width=14)
        self.btn_sortie.grid(row=1, column=3, pady=(8, 0))

        options = ttk.Frame(sortie)
        options.grid(row=2, column=0, columnspan=4, sticky="w", pady=(10, 0))
        self.var_highlight = tk.BooleanVar(value=self.profil.output.highlight)
        self.var_rapport_txt = tk.BooleanVar(value=self.profil.output.report_txt)
        self.var_rapport_xlsx = tk.BooleanVar(value=self.profil.output.report_xlsx)
        self.var_ouvrir = tk.BooleanVar(value=self.profil.output.open_after)
        ttk.Checkbutton(options, text="Surligner les cellules modifiees",
                        variable=self.var_highlight).pack(side="left", padx=(0, 20))
        ttk.Checkbutton(options, text="Rapport .txt", variable=self.var_rapport_txt).pack(
            side="left", padx=(0, 20))
        ttk.Checkbutton(options, text="Rapport .xlsx", variable=self.var_rapport_xlsx).pack(
            side="left", padx=(0, 20))
        ttk.Checkbutton(options, text="Ouvrir le resultat a la fin", variable=self.var_ouvrir).pack(
            side="left")
        self.btn_ecrire = ttk.Button(sortie, text="ECRIRE LES DONNEES", style="Accent.TButton",
                                     command=self.ecrire, state="disabled")
        self.btn_ecrire.grid(row=2, column=3, sticky="e", pady=(10, 0))

    def _maj_sortie(self):
        surplace = self.var_sortie.get() == "surplace"
        etat = "disabled" if surplace else "normal"
        self.ent_sortie.configure(state=etat)
        self.btn_sortie.configure(state=etat)

    def choisir_sortie(self):
        cible = self.pan_cible.var_path.get()
        ext = os.path.splitext(cible)[1] or ".xlsx"
        chemin = filedialog.asksaveasfilename(
            title="Enregistrer le resultat sous", defaultextension=ext,
            initialfile=os.path.basename(tables.chemin_auto(cible)) if cible else "",
            filetypes=TYPES_TABLEAU)
        if chemin:
            self.var_chemin_sortie.set(chemin)

    # --- resultats ---------------------------------------------------------

    def rafraichir_resultats(self):
        self.tv_res.delete(*self.tv_res.get_children())
        if not self.analyse:
            return
        cols = ["chk", "ligne", "cle", "statut"]
        entetes = [("chk", "", 34), ("ligne", "Ligne", 60), ("cle", "Cle", 220), ("statut", "Statut", 110)]
        for n, inj in enumerate(self.profil.injections):
            cols += ["a%d" % n, "n%d" % n]
            nom = inj.source.header or inj.source.label()
            entetes += [("a%d" % n, "%s : actuel" % nom, 160), ("n%d" % n, "%s : propose" % nom, 160)]
        self.tv_res["columns"] = cols
        for cid, texte, largeur in entetes:
            self.tv_res.heading(cid, text=texte)
            self.tv_res.column(cid, width=largeur, anchor="center" if cid == "chk" else "w",
                               stretch=False)

        filtre = self.var_filtre.get()
        recherche = self.var_recherche.get().strip().lower()
        affiches = 0
        for i, res in enumerate(self.analyse.lignes):
            if filtre != "Tout" and res.statut != filtre:
                continue
            if recherche and recherche not in res.cle.lower():
                continue
            if affiches >= MAX_AFFICHAGE:
                break
            valeurs = [COCHE if res.inclure else VIDE, res.row, res.cle, res.statut]
            for n in range(len(self.profil.injections)):
                cell = next((c for c in res.cells if c.injection == n), None)
                valeurs += [afficher(cell.ancien) if cell else "",
                            afficher(cell.nouveau) if cell else ""]
            self.tv_res.insert("", "end", iid=str(i), values=valeurs, tags=(res.statut,))
            affiches += 1

        total = len(self.analyse.lignes)
        suffixe = "  (affichage limite aux %d premieres)" % MAX_AFFICHAGE if affiches >= MAX_AFFICHAGE else ""
        self.lbl_selection.configure(
            text="%d ligne(s) affichee(s) sur %d%s" % (affiches, total, suffixe))
        self._maj_compte_retenu()

    def clic_resultat(self, event):
        if self.tv_res.identify_region(event.x, event.y) != "cell":
            return
        if self.tv_res.identify_column(event.x) != "#1":
            return
        iid = self.tv_res.identify_row(event.y)
        if iid:
            self._basculer(iid)
            return "break"

    def basculer_selection(self):
        for iid in self.tv_res.selection():
            self._basculer(iid)

    def _basculer(self, iid):
        res = self.analyse.lignes[int(iid)]
        if res.statut != ST_A_ECRIRE:
            return
        res.inclure = not res.inclure
        self.tv_res.set(iid, "chk", COCHE if res.inclure else VIDE)
        self._maj_compte_retenu()

    def cocher_tout(self, etat):
        if not self.analyse:
            return
        for iid in self.tv_res.get_children():
            res = self.analyse.lignes[int(iid)]
            if res.statut != ST_A_ECRIRE:
                continue
            res.inclure = etat
            self.tv_res.set(iid, "chk", COCHE if etat else VIDE)
        self._maj_compte_retenu()

    def _maj_compte_retenu(self):
        if not self.analyse:
            return
        n = self.analyse.cellules_retenues()
        self.btn_ecrire.configure(state="normal" if n else "disabled",
                                  text="ECRIRE %d CELLULE(S)" % n if n else "ECRIRE LES DONNEES")

    # --- analyse en tache de fond -----------------------------------------

    def collecter(self):
        """Recopie l'etat des widgets dans le profil. Retourne un message d'erreur ou ''."""
        p = self.profil
        p.source = self.pan_source.vers_spec()
        p.target = self.pan_cible.vers_spec()
        for cle, var in self.var_opts.items():
            setattr(p.match, cle, bool(var.get()))
        p.match.on_duplicate = next((k for k, v in DUP_LIBELLES.items() if v == self.var_dup.get()),
                                    DUP_IGNORER)
        p.output.mode = self.var_sortie.get()
        p.output.path = self.var_chemin_sortie.get().strip()
        p.output.backup = bool(self.var_backup.get())
        p.output.highlight = bool(self.var_highlight.get())
        p.output.report_txt = bool(self.var_rapport_txt.get())
        p.output.report_xlsx = bool(self.var_rapport_xlsx.get())
        p.output.open_after = bool(self.var_ouvrir.get())

        if not p.source.path or not os.path.isfile(p.source.path):
            return "Choisissez un tableau source valide (onglet 1)."
        if not p.target.path or not os.path.isfile(p.target.path):
            return "Choisissez un tableau cible valide (onglet 1)."
        if os.path.abspath(p.source.path) == os.path.abspath(p.target.path):
            return "La source et la cible designent le meme fichier."
        if tables.type_fichier(p.target.path) == "xls":
            return "Le format .xls ne peut pas etre modifie. Enregistrez la cible en .xlsx."
        if not p.keys:
            return "Definissez au moins une cle d'association (onglet 2)."
        if not p.injections:
            return "Definissez au moins une colonne a injecter (onglet 3)."
        return ""

    def lancer_analyse(self):
        if self.travail_en_cours:
            return
        erreur = self.collecter()
        if erreur:
            messagebox.showwarning(APP_NAME, erreur)
            return
        self.onglets.select(3)
        self.analyse = None
        self.tv_res.delete(*self.tv_res.get_children())
        for v in self.chiffres.values():
            v.set("-")
        self.btn_ecrire.configure(state="disabled")
        self.travail_en_cours = True
        self.annulation = False
        self.btn_analyse.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.lbl_resume.configure(text="Analyse en cours...")
        self.moteur = Moteur(self.profil)
        threading.Thread(target=self._analyser_fond, daemon=True).start()

    def _analyser_fond(self):
        try:
            def progression(phase, fait, total):
                self.file_msg.put(("progression", (phase, fait, total)))
            self.moteur.ouvrir()
            analyse = self.moteur.analyser(progression, lambda: self.annulation)
            self.file_msg.put(("analyse_ok", analyse))
        except Annulation:
            self.file_msg.put(("annule", None))
        except Exception as exc:
            self.file_msg.put(("erreur", (str(exc), traceback.format_exc())))

    def interrompre(self):
        self.annulation = True
        self.statut("Interruption demandee...")

    def _pomper(self):
        """Boucle de reception des messages du thread de travail."""
        try:
            while True:
                genre, charge = self.file_msg.get_nowait()
                if genre == "progression":
                    phase, fait, total = charge
                    self.progression.configure(maximum=max(total, 1), value=fait)
                    self.statut("%s : %d / %d" % (phase, fait, total))
                elif genre == "analyse_ok":
                    self._analyse_terminee(charge)
                elif genre == "ecriture_ok":
                    self._ecriture_terminee(charge)
                elif genre == "annule":
                    self._fin_travail("Analyse interrompue. Aucune donnee ecrite.")
                elif genre == "erreur":
                    message, detail = charge
                    self._fin_travail("Echec.")
                    messagebox.showerror(APP_NAME, "Erreur :\n\n%s" % message)
                    sys.stderr.write(detail)
        except queue.Empty:
            pass
        self.after(120, self._pomper)

    def _fin_travail(self, message=""):
        self.travail_en_cours = False
        self.btn_analyse.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.progression.configure(value=0)
        self.statut("")
        if message:
            self.lbl_resume.configure(text=message)

    def _analyse_terminee(self, analyse: Analyse):
        self.analyse = analyse
        self._fin_travail()
        s = analyse.stats
        for cle, var in self.chiffres.items():
            var.set(str(s.get("statut_" + cle, 0)))
        self.lbl_resume.configure(
            text="%d cellule(s) a ecrire  |  source : %d cle(s) exploitable(s) sur %d"
                 % (s.get("cellules_a_ecrire", 0), s.get("source_exploitables", 0),
                    s.get("source_cles", 0)))
        if not self.var_chemin_sortie.get().strip():
            self.var_chemin_sortie.set(tables.chemin_auto(self.profil.target.path))
        self.rafraichir_resultats()
        messages = analyse.avertissements + [a for a in analyse.anomalies if a.startswith("[!!]")]
        if messages:
            messagebox.showwarning(APP_NAME, "Points de vigilance :\n\n" + "\n".join(messages[:12]))
        elif not s.get("cellules_a_ecrire", 0):
            messagebox.showinfo(APP_NAME, "Analyse terminee : aucune cellule a modifier.")

    # --- ecriture ----------------------------------------------------------

    def ecrire(self):
        if not self.analyse or self.travail_en_cours:
            return
        erreur = self.collecter()
        if erreur:
            messagebox.showwarning(APP_NAME, erreur)
            return
        n = self.analyse.cellules_retenues()
        if not n:
            return
        if self.profil.output.mode == "surplace":
            destination = os.path.abspath(self.profil.target.path)
            avertissement = ("\n\nLe fichier cible sera MODIFIE.%s"
                             % ("\nUne sauvegarde horodatee sera creee d'abord."
                                if self.profil.output.backup else
                                "\nAUCUNE sauvegarde ne sera creee."))
        else:
            destination = self.profil.output.path or tables.chemin_auto(self.profil.target.path)
            avertissement = "\n\nLe fichier cible d'origine ne sera pas touche."
        if not messagebox.askyesno(
                APP_NAME, "Ecrire %d cellule(s) ?\n\nDestination :\n%s%s"
                          % (n, destination, avertissement), icon="question"):
            return
        self.travail_en_cours = True
        self.btn_ecrire.configure(state="disabled")
        self.btn_analyse.configure(state="disabled")
        self.statut("Ecriture en cours...")
        threading.Thread(target=self._ecrire_fond, daemon=True).start()

    def _ecrire_fond(self):
        try:
            def progression(phase, fait, total):
                self.file_msg.put(("progression", (phase, fait, total)))
            resultat = self.moteur.appliquer(progression)
            base = resultat["destination"]
            resultat["rapport_txt"] = ""
            resultat["rapport_xlsx"] = ""
            if self.profil.output.report_txt:
                texte = self.moteur.rapport_texte(resultat)
                resultat["rapport_txt"] = self.moteur.ecrire_rapport_txt(texte, base)
            if self.profil.output.report_xlsx:
                resultat["rapport_xlsx"] = self.moteur.ecrire_rapport_xlsx(base)
            self.file_msg.put(("ecriture_ok", resultat))
        except PermissionError as exc:
            self.file_msg.put(("erreur", ("Fichier verrouille (ouvert dans Excel ?)\n%s" % exc,
                                          traceback.format_exc())))
        except Exception as exc:
            self.file_msg.put(("erreur", (str(exc), traceback.format_exc())))

    def _ecriture_terminee(self, resultat):
        self._fin_travail()
        self.btn_ecrire.configure(state="disabled")
        lignes = ["%d cellule(s) ecrite(s) sur %d ligne(s)." % (resultat["cellules"], resultat["lignes"]),
                  "", "Fichier : %s" % resultat["destination"]]
        if resultat.get("sauvegarde"):
            lignes.append("Sauvegarde : %s" % resultat["sauvegarde"])
        if resultat.get("rapport_txt"):
            lignes.append("Rapport : %s" % resultat["rapport_txt"])
        if resultat.get("rapport_xlsx"):
            lignes.append("Rapport Excel : %s" % resultat["rapport_xlsx"])
        self.lbl_resume.configure(text=lignes[0])
        messagebox.showinfo(APP_NAME, "\n".join(lignes))
        if self.profil.output.open_after:
            ouvrir_dans_explorateur(resultat["destination"])
        # le classeur cible a ete modifie en memoire : une nouvelle analyse s'impose
        try:
            self.moteur.fermer()
        except Exception:
            pass
        self.analyse = None
        self.tv_res.delete(*self.tv_res.get_children())

    # =========================================================================
    #  Profils
    # =========================================================================

    def colonnes_modifiees(self):
        """Met a jour toutes les listes de colonnes apres (re)chargement d'un tableau."""
        src = self.pan_source.libelles()
        cib = self.pan_cible.libelles()
        self.cb_cle_src["values"] = src
        self.cb_cle_cib["values"] = cib
        self.cb_inj_src["values"] = src
        self.cb_inj_cib["values"] = ["+  Nouvelle colonne (ajoutee a la fin)"] + cib

    def appliquer_profil(self, p: Profile):
        self.profil = p
        self.var_profil.set("Profil : %s" % p.name)
        if p.source.path and os.path.isfile(p.source.path):
            self.pan_source.definir(p.source.path, p.source.sheet, p.source.header_row)
        else:
            self.pan_source.var_header.set(p.source.header_row)
        if p.target.path and os.path.isfile(p.target.path):
            self.pan_cible.definir(p.target.path, p.target.sheet, p.target.header_row)
        else:
            self.pan_cible.var_header.set(p.target.header_row)
        for cle, var in self.var_opts.items():
            var.set(getattr(p.match, cle))
        self.var_dup.set(DUP_LIBELLES.get(p.match.on_duplicate, DUP_LIBELLES[DUP_IGNORER]))
        self.var_sortie.set(p.output.mode)
        self.var_chemin_sortie.set(p.output.path)
        self.var_backup.set(p.output.backup)
        self.var_highlight.set(p.output.highlight)
        self.var_rapport_txt.set(p.output.report_txt)
        self.var_rapport_xlsx.set(p.output.report_xlsx)
        self.var_ouvrir.set(p.output.open_after)
        self._maj_sortie()
        self.rafraichir_cles()
        self.rafraichir_inj()
        self.colonnes_modifiees()

    def profil_nouveau(self):
        if not messagebox.askyesno(APP_NAME, "Repartir d'une configuration vierge ?"):
            return
        self.chemin_profil = ""
        self.appliquer_profil(Profile())

    def profil_ouvrir(self, chemin=""):
        if not chemin:
            initial = self.recents.get("dossier_profils", os.path.join(os.getcwd(), "profils"))
            chemin = filedialog.askopenfilename(
                title="Ouvrir un profil INJXL", filetypes=TYPES_PROFIL,
                initialdir=initial if os.path.isdir(initial) else None)
        if not chemin:
            return
        try:
            p = Profile.load(chemin)
        except Exception as exc:
            messagebox.showerror(APP_NAME, "Profil illisible :\n%s" % exc)
            return
        self.chemin_profil = chemin
        self.recents["dossier_profils"] = os.path.dirname(chemin)
        self._ajouter_recent(chemin)
        self.appliquer_profil(p)
        self.statut("Profil charge : %s" % os.path.basename(chemin))

    def profil_enregistrer(self):
        if not self.chemin_profil:
            return self.profil_enregistrer_sous()
        self.collecter()
        try:
            self.profil.save(self.chemin_profil)
            self._ajouter_recent(self.chemin_profil)
            self.statut("Profil enregistre : %s" % os.path.basename(self.chemin_profil))
        except Exception as exc:
            messagebox.showerror(APP_NAME, "Enregistrement impossible :\n%s" % exc)

    def profil_enregistrer_sous(self):
        chemin = filedialog.asksaveasfilename(
            title="Enregistrer le profil", defaultextension=".json", filetypes=TYPES_PROFIL,
            initialfile="%s.json" % (self.profil.name or "profil"))
        if not chemin:
            return
        self.profil.name = os.path.splitext(os.path.basename(chemin))[0]
        self.var_profil.set("Profil : %s" % self.profil.name)
        self.chemin_profil = chemin
        self.profil_enregistrer()

    def _ajouter_recent(self, chemin):
        liste = [x for x in self.recents.get("profils", []) if x != chemin]
        liste.insert(0, chemin)
        self.recents["profils"] = liste[:10]
        enregistrer_recents(self.recents)
        self._rafraichir_recents()

    def _rafraichir_recents(self):
        self.menu_recents.delete(0, "end")
        liste = [x for x in self.recents.get("profils", []) if os.path.isfile(x)]
        if not liste:
            self.menu_recents.add_command(label="(aucun)", state="disabled")
            return
        for chemin in liste:
            self.menu_recents.add_command(
                label=os.path.basename(chemin),
                command=lambda c=chemin: self.profil_ouvrir(c))

    # =========================================================================
    #  Divers
    # =========================================================================

    def ouvrir_dossier_cible(self):
        chemin = self.pan_cible.var_path.get()
        if chemin and os.path.isfile(chemin):
            ouvrir_dans_explorateur(os.path.dirname(os.path.abspath(chemin)))

    def suivant(self):
        i = self.onglets.index(self.onglets.select())
        if i < 3:
            self.onglets.select(i + 1)

    def precedent(self):
        i = self.onglets.index(self.onglets.select())
        if i > 0:
            self.onglets.select(i - 1)

    def statut(self, texte):
        self.var_statut.set(texte or "Pret.")
        self.update_idletasks()

    def aide(self):
        texte = (
            "INJXL en quatre etapes\n"
            "\n"
            "1. Tableaux\n"
            "   Choisissez le tableau SOURCE (celui qui contient les donnees a rapatrier) et le\n"
            "   tableau CIBLE (celui a completer). Pour chacun : la feuille et la ligne d'en-tete.\n"
            "   Formats : .xlsx, .xlsm, .csv, .txt, .tsv, et .xls en lecture seule.\n"
            "\n"
            "2. Cle d'association\n"
            "   Indiquez quelles colonnes identifient une meme ligne des deux cotes. Plusieurs\n"
            "   couples = cle composite : toutes les parties doivent correspondre.\n"
            "   Les options de comparaison gerent casse, espaces, accents, zeros de tete...\n"
            "\n"
            "3. Colonnes a injecter\n"
            "   Pour chaque colonne source, choisissez la colonne de destination (existante ou\n"
            "   nouvelle) et ce qu'il faut faire si la cellule cible est deja remplie.\n"
            "\n"
            "4. Verification et ecriture\n"
            "   L'analyse est une simulation : elle n'ecrit rien. Elle liste ce qui serait ecrit,\n"
            "   les conflits, les orphelins et les doublons ambigus. Decochez les lignes a exclure,\n"
            "   puis lancez l'ecriture.\n"
            "\n"
            "Profils\n"
            "   Menu Profil > Enregistrer sous : toute la configuration part dans un .json\n"
            "   rejouable, y compris en ligne de commande :  INJXL.exe --profil mon_profil.json\n"
            "\n"
            "Statuts\n"
            "   A ecrire   : la cellule cible sera renseignee\n"
            "   Identique  : la cible contient deja la bonne valeur\n"
            "   Conflit    : la cible contient une AUTRE valeur, conservee (selon le mode choisi)\n"
            "   Orphelin   : la cle de la cible n'existe pas dans la source\n"
            "   Ambigu     : la cle apparait plusieurs fois dans la source avec des valeurs\n"
            "                differentes ; elle est ecartee par securite\n"
        )
        self._fenetre_texte("Mode d'emploi", texte)

    def a_propos(self):
        messagebox.showinfo(
            APP_NAME,
            "%s %s\n\nInjection de colonnes entre deux tableaux, par cle d'association.\n"
            "Analyse d'impact complete avant toute ecriture.\n\n"
            "Python %s - Tkinter" % (APP_NAME, __version__, sys.version.split()[0]))

    def _fenetre_texte(self, titre, texte):
        fen = tk.Toplevel(self)
        fen.title(titre)
        fen.geometry("820x620")
        fen.configure(bg=FOND)
        cadre = ttk.Frame(fen, padding=10)
        cadre.pack(fill="both", expand=True)
        zone = tk.Text(cadre, wrap="word", font=("Consolas", 10), relief="flat",
                       bg="white", fg=ENCRE, padx=12, pady=12)
        barre = ttk.Scrollbar(cadre, orient="vertical", command=zone.yview)
        zone.configure(yscrollcommand=barre.set)
        zone.pack(side="left", fill="both", expand=True)
        barre.pack(side="right", fill="y")
        zone.insert("1.0", texte)
        zone.configure(state="disabled")
        ttk.Button(fen, text="Fermer", command=fen.destroy).pack(pady=(0, 10))

    def quitter(self):
        if self.travail_en_cours:
            if not messagebox.askyesno(APP_NAME, "Un traitement est en cours. Quitter quand meme ?"):
                return
            self.annulation = True
        enregistrer_recents(self.recents)
        try:
            if self.moteur:
                self.moteur.fermer()
        except Exception:
            pass
        self.destroy()


def lancer(profil=None, chemin_profil=""):
    app = Application(profil, chemin_profil)
    app.mainloop()
