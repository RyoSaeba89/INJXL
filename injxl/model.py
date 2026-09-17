# -*- coding: utf-8 -*-
"""Modele de donnees d'INJXL : description d'un mappage et serialisation JSON."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional

SCHEMA = "injxl/profil/1"

# --- modes d'ecriture d'une injection -----------------------------------------
MODE_VIDES = "vides"              # ne remplit que les cellules cibles vides
MODE_TOUT = "tout"                # ecrase systematiquement
MODE_DIFFERENTES = "differentes"  # ecrase uniquement si la valeur differe

MODES_LIBELLES = {
    MODE_VIDES: "Remplir uniquement les cellules vides",
    MODE_TOUT: "Ecraser systematiquement",
    MODE_DIFFERENTES: "Ecraser seulement si la valeur differe",
}

# --- arbitrage des doublons dans la source ------------------------------------
DUP_IGNORER = "ignorer"   # cle en doublon avec valeurs differentes -> cle ecartee
DUP_PREMIER = "premier"   # on garde la premiere occurrence
DUP_DERNIER = "dernier"   # on garde la derniere occurrence

DUP_LIBELLES = {
    DUP_IGNORER: "Ecarter la cle (aucune injection)",
    DUP_PREMIER: "Garder la premiere occurrence",
    DUP_DERNIER: "Garder la derniere occurrence",
}

# --- statuts de ligne ----------------------------------------------------------
ST_A_ECRIRE = "A ecrire"
ST_IDENTIQUE = "Identique"
ST_CONFLIT = "Conflit"          # cible deja remplie differemment, conservee
ST_ORPHELIN = "Orphelin"        # cle de la cible absente de la source
ST_AMBIGU = "Source ambigue"    # cle en doublon non arbitre dans la source
ST_SANS_CLE = "Sans cle"
ST_RIEN = "Sans effet"

STATUTS = [ST_A_ECRIRE, ST_IDENTIQUE, ST_CONFLIT, ST_ORPHELIN,
           ST_AMBIGU, ST_SANS_CLE, ST_RIEN]


@dataclass
class ColumnRef:
    """Reference a une colonne : index 1-base + libelle d'en-tete.

    Le libelle sert a retrouver la colonne quand un profil est rejoue sur un
    fichier dont les colonnes ont bouge ; l'index sert de repli.
    """
    index: int
    header: str = ""

    def label(self) -> str:
        from .tables import lettre_colonne
        base = lettre_colonne(self.index)
        return "%s - %s" % (base, self.header) if self.header else base

    def to_dict(self) -> Dict[str, Any]:
        return {"index": self.index, "header": self.header}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "ColumnRef":
        return ColumnRef(int(d.get("index", 1)), str(d.get("header", "") or ""))


@dataclass
class FileSpec:
    """Un tableau : fichier + feuille + ligne d'en-tete."""
    path: str = ""
    sheet: str = ""
    header_row: int = 1

    def to_dict(self) -> Dict[str, Any]:
        return {"path": self.path, "sheet": self.sheet, "header_row": self.header_row}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "FileSpec":
        return FileSpec(str(d.get("path", "")), str(d.get("sheet", "") or ""),
                        int(d.get("header_row", 1) or 1))


@dataclass
class KeyPair:
    """Un couple de colonnes formant (une partie de) la cle d'association."""
    source: ColumnRef
    target: ColumnRef

    def to_dict(self):
        return {"source": self.source.to_dict(), "target": self.target.to_dict()}

    @staticmethod
    def from_dict(d):
        return KeyPair(ColumnRef.from_dict(d["source"]), ColumnRef.from_dict(d["target"]))


@dataclass
class Injection:
    """Une colonne source rapatriee vers une colonne cible."""
    source: ColumnRef
    target: Optional[ColumnRef] = None      # None => nouvelle colonne
    new_header: str = ""                    # en-tete si nouvelle colonne
    mode: str = MODE_VIDES
    default: str = ""                       # valeur si la cle est introuvable
    apply_default: bool = False

    def titre(self) -> str:
        cible = self.target.label() if self.target else "[nouvelle] " + self.new_header
        return "%s  ->  %s" % (self.source.label(), cible)

    def to_dict(self):
        return {
            "source": self.source.to_dict(),
            "target": self.target.to_dict() if self.target else None,
            "new_header": self.new_header,
            "mode": self.mode,
            "default": self.default,
            "apply_default": self.apply_default,
        }

    @staticmethod
    def from_dict(d):
        t = d.get("target")
        return Injection(
            ColumnRef.from_dict(d["source"]),
            ColumnRef.from_dict(t) if t else None,
            str(d.get("new_header", "") or ""),
            str(d.get("mode", MODE_VIDES)),
            str(d.get("default", "") or ""),
            bool(d.get("apply_default", False)),
        )


@dataclass
class MatchOptions:
    """Regles de normalisation des cles et d'arbitrage des doublons."""
    ignore_case: bool = True
    trim: bool = True
    collapse_spaces: bool = True
    ignore_accents: bool = False
    ignore_separators: bool = False     # retire - _ . / et espaces
    ignore_leading_zeros: bool = False
    numeric_keys: bool = True           # 1234.0 -> 1234
    on_duplicate: str = DUP_IGNORER

    def to_dict(self):
        return asdict(self)

    @staticmethod
    def from_dict(d):
        o = MatchOptions()
        for k, v in (d or {}).items():
            if hasattr(o, k):
                setattr(o, k, v)
        return o


@dataclass
class OutputOptions:
    """Ce qui est produit a la fin."""
    mode: str = "copie"              # "copie" (nouveau fichier) ou "surplace"
    path: str = ""                   # chemin impose ; vide => nom horodate auto
    backup: bool = True              # sauvegarde avant ecriture sur place
    highlight: bool = True
    highlight_color: str = "FFC6EFCE"
    report_txt: bool = True
    report_xlsx: bool = False
    open_after: bool = False

    def to_dict(self):
        return asdict(self)

    @staticmethod
    def from_dict(d):
        o = OutputOptions()
        for k, v in (d or {}).items():
            if hasattr(o, k):
                setattr(o, k, v)
        return o


@dataclass
class Profile:
    """Configuration complete, sauvegardable en .json."""
    name: str = "Sans titre"
    source: FileSpec = field(default_factory=FileSpec)
    target: FileSpec = field(default_factory=FileSpec)
    keys: List[KeyPair] = field(default_factory=list)
    injections: List[Injection] = field(default_factory=list)
    match: MatchOptions = field(default_factory=MatchOptions)
    output: OutputOptions = field(default_factory=OutputOptions)
    remember_paths: bool = True

    def to_dict(self) -> Dict[str, Any]:
        src = self.source.to_dict()
        tgt = self.target.to_dict()
        if not self.remember_paths:
            src = dict(src, path="")
            tgt = dict(tgt, path="")
        return {
            "schema": SCHEMA,
            "name": self.name,
            "source": src,
            "target": tgt,
            "keys": [k.to_dict() for k in self.keys],
            "injections": [i.to_dict() for i in self.injections],
            "match": self.match.to_dict(),
            "output": self.output.to_dict(),
            "remember_paths": self.remember_paths,
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Profile":
        p = Profile()
        p.name = str(d.get("name", "Sans titre"))
        p.source = FileSpec.from_dict(d.get("source", {}))
        p.target = FileSpec.from_dict(d.get("target", {}))
        p.keys = [KeyPair.from_dict(x) for x in d.get("keys", [])]
        p.injections = [Injection.from_dict(x) for x in d.get("injections", [])]
        p.match = MatchOptions.from_dict(d.get("match", {}))
        p.output = OutputOptions.from_dict(d.get("output", {}))
        p.remember_paths = bool(d.get("remember_paths", True))
        return p

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @staticmethod
    def load(path: str) -> "Profile":
        with open(path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        schema = data.get("schema")
        if schema and not str(schema).startswith("injxl/profil/"):
            raise ValueError("Ce fichier n'est pas un profil INJXL.")
        p = Profile.from_dict(data)
        if not p.name or p.name == "Sans titre":
            p.name = os.path.splitext(os.path.basename(path))[0]
        return p
