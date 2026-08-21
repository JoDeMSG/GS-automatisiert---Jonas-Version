"""
Datenmodell: OSCAL-Teilmenge (Eingang) + flaches Anforderungsmodell (intern).

Bewusst STRIKT validiert: wenn das BSI die Katalogstruktur aendert, bricht der
Lauf hier kontrolliert ab - statt still falsche Excel-Dateien zu erzeugen.
"""
from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

# ---------------------------------------------------------------- OSCAL-Eingang


class OscalBase(BaseModel):
    # extra="allow": unbekannte Felder werden toleriert (BSI ergaenzt sukzessive),
    # aber fehlende Pflichtfelder brechen den Lauf ab.
    model_config = ConfigDict(extra="allow", populate_by_name=True)


class Prop(OscalBase):
    name: str
    value: str
    ns: str | None = None
    clazz: str | None = Field(default=None, alias="class")


class Link(OscalBase):
    href: str
    rel: str | None = None
    text: str | None = None


class Param(OscalBase):
    id: str
    label: str | None = None
    values: list[str] = Field(default_factory=list)


class Part(OscalBase):
    id: str | None = None
    name: str
    prose: str | None = None
    props: list[Prop] = Field(default_factory=list)
    parts: list["Part"] = Field(default_factory=list)


class Control(OscalBase):
    id: str
    title: str
    clazz: str | None = Field(default=None, alias="class")
    params: list[Param] = Field(default_factory=list)
    props: list[Prop] = Field(default_factory=list)
    links: list[Link] = Field(default_factory=list)
    parts: list[Part] = Field(default_factory=list)
    controls: list["Control"] = Field(default_factory=list)


class Group(OscalBase):
    id: str | None = None
    title: str
    props: list[Prop] = Field(default_factory=list)
    groups: list["Group"] = Field(default_factory=list)
    controls: list[Control] = Field(default_factory=list)


class Metadata(OscalBase):
    title: str
    version: str
    oscal_version: str = Field(alias="oscal-version")
    last_modified: str = Field(alias="last-modified")
    props: list[Prop] = Field(default_factory=list)


class Catalog(OscalBase):
    uuid: str
    metadata: Metadata
    groups: list[Group] = Field(default_factory=list)
    back_matter: dict[str, Any] | None = Field(default=None, alias="back-matter")


class CatalogDocument(OscalBase):
    catalog: Catalog


# ------------------------------------------------------- Internes flaches Modell

SecLevel = Literal["normal-SdT", "erhoeht"]


class Requirement(BaseModel):
    """Eine Zeile in der Excel-Vorlage."""

    model_config = ConfigDict(extra="forbid")

    # Verortung
    praktik_id: str
    praktik_titel: str
    thema_id: str
    thema_titel: str
    anforderung_id: str
    parent_id: str | None
    ebene: int  # 0 = Top-Level-Anforderung, 1..n = untergeordnet

    # Baumsemantik (siehe parser.klassifiziere)
    knotentyp: str  # Sammelanforderung | Anforderung m. Teilanforderungen | Einzelanforderung
    pruefpflichtig: bool  # False = Status wird aus Teilanforderungen aggregiert
    anzahl_teilanforderungen: int  # nur direkte Kinder
    pfad: str  # "GC.5.1 > GC.5.1.1"

    # Kern
    titel: str
    anforderungstext: str  # Parameter aufgeloest
    anforderungstext_roh: str  # mit {{ insert: param, ... }}
    erlaeuterung: str

    # Steuerungsattribute
    schutzbedarfsstufe: str  # normal-SdT | erhoeht
    modalverb: str  # MUSS | SOLLTE | KANN
    aufwand: int | None  # effort_level 0..5

    # Semantik
    aktionswort: str
    ergebnis: str
    ergebnis_spezifikation: str
    dokumentation: str
    zielobjekt_kategorien: str

    # Schutzziele (0/1/2)
    vertraulichkeit: int | None
    integritaet: int | None
    verfuegbarkeit: int | None
    authentizitaet: int | None

    # Kontext
    gefaehrdungen: str  # "G 0.18, G 0.31"
    gefaehrdungen_lang: str = ""  # "G 0.18 – Fehlplanung...; G 0.19 – Offenlegung..."
    tags: str
    verweise_related: str
    verweise_required: str

    # Provenienz
    alt_identifier: str

    # Zuordnung zum alten IT-Grundschutz 2023 (offizielles BSI-Mapping, unvollstaendig)
    alte_anforderungen: str = ""  # "APP.3.6.A1-UA.2; NET.1.1.A16-UA.2"
    alte_beziehung: str = ""      # "subset-of; superset-of" - positionsgleich zu oben

    def key(self) -> str:
        return self.anforderung_id


class CatalogSnapshot(BaseModel):
    """Katalog + Herkunftsnachweis. Basis fuer Diffing und Revisionssicherheit."""

    model_config = ConfigDict(extra="forbid")

    quelle_url: str
    commit_sha: str | None
    commit_datum: str | None
    katalog_uuid: str
    katalog_version: str
    katalog_last_modified: str
    oscal_version: str
    sha256_quelle: str
    abgerufen_am: str
    tool_version: str
    requirements: list[Requirement]
