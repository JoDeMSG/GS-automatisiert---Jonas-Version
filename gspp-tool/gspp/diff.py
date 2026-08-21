"""
Revisionsmanagement: Vergleich zweier Katalog-Snapshots.

Das ist der eigentliche Mehrwert gegenueber einem reinen Konverter. Das BSI
liefert Aenderungen als Commits, nicht als Releases - ohne Diffing entsteht
unbemerkte Abweichung der eigenen Dokumentation vom Stand der Technik.

Feldgewichtung:
  KRITISCH  -> Umsetzung muss neu bewertet werden (Text, Modalverb, Stufe)
  RELEVANT  -> Kontext veraendert, Pruefung empfohlen
  REDAKTION -> rein informativ

Umnummerierungs-Erkennung:
  Wenn das BSI eine Praktik intern umnummeriert (z.B. neue Anforderung
  eingefuegt, alles danach rutscht eine Nummer weiter), erscheint jede
  verschobene Anforderung als "entfallen" unter der alten ID UND als "neu"
  unter der neuen ID - obwohl der Inhalt identisch oder fast identisch ist.
  Ohne Gegenmassnahme ertrinkt eine echte inhaltliche Aenderung in Dutzenden
  falschen neu/entfallen-Meldungen (beobachtet am 2026-07-29-Katalogstand:
  BSI hat die Praktik BER umnummeriert und dabei 47 Anforderungen verschoben).

  _erkenne_umzuege() prueft daher vor der eigentlichen Diff-Berechnung, ob
  sich ein "entfallener" Text nahezu unveraendert unter neuer ID im selben
  Themenbereich wiederfindet (SequenceMatcher-Aehnlichkeit >= SCHWELLE) und
  fuehrt das Paar als eigene Kategorie "umnummeriert" statt als neu+entfallen.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from difflib import SequenceMatcher
from enum import Enum

from .models import CatalogSnapshot, Requirement


class Gewicht(str, Enum):
    KRITISCH = "kritisch"
    RELEVANT = "relevant"
    REDAKTION = "redaktionell"


FELD_GEWICHT: dict[str, Gewicht] = {
    "anforderungstext": Gewicht.KRITISCH,
    "modalverb": Gewicht.KRITISCH,
    "schutzbedarfsstufe": Gewicht.KRITISCH,
    "titel": Gewicht.KRITISCH,
    "verweise_required": Gewicht.KRITISCH,
    "ergebnis": Gewicht.RELEVANT,
    "ergebnis_spezifikation": Gewicht.RELEVANT,
    "zielobjekt_kategorien": Gewicht.RELEVANT,
    "dokumentation": Gewicht.RELEVANT,
    "aufwand": Gewicht.RELEVANT,
    "vertraulichkeit": Gewicht.RELEVANT,
    "integritaet": Gewicht.RELEVANT,
    "verfuegbarkeit": Gewicht.RELEVANT,
    "authentizitaet": Gewicht.RELEVANT,
    "gefaehrdungen": Gewicht.RELEVANT,
    "praktik_id": Gewicht.RELEVANT,
    "thema_id": Gewicht.RELEVANT,
    "erlaeuterung": Gewicht.REDAKTION,
    "aktionswort": Gewicht.REDAKTION,
    "tags": Gewicht.REDAKTION,
    "verweise_related": Gewicht.REDAKTION,
}

IGNORIERT = {"anforderungstext_roh", "alt_identifier", "ebene", "parent_id",
             "praktik_titel", "thema_titel"}


@dataclass
class Feldaenderung:
    feld: str
    alt: str
    neu: str
    gewicht: Gewicht


@dataclass
class Aenderung:
    anforderung_id: str
    art: str  # neu | entfallen | geaendert | umnummeriert
    titel: str
    praktik_id: str
    felder: list[Feldaenderung] = field(default_factory=list)
    alte_id: str | None = None  # nur bei art == "umnummeriert"

    @property
    def max_gewicht(self) -> Gewicht:
        if self.art == "umnummeriert":
            return Gewicht.KRITISCH if self.felder else Gewicht.REDAKTION
        if self.art in ("neu", "entfallen"):
            return Gewicht.KRITISCH
        if any(f.gewicht is Gewicht.KRITISCH for f in self.felder):
            return Gewicht.KRITISCH
        if any(f.gewicht is Gewicht.RELEVANT for f in self.felder):
            return Gewicht.RELEVANT
        return Gewicht.REDAKTION


SCHWELLE_UMZUG = 0.85  # Textaehnlichkeit, ab der zwei IDs als "dieselbe Anforderung" gelten


def _erkenne_umzuege(
    entfallen_map: dict[str, Requirement],
    neu_map: dict[str, Requirement],
) -> list[Aenderung]:
    """
    Paart entfallene und neue IDs, deren Anforderungstext (fast) identisch ist.

    Suche ist auf dieselbe Praktik eingeschraenkt (Umnummerierungen passieren
    praktikweise, nicht quer durch den Katalog) - haelt die Suche schnell und
    vermeidet zufaellige Treffer zwischen unverwandten kurzen Texten.
    Jede ID wird hoechstens einmal verbraucht (bestes Paar zuerst).
    """
    kandidaten: list[tuple[float, str, str]] = []
    for eid, ereq in entfallen_map.items():
        for nid, nreq in neu_map.items():
            if ereq.praktik_id != nreq.praktik_id:
                continue
            ratio = SequenceMatcher(None, ereq.anforderungstext, nreq.anforderungstext).ratio()
            if ratio >= SCHWELLE_UMZUG:
                kandidaten.append((ratio, eid, nid))
    kandidaten.sort(reverse=True)

    verbraucht_alt: set[str] = set()
    verbraucht_neu: set[str] = set()
    umzuege: list[Aenderung] = []
    for ratio, eid, nid in kandidaten:
        if eid in verbraucht_alt or nid in verbraucht_neu:
            continue
        verbraucht_alt.add(eid)
        verbraucht_neu.add(nid)
        alt_req, neu_req = entfallen_map[eid], neu_map[nid]
        felder = _vergleiche(alt_req, neu_req)
        # ID-Wechsel selbst ist keine inhaltliche Aenderung, ebenso wenig die
        # daraus zwangslaeufig folgende neue Position im Themenbaum.
        felder = [f for f in felder if f.feld not in ("anforderung_id", "thema_id", "pfad")]
        umzuege.append(
            Aenderung(
                anforderung_id=nid,
                art="umnummeriert",
                titel=neu_req.titel,
                praktik_id=neu_req.praktik_id,
                felder=felder,
                alte_id=eid,
            )
        )
    return umzuege, verbraucht_alt, verbraucht_neu


@dataclass
class Diffbericht:
    von_version: str
    nach_version: str
    von_commit: str | None
    nach_commit: str | None
    aenderungen: list[Aenderung]

    def zaehle(self, art: str) -> int:
        return sum(1 for a in self.aenderungen if a.art == art)

    @property
    def kritisch(self) -> list[Aenderung]:
        return [a for a in self.aenderungen if a.max_gewicht is Gewicht.KRITISCH]

    def zusammenfassung(self) -> str:
        umgezogen = self.zaehle("umnummeriert")
        zusatz = f", {umgezogen} umnummeriert" if umgezogen else ""
        return (
            f"{self.zaehle('neu')} neu, {self.zaehle('entfallen')} entfallen, "
            f"{self.zaehle('geaendert')} geaendert{zusatz} "
            f"({len(self.kritisch)} davon kritisch)"
        )


def _vergleiche(alt: Requirement, neu: Requirement) -> list[Feldaenderung]:
    out: list[Feldaenderung] = []
    a, n = alt.model_dump(), neu.model_dump()
    for feld in a:
        if feld in IGNORIERT:
            continue
        if a[feld] == n[feld]:
            continue
        out.append(
            Feldaenderung(
                feld=feld,
                alt="" if a[feld] is None else str(a[feld]),
                neu="" if n[feld] is None else str(n[feld]),
                gewicht=FELD_GEWICHT.get(feld, Gewicht.REDAKTION),
            )
        )
    return out


def diff(alt: CatalogSnapshot, neu: CatalogSnapshot) -> Diffbericht:
    """Vergleicht zwei Snapshots ueber die Anforderungs-ID als Schluessel."""
    alt_map = {r.anforderung_id: r for r in alt.requirements}
    neu_map = {r.anforderung_id: r for r in neu.requirements}

    nur_alt_ids = alt_map.keys() - neu_map.keys()
    nur_neu_ids = neu_map.keys() - alt_map.keys()
    nur_alt = {i: alt_map[i] for i in nur_alt_ids}
    nur_neu = {i: neu_map[i] for i in nur_neu_ids}

    umzuege, verbraucht_alt, verbraucht_neu = _erkenne_umzuege(nur_alt, nur_neu)

    aenderungen: list[Aenderung] = list(umzuege)

    for rid in sorted(nur_neu_ids - verbraucht_neu):
        r = neu_map[rid]
        aenderungen.append(Aenderung(rid, "neu", r.titel, r.praktik_id))

    for rid in sorted(nur_alt_ids - verbraucht_alt):
        r = alt_map[rid]
        aenderungen.append(Aenderung(rid, "entfallen", r.titel, r.praktik_id))

    for rid in sorted(alt_map.keys() & neu_map.keys()):
        felder = _vergleiche(alt_map[rid], neu_map[rid])
        if felder:
            r = neu_map[rid]
            aenderungen.append(Aenderung(rid, "geaendert", r.titel, r.praktik_id, felder))

    return Diffbericht(
        von_version=alt.katalog_version,
        nach_version=neu.katalog_version,
        von_commit=alt.commit_sha,
        nach_commit=neu.commit_sha,
        aenderungen=aenderungen,
    )
