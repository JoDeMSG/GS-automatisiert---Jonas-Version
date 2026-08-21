"""
Elementare Gefaehrdungen des BSI (G 0.x).

Quelle: documentation/namespaces/basethreats.csv im BSI-Repository.
Der Katalog selbst verweist in jeder threats-Eigenschaft per Namespace-URL
genau auf diese Datei - die Verknuepfung ist also vom BSI so vorgesehen und
nicht von uns erfunden.

Umfang (Stand 2026-08): 47 Gefaehrdungen, je mit Kurzbegriff und
ausfuehrlicher Definition (bis ueber 2.400 Zeichen Fliesstext).

Darstellungsentscheidung:
  * In der Anforderungszeile steht "G 0.18 – Fehlplanung oder fehlende
    Anpassung" - Kuerzel plus Begriff, lesbar ohne die Zeile zu sprengen.
  * Die vollstaendigen Definitionen stehen einmalig auf einem eigenen
    Nachschlageblatt. Sie in jede Zeile zu kopieren waere sinnlos: dieselbe
    Gefaehrdung taucht in dutzenden Anforderungen auf, das Blatt wuerde um
    ein Vielfaches wachsen ohne Informationsgewinn.
"""
from __future__ import annotations

import csv
import io
import logging
from pathlib import Path

import requests

from .fetch import TIMEOUT, sha256_bytes

log = logging.getLogger(__name__)

REPO = "BSI-Bund/Stand-der-Technik-Bibliothek"
BRANCH = "main"
THREATS_PFAD = "documentation/namespaces/basethreats.csv"
RAW_URL = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/{THREATS_PFAD}"


class ThreatsFetchError(RuntimeError):
    pass


def hole_gefaehrdungen(
    *, lokale_datei: Path | None = None, cache_dir: Path | None = None
) -> dict[str, tuple[str, str]]:
    """
    Liefert: Gefaehrdungs-ID -> (Begriff, Definition).

    Beispiel: "G 0.18" -> ("Fehlplanung oder fehlende Anpassung", "Wenn ...")
    """
    if lokale_datei is not None:
        roh = Path(lokale_datei).read_bytes()
    else:
        r = requests.get(RAW_URL, timeout=TIMEOUT)
        if r.status_code != 200:
            raise ThreatsFetchError(
                f"Download der Gefaehrdungsliste fehlgeschlagen: HTTP {r.status_code}")
        roh = r.content
        if cache_dir:
            cache_dir = Path(cache_dir)
            cache_dir.mkdir(parents=True, exist_ok=True)
            ziel = cache_dir / f"basethreats_{sha256_bytes(roh)[:12]}.csv"
            if not ziel.exists():
                ziel.write_bytes(roh)

    text = roh.decode("utf-8-sig")
    leser = csv.DictReader(io.StringIO(text))
    tabelle: dict[str, tuple[str, str]] = {}
    for zeile in leser:
        gid = (zeile.get("ID") or "").strip()
        if not gid:
            continue
        tabelle[gid] = (
            (zeile.get("Begriff") or "").strip(),
            (zeile.get("Definition") or "").strip(),
        )

    if not tabelle:
        raise ThreatsFetchError("Gefaehrdungsliste ist leer oder hat unerwartete Spalten.")
    log.info("Gefaehrdungen geladen: %d Eintraege", len(tabelle))
    return tabelle


def formatiere(ids_roh: str, tabelle: dict[str, tuple[str, str]]) -> str:
    """
    "G 0.18, G 0.19" -> "G 0.18 – Fehlplanung...; G 0.19 – Offenlegung..."

    Unbekannte IDs bleiben unveraendert stehen, statt still zu verschwinden -
    so faellt auf, wenn das BSI eine neue Gefaehrdung ergaenzt.
    """
    if not ids_roh:
        return ""
    teile = []
    for gid in [x.strip() for x in ids_roh.split(",") if x.strip()]:
        begriff = tabelle.get(gid, ("", ""))[0]
        teile.append(f"{gid} – {begriff}" if begriff else gid)
    return "; ".join(teile)


def reichere_an(reqs: list, tabelle: dict[str, tuple[str, str]]) -> list:
    """Ergaenzt gefaehrdungen_lang je Requirement. Reine Kopie, keine Mutation."""
    out = []
    treffer = 0
    for r in reqs:
        lang = formatiere(r.gefaehrdungen, tabelle)
        if lang:
            treffer += 1
        out.append(r.model_copy(update={"gefaehrdungen_lang": lang}))
    log.info("Gefaehrdungen zugeordnet: %d von %d Anforderungen", treffer, len(reqs))
    return out


def verwendete_ids(reqs: list) -> list[str]:
    """Alle im Katalog tatsaechlich referenzierten Gefaehrdungs-IDs, sortiert."""
    gesehen: set[str] = set()
    for r in reqs:
        for gid in [x.strip() for x in (r.gefaehrdungen or "").split(",") if x.strip()]:
            gesehen.add(gid)

    def sortier(g: str):
        """
        Sortiert nach Haupt- und Unternummer als GANZZAHLEN.

        Nicht als float: "G 0.1" und "G 0.10" ergaeben beide 0.1 und
        kollidierten - die Liste stand dann als G 0.1, G 0.10, G 0.11, G 0.2
        da statt in numerischer Reihenfolge.
        """
        rest = g.replace("G", "").strip()
        try:
            teile = [int(x) for x in rest.split(".")]
            return (0, teile)
        except ValueError:
            return (1, [0])

    return sorted(gesehen, key=sortier)
