"""
Offizielle BSI-Zuordnungstabelle IT-Grundschutz 2023 -> Grundschutz++.

Quelle: control_layer/Mappings/IT-GS2023-zu-GSpp/ITGS-to-GS++-mapping_collection.json
Format: OSCAL Mapping Collection (NIST-Schema), Beziehungstypen laut OSCAL-Spezifikation:
    equal-to, equivalent-to, subset-of, superset-of, intersects-with

WICHTIG - Deckungsgrad ist unvollstaendig, kein Fehler:
    Von 1000 GS++-Anforderungen im Stand 2026-07-29 haben nur 306 (~31%) ueberhaupt
    eine Zuordnung zum alten IT-Grundschutz 2023. Das Mapping befindet sich laut
    BSI selbst noch im Aufbau (Pilotierungsphase). Anforderungen ohne Zuordnung
    bekommen in der Vorlage einen leeren Wert, keinen Fehler.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from pathlib import Path

import requests

from .fetch import TIMEOUT, sha256_bytes

log = logging.getLogger(__name__)

REPO = "BSI-Bund/Stand-der-Technik-Bibliothek"
BRANCH = "main"
MAPPING_PFAD = "control_layer/Mappings/IT-GS2023-zu-GSpp/ITGS-to-GS%2B%2B-mapping_collection.json"
RAW_URL = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/{MAPPING_PFAD}"


class MappingFetchError(RuntimeError):
    pass


def hole_mapping_roh(
    *, lokale_datei: Path | None = None, cache_dir: Path | None = None
) -> dict:
    """Holt die rohe Mapping-Collection - lokal oder von GitHub. Wirft bei Fehlschlag."""
    if lokale_datei is not None:
        import json

        return json.loads(Path(lokale_datei).read_bytes())

    r = requests.get(RAW_URL, timeout=TIMEOUT)
    if r.status_code != 200:
        raise MappingFetchError(f"Mapping-Download fehlgeschlagen: HTTP {r.status_code}")
    raw = r.content
    if cache_dir:
        cache_dir = Path(cache_dir)
        cache_dir.mkdir(parents=True, exist_ok=True)
        ziel = cache_dir / f"mapping_{sha256_bytes(raw)[:12]}.json"
        if not ziel.exists():
            ziel.write_bytes(raw)
    import json

    return json.loads(raw)


def parse_mapping(roh: dict) -> dict[str, list[tuple[str, str]]]:
    """
    Liefert: GS++-Anforderungs-ID -> Liste von (alte_2023_id, beziehungstyp).

    Mehrfachzuordnungen (eine neue Anforderung buendelt mehrere alte) sind haeufig -
    bis zu 23 alte IDs auf eine neue in der aktuellen Zuordnungstabelle.
    """
    je_neue_id: dict[str, list[tuple[str, str]]] = defaultdict(list)
    try:
        maps = roh["mapping-collection"]["mappings"][0]["maps"]
    except (KeyError, IndexError, TypeError) as exc:
        raise MappingFetchError(f"Unerwartete Struktur der Mapping-Datei: {exc}") from exc

    for eintrag in maps:
        beziehung = eintrag.get("relationship", "")
        for target in eintrag.get("targets", []):
            neu_id = target.get("id-ref")
            if not neu_id:
                continue
            for source in eintrag.get("sources", []):
                alt_id = source.get("id-ref")
                if alt_id:
                    je_neue_id[neu_id].append((alt_id, beziehung))

    log.info("Mapping geladen: %d GS++-Anforderungen mit Zuordnung zu IT-GS 2023",
             len(je_neue_id))
    return dict(je_neue_id)


def formatiere_spalten(paare: list[tuple[str, str]]) -> tuple[str, str]:
    """Zwei parallele, positionsgleiche Strings fuer die Excel-Spalten."""
    if not paare:
        return "", ""
    return (
        "; ".join(p[0] for p in paare),
        "; ".join(p[1] for p in paare),
    )


def reichere_an(reqs: list, mapping: dict[str, list[tuple[str, str]]]) -> list:
    """Ergaenzt alte_anforderungen/alte_beziehung je Requirement. Reine Kopie, keine Mutation."""
    angereichert = []
    treffer = 0
    for r in reqs:
        paare = mapping.get(r.anforderung_id)
        if paare:
            treffer += 1
            alte, bez = formatiere_spalten(paare)
            angereichert.append(r.model_copy(update={"alte_anforderungen": alte, "alte_beziehung": bez}))
        else:
            angereichert.append(r)
    log.info("Mapping angewendet: %d von %d Anforderungen mit Zuordnung zu IT-GS 2023 (%.0f%%)",
             treffer, len(reqs), 100 * treffer / len(reqs) if reqs else 0)
    return angereichert
