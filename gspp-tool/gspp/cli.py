"""
Kommandozeile.

  gspp build     Katalog holen -> Excel-Vorlage(n) erzeugen
  gspp snapshot  Katalog holen -> versionierten JSON-Snapshot ablegen
  gspp diff      Zwei Snapshots vergleichen -> Aenderungsbericht
  gspp watch     Snapshot + Diff gegen den letzten Stand in einem Lauf (fuer CI)
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

from . import __version__
from .diff import diff as diff_snapshots
from .excel import befuelle_template, erzeuge_mappe
from .parser import filter_mit_ahnen
from .schema2023 import erzeuge_mappe_2023
from .fetch import hole_katalog
from .models import CatalogSnapshot
from .parser import parse
from .report import schreibe_excel, schreibe_markdown

log = logging.getLogger("gspp")

# Gefaehrdungstabelle des aktuellen Laufs - wird vom Excel-Renderer fuer das
# Nachschlageblatt gebraucht. Bewusst hier und nicht im Snapshot: es sind
# Stammdaten des BSI, keine Eigenschaft eines einzelnen Katalogstands.
_GEFAEHRDUNGSTABELLE: dict[str, tuple[str, str]] = {}


def _snapshot(args) -> CatalogSnapshot:
    roh, herkunft = hole_katalog(
        lokale_datei=Path(args.katalog) if args.katalog else None,
        cache_dir=Path(args.cache) if args.cache else None,
        token=args.token,
    )
    doc, reqs = parse(roh)

    if not getattr(args, "ohne_mapping", False):
        try:
            from . import mapping as mapping_mod
            mroh = mapping_mod.hole_mapping_roh(
                lokale_datei=Path(args.mapping) if getattr(args, "mapping", None) else None,
                cache_dir=Path(args.cache) if args.cache else None,
            )
            zuordnung = mapping_mod.parse_mapping(mroh)
            reqs = mapping_mod.reichere_an(reqs, zuordnung)
        except Exception as exc:  # Mapping ist Zusatzinfo - darf den Lauf nie blockieren
            log.warning("Zuordnung zu IT-GS 2023 nicht verfuegbar (%s) - Vorlage wird ohne "
                        "diese Spalten erzeugt.", exc)

    if not getattr(args, "ohne_gefaehrdungen", False):
        try:
            from . import threats as threats_mod
            tabelle = threats_mod.hole_gefaehrdungen(
                lokale_datei=Path(args.gefaehrdungen) if getattr(args, "gefaehrdungen", None) else None,
                cache_dir=Path(args.cache) if args.cache else None,
            )
            reqs = threats_mod.reichere_an(reqs, tabelle)
            _GEFAEHRDUNGSTABELLE.clear()
            _GEFAEHRDUNGSTABELLE.update(tabelle)
        except Exception as exc:  # ebenfalls nur Zusatzinfo
            log.warning("Gefaehrdungsbeschreibungen nicht verfuegbar (%s) - es bleiben "
                        "die Kuerzel wie 'G 0.18'.", exc)

    m = doc.catalog.metadata
    return CatalogSnapshot(
        katalog_uuid=doc.catalog.uuid,
        katalog_version=m.version,
        katalog_last_modified=m.last_modified,
        oscal_version=m.oscal_version,
        tool_version=__version__,
        requirements=reqs,
        **herkunft,
    )


def _speichere_snapshot(snap: CatalogSnapshot, verzeichnis: Path) -> Path:
    verzeichnis.mkdir(parents=True, exist_ok=True)
    stempel = datetime.now(timezone.utc).strftime("%Y%m%d")
    kennung = (snap.commit_sha or snap.sha256_quelle)[:12]
    ziel = verzeichnis / f"snapshot_{stempel}_{kennung}.json"
    ziel.write_text(snap.model_dump_json(indent=1), encoding="utf-8")
    log.info("Snapshot: %s", ziel)
    return ziel


def _lade_snapshot(pfad: Path) -> CatalogSnapshot:
    return CatalogSnapshot.model_validate_json(Path(pfad).read_text(encoding="utf-8"))


def _neuester_snapshot(verzeichnis: Path) -> Path | None:
    """
    Juengster Snapshot nach Aenderungszeit der Datei.

    Nicht nach Dateiname sortieren: der traegt das Laufdatum, nicht die
    Katalogversion. Zwei Laeufe am selben Tag wuerden sonst ueber den
    Hash-Suffix alphabetisch entschieden - also zufaellig.
    """
    kandidaten = list(Path(verzeichnis).glob("snapshot_*.json"))
    if not kandidaten:
        return None
    return max(kandidaten, key=lambda f: f.stat().st_mtime)


# ------------------------------------------------------------------- Befehle


def cmd_build(args) -> int:
    if getattr(args, "design", None):
        from . import design as design_mod
        profil = design_mod.setze_profil(args.design)
        log.info("Designprofil: %s (Akzent #%s, Schrift %s)",
                 profil.name, profil.akzent, profil.schrift)
    snap = _snapshot(args)
    out = Path(args.out)

    if args.template:
        befuelle_template(snap, Path(args.template), out,
                          blattname=args.blatt, kopfzeile=args.kopfzeile,
                          filter_stufe=args.stufe)
    elif args.ziel_schema == "2023":
        reqs = snap.requirements
        if args.stufe:
            reqs = filter_mit_ahnen(reqs, lambda r: r.schutzbedarfsstufe == args.stufe)
        erzeuge_mappe_2023(snap, out, reqs=reqs, gefaehrdungen=_GEFAEHRDUNGSTABELLE)
    else:
        erzeuge_mappe(snap, out, filter_stufe=args.stufe)

    if getattr(args, "makro", False):
        # Zweiter Anlauf nach Analyse gegen eine echte, von Excel erzeugte
        # Datei (BSI-Vorlage A.3.4). Korrigiert wurden drei Abweichungen:
        #   1. _VBA_PROJECT-Version auf 0xFFFF -> Excel verwirft den
        #      P-Code-Cache und kompiliert aus dem Quelltext neu
        #   2. fehlender Record 0x004A (PROJECTCOMPATVERSION) ergaenzt
        #   3. PROJECTVERSION-Sonderfall: Length-Feld 4, Daten aber 6 Byte
        # Die .xlsx bleibt in jedem Fall erhalten und ist voll nutzbar.
        try:
            from .makro import VBA_CODE_AUTOOPEN, _vba_code, erzeuge_bas_datei, zu_xlsm
            from .vba_project import baue_vbaproject
            vbabin = baue_vbaproject(_vba_code() + VBA_CODE_AUTOOPEN)
            xlsm = out.with_suffix(".xlsm")
            zu_xlsm(out, xlsm, vbabin)
            bas = erzeuge_bas_datei(out.parent / "GSppAnsicht.bas")
            print(f"Makro-Variante: {xlsm}")
            print(f"Makro-Quelltext (Rueckfallebene): {bas}")
            print("  Falls Excel die .xlsm reklamiert: die .xlsx nutzen und "
                  "die .bas ueber Alt+F11 importieren (ANLEITUNG_MAKRO.md).")
        except Exception as exc:
            log.warning("Makro-Variante konnte nicht erzeugt werden (%s). "
                        "Die .xlsx ist davon unberuehrt und voll nutzbar.", exc)

    if args.snapshot_dir:
        _speichere_snapshot(snap, Path(args.snapshot_dir))
    anker = f"Commit {snap.commit_sha[:8]}" if snap.commit_sha \
        else f"SHA-256 {snap.sha256_quelle[:8]}"
    print(f"OK: {out}  ({len(snap.requirements)} Anforderungen, {anker})")
    return 0


def cmd_snapshot(args) -> int:
    snap = _snapshot(args)
    ziel = _speichere_snapshot(snap, Path(args.snapshot_dir))
    print(f"OK: {ziel}")
    return 0


def cmd_diff(args) -> int:
    alt = _lade_snapshot(Path(args.alt))
    neu = _lade_snapshot(Path(args.neu))
    bericht = diff_snapshots(alt, neu)
    print(bericht.zusammenfassung())

    if args.out_excel:
        print("Excel:", schreibe_excel(bericht, Path(args.out_excel)))
    if args.out_md:
        print("Markdown:", schreibe_markdown(bericht, Path(args.out_md)))

    if args.fail_on_change and bericht.aenderungen:
        return 2
    return 0


def cmd_watch(args) -> int:
    """Fuer den geplanten Lauf: Snapshot ziehen, gegen letzten vergleichen."""
    verzeichnis = Path(args.snapshot_dir)
    vorher = _neuester_snapshot(verzeichnis)
    neu = _snapshot(args)

    if vorher is None:
        _speichere_snapshot(neu, verzeichnis)
        print("Erster Snapshot angelegt - kein Vergleich moeglich.")
        return 0

    alt = _lade_snapshot(vorher)
    if alt.sha256_quelle == neu.sha256_quelle:
        print(f"Keine Aenderung (SHA-256 identisch zu {vorher.name}).")
        return 0

    bericht = diff_snapshots(alt, neu)
    _speichere_snapshot(neu, verzeichnis)
    print(f"Aenderung erkannt gegenueber {vorher.name}: {bericht.zusammenfassung()}")

    stempel = datetime.now(timezone.utc).strftime("%Y%m%d")
    schreibe_excel(bericht, verzeichnis / f"aenderungen_{stempel}.xlsx")
    schreibe_markdown(bericht, verzeichnis / f"aenderungen_{stempel}.md")
    return 2 if args.fail_on_change else 0


def cmd_doctor(args) -> int:
    """Umgebungspruefung vor dem ersten echten Lauf."""
    import sys as _sys

    ok = True
    print(f"Python            {_sys.version.split()[0]}", end="")
    if _sys.version_info < (3, 10):
        print("   FEHLT: 3.10+ erforderlich (Union-Syntax)")
        ok = False
    else:
        print("   ok")

    for modul, mindest in [("requests", None), ("pydantic", "2"), ("openpyxl", None)]:
        try:
            m = __import__(modul)
            ver = getattr(m, "__version__", getattr(m, "VERSION", "?"))
            passt = (mindest is None) or str(ver).startswith(mindest)
            print(f"{modul:<18}{ver}   {'ok' if passt else 'FALSCHE VERSION, ' + mindest + '.x noetig'}")
            ok = ok and passt
        except ImportError:
            print(f"{modul:<18}FEHLT   -> pip install -r requirements.txt")
            ok = False

    print()
    if args.katalog:
        pfad = Path(args.katalog)
        print(f"Lokaler Katalog   {'ok' if pfad.exists() else 'NICHT GEFUNDEN'}  {pfad}")
        ok = ok and pfad.exists()
    else:
        import requests as _rq
        from .fetch import API_COMMITS, RAW_URL
        for name, url in [("Katalog-Download", RAW_URL), ("GitHub-API", API_COMMITS)]:
            try:
                r = _rq.head(url, timeout=15, allow_redirects=True)
                hinweis = ""
                if r.status_code == 403 and "api" in url:
                    hinweis = "  (Ratelimit - unkritisch, siehe --token)"
                elif r.status_code >= 400:
                    hinweis = "  -> Proxy/Firewall pruefen oder --katalog nutzen"
                    if "api" not in url:
                        ok = False
                print(f"{name:<18}HTTP {r.status_code}{hinweis}")
            except Exception as exc:
                print(f"{name:<18}nicht erreichbar ({type(exc).__name__})"
                      f"{'  -> --katalog nutzen' if 'api' not in url else ''}")
                if "api" not in url:
                    ok = False

    print()
    print("Bereit." if ok else "Nicht bereit - siehe Meldungen oben.")
    return 0 if ok else 1


# --------------------------------------------------------------------- Parser


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="gspp", description="Grundschutz++ OSCAL -> Excel")
    p.add_argument("-v", "--verbose", action="store_true")
    sub = p.add_subparsers(dest="befehl", required=True)

    def quelle(sp):
        sp.add_argument("--katalog", help="lokale OSCAL-JSON statt GitHub-Download")
        sp.add_argument("--cache", default=".cache", help="Cache-Verzeichnis fuer Downloads")
        sp.add_argument("--token", help="GitHub-Token (nur gegen API-Ratelimit)")
        sp.add_argument("--mapping", help="lokale Mapping-JSON (IT-GS 2023 -> GS++) statt Download")
        sp.add_argument("--ohne-mapping", action="store_true", dest="ohne_mapping",
                        help="Zuordnungsspalten zu IT-GS 2023 weglassen")
        sp.add_argument("--gefaehrdungen", help="lokale basethreats.csv statt Download")
        sp.add_argument("--ohne-gefaehrdungen", action="store_true",
                        dest="ohne_gefaehrdungen",
                        help="Gefaehrdungsbeschreibungen und Nachschlageblatt weglassen")

    b = sub.add_parser("build", help="Excel-Vorlage erzeugen")
    quelle(b)
    b.add_argument("-o", "--out", default="out/Grundschutz++_Vorlage.xlsx")
    b.add_argument("--template", help="bestehende Vorlage befuellen statt neu erzeugen")
    b.add_argument("--blatt", default="Anforderungen", help="Zielblatt in der Vorlage")
    b.add_argument("--kopfzeile", type=int, default=1)
    b.add_argument("--stufe", choices=["normal-SdT", "erhöht"], help="nach Schutzbedarfsstufe filtern")
    b.add_argument("--ziel-schema", choices=["gspp", "2023"], default="gspp",
                   dest="ziel_schema",
                   help="gspp = katalognahes Layout; 2023 = Layout der BSI-Vorlage A.3.4")
    b.add_argument("--snapshot-dir", help="zusaetzlich Snapshot ablegen")
    b.add_argument("--makro", action="store_true",
                   help="zusaetzlich .xlsm mit Deckblatt-Steuerung erzeugen "
                        "(Spaltengruppierung bleibt als Rueckfallebene erhalten)")
    b.add_argument("--design", choices=["bsi", "msg"], default="bsi",
                   help="Designprofil: bsi = Referenzvorlage A.1 (Petrol #56A3BC), "
                        "msg = msg-Hausfarben (aktuell Platzhalterwerte)")
    b.set_defaults(func=cmd_build)

    s = sub.add_parser("snapshot", help="versionierten Snapshot ablegen")
    quelle(s)
    s.add_argument("--snapshot-dir", default="snapshots")
    s.set_defaults(func=cmd_snapshot)

    d = sub.add_parser("diff", help="zwei Snapshots vergleichen")
    d.add_argument("alt")
    d.add_argument("neu")
    d.add_argument("--out-excel")
    d.add_argument("--out-md")
    d.add_argument("--fail-on-change", action="store_true", help="Exitcode 2 bei Aenderungen (CI)")
    d.set_defaults(func=cmd_diff)

    d2 = sub.add_parser("doctor", help="Umgebung pruefen (vor dem ersten Lauf)")
    d2.add_argument("--katalog", help="lokale Datei statt Netzzugriff pruefen")
    d2.set_defaults(func=cmd_doctor)

    w = sub.add_parser("watch", help="Snapshot + Diff gegen letzten Stand")
    quelle(w)
    w.add_argument("--snapshot-dir", default="snapshots")
    w.add_argument("--fail-on-change", action="store_true")
    w.set_defaults(func=cmd_watch)

    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s %(name)s: %(message)s",
    )
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001
        log.error("Abbruch: %s", exc)
        if args.verbose:
            raise
        return 1


if __name__ == "__main__":
    sys.exit(main())
