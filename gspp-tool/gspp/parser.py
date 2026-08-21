"""
Parse-Layer: OSCAL-Baum -> flache Anforderungsliste.

Zu loesende Eigenheiten des GS++-Katalogs:
  * Zwei Gruppenebenen: Praktik (z.B. GC) -> Thema (GC.1) -> Control (GC.1.1)
  * Controls koennen Controls enthalten (348 von 999 sind verschachtelt)
  * Der Anforderungstext enthaelt Platzhalter "{{ insert: param, gc.1.1-prm1 }}",
    die gegen die params des Controls aufgeloest werden muessen
  * Attribute stecken teils am Control (props), teils am statement-Part (props)
"""
from __future__ import annotations

import logging
import re
from typing import Iterable

from .models import CatalogDocument, Control, Group, Param, Prop, Requirement

log = logging.getLogger(__name__)

INSERT_RE = re.compile(r"\{\{\s*insert:\s*param,\s*([^}\s]+)\s*\}\}")

# --- Baumsemantik -----------------------------------------------------------
# Der Katalog kennt zwei voellig verschiedene Eltern-Kind-Beziehungen. Die
# Unterscheidung laeuft ueber effort_level des Elternteils:
#
#   Aufwand 0  -> SAMMEL: Der Elternteil ist eine Klammer ohne eigenen Aufwand.
#                 Er ist erfuellt, wenn seine Teilanforderungen erfuellt sind
#                 (UND-Verknuepfung). Nicht eigenstaendig pruefbar - sonst
#                 zaehlt man dieselbe Leistung zweimal.
#                 Beispiel ASST.1.1 "Verfahren verankern"
#                          -> .1 dokumentieren / .2 zuweisen / .3 bekanntgeben
#
#   Aufwand >0 -> EIGEN: Der Elternteil ist selbst pruefbar; die Kinder sind
#                 zusaetzliche, meist auf ein engeres Zielobjekt verengte
#                 Anforderungen.
#                 Beispiel BES.2.1 "Bedarf dokumentieren" (Einkaeufe, Aufwand 2)
#                          -> .1 fuer IT-Produkte, .2 fuer Outsourcing (je eigener Aufwand)
#
# Verteilung im Stand 2026-07-16: 26x SAMMEL, 100x EIGEN, 873 Blaetter.
TYP_SAMMEL = "Sammelanforderung"
TYP_EIGEN = "Anforderung m. Teilanforderungen"
TYP_BLATT = "Einzelanforderung"

# Erwartete Attribute. Fehlt eines flaechendeckend, hat das BSI etwas umgebaut.
ERWARTETE_CONTROL_PROPS = {"sec_level", "effort_level", "alt-identifier"}
ERWARTETE_PART_PROPS = {"result", "action_word", "modal_verb"}


class SchemaAbweichung(RuntimeError):
    pass


def _prop(props: Iterable[Prop], name: str, default: str = "") -> str:
    for p in props:
        if p.name == name:
            return p.value
    return default


def _int_or_none(wert: str) -> int | None:
    try:
        return int(wert)
    except (TypeError, ValueError):
        return None


def _part(control: Control, name: str):
    for p in control.parts:
        if p.name == name:
            return p
    return None


def _loese_parameter(text: str, params: list[Param]) -> str:
    """Ersetzt {{ insert: param, id }} durch den Parameterwert."""
    if not text:
        return ""
    lookup = {p.id.lower(): (p.values[0] if p.values else (p.label or p.id)) for p in params}

    def ersetze(m: re.Match) -> str:
        pid = m.group(1).lower()
        if pid not in lookup:
            log.debug("Parameter %s nicht aufloesbar", pid)
            return m.group(0)
        return lookup[pid]

    return INSERT_RE.sub(ersetze, text)


def _links(control: Control, rel: str) -> str:
    return ", ".join(l.href.lstrip("#") for l in control.links if l.rel == rel)


def klassifiziere(ctrl: Control) -> tuple[str, bool]:
    """Liefert (Knotentyp, pruefpflichtig)."""
    if not ctrl.controls:
        return TYP_BLATT, True
    aufwand = _int_or_none(_prop(ctrl.props, "effort_level"))
    if aufwand in (0, None):
        return TYP_SAMMEL, False
    return TYP_EIGEN, True


def _control_zu_requirement(
    ctrl: Control,
    *,
    praktik: Group,
    thema: Group,
    parent_id: str | None,
    ebene: int,
    pfad: str,
) -> Requirement:
    knotentyp, pruefpflichtig = klassifiziere(ctrl)
    stm = _part(ctrl, "statement")
    gdn = _part(ctrl, "guidance")
    stm_props = stm.props if stm else []
    roh = stm.prose if stm and stm.prose else ""

    return Requirement(
        praktik_id=praktik.id or "",
        praktik_titel=praktik.title,
        thema_id=thema.id or "",
        thema_titel=thema.title,
        anforderung_id=ctrl.id,
        parent_id=parent_id,
        ebene=ebene,
        knotentyp=knotentyp,
        pruefpflichtig=pruefpflichtig,
        anzahl_teilanforderungen=len(ctrl.controls),
        pfad=pfad,
        titel=ctrl.title,
        anforderungstext=_loese_parameter(roh, ctrl.params),
        anforderungstext_roh=roh,
        erlaeuterung=(gdn.prose or "") if gdn else "",
        schutzbedarfsstufe=_prop(ctrl.props, "sec_level"),
        modalverb=_prop(stm_props, "modal_verb"),
        aufwand=_int_or_none(_prop(ctrl.props, "effort_level")),
        aktionswort=_prop(stm_props, "action_word"),
        ergebnis=_prop(stm_props, "result"),
        ergebnis_spezifikation=_prop(stm_props, "result_specification"),
        dokumentation=_prop(stm_props, "documentation"),
        zielobjekt_kategorien=_prop(stm_props, "target_object_categories"),
        vertraulichkeit=_int_or_none(_prop(ctrl.props, "confidentiality")),
        integritaet=_int_or_none(_prop(ctrl.props, "integrity")),
        verfuegbarkeit=_int_or_none(_prop(ctrl.props, "availability")),
        authentizitaet=_int_or_none(_prop(ctrl.props, "authenticity")),
        gefaehrdungen=_prop(ctrl.props, "threats"),
        tags=_prop(ctrl.props, "tags"),
        verweise_related=_links(ctrl, "related"),
        verweise_required=_links(ctrl, "required"),
        alt_identifier=_prop(ctrl.props, "alt-identifier"),
    )


def _sammle_controls(
    controls: list[Control],
    *,
    praktik: Group,
    thema: Group,
    parent_id: str | None = None,
    ebene: int = 0,
    pfad_prefix: str = "",
) -> list[Requirement]:
    """Tiefensuche - die Reihenfolge garantiert, dass Nachfahren zusammenhaengend
    unter ihrem Elternteil liegen. Darauf beruht die Rollup-Formel in Excel."""
    out: list[Requirement] = []
    for c in controls:
        pfad = f"{pfad_prefix} > {c.id}" if pfad_prefix else c.id
        out.append(
            _control_zu_requirement(
                c, praktik=praktik, thema=thema, parent_id=parent_id, ebene=ebene, pfad=pfad
            )
        )
        if c.controls:
            out.extend(
                _sammle_controls(
                    c.controls, praktik=praktik, thema=thema, parent_id=c.id,
                    ebene=ebene + 1, pfad_prefix=pfad,
                )
            )
    return out


def filter_mit_ahnen(reqs: list[Requirement], praedikat) -> list[Requirement]:
    """
    Filtert, behaelt aber alle Vorfahren der Treffer als Kontext.

    Ohne das erzeugt --stufe erhoeht 88 Waisen: Teilanforderungen mit erhoehtem
    Schutzbedarf unter Eltern mit normalem. Vorfahren, die selbst nicht auf das
    Praedikat passen, werden auf pruefpflichtig=False gesetzt - sie stehen als
    Lesehilfe da und verfaelschen den Erfuellungsgrad nicht.
    """
    nach_id = {r.anforderung_id: r for r in reqs}
    behalten: set[str] = set()
    treffer: set[str] = set()
    for r in reqs:
        if not praedikat(r):
            continue
        treffer.add(r.anforderung_id)
        rid: str | None = r.anforderung_id
        while rid:
            behalten.add(rid)
            rid = nach_id[rid].parent_id if rid in nach_id else None

    out: list[Requirement] = []
    for r in reqs:
        if r.anforderung_id not in behalten:
            continue
        if r.anforderung_id in treffer:
            out.append(r)
        else:
            out.append(r.model_copy(update={"pruefpflichtig": False,
                                            "knotentyp": TYP_SAMMEL}))
    return out


def parse(doc_dict: dict) -> tuple[CatalogDocument, list[Requirement]]:
    """Validiert den Katalog und liefert (Dokument, flache Anforderungsliste)."""
    doc = CatalogDocument.model_validate(doc_dict)
    cat = doc.catalog

    reqs: list[Requirement] = []
    for praktik in cat.groups:
        # Direkte Controls an der Praktik (aktuell 0, aber strukturell moeglich)
        if praktik.controls:
            reqs.extend(_sammle_controls(praktik.controls, praktik=praktik, thema=praktik))
        for thema in praktik.groups:
            reqs.extend(_sammle_controls(thema.controls, praktik=praktik, thema=thema))
            if thema.groups:
                log.warning(
                    "Unerwartete dritte Gruppenebene unter %s - wird flach eingehaengt.", thema.id
                )
                for unter in thema.groups:
                    reqs.extend(_sammle_controls(unter.controls, praktik=praktik, thema=unter))

    _plausibilitaet(reqs)
    log.info("Geparst: %d Praktiken, %d Anforderungen", len(cat.groups), len(reqs))
    return doc, reqs


def _plausibilitaet(reqs: list[Requirement]) -> None:
    """Frueherkennung von Schemaaenderungen: Attribute duerfen nicht flaechendeckend leer sein."""
    if not reqs:
        raise SchemaAbweichung("Katalog enthaelt keine Anforderungen.")

    n = len(reqs)
    pruefungen = {
        "schutzbedarfsstufe": sum(1 for r in reqs if r.schutzbedarfsstufe),
        "modalverb": sum(1 for r in reqs if r.modalverb),
        "anforderungstext": sum(1 for r in reqs if r.anforderungstext),
        "ergebnis": sum(1 for r in reqs if r.ergebnis),
    }
    for feld, treffer in pruefungen.items():
        if treffer / n < 0.9:
            raise SchemaAbweichung(
                f"Feld '{feld}' nur bei {treffer}/{n} Anforderungen belegt - "
                f"vermutlich hat sich die Katalogstruktur geaendert. Bitte Parser pruefen."
            )

    offen = [r.anforderung_id for r in reqs if INSERT_RE.search(r.anforderungstext)]
    if offen:
        log.warning("%d Anforderungen mit nicht aufloesbaren Parametern: %s",
                    len(offen), ", ".join(offen[:5]))

    ids = [r.anforderung_id for r in reqs]
    if len(ids) != len(set(ids)):
        doppelt = {i for i in ids if ids.count(i) > 1}
        raise SchemaAbweichung(f"Doppelte Anforderungs-IDs: {sorted(doppelt)[:10]}")
