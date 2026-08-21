# gspp-tool — Grundschutz++ OSCAL → Excel

Zieht den Anwenderkatalog Grundschutz++ aus dem BSI-GitHub, überführt ihn in
Excel-Vorlagen und erkennt Änderungen zwischen zwei Katalogständen.

Quelle: `BSI-Bund/Stand-der-Technik-Bibliothek`,
`Anwenderkataloge/Grundschutz++/Grundschutz++-catalog.json` (OSCAL 1.1.3).

---

## Was der Katalog tatsächlich enthält

Stand 2026-07-16, verifiziert gegen die Live-Datei (5,4 MB):

| Größe | Wert |
|---|---|
| Praktiken (Top-Level-Gruppen) | 20 |
| Themen (2. Ebene) | 139 |
| Anforderungen (Controls) | 999 — davon 348 als Unter-Controls verschachtelt |
| Schutzbedarfsstufen | 2 — `normal-SdT` (772), `erhöht` (227) |
| Modalverben | MUSS 149 · SOLLTE 626 · KANN 224 |
| Aufwandsstufen | 0–5 |
| Schutzziele je Anforderung | C/I/A/Authentizität, jeweils 0/1/2 |
| Gefährdungsbezug | Verweise auf elementare Gefährdungen (`G 0.18` …) |
| Querverweise | `related` (199) und `required` (67) |

Praktisch relevant für den Parser:

* **Zwei Gruppenebenen, dann Controls** — und Controls können Controls enthalten.
  Wer nur `catalog.groups[].controls` liest, verliert ein Drittel des Katalogs.
* **Attribute liegen an zwei Orten**: Schutzbedarfsstufe, Aufwand, Schutzziele,
  Gefährdungen am Control; Modalverb, Aktionswort, Ergebnis, Zielobjekt-Kategorien
  und Dokumentationsart am `statement`-Part.
* **Der Anforderungstext ist ein Template.** Er enthält
  `{{ insert: param, gc.1.1-prm1 }}`, das gegen `control.params` aufzulösen ist.
  Ohne Auflösung stehen Platzhalter in der Kundenvorlage.
* **Es gibt keine Release-Tags.** Der einzige belastbare Versionsanker ist der
  Commit-SHA plus `metadata.last-modified`.

---

## Ablauf

Es sind **nicht drei Skripte hintereinander**. Holen und Parsen passiert einmal
pro Katalogstand und erzeugt einen *Snapshot*. Excel-Ausgabe und Diff sind zwei
unabhängige Abnehmer dieses Snapshots:

```
                        einmal je Katalogstand
   GitHub ──fetch──> OSCAL-JSON ──parse──> Snapshot (JSON, versioniert)
                                              │
                       ┌──────────────────────┼──────────────────────┐
                       ▼                      ▼                      ▼
              Excel 2023-Schema       Excel GS++-nativ       diff gegen
              (A.3.4-Layout)          (katalognah)           Vor-Snapshot
                                                                    │
                                                                    ▼
                                                        Änderungsbericht
                                                        (.xlsx + .md)
```

Warum der Snapshot in der Mitte steht:

* **Der Diff braucht zwei Stände.** Beim ersten Lauf gibt es nichts zu
  vergleichen — er läuft erst ab dem zweiten. Ohne abgelegten Snapshot müsste
  man alte Katalogversionen aus der Git-Historie rekonstruieren.
* **Excel ist nur eine Darstellung.** Beide Layouts entstehen aus demselben
  Snapshot, ohne erneuten Download und ohne Risiko, dass zwei Ausgaben auf
  unterschiedlichen Katalogständen beruhen.
* **Der Snapshot ist das prüfbare Artefakt.** Er trägt Commit-SHA und SHA-256
  der Quelldatei. Im Audit ist er der Beleg, auf welchem Stand eine Bewertung
  beruht — eine Excel-Datei allein belegt das nicht.

In der Praxis:

```bash
# Erstmalig
python -m gspp.cli build --ziel-schema 2023 -o out/V.xlsx --snapshot-dir snapshots

# Danach wöchentlich (holt, vergleicht, schreibt Bericht nur bei Änderung)
python -m gspp.cli watch --snapshot-dir snapshots --fail-on-change
```

`watch` verkettet die Schritte für den automatisierten Lauf und vergleicht
zuerst den SHA-256 — bei unverändertem Katalog kostet der Durchlauf nichts.

---

## Architektur

```
        ┌──────────────┐
        │  fetch.py    │  GitHub-Raw + Commits-API · Offline-Modus · SHA-256 · Cache
        └──────┬───────┘
               │ dict
        ┌──────▼───────┐
        │  models.py   │  Pydantic-Modelle der OSCAL-Teilmenge
        │  parser.py   │  Baum → 999 flache Requirement-Objekte
        └──────┬───────┘         + Plausibilitätsprüfung (bricht bei Schemabruch ab)
               │ list[Requirement]
      ┌────────┴─────────┬──────────────────┐
┌─────▼──────┐   ┌───────▼──────┐   ┌───────▼──────┐
│  excel.py  │   │ CatalogSnap- │   │   diff.py    │
│ Template   │   │ shot (JSON)  │──▶│  + report.py │
│ / Generate │   │ versioniert  │   │ Änderungsber.│
└────────────┘   └──────────────┘   └──────────────┘
```

Die Trennung ist bewusst: Der Snapshot ist das revisionssichere Artefakt, Excel
ist nur eine Darstellung davon. Wenn später ein OSCAL-fähiges ISMS-Tool kommt,
bleibt alles außer `excel.py` bestehen.

### Warum Pydantic dazwischen

Das BSI ergänzt Inhalte laufend. `extra="allow"` toleriert neue Felder, aber
fehlende Pflichtfelder brechen den Lauf ab. Zusätzlich prüft `_plausibilitaet()`,
ob Kernattribute bei ≥90 % der Anforderungen belegt sind — ein stiller Umbau am
BSI-Schema erzeugt so einen Fehler statt einer halbleeren Kundenvorlage.

### Warum Template-Modus statt Generieren

`befuelle_template()` öffnet eine leergeräumte Kopie der bestehenden
2023er-Vorlage und schreibt nur Datenzeilen. Corporate Design, Spaltenbreiten,
bestehende Dropdowns und eigene Bewertungsspalten bleiben erhalten. Das Mapping
läuft über **Kopfzeilentexte**, nicht über Spaltenindizes — Spalten dürfen also
umsortiert werden, und unbekannte Kopftexte werden nicht angefasst:

```
Template-Mapping: 7 Spalten zugeordnet, 2 bleiben unberührt
                  (Kundenspezifische Bewertung, Interner Kommentar)
```

---

## Erzeugte Arbeitsmappe

| Blatt | Inhalt |
|---|---|
| `00_Metadaten` | Herkunftsnachweis: Quell-URL, Commit-SHA, SHA-256, Abrufzeitpunkt, Katalogversion. Der Audit-Beleg, auf welchem Stand die Bewertung beruht. |
| `01_Uebersicht` | Auswertung je Praktik über `COUNTIFS` — rechnet beim Ausfüllen mit. Erfüllungsgrad in Prozent. |
| `02_Anforderungen` | Vollständiger Katalog, 23 Spalten. |
| `03_Grundschutz-Check` | Arbeitsblatt: Katalogdaten + gelbe Eingabespalten mit Dropdown-Validierung (`entbehrlich/ja/teilweise/nein/offen`), Fristdatum, Verantwortlich, Begründung, Nachweis. |
| `04_Legende` | Modalverben, Schutzbedarfsstufen, Schutzziel-Skala, Bedienhinweise. |

Formeln sind mit LibreOffice gegengerechnet: 252 Formeln, 0 Fehler.

---

## Verwendung

```bash
pip install -r requirements.txt

# Umgebung prüfen (Python-Version, Pakete, Netzzugang)
python -m gspp.cli doctor

# Vollständige Vorlage aus dem Live-Katalog
python -m gspp.cli build -o out/GSpp_Vorlage.xlsx

# Bestehende Vorlage befüllen, nur erhöhter Schutzbedarf
python -m gspp.cli build \
    --template templates/msg_Vorlage_leer.xlsx \
    --blatt Anforderungskatalog --kopfzeile 3 \
    --stufe erhöht \
    -o out/msg_GSpp_erhoeht.xlsx

# Offline (abgeschottete Umgebung)
python -m gspp.cli build --katalog ./Grundschutz++-catalog.json -o out/V.xlsx

# Revisionsmanagement
python -m gspp.cli snapshot --snapshot-dir snapshots
python -m gspp.cli diff snapshots/alt.json snapshots/neu.json \
    --out-excel out/Aenderungsbericht.xlsx --out-md out/Aenderungsbericht.md

# Für den geplanten Lauf: holt, vergleicht, schreibt Bericht nur bei Änderung
python -m gspp.cli watch --snapshot-dir snapshots --fail-on-change
```

### Voraussetzungen

Python 3.10+ (wegen `str | None`-Syntax), die drei Pakete aus
`requirements.txt`, und für den Live-Abruf Zugriff auf
`raw.githubusercontent.com`. `api.github.com` ist optional — ohne die API fehlt
nur der Commit-SHA, als Versionsanker dient dann die SHA-256-Summe der
Quelldatei. Unauthentifiziert erlaubt GitHub 60 API-Anfragen pro Stunde und IP;
hinter einem Firmen-NAT ist das schnell erschöpft, dafür gibt es `--token`.

Ist GitHub gar nicht erreichbar: Datei manuell laden, `--katalog datei.json`.
`python -m gspp.cli doctor` prüft all das vorab.

`watch` vergleicht zuerst den SHA-256 — bei unverändertem Katalog kostet der Lauf
nichts. Exitcode 2 bei Änderungen macht ihn CI-tauglich.

Beispielausgabe des Änderungsberichts:

```
1 neu, 1 entfallen, 4 geaendert (4 davon kritisch)
```

Die Gewichtung ist in `diff.FELD_GEWICHT` hinterlegt:
**kritisch** (Text, Modalverb, Schutzbedarfsstufe, Titel, required-Verweise) →
Umsetzung neu bewerten · **relevant** (Ergebnis, Zielobjekte, Aufwand,
Schutzziele, Gefährdungen) → prüfen · **redaktionell** (Erläuterung, Tags) →
informativ.

---

## Offene Punkte für die Austauschplattform

1. **Praktiken ↔ Bausteine.** 20 prozessorientierte Praktiken lassen sich nicht
   sauber auf 111 zielobjektorientierte Bausteine abbilden. Die Kandidaten für
   eine Brücke sind `target_object_categories` (60 verschiedene Werte) und die
   `tags`. Ob wir eine eigene Mapping-Tabelle pflegen oder auf ein BSI-Mapping
   warten, ist eine Produktentscheidung, keine technische.
2. **Drei → zwei Schutzbedarfsstufen.** Bestehende Kundenbewertungen nach
   B/S/H brauchen eine dokumentierte Überleitungsregel.
3. **Verschachtelte Anforderungen** — gelöst, siehe eigenen Abschnitt unten.
   Zur Diskussion steht nur noch, ob die Heuristik `effort_level == 0` durch
   eine gepflegte Ausnahmeliste ergänzt werden soll.
4. **Zielformat.** Excel ist die Kundensicht. Ob wir mittelfristig OSCAL-nativ
   arbeiten (Assessment-Plan / Assessment-Results statt Excel-Rückkanal),
   entscheidet, wie viel wir in die Excel-Schiene investieren.

---

## Ausgabe im 2023-Schema (`--ziel-schema 2023`)

Erzeugt das Layout der BSI-Vorlage **A.3.4** (Version 1.1.1, Edition 2023):
Kopfzeile in Zeile 6, zwölf Spalten `Anforderung | Beschreibung |
Teilanforderung | Schutz | Anforderungstext | Umsetzung | Nachweis | Status |
Maßnahme | Befragte Person | Prüfende Person | Datum der Prüfung`, ein Blatt je
Praktik analog zu einem Blatt je Baustein, dazu Deckblatt, Praktiken-Übersicht
(entspricht `Modellierung`), Dashboard und Hilfstabelle.

**Die 2023er-Vorlage nutzt bereits dieselbe Eltern-Kind-Mechanik wie GS++:**
Die Anforderungszeile (`APP.3.3.A2`) hat kein Eingabefeld, ihr Status wird per
Formel aus den Teilanforderungen (`A2.1`–`A2.4`) aggregiert. Vokabular und
Aggregationsregel sind 1:1 von dort übernommen — inklusive `zu klären`,
`verwiesen` und `Bitte Teilanforderungen ausfüllen`.

Strukturabbildung:

| 2023 | GS++ |
|---|---|
| Anforderungszeile | jeder Knoten auf Ebene 0 — **651 Zeilen** |
| Teilanforderung | jeder prüfpflichtige Knoten darunter — **973 Zeilen** |

Sammelanforderungen erscheinen nicht als Teilanforderung; ihre Kinder hängen
direkt an der Anforderungszeile. Damit bleibt die Zahl bewertbarer Zeilen exakt
bei 973.

Größenvergleich mit der Originalvorlage:

| | 2023 | GS++ |
|---|---|---|
| Blätter | 121 Bausteine | 20 Praktiken |
| Anforderungen | 2.508 | 651 |
| **Bewertbare Zeilen** | **7.889** | **973** |

### Offene fachliche Annahme

Die Spalte `Schutz` kennt in 2023 drei Kategorien, GS++ nur zwei
Schutzbedarfsstufen. Abgebildet wird:

```
erhöht                      -> Anforderungen bei erhöhtem Schutzbedarf
normal-SdT + MUSS           -> Basis-Anforderungen
normal-SdT + SOLLTE/KANN    -> Standard-Anforderungen
```

Das ist **nicht vom BSI dokumentiert**. Die Annahme steht deshalb auch auf dem
Deckblatt der erzeugten Mappe und ist fachlich zu bestätigen.

---

## Umgang mit verschachtelten Anforderungen

348 der 999 Anforderungen sind Unter-Controls. Naiv flach ausgerollt zählt man
dieselbe Leistung zweimal und der Erfüllungsgrad wird geschönt. Die Auswertung
des Katalogs zeigt aber, dass es **zwei verschiedene Eltern-Kind-Beziehungen**
gibt, die unterschiedlich behandelt werden müssen.

### Die Unterscheidung

Das trennende Merkmal ist `effort_level` des Elternteils:

**Typ A — Sammelanforderung** (`effort_level == 0`, 26 Fälle)

Der Elternteil ist eine Klammer ohne eigenen Aufwand; die Kinder sind die
konstitutiven Schritte. UND-Verknüpfung.

```
ASST.1.1  Verfahren und Regelungen verankern            Aufwand 0
  ├─ .1   … dokumentieren                               Aufwand 0
  ├─ .2   … Aufgaben zuständigen Rollen zuweisen        Aufwand 0
  └─ .3   … zuständige Rollen informieren               Aufwand 0
```

**Typ B — eigenständige Anforderung mit Teilanforderungen** (`effort_level > 0`, 100 Fälle)

Der Elternteil ist selbst prüfbar; die Kinder sind zusätzliche, meist auf ein
engeres Zielobjekt verengte Anforderungen mit eigenem Aufwand.

```
BES.2.1   Bedarf dokumentieren  (für Einkäufe)          Aufwand 2
  ├─ .1   … Verwendungszweck    (für IT-Produkte)       Aufwand 3
  ├─ .2   … Geschäftsprozessprofile (für Outsourcing)   Aufwand 4
  └─ .3   … Systemvoraussetzungen (für IT-Produkte)     Aufwand 3
```

Das Subjekt des Satzes verengt sich beim Kind — „Beschaffungsmanagement für
Einkäufe" wird zu „… für IT-Produkte". Das ist eine Präzisierung, keine Zerlegung.

Bestätigt wird die Lesart durch eine Monotonie-Eigenschaft: Kinder verschärfen
die Verbindlichkeit nie. MUSS→SOLLTE und SOLLTE→KANN kommen vor, SOLLTE→MUSS nie.

### Umsetzung

| Knotentyp | Statuszelle | in Zählung |
|---|---|---|
| Sammelanforderung | **grau, Formel** — aggregiert aus Teilanforderungen | nein |
| Anforderung m. Teilanforderungen | gelb, Eingabe | ja |
| Einzelanforderung | gelb, Eingabe | ja |

Die Spalte `Prüfpflichtig` ist das Kriterium, über das `01_Uebersicht` zählt.
Ergebnis: **973 prüfpflichtige Anforderungen** statt 999 — die Differenz sind
exakt die 26 Sammelknoten.

Aggregationsregel für Sammelknoten (`_rollup_formel`):

```
ein  "nein"                → nein
ein  "teilweise"           → teilweise
alle "ja"/"entbehrlich"    → ja
teils bewertet             → teilweise
nichts bewertet            → offen
```

Die Formel greift auf einen zusammenhängenden Zeilenbereich zu. Das funktioniert,
weil der Parser tiefensuchend ausgibt — Nachfahren liegen garantiert direkt unter
ihrem Knoten. Verifiziert: 2 von 3 Kindern erfüllt → `teilweise`; alle drei →
`ja`; eines `entbehrlich` → `ja`; eines `nein` → `nein`.

Zusätzlich:

* **Excel-Gliederungsebenen** (`outlineLevel`) — ganze Anforderungsbäume lassen
  sich ein- und ausklappen. Titel zusätzlich je Ebene eingerückt.
* **Spalte `Pfad`** (`GC.9.1 > GC.9.1.1 > GC.9.1.1.1`) für Filter und Pivots,
  wo Einrückung nicht trägt.
* **Waisen beim Filtern.** 88 Teilanforderungen mit `erhöht` hängen unter Eltern
  mit `normal-SdT`. `filter_mit_ahnen()` behält die Vorfahren als Kontextzeilen
  und setzt sie auf `prüfpflichtig = nein` — sie sind Lesehilfe und verfälschen
  den Erfüllungsgrad nicht. Bei `--stufe erhöht`: 280 Zeilen, davon 227
  prüfpflichtig und 53 Kontext.

### Grenze der Heuristik

`effort_level == 0` ist ein Indiz, keine vom BSI dokumentierte Semantik. Bei 26
Fällen ist eine manuelle Durchsicht in etwa einer Stunde machbar; bestätigte
Abweichungen gehören dann in eine Ausnahmeliste in `parser.klassifiziere()`.
Der Diff meldet Änderungen an `aufwand` als *relevant* — ein Knoten, der von 0
auf >0 wechselt, wird also nicht stillschweigend umklassifiziert.

---

## Zuordnung zu IT-Grundschutz 2023 (offizielles BSI-Mapping)

Das BSI veröffentlicht seit Kurzem eine requirement-genaue Zuordnungstabelle
zwischen dem alten IT-Grundschutz-Kompendium 2023 und Grundschutz++
(`control_layer/Mappings/IT-GS2023-zu-GSpp/`). 1013 Einzelzuordnungen, fünf
OSCAL-Beziehungstypen (`equal-to`, `equivalent-to`, `subset-of`,
`superset-of`, `intersects-with`), 718 alte auf 306 neue Anforderungen.

Beide Excel-Layouts bekommen dadurch zwei zusätzliche Spalten:

| Spalte | Inhalt |
|---|---|
| Alte Anforderung (IT-GS 2023) | z. B. `NET.1.1.A16-UA.2; NET.1.1.A17-UA.2; …` |
| Beziehung (IT-GS 2023) | positionsgleich, z. B. `superset-of; superset-of; …` |

**Abdeckung ist unvollständig — kein Fehler.** Nur ca. 30 % der GS++-Anforderungen
haben aktuell eine Zuordnung; das BSI selbst bezeichnet das Mapping als in
Pilotierung befindlich. Eine leere Zelle heißt „noch nicht kartiert", nicht
„keine Beziehung". Steht auch auf dem Deckblatt jeder erzeugten Datei.

Automatisch aktiv, kein zusätzliches Flag nötig. Abschalten mit
`--ohne-mapping`, offline mit `--mapping datei.json` (gleiches Prinzip wie
`--katalog`). Ein Fehlschlag beim Laden der Mapping-Datei blockiert den Lauf
nicht — die Vorlage entsteht dann einfach ohne diese beiden Spalten.

Ein KPI-/Punktesystem gibt es dagegen (noch) nicht: Der dafür vorgesehene
`assessment_layer` im BSI-Repository ist offiziell als leer markiert
("Derzeit sind in diesem Bereich noch keine Artefakte veröffentlicht").

## Spalten ein-/ausblenden (Excel-Gruppierung)

Über Excels native Spaltengruppierung — kleine `+`/`−` Schaltflächen über den
Spaltenköpfen. Kein VBA, keine Makro-Sicherheitswarnung, funktioniert in jeder
Excel-Version. **Zwei unabhängig steuerbare Blöcke:**

| Block | Spalten | Startzustand |
|---|---|---|
| Guidance | Erläuterung, Zielobjekt-Kategorien, Dokumentation, Ergebnis, Gefährdungen, Aufwand | **eingeklappt** |
| Zuordnung IT-GS 2023 | Alte Anforderung, Beziehung | sichtbar |

Guidance startet eingeklappt, damit die Vorlage beim Öffnen nah an der
gewohnten A.3.4-Ansicht ist — ein Klick auf `+` blendet den kompletten
Erklärungsblock ein. Die Bedienung ist zusätzlich auf dem Deckblatt erklärt.

Im katalognahen Layout (`--ziel-schema gspp`) sind Erläuterung sowie die
beiden Zuordnungsspalten ebenfalls gruppiert.

**Warum keine Checkboxen:** Echte Ankreuzfelder, die Spalten ein- und
ausblenden, brauchen zwingend VBA-Makros. Das bedeutet `.xlsm`-Format,
Makro-Sicherheitswarnung bei jedem Öffnen und je nach Gruppenrichtlinie
gar keine Ausführung. Die Gruppierung leistet dasselbe ohne diese Nachteile.
Falls Checkboxen dennoch gewünscht sind, ist das umsetzbar — dann aber
bewusst als Makro-Variante.

## Designprofile

Alle Farb-, Schrift- und Layoutwerte liegen zentral in `gspp/design.py`.
Umschalten über `--design bsi` (Vorgabe) oder `--design msg`.

Das Profil **BSI** ist aus der Referenzdatei
`A_1_Vorlage_Strukturanalyse_1_1_0.xlsx` ausgelesen, nicht geschätzt:

| Merkmal | Wert |
|---|---|
| Schrift | Aptos, 11 pt (Daten) / 12 pt (Deckblatt) |
| Akzentfarbe | `#56A3BC` (Petrolblau, Kopfzeilen) |
| Sekundärtext | `#6F6F6F` |
| Zellrahmen | keine |
| Gitternetzlinien | ausgeblendet |
| Randspalten | links/rechts je 2,1 Breite |

Das Profil **msg** enthält derzeit **Platzhalterwerte** (`#9C1006`) — die
offiziellen CI-Hexcodes fehlen noch. Vor Kundenauslieferung ersetzen; es sind
genau zwei Konstanten in `design.py`.

**Bewusste Abweichung von der Referenz:** Eingabefelder und berechnete Zellen
werden dezent hinterlegt (`#FDF6E3` bzw. `#F0F0F0`). Das Original kommt ohne
aus, weil dort jede Zelle im Datenbereich ein Eingabefeld ist. Im
Grundschutz-Check stehen Katalogdaten, Eingabefelder und Rollup-Formeln
nebeneinander — ohne Unterscheidung würde in die falschen Zellen geschrieben.

## Elementare Gefährdungen (G 0.x)

Der Katalog nennt in der `threats`-Eigenschaft nur Kürzel wie `G 0.18` und
verweist per Namespace-URL auf `documentation/namespaces/basethreats.csv`.
Diese Datei wird automatisch mitgeladen — 47 Gefährdungen mit Kurzbegriff
und ausführlicher Definition.

**In der Anforderungszeile** steht dadurch statt `G 0.18, G 0.19` jetzt:

```
G 0.18 – Fehlplanung oder fehlende Anpassung; G 0.19 – Offenlegung schützenswerter Informationen
```

**Blatt „Gefährdungen"** listet die vollständigen Definitionstexte einmalig
auf — nur die im Katalog tatsächlich referenzierten (aktuell 43 von 47).
Die Texte in jede Zeile zu kopieren wäre sinnlos: dieselbe Gefährdung taucht
in dutzenden Anforderungen auf.

Automatisch aktiv. Abschalten mit `--ohne-gefaehrdungen`, offline mit
`--gefaehrdungen basethreats.csv`. Schlägt das Laden fehl, bleiben die
Kürzel stehen — der Lauf bricht nicht ab.

Hinweis: 99 Anforderungen haben im Katalog gar keine Gefährdungsangabe,
darunter die komplette Praktik GC. Leere Zellen sind dort korrekt.

## Punkte je Schutzziel

Blatt „Praktiken" und Dashboard summieren die Schutzzielwerte (0–2) je
Praktik bzw. insgesamt. Gezählt wird **nur bei Status „ja"** — entbehrliche
und verwiesene Anforderungen zählen nicht mit, weil der Schutzbeitrag dort
nicht tatsächlich erbracht wird (Festlegung Sven, 07.08.).

Sammelzeilen sind über das Kriterium `Teilanforderung <> ""` ausgeschlossen,
sonst würde jede Anforderung doppelt gewichtet.

## Bedingte Formatierung der Statusspalte

Ampelfarben per Excel-Regel, kein Makro. Farben zentral in
`design.py` → `status_farben`:

| Status | Farbe |
|---|---|
| ja | grün |
| nein | rot |
| teilweise | gelb |
| entbehrlich | grau |
| verwiesen | blau |
| zu klären | orange |
| unbearbeitet | keine Füllung |

Als Regel statt fester Zellfüllung umgesetzt: dadurch färben sich auch
Sammelzeilen, deren Status aus der Rollup-Formel stammt.

## Zielobjekt und Zuständig

Werden zentral im Blatt „Praktiken" gepflegt. Die Praktikblätter spiegeln
sie per Formel (grau = berechnet, nicht eintippen) — keine Doppelpflege.

## Nicht enthalten (bewusst)

* Rückkanal Excel → OSCAL (ausgefüllter Check als `assessment-results`)
* Mapping-Layer auf ISO 27001 / NIS-2 / C5
* Word-Export der Anforderungen
