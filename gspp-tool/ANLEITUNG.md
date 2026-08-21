# Eigener Durchgang — Schritt für Schritt

Getestet aus dem ausgelieferten ZIP heraus, Python 3.12 unter Linux.
Windows-Abweichungen stehen jeweils dabei.

---

## Voraussetzungen

| | |
|---|---|
| Python | **3.10 oder neuer** (das Tool nutzt `str \| None`-Syntax) |
| Pakete | `requests`, `pydantic 2.x`, `openpyxl` |
| Netz | `raw.githubusercontent.com` — zwingend, außer im Offline-Modus |
| Netz | `api.github.com` — optional, liefert nur den Commit-SHA |

Prüfen, was installiert ist:

```bash
python3 --version        # Linux/macOS
py --version             # Windows
```

Ist die Version älter als 3.10, hilft nur eine neuere Installation —
Rückportieren lohnt nicht.

---

## Schritt 1 — Entpacken

```bash
unzip gspp-tool.zip
cd gspp-tool
```

Im Paket liegen bereits `out/` und `snapshots/` mit meinen Ergebnissen. Die
kannst du als Vergleichsmaßstab behalten oder löschen — deine Läufe schreiben
ohnehin in dieselben Verzeichnisse und überschreiben nur gleichnamige Dateien.

---

## Schritt 2 — Virtuelle Umgebung

Nicht zwingend, aber sinnvoll: `pydantic 2.x` kollidiert sonst womöglich mit
etwas, das schon systemweit installiert ist.

```bash
python3 -m venv .venv
source .venv/bin/activate          # Linux/macOS
.venv\Scripts\activate             # Windows (cmd)
.venv\Scripts\Activate.ps1         # Windows (PowerShell)
```

Scheitert das unter Debian/Ubuntu mit „ensurepip is not available":
`sudo apt install python3-venv`.

---

## Schritt 3 — Pakete installieren

```bash
pip install -r requirements.txt
```

Hinter einem Firmen-Proxy:

```bash
pip install -r requirements.txt --proxy http://proxy.firma.de:8080
```

---

## Schritt 4 — Umgebung prüfen

```bash
python -m gspp.cli doctor
```

Erwartete Ausgabe:

```
Python            3.12.3   ok
requests          2.33.1   ok
pydantic          2.13.4   ok
openpyxl          3.1.5    ok

Katalog-Download  HTTP 200
GitHub-API        HTTP 403  (Ratelimit - unkritisch, siehe --token)

Bereit.
```

`HTTP 403` bei der GitHub-API ist **kein Fehler**. Unauthentifiziert erlaubt
GitHub 60 Anfragen pro Stunde und IP; hinter einem Firmen-NAT teilen sich das
viele Kolleginnen und Kollegen. Der Katalog-Download ist davon nicht betroffen.
Es fehlt dann lediglich der Commit-SHA im Herkunftsnachweis — als Versionsanker
dient stattdessen die SHA-256-Summe der Quelldatei.

Steht bei **Katalog-Download** etwas anderes als `HTTP 200`, weiter bei
Schritt 5b.

---

## Schritt 5a — Erster Lauf (mit Netz)

```bash
python -m gspp.cli build \
    --ziel-schema 2023 \
    -o out/GSpp_Vorlage_2023-Schema.xlsx \
    --snapshot-dir snapshots \
    --cache .cache
```

Windows (eine Zeile, oder `^` statt `\` als Fortsetzungszeichen):

```cmd
python -m gspp.cli build --ziel-schema 2023 -o out\GSpp_Vorlage_2023-Schema.xlsx --snapshot-dir snapshots --cache .cache
```

Erwartete Ausgabe:

```
INFO  gspp.fetch: Lade Katalog von GitHub (commit ...)
INFO  gspp.parser: Geparst: 20 Praktiken, 999 Anforderungen
INFO  gspp.schema2023: 2023-Schema geschrieben: ... (651 Anforderungszeilen,
      973 Teilanforderungen, 20 Praktik-Blätter)
INFO  gspp: Snapshot: snapshots/snapshot_JJJJMMTT_xxxxxxxxxxxx.json
OK:   out/GSpp_Vorlage_2023-Schema.xlsx  (999 Anforderungen, SHA-256 xxxxxxxx)
```

**Stimmen 20 / 999 / 651 / 973 nicht?** Dann hat das BSI den Katalog geändert.
Das ist kein Fehler, sondern genau der Fall, für den das Diffing existiert —
siehe Schritt 8.

### Mit GitHub-Token (empfohlen im Firmennetz)

Ein Personal Access Token ohne jeden Scope reicht; es geht nur um das
Ratelimit (60 → 5.000 Anfragen/Stunde). Auf github.com unter
*Settings → Developer settings → Personal access tokens → Fine-grained*
erzeugen, „Public repositories (read-only)" genügt.

```bash
python -m gspp.cli build --ziel-schema 2023 -o out/V.xlsx \
    --snapshot-dir snapshots --token ghp_xxxxxxxxxxxx
```

Dann steht der echte Commit-SHA im Herkunftsnachweis.

---

## Schritt 5b — Erster Lauf (ohne Netz)

Für abgeschottete Umgebungen. Katalogdatei einmalig auf einem Rechner mit
Internetzugang holen:

```
https://raw.githubusercontent.com/BSI-Bund/Stand-der-Technik-Bibliothek/main/Anwenderkataloge/Grundschutz%2B%2B/Grundschutz%2B%2B-catalog.json
```

Datei übertragen, dann:

```bash
python -m gspp.cli build --ziel-schema 2023 \
    --katalog ./Grundschutz++-catalog.json \
    -o out/GSpp_Vorlage_2023-Schema.xlsx \
    --snapshot-dir snapshots
```

Das Deckblatt weist die Datei dann korrekt als lokale Quelle aus.

---

## Schritt 6 — Ergebnis prüfen

Öffne `out/GSpp_Vorlage_2023-Schema.xlsx`. Vier Kontrollen, in dieser
Reihenfolge:

**1. Deckblatt.** Katalog-Version, SHA-256, Erzeugungszeitpunkt. Ganz unten
steht die offene fachliche Annahme zur Spalte `Schutz` — die solltest du gelesen
haben, bevor die Datei jemand anderes sieht.

**2. Blatt `GC`, Zeile 6 abwärts.** Layout muss der 2023er-Vorlage entsprechen:
`Anforderung | Beschreibung | Teilanforderung | Schutz | Anforderungstext | …`
Sammelzeilen erkennst du an leerer Spalte C und `Gesamt:` in Spalte E.

**3. Rollup testen.** Such im Blatt `GC` die Sammelzeile `GC.5.1`; darunter
liegen vier Teilanforderungen. Setz in Spalte H (`Status`) zwei davon auf `ja`
— die Sammelzeile muss auf `teilweise` springen. Alle vier auf `ja` → `ja`.
Eine auf `zu klären` → `zu klären`.

**4. Blatt `Praktiken`.** Setz Spalte C (`Relevanz`) auf `Ja`. Die Zähler und
der Umsetzungsstand müssen anspringen, und das `Dashboard` mitziehen.

Falls die Zellen nach dem Öffnen leer wirken: Excel neu berechnen lassen
(`Strg + Alt + F9`). openpyxl schreibt Formeln ohne zwischengespeicherte
Ergebnisse.

---

## Schritt 7 — Weitere Varianten

```bash
# Katalognahes Layout: mehr Attribute je Zeile (Schutzziele, Gefährdungen,
# Querverweise), die im 2023er Zwölf-Spalten-Korsett keinen Platz haben
python -m gspp.cli build -o out/GSpp_katalognah.xlsx

# Nur erhöhter Schutzbedarf (227 Anforderungen + 53 Kontextzeilen)
python -m gspp.cli build --ziel-schema 2023 --stufe erhöht -o out/GSpp_erhoeht.xlsx

# Eine eigene msg-Vorlage befüllen statt neu erzeugen.
# Das Mapping läuft über die Kopfzeilentexte; eigene Spalten bleiben unberührt.
python -m gspp.cli build \
    --template vorlagen/msg_Vorlage_leer.xlsx \
    --blatt Anforderungskatalog --kopfzeile 3 \
    -o out/msg_GSpp.xlsx
```

---

## Schritt 8 — Änderungen verfolgen

Der Diff braucht **zwei** Snapshots und läuft deshalb erst ab dem zweiten
Durchgang.

```bash
# Holen, mit letztem Stand vergleichen, Bericht nur bei Änderung schreiben
python -m gspp.cli watch --snapshot-dir snapshots
```

Unverändert:

```
Keine Aenderung (SHA-256 identisch zu snapshot_20260723_e4c66cbc76fb.json).
```

Verändert: es entstehen `snapshots/aenderungen_JJJJMMTT.xlsx` und `.md`, dazu
eine Zeile wie

```
Aenderung erkannt gegenueber snapshot_...: 1 neu, 0 entfallen, 1 geaendert
(2 davon kritisch)
```

Zwei beliebige Stände vergleichen:

```bash
python -m gspp.cli diff snapshots/alt.json snapshots/neu.json \
    --out-md out/Aenderungen.md --out-excel out/Aenderungen.xlsx
```

Mit `--fail-on-change` liefert der Befehl Exitcode 2 bei Änderungen — dadurch
lässt er sich als geplanter Task oder CI-Job auswerten (siehe
`.github/workflows/katalog-watch.yml`).

---

## Wenn etwas klemmt

| Meldung | Ursache und Abhilfe |
|---|---|
| `ModuleNotFoundError: gspp` | Falsches Verzeichnis. Muss dort laufen, wo der Ordner `gspp/` liegt. |
| `SSLError` / `ProxyError` | Firmen-Proxy. `HTTPS_PROXY` setzen oder Schritt 5b (Offline) nutzen. |
| `Commit-Metadaten nicht abrufbar` | Ratelimit. Unkritisch — der Katalog kommt trotzdem, nur der Commit-SHA fehlt. `--token` behebt es. |
| `Download fehlgeschlagen: HTTP 404` | Das BSI hat die Datei verschoben. Pfad in `gspp/fetch.py`, Konstante `KATALOG_PFAD`, prüfen. |
| `Feld 'x' nur bei n/999 Anforderungen belegt` | Das BSI hat die Katalogstruktur geändert. Absichtlicher Abbruch — der Parser muss angepasst werden, bevor Excel entsteht. |
| `Doppelte Anforderungs-IDs` | Ebenfalls Schemaänderung. Meldung enthält die betroffenen IDs. |
| `Keine Spaltenköpfe der Vorlage konnten zugeordnet werden` | Bei `--template`: falsche `--kopfzeile` oder abweichende Spaltenüberschriften. Die Meldung listet, was gefunden wurde. |
| Excel zeigt leere Formelzellen | `Strg + Alt + F9`. openpyxl schreibt keine zwischengespeicherten Ergebnisse. |

---

## Was NICHT tun

**Die Original-`.xlsm` des BSI nicht durch das Tool schreiben lassen.** Sie
enthält ein 322 KB großes VBA-Projekt und zwei `XLOOKUP`-Formeln auf dem
Deckblatt. Ein Schreibvorgang mit openpyxl ohne `keep_vba=True` verwirft die
Makros; ein anschließender LibreOffice-Durchlauf zerstört die `XLOOKUP`-Zellen.
Der `--template`-Modus ist für **leergeräumte Kopien eurer eigenen Vorlagen**
gedacht, nicht für die Original-Hilfsmitteldatei.
