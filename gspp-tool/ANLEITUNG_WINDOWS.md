# gspp-tool unter Windows starten — Schritt für Schritt

Für Windows 10/11, ohne Vorkenntnisse in der Kommandozeile. Jeder Schritt ist
einzeln nachvollziehbar; wenn einer nicht klappt, findest du unten eine
passende Fehlerbehebung.

---

## Schritt 1 — Python installieren (falls noch nicht vorhanden)

Prüfen, ob Python schon da ist: **Startmenü → „cmd" eingeben → Eingabeaufforderung öffnen**, dann:

```
py --version
```

Kommt eine Versionsnummer **3.10 oder höher** zurück (z. B. `Python 3.12.3`), weiter mit Schritt 2.

Kommt `'py' ist nicht als interner oder externer Befehl erkannt`, Python fehlt:

1. Auf **python.org/downloads** die aktuelle Version herunterladen (nicht den Microsoft-Store-Eintrag — der macht bei PATH gelegentlich Ärger).
2. Installer starten. **Ganz wichtig, unten im ersten Fenster:** Häkchen bei **„Add python.exe to PATH"** setzen, bevor auf „Install Now" geklickt wird. Das ist der häufigste Stolperstein — ohne dieses Häkchen findet die Kommandozeile Python später nicht.
3. Nach der Installation ein **neues** Eingabeaufforderungsfenster öffnen (das alte kennt die PATH-Änderung noch nicht) und `py --version` erneut prüfen.

---

## Schritt 2 — ZIP entpacken

1. `gspp-tool.zip` in den Windows-Explorer legen, z. B. nach `Dokumente`.
2. Rechtsklick auf die Datei → **„Alle extrahieren…"** → Zielordner bestätigen → **Extrahieren**.
3. Ergebnis ist ein Ordner `gspp-tool` mit den Unterordnern `gspp`, `out`, `snapshots` und der Datei `requirements.txt`.

Wichtig: **nicht** direkt aus dem gezippten Zustand heraus arbeiten (Doppelklick auf die ZIP-Datei öffnet nur eine Voransicht) — erst „Alle extrahieren" ausführen.

---

## Schritt 3 — Eingabeaufforderung im richtigen Ordner öffnen

Der bequemste Weg:

1. Im Explorer in den entpackten Ordner `gspp-tool` hineinklicken (sodass `gspp`, `out`, `requirements.txt` sichtbar sind).
2. In die Adressleiste oben im Explorer-Fenster klicken (dort wo der Pfad steht), `cmd` eingeben, **Enter**.

Es öffnet sich eine Eingabeaufforderung, die bereits im richtigen Ordner steht — erkennbar am Pfad vor dem Blinkzeichen, z. B. `C:\Users\deutsj\Dokumente\gspp-tool>`.

*(Alternative: Eingabeaufforderung normal öffnen und `cd Dokumente\gspp-tool` eintippen — der genaue Pfad hängt davon ab, wo entpackt wurde.)*

---

## Schritt 4 — Virtuelle Umgebung anlegen

Das hält die für dieses Tool installierten Pakete von allem anderen auf dem Rechner getrennt. In der Eingabeaufforderung, die aus Schritt 3 offen ist:

```
py -m venv .venv
```

Das dauert ein paar Sekunden und erzeugt einen neuen Unterordner `.venv`. Danach aktivieren:

```
.venv\Scripts\activate
```

Am Anfang der Zeile erscheint jetzt `(.venv)` — das zeigt, dass die aktivierte Umgebung genutzt wird.

**Falls stattdessen PowerShell verwendet wird** (statt der klassischen Eingabeaufforderung — erkennbar an blauer statt schwarzer Konsole) und eine Fehlermeldung über „Ausführung von Skripts deaktiviert" erscheint:

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.venv\Scripts\Activate.ps1
```

Das setzt die Einschränkung nur für dieses eine Fenster herab, dauerhaft und systemweit ändert sich nichts.

---

## Schritt 5 — Pakete installieren

Mit weiterhin aktivierter `(.venv)`-Umgebung:

```
pip install -r requirements.txt
```

Das lädt drei kleine Pakete herunter (`requests`, `pydantic`, `openpyxl`) und dauert normalerweise unter einer Minute.

**Hinter einem Firmen-Proxy**, falls eine Fehlermeldung mit `ProxyError` oder `SSLError` erscheint:

```
pip install -r requirements.txt --proxy http://proxy.firma.de:8080
```

(die tatsächliche Proxy-Adresse bei der IT erfragen, falls unbekannt).

---

## Schritt 6 — Umgebung prüfen

```
python -m gspp.cli doctor
```

Erwartetes Bild:

```
Python            3.12.3   ok
requests          2.33.1   ok
pydantic          2.13.4   ok
openpyxl          3.1.5    ok

Katalog-Download  HTTP 200
GitHub-API        HTTP 403  (Ratelimit - unkritisch, siehe --token)

Bereit.
```

`HTTP 403` bei der GitHub-API ist normal (Ratelimit, kein Fehler) — siehe Anhang unten. Steht bei **Katalog-Download** etwas anderes als `HTTP 200`, ist vermutlich ein Firmen-Proxy oder eine Firewall im Weg; dann Abschnitt „Offline-Modus" weiter unten nutzen.

---

## Schritt 7 — Die Vorlage erzeugen

```
python -m gspp.cli build --ziel-schema 2023 -o out\GSpp_Vorlage.xlsx --snapshot-dir snapshots --cache .cache
```

Nach ein paar Sekunden erscheint:

```
OK: out\GSpp_Vorlage.xlsx  (999 Anforderungen, SHA-256 ...)
```

---

## Schritt 8 — Ergebnis öffnen

Im Explorer in den Unterordner `out` wechseln — dort liegt jetzt `GSpp_Vorlage.xlsx`. Einfach per Doppelklick in Excel öffnen.

Falls Formelzellen zunächst leer aussehen: **Strg + Alt + F9** drücken, um Excel zum Neuberechnen zu zwingen (openpyxl schreibt Formeln ohne zwischengespeichertes Ergebnis — das ist normal und kein Fehler).

---

## Beim nächsten Mal

Die Einrichtung (Schritte 1, 4, 5) ist einmalig. Für jeden weiteren Durchgang reicht:

1. Eingabeaufforderung im `gspp-tool`-Ordner öffnen (Schritt 3)
2. `.venv\Scripts\activate`
3. Der gewünschte Befehl, z. B. erneut Schritt 7, oder für die Änderungsverfolgung:

```
python -m gspp.cli watch --snapshot-dir snapshots
```

---

## Fehlerbehebung

| Meldung | Ursache und Abhilfe |
|---|---|
| `'py' ist nicht als interner oder externer Befehl erkannt` | Python fehlt oder PATH-Häkchen bei der Installation vergessen. Siehe Schritt 1. Nach der Installation ein neues Fenster öffnen. |
| `'python' ist nicht als interner oder externer Befehl erkannt`, aber `py` funktioniert | `py` statt `python` verwenden — beide Befehle sind im Grunde gleichwertig, aber nur `py` ist bei manchen Installationen automatisch verfügbar. |
| PowerShell: „...kann nicht geladen werden, da die Ausführung von Skripts auf diesem System deaktiviert ist" | Siehe Schritt 4, `Set-ExecutionPolicy`-Befehl. Oder einfach die klassische Eingabeaufforderung (`cmd`) statt PowerShell nutzen — dort tritt das Problem nicht auf. |
| `ModuleNotFoundError: No module named 'gspp'` | Zwei mögliche Ursachen: (a) `.venv` ist nicht aktiviert — Zeile beginnt nicht mit `(.venv)`, dann Schritt 4 wiederholen. (b) Die Eingabeaufforderung steht im falschen Ordner — mit `dir` prüfen, ob `gspp` als Unterordner aufgelistet wird. |
| `SSLError` / `ProxyError` beim `pip install` oder beim `build` | Firmenproxy. Siehe Schritt 5 für `pip`. Für den Katalog-Download selbst: Abschnitt „Offline-Modus" unten. |
| Downloads werden von Windows Defender / SmartScreen blockiert oder als „nicht häufig heruntergeladen" markiert | Bei der ZIP-Datei kann ein Rechtsklick → Eigenschaften → Haken bei „Zulassen" nötig sein, **bevor** entpackt wird. |
| Excel zeigt leere Formelzellen | `Strg + Alt + F9` in Excel. Kein Fehler in der Datei. |
| Pfade mit Leerzeichen (z. B. `Eigene Dateien`) verursachen Fehler | In Anführungszeichen setzen, z. B. `cd "C:\Users\deutsj\Eigene Dateien\gspp-tool"`. |

---

## Offline-Modus (kein Zugriff auf GitHub)

Falls Schritt 6 bei **Katalog-Download** dauerhaft einen Fehler zeigt (Firewall blockiert `raw.githubusercontent.com`):

1. Auf einem Rechner mit Internetzugang diese Adresse im Browser öffnen und die Datei speichern:
   ```
   https://raw.githubusercontent.com/BSI-Bund/Stand-der-Technik-Bibliothek/main/Anwenderkataloge/Grundschutz%2B%2B/Grundschutz%2B%2B-catalog.json
   ```
2. Die gespeicherte Datei (z. B. `Grundschutz++-catalog.json`) in den `gspp-tool`-Ordner legen.
3. In Schritt 7 den Befehl anpassen:
   ```
   python -m gspp.cli build --ziel-schema 2023 --katalog Grundschutz++-catalog.json -o out\GSpp_Vorlage.xlsx --snapshot-dir snapshots
   ```

---

## Warum das GitHub-API-„403" in Schritt 6 keine Sorge ist

GitHub erlaubt ohne Anmeldung 60 Abfragen pro Stunde und IP-Adresse. Hinter
dem msg-Netzwerk teilen sich das viele Kolleg:innen — das Limit ist schnell
erreicht. Der eigentliche Katalog-Download läuft über eine andere, unlimitierte
Adresse und ist davon nicht betroffen. Es fehlt lediglich der exakte
Commit-Hash im Herkunftsnachweis der erzeugten Datei; als Ersatz dient die
Prüfsumme (SHA-256) der heruntergeladenen Datei. Wer den Commit-Hash trotzdem
braucht: ein GitHub-Konto mit einem Personal-Access-Token (ohne besondere
Rechte, nur „Public repositories read-only") erhöht das Limit auf 5.000
Anfragen pro Stunde:

```
python -m gspp.cli build --ziel-schema 2023 -o out\GSpp_Vorlage.xlsx --snapshot-dir snapshots --token ghp_xxxxxxxxxxxx
```
