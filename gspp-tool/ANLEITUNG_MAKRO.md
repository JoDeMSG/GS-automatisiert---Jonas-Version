# Makro-Variante: Ansicht vom Deckblatt steuern

## Status

Der erste Versuch (06.08.) scheiterte: Excel meldete beim Öffnen einen
Reparaturbedarf und entfernte das VBA-Projekt ("Entfernter Teil:
/xl/vbaProject.bin").

**Ursachenanalyse gegen eine echte, von Excel erzeugte Datei** (die
BSI-Vorlage A.3.4 enthält ein 322-KB-VBA-Projekt) ergab drei konkrete
Abweichungen, die inzwischen korrigiert sind:

| # | Problem | Korrektur |
|---|---|---|
| 1 | `_VBA_PROJECT` gab vor, gültigen P-Code-Cache zu enthalten | Versionsfeld auf `0xFFFF` — Excel verwirft den Cache und kompiliert aus dem Quelltext neu |
| 2 | Record `0x004A` (PROJECTCOMPATVERSION) fehlte ganz | ergänzt |
| 3 | `PROJECTVERSION` als normaler Record gebaut | Sonderfall umgesetzt: Length-Feld 4, Daten aber 6 Byte |

Die ersten 15 Records des `dir`-Streams stimmen jetzt byte-genau mit der
echten Excel-Datei überein.

**Trotzdem ungeprüft in echtem Excel** — hier steht keine Excel-Installation
zur Verfügung. Deshalb bleibt die Rückfallebene unten bestehen. Falls die
`.xlsm` erneut reklamiert wird: Der manuelle Import funktioniert garantiert,
weil dort Excel selbst den Container baut.

## So geht's

1. `python -m gspp.cli build --ziel-schema 2023 --makro -o out\GSpp_Vorlage.xlsx ...`
   erzeugt neben der `.xlsx` eine Datei `GSppAnsicht.bas` im selben Ordner.
2. `GSpp_Vorlage.xlsx` in Excel öffnen (nicht neu erzeugen).
3. `Alt` + `F11` — der VBA-Editor öffnet sich.
4. Menü **Datei → Datei importieren…**, `GSppAnsicht.bas` auswählen.
5. VBA-Editor schließen.
6. **Datei → Speichern unter…**, Dateityp **Excel-Arbeitsmappe mit
   Makros (\*.xlsm)** wählen.

Fertig. `AnsichtAnwenden` steht ab jetzt über `Alt` + `F8` zur Verfügung und
lässt sich zusätzlich an die beiden Ja/Nein-Zellen auf dem Deckblatt
koppeln (siehe unten).

## Was die Datei ohne Makro kann

Auch ohne den Import bleibt die `.xlsx` voll bedienbar — die Ja/Nein-Zellen
auf dem Deckblatt bewirken dann nichts, aber die Spalten lassen sich über
die `+`/`−` Schaltflächen oberhalb der Spaltenköpfe ein- und ausblenden.
Das ist keine Notlösung, sondern die reguläre Bedienung, wenn ihr auf den
Makro-Import verzichten wollt oder Makros bei msg ohnehin gesperrt sind.

## Verbindung zu den Deckblatt-Zellen herstellen (optional)

Nach dem Import lässt sich das Makro an die Auswahlzellen koppeln, damit
eine Änderung dort automatisch wirkt:

1. Im VBA-Editor (`Alt+F11`) links auf **Deckblatt** doppelklicken
2. Folgenden Code einfügen:
   ```vba
   Private Sub Worksheet_Change(ByVal Target As Range)
       If Not Intersect(Target, Range("GUIDANCE_SICHTBAR")) Is Nothing _
          Or Not Intersect(Target, Range("ZUORDNUNG_SICHTBAR")) Is Nothing Then
           AnsichtAnwenden
       End If
   End Sub
   ```
3. Speichern (als `.xlsm`)

Ohne diesen Schritt lässt sich `AnsichtAnwenden` weiterhin manuell über
`Alt+F8` ausführen — die Zellen dienen dann nur als Merkzettel für den
gewünschten Zustand.

## Warum nicht einfach nachbessern?

Ohne Zugriff auf echtes Excel kann ich Änderungen am Binärformat nicht
verlässlich prüfen — jede weitere Reparatur wäre eine Vermutung, keine
geprüfte Korrektur. Der manuelle Import umgeht das Problem vollständig,
weil an der kritischen Stelle Excel selbst arbeitet statt eines
Nachbaus. Das ist der Grund für den Kurswechsel.
