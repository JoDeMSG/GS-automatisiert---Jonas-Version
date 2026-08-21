"""
Makro-Variante: Spaltensteuerung vom Deckblatt aus.

Wichtig - mehrstufige Rueckfallebene, damit die Datei NIE unbrauchbar wird:

  Stufe 1  Makros aktiviert -> Auswahl auf dem Deckblatt (Zellen mit
           Ja/Nein-Dropdown) steuert per Worksheet_Change alle Praktikblaetter.
  Stufe 2  Makros blockiert -> die Excel-Spaltengruppierung (+/- ueber den
           Spaltenkoepfen) bleibt vollstaendig funktionsfaehig. Identisch zur
           makrofreien .xlsx-Variante, es geht also nichts verloren.
  Stufe 3  VBA-Projekt beschaedigt/abgelehnt -> die mitgelieferte .bas-Datei
           laesst sich in 30 Sekunden von Hand importieren (Anleitung im
           Deckblatt und in der ANLEITUNG_MAKRO.md).

Der erzeugte vbaProject.bin-Container wird von diesem Werkzeug selbst gebaut
(MS-CFB + MS-OVBA). Er ist gegen olefile und per Quelltext-Round-Trip
geprueft, konnte hier aber NICHT gegen echtes Excel getestet werden.
Deshalb die Stufen 2 und 3.
"""
from __future__ import annotations

import logging
import shutil
import zipfile
from pathlib import Path

log = logging.getLogger(__name__)

# Steuerzellen auf dem Deckblatt
STEUER_ZEILE_GUIDANCE = 0  # wird beim Schreiben gesetzt
STEUER_SPALTE = 3          # Spalte C

def _spaltenbereiche() -> tuple[str, str, str]:
    """
    Leitet die VBA-Spaltenbereiche aus dem tatsaechlichen Layout ab.

    Verhindert, dass das Makro nach einer Layout-Aenderung die falschen
    Spalten aus- und einblendet - genau das waere beim Ergaenzen der
    Schutzziel-Spalten sonst passiert.
    """
    from openpyxl.utils import get_column_letter as _g
    from .schema2023 import GRUPPE_GUIDANCE, GRUPPE_ZUORDNUNG, SPALTEN_2023

    zuo = f"{_g(min(GRUPPE_ZUORDNUNG))}:{_g(max(GRUPPE_ZUORDNUNG))}"
    gui = f"{_g(min(GRUPPE_GUIDANCE))}:{_g(max(GRUPPE_GUIDANCE))}"
    alle = f"A:{_g(len(SPALTEN_2023))}"
    return zuo, gui, alle


VBA_CODE_VORLAGE = '''Option Explicit

' Steuert die Sichtbarkeit der Zusatzspalten auf allen Praktikblaettern.
' Aufgerufen automatisch bei Aenderung der Steuerzellen auf dem Deckblatt.
' Faellt das Makro aus (Makros deaktiviert), bleibt die Spaltengruppierung
' als vollwertige Bedienalternative erhalten.

Public Sub AnsichtAnwenden()
    Dim ws As Worksheet
    Dim deck As Worksheet
    Dim zeigeGuidance As Boolean
    Dim zeigeZuordnung As Boolean

    On Error GoTo Fehler
    Set deck = ThisWorkbook.Worksheets("Deckblatt")

    zeigeGuidance = (UCase$(Trim$(CStr(deck.Range("GUIDANCE_SICHTBAR").Value))) = "JA")
    zeigeZuordnung = (UCase$(Trim$(CStr(deck.Range("ZUORDNUNG_SICHTBAR").Value))) = "JA")

    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If IstPraktikblatt(ws) Then
            ws.Range("{ZUORDNUNG}").EntireColumn.Hidden = Not zeigeZuordnung
            ws.Range("{GUIDANCE}").EntireColumn.Hidden = Not zeigeGuidance
        End If
    Next ws
    Application.ScreenUpdating = True
    Exit Sub

Fehler:
    Application.ScreenUpdating = True
    MsgBox "Ansicht konnte nicht angewendet werden: " & Err.Description & vbCrLf & vbCrLf & _
           "Die Spalten lassen sich weiterhin ueber die +/- Schaltflaechen " & _
           "oberhalb der Spaltenkoepfe ein- und ausblenden.", vbInformation
End Sub

' Ein Praktikblatt erkennt man daran, dass in Zeile 6 Spalte A "Anforderung" steht.
Private Function IstPraktikblatt(ws As Worksheet) As Boolean
    On Error Resume Next
    IstPraktikblatt = (Trim$(CStr(ws.Cells(6, 1).Value)) = "Anforderung")
End Function

Public Sub AllesEinblenden()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If IstPraktikblatt(ws) Then ws.Columns("{ALLE}").EntireColumn.Hidden = False
    Next ws
    Application.ScreenUpdating = True
End Sub
'''

# Wird als Klassenmodul-Code auf dem Deckblatt gebraucht; da wir nur ein
# Standardmodul einbetten koennen, wird stattdessen ein Auto-Open genutzt.
VBA_CODE_AUTOOPEN = '''
Public Sub Auto_Open()
    AnsichtAnwenden
End Sub
'''


def _vba_code() -> str:
    """VBA-Quelltext mit eingesetzten, aus dem Layout abgeleiteten Bereichen."""
    zuo, gui, alle = _spaltenbereiche()
    return (VBA_CODE_VORLAGE
            .replace("{ZUORDNUNG}", zuo)
            .replace("{GUIDANCE}", gui)
            .replace("{ALLE}", alle))


def erzeuge_bas_datei(ziel: Path) -> Path:
    """
    Schreibt den Makrocode als importierbare .bas-Datei.

    WICHTIG - binaer schreiben, nicht als Text:
    write_text() wandelt unter Windows jedes '\\n' zusaetzlich in '\\r\\n' um.
    Da der Quelltext bereits '\\r\\n' enthaelt, entstuende '\\r\\r\\n'. VBA
    meldet darauf "Fehler beim Kompilieren: Syntaxfehler", vor allem an
    Zeilenfortsetzungen mit '_'. Unter Linux faellt das nicht auf, weil dort
    keine Umwandlung stattfindet - der Fehler tritt also nur beim Erzeugen
    auf einem Windows-Rechner auf.
    """
    inhalt = 'Attribute VB_Name = "GSppAnsicht"\r\n' + \
             (_vba_code() + VBA_CODE_AUTOOPEN).replace("\r\n", "\n").replace("\n", "\r\n")
    ziel.parent.mkdir(parents=True, exist_ok=True)
    ziel.write_bytes(inhalt.encode("cp1252"))
    return ziel


def zu_xlsm(xlsx_pfad: Path, ziel: Path, vbaproject: bytes) -> Path:
    """
    Wandelt eine erzeugte .xlsx in eine .xlsm mit eingebettetem VBA-Projekt.

    Arbeitet direkt auf der ZIP-Struktur: Content-Types anpassen, Beziehung
    ergaenzen, vbaProject.bin einlegen.
    """
    ziel.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(xlsx_pfad, "r") as zin:
        eintraege = {n: zin.read(n) for n in zin.namelist()}

    # 1) Content-Types: xlsx -> xlsm, vbaProject registrieren
    ct = eintraege["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    )
    if "vbaProject" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/vbaProject.bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
        )
    eintraege["[Content_Types].xml"] = ct.encode("utf-8")

    # 2) Beziehung von workbook.xml auf vbaProject.bin
    rels_pfad = "xl/_rels/workbook.xml.rels"
    rels = eintraege[rels_pfad].decode("utf-8")
    if "vbaProject" not in rels:
        import re
        vorhandene = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
        neue_id = f"rId{max(vorhandene) + 1 if vorhandene else 1}"
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="{neue_id}" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
            'Target="vbaProject.bin"/></Relationships>',
        )
        eintraege[rels_pfad] = rels.encode("utf-8")

    # 3) VBA-Projekt einlegen
    eintraege["xl/vbaProject.bin"] = vbaproject

    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, daten in eintraege.items():
            zout.writestr(name, daten)

    log.info("Makro-Variante geschrieben: %s", ziel)
    return ziel
