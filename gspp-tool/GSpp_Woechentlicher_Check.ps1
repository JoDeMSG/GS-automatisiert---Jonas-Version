# ============================================================
# GSpp-Tool - Woechentlicher Katalog-Check (fuer Windows-Aufgabenplanung)
#
# Prueft, ob sich der BSI-Katalog seit dem letzten Lauf geaendert hat.
# Bei Aenderung: Aenderungsbericht (.xlsx + .md) im snapshots-Ordner.
# Ohne Aenderung: Log-Eintrag, sonst passiert nichts.
#
# WICHTIG - so registrieren, dass die Ausfuehrungsrichtlinie nicht
# im Weg steht (siehe Abschnitt "Einrichtung" unten in dieser Datei):
#   Aktion in der Aufgabenplanung:
#     Programm/Skript:  powershell.exe
#     Argumente:        -ExecutionPolicy Bypass -File "C:\Pfad\zu\GSpp_Woechentlicher_Check.ps1"
# ============================================================

param(
    # Wenn gesetzt, meldet das Skript bei gefundenen Aenderungen einen
    # Fehlercode (2) an die Aufgabenplanung zurueck - dort erscheint der
    # Lauf dann als "fehlgeschlagen", was manche als hilfreiches visuelles
    # Signal nutzen wollen ("etwas hat sich geaendert, bitte pruefen").
    # Standardmaessig AUS, damit ein normaler, erfolgreicher Lauf mit
    # gefundenen Aenderungen nicht wie ein echter Fehler aussieht.
    [switch]$FailOnChange
)

$ErrorActionPreference = "Stop"

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python     = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$Snapshots  = Join-Path $ScriptDir "snapshots"
$CacheDir   = Join-Path $ScriptDir ".cache"
$LogDatei   = Join-Path $ScriptDir "logs\woechentlicher_check.log"

New-Item -ItemType Directory -Force -Path (Split-Path $LogDatei) | Out-Null

function Schreibe-Log {
    param([string]$Text)
    $Zeitstempel = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Zeitstempel  $Text" | Add-Content -Path $LogDatei -Encoding UTF8
}

Schreibe-Log "=== Lauf gestartet ==="

# Direkt die python.exe der virtuellen Umgebung aufrufen statt sie ueber
# Activate.ps1 zu "aktivieren". Vermeidet, dass ein per Gruppenrichtlinie
# eingeschraenktes ExecutionPolicy-Setting den ganzen Lauf blockiert -
# eine .exe direkt aufzurufen ist davon nicht betroffen, nur das
# Ausfuehren von .ps1-Skripten selbst.
if (-not (Test-Path $Python)) {
    Schreibe-Log "FEHLER: Virtuelle Umgebung nicht gefunden unter $Python"
    Schreibe-Log "Einmalige Einrichtung noetig (siehe ANLEITUNG_WINDOWS.md)."
    exit 1
}

try {
    $Ausgabe = & $Python -m gspp.cli watch `
        --snapshot-dir $Snapshots `
        --cache $CacheDir 2>&1 | Out-String

    $ExitCode = $LASTEXITCODE
    Schreibe-Log $Ausgabe.Trim()

    switch ($ExitCode) {
        0 {
            Schreibe-Log "Ergebnis: keine Aenderung am Katalog."
        }
        2 {
            Schreibe-Log "Ergebnis: AENDERUNG ERKANNT. Bericht liegt in $Snapshots"
            if ($FailOnChange) {
                Schreibe-Log "=== Lauf beendet (Exitcode 2, Aenderung) ==="
                exit 2
            }
        }
        default {
            Schreibe-Log "Ergebnis: unerwarteter Exitcode $ExitCode - bitte Log pruefen."
            Schreibe-Log "=== Lauf beendet (Exitcode $ExitCode, Fehler) ==="
            exit 1
        }
    }
}
catch {
    Schreibe-Log "FEHLER: $($_.Exception.Message)"
    Schreibe-Log "=== Lauf beendet (Ausnahme) ==="
    exit 1
}

Schreibe-Log "=== Lauf beendet (erfolgreich) ==="
exit 0

# ============================================================
# Einrichtung als woechentliche Aufgabe (einmalig, ueber die GUI):
#
# 1. Windows-Suche: "Aufgabenplanung" oeffnen
# 2. Rechts: "Einfache Aufgabe erstellen..."
# 3. Name: z. B. "GSpp Woechentlicher Check"
# 4. Trigger: Woechentlich, Wochentag/Uhrzeit nach Wahl
# 5. Aktion: "Programm starten"
#      Programm/Skript:  powershell.exe
#      Argumente:         -ExecutionPolicy Bypass -File "VOLLER_PFAD\GSpp_Woechentlicher_Check.ps1"
#      Starten in:        VOLLER_PFAD_ZUM_gspp-tool-ORDNER
# 6. Fertigstellen
#
# Alternativ per Kommandozeile (einmalig, als Admin):
#   schtasks /create /tn "GSpp Woechentlicher Check" /tr "powershell.exe -ExecutionPolicy Bypass -File \"VOLLER_PFAD\GSpp_Woechentlicher_Check.ps1\"" /sc weekly /d MON /st 07:00
#
# Log-Datei liegt danach unter logs\woechentlicher_check.log - dort laesst
# sich nachvollziehen, ob und wann der Task tatsaechlich gelaufen ist.
# ============================================================
