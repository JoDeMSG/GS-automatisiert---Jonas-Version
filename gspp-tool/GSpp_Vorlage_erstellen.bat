@echo off
REM ============================================================
REM GSpp-Tool - Vorlage per Doppelklick erstellen
REM
REM Erwartet .venv im selben Ordner wie diese Datei (siehe
REM ANLEITUNG_WINDOWS.md, Schritt 4). Erzeugt eine Vorlage mit
REM Zeitstempel im Namen, damit alte Stände nicht ueberschrieben
REM werden.
REM ============================================================

setlocal enabledelayedexpansion

REM %~dp0 = Ordner dieser .bat-Datei, MIT Backslash am Ende.
REM Dadurch funktioniert das Skript unabhaengig davon, von wo aus
REM es gestartet wird - Voraussetzung ist nur, dass es im
REM gspp-tool-Ordner selbst liegt.
set "SCRIPT_DIR=%~dp0"
set "VENV_ACTIVATE=%SCRIPT_DIR%.venv\Scripts\activate.bat"

REM Zeitstempel fuer den Dateinamen (Format JJJJMMTT_HHMM).
REM %date%/%time% sind vom Windows-Regionalformat abhaengig - diese
REM Zusammensetzung ist auf deutschen Systemen (TT.MM.JJJJ) korrekt.
REM Falls der Dateiname seltsam aussieht, ist das die Ursache; die
REM Vorlage selbst ist davon nicht betroffen.
set "ZEITSTEMPEL=%date:~6,4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%"
set "ZEITSTEMPEL=%ZEITSTEMPEL: =0%"

set "AUSGABE=%SCRIPT_DIR%out\GSpp_Vorlage_%ZEITSTEMPEL%.xlsx"

echo ============================================================
echo   GSpp-Tool - Vorlage erstellen
echo ============================================================
echo.

if not exist "%VENV_ACTIVATE%" (
    echo FEHLER: Virtuelle Umgebung nicht gefunden unter:
    echo   %VENV_ACTIVATE%
    echo.
    echo Einmalige Einrichtung noetig ^(siehe ANLEITUNG_WINDOWS.md^):
    echo   py -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

call "%VENV_ACTIVATE%"
if errorlevel 1 (
    echo FEHLER: Virtuelle Umgebung konnte nicht aktiviert werden.
    pause
    exit /b 1
)

echo Erzeuge: %AUSGABE%
echo.

python -m gspp.cli build --ziel-schema 2023 --design bsi ^
    -o "%AUSGABE%" ^
    --snapshot-dir "%SCRIPT_DIR%snapshots" ^
    --cache "%SCRIPT_DIR%.cache"

set "ERGEBNIS=%errorlevel%"

echo.
if "%ERGEBNIS%"=="0" (
    echo ============================================================
    echo   Fertig: %AUSGABE%
    echo ============================================================
) else (
    echo ============================================================
    echo   FEHLER - siehe Meldung oben. Haeufige Ursachen:
    echo     - Kein Netzzugang zu raw.githubusercontent.com
    echo       ^-^> ggf. "pip install pip-system-certs" pruefen
    echo     - .venv veraltet ^-^> pip install -r requirements.txt
    echo ============================================================
)

echo.
pause
endlocal
