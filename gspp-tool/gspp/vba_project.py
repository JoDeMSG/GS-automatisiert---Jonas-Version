"""
Erzeugt ein vbaProject.bin mit einem einzigen Standardmodul.

Aufbau nach [MS-OVBA]:
    /PROJECT            Textkonfiguration (unkomprimiert)
    /PROJECTwm          Modulnamen-Zuordnung ANSI/Unicode
    /VBA/_VBA_PROJECT   Versionskennung
    /VBA/dir            komprimierte Projektstruktur (Records)
    /VBA/<Modul>        komprimierter Quelltext
"""
from __future__ import annotations

import struct

from .cfb_writer import Eintrag, schreibe_cfb
from .ovba_compress import compress

CODEPAGE = 1252  # Windows-1252


def _rec(id_: int, daten: bytes) -> bytes:
    return struct.pack("<HI", id_, len(daten)) + daten


def _dir_stream(projektname: str, modulname: str, modul_offset: int) -> bytes:
    n = projektname.encode("cp1252")
    m = modulname.encode("cp1252")
    d = bytearray()
    # PROJECTINFORMATION
    d += _rec(0x0001, struct.pack("<I", 0x00000003))          # SysKind (32-bit)
    # PROJECTCOMPATVERSION - in echten Excel-Dateien direkt nach SysKind.
    # Fehlte im ersten Versuch; aus der BSI-Referenzdatei uebernommen.
    d += _rec(0x004A, struct.pack("<I", 0x00000006))
    d += _rec(0x0002, struct.pack("<I", 0x00000409))          # Lcid
    d += _rec(0x0014, struct.pack("<I", 0x00000409))          # LcidInvoke
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))            # CodePage
    d += _rec(0x0004, n)                                      # Name
    d += _rec(0x0005, b"")                                    # DocString
    d += struct.pack("<HI", 0x0040, 0)                        # DocStringUnicode
    d += _rec(0x0006, b"")                                    # HelpFile1
    d += _rec(0x003D, b"")                                    # HelpFile2
    d += _rec(0x0007, struct.pack("<I", 0))                   # HelpContext
    d += _rec(0x0008, struct.pack("<I", 0))                   # LibFlags
    # PROJECTVERSION - Sonderfall laut [MS-OVBA] 2.3.4.2.1.9:
    # Das Length-Feld traegt IMMER 4, die Daten sind aber 6 Byte
    # (VersionMajor 4 Byte + VersionMinor 2 Byte). Im ersten Versuch war
    # hier faelschlich ein normaler Record mit 10 Byte Daten.
    d += struct.pack("<HI", 0x0009, 4) + struct.pack("<IH", 0x0001322B, 0x0001)
    d += _rec(0x000C, b"")                                    # Constants
    d += struct.pack("<HI", 0x003C, 0)                        # ConstantsUnicode

    # PROJECTMODULES
    d += _rec(0x000F, struct.pack("<H", 1))                   # Count = 1 Modul
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))              # ProjectCookie

    # MODULE
    d += _rec(0x0019, m)                                      # ModuleName
    d += struct.pack("<HI", 0x0047, len(m) * 2) + modulname.encode("utf-16-le")
    d += _rec(0x001A, m)                                      # StreamName
    d += struct.pack("<HI", 0x0032, len(m) * 2) + modulname.encode("utf-16-le")
    d += _rec(0x001C, b"")                                    # DocString
    d += struct.pack("<HI", 0x0048, 0)
    d += _rec(0x0031, struct.pack("<I", modul_offset))        # TextOffset
    d += _rec(0x001E, struct.pack("<I", 0))                   # HelpContext
    d += _rec(0x002C, struct.pack("<H", 0xFFFF))              # Cookie
    d += _rec(0x0021, b"")                                    # Type: Standardmodul
    d += struct.pack("<HI", 0x002B, 0)                        # Terminator Modul
    d += struct.pack("<HI", 0x0010, 0)                        # Terminator dir
    return bytes(d)


def _project_stream(projektname: str, modulname: str) -> bytes:
    # Echte Excel-Dateien tragen hier eine gueltige GUID (Beispiel aus der
    # BSI-Vorlage: {CC61483E-...}). Nullen wie im ersten Versuch weichen von
    # allem ab, was Excel selbst schreibt - deshalb eine feste, gueltige GUID.
    zeilen = [
        'ID="{9A3B7C21-4E5D-4F8A-B6C2-1D7E9F3A5B84}"',
        f"Module={modulname}",
        f'Name="{projektname}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="0000000000000000000000000000"',
        'DPB="0000000000000000000000000000"',
        'GC="0000000000000000000000000000"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
    ]
    return "\r\n".join(zeilen).encode("cp1252")


def _projectwm_stream(modulname: str) -> bytes:
    return (modulname.encode("cp1252") + b"\x00"
            + modulname.encode("utf-16-le") + b"\x00\x00\x00\x00")


def _vba_project_stream() -> bytes:
    """
    _VBA_PROJECT-Stream - [MS-OVBA] 2.3.4.1.

    Dieser Stream ist ein Performance-Cache mit vorkompiliertem P-Code.
    Wir koennen keinen gueltigen P-Code erzeugen (das macht nur der
    VBA-Compiler selbst) - und duerfen es auch nicht vortaeuschen:

      Der erste Versuch scheiterte genau daran. Dort stand im Versionsfeld
      ein Wert, den Excel als "Cache passt zu meiner Version" gelesen hat.
      Excel hat daraufhin versucht, den nicht vorhandenen P-Code zu laden,
      ist gescheitert und hat den kompletten Teil verworfen
      ("Entfernter Teil: /xl/vbaProject.bin").

    Loesung: Das Versionsfeld wird bewusst auf einen Wert gesetzt, der zu
    keiner realen Excel-Version passt. Excel erkennt den Cache dann als
    veraltet, verwirft ihn und kompiliert aus dem Klartext-Quelltext neu -
    der liegt korrekt im Modul-Stream und im dir-Stream.

    Zum Vergleich: Eine echte, von Excel erzeugte Datei (BSI-Vorlage A.3.4)
    hat hier Version 181 (0x00B5) und 28.536 Byte P-Code. Wir setzen 0xFFFF.

    Diese Technik ist in der Sicherheitsforschung als "VBA Stomping" bekannt
    und wird dort genutzt, um Quelltext und P-Code auseinanderzuhalten.
    """
    return (
        b"\xcc\x61"                  # Reserved1 - Kennzeichen
        + b"\xff\xff"                # Version - absichtlich ungueltig
        + b"\x00"                    # Reserved2
        + b"\x03\x00"                # Reserved3
        + b"\xff"                    # Reserved4
        + b"\x01\x00\x00\x00"        # Reserved5
        + b"\x00" * 4                # Reserved6
        + b"\x00" * 2                # Reserved7
    )


def baue_vbaproject(vba_code: str, projektname: str = "VBAProject",
                    modulname: str = "GSppAnsicht") -> bytes:
    """Liefert den kompletten Inhalt einer vbaProject.bin."""
    # Modulstream: Attribut-Kopf + Quelltext, dann komprimiert
    kopf = f'Attribute VB_Name = "{modulname}"\r\n'
    voller_text = kopf + vba_code.replace("\n", "\r\n")
    modul_komprimiert = compress(voller_text.encode("cp1252"))

    dir_roh = _dir_stream(projektname, modulname, 0)
    dir_komprimiert = compress(dir_roh)

    vba_storage = Eintrag("VBA", is_storage=True, kinder=[
        Eintrag("_VBA_PROJECT", data=_vba_project_stream()),
        Eintrag("dir", data=dir_komprimiert),
        Eintrag(modulname, data=modul_komprimiert),
    ])
    return schreibe_cfb([
        vba_storage,
        Eintrag("PROJECT", data=_project_stream(projektname, modulname)),
        Eintrag("PROJECTwm", data=_projectwm_stream(modulname)),
    ])
