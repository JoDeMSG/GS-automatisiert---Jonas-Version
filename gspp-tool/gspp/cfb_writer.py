"""
Minimaler Writer fuer Compound File Binary (CFB / OLE2), [MS-CFB].

Nur so viel Funktionsumfang wie fuer ein vbaProject.bin noetig:
  * 512-Byte-Sektoren, Version 3
  * FAT + MiniFAT (Streams < 4096 Byte liegen im MiniStream)
  * flache Verzeichnisstruktur mit Unterordnern (rot-schwarz-Baum
    vereinfacht als Kette - Excel akzeptiert unbalancierte Baeume)
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field

SEKTOR = 512
MINI_SEKTOR = 64
MINI_GRENZE = 4096

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


@dataclass
class Eintrag:
    name: str
    is_storage: bool = False
    data: bytes = b""
    kinder: list = field(default_factory=list)
    # wird beim Schreiben gefuellt
    idx: int = -1
    child_id: int = NOSTREAM
    left_id: int = NOSTREAM
    right_id: int = NOSTREAM
    start_sect: int = ENDOFCHAIN
    size: int = 0


def _kette(daten: bytes, sektorgroesse: int) -> list[bytes]:
    out = []
    for i in range(0, len(daten), sektorgroesse):
        s = daten[i:i + sektorgroesse]
        out.append(s + b"\x00" * (sektorgroesse - len(s)))
    return out


def schreibe_cfb(root_kinder: list[Eintrag]) -> bytes:
    # --- Verzeichnis linearisieren (Root + rekursiv) ---
    eintraege: list[Eintrag] = []
    root = Eintrag("Root Entry", is_storage=True, kinder=root_kinder)

    def sammle(e: Eintrag):
        e.idx = len(eintraege)
        eintraege.append(e)
        for k in e.kinder:
            sammle(k)

    sammle(root)

    # Kinder als einfache Kette (linkes Kind = erstes, rechts = naechstes Geschwister)
    for e in eintraege:
        if e.kinder:
            e.child_id = e.kinder[0].idx
            for a, b in zip(e.kinder, e.kinder[1:]):
                a.right_id = b.idx

    # --- Stream-Daten aufteilen: gross -> FAT, klein -> MiniFAT ---
    mini_stream = bytearray()
    minifat: list[int] = []
    gross: list[tuple[Eintrag, bytes]] = []

    for e in eintraege:
        if e.is_storage:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start_sect = ENDOFCHAIN
        elif e.size < MINI_GRENZE:
            start = len(mini_stream) // MINI_SEKTOR
            e.start_sect = start
            sektoren = _kette(e.data, MINI_SEKTOR)
            for j, s in enumerate(sektoren):
                mini_stream += s
                minifat.append(start + j + 1 if j < len(sektoren) - 1 else ENDOFCHAIN)
        else:
            gross.append((e, e.data))

    # --- Sektorbelegung planen ---
    fat: list[int] = []
    sektoren: list[bytes] = []

    def belege(daten: bytes) -> int:
        """Legt Daten als FAT-Kette ab, liefert Startsektor."""
        if not daten:
            return ENDOFCHAIN
        teile = _kette(daten, SEKTOR)
        start = len(sektoren)
        for j, t in enumerate(teile):
            sektoren.append(t)
            fat.append(start + j + 1 if j < len(teile) - 1 else ENDOFCHAIN)
        return start

    for e, d in gross:
        e.start_sect = belege(d)

    mini_start = belege(bytes(mini_stream))
    root.start_sect = mini_start
    root.size = len(mini_stream)

    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    minifat_start = belege(minifat_bytes) if minifat_bytes else ENDOFCHAIN
    minifat_count = max(1, (len(minifat_bytes) + SEKTOR - 1) // SEKTOR) if minifat_bytes else 0

    # --- Verzeichnis-Sektoren ---
    dir_bytes = bytearray()
    for e in eintraege:
        name_utf16 = e.name.encode("utf-16-le") + b"\x00\x00"
        eintrag = bytearray(128)
        eintrag[0:len(name_utf16)] = name_utf16
        struct.pack_into("<H", eintrag, 64, len(name_utf16))
        eintrag[66] = 5 if e.idx == 0 else (1 if e.is_storage else 2)  # Root/Storage/Stream
        eintrag[67] = 1  # schwarz
        struct.pack_into("<I", eintrag, 68, e.left_id)
        struct.pack_into("<I", eintrag, 72, e.right_id)
        struct.pack_into("<I", eintrag, 76, e.child_id)
        struct.pack_into("<I", eintrag, 116, e.start_sect if e.start_sect != ENDOFCHAIN else ENDOFCHAIN)
        struct.pack_into("<Q", eintrag, 120, e.size)
        dir_bytes += eintrag
    while len(dir_bytes) % SEKTOR:
        dir_bytes += b"\x00" * 128
    dir_start = belege(bytes(dir_bytes))
    dir_count = len(dir_bytes) // SEKTOR

    # --- FAT-Sektoren selbst einplanen (iterativ, da sie sich selbst enthalten) ---
    eintraege_pro_fat = SEKTOR // 4
    while True:
        benoetigt = (len(fat) + eintraege_pro_fat - 1) // eintraege_pro_fat
        if len(fat) + benoetigt <= benoetigt * eintraege_pro_fat:
            break
        benoetigt += 1
        break
    fat_anzahl = max(1, (len(fat) + eintraege_pro_fat - 1) // eintraege_pro_fat)
    while len(fat) + fat_anzahl > fat_anzahl * eintraege_pro_fat:
        fat_anzahl += 1

    fat_sektor_ids = list(range(len(sektoren), len(sektoren) + fat_anzahl))
    for _ in range(fat_anzahl):
        fat.append(FATSECT)
        sektoren.append(b"\x00" * SEKTOR)  # Platzhalter, gleich ersetzt

    while len(fat) < fat_anzahl * eintraege_pro_fat:
        fat.append(FREESECT)

    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    for j, sid in enumerate(fat_sektor_ids):
        sektoren[sid] = fat_bytes[j * SEKTOR:(j + 1) * SEKTOR]

    # --- Header ---
    header = bytearray(SEKTOR)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", header, 24, 0x003E)   # Minor Version
    struct.pack_into("<H", header, 26, 0x0003)   # Major Version 3
    struct.pack_into("<H", header, 28, 0xFFFE)   # Little Endian
    struct.pack_into("<H", header, 30, 9)        # Sektorgroesse 2^9 = 512
    struct.pack_into("<H", header, 32, 6)        # Mini-Sektorgroesse 2^6 = 64
    struct.pack_into("<I", header, 44, fat_anzahl)
    struct.pack_into("<I", header, 48, dir_start)
    struct.pack_into("<I", header, 56, MINI_GRENZE)
    struct.pack_into("<I", header, 60, minifat_start)
    struct.pack_into("<I", header, 64, minifat_count)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)  # DIFAT-Start
    struct.pack_into("<I", header, 72, 0)           # DIFAT-Anzahl
    for j in range(109):
        wert = fat_sektor_ids[j] if j < len(fat_sektor_ids) else FREESECT
        struct.pack_into("<I", header, 76 + j * 4, wert)

    return bytes(header) + b"".join(sektoren)
