"""
MS-OVBA Kompression (RLE-Variante), [MS-OVBA] 2.4.1.

Bewusste Vereinfachung: Es werden ausschliesslich komprimierte Chunks erzeugt
und die Eingabe in Bloecke von 3600 Byte zerlegt. Begruendung:

  Ein CompressedChunk darf hoechstens 4096 Byte Nutzdaten tragen. Im
  schlechtesten Fall (nicht komprimierbare Daten) braucht ein Block aus n
  Bytes n + ceil(n/8) Token-Bytes. Bei n = 4096 waeren das 4608 Byte - der
  Groessenzaehler im Chunk-Header laeuft ueber. Reale Implementierungen
  weichen dann auf unkomprimierte RawChunks aus, die immer exakt auf 4096
  Byte aufgefuellt werden.

  Mit n = 3600 gilt 3600 + 450 = 4050 < 4096, der Ueberlauf ist also
  konstruktiv ausgeschlossen und der fehleranfaellige RawChunk-Pfad
  entfaellt vollstaendig. Kosten: minimal schlechtere Kompressionsrate,
  bei VBA-Quelltext (stark redundanter Text) irrelevant.

decompress() ist vollstaendig implementiert, damit jeder erzeugte Container
per Round-Trip geprueft werden kann. Ohne Zugriff auf echtes Excel ist das
die beste verfuegbare Korrektheitspruefung.
"""
from __future__ import annotations

BLOCKGROESSE = 3600


def _bit_count(difference: int) -> int:
    """CopyToken-Bitbreite - [MS-OVBA] 2.4.1.3.19.1."""
    bc = 4
    while (1 << bc) < difference and bc < 12:
        bc += 1
    return bc


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")  # SignatureByte
    pos = 0
    n = len(data)

    while pos < n:
        block = data[pos:pos + BLOCKGROESSE]
        tokens = bytearray()
        i = 0
        while i < len(block):
            flag_pos = len(tokens)
            tokens.append(0)  # Platzhalter fuer FlagByte
            flag = 0
            for bit in range(8):
                if i >= len(block):
                    break
                bc = _bit_count(i) if i > 0 else 4
                length_mask = 0xFFFF >> bc
                max_len = min(length_mask + 3, len(block) - i)

                best_len, best_off = 0, 0
                if i > 0:
                    max_offset = min(i, (0xFFFF >> (16 - bc)) + 1)
                    for cand in range(max(0, i - max_offset), i):
                        ln = 0
                        while ln < max_len and block[cand + ln] == block[i + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len, best_off = ln, i - cand
                            if ln == max_len:
                                break

                if best_len >= 3:
                    token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                    tokens += token.to_bytes(2, "little")
                    flag |= (1 << bit)
                    i += best_len
                else:
                    tokens.append(block[i])
                    i += 1
            tokens[flag_pos] = flag

        assert len(tokens) <= 4096, f"Token-Ueberlauf: {len(tokens)}"
        header = (len(tokens) - 1) | 0xB000
        out += header.to_bytes(2, "little")
        out += tokens
        pos += BLOCKGROESSE

    return bytes(out)


def decompress(data: bytes) -> bytes:
    if not data or data[0] != 0x01:
        raise ValueError("Kein gueltiger CompressedContainer (SignatureByte fehlt)")
    out = bytearray()
    pos = 1
    while pos + 2 <= len(data):
        header = int.from_bytes(data[pos:pos + 2], "little")
        pos += 2
        size = (header & 0x0FFF) + 1
        komprimiert = (header & 0x8000) != 0
        chunk = data[pos:pos + size]
        pos += size

        if not komprimiert:
            out += chunk
            continue

        chunk_start = len(out)
        i = 0
        while i < len(chunk):
            flag = chunk[i]
            i += 1
            for bit in range(8):
                if i >= len(chunk):
                    break
                if flag & (1 << bit):
                    if i + 2 > len(chunk):
                        break
                    token = int.from_bytes(chunk[i:i + 2], "little")
                    i += 2
                    diff = len(out) - chunk_start
                    bc = _bit_count(diff) if diff > 0 else 4
                    length_mask = 0xFFFF >> bc
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bc)) + 1
                    src = len(out) - offset
                    for k in range(length):
                        out.append(out[src + k])
                else:
                    out.append(chunk[i])
                    i += 1
    return bytes(out)
