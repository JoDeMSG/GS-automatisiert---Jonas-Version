"""
Fetch-Layer: holt den OSCAL-Katalog aus dem BSI-GitHub - oder liest ihn lokal.

Designentscheidungen:
  * Offline-Modus ist gleichberechtigt (--katalog datei.json), damit das Tool
    auch in abgeschotteten Umgebungen laeuft.
  * Der Commit-SHA wird ueber die GitHub-Contents-API mitgezogen. Er ist der
    einzige belastbare Versionsanker - das BSI liefert keine Release-Tags.
  * Jeder Download wird per SHA-256 fixiert und im Cache abgelegt.
"""
from __future__ import annotations

import hashlib
import json
import logging
from datetime import datetime, timezone
from pathlib import Path

import requests

log = logging.getLogger(__name__)

REPO = "BSI-Bund/Stand-der-Technik-Bibliothek"
BRANCH = "main"
# Stand 2026-07-31: Das BSI hat das Repository umstrukturiert.
# Alter Pfad (bis mind. 2026-07-16, seither HTTP 404):
#   Anwenderkataloge/Grundschutz++/Grundschutz++-catalog.json
# Neuer Pfad:
KATALOG_PFAD = "control_layer/Grundschutz++/Grundschutz++-resolved_catalog.json"

RAW_URL = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/{KATALOG_PFAD}"
API_COMMITS = f"https://api.github.com/repos/{REPO}/commits"

TIMEOUT = 60


class FetchError(RuntimeError):
    pass


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def letzter_commit(pfad: str = KATALOG_PFAD, token: str | None = None) -> tuple[str | None, str | None]:
    """Liefert (sha, iso-datum) des letzten Commits, der die Datei beruehrt hat."""
    headers = {"Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    try:
        r = requests.get(
            API_COMMITS,
            params={"path": pfad, "sha": BRANCH, "per_page": 1},
            headers=headers,
            timeout=TIMEOUT,
        )
        if r.status_code == 403:
            rest = r.headers.get("X-RateLimit-Remaining")
            log.warning(
                "GitHub-API-Ratelimit erreicht (verbleibend: %s). Der Katalog wird "
                "trotzdem geladen; als Versionsanker dient dann die SHA-256-Summe "
                "der Quelldatei. Fuer den Commit-SHA: --token <PAT> setzen "
                "(hebt das Limit von 60 auf 5000 Anfragen/Stunde).", rest or "0")
            return None, None
        r.raise_for_status()
        commits = r.json()
        if not commits:
            log.warning("Kein Commit zu %s gefunden.", pfad)
            return None, None
        c = commits[0]
        return c["sha"], c["commit"]["committer"]["date"]
    except Exception as exc:  # Netzfehler darf den Lauf nicht killen
        log.warning("Commit-Metadaten nicht abrufbar (%s). SHA-256 der Quelldatei "
                    "bleibt als Versionsanker.", exc)
        return None, None


def hole_katalog(
    *,
    lokale_datei: Path | None = None,
    cache_dir: Path | None = None,
    token: str | None = None,
) -> tuple[dict, dict]:
    """
    Gibt (katalog_dict, herkunft_dict) zurueck.

    herkunft_dict: quelle_url, commit_sha, commit_datum, sha256_quelle, abgerufen_am
    """
    if lokale_datei is not None:
        raw = Path(lokale_datei).read_bytes()
        log.info("Katalog lokal gelesen: %s (%.1f MB)", lokale_datei, len(raw) / 1e6)
        herkunft = {
            "quelle_url": f"file://{Path(lokale_datei).resolve()}",
            "commit_sha": None,
            "commit_datum": None,
            "sha256_quelle": sha256_bytes(raw),
            "abgerufen_am": datetime.now(timezone.utc).isoformat(),
        }
        return json.loads(raw), herkunft

    sha, datum = letzter_commit(token=token)
    log.info("Lade Katalog von GitHub (commit %s)", (sha or "unbekannt")[:8])
    r = requests.get(RAW_URL, timeout=TIMEOUT)
    if r.status_code != 200:
        raise FetchError(f"Download fehlgeschlagen: HTTP {r.status_code} - {RAW_URL}")
    raw = r.content
    digest = sha256_bytes(raw)

    if cache_dir:
        cache_dir = Path(cache_dir)
        cache_dir.mkdir(parents=True, exist_ok=True)
        ziel = cache_dir / f"catalog_{(sha or digest)[:12]}.json"
        if not ziel.exists():
            ziel.write_bytes(raw)
            log.info("Cache geschrieben: %s", ziel)

    herkunft = {
        "quelle_url": RAW_URL,
        "commit_sha": sha,
        "commit_datum": datum,
        "sha256_quelle": digest,
        "abgerufen_am": datetime.now(timezone.utc).isoformat(),
    }
    return json.loads(raw), herkunft
