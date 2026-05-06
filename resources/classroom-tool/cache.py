"""Disk-cache för Drive-exporterad text. TTL 24h. Lagras i cache/ (gitignored).

Cache-nyckel = Drive file id. En enkel TTL räcker — Classroom-inlämningar är
sällan föränderliga efter deadline. Vid behov: radera cache/ för full re-fetch.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CACHE_DIR = ROOT / "cache"
TTL_SECONDS = 24 * 3600


def _safe(file_id: str) -> str:
    # Drive id:n är base64-url-säkra men vi normaliserar för säkerhet.
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in file_id)


def _path(file_id: str) -> Path:
    return CACHE_DIR / f"{_safe(file_id)}.txt"


def get(file_id: str) -> str | None:
    p = _path(file_id)
    if not p.exists():
        return None
    age = time.time() - p.stat().st_mtime
    if age > TTL_SECONDS:
        p.unlink(missing_ok=True)
        return None
    return p.read_text(encoding="utf-8")


def put(file_id: str, text: str) -> None:
    CACHE_DIR.mkdir(exist_ok=True)
    p = _path(file_id)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, p)


def purge_expired() -> int:
    if not CACHE_DIR.exists():
        return 0
    now = time.time()
    n = 0
    for p in CACHE_DIR.glob("*.txt"):
        try:
            if now - p.stat().st_mtime > TTL_SECONDS:
                p.unlink(missing_ok=True)
                n += 1
        except FileNotFoundError:
            pass
    return n


def clear_all() -> int:
    if not CACHE_DIR.exists():
        return 0
    n = 0
    for p in CACHE_DIR.glob("*.txt"):
        p.unlink(missing_ok=True)
        n += 1
    return n
