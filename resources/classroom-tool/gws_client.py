"""Tunn subprocess-wrapper kring `gws` (Google Workspace CLI).

Alla funktioner returnerar parsad JSON. Pagination hanteras transparent.
"""

from __future__ import annotations

import json
import subprocess
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent
# Skriv tempfiler till project-dir istället för $TMPDIR. Anledning: under Flatpak
# (Obsidian) sätts $TMPDIR till /run/user/<uid>/app/<id>/ som är privat för
# sandboxen. gws (subprocess) kan inte skriva dit. Project-dir ligger på vanliga
# filsystemet och är synlig från båda håll.
_TMP_DIR = ROOT / ".tmp"


def _run(args: list[str], params: dict | None = None) -> dict:
    cmd = ["gws", *args, "--format", "json"]
    if params is not None:
        cmd += ["--params", json.dumps(params)]
    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(
            f"gws kommando misslyckades (exit {result.returncode}):\n"
            f"  cmd: {' '.join(cmd)}\n"
            f"  stderr: {result.stderr.strip()}\n"
            f"  stdout: {result.stdout.strip()[:500]}"
        )
    out = result.stdout
    start = out.find("{")
    if start < 0:
        raise RuntimeError(f"gws returnerade ingen JSON: {out!r}")
    return json.loads(out[start:])


def _paginate(args: list[str], params: dict, key: str) -> list[dict]:
    items: list[dict] = []
    page_token: str | None = None
    while True:
        p = dict(params)
        if page_token:
            p["pageToken"] = page_token
        resp = _run(args, p)
        items.extend(resp.get(key, []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def list_courses(active_only: bool = True) -> list[dict]:
    params: dict = {"pageSize": 100}
    if active_only:
        params["courseStates"] = "ACTIVE"
    return _paginate(["classroom", "courses", "list"], params, "courses")


def get_course(course_id: str) -> dict:
    return _run(["classroom", "courses", "get"], {"id": course_id})


def list_students(course_id: str) -> list[dict]:
    return _paginate(
        ["classroom", "courses", "students", "list"],
        {"courseId": course_id, "pageSize": 100},
        "students",
    )


def list_teachers(course_id: str) -> list[dict]:
    return _paginate(
        ["classroom", "courses", "teachers", "list"],
        {"courseId": course_id, "pageSize": 100},
        "teachers",
    )


def list_coursework(course_id: str) -> list[dict]:
    return _paginate(
        ["classroom", "courses", "courseWork", "list"],
        {"courseId": course_id, "pageSize": 100, "orderBy": "dueDate desc"},
        "courseWork",
    )


def list_submissions(course_id: str, coursework_id: str) -> list[dict]:
    return _paginate(
        ["classroom", "courses", "courseWork", "studentSubmissions", "list"],
        {"courseId": course_id, "courseWorkId": coursework_id, "pageSize": 100},
        "studentSubmissions",
    )


def list_announcements(course_id: str, limit: int = 5) -> list[dict]:
    resp = _run(
        ["classroom", "courses", "announcements", "list"],
        {"courseId": course_id, "pageSize": limit, "orderBy": "updateTime desc"},
    )
    return resp.get("announcements", [])


def get_coursework(course_id: str, coursework_id: str) -> dict:
    return _run(
        ["classroom", "courses", "courseWork", "get"],
        {"courseId": course_id, "id": coursework_id},
    )


def drive_get_metadata(file_id: str) -> dict:
    """Hämtar fil-metadata (mimeType, name, size, etc.) via `gws drive files get`."""
    return _run(
        ["drive", "files", "get"],
        {"fileId": file_id, "fields": "id,name,mimeType,size"},
    )


def drive_download_binary(file_id: str, output_path: Path) -> None:
    """Laddar ner råinnehållet i en uppladdad Drive-fil (PDF/DOCX/etc).

    Använd för icke-Workspace-filer. Workspace-filer (Docs/Sheets/Slides)
    måste exporteras via drive_export_text — de kan inte laddas ner direkt.

    Drive API:s standardväg för att hämta media är `files.get` med
    `alt=media`. `gws drive files download`-subkommandot returnerar 500
    backendError för normala filer, så vi använder `get` istället.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        "gws",
        "drive",
        "files",
        "get",
        "--params",
        json.dumps({"fileId": file_id, "alt": "media"}),
        "-o",
        str(output_path),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        raise RuntimeError(
            f"gws drive get (alt=media) misslyckades (exit {result.returncode}): "
            f"{result.stderr.strip() or result.stdout.strip()[:300]}"
        )


def drive_export_text(file_id: str) -> str:
    """Exporterar en Google Doc till klartext (UTF-8) via `gws drive files export`.

    Höjer RuntimeError om filen inte är en exporterbar Google Workspace-fil
    (t.ex. uppladdade .docx/.pdf — Drive vägrar export och returnerar 400).
    Felmeddelandet från gws bubblas upp.
    """
    _TMP_DIR.mkdir(exist_ok=True)
    tmp_path = _TMP_DIR / f"export-{uuid.uuid4().hex}.txt"
    try:
        cmd = [
            "gws",
            "drive",
            "files",
            "export",
            "--params",
            json.dumps({"fileId": file_id, "mimeType": "text/plain"}),
            "-o",
            str(tmp_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            raise RuntimeError(
                f"gws drive export misslyckades (exit {result.returncode}): "
                f"{result.stderr.strip() or result.stdout.strip()[:300]}"
            )
        text = tmp_path.read_text(encoding="utf-8", errors="replace")
        # Drive prefixar text-export med BOM. Strippa.
        if text.startswith("﻿"):
            text = text[1:]
        return text
    finally:
        tmp_path.unlink(missing_ok=True)
