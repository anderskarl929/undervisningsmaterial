"""Hämta inlämnings-text via Drive. Tunn orkestrering ovanpå gws_client + cache.

En Classroom-inlämning kan ha:
  - assignmentSubmission.attachments[]   (Drive-fil, länk, YouTube, Form)
  - shortAnswerSubmission.answer         (kort textsvar)
  - multipleChoiceSubmission.answer      (flerval)

Strategi för Drive-bilagor:
  - Google Workspace-filer (Doc/Sheet/Slide) → exporteras till klartext
  - Uppladdade filer (PDF/DOCX/ODT/text) → laddas ner och konverteras lokalt
  - Övrigt (legacy .doc, RTF, bild, video, …) → noteras som skipped
"""

from __future__ import annotations

import uuid
from pathlib import Path

import gws_client
import cache
import converters

ROOT = Path(__file__).resolve().parent
_TMP_DIR = ROOT / ".tmp"

# MIME-prefix för Google Workspace-filer som kräver export
_WORKSPACE_MIME_PREFIX = "application/vnd.google-apps."


def _fetch_text_for_file(file_id: str, fallback_title: str) -> tuple[str, str]:
    """Returnerar (title, text). Höjer RuntimeError om allt misslyckas."""
    meta = gws_client.drive_get_metadata(file_id)
    mime = (meta.get("mimeType") or "").lower()
    title = meta.get("name") or fallback_title

    if mime.startswith(_WORKSPACE_MIME_PREFIX):
        # Google Workspace — Docs/Sheets/Slides måste exporteras
        text = gws_client.drive_export_text(file_id)
        return title, text

    # Uppladdad fil — ladda ner binärt och konvertera lokalt
    suffix = Path(title).suffix or ".bin"
    tmp_path = _TMP_DIR / f"dl-{uuid.uuid4().hex}{suffix}"
    try:
        gws_client.drive_download_binary(file_id, tmp_path)
        text = converters.extract_text(tmp_path, mime)
        return title, text
    finally:
        tmp_path.unlink(missing_ok=True)


def fetch_submission_text(submission: dict, use_cache: bool = True) -> dict:
    """Returnerar struktur:
    {
      'parts':   [{'title': str, 'text': str, 'cached': bool}],
      'skipped': [{'title': str, 'reason': str}],
    }
    """
    parts: list[dict] = []
    skipped: list[dict] = []

    short = submission.get("shortAnswerSubmission") or {}
    if short.get("answer"):
        parts.append({"title": "(kort svar)", "text": short["answer"], "cached": False})

    mcq = submission.get("multipleChoiceSubmission") or {}
    if mcq.get("answer"):
        parts.append({"title": "(flerval)", "text": mcq["answer"], "cached": False})

    asg = submission.get("assignmentSubmission") or {}
    for att in asg.get("attachments") or []:
        df = att.get("driveFile")
        if not df:
            kind = next(iter(k for k in att.keys() if k != "driveFile"), "okänd")
            skipped.append({"title": f"({kind})", "reason": "ej-Drive-fil"})
            continue
        file_id = df.get("id")
        fallback_title = df.get("title", file_id or "(utan titel)")
        if not file_id:
            skipped.append({"title": fallback_title, "reason": "saknar Drive-id"})
            continue

        text = cache.get(file_id) if use_cache else None
        was_cached = text is not None
        title = fallback_title
        if text is None:
            try:
                title, text = _fetch_text_for_file(file_id, fallback_title)
            except converters.UnsupportedFormatError as e:
                skipped.append({"title": fallback_title, "reason": str(e)[:200]})
                continue
            except RuntimeError as e:
                msg = str(e)
                if "Export only supports" in msg or "not supported" in msg.lower():
                    skipped.append({"title": fallback_title, "reason": "ej Google Doc"})
                else:
                    skipped.append({"title": fallback_title, "reason": msg[:200]})
                continue
            if use_cache:
                cache.put(file_id, text)
        parts.append({"title": title, "text": text, "cached": was_cached})

    return {"parts": parts, "skipped": skipped}
