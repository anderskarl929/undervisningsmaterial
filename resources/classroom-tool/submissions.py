"""Genererar markdown-output för inlämnings-text.

Två lägen:
  - read(course_id, coursework_id, alias)  — en elevs inlämning, anonymiserad
  - dump(course_id, coursework_id)         — alla inlämningar i klassen, anonymiserade

GDPR: Output sparas under aliaset (`Elev N`). Fri text kan ändå innehålla
identifierande info ("jag heter X"). En banner högst upp gör risken explicit.
"""

from __future__ import annotations

import re
from datetime import datetime

import gws_client
import drive
import forms
from anonymize import CourseAliases

GDPR_BANNER = (
    "> ⚠️ **Inlämningstext kan innehålla identifierande info** "
    "(\"jag heter X\", \"min mamma X\"). Pseudonymerna `Elev N` bryter "
    "kopplingen bakåt men fritext är inte filtrerad. Behandla denna fil "
    "därefter — radera när du är klar."
)


def _normalize_alias(s: str) -> str:
    """'7' / 'elev 7' / 'Elev7' → 'Elev 7'."""
    s = s.strip()
    if not s:
        return s
    digits = "".join(c for c in s if c.isdigit())
    if digits:
        return f"Elev {int(digits)}"
    return s


def _fmt_dt(s: str | None) -> str:
    if not s:
        return "—"
    # Classroom timestamp: "2026-04-30T14:32:11.123Z"
    m = re.match(r"(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})", s)
    return f"{m.group(1)} {m.group(2)}" if m else s


def _safe_label(part: dict, idx: int, total_attachments: int) -> str:
    """Generera ofarlig sektions-rubrik. Classroom auto-namnger kopior med
    elevens fulla namn ('Förnamn Efternamn - Original'), så vi visar aldrig
    den råa filtiteln. Kort-svar / flerval har egna säkra labels redan.
    """
    raw = part.get("title", "")
    if raw in ("(kort svar)", "(flerval)"):
        return raw
    if total_attachments == 1:
        return "Inlämning"
    return f"Bilaga {idx}"


def _format_submission_block(
    alias: str,
    submission: dict,
    body: dict,
) -> list[str]:
    out: list[str] = []
    out.append(f"## {alias}")
    state = submission.get("state", "—")
    submitted = _fmt_dt(submission.get("updateTime"))
    late = " · _sent_" if submission.get("late") else ""
    grade = submission.get("assignedGrade")
    grade_part = f" · Bedömt: {grade}" if grade is not None else ""
    out.append(f"_Status: {state} · Uppdaterad: {submitted}{late}{grade_part}_")
    out.append("")

    parts = body["parts"]
    skipped = body["skipped"]

    if not parts and not skipped:
        out.append("_(ingen text inlämnad)_")
        out.append("")
        return out

    attachment_count = sum(
        1 for p in parts if p.get("title") not in ("(kort svar)", "(flerval)")
    ) + len(skipped)

    attach_idx = 0
    for p in parts:
        is_attachment = p.get("title") not in ("(kort svar)", "(flerval)")
        if is_attachment:
            attach_idx += 1
            label = _safe_label(p, attach_idx, attachment_count)
        else:
            label = p["title"]
        cache_marker = " _(cache)_" if p.get("cached") else ""
        out.append(f"### {label}{cache_marker}")
        text = p["text"].rstrip()
        out.append(text if text else "_(tom fil)_")
        out.append("")

    for _ in skipped:
        attach_idx += 1
    if skipped:
        # Visa endast antal + reason-summering, aldrig råa titlar (kan innehålla namn)
        reasons = ", ".join(sorted({s["reason"] for s in skipped}))
        out.append(f"_{len(skipped)} bilaga(or) ej läst — {reasons}_")
        out.append("")

    return out


def _coursework_header(course: dict, work: dict) -> list[str]:
    out: list[str] = []
    title = work.get("title", "(uppgift utan titel)")
    out.append(f"# {title}")
    out.append("")
    out.append(GDPR_BANNER)
    out.append("")
    out.append(f"_Kurs:_ {course.get('name', '—')} · `{course.get('id', '—')}`  ")
    out.append(f"_Uppgift-ID:_ `{work.get('id', '—')}`  ")
    out.append(f"_Genererad:_ {datetime.now().strftime('%Y-%m-%d %H:%M')}  ")
    if work.get("alternateLink"):
        out.append(f"_Länk:_ {work['alternateLink']}  ")
    out.append("")
    return out


def build_read(
    course_id: str,
    coursework_id: str,
    alias_input: str,
    use_cache: bool = True,
) -> str:
    alias = _normalize_alias(alias_input)

    # Om uppgiften är en Forms-uppgift: delegera till forms.py
    work = gws_client.get_coursework(course_id, coursework_id)
    form_id = forms.extract_form_id(work)
    if form_id:
        return forms.build_form_read(course_id, coursework_id, form_id, alias)

    aliases = CourseAliases(course_id)
    user_id = aliases.user_id_for_alias(alias)
    if not user_id:
        # Försök synka roster ifall aliaset inte byggts än
        for s in gws_client.list_students(course_id):
            aliases.alias_for(s["userId"])
        aliases.save()
        user_id = aliases.user_id_for_alias(alias)
    if not user_id:
        raise SystemExit(
            f"Hittar inte aliaset {alias!r} i kurs {course_id}. "
            f"Kör `summary {course_id}` först eller kontrollera alias-listan."
        )

    course = gws_client.get_course(course_id)
    submissions = gws_client.list_submissions(course_id, coursework_id)
    match = next((s for s in submissions if s.get("userId") == user_id), None)
    if not match:
        raise SystemExit(f"Ingen inlämning hittad för {alias} på uppgift {coursework_id}.")

    body = drive.fetch_submission_text(match, use_cache=use_cache)

    lines = _coursework_header(course, work)
    lines.extend(_format_submission_block(alias, match, body))
    return "\n".join(lines)


def build_dump(
    course_id: str,
    coursework_id: str,
    use_cache: bool = True,
) -> str:
    work = gws_client.get_coursework(course_id, coursework_id)

    # Forms-uppgift? Delegera. Forms har all elevdata via Forms API,
    # inte via studentSubmissions/Drive.
    form_id = forms.extract_form_id(work)
    if form_id:
        return forms.build_form_dump(course_id, coursework_id, form_id)

    course = gws_client.get_course(course_id)
    submissions = gws_client.list_submissions(course_id, coursework_id)

    aliases = CourseAliases(course_id)
    # Säkerställ att alla inlämnande elever har alias innan vi formaterar.
    for s in submissions:
        if s.get("userId"):
            aliases.alias_for(s["userId"])
    aliases.save()

    # Sortera på alias-nummer
    annotated = []
    for s in submissions:
        uid = s.get("userId")
        if not uid:
            continue
        alias = aliases.alias_for(uid)
        annotated.append((int(alias.split()[1]), alias, s))
    annotated.sort(key=lambda t: t[0])

    lines = _coursework_header(course, work)

    submitted = sum(1 for _, _, s in annotated if s.get("state") in ("TURNED_IN", "RETURNED"))
    graded = sum(1 for _, _, s in annotated if s.get("assignedGrade") is not None)
    missing = [a for _, a, s in annotated if s.get("state") in ("CREATED", "NEW")]
    lines.append(
        f"_Status:_ {submitted}/{len(annotated)} inlämnade · "
        f"{graded} bedömda · {len(missing)} saknar inlämning"
    )
    lines.append("")
    lines.append("---")
    lines.append("")

    for _, alias, s in annotated:
        if s.get("state") in ("CREATED", "NEW"):
            # Hoppa över tomma — listas i botten istället
            continue
        body = drive.fetch_submission_text(s, use_cache=use_cache)
        lines.extend(_format_submission_block(alias, s, body))
        lines.append("---")
        lines.append("")

    if missing:
        lines.append(f"## Saknar inlämning ({len(missing)})")
        lines.append(", ".join(missing))
        lines.append("")

    return "\n".join(lines)
