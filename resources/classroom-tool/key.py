"""Genererar en utskrivbar HTML-nyckel som mappar Elev N -> riktiga namn.

Filen skrivs till keys/<course>-<timestamp>.html (gitignored). Du öppnar den
i browser, skriver ut fysiskt, och raderar filen manuellt efteråt.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import gws_client
from anonymize import CourseAliases

ROOT = Path(__file__).resolve().parent
KEYS_DIR = ROOT / "keys"


def _student_name(student: dict) -> str:
    profile = student.get("profile", {}) or {}
    name = profile.get("name", {}) or {}
    return name.get("fullName") or profile.get("emailAddress") or student.get("userId", "—")


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="utf-8">
<title>Nyckel — {course_name}</title>
<style>
  @page {{ size: A4; margin: 1.5cm; }}
  body {{ font-family: Georgia, "Times New Roman", serif; color: #2a2a2a; max-width: 18cm; margin: 0 auto; }}
  h1 {{ font-size: 18pt; border-bottom: 2px solid #2a2a2a; padding-bottom: .3em; }}
  .meta {{ color: #666; font-size: 10pt; margin-bottom: 1em; }}
  .warn {{ background: #fff3d6; border: 1px solid #d4a017; padding: .6em 1em; margin: 1em 0; font-size: 10pt; }}
  table {{ border-collapse: collapse; width: 100%; margin-top: 1em; font-size: 11pt; }}
  th, td {{ text-align: left; padding: .35em .6em; border-bottom: 1px solid #ccc; }}
  th {{ background: #f3eee4; }}
  td.alias {{ font-variant-numeric: tabular-nums; width: 6em; font-weight: bold; }}
  .footer {{ margin-top: 2em; color: #888; font-size: 9pt; }}
  @media print {{ .warn {{ display: none; }} }}
</style>
</head>
<body>
<h1>{course_name} — Nyckel</h1>
<div class="meta">Sektion: {section} · Kurs-ID: {course_id} · Genererad: {generated}</div>
<div class="warn"><strong>Hantering:</strong> Skriv ut fysiskt och radera HTML-filen efteråt.
Behåll inte digital kopia. Förvara utskriften säkert.</div>
<table>
<thead><tr><th>Alias</th><th>Riktigt namn</th></tr></thead>
<tbody>
{rows}
</tbody>
</table>
<div class="footer">Antal elever: {count}</div>
</body>
</html>
"""


def build_key(course_id: str) -> Path:
    course = gws_client.get_course(course_id)
    students = gws_client.list_students(course_id)
    aliases = CourseAliases(course_id)
    for s in students:
        aliases.alias_for(s["userId"])
    aliases.save()

    rows_data: list[tuple[str, str]] = []
    for s in students:
        alias = aliases.alias_for(s["userId"])
        rows_data.append((alias, _student_name(s)))
    rows_data.sort(key=lambda r: int(r[0].split()[1]))

    rows_html = "\n".join(
        f"<tr><td class='alias'>{alias}</td><td>{name}</td></tr>"
        for alias, name in rows_data
    )

    html = HTML_TEMPLATE.format(
        course_name=_html_escape(course.get("name", "Kurs")),
        section=_html_escape(course.get("section") or "—"),
        course_id=course_id,
        generated=datetime.now().strftime("%Y-%m-%d %H:%M"),
        rows=rows_html,
        count=len(rows_data),
    )

    KEYS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    out_path = KEYS_DIR / f"{course_id}-{timestamp}.html"
    out_path.write_text(html, encoding="utf-8")
    return out_path


def _html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
