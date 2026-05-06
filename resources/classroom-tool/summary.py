"""Genererar anonymiserad markdown-sammanställning för en kurs.

Inga elev-namn eller userId:n förekommer i output — endast pseudonymer (Elev N).
Inget innehåll i elevernas inlämningar visas; bara metadata (titel, deadline,
inlämningsstatus, antal saknas).
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime, timezone

import gws_client
from anonymize import CourseAliases


def _fmt_due(work: dict) -> str:
    due = work.get("dueDate")
    if not due:
        return "—"
    y, m, d = due.get("year"), due.get("month"), due.get("day")
    if not (y and m and d):
        return "—"
    t = work.get("dueTime") or {}
    hh, mm = t.get("hours"), t.get("minutes")
    base = f"{y:04d}-{m:02d}-{d:02d}"
    if hh is not None:
        return f"{base} {hh:02d}:{mm or 0:02d}"
    return base


def _is_overdue(work: dict) -> bool:
    due = work.get("dueDate")
    if not due:
        return False
    y, m, d = due.get("year"), due.get("month"), due.get("day")
    if not (y and m and d):
        return False
    t = work.get("dueTime") or {}
    hh = t.get("hours") or 23
    mm = t.get("minutes") or 59
    try:
        when = datetime(y, m, d, hh, mm, tzinfo=timezone.utc)
    except ValueError:
        return False
    return when < datetime.now(timezone.utc)


def build_summary(course_id: str) -> str:
    course = gws_client.get_course(course_id)
    students = gws_client.list_students(course_id)
    coursework = gws_client.list_coursework(course_id)
    announcements = gws_client.list_announcements(course_id, 5)

    aliases = CourseAliases(course_id)
    for s in students:
        aliases.alias_for(s["userId"])
    aliases.save()

    lines: list[str] = []
    lines.append(f"# {course.get('name', 'Kurs')}\n")
    section = course.get("section")
    if section:
        lines.append(f"_Sektion:_ {section}  ")
    lines.append(f"_Kurs-ID:_ `{course_id}`  ")
    lines.append(f"_Status:_ {course.get('courseState', '—')}  ")
    if course.get("alternateLink"):
        lines.append(f"_Länk:_ {course['alternateLink']}\n")

    lines.append(f"## Elever ({len(students)})")
    lines.append(f"_Anonymiserade pseudonymer. Kör `key {course_id}` för utskrivbar nyckel._\n")
    aliases_list = sorted(
        (aliases.alias_for(s["userId"]) for s in students),
        key=lambda a: int(a.split()[1]),
    )
    for alias in aliases_list:
        lines.append(f"- {alias}")
    lines.append("")

    lines.append(f"## Uppgifter ({len(coursework)})")
    if not coursework:
        lines.append("_(inga publicerade uppgifter)_")
    student_count = len(students)

    for work in coursework:
        title = work.get("title", "(utan titel)")
        state = work.get("state", "")
        due = _fmt_due(work)
        overdue_marker = " ⚠️ förfallen" if _is_overdue(work) and state == "PUBLISHED" else ""
        wtype = work.get("workType", "ASSIGNMENT")
        lines.append(f"\n### {title}")
        lines.append(f"- **Typ:** {wtype}")
        lines.append(f"- **Status:** {state}")
        lines.append(f"- **Deadline:** {due}{overdue_marker}")
        if work.get("maxPoints") is not None:
            lines.append(f"- **Poäng:** {work['maxPoints']}")
        if work.get("alternateLink"):
            lines.append(f"- **Länk:** {work['alternateLink']}")

        if state != "PUBLISHED" or student_count == 0:
            continue
        try:
            subs = gws_client.list_submissions(course_id, work["id"])
        except RuntimeError as e:
            lines.append(f"- _Kunde inte läsa inlämningar:_ {e}")
            continue

        states_count = Counter(s.get("state", "UNKNOWN") for s in subs)
        late = sum(1 for s in subs if s.get("late"))
        turned_in = states_count.get("TURNED_IN", 0) + states_count.get("RETURNED", 0)
        graded = sum(1 for s in subs if s.get("assignedGrade") is not None)

        lines.append(
            f"- **Inlämnat:** {turned_in}/{student_count}"
            + (f" (varav {late} sent)" if late else "")
        )
        if graded:
            lines.append(f"- **Bedömt:** {graded}/{turned_in}")

        missing = [s for s in subs if s.get("state") in ("CREATED", "NEW")]
        if missing:
            missing_aliases = sorted(
                (aliases.alias_for(s["userId"]) for s in missing),
                key=lambda a: int(a.split()[1]),
            )
            preview = ", ".join(missing_aliases[:8])
            more = f" +{len(missing_aliases) - 8} till" if len(missing_aliases) > 8 else ""
            lines.append(f"- **Saknas inlämning:** {preview}{more}")

    lines.append("")
    lines.append("## Senaste meddelanden")
    if not announcements:
        lines.append("_(inga eller ingen åtkomst)_")
    for a in announcements:
        text = (a.get("text") or "").strip().splitlines()[0][:120] if a.get("text") else ""
        updated = a.get("updateTime", "")[:10]
        link = a.get("alternateLink", "")
        lines.append(f"- _{updated}_ {text} — {link}")
    lines.append("")

    return "\n".join(lines)
