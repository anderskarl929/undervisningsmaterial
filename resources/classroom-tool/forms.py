"""Hämta och formatera svar från Google Forms-uppgifter i Classroom.

Detektion: en Classroom-uppgift med en bifogad Form har
`coursework.materials[].form.formUrl`. Form-id parsas från URL:n.

Mapping av svar till elev-alias kräver att Forms-formuläret samlar
email-adresser ("Samla in e-postadresser" på i Forms-inställningar).
Vi mappar respondentEmail → userId via students.list (kräver scope
`classroom.profile.emails`) → alias via aliases.json.

Quiz-poäng (om formuläret är ett quiz) visas per-fråga och totalt.
"""

from __future__ import annotations

import re
from datetime import datetime

import gws_client
from anonymize import CourseAliases

GDPR_BANNER = (
    "> ⚠️ **Forms-svar kan innehålla identifierande info** i fritext-fält "
    "(\"jag heter X\", \"min mamma X\"). Pseudonymerna `Elev N` bryter "
    "kopplingen bakåt men fritext är inte filtrerad. Behandla denna fil "
    "därefter — radera när du är klar."
)

_FORM_URL_RE = re.compile(r"/forms/d/(?:e/)?([a-zA-Z0-9_-]{20,})")


def extract_form_id(coursework: dict) -> str | None:
    """Returnerar form-id från en coursework-payload, eller None.

    Letar efter `materials[].form.formUrl` (dedikerat Forms-attachment) och
    parsar id ur URL:en. Exempel:
      https://docs.google.com/forms/d/<ID>/edit  ->  <ID>
    """
    for m in coursework.get("materials") or []:
        form = m.get("form") or {}
        url = form.get("formUrl") or ""
        match = _FORM_URL_RE.search(url)
        if match:
            return match.group(1)
    return None


def _question_type_label(q: dict) -> str:
    """Kort etikett för fråge-typ."""
    if "textQuestion" in q:
        return "långsvar" if q["textQuestion"].get("paragraph") else "kortsvar"
    if "choiceQuestion" in q:
        kind = (q["choiceQuestion"].get("type") or "").lower()
        return {"radio": "flerval", "checkbox": "kryssval", "drop_down": "rullista"}.get(
            kind, kind or "val"
        )
    if "scaleQuestion" in q:
        s = q["scaleQuestion"]
        return f"skala {s.get('low','?')}-{s.get('high','?')}"
    if "dateQuestion" in q:
        return "datum"
    if "timeQuestion" in q:
        return "tid"
    if "fileUploadQuestion" in q:
        return "fil-uppladdning"
    if "ratingQuestion" in q:
        return "betyg"
    return "fråga"


def _extract_questions(form: dict) -> list[dict]:
    """Plattar items[] till en lista av frågor i ordning.

    Returnerar bara items som har `questionItem.question` — section-headers
    och rena bilder/text hoppas över.
    """
    out: list[dict] = []
    for idx, item in enumerate(form.get("items") or [], start=1):
        qi = item.get("questionItem") or {}
        q = qi.get("question") or {}
        qid = q.get("questionId")
        if not qid:
            continue
        out.append(
            {
                "index": idx,
                "question_id": qid,
                "title": item.get("title") or "(utan titel)",
                "type": _question_type_label(q),
                "max_points": (q.get("grading") or {}).get("pointValue"),
            }
        )
    return out


def _format_answer(answer: dict | None) -> str:
    """Formaterar en svars-payload till sträng. Inget svar -> '_(tomt)_'."""
    if not answer:
        return "_(tomt)_"
    text_answers = (answer.get("textAnswers") or {}).get("answers") or []
    if text_answers:
        values = [a.get("value", "") for a in text_answers if a.get("value")]
        if not values:
            return "_(tomt)_"
        if len(values) == 1:
            return values[0]
        # Flera värden (checkbox-fråga): semikolon-separerad lista
        return "; ".join(values)
    if "fileUploadAnswers" in answer:
        files = (answer["fileUploadAnswers"] or {}).get("answers") or []
        return f"_(filuppladdning: {len(files)} fil(er) — visas inte)_"
    return "_(okänd svarstyp)_"


def _fmt_dt(s: str | None) -> str:
    if not s:
        return "—"
    m = re.match(r"(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})", s)
    return f"{m.group(1)} {m.group(2)}" if m else s


def _fetch_data(form_id: str) -> tuple[dict, list[dict]]:
    """Hämtar form-strukturen och alla svar."""
    try:
        form = gws_client.forms_get(form_id)
        responses = gws_client.forms_responses_list(form_id)
    except RuntimeError as e:
        msg = str(e)
        if "insufficient authentication scopes" in msg or "403" in msg:
            raise SystemExit(
                "Forms API kräver scopes som saknas i din nuvarande inloggning. "
                "Re-logga in:\n\n"
                "  gws auth login --scopes \\\n"
                "    https://www.googleapis.com/auth/forms.body.readonly,\\\n"
                "    https://www.googleapis.com/auth/forms.responses.readonly,\\\n"
                "    https://www.googleapis.com/auth/classroom.profile.emails,\\\n"
                "    <plus dina befintliga scopes — se README>\n\n"
                f"(orsak: {msg.splitlines()[0]})"
            )
        raise
    return form, responses


def _is_quiz(form: dict) -> bool:
    """Heuristik: om någon fråga har grading.pointValue är formuläret ett quiz."""
    for item in form.get("items") or []:
        q = ((item.get("questionItem") or {}).get("question")) or {}
        if q.get("grading", {}).get("pointValue") is not None:
            return True
    return False


def _coursework_header(course: dict, work: dict, form: dict, form_id: str) -> list[str]:
    out: list[str] = []
    title = work.get("title") or form.get("info", {}).get("title") or "(uppgift utan titel)"
    out.append(f"# {title}")
    out.append("")
    out.append(GDPR_BANNER)
    out.append("")
    out.append(f"_Kurs:_ {course.get('name', '—')} · `{course.get('id', '—')}`  ")
    out.append(f"_Uppgift-ID:_ `{work.get('id', '—')}`  ")
    out.append(f"_Form:_ {form.get('info', {}).get('title', '—')} · `{form_id}`  ")
    out.append(f"_Genererad:_ {datetime.now().strftime('%Y-%m-%d %H:%M')}  ")
    if work.get("alternateLink"):
        out.append(f"_Länk:_ {work['alternateLink']}  ")
    out.append("")
    return out


def _format_questions_overview(questions: list[dict], is_quiz: bool) -> list[str]:
    out = [f"## Frågor ({len(questions)})", ""]
    for q in questions:
        suffix = f" · {q['type']}"
        if is_quiz and q["max_points"] is not None:
            suffix += f" · {q['max_points']}p"
        out.append(f"{q['index']}. **{q['title']}**{suffix}")
    out.append("")
    return out


def _format_response_block(
    alias: str,
    response: dict,
    questions: list[dict],
    is_quiz: bool,
) -> list[str]:
    out: list[str] = []
    out.append(f"## {alias}")
    submitted = _fmt_dt(response.get("lastSubmittedTime") or response.get("createTime"))
    score_part = ""
    if is_quiz and response.get("totalScore") is not None:
        score_part = f" · Poäng: {response['totalScore']}"
    out.append(f"_Inlämnat: {submitted}{score_part}_")
    out.append("")

    answers_by_qid = response.get("answers") or {}
    for q in questions:
        ans = answers_by_qid.get(q["question_id"])
        text = _format_answer(ans)
        score_suffix = ""
        if is_quiz and ans:
            grade = ans.get("grade") or {}
            if "score" in grade and q["max_points"] is not None:
                mark = "✓" if grade.get("correct") else "✗"
                score_suffix = f" _({grade['score']}/{q['max_points']}p {mark})_"
            elif "correct" in grade:
                score_suffix = " _(✓)_" if grade["correct"] else " _(✗)_"
        out.append(f"**{q['index']}. {q['title']}**{score_suffix}")
        out.append("")
        out.append(text)
        out.append("")
    return out


def _annotate_responses(
    course_id: str,
    responses: list[dict],
) -> tuple[list[tuple[int, str, dict]], list[dict]]:
    """Returnerar (annotated, orphans).

    annotated: [(alias_num, alias, response), ...] sorterad — svar vi kunde
    mappa mot en elev i kursen.
    orphans: svar med email som inte finns i kursens roster (externa svar
    eller saknad scope) — listas separat utan att slänga datan.
    """
    aliases = CourseAliases(course_id)
    email_index = gws_client.students_email_index(course_id)

    # Säkerställ att alla studenter har alias innan vi formaterar
    for uid in email_index.values():
        aliases.alias_for(uid)
    aliases.save()

    annotated: list[tuple[int, str, dict]] = []
    orphans: list[dict] = []
    for r in responses:
        email = (r.get("respondentEmail") or "").lower()
        uid = email_index.get(email) if email else None
        if not uid:
            orphans.append(r)
            continue
        alias = aliases.alias_for(uid)
        annotated.append((int(alias.split()[1]), alias, r))
    annotated.sort(key=lambda t: t[0])
    return annotated, orphans


def build_form_dump(course_id: str, coursework_id: str, form_id: str) -> str:
    course = gws_client.get_course(course_id)
    work = gws_client.get_coursework(course_id, coursework_id)
    form, responses = _fetch_data(form_id)
    questions = _extract_questions(form)
    is_quiz = _is_quiz(form)

    annotated, orphans = _annotate_responses(course_id, responses)

    # Räkna saknas-status från Classroom-submissions (TURNED_IN/CREATED)
    submissions = gws_client.list_submissions(course_id, coursework_id)
    aliases = CourseAliases(course_id)
    missing: list[str] = []
    for s in submissions:
        if s.get("state") in ("CREATED", "NEW"):
            uid = s.get("userId")
            if uid:
                missing.append(aliases.alias_for(uid))

    lines = _coursework_header(course, work, form, form_id)
    total = len(annotated) + len(orphans)
    summary_parts = [f"{total} svar"]
    if orphans:
        summary_parts.append(f"{len(orphans)} extern(a) (ej kursmedlem)")
    if missing:
        summary_parts.append(f"{len(missing)} saknar svar")
    if is_quiz:
        summary_parts.append("quiz-läge")
    lines.append(f"_Status:_ " + " · ".join(summary_parts))
    lines.append("")

    lines.extend(_format_questions_overview(questions, is_quiz))
    lines.append("---")
    lines.append("")

    for _, alias, r in annotated:
        lines.extend(_format_response_block(alias, r, questions, is_quiz))
        lines.append("---")
        lines.append("")

    for i, r in enumerate(orphans, start=1):
        email_hint = (r.get("respondentEmail") or "").split("@")[0]
        anon_label = f"Externt svar {i}"
        # Lägg INTE in raw email — bara markera att det är externt
        lines.extend(_format_response_block(anon_label, r, questions, is_quiz))
        lines.append("---")
        lines.append("")

    if missing:
        lines.append(f"## Saknar svar ({len(missing)})")
        lines.append(", ".join(sorted(missing, key=lambda a: int(a.split()[1]))))
        lines.append("")

    return "\n".join(lines)


def build_form_read(
    course_id: str,
    coursework_id: str,
    form_id: str,
    alias: str,
) -> str:
    """En elevs svar på Form-uppgift."""
    aliases = CourseAliases(course_id)
    user_id = aliases.user_id_for_alias(alias)
    if not user_id:
        # Synka roster och försök igen
        for s in gws_client.list_students(course_id):
            aliases.alias_for(s["userId"])
        aliases.save()
        user_id = aliases.user_id_for_alias(alias)
    if not user_id:
        raise SystemExit(
            f"Hittar inte aliaset {alias!r} i kurs {course_id}. "
            f"Kör `summary {course_id}` först."
        )

    email_index = gws_client.students_email_index(course_id)
    student_email = next(
        (e for e, uid in email_index.items() if uid == user_id), None
    )
    if not student_email:
        raise SystemExit(
            f"Hittar ingen email för {alias} (userId {user_id}). "
            f"Kontrollera att scopet `classroom.profile.emails` är aktiverat."
        )

    course = gws_client.get_course(course_id)
    work = gws_client.get_coursework(course_id, coursework_id)
    form, responses = _fetch_data(form_id)
    questions = _extract_questions(form)
    is_quiz = _is_quiz(form)

    match = next(
        (r for r in responses if (r.get("respondentEmail") or "").lower() == student_email),
        None,
    )
    if not match:
        raise SystemExit(f"Inget Form-svar hittat för {alias} på uppgift {coursework_id}.")

    lines = _coursework_header(course, work, form, form_id)
    lines.extend(_format_response_block(alias, match, questions, is_quiz))
    return "\n".join(lines)
