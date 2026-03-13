---
phase: 03-verktygslektioner-l2-l3
plan: 01
subsystem: undervisningsmaterial
tags: [docx, pptx, kallkritik, SIFT, lateral-reading, harvard-referering]

requires:
  - phase: 02-andpunkter-l1-l6-l8
    provides: "L1 lektionsplan med exit ticket som L2 retrieval review baseras pa"
provides:
  - "L2 lektionsplan (lektion-2.docx) med 6-fas struktur och 80 min tidsbudget"
  - "L2 presentation (lektion-2.pptx) med 9 slides for instruktionsmoment"
  - "generate-lektion-2.js omskriven fran grunden med hjalpfunktioner fran generate-lektion-1.js"
affects: [03-verktygslektioner-l2-l3, 04-tillampning]

tech-stack:
  added: []
  patterns: ["6-fas lektionsstruktur med retrieval review och exit ticket"]

key-files:
  created:
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-2.js"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-2.docx"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-2.pptx"
  modified: []

key-decisions:
  - "L2 anvander 6-fas struktur med SIFT, grundfragor, lateral reading, omvand bildsokning och Harvard-referering intro"
  - "Retrieval review baseras pa L1 exit ticket om AI-trovardighetsmorkorer"

patterns-established:
  - "generate-lektion-X.js ateranvander hjalpfunktioner fran generate-lektion-1.js"

requirements-completed: [KRIT-01, KRIT-02, KRIT-03, KRIT-04]

duration: 15min
completed: 2026-03-13
---

# Phase 3 Plan 1: L2 Kallkritiska verktyg Summary

**L2 lektionsplan med SIFT-metoden, fyra kallkritiska grundfragor, lateral reading, omvand bildsokning och Harvard-referering intro som docx + pptx**

## Performance

- **Duration:** ~15 min (across continuation)
- **Started:** 2026-03-13
- **Completed:** 2026-03-13
- **Tasks:** 2 (1 auto + 1 checkpoint)
- **Files modified:** 3

## Accomplishments
- L2 lektionsplan (lektion-2.docx) med komplett 6-fas struktur och 80 min tidsbudget
- L2 presentation (lektion-2.pptx) med 9 slides for instruktionsmoment (SIFT, grundfragor, lateral reading, Harvard)
- Retrieval review kopplad till L1 exit ticket (AI-trovardighetsmorkorer)
- Exit ticket matchar exit-ticket-mall.md exakt
- Tredjepersons-framing paminnelse inkluderad

## Task Commits

Each task was committed atomically:

1. **Task 1: Skriv om generate-lektion-2.js och generera lektion-2.docx + lektion-2.pptx** - `1985767` (feat)
2. **Task 2: Verifiera L2 lektionsplan** - checkpoint:human-verify (godkant)

## Files Created/Modified
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-2.js` - Node.js-script som genererar lektion-2.docx och lektion-2.pptx
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-2.docx` - Komplett lektionsplan for L2 med 6-fas struktur
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-2.pptx` - Presentation for L2 instruktionsmoment

## Decisions Made
- L2 anvander 6-fas struktur med alla kallkritiska grundverktyg samlade i en lektion
- Retrieval review baseras pa L1 exit ticket for att skapa koppling mellan lektionerna
- Harvard-referering begransas till max 9 min intro - byggs vidare i L3 och L5

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- L2 verktygslektion klar - redo for Plan 03-02 (L3)
- L3 kan bygga pa SIFT och grundfragor fran L2
- Harvard-referering intro i L2 forbereds for L3 exit ticket och L5 fordjupning

## Self-Check: PASSED

- FOUND: generate-lektion-2.js
- FOUND: lektion-2.docx
- FOUND: lektion-2.pptx
- FOUND: commit 1985767

---
*Phase: 03-verktygslektioner-l2-l3*
*Completed: 2026-03-13*
