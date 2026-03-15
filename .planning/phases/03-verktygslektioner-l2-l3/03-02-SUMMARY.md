---
phase: 03-verktygslektioner-l2-l3
plan: 02
subsystem: undervisningsmaterial
tags: [docx, pptx, kallkritik, AI, konspirationsteorier, stationsarbete, SIFT, gateway-check]

# Dependency graph
requires:
  - phase: 03-verktygslektioner-l2-l3
    provides: "L2 lektionsplan med SIFT, grundfragor, lateral reading (03-01)"
provides:
  - "L3 lektionsplan med 4 stationer (deepfakes, filterbubbla, AI-kallor, bottar)"
  - "L3 gateway exit ticket med 3-punkts checklista for L4-gruppering"
  - "L3 PPTX-presentation for instruktions- och stationsmoment"
affects: [04-seminarier-l4-l5]

# Tech tracking
tech-stack:
  added: []
  patterns: [stationsarbete-4-stationer, gateway-check-checklista, tredjepersons-framing-per-station]

key-files:
  created:
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-3.docx"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-3.pptx"
  modified:
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-3.js"

key-decisions:
  - "Stationsarbete i par (inte grupper) for att sakerstalla att varje elev ar aktiv"
  - "Gateway check med 3-punkts checklista avgur L4-gruppsammansattning (stark/medel/behover stod)"
  - "Mix av aterkommande case (deepfakes valet 2026, DeepSeek) och nya exempel (filterbubbla, bottar)"

patterns-established:
  - "Gateway check: exit ticket som informerar nasta lektions gruppering"
  - "Stationsblad som bilagor med bakgrundstext + fragor + reflektionsfraga"

requirements-completed: [KONSP-01, KONSP-04, AI-01]

# Metrics
duration: 18min
completed: 2026-03-15
---

# Phase 3 Plan 2: L3 AI och konspirationsteorier Summary

**L3 stationsarbete med 4 AI-mekanismstationer (deepfakes, filterbubbla, AI-kallor, bottar), gateway exit ticket utan scaffolding, och PPTX-presentation**

## Performance

- **Duration:** 18 min
- **Started:** 2026-03-15T00:00:00Z
- **Completed:** 2026-03-15T00:18:00Z
- **Tasks:** 2
- **Files modified:** 3

## Accomplishments
- Komplett L3 lektionsplan med 6-fas struktur (80 min) och stationsarbete
- 4 stationsblad med explicit SIFT/grundfrage-koppling och tredjepersons-framing vid varje station
- Gateway exit ticket utan scaffolding + Harvard-referering + 3-punkts checklista for lararsortering
- PPTX-presentation med slides for mekanismer, stationsarbete och exit ticket

## Task Commits

Each task was committed atomically:

1. **Task 1: Skriv om generate-lektion-3.js och generera lektion-3.docx + lektion-3.pptx** - `782164a` (feat)
2. **Task 2: Verifiera L3 lektionsplan och stationsarbete** - checkpoint:human-verify (godkant)

## Files Created/Modified
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-3.js` - Omskrivet fran grunden med 6-fas struktur och stationsarbete
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-3.docx` - Komplett lektionsplan med 6 bilagor (stationsblad, exit ticket, checklista)
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-3.pptx` - Presentation for instruktion och stationsarbete

## Decisions Made
- Stationsarbete i par (inte storre grupper) for att sakerstalla aktiv deltagande
- Gateway check med 3-punkts checklista avgur L4-gruppsammansattning
- Mix av aterkommande case (deepfakes valet 2026, DeepSeek) och nya exempel (filterbubbla, bottar)

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- L3 komplett med gateway check som informerar L4-gruppering
- Redo for fas 4 (seminarier L4-L5)
- Gateway checklistan ger lararen underlag for att skapa heterogena seminariegrupper

## Self-Check: PASSED

- FOUND: generate-lektion-3.js
- FOUND: lektion-3.docx
- FOUND: lektion-3.pptx
- FOUND: commit 782164a

---
*Phase: 03-verktygslektioner-l2-l3*
*Completed: 2026-03-15*
