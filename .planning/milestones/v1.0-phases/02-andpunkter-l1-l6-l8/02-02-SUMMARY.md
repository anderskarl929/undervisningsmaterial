---
phase: 02-andpunkter-l1-l6-l8
plan: 02
subsystem: undervisningsmaterial
tags: [docx, examination, seminarium, backward-design, equity-safeguards]

requires:
  - phase: 01-strukturell-grund
    provides: momentplan, bedomningskriterier, exit-ticket-mall, tredjepersons-framing
  - phase: 02-andpunkter-l1-l6-l8-plan01
    provides: generate-lektion-1.js, lektion-1.docx, aterkommande case

provides:
  - generate-lektion-6.js och lektion-6.docx (summativ skriftlig examination)
  - generate-lektion-8.js och lektion-8.docx (summativ seminarie-examination)

affects: [03-verktyg-l2-l3, 04-tillampning-l4-l5, 05-avslutning-l7]

tech-stack:
  added: []
  patterns: [6-fas lektionsstruktur for examination, equity-safeguards i seminarium]

key-files:
  created:
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-6.js"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-6.docx"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-8.js"
    - "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-8.docx"
  modified: []

key-decisions:
  - "L6 rubrikforslag kraver analys med tredjepersons-framing, inte beskrivning"
  - "L8 seminarium med 5 tilldelade perspektiv och equity-safeguards fran bedomningskriterier.md"

patterns-established:
  - "Examinationsanpassad 6-fas struktur: kortare intro, langre sjalvstandig fas"
  - "Equity-safeguards som obligatoriskt element i muntlig examination"

requirements-completed: [BED-02, KONSP-03]

duration: 12min
completed: 2026-03-12
---

# Phase 2 Plan 2: Summativa bedomningsandpunkter L6 och L8 Summary

**Summativ skriftlig examination (L6) med rubrikforslag och givet material, plus seminarie-examination (L8) med 5 tilldelade perspektiv och equity-safeguards**

## Performance

- **Duration:** ~12 min (inklusive verifieringspaus)
- **Started:** 2026-03-12
- **Completed:** 2026-03-12
- **Tasks:** 2 (1 auto + 1 checkpoint)
- **Files modified:** 4

## Accomplishments
- Skapat L6 lektionsplan med summativ skrivuppgift (500-800 ord, givet material, E/C/A-kriterier)
- Skapat L8 lektionsplan med seminarium i smagrupper (5 perspektiv, equity-safeguards, inspelning)
- Bada lektionsplanerna refererar aterkommande case fran L1 for koherens
- Anvandaren har godkant bada lektionsplanerna

## Task Commits

Each task was committed atomically:

1. **Task 1: Skapa generate-lektion-6.js och generate-lektion-8.js, generera bada .docx-filer** - `b235042` (feat)
2. **Task 2: Verifiera L6 och L8 lektionsplaner** - checkpoint:human-verify, godkant av anvandaren

## Files Created/Modified
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-6.js` - Node.js-script som genererar L6 lektionsplan
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-6.docx` - Komplett lektionsplan for summativ skrivuppgift
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-8.js` - Node.js-script som genererar L8 lektionsplan
- `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/lektion-8.docx` - Komplett lektionsplan for seminarie-examination

## Decisions Made
- L6 rubrikforslag kraver analys med tredjepersons-framing (inte beskrivning)
- L8 seminarium med 5 tilldelade perspektiv (medieforskare, AI-utvecklare, psykolog, samhallsdebattor, ungdomsrepresentant)
- Equity-safeguards kopierade exakt fran bedomningskriterier.md (tilldelade roller, tredjepersons-framing, privat uppfoljningsplan)

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- Alla tre andpunkter (L1, L6, L8) ar klara
- Fas 2 komplett - redo for Fas 3 (Verktyg L2-L3)
- L1 introducerar case, L6 examinerar skriftligt, L8 examinerar muntligt - backward design-kedjan ar komplett

## Self-Check: PASSED

All 4 files found. Commit b235042 verified.

---
*Phase: 02-andpunkter-l1-l6-l8*
*Completed: 2026-03-12*
