# Project Retrospective

*A living document updated after each milestone. Lessons feed forward into future planning.*

## Milestone: v1.0 — Källkritik - Konspirationsteorier och AI

**Shipped:** 2026-03-15
**Phases:** 5 | **Plans:** 9

### What Was Built
- Komplett undervisningsmoment: 8 lektionsplaner (.docx), 4 presentationer (.pptx), momentöversikt (HTML), exit ticket CSV-export
- 8 Node.js-generatorer som producerar .docx/.pptx programmatiskt
- Backward design-struktur: bedömning (L6/L8) definierad före verktygsbygge (L2/L3) och tillämpning (L4/L5)
- Genomgående pedagogiska mönster: exit ticket-kedja, scaffolding-fade, tredjepersons-framing

### What Worked
- Backward design-ordningen (bedömning → verktyg → tillämpning → avslutning) gav tydlig beroendekedja - varje fas visste exakt vad den byggde mot
- Parallelisering i Wave 1 (två planer samtidigt) sparade tid i fas 2-5
- discuss-phase innan plan-phase gav bra design decisions som styrde planerna
- Generator-mönstret (JS → .docx) möjliggjorde snabb iteration och konsistent format

### What Was Inefficient
- Verifier var disabled - innebar att VERIFICATION.md aldrig skapades, vilket ledde till Nyquist-artefakter som aldrig signerades av
- Lektionstitlar i momentoversikt.html matchade inte alltid momentplan.md (L7 trunkerat) - borde valideras automatiskt
- SUMMARYs saknade one_liner-fält - gjorde accomplishment-extraktion svårare vid milestone completion

### Patterns Established
- 6-fas Rosenshine-lektionsstruktur som standard för alla lektioner
- Exit ticket → retrieval review-kedja mellan alla lektioner
- Tredjepersons-framing som genomgående norm för konspirationskänsligt ämne
- Generator-skript (generate-lektion-N.js) som kanonisk källa, .docx som output

### Key Lessons
1. Discuss-phase är värd investeringen - design decisions tidigt förhindrar omarbete i planerings- och exekveringsfas
2. Backward design fungerar utmärkt för undervisningsmoment - bedömningspunkterna definierar vad allt annat ska leda till
3. Generatorer (JS → .docx) är bättre än manuell docx-skapning - konsistent format, versionerbart, reproducerbart

### Cost Observations
- Model mix: Opus för orchestration och execution, Sonnet för plan-checking och verification
- 59 commits på 5 dagar
- Notable: Fas 5 (avslutning) var snabbast - mönster etablerade, bara 6 min total exekveringstid

---

## Cross-Milestone Trends

### Process Evolution

| Milestone | Phases | Plans | Key Change |
|-----------|--------|-------|------------|
| v1.0 | 5 | 9 | First milestone - established backward design + generator pattern |

### Top Lessons (Verified Across Milestones)

1. (First milestone - lessons to be validated in v2.0)
