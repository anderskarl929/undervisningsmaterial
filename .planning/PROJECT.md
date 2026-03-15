# Källkritik: Konspirationsteorier och AI

## What This Is

Ett komplett undervisningsmoment i Samhällskunskap 3 om källkritik, med konspirationsteorier som huvudingång och AI:s roll i informationslandskapet som kompletterande perspektiv. Momentet omfattar 8 lektioner à 80 minuter med backward design, 6-fas Rosenshine-struktur, exit ticket-kedja och scaffolding-fade. Riktar sig till gymnasieelever och ger dem analytiska modeller och praktiska verktyg för att granska konspirationsteorier och AI-genererat innehåll - utan att avfärda fenomenen.

## Core Value

Eleverna ska kunna granska konspirationsteorier och AI-genererat innehåll med intellektuell ärlighet - förstå varför de uppstår och sprids, inte bara avfärda dem.

## Requirements

### Validated

- ✓ KRIT-01: Källkritiska grundfrågor på digitalt innehåll - v1.0
- ✓ KRIT-02: SIFT-metoden/lateral reading - v1.0
- ✓ KRIT-03: Harvard-referering - v1.0
- ✓ KRIT-04: Digitala verifieringsverktyg (omvänd bildsökning, Google Lens) - v1.0
- ✓ KONSP-01: Psykologiska mekanismer bakom konspirationstro - v1.0
- ✓ KONSP-02: Inokulering/prebunking genom AI-labb - v1.0
- ✓ KONSP-03: Aktuella svenska/internationella exempel 2025-2026 - v1.0
- ✓ KONSP-04: AI:s roll i konspirationsspridning - v1.0
- ✓ AI-01: Källkritisk granskning av AI-genererat innehåll - v1.0
- ✓ AI-02: Metakognitiv reflektion - v1.0
- ✓ BED-01: Exit ticket-kedja med retrieval review - v1.0
- ✓ BED-02: Seminarium med equity-safeguards - v1.0
- ✓ BED-03: Tredjepersons-framing genomgående - v1.0
- ✓ MAT-01: 7 lektionsplaner som .docx med 6-fas struktur - v1.0
- ✓ MAT-02: Presentationer som .pptx - v1.0
- ✓ MAT-03: Momentöversikt som HTML-sida - v1.0

### Active

(None - v1.0 complete. Define new requirements with `/gsd:new-milestone`)

### Out of Scope

- Djupgående mediekunskap/journalistik - fokus är källkritik, inte medieproduktion
- Teknikdjupdykning i hur AI fungerar - fokus är på att granska output, inte förstå modellerna
- Debatt om yttrandefrihet/censur - relevant men ett eget moment
- Logiska felslut och argumentationsanalys som explicit verktyg (v2 kandidat: FORD-01)
- Bad News-spelet som komplement till AI-labben (v2 kandidat: FORD-02)
- Scaffolding-fade med explicita per-elev-triggar (v2 kandidat: FORD-03)

## Context

Shipped v1.0 med 8 lektionsplaner, 4 presentationer, momentöversikt HTML och exit ticket CSV.
Material i `Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/`.
8 Node.js-generatorer (generate-lektion-1.js - generate-lektion-8.js) producerar .docx/.pptx.

## Constraints

- **Format:** Lektionsplaner som .docx, presentationer som .pptx, momentöversikt som .html
- **Skill-integration:** Materialet genereras med /planera-moment-skillens struktur (6-fas lektionsmodell, exit ticket-slinga, AI-svaghetskontroll)
- **Kurs:** Samhällskunskap 3 - kunskapskrav och centralt innehåll enligt Gy11
- **Omfång:** 8 lektioner à 80 minuter

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Seminarium som examination | Källkritik kräver resonemang och argumentation - passar muntlig form | ✓ Good - L8 med 5 perspektiv och equity-safeguards |
| Konspirationsteorier med respekt | Avfärdande stänger lärande - intellektuell ärlighet öppnar | ✓ Good - tredjepersons-framing genomgående |
| AI som källa att granska | Mest relevant vinkel - eleverna möter AI-innehåll dagligen | ✓ Good - AI-labb i L1, stationsarbete i L3 |
| GSD-ramverk runt /planera-moment | Löser context rot och möjliggör forskningsfas | ✓ Good - 5 faser, 9 planer, rent kontextfönster |
| Backward design (bedömning före instruktion) | Forskning visar bättre alignment | ✓ Good - L6/L8 definierade i fas 2 |
| L3-L4 gateway check | Exit ticket avgör scaffolding-nivå | ✓ Good - differentiering utan stigmatisering |
| Tredjepersons-framing som norm | Konspirationskänsligt ämne kräver distans | ✓ Good - alla 8 lektioner, graded criterion i L8 |
| Scaffolding-fade L2→L5 | Gradvis ökat ansvar för eleverna | ✓ Good - L2 full → L3 guidad → L4/L5 på begäran |

---
*Last updated: 2026-03-15 after v1.0 milestone*
