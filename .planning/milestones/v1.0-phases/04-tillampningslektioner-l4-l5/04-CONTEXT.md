# Phase 4: Tillämpningslektioner (L4 + L5) - Context

**Gathered:** 2026-03-15
**Status:** Ready for planning

<domain>
## Phase Boundary

Två lektionsplaner som .docx med 6-fas struktur och tillhörande .pptx: L4 (formativt seminarium med tilldelade perspektiv) och L5 (skriftlig perspektivanalys som dress rehearsal för L6). Scaffolding tillgänglig på begäran men inte utdelad som standard.

</domain>

<decisions>
## Implementation Decisions

### Koppling L4→L5→L6
- L4 och L5 använder samma case/tema - den muntliga diskussionen i L4 förbereder direkt det skriftliga arbetet i L5
- L5 är en mini-L6: samma format (perspektivanalys av givet material) men kortare, ca 300-400 ord
- L5-material: bekant tema från ett av de återkommande casen men en ny källa/text - balans mellan igenkänning och ny utmaning
- Lärarfeedback mellan L5 och L6: kort formativ feedback på varje elevtext (inte kamratrespons)

### Claude's Discretion
- L4 seminariedesign: antal perspektiv, specifika roller, seminariefråga, tidsstruktur (inom ramen för tilldelade perspektiv och equity-safeguards från fas 2)
- L4 vs L8 differentiering: hur det formativa seminariet skiljer sig från det summativa (L8 har 5 perspektiv, inspelning, smågrupper om 4-5)
- Scaffolding-fade: hur avtagande stöd görs konkret, vilka stödstrukturer som finns tillgängliga på begäran
- L5 specifik rubrik och källmaterial (inom ramen för bekant tema + ny källa)
- Presentationernas (.pptx) design och innehåll
- Exit ticket-design för L4 och L5

</decisions>

<specifics>
## Specific Ideas

- Progressionen muntligt→skriftligt (L4→L5) med samma case skapar en naturlig övergång: eleverna har redan tänkt igenom argumenten innan de skriver
- L5 som dress rehearsal innebär att eleverna övar exakt det format de möter i L6, men med lägre insats (formativt, kortare)
- Lärarfeedback mellan L5 och L6 ger eleverna konkreta förbättringsområden att arbeta med inför examinationen

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `/docx` skill: generera Word-dokument för lektionsplaner
- `/pptx` skill: generera PowerPoint-presentationer
- Fas 1-dokument: momentplan.md, bedomningskriterier.md, tredjepersons-framing.md, exit-ticket-mall.md
- Fas 2-dokument: lektion-1.docx (L1), lektion-6.docx (L6 summativ), lektion-8.docx (L8 seminarium)
- Fas 3-dokument: lektion-2.docx (L2), lektion-3.docx (L3) med gateway check
- Befintliga generate-scripts: generate-lektion-4.js, generate-lektion-5.js finns redan

### Established Patterns
- 6-fas lektionsstruktur (Rosenshine): retrieval review -> instruktion -> guidad övning -> självständig övning -> exit ticket -> förhandsvisning
- Exit ticket-slinga: varje lektions exit ticket informerar nästa lektions retrieval review
- Tredjepersons-framing genomgående med förstärkta påminnelser i L4
- Tilldelade perspektiv + equity-safeguards (definierade i fas 2)
- Återkommande case (2-3) genom hela momentet

### Integration Points
- Material sparas i `Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/`
- L4 retrieval review baseras på L3 exit ticket (gateway check-resultat informerar gruppsammansättning)
- L5 retrieval review baseras på L4 exit ticket
- L5 perspektivanalys förbereder specifikt för L6-formatet (samma uppgiftstyp)
- Lärarfeedback på L5 ges mellan L5 och L6

</code_context>

<deferred>
## Deferred Ideas

None - discussion stayed within phase scope

</deferred>

---

*Phase: 04-tillampningslektioner-l4-l5*
*Context gathered: 2026-03-15*
