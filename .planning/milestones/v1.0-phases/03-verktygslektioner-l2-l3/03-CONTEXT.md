# Phase 3: Verktygslektioner (L2 + L3) - Context

**Gathered:** 2026-03-12
**Status:** Ready for planning

<domain>
## Phase Boundary

Två lektionsplaner som .docx med 6-fas struktur och tillhörande .pptx: L2 (källkritiska verktyg, SIFT, lateral reading, Harvard-referering) och L3 (AI+konspirationsteori-intersection med stationsarbete). L3 exit ticket fungerar som gateway check inför L4.

</domain>

<decisions>
## Implementation Decisions

### L3 AI+konspirations-intersection
- Mekanismfokus: presentera 3-4 mekanismer (t.ex. deepfakes, filterbubbla, AI-genererade källor, automatiserade bottar) som eleverna utforskar
- Stationsarbete: varje station = en mekanism, eleverna roterar i par/grupper och analyserar exempel vid varje station med strukturerade frågor
- Mix av case: några stationer använder de återkommande casen från L1, andra har nya träffsäkra exempel som bättre illustrerar just den mekanismen
- Strukturerad koppling till källkritik: varje station har en explicit fråga där eleverna använder SIFT/grundfrågorna för att granska det AI-genererade exemplet - de övar L2-verktygen på nytt material

### Gateway check (L3 exit ticket)
- Mäter verktygstillämpning: kan eleven självständigt tillämpa källkritiska verktyg på ett nytt exempel utan scaffolding?
- Resultaten används för gruppsammansättning inför L4-seminariet - blanda starkare/svagare elever strategiskt
- Två delar: Del 1 = gateway-fråga (källkritisk tillämpning på nytt exempel), Del 2 = kort Harvard-refereringsövning (formatera källor korrekt)
- Enkel checklista som tröskel: 2-3 konkreta kriterier (t.ex. "använder minst 2 källkritiska verktyg", "identifierar rätt problem") - snabbt för läraren att sortera

### Claude's Discretion
- L2 övningsdesign: hur hands-on övningen med SIFT/lateral reading struktureras, vilka verkliga källor som används, balans mellan grundfrågor och SIFT
- Harvard-referering introduktion i L2: hur mycket tid, förenklad eller fullständig modell
- Antal och val av specifika mekanismer för L3-stationerna (inom ramen för mekanismfokus)
- Stationsarbetets praktiska logistik (tid per station, antal stationer, rotationsordning)
- Presentationernas (.pptx) design och innehåll
- Exakta gateway check-kriterier (inom ramen för enkel checklista)

</decisions>

<specifics>
## Specific Ideas

- L3 är "ny pedagogisk mark" (från STATE.md) - AI+konspirationsteori-intersektionen har få pedagogiska modeller att följa, research-agenten bör söka brett
- Stationernas strukturerade källkritik-koppling skapar en naturlig brygga mellan L2 (verktyg) och L3 (tillämpning) - eleverna ser att verktygen fungerar på AI-innehåll
- Gateway check med gruppsammansättning istället för scaffolding-differentiering - läraren blandar grupper strategiskt inför L4

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `/docx` skill: generera Word-dokument för lektionsplaner
- `/pptx` skill: generera PowerPoint-presentationer
- Fas 1-dokument: momentplan.md, bedomningskriterier.md, tredjepersons-framing.md, exit-ticket-mall.md
- Fas 2-dokument: lektion-1.docx (L1 med AI-labb), lektion-6.docx (L6 summativ), lektion-8.docx (L8 seminarium)
- Befintliga generate-scripts: generate-lektion-2.js, generate-lektion-3.js finns redan

### Established Patterns
- 6-fas lektionsstruktur (Rosenshine): retrieval review -> instruktion -> guidad övning -> självständig övning -> exit ticket -> förhandsvisning
- Exit ticket-slinga: varje lektions exit ticket informerar nästa lektions retrieval review
- Tredjepersons-framing genomgående
- Återkommande case (2-3) introducerade i L1

### Integration Points
- Material sparas i `Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/`
- L2 retrieval review baseras på L1 exit ticket
- L3 retrieval review baseras på L2 exit ticket
- L3 exit ticket informerar L4 gruppsammansättning (gateway)

</code_context>

<deferred>
## Deferred Ideas

None - discussion stayed within phase scope

</deferred>

---

*Phase: 03-verktygslektioner-l2-l3*
*Context gathered: 2026-03-12*
