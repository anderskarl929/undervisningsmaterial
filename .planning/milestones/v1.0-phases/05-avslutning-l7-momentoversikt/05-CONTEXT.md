# Phase 5: Avslutning (L7 + momentöversikt) - Context

**Gathered:** 2026-03-15
**Status:** Ready for planning

<domain>
## Phase Boundary

Momentets avslutningsfas: L7 (reflekterande lektion med metakognitiv reflektion och självbedömning) som .docx med 6-fas struktur, momentöversikt som self-contained HTML-sida för eleverna, och CSV-export av exit ticket-frågor för alla 7 lektioner till lärarens applikation.

</domain>

<decisions>
## Implementation Decisions

### Momentöversikt (HTML)
- Huvudsyfte: förberedelseverktyg - eleverna ser vad som kommer härnäst och var de är i momentet
- Bara lektionsnummer (inga datum) - flexibelt om schemat ändras
- Per lektion: titel + kort beskrivning av vad lektionen handlar om
- En kort bedömningssammanfattning: förklarar att L6 är skriftlig examination och L8 är seminarium, med vad som bedöms i korta drag
- Genereras med `/html-momentoversikt` skill

### CSV-export exit tickets
- Format matchar appens importspecifikation: `topic,type,text,option1,option2,option3,option4,correctAnswer`
- Två frågetyper: MULTIPLE_CHOICE (4 alternativ + korrekt svar) och FREE_TEXT (tomma alternativ)
- Ett topic per lektion, t.ex. "L1 - Konspirationsteorier & AI-labb"
- 1-2 frågor per lektion (matchar exit ticket-mallens minimala design från fas 1)
- UTF-8-kodning, fält med komma i dubbla citattecken
- Frågor hämtas från de befintliga lektionsplanerna (L1-L7) och exit ticket-mallen

### Claude's Discretion
- L7 reflektionsdesign: hur den metakognitiva reflektionen struktureras (individuellt/grupp, skriftligt/muntligt, specifika prompts) - inom ramen för roadmap-kravet "specifika prompts, inte ytlig vad lärde vi oss"
- L7 självbedömning: form och kriterier för elevernas bedömning av sin egen utveckling
- L7 exit ticket-design och retrieval review (baserad på L6)
- HTML-sidans visuella design och layout
- Val av specifika exit ticket-frågor per lektion (inom ramen för 1-2 frågor, blandformat, koppling till lärandemål)
- L7 presentation (.pptx) om det behövs

</decisions>

<specifics>
## Specific Ideas

- CSV-exporten är en uppskjuten leverans från fas 1 - exit ticket-frågorna ska matcha appens importformat exakt (dokumenterat i `Kod/survey-platform/prompts/generera-fragor.md`)
- Momentöversikten ska vara framåtblickande - eleven ska kunna se "vad kommer härnäst?" snarare än ett arkiv
- L7 ska avsluta momentet med genuin metakognitiv reflektion, inte rituell "vad lärde vi oss"-avslutning

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `/docx` skill: generera Word-dokument för L7 lektionsplan
- `/pptx` skill: generera PowerPoint-presentationer
- `/html-momentoversikt` skill: generera self-contained HTML-elevsida
- Fas 1-dokument: momentplan.md, bedomningskriterier.md, tredjepersons-framing.md, exit-ticket-mall.md
- Alla befintliga lektionsplaner: lektion-1.docx till lektion-6.docx, lektion-8.docx
- generate-lektion-7.js finns redan (kan behöva uppdateras)
- CSV-importspecifikation: `Kod/survey-platform/prompts/generera-fragor.md`

### Established Patterns
- 6-fas lektionsstruktur (Rosenshine): retrieval review -> instruktion -> guidad övning -> självständig övning -> exit ticket -> förhandsvisning
- Exit ticket-slinga: varje lektions exit ticket informerar nästa lektions retrieval review
- Tredjepersons-framing genomgående
- Återkommande case (2-3) genom hela momentet
- E/C/A-bedömning med explicit C-nivå

### Integration Points
- Material sparas i `Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/`
- L7 retrieval review baseras på L6 exit ticket
- CSV-export behöver läsa exit ticket-frågor från alla 7 lektionsplaner
- HTML-momentöversikt behöver lektionstitlar och beskrivningar från momentplan.md

</code_context>

<deferred>
## Deferred Ideas

None - discussion stayed within phase scope

</deferred>

---

*Phase: 05-avslutning-l7-momentoversikt*
*Context gathered: 2026-03-15*
