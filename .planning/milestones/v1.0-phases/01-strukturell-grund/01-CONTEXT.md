# Phase 1: Strukturell grund - Context

**Gathered:** 2026-03-10
**Status:** Ready for planning

<domain>
## Phase Boundary

Momentets ramverk definieras - momentplan, bedomningskriterier, exit ticket-mall och epistemisk kultur (tredjepersons-framing) - innan nagon lektion skrivs. Befintlig momentplan.md ska INTE anvandas som utgangspunkt; ny plan byggs fran ROADMAP.md och REQUIREMENTS.md.

</domain>

<decisions>
## Implementation Decisions

### Exit ticket-design
- Blandformat: variera mellan oppna reflektionsfragor och strukturerade fragor beroende pa lektionens innehall
- Minimal mall: fraga + larandemal den mater (inget mer)
- Koppling till retrieval practice: explicit i varje lektionsplan - en sektion "Retrieval review (5 min)" som specificerar vad som aterkopplas baserat pa foregaende exit ticket
- Medium: Egen applikation - fragor levereras som .csv-export i slutet av processen (efter alla lektionsplaner ar klara)

### Momentplan
- Byggs fran grunden baserat pa ROADMAP.md (8 lektioner: 7 undervisning + 1 examination)
- Befintlig momentplan.md ignoreras helt
- Ska beskriva alla 8 lektioner med mal, innehall och progression

### Claude's Discretion
- E/C/A-bedomningskriterier: formulering och detaljeringsniva for skriftlig och muntlig examination, med explicit C-niva
- Tredjepersons-framing: definition av normen, exempelformuleringar, och om det galler alla lektioner eller framst konspirationsteorimomenten
- Momentplanens interna progressionslogik
- Exit ticket-fragebankens specifika fragor per lektion

</decisions>

<specifics>
## Specific Ideas

- Lararen har en egen applikation dar elever svarar pa fragor - exit ticket-fragor ska exporteras som .csv i slutet av hela momentplaneringsprocessen (inte i fas 1)
- CSV-export ar en leverans som sker sist, efter alla lektionsplaner ar fardiga

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `/planera-moment` skill: 7-stegs dialogdriven momentplanering med 6-fas lektionsmodell (Rosenshine), exit ticket-slinga, AI-svaghetsskydd
- `/docx` skill: generera Word-dokument for lektionsplaner
- `/pptx` skill: generera PowerPoint-presentationer
- `/html-momentoversikt` skill: generera HTML-elevsidor

### Established Patterns
- 6-fas lektionsstruktur (Rosenshine): retrieval review -> instruktion -> guidad ovning -> sjalvstandig ovning -> exit ticket -> forhandsvisning
- Exit ticket-slinga: varje lektions exit ticket informerar nasta lektions retrieval review
- E/C/A-bedomning: kunskapskravens tre nivaer (Gy11)

### Integration Points
- Materialoutput sparas i `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/`
- Momentplan som .md, lektionsplaner som .docx, presentationer som .pptx, momentoversikt som .html

</code_context>

<deferred>
## Deferred Ideas

- CSV-export av exit ticket-fragor till lararens egen applikation - levereras efter alla lektionsplaner ar klara (troligen fas 5 eller separat steg)

</deferred>

---

*Phase: 01-strukturell-grund*
*Context gathered: 2026-03-10*
