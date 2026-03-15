# Phase 2: Ändpunkter (L1 + L6 + L8) - Context

**Gathered:** 2026-03-12
**Status:** Ready for planning

<domain>
## Phase Boundary

Tre lektionsplaner som .docx med 6-fas struktur: öppningslektion med AI-labb och inokulering (L1), summativ skriftlig examination (L6), och seminarie-examination (L8). Dessa definierar momentets start och två bedömningsändpunkter som resten av momentet ska lära mot.

</domain>

<decisions>
## Implementation Decisions

### Seminariet (L8)
- Smågrupper om 4-5 elever
- Tilldelade perspektiv (t.ex. medieforskare, AI-utvecklare, psykolog) - stärker tredjepersons-framing och skyddar personliga åsikter
- Seminariefrågeställning: analys av ett aktuellt case (specifik konspirationsteori/AI-fenomen) från de tilldelade perspektiven
- Bedömning via inspelning + efterbedömning - seminariet spelas in (ljud) så läraren kan bedöma i efterhand
- Equity-safeguards enligt fas 1: tilldelade roller, tredjepersons-framing, privat uppföljningsplan

### Skrivuppgiften (L6)
- Analys av givet material - eleverna får en eller flera källor (t.ex. konspirationsteori-text + AI-genererat innehåll) och skriver en källkritisk analys
- Omfattning: 500-800 ord
- Hjälpmedel: egna anteckningar + utdelat material, ingen internet
- Rubriken ges först på examinationsdagen (eleverna vet temat men inte exakt rubrik)
- Koppling till E/C/A-kriterierna från fas 1 (skriftlig examination)

### Aktuella exempel (2025-2026)
- Mix av svenska och internationella exempel i jämn fördelning
- Research-agenten söker de mest aktuella och pedagogiskt lämpliga exemplen
- Inga särskilda begränsningar på ämnen - alla är ok med tredjepersons-framing och analytisk distans
- Återkommande case: 2-3 huvudexempel som följer med genom hela momentet (L1-L8), eleverna fördjupas och ser samma case från nya vinklar

### Claude's Discretion
- AI-labben i L1: val av AI-verktyg, uppgiftsdesign, säkerhetsramar, typ av innehåll eleverna skapar (inom ramen för prebunking/inokulering)
- Seminariets specifika perspektiv/roller (inom ramen för tilldelade perspektiv)
- Skrivuppgiftens exakta rubrikformulering
- Val av specifika 2025-2026-exempel (research-fasen tar fram förslag)

</decisions>

<specifics>
## Specific Ideas

- Seminariet ska bygga på ett aktuellt case, inte en abstrakt frågeställning
- Skrivuppgiften ska vara analys av givet material så alla har samma utgångspunkt - jämförbar bedömning
- Exemplen ska vara återkommande (2-3 case genom hela momentet) för att ge fördjupning snarare än bredd

</specifics>

<code_context>
## Existing Code Insights

### Reusable Assets
- `/docx` skill: generera Word-dokument för lektionsplaner
- `/planera-moment` skill: 6-fas lektionsmodell (Rosenshine), exit ticket-slinga
- Fas 1-dokument: momentplan.md (8-lektionsöversikt), bedomningskriterier.md (E/C/A), tredjepersons-framing.md, exit-ticket-mall.md

### Established Patterns
- 6-fas lektionsstruktur: retrieval review -> instruktion -> guidad övning -> självständig övning -> exit ticket -> förhandsvisning
- E/C/A-bedömning med explicit C-nivå (definierad i fas 1)
- Tredjepersons-framing som genomgående norm (definierad i fas 1)

### Integration Points
- Lektionsplaner sparas som .docx i `Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/`
- Ska följa momentplanens struktur för L1, L6 och L8
- Exit ticket-frågor ska matcha exit-ticket-mall.md (blandformat)

</code_context>

<deferred>
## Deferred Ideas

None - discussion stayed within phase scope

</deferred>

---

*Phase: 02-andpunkter-l1-l6-l8*
*Context gathered: 2026-03-12*
