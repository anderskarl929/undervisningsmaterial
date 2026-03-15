# Phase 4: Tillämpningslektioner (L4 + L5) - Research

**Researched:** 2026-03-15
**Domain:** Lektionsplanering - seminarium och skriftlig perspektivanalys med scaffolding-fade
**Confidence:** HIGH

## Summary

Fas 4 handlar om att skapa två lektionsplaner (L4 formativt seminarium, L5 skriftlig perspektivanalys) med tillhörande presentationer (.pptx). Befintliga generate-scripts (`generate-lektion-4.js` och `generate-lektion-5.js`) innehåller redan fullständiga lektionsplaner med detaljerat innehåll - bilagor, rollkort, tidsplanering, lärarinstruktioner. Dessa scripts genererar .docx-filer. Inga .pptx-filer finns ännu för L4 eller L5 (bara L2 och L3 har .pptx).

Huvuduppgiften är att (1) revidera de befintliga generate-scripts för att anpassa dem till CONTEXT.md-besluten (gemensamt case L4/L5, scaffolding-fade, L4 vs L8 differentiering), (2) skapa .pptx-presentationer för båda lektionerna, och (3) säkerställa att AI-02 (metakognition) och MAT-02 (presentationer) adresseras.

**Primary recommendation:** Utgå från befintliga generate-scripts som bas, revidera innehållet enligt kontextbesluten, och skapa nya .pptx-filer med pptxgenjs.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
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

### Deferred Ideas (OUT OF SCOPE)
None - discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| AI-02 | Elev kan reflektera över sitt eget tänkande och blinda fläckar (metakognition) | L4 exit ticket och lärpar-återgivning tränar metakognition. L5 exit ticket ("markera en mening där du kunde ha stärkt din argumentation") är explicit metakognitiv. Båda lektionerna har moment där eleverna reflekterar över sitt eget tänkande. |
| MAT-02 | Presentationer som PowerPoint (.pptx) för lektioner med instruktionsmoment | L4 och L5 behöver varsin .pptx. L4 har instruktionsmoment (seminarieintro, spelregler, tredjepersons-framing-påminnelse). L5 har instruktionsmoment (skrivuppgiftens förväntningar, Harvard-påminnelse). |
</phase_requirements>

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| docx (npm) | latest | Generera .docx lektionsplaner | Redan installerat, alla tidigare lektioner använder det |
| pptxgenjs (npm) | latest | Generera .pptx presentationer | Redan installerat globalt, använt för L2 och L3 |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| fs (node built-in) | - | Skriva filer till disk | Alltid, för att spara genererade filer |

**Installation:**
```bash
# Redan installerat i projektets package.json
cd "Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/"
npm install
```

## Architecture Patterns

### Recommended Project Structure
```
Undervisningsmaterial/Samhällskunskap/Källkritik AI och konspirationsteorier/
├── generate-lektion-4.js     # Genererar lektion-4.docx (BEFINTLIG - ska revideras)
├── generate-lektion-5.js     # Genererar lektion-5.docx (BEFINTLIG - ska revideras)
├── lektion-4.docx            # Output (BEFINTLIG - ska regenereras)
├── lektion-4.pptx            # Output (NY - ska skapas)
├── lektion-5.docx            # Output (BEFINTLIG - ska regenereras)
└── lektion-5.pptx            # Output (NY - ska skapas)
```

### Pattern 1: Generate-script pattern (ESTABLISHED)
**What:** Varje lektion har ett `generate-lektion-N.js` som skapar en .docx via docx-biblioteket
**When to use:** Alltid - alla 8 lektioner följer detta mönster
**Example structure:**
```javascript
// Standardmönster från befintliga scripts
const { Document, Packer, Paragraph, TextRun, Table, ... } = require("docx");

// Hjälpfunktioner (identiska i alla scripts)
function headerCell(text, width) { ... }
function cell(text, width) { ... }
function bullet(text) { ... }
function heading2(text) { ... }
function heading3(text) { ... }
function bodyText(text) { ... }
function italicText(text) { ... }
function boldBodyText(label, text) { ... }

// Dokumentstruktur
const doc = new Document({
  styles: { ... },       // Arial, headings
  numbering: { ... },    // bullets, numbers
  sections: [{
    properties: { page: { size: A4, margin: 1440 } },
    footers: { ... },
    children: [
      // TITEL
      // LÄRANDEMÅL (med E/C/A)
      // CENTRALT INNEHÅLL
      // FÖRBEREDELSE
      // TIDSPLANERING (tabell)
      // LÄRARINSTRUKTIONER (per fas)
      // ELEVAKTIVITETER
      // DIFFERENTIERING
      // MATERIAL
      // KOPPLING TILL KUNSKAPSKRAV
      // KOPPLINGAR
      // BILAGOR (PageBreak mellan)
    ]
  }]
});
```

### Pattern 2: PPTX-presentation (ESTABLISHED for L2, L3)
**What:** Separat .pptx-fil skapad med pptxgenjs, innehåller slides för instruktionsmoment
**When to use:** Lektioner med instruktionsmoment som behöver visuellt stöd

### Pattern 3: 6-fas lektionsstruktur (Rosenshine)
**What:** Varje lektion följer: retrieval review -> instruktion -> guidad övning -> självständig övning -> exit ticket -> framåtkoppling
**When to use:** Genomgående i alla lektioner

### Anti-Patterns to Avoid
- **Skapa .docx från scratch utan att utgå från befintliga scripts:** De befintliga generate-lektion-4.js och generate-lektion-5.js innehåller redan fullständiga lektionsplaner. Revidera, bygg inte om.
- **Avvikande dokumentstruktur:** Alla lektionsplaner måste följa identisk rubrikstruktur och formatering.
- **Kamratrespons i L5:** CONTEXT.md specificerar explicit lärarfeedback, inte kamratrespons. Befintliga generate-lektion-5.js har en kamratläsningsfas som behöver tas bort eller ändras.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Word-dokument | Skriv XML manuellt | docx npm-biblioteket | Etablerat mönster, alla scripts använder det |
| Presentationer | Bygg slides manuellt | pptxgenjs | Etablerat mönster från L2/L3 |
| Lektionsinnehåll | Skriv helt nytt innehåll | Revidera befintliga generate-scripts | Scripts har redan genomarbetade bilagor, rollkort, källhänvisningar |

## Common Pitfalls

### Pitfall 1: L5 har kamratrespons i befintligt script men CONTEXT.md säger lärarfeedback
**What goes wrong:** Befintliga generate-lektion-5.js har en kamratläsningsfas (50-65 min). Men kontextbeslutet specificerar "Lärarfeedback mellan L5 och L6: kort formativ feedback på varje elevtext (inte kamratrespons)".
**Why it happens:** Scriptet skrevs före kontextdiskussionen.
**How to avoid:** Ta bort kamratläsningsfasen från L5-scriptet. Ersätt med mer skrivtid eller en kort helklassdel. Lärarfeedback sker EFTER lektionen, inte under.
**Warning signs:** Om L5 fortfarande har "kamratläsning" i tidsplaneringen.

### Pitfall 2: L4 framåtkoppling refererar till L5 som "debatt" - men L5 är skriftlig analys
**What goes wrong:** Befintliga generate-lektion-4.js har en framåtkoppling som säger "Nästa gång höjer vi insatsen: ni får tilldelade positioner som ni kanske INTE håller med om - och ska debattera dem." Men L5 är en skriftlig perspektivanalys, inte en debatt.
**Why it happens:** Scriptet skrevs före momentplanens slutliga design.
**How to avoid:** Ändra framåtkopplingen i L4 till att referera till den skriftliga perspektivanalysen.
**Warning signs:** Om L4 framåtkoppling nämner "debatt" istället för "skriftlig perspektivanalys".

### Pitfall 3: L4 och L5 använder inte samma case/tema
**What goes wrong:** CONTEXT.md specificerar "L4 och L5 använder samma case/tema - den muntliga diskussionen i L4 förbereder direkt det skriftliga arbetet i L5". Befintliga scripts måste verifiera tematisk koherens.
**Why it happens:** Scripts kan ha skrivits oberoende.
**How to avoid:** Säkerställ att L4-seminariets fråga och L5-skrivuppgiftens fråga använder samma tema/case. Det befintliga L4 använder "Informationskriget" (AI, desinformation, demokrati) och L5 använder "Vem bär ansvaret för att stoppa AI-driven desinformation?" - detta är redan samma tema. Behåll tematisk koherens.
**Warning signs:** Om L4 och L5 har helt olika ämnesområden.

### Pitfall 4: L4 vs L8 differentiering saknas
**What goes wrong:** CONTEXT.md flaggar att L4 (formativt seminarium) måste differentieras tydligt från L8 (summativt seminarium).
**Why it happens:** Båda är seminarier med tilldelade perspektiv.
**How to avoid:** L8 (från fas 2) har: 5 tilldelade perspektiv, inspelning, smågrupper om 4-5. L4 bör därför ha: färre perspektiv (3-4), helklassformat, inget inspelningsmoment, lägre insats, formativt syfte med lärpar-återgivning som metakognitiv komponent.

### Pitfall 5: Scaffolding helt borttagen istället för tillgänglig på begäran
**What goes wrong:** CONTEXT.md och success criteria specificerar "Scaffolding tillgänglig på begäran men inte utdelad som standard". Det befintliga L4-scriptet säger redan "Inga meningsstartare på tavlan under seminariet" - men scaffolding bör finnas i en bilaga eller lärarinstruktionen.
**Why it happens:** Missförstånd av "explicit fade" - det betyder inte borttagning utan tillgänglighet.
**How to avoid:** Ha scaffolding-material (meningsstartare, begreppslistor) i en bilaga eller lärarens instruktion, men dela inte ut det proaktivt. Instruera läraren att ge stöd vid behov, inte som standard.

### Pitfall 6: Exit ticket-mallen i exit-ticket-mall.md har redan definierade frågor
**What goes wrong:** Man designar nya exit tickets utan att kolla exit-ticket-mall.md.
**Why it happens:** Mallen skapades i fas 1 och kan missas.
**How to avoid:** L4 exit ticket finns redan: "Beskriv ett ögonblick under seminariet då du ändrade eller fördjupade din analys..." L5 exit ticket finns redan: "Läs igenom din perspektivanalys. Markera en mening..."

## Code Examples

### Befintlig L4 generate-script (NYCKELSTRUKTUR)
```javascript
// Befintlig generate-lektion-4.js - 100 min lektion
// Titel: "Lektion 4: Informationskriget"
// Undertitel: "Seminarium - AI, desinformation och demokrati"
// Struktur:
//   0-10: Retrieval practice + ramasättning
//   10-25: Gruppförberedelse av positioner
//   25-70: Seminarium (öppningsanföranden + fri diskussion + syntes)
//   70-80: Lärpar-återgivning
//   80-95: Summering + exit ticket
//   95-100: Framåtkoppling
// Bilagor: Seminarieunderlag, Bedömningsmatris, Positionskort (3 positioner)
```

### Befintlig L5 generate-script (NYCKELSTRUKTUR)
```javascript
// Befintlig generate-lektion-5.js - 85 min lektion
// Titel: "Lektion 5: Perspektiven krockar"
// Undertitel: "Skriftlig perspektivanalys"
// Fråga: "Vem bär ansvaret för att stoppa AI-driven desinformation?"
// Struktur:
//   0-8: Retrieval practice
//   8-15: Instruktion - presentera skrivuppgiften
//   15-50: Skriftlig perspektivanalys (35 min)
//   50-65: Kamratläsning (BEHÖVER ÄNDRAS - ska bli mer skrivtid)
//   65-77: Helklassdiskussion
//   77-82: Exit ticket
//   82-85: Framåtkoppling
// Bilagor: Rollkort (3 positioner med källor), Bedömningsstöd, Exit ticket, Skrivuppgift, Källmaterial
```

### Ändringar som behövs i L4
```
1. Framåtkoppling: Ändra referens från "debatt" till "skriftlig perspektivanalys"
2. Scaffolding: Lägg till bilaga med "scaffolding på begäran" (meningsstartare, begreppslistor)
3. Verifiera att det har 3-4 perspektiv (inte 5 som L8)
4. Exit ticket: Verifiera mot exit-ticket-mall.md
```

### Ändringar som behövs i L5
```
1. Ta bort kamratläsningsfasen (50-65 min)
2. Ersätt med mer skrivtid eller annan aktivitet
3. Lägg till lärarinstruktion om formativ feedback EFTER lektionen
4. Verifiera att temat matchar L4 (redan ok: samma ansvarsfråga)
5. Exit ticket: Verifiera mot exit-ticket-mall.md
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| L5 med kamratrespons | L5 med lärarfeedback (efter lektion) | CONTEXT.md 2026-03-15 | Kamratläsningsfasen i scriptet måste tas bort |
| L4 framåtkoppling till "debatt" | L4 framåtkoppling till "skriftlig perspektivanalys" | Momentplanens slutliga design | Framåtkopplingen måste ändras |

## L4 vs L8 Differentiering (Claude's Discretion - Rekommendation)

| Aspekt | L4 (Formativt) | L8 (Summativt) |
|--------|-----------------|-----------------|
| Syfte | Formativt - träna och identifiera gap | Summativt - bedömning |
| Antal perspektiv | 3 positioner | 5 perspektiv |
| Format | Helklass med alla grupper | Smågrupper om 4-5 |
| Inspelning | Nej | Ja |
| Scaffolding | Tillgänglig på begäran | Ingen |
| Lärpar-återgivning | Ja (metakognitiv komponent) | Nej |
| Förberedelse | Positionskort + seminarieunderlag | Tilldelade perspektiv |
| Tidsram | 100 min | Examinationstid |

**Rekommendation:** L4 bör behålla sin befintliga design med 3 positioner i helklassformat. Det som differentierar mest är: formativt syfte (inga betyg), lärpar-återgivning som metakognitiv tillägg, och scaffolding tillgänglig på begäran.

## Scaffolding-fade Design (Claude's Discretion - Rekommendation)

### Progressionen L1-L3 -> L4-L5

| Lektion | Scaffolding-nivå | Konkret |
|---------|-----------------|---------|
| L1-L3 | Full scaffolding | Meningsstartare på tavlan, begreppslistor utdelade, stödstrukturer som standard |
| L4 | Scaffolding på begäran | Inga meningsstartare på tavlan. Positionskort ger viss struktur. Läraren ger riktad hjälp till elever som behöver det. Scaffolding-bilaga finns för läraren att använda vid behov. |
| L5 | Scaffolding på begäran | Rollkort som skrivstöd (redan i befintligt script). Harvard-guide tillgänglig. Meningsstartare på tavlan om läraren bedömer att det behövs, annars inte. |

**Rekommendation:** Skapa en "Bilaga: Scaffolding på begäran" i båda lektionsplanerna som innehåller meningsstartare och stödfrågor. Instruera läraren att använda dessa vid behov, inte dela ut dem proaktivt. Befintliga scripts har redan detta delvis implementerat.

## PPTX Design (Claude's Discretion - Rekommendation)

### L4 Presentation (Seminarium)
Slides behövs för:
1. Titel + lektionens mål
2. Retrieval practice-fråga
3. Seminariestruktur och spelregler
4. Tredjepersons-framing-påminnelse
5. Seminariefrågor (referens för läraren)
6. Exit ticket

### L5 Presentation (Skriftlig perspektivanalys)
Slides behövs för:
1. Titel + lektionens mål
2. Retrieval practice-fråga
3. Skrivuppgiftens instruktion och förväntningar
4. Harvard-påminnelse
5. Textstruktur (stöd)
6. Exit ticket
7. Framåtkoppling till L6

**Designprinciper:** Följ mönstret från L2/L3 .pptx. Använd pptxgenjs. Mörk bakgrund för titel/avslut, ljus för innehåll. Tydlig typografi.

## Open Questions

1. **L4 lektionslängd: 100 min vs 85 min**
   - What we know: Befintligt L4-script har 100 min. Momentplanen anger 80 min per lektion.
   - What's unclear: Om 100 min är korrekt eller om det ska vara 85 min.
   - Recommendation: Behåll 100 min för seminariet - det behöver tid. Momentplanen anger "80 minuter" men befintliga scripts har varierat (85 min för L5).

2. **L5 utan kamratrespons - hur fylla tiden?**
   - What we know: Kamratläsning tog 15 min (50-65 min). Den ska bort.
   - What's unclear: Vad som ersätter den.
   - Recommendation: Ge mer skrivtid (förläng från 35 min till 45-50 min) och/eller lägg till kort individuell revision innan helklassdiskussion.

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Node.js script execution |
| Config file | package.json i projektmappen |
| Quick run command | `node generate-lektion-4.js && node generate-lektion-5.js` |
| Full suite command | `node generate-lektion-4.js && node generate-lektion-5.js` |

### Phase Requirements -> Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| AI-02 | Metakognitiv reflektion i exit tickets och lärpar-återgivning | manual-only | Granska .docx-innehåll efter generering | N/A |
| MAT-02 | Presentationer finns som .pptx | smoke | `ls lektion-4.pptx lektion-5.pptx` | Nej - Wave 0 |

### Sampling Rate
- **Per task commit:** `node generate-lektion-4.js` eller `node generate-lektion-5.js`
- **Per wave merge:** Kör alla generate-scripts + verifiera att .docx och .pptx skapas
- **Phase gate:** Alla filer genereras utan fel, innehåll verifierat

### Wave 0 Gaps
- [ ] `lektion-4.pptx` - ny fil som ska skapas
- [ ] `lektion-5.pptx` - ny fil som ska skapas
- [ ] Revidering av `generate-lektion-4.js` (framåtkoppling, scaffolding-bilaga)
- [ ] Revidering av `generate-lektion-5.js` (ta bort kamratrespons, lägga till lärarfeedback-instruktion)

## Sources

### Primary (HIGH confidence)
- Befintliga generate-lektion-4.js och generate-lektion-5.js - fullständigt analyserade
- momentplan.md - lektionsöversikt och progressionslogik
- bedomningskriterier.md - E/C/A-kriterier och equity-safeguards
- tredjepersons-framing.md - genomgående norm
- exit-ticket-mall.md - frågor per lektion

### Secondary (MEDIUM confidence)
- CONTEXT.md (04-CONTEXT.md) - användarbeslut
- docx SKILL.md och pptx SKILL.md - verktygsreferens

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - samma verktyg som fas 2-3, redan etablerat
- Architecture: HIGH - identiskt mönster, befintliga scripts finns
- Pitfalls: HIGH - baserat på direkt jämförelse mellan CONTEXT.md och befintliga scripts
- Content changes: HIGH - specifika ändringar identifierade genom kodanalys

**Research date:** 2026-03-15
**Valid until:** 2026-04-15 (stabil domän, inga externa beroenden)
