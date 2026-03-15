# Phase 5: Avslutning (L7 + momentoversikt) - Research

**Researched:** 2026-03-15
**Domain:** Lesson plan generation (docx), HTML student overview, CSV export
**Confidence:** HIGH

## Summary

This phase completes the teaching module by delivering three distinct outputs: (1) an updated L7 lesson plan as .docx with proper 6-phase structure and metakognitive reflection, (2) a self-contained HTML student overview page ("momentoversikt") generated via the `/html-momentoversikt` skill, and (3) a CSV export of exit ticket questions from all 7 lessons matching the survey platform's import format.

The existing L7 lesson plan (`generate-lektion-7.js`) already exists and contains solid content with metacognitive reflection, self-assessment matrix, and retrieval review. The main work is to verify it meets phase 5 requirements (6-phase structure with specific prompts, not superficial reflection), ensure consistency with updates from phases 2-4, and generate the momentoversikt HTML and CSV export.

**Primary recommendation:** Treat L7 as a review-and-update task (the generate script exists), focus effort on the momentoversikt HTML generation and the CSV exit ticket export which are new deliverables.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- Momentoversikt: foerberedelseverktyg (framaaatblickande, inte arkiv), bara lektionsnummer (inga datum), per lektion: titel + kort beskrivning, kort bedomningssammanfattning (L6 skriftlig, L8 seminarium), genereras med `/html-momentoversikt` skill
- CSV-export: format matchar `topic,type,text,option1,option2,option3,option4,correctAnswer`, tva fraaagetyper (MULTIPLE_CHOICE, FREE_TEXT), ett topic per lektion, 1-2 fraagor per lektion, UTF-8, falt med komma i dubbla citattecken, fraagor haemtas fraan befintliga lektionsplaner

### Claude's Discretion
- L7 reflektionsdesign: hur den metakognitiva reflektionen struktureras (individuellt/grupp, skriftligt/muntligt, specifika prompts) - inom ramen for kravet "specifika prompts, inte ytlig vad laerde vi oss"
- L7 sjalvbedomning: form och kriterier
- L7 exit ticket-design och retrieval review (baserad paa L6)
- HTML-sidans visuella design och layout
- Val av specifika exit ticket-fraagor per lektion (inom ramen for 1-2 fraagor, blandformat, koppling till laarandemaal)
- L7 presentation (.pptx) om det behoevs

### Deferred Ideas (OUT OF SCOPE)
None - discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| MAT-01 | 7 lektionsplaner som Word-dokument (.docx) med 6-fas struktur | L7 generate script exists (`generate-lektion-7.js`). L1-L6 already complete as .docx. Needs verification that L7 has proper 6-phase structure and reflects phase 2-4 updates. All 7 must be confirmed present and consistent. |
| MAT-03 | Momentoversikt som HTML-sida for eleverna | `/html-momentoversikt` skill provides full generation pipeline. Data sourced from momentplan.md. Locked decision: no dates, lektionsnummer + titel + beskrivning, bedomningssammanfattning. |
</phase_requirements>

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| docx (npm) | latest | Generate L7 .docx lesson plan | Established pattern from all previous phases, `npm install -g docx` |
| Node.js | system | Run generate scripts | All lesson generators use Node.js |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| `/html-momentoversikt` skill | built-in | Generate self-contained HTML student page | For momentoversikt - follows skill's design spec |
| `/docx` skill | built-in | Pattern reference for .docx generation | For L7 lesson plan structure |

### No External Dependencies Needed
The CSV export is pure text generation - no library needed. Write directly with `fs.writeFileSync` or inline in the generate script.

## Architecture Patterns

### Recommended Project Structure
```
Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/
  lektion-7.docx           # Updated L7 (from generate-lektion-7.js)
  momentoversikt.html       # NEW: student overview page
  exit-tickets.csv          # NEW: CSV export for survey platform
  generate-lektion-7.js     # Updated generate script
```

### Pattern 1: Existing Generate Script Pattern
**What:** Each lesson has a `generate-lektion-N.js` that produces a .docx using the `docx` npm library
**When to use:** For L7 - the script already exists and contains the full lesson content
**Key structure from generate-lektion-7.js:**
- Helper functions: `headerCell`, `cell`, `bullet`, `heading2`, `heading3`, `bodyText`, `italicText`, `boldBodyText`, `spacer`
- Document styles: Arial font, Heading1/2/3 with color scheme (1A1A2E, 2C3E50, 34495E)
- Numbering config: bullets and numbers references
- Page size: A4 (11906 x 16838 DXA) with 1440 margins
- Footer: course name and module title
- Content: Titel, Larandemal, Centralt innehall, Forberedelse, Tidsplanering (table), Lararinstruktioner, Elevaktiviteter, Differentiering, Material, Kopplingar, Bilagor

### Pattern 2: 6-Phase Lesson Structure (Rosenshine)
**What:** Every lesson follows: Retrieval review -> Instruktion -> Guidad ovning -> Sjalvstandig ovning -> Exit ticket -> Forhandsvisning
**When to use:** L7 must conform to this. Current L7 has a slightly different phase naming (Uppstart, Instruktion, Bearbetning, Reflektion, Summering, Framtidblick) which maps reasonably but should be verified for alignment.

**L7 current phases vs 6-phase model:**
| Current L7 | 6-phase model | Notes |
|------------|---------------|-------|
| Uppstart (Aterblick) 0-12 min | Retrieval review | Good - uses L6 images as retrieval |
| Instruktion (Momentets resa) 12-15 min | Instruktion | Brief summary - appropriate |
| Bearbetning (Syntesreflektion) 15-30 min | Guidad ovning | Individual reflection with prompts |
| Bearbetning (Par-samtal + helklass) 30-40 min | Sjalvstandig ovning | Pair + whole-class discussion |
| Reflektion (Sjalvvardering) 40-50 min | Exit ticket (formativ) | Self-assessment matrix |
| Summering + Framtidblick 50-65 min | Forhandsvisning | Closure + forward look to L8 |

### Pattern 3: HTML Momentoversikt Generation
**What:** Self-contained HTML file following `/html-momentoversikt` skill spec
**Source data:** momentplan.md lektionsoversikt table
**Key constraints from CONTEXT.md:**
- No dates (bara lektionsnummer)
- Per lektion: titel + kort beskrivning
- Bedomningssammanfattning: L6 skriftlig examination, L8 seminarium
- Forward-looking: "vad kommer harnast?"
- Samhallskunskap color palette: teal, slate, vit (from skill spec)

### Pattern 4: CSV Exit Ticket Export
**What:** CSV file matching survey platform import format
**Format:** `topic,type,text,option1,option2,option3,option4,correctAnswer`
**Source:** Exit ticket questions from exit-ticket-mall.md + individual lesson plans
**Rules:**
- MULTIPLE_CHOICE: 4 options + correctAnswer
- FREE_TEXT: empty options and correctAnswer
- Topic per lesson: e.g., "L1 - Konspirationsteorier & AI-labb"
- 1-2 questions per lesson
- UTF-8, commas in fields quoted

### Anti-Patterns to Avoid
- **Rewriting L7 from scratch:** The generate script already has comprehensive content. Update, don't rebuild.
- **Including teacher-internal info in momentoversikt:** The HTML page is for students - no lesson phases, scaffolding notes, or assessment criteria details.
- **Using dates in momentoversikt:** Locked decision - only lesson numbers for flexibility.
- **Complex exit ticket questions for CSV:** Keep to 1-2 per lesson, matching the minimal exit-ticket-mall design.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| HTML student page | Custom HTML from scratch | `/html-momentoversikt` skill | Skill has design spec, color palettes, responsive layout, Google Sites compatibility |
| .docx generation | Manual XML editing | `docx` npm library via generate script | Established pattern, all 8 lessons use this |
| CSV formatting | Manual string concatenation | Simple fs.writeFileSync with proper quoting | Format is simple enough, but quote fields with commas |

## Common Pitfalls

### Pitfall 1: L7 Inconsistency with Phase 2-4 Updates
**What goes wrong:** L7 references lesson content that changed during phases 2-4 (e.g., L1 is now AI-labb, L4 is formativt seminarium, L5 is skriftlig perspektivanalys)
**Why it happens:** generate-lektion-7.js was written early and may reference outdated lesson descriptions
**How to avoid:** Cross-reference L7's "Momentets resa" summary (lines 276-282 of generate script) against the current momentplan.md lektionsoversikt
**Warning signs:** L7 mentions "kamratgranskning" in L6 (was removed in Phase 4), or wrong lesson titles

### Pitfall 2: Superficial Metacognitive Reflection
**What goes wrong:** L7 reflection becomes "vad larde vi oss" ritual instead of genuine metacognition
**Why it happens:** Easy to write generic reflection prompts
**How to avoid:** The existing L7 already has good prompts ("Hur tanker du annorlunda nu?", "Beskriv en situation dar du redan har anvant det du lart dig"). Verify these remain specific and require concrete examples.
**Warning signs:** Prompts that can be answered with one word or generic statements

### Pitfall 3: CSV Encoding Issues
**What goes wrong:** Swedish characters (aa, ae, oe) break in CSV import
**Why it happens:** File not saved as UTF-8, or BOM issues
**How to avoid:** Write with explicit UTF-8 encoding, test with Swedish characters, add UTF-8 BOM if the platform expects it
**Warning signs:** The survey platform docs don't mention BOM - test without first

### Pitfall 4: Momentoversikt Content Bloat
**What goes wrong:** HTML page includes too much detail (lesson phases, teacher notes, full bedomningskriterier)
**Why it happens:** Rich source data tempts overinclusion
**How to avoid:** Stick to locked decisions: lektionsnummer + titel + kort beskrivning. Bedomningssammanfattning is brief ("L6 ar skriftlig examination, L8 ar seminarium, med vad som bedms i korta drag")

### Pitfall 5: Missing L7 Exit Ticket in CSV
**What goes wrong:** CSV only includes L1-L6, forgetting L7 has its own exit ticket
**Why it happens:** L8 has no exit ticket (summativ), easy to conflate L7 and L8
**How to avoid:** exit-ticket-mall.md explicitly lists L7's exit ticket. Include it. L8 has none.

## Code Examples

### CSV Export Format (verified from generera-fragor.md)
```csv
topic,type,text,option1,option2,option3,option4,correctAnswer
"L1 - Oppning med AI-labb",MULTIPLE_CHOICE,"Vilken av dessa ar en trovardighetsmarkor som AI-genererat innehall ofta anvander?",Grammatiska fel,Akademisk stil och kallhanvisningar,Korta meningar utan sammanhang,Felstavade ord,Akademisk stil och kallhanvisningar
"L1 - Oppning med AI-labb",FREE_TEXT,"Namm tva saker som gjorde AI-genererat innehall overtygande under labben - och en sak som en kallkritisk granskare borde ha reagerat pa.",,,,,
```

### Generate Script Pattern (from existing generate-lektion-7.js)
```javascript
// Key pattern: helper functions + Document config + sections array
const doc = new Document({
  styles: { /* Heading1, Heading2, Heading3 with Arial */ },
  numbering: { config: [{ reference: "bullets", ... }, { reference: "numbers", ... }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, ... } } },
    footers: { default: new Footer({ ... }) },
    children: [ /* content paragraphs, tables, page breaks, bilagor */ ]
  }]
});
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("lektion-7.docx", buffer));
```

### Momentoversikt Data Extraction (from momentplan.md)
```
L1: Oppning med AI-labb och inokulering
  -> Eleverna skapar sjalva overtygande AI-genererat innehall (prebunking/inokulering).
L2: Kallkritiska verktyg
  -> De fyra kallkritiska grundfragorna, SIFT-metoden, lateral reading. Hands-on ovning.
L3: AI och konspirationsteorier
  -> Hur AI forstarker konspirationsteoriers spridning. Integration av kallkritik och konspirationsteori.
L4: Formativt seminarium
  -> Tillampning av verktyg med tilldelade perspektiv.
L5: Skriftlig perspektivanalys
  -> Dress rehearsal for L6. Kortare perspektivanalys med kallhanvisningar.
L6: Summativ skrivuppgift
  -> Summativ examination: skriftlig perspektivanalys.
L7: Metakognitiv reflektion och sjalvbedomning
  -> Reflektera over sitt eget larande, identifiera styrkor och utvecklingsomraden.
L8: Seminarie-examination
  -> Summativ examination: seminarium med tilldelade perspektiv.
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| L7 written early (Phase 1 era) | Updated with Phase 2-4 content | Phase 5 | Must cross-reference lesson descriptions |
| No momentoversikt | HTML-momentoversikt skill available | Skill exists | Use skill pipeline, don't hand-roll |
| Exit tickets in lesson plans only | CSV export for survey platform | Phase 5 | New deliverable - extract from existing plans |

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Manual validation (document generation) |
| Config file | none |
| Quick run command | `node generate-lektion-7.js && python resources/office-scripts/validate.py lektion-7.docx` |
| Full suite command | Validate all .docx + open HTML in browser + verify CSV encoding |

### Phase Requirements -> Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| MAT-01 | All 7 lesson plans exist as .docx with 6-phase structure | smoke | `ls -la lektion-{1,2,3,4,5,6,7}.docx` | 6 of 7 exist, L7 needs regeneration |
| MAT-01 | L7 .docx validates | unit | `python resources/office-scripts/validate.py lektion-7.docx` | Script exists |
| MAT-03 | Momentoversikt HTML is self-contained and renders | smoke | Open `momentoversikt.html` in browser | Will be created |
| MAT-03 | Momentoversikt has all 8 lessons listed | manual-only | Visual inspection | N/A |

### Sampling Rate
- **Per task commit:** Validate .docx after generation, open HTML in browser
- **Per wave merge:** All 7 .docx present, HTML renders, CSV imports correctly
- **Phase gate:** All deliverables present and validated

### Wave 0 Gaps
None - existing infrastructure (docx npm, validate.py, html-momentoversikt skill) covers all needs.

## Open Questions

1. **L7 Content Currency**
   - What we know: generate-lektion-7.js exists with full content
   - What's unclear: Whether the "Momentets resa" summary and lesson descriptions in L7 match the final versions from phases 2-4
   - Recommendation: Diff L7's lesson references against momentplan.md during implementation

2. **CSV BOM Requirement**
   - What we know: Format spec says UTF-8
   - What's unclear: Whether the survey platform needs a UTF-8 BOM for Swedish characters
   - Recommendation: Generate without BOM first, add if import fails

3. **L7 Presentation Need**
   - What we know: Claude's discretion whether L7 needs a .pptx
   - What's unclear: Whether the "Aterblick" opening with 3 images needs a presentation
   - Recommendation: L7 references reusing L1's images - a simple slide deck may help but is not strictly required. Skip unless specifically needed.

## Sources

### Primary (HIGH confidence)
- `generate-lektion-7.js` - Full existing L7 generate script (497 lines)
- `momentplan.md` - Lesson overview table with all 8 lessons
- `exit-ticket-mall.md` - Exit ticket questions for L1-L7
- `Kod/survey-platform/prompts/generera-fragor.md` - CSV import format specification
- `.claude/skills/html-momentoversikt/SKILL.md` - HTML generation skill spec
- `.claude/skills/docx/SKILL.md` - Docx generation patterns and rules

### Secondary (MEDIUM confidence)
- `bedomningskriterier.md` - Assessment criteria for bedomningssammanfattning in HTML

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - all tools and libraries are established in prior phases
- Architecture: HIGH - patterns copied from 6 completed lesson generators
- Pitfalls: HIGH - based on direct inspection of existing code and content

**Research date:** 2026-03-15
**Valid until:** 2026-04-15 (stable - no external dependencies changing)
