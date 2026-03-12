# Phase 1: Strukturell grund - Research

**Researched:** 2026-03-12
**Domain:** Pedagogisk planering - momentstruktur, bedomningskriterier, formativ bedomning
**Confidence:** HIGH

## Summary

Fas 1 handlar om att skapa det strukturella ramverket for ett 8-lektionsmoment om kallkritik, konspirationsteorier och AI. Inga lektionsplaner ska skrivas - bara de fyra dokumenten som alla framtida faser beror pa: momentplan, E/C/A-bedomningskriterier, exit ticket-mall och tredjepersons-framing-norm.

Projektet har redan ett valetablerat pedagogiskt ramverk i `/planera-moment`-skillen med 6-fas lektionsmodell (Rosenshine), Gy11-progressionsord och exit ticket-slinga. En befintlig momentplan.md finns men ska INTE anvandas som utgangspunkt - ny plan byggs fran ROADMAP.md och REQUIREMENTS.md. Det befintliga dokumentet ar dock vardefullt som formatreferens.

**Primary recommendation:** Skapa fyra sammanhangande dokument i `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/`: (1) ny momentplan.md, (2) bedomningskriterier.md, (3) exit-ticket-mall.md, (4) tredjepersons-framing.md. Alla fyra maste vara inbordes koherenta och folja backward design.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- Exit ticket-design: Blandformat (oppna + strukturerade fragor), minimal mall (fraga + larandemal), explicit retrieval review-koppling i varje lektionsplan, CSV-export som sista leverans (inte i fas 1)
- Momentplan: Byggs fran grunden baserat pa ROADMAP.md (8 lektioner: 7 undervisning + 1 examination), befintlig momentplan.md ignoreras helt
- Momentplanen ska beskriva alla 8 lektioner med mal, innehall och progression

### Claude's Discretion
- E/C/A-bedomningskriterier: formulering och detaljeringsniva for skriftlig och muntlig examination, med explicit C-niva
- Tredjepersons-framing: definition av normen, exempelformuleringar, och om det galler alla lektioner eller framst konspirationsteorimomenten
- Momentplanens interna progressionslogik
- Exit ticket-fragebankens specifika fragor per lektion

### Deferred Ideas (OUT OF SCOPE)
- CSV-export av exit ticket-fragor till lararens egen applikation - levereras efter alla lektionsplaner ar klara (troligen fas 5 eller separat steg)
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| BED-01 | Exit ticket varje lektion som informerar nasta lektions retrieval review | Exit ticket-mall med blandformat, minimal struktur (fraga + larandemal), kopplad till retrieval review-slinga fran pedagogik-ramverket |
| BED-02 | Seminarium som examination med equity-safeguards for konspirationskansligt amne | E/C/A-bedomningskriterier for muntlig examination med explicit C-niva; seminarie-equity-safeguards dokumenterade i bedomningskriterierna |
| BED-03 | Tredjepersons-framing genomgaende ("varfor kan nagon finna detta overtygande?") | Tredjepersons-framing dokumenterad som norm med konkreta exempelformuleringar |
</phase_requirements>

## Standard Stack

### Core
| Komponent | Format | Syfte | Varfor |
|-----------|--------|-------|--------|
| Momentplan | .md (Markdown) | Oversiktsplanering av alla 8 lektioner | Etablerat format i projektet, las av bade manniska och agent |
| Bedomningskriterier | .md | E/C/A-kriterier for skriftlig och muntlig examination | Behover vara sökbart och referensbart fran lektionsplaner |
| Exit ticket-mall | .md | Mall for formativa fragor per lektion | Minimal struktur som aterkommande leveranser fyller |
| Tredjepersons-framing | .md | Normdokument med exempelformuleringar | Referensdokument for alla framtida lektionsplaner |

### Supporting
| Komponent | Syfte | Nar |
|-----------|-------|-----|
| Gy11-progressionsord | E/C/A-formulering | Nar bedomningskriterier skrivs |
| pedagogik-ramverk.md | 6-fas struktur, exit ticket-slinga | Nar momentplanens progression definieras |
| lektionsplanering.md | Mall och principer for lektionsplaner | Formatreferens for momentplanens lektionsbeskrivningar |

### Outputkatalog
Alla filer sparas i: `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/`

## Architecture Patterns

### Rekommenderad filstruktur for fas 1
```
Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/
  momentplan.md          # NY - ersatter befintlig (8 lektioner, mal, innehall, progression)
  bedomningskriterier.md # NY - E/C/A for skriftlig + muntlig examination
  exit-ticket-mall.md    # NY - minimal mall + fragor per lektion
  tredjepersons-framing.md # NY - normdokument med exempelformuleringar
```

### Pattern 1: Backward Design (Wiggins & McTighe)
**What:** Bedomning definieras fore instruktion. Fas 1 skapar bedomningskriterierna som alla lektioner sedan bygger mot.
**When to use:** Alltid vid momentplanering.
**Implikation for fas 1:** Bedomningskriterierna (BED-02) maste vara klara FORE momentplanens progression detaljeras, eftersom progressionen ska leda mot bedomningsandpunkterna.

### Pattern 2: Exit ticket-slinga (fran pedagogik-ramverket)
**What:** Varje lektions exit ticket informerar nasta lektions retrieval review (fas 1 i 6-fas modellen).
**Implikation for fas 1:** Exit ticket-mallen maste ha tva falt: (1) fragan som mater forstaelse, (2) larandemalet den mater. Retrieval review-kopplingen specificeras i varje lektionsplan (fas 2-5), inte i mallen.

### Pattern 3: E/C/A-progressionsmodell (Gy11)
**What:** Kvalitetsnivaaer pa elevernas analys och resonemang, INTE svarighetsnivaaer pa innehallet.
**Implikation for fas 1:** C-nivan maste vara explicit - den saknas ofta i lararproducerade kriterier. Anvand progressionsorden fran gy11-struktur.md: "valgrundade resonemang", "med visst kritiskt perspektiv", "utforligt".

### Pattern 4: Momentplan med 8 lektioner (ROADMAP-driven)
**What:** ROADMAP.md definierar 8 lektioner: 7 undervisning + 1 examination (L8 = seminarium).
**Lektionsfordelning fran ROADMAP:**
- L1: Oppning med AI-labb/inokulering (Fas 2)
- L2: Kallkritiska verktyg (Fas 3)
- L3: AI + konspirationsteorier integration (Fas 3)
- L4: Formativt seminarium (Fas 4)
- L5: Skriftlig perspektivanalys (Fas 4)
- L6: Summativ skrivuppgift (Fas 2)
- L7: Reflektion och sjalvbedomning (Fas 5)
- L8: Seminarie-examination (Fas 2)

**Implikation for fas 1:** Momentplanen ska beskriva alla 8 med mal, innehall och progression. Detaljerade lektionsplaner skrivs i senare faser.

### Anti-Patterns to Avoid
- **Kopiera befintlig momentplan:** CONTEXT.md sager explicit att befintlig momentplan.md ignoreras. Den har 7 lektioner, det nya formatet har 8.
- **Overdetaljerade lektionsbeskrivningar i momentplanen:** Fas 1 skapar ramverket. Detaljerade lektionsplaner (6-fas struktur, tidsplanering) skrivs i fas 2-5.
- **E/C/A utan explicit C-niva:** Vanlig svaghet - C-nivan hamnar "mellan" E och A istallet for att vara tydligt formulerad.
- **Exit ticket som efterhandskonstruktion:** Mallen maste designas som en koppling mellan formativ och summativ bedomning, inte som ett losryckt avslutningsmoment.

## Don't Hand-Roll

| Problem | Bygg inte | Anvand istallet | Varfor |
|---------|-----------|-----------------|--------|
| E/C/A-formuleringar | Egna progressionsord | Gy11:s officiella progressionsord fran `gy11-struktur.md` | Juridiskt bindande dokument, lararen maste kunna koppla till styrdokument |
| Lektionsstruktur | Egen modell | 6-fas modellen fran `pedagogik-ramverk.md` | Redan etablerad i projektet, evidensbaserad |
| Exit ticket-format | Komplexa formuler | Minimal mall: fraga + larandemal (beslut i CONTEXT.md) | Anvandaren har beslutat om minimal design |

## Common Pitfalls

### Pitfall 1: Otydlig C-niva
**What goes wrong:** C-nivan beskrivs vagt eller som "mittemellan E och A" istallet for med egna formuleringar.
**Why it happens:** Gy11:s kunskapskrav ar ibland knappmandiga pa C-nivan.
**How to avoid:** Anvand alltid tre separata, fullstandiga meningar for E, C och A. C ska ha egna progressionsord: "valgrundade resonemang", "utforligt", "med visst kritiskt perspektiv".
**Warning signs:** C-formuleringen ar kortare an E- eller A-formuleringen.

### Pitfall 2: Momentplan utan progression
**What goes wrong:** Lektionerna beskrivs som isolerade enheter utan tydlig rorelse fran grundlaggande till avancerat.
**Why it happens:** Fokus pa innehall per lektion istallet for analytisk utveckling genom momentet.
**How to avoid:** Beskriv explicit i momentplanen: (1) var eleverna borjar, (2) hur analytisk fordjupning sker, (3) vad andpunkterna kraver. Anvand progressionslogiken fran ROADMAP.md (verktygsfas -> tillampningsfas -> examinationsfas).

### Pitfall 3: Tredjepersons-framing som ytlig instruktion
**What goes wrong:** "Anvand tredje person" utan att forklara varfor och hur det forhandrar kognitiv forsvarshallning.
**Why it happens:** Normen ar rattfram att beskriva men svar att operationalisera i klassrummet.
**How to avoid:** Dokumentet maste innehalla: (1) varfor tredjepersons-framing ar viktigt vid konspirationsteorier (minskar personligt hot, oppnar analytiskt utrymme), (2) konkreta exempelformuleringar ("varfor kan nagon finna detta overtygande?" istallet for "tror du pa detta?"), (3) nar det ar sarskilt viktigt (konspirationsteorilektioner) vs genomgaende.

### Pitfall 4: Exit ticket utan retrieval review-koppling
**What goes wrong:** Exit ticket-fragor samlas in men anvands aldrig for att styra nasta lektion.
**Why it happens:** Kopplingen ar implicit, inte explicit i dokumenten.
**How to avoid:** Mallen ska ha ett falt som explicit fragar: "Hur informerar denna exit ticket nasta lektions retrieval review?" Momentplanen ska ha en kolumn/sektion som visar slinga per lektion.

### Pitfall 5: Examination utan equity-safeguards
**What goes wrong:** Seminariet om konspirationsteorier riskerar att elever utpekas for personliga asikter.
**Why it happens:** Standard seminariestruktur antar neutrala amnen.
**How to avoid:** BED-02 kraver explicit: tilldelade roller/perspektiv (inte personliga asikter), tredjepersons-framing, privat uppfoljningsplan vid behov.

## Code Examples

### Momentplan-format (ny, ROADMAP-driven)
```markdown
# Momentplan: Kallkritik - Konspirationsteorier och AI

## Grundinformation
- **Amne/Kurs:** Samhallskunskap 3 / Internationella relationer
- **Centralt innehall:** [fran amnesplanen]
- **Antal lektioner:** 8 (7 undervisning + 1 examination)
- **Tema:** Kallkritik, AI-genererat innehall och konspirationsteorier

## Larandemal
### 1. [Mal med E/C/A-progression]
...

## Lektionsoversikt
| Lektion | Titel | Huvudinnehall | Larandemal | Fas i projektet |
|---------|-------|---------------|------------|-----------------|
| L1 | ... | ... | ... | Fas 2 |
| ... | ... | ... | ... | ... |

## Progressionslogik
[Explicit beskrivning av hur analytiskt djup okar genom momentet]

## Exit ticket-slinga
[Oversikt av hur exit tickets kopplar till retrieval review]
```

### E/C/A-bedomningskriterier-format
```markdown
# Bedomningskriterier: Kallkritik - Konspirationsteorier och AI

## Skriftlig examination (L6)
### E-niva
Eleven gor en **oversiktlig** analys av [innehall] och for **enkla** resonemang...

### C-niva
Eleven gor en **utforlig** analys av [innehall] och for **valgrundade** resonemang
med **visst kritiskt perspektiv**...

### A-niva
Eleven gor en **utforlig och nyanserad** analys av [innehall] och for
**valgrundade och nyanserade** resonemang **ur flera perspektiv**...

## Muntlig examination / Seminarium (L8)
### E-niva
...
### C-niva
...
### A-niva
...

## Equity-safeguards for seminarium
- Tilldelade roller/perspektiv
- Tredjepersons-framing obligatorisk
- Privat uppfoljningsplan vid behov
```

### Exit ticket-mall-format
```markdown
# Exit ticket-mall

## Format
Varje exit ticket bestar av:
1. **Fraga** - oppen reflektion ELLER strukturerad fraga (blandformat per CONTEXT.md)
2. **Larandemal** - vilket av momentets larandemal fragan mater

## Mall per lektion
| Lektion | Fraga | Typ (oppen/strukturerad) | Larandemal | Retrieval review nasta lektion |
|---------|-------|--------------------------|------------|-------------------------------|
| L1 | [specificeras vid genomforande] | ... | ... | -> L2 retrieval review |
| ... | ... | ... | ... | ... |
```

### Tredjepersons-framing-format
```markdown
# Tredjepersons-framing: Norm for momentet

## Varfor tredjepersons-framing?
[Pedagogisk motivering: minskar personligt hot, oppnar analytiskt utrymme]

## Normen
[Definition: analysera varfor nagon KAN finna X overtygande, inte om DU tror pa X]

## Exempelformuleringar
- Istallet for: "Tror du pa denna konspirationsteori?"
- Anvand: "Varfor kan nagon finna denna forklaring overtygande?"
- Istallet for: "Ar detta sant?"
- Anvand: "Vilka behov fyller denna forklaring for den som tror pa den?"

## Tillampning
[Nar det galler: alla lektioner eller framst konspirationsteorimomenten - Claude's discretion]
```

## State of the Art

| Aspekt | Projektets angreppssatt | Stod |
|--------|-------------------------|------|
| Backward design | Bedomning fore instruktion | Wiggins & McTighe, etablerat i projektet |
| 6-fas lektionsmodell | Rosenshine-baserad | pedagogik-ramverk.md |
| Exit ticket-slinga | Formativ -> nasta retrieval review | lektionsplanering.md |
| Tredjepersons-framing | Kognitiv sakerhet vid konspirationsteorier | Forskning om inokulering/prebunking |
| E/C/A-progression | Gy11 progressionsord | gy11-struktur.md |

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Manuell granskning (pedagogiskt innehall) |
| Config file | Ingen - dokumentgranskningsbaserad validering |
| Quick run command | `cat [fil] \| head -50` (snabbkoll av format) |
| Full suite command | Manuell granskning mot success criteria |

### Phase Requirements -> Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| BED-01 | Exit ticket-mall med fraga+larandemal, koppling till retrieval review | manual-only | Granska att exit-ticket-mall.md har alla 8 lektioner med fraga, larandemal och retrieval review-koppling | N/A Wave 0 |
| BED-02 | E/C/A-kriterier for seminarium med equity-safeguards | manual-only | Granska att bedomningskriterier.md har explicit E/C/A for muntlig examination + equity-sektion | N/A Wave 0 |
| BED-03 | Tredjepersons-framing dokumenterad med exempelformuleringar | manual-only | Granska att tredjepersons-framing.md har definition + minst 4 exempelformuleringar | N/A Wave 0 |

### Sampling Rate
- **Per task commit:** Manuell granskning av skapade .md-filer mot success criteria
- **Per wave merge:** Alla fyra dokument koherenta med varandra
- **Phase gate:** Alla success criteria verifierade fore fas 2

### Wave 0 Gaps
None - fas 1 producerar Markdown-dokument, ingen testkod behovs.

## Open Questions

1. **Lektionstider for det nya 8-lektionsformatet**
   - What we know: Befintlig momentplan har 7 lektioner med varierande tider (65-100 min). ROADMAP har 8 lektioner.
   - What's unclear: Vilka tider gar L8 (examination) ha? Andras nagra befintliga tider?
   - Recommendation: Anvand befintliga tider fran momentplan.md som referens for L1-L7, tilldela L8 en rimlig tid (80-100 min for seminarium). Lagg som Claude's discretion.

2. **Tredjepersons-framing: alla lektioner eller enbart konspirationsteorier?**
   - What we know: CONTEXT.md listar detta som Claude's discretion.
   - Recommendation: Genomgaende norm med forstarkta paminelser i konspirationsteorilektioner (L2, L3, L4, L8). Motivering: principen ar lika relevant for AI-genererat innehall.

3. **Bedomningskriterier: separata per examination eller integrerade?**
   - What we know: ROADMAP definierar tva andpunkter: L6 (skriftlig) och L8 (muntlig seminarium).
   - Recommendation: Ett dokument med tva sektioner (skriftlig + muntlig), vardera med fullstandig E/C/A.

## Sources

### Primary (HIGH confidence)
- `.claude/skills/planera-moment/pedagogik-ramverk.md` - 6-fas modell, E/C/A, backward design
- `.claude/skills/planera-moment/references/gy11-struktur.md` - Progressionsord, kunskapskravs-mall
- `.claude/skills/planera-moment/references/lektionsplanering.md` - Exit ticket-slinga, lektionsmall
- `.planning/ROADMAP.md` - 8-lektionsstruktur, fasindelning, bedomningsandpunkter
- `.planning/phases/01-strukturell-grund/01-CONTEXT.md` - Anvandarbeslut och begransningar

### Secondary (MEDIUM confidence)
- Befintlig `momentplan.md` - Formatreferens (inte innehallsreferens per CONTEXT.md)

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - Alla verktyg och format ar etablerade i projektet
- Architecture: HIGH - Backward design, 6-fas modell och Gy11-progressionsord ar val dokumenterade
- Pitfalls: HIGH - Baserat pa kanda svagheter i AI-genererad pedagogisk planering (dokumenterat i SKILL.md)

**Research date:** 2026-03-12
**Valid until:** 2026-04-12 (stabil doman, inga snabba forandringar)
