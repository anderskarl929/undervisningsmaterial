# Phase 2: Andpunkter (L1 + L6 + L8) - Research

**Researched:** 2026-03-12
**Domain:** Lektionsplanering - inokulering/prebunking, summativ examination, seminariestruktur med equity-safeguards
**Confidence:** HIGH

## Summary

Fas 2 ska producera tre detaljerade lektionsplaner som .docx: en oppningslektion med AI-labb och inokulering (L1), en summativ skrivuppgift (L6) och en seminarie-examination (L8). Dessa tre lektioner definierar momentets start och tva bedomningsandpunkter. Alla tre beror pa fas 1:s strukturella ramverk (momentplan, bedomningskriterier, exit ticket-mall, tredjepersons-framing).

Inokulering/prebunking ar valetablerat i forskning - elever som sjalva skapar overtygandeinnehall bygger "mentala antikroppar" mot desinformation. Bad News-spelet har testats pa 516 svenska gymnasieelever med positiva resultat pa kort sikt. For L1 ska eleverna sjalva skapa AI-genererat innehall i en kontrollerad labb-miljo (inte bara spela ett spel), vilket ar en starkare form av aktiv inokulering. Aktuella svenska exempel 2025-2026 inkluderar AI-genererade TikTok-konton som spred hogerpropaganda, deepfake-hotet infor valet 2026, och att 74% av nyskapat natiinnehall 2025 genereras av AI/botar.

**Primary recommendation:** Skapa tre Node.js-script (generate-lektion-1.js, generate-lektion-6.js, generate-lektion-8.js) som genererar .docx-filer med docx-js, foljande samma monster som befintliga generate-lektion-*.js-filer. Varje lektionsplan foljer 6-fas strukturen fran pedagogik-ramverket med 80 minuters tidsbudget.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- Seminariet (L8): Smagrupper om 4-5 elever, tilldelade perspektiv (t.ex. medieforskare, AI-utvecklare, psykolog), seminariefragestallning baserad pa aktuellt case, bedomning via inspelning + efterbedomning, equity-safeguards enligt fas 1
- Skrivuppgiften (L6): Analys av givet material (konspirationsteori-text + AI-genererat innehall), 500-800 ord, hjalpmedel: egna anteckningar + utdelat material (ingen internet), rubriken ges forst pa examinationsdagen
- Aktuella exempel (2025-2026): Mix av svenska och internationella, 2-3 huvudexempel som foljer med genom hela momentet (L1-L8), inga amnesbegransningar med tredjepersons-framing
- Seminariet ska bygga pa ett aktuellt case, inte en abstrakt fragestallning
- Skrivuppgiften ska vara analys av givet material for jamforbar bedomning
- Exemplen ska vara aterkommande (2-3 case genom hela momentet)

### Claude's Discretion
- AI-labben i L1: val av AI-verktyg, uppgiftsdesign, sakerhetsramar, typ av innehall eleverna skapar (inom ramen for prebunking/inokulering)
- Seminariets specifika perspektiv/roller (inom ramen for tilldelade perspektiv)
- Skrivuppgiftens exakta rubrikformulering
- Val av specifika 2025-2026-exempel (research-fasen tar fram forslag)

### Deferred Ideas (OUT OF SCOPE)
None - discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| KONSP-02 | Elev har genomgatt inokulering/prebunking genom AI-labb i kontrollerad miljo | L1 AI-labb dar elever sjalva skapar overtygandeinnehall; forskningsstod fran inokulationsteorin; Bad News-spelet testat pa svenska gymnasieelever |
| KONSP-03 | Momentet anvander aktuella svenska och internationella exempel (2025-2026) | Research identifierar 2-3 aterkommande case: AI-TikTok-konton i Sverige, deepfakes infor valet 2026, AI-genererat innehall pa natet (74% 2025) |
| BED-02 | Seminarium som examination med equity-safeguards for konspirationskansligt amne | L8 seminariestruktur med tilldelade perspektiv, tredjepersons-framing, privat uppfoljningsplan; bedomning via inspelning; E/C/A-kriterier fran fas 1 |
</phase_requirements>

## Standard Stack

### Core
| Komponent | Format | Syfte | Varfor |
|-----------|--------|-------|--------|
| Lektionsplaner | .docx (via docx-js) | Detaljerade lektionsplaner for L1, L6, L8 | Lasbara i Word, Google Docs; etablerat format i projektet |
| Generate-scripts | .js (Node.js) | Genererar .docx-filer programmatiskt | Befintligt monster (generate-lektion-2.js etc.) - ateranvandbart |
| docx-js | npm-paket | Skapar Word-dokument | Redan installerat, beprovat i projektet |

### Supporting
| Komponent | Syfte | Nar |
|-----------|-------|-----|
| momentplan.md | Lektionsinnehall och progressionslogik | Referens for varje lektionsplan |
| bedomningskriterier.md | E/C/A-kriterier for L6 och L8 | Kopieras in/refereras i lektionsplanerna |
| exit-ticket-mall.md | Exit ticket-fragor for L1, L6 | Kopieras in i lektionsplanerna |
| tredjepersons-framing.md | Norm och exempelformuleringar | Refereras sarskilt i L1 (intro) och L8 (seminarium) |
| validate.py | Validerar .docx-filer | Kor efter varje genererad fil |

### Outputkatalog
Alla filer sparas i: `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/`

## Architecture Patterns

### Rekommenderad filstruktur for fas 2
```
Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/
  generate-lektion-1.js   # NY - genererar L1 .docx
  lektion-1.docx          # NY - oppningslektion med AI-labb
  generate-lektion-6.js   # NY - genererar L6 .docx
  lektion-6.docx          # NY - summativ skrivuppgift
  generate-lektion-8.js   # NY - genererar L8 .docx
  lektion-8.docx          # NY - seminarie-examination
```

### Pattern 1: Befintligt generate-script-monster
**What:** Varje lektionsplan genereras av ett fristaende Node.js-script som anvander docx-js. Scriptet definierar hjalpfunktioner (headerCell, cell, bullet, heading2, heading3) och bygger dokumentet med 6-fas tidsplaneringstabell.
**When to use:** Alltid vid docx-generering i detta projekt.
**Implikation for fas 2:** Ateranvand exakt samma hjalpfunktioner och styling som i generate-lektion-2.js. Konsistens i dokumentutseende ar viktigare an kreativ formatering. A4-format (11906 x 16838 DXA), 1"-marginaler, Arial typsnitt, mork header-rad i tabeller.

### Pattern 2: 6-fas lektionsstruktur (80 minuter)
**What:** Varje lektionsplan foljer Rosenshines evidensbaserade modell med sex faser.
**Tidsbudget for 80 min:**

| Fas | Tid | Beskrivning |
|-----|-----|-------------|
| 1. Retrieval review | 8 min | Aktiv aterkallelse (L1: forkunskapsaktivering istallet) |
| 2. Malaktivering | 3 min | Larandemal + kognitiv krok |
| 3. Explicit instruktion | 12 min | Nytt innehall med worked examples |
| 4. Guidad ovning | 25 min | Langsta fasen - eleverna arbetar med stod |
| 5. Sjalvstandig ovning | 15 min | Eleven demonstrerar utan stod |
| 6. Avslut | 5 min | Exit ticket + preview |

**Undantag:**
- L1 har ingen retrieval review (forsta lektionen) - istallet forkunskapsaktivering
- L6 (examination) har annorlunda fasfordelning - fas 4+5 slat ihop till sjalvstandigt skrivande
- L8 (seminarium) har annorlunda fasfordelning - fas 3+4+5 ar seminariets genomforande

### Pattern 3: Aterkommande case (2-3 genom hela momentet)
**What:** 2-3 huvudexempel som introduceras i L1 och aterkomeri L2-L8 fran nya vinklar.
**Rekommenderade case (Claude's discretion):**

**Case 1: AI-genererade TikTok-konton i Sverige (2025)**
- Fyra konton utgav sig for att vara svenska tjejer, men var AI-genererade
- Spred hogerpropaganda, tusentals foljare
- En enskild person bakom (erkande)
- Bra for: inokulering (L1), kallkritiska verktyg (L2), AI+konspiration (L3), skrivuppgift (L6)

**Case 2: Deepfakes infor valet 2026**
- Manipulerade filmer och bilder av politiska foretraddare
- Syftar till att skapa misstro och paverka opinion
- 74% av nyskapat natiinnehall 2025 genereras av AI/botar
- Bra for: AI-labb (L1), konspirationsteori (L3), seminarium (L8)

**Case 3 (internationellt): DeepSeek och konspirationsspridning (2025)**
- Kinesisk AI-modell testades pa konspirationsteorier
- Visar hur AI-modeller kan generera overtygande konspirationstexter
- Bra for: AI-labb (L1), AI+konspiration (L3), seminarium (L8)

### Pattern 4: L1 AI-labb - Prebunking/inokulering
**What:** Eleverna skapar sjalva overtygandeinnehall med AI-verktyg i kontrollerad miljo.
**Design (Claude's discretion):**

1. **Uppgift:** Eleverna far i uppdrag att anvanda en AI-chatbot (t.ex. ChatGPT, Copilot) for att skapa ett socialt medie-inlagg som ser trovaardigt ut men innehaller felaktig eller vilseledande information om ett av de aterkommande casen
2. **Sakerhetsramar:** Tredjepersons-framing introduceras fore uppgiften. Tydlig markering att allt stannar i klassrummet. Analys av VARFOR innehallet ar overtygande, inte spridning av det
3. **Reflektion:** Eleverna granskar varandras skapelser - vad gor dem overtygande? Vilka kallkritiska reaktioner borde utlosas?
4. **Inokuleringseffekt:** Genom att sjalva skapa desinformation upplever eleverna hur latt det ar, vilket bygger motstandskraft ("mental antikropp")

### Pattern 5: L6 - Summativ skrivuppgift
**What:** Examinationsstruktur for skriftlig bedomning.
**Design (locked decisions):**
- Eleverna far utdelat material (konspirationsteori-text + AI-genererat innehall) - givet, inte egetvalt
- 500-800 ord kallkritisk analys
- Rubriken avslojars forst pa examinationsdagen (eleverna vet temat men inte exakt rubrik)
- Hjalpmedel: egna anteckningar + utdelat material, ingen internet
- Bedomning mot E/C/A-kriterierna fran bedomningskriterier.md (skriftlig examination)

**6-fas anpassning:**
- Fas 1 (8 min): Retrieval review - kort repetition av centrala begrepp
- Fas 2 (3 min): Uppgiftsinstruktion och rubrikutdelning
- Fas 3 (5 min): Genomgang av materialet tillsammans
- Fas 4+5 (55 min): Sjalvstandigt skrivande (sammanslagen guidad+sjalvstandig)
- Fas 6 (5 min): Insamling + exit ticket (reflektion over upplevd svarighet)

### Pattern 6: L8 - Seminarie-examination
**What:** Muntlig examination med equity-safeguards.
**Design (locked decisions):**
- Smagrupper om 4-5 elever
- Tilldelade perspektiv (t.ex. medieforskare, AI-utvecklare, psykolog, samhallsdebattor, ungdomsrepresentant)
- Seminariefragestallning baserad pa aktuellt case (ett av de aterkommande)
- Inspelning (ljud) for efterbedomning
- Equity-safeguards: tilldelade roller, tredjepersons-framing, privat uppfoljningsplan

**6-fas anpassning:**
- Fas 1 (8 min): Retrieval review infor seminarium (tredjepersons-framing, analytiska perspektiv)
- Fas 2 (5 min): Presentation av seminariefragestallning och roller, normpaminnelse
- Fas 3+4+5 (55 min): Seminariets genomforande i smagrupper (lararen observerar och spelar in)
- Fas 6 (12 min): Avslut och sammanfattning (ingen exit ticket vid examination)

**Observera:** L8 har ingen exit ticket (summativ bedomning). Tidsfordelningen avviker fran standard for att ge tillracklig seminarietid.

### Anti-Patterns to Avoid
- **Anvand INTE ChatGPT/AI direkt i examinationer (L6, L8):** AI-labben ar L1, inte examination. L6 och L8 testar elevens egen formaga.
- **Generera INTE lektionsplaner som markdown:** Fas 2 levererar .docx, inte .md. Markdown-versioner behovs inte.
- **Skapa INTE nya generate-scripts fran grunden:** Ateranvand hjalpfunktioner och styling fran generate-lektion-2.js.
- **Ge INTE eleverna fritt val av AI-verktyg i L1:** Anvand ett specificerat verktyg for kontrollerbar miljo.
- **Inkludera INTE personliga asikter i seminariefragestallningen (L8):** Alla fragor ska vara i tredjepersons-form.

## Don't Hand-Roll

| Problem | Bygg inte | Anvand istallet | Varfor |
|---------|-----------|-----------------|--------|
| Word-dokument | Manuell XML | docx-js via generate-script | Beprovat monster i projektet, generate-lektion-2.js som mall |
| E/C/A-kriterier i lektionsplanerna | Nya kriterier | Kopiera/referera fran bedomningskriterier.md | Fas 1 har redan definierat dessa - backward design |
| Exit ticket-fragor | Nya fragor | Kopiera fran exit-ticket-mall.md | Redan definierade i fas 1, koherens kraver identiska fragor |
| Tredjepersons-framing-formuleringar | Ny formulering | Referera tredjepersons-framing.md | Konsistent norm genom hela momentet |
| 6-fas tidsplanering | Egna tider | Anvand tidsbudgeten fran pedagogik-ramverket (80 min) | Evidensbaserad fordelning |

## Common Pitfalls

### Pitfall 1: AI-labben i L1 tappar prebunking-syftet
**What goes wrong:** Aktiviteten blir "kul att testa AI" istallet for inokulering - eleverna upplever inte den kallkritiska dimensionen.
**Why it happens:** AI-verktyg ar engagerande i sig, och det ar latt att focusera pa skapandet snarare an reflektionen.
**How to avoid:** Aktiviteten MASTE ha en tydlig reflektionsfas dar eleverna analyserar VARFOR deras skapelser ar overtygande. Reflektionsfasen ar lika lang som skapandefasen. Exit ticket fran exit-ticket-mall.md kraver att eleverna namnger trovardighetsfaktorer.
**Warning signs:** Lektionsplanen har >15 min skapande men <10 min reflektion/analys.

### Pitfall 2: L6 skrivuppgift blir for oppen eller for sluten
**What goes wrong:** Antingen ar uppgiften sa oppen att eleverna inte vet vad de ska analysera, eller sa sluten att det bara kraver aterapgivning.
**Why it happens:** Balansen mellan struktur och oppethet ar svar.
**How to avoid:** Uppgiften ska ge ett specifikt material att analysera (locked decision) men en oppen analytisk fraga som tillater E/C/A-differentiering. Rubriken ska vara en fraga som kraver analys (inte "beskriv") och som kan besvaras pa alla tre nivaaer.
**Warning signs:** Rubriken borjar med "Beskriv..." eller "Lista..." istallet for "Analysera..." eller "Varfor...".

### Pitfall 3: Seminariet i L8 domineras av 1-2 elever
**What goes wrong:** Trots smagrupper talar nagra elever hela tiden medan andra ar passiva.
**Why it happens:** Standard seminarieformat gynnar verbalt starka elever.
**How to avoid:** Tilldelade perspektiv ger varje elev ett "uppdrag" att representera. Seminariefragestallningen kan ha delfragor som riktar sig till specifika perspektiv. Lararens facilitering kan inkludera "Nu vill jag hora fran [perspektiv]-experten."
**Warning signs:** Lektionsplanen saknar specifik faciliterings-instruktion for att sakerstalla att alla hors.

### Pitfall 4: Inkoherens med fas 1-dokumenten
**What goes wrong:** Lektionsplanerna refererar bedomningskriterier, exit tickets eller tredjepersons-framing som inte overensstammer med fas 1:s dokument.
**Why it happens:** Fas 2 skrivs i en annan session an fas 1, och detaljerna driftar.
**How to avoid:** Varje lektionsplan ska EXPLICIT kopiera relevanta E/C/A-kriterier fran bedomningskriterier.md och exit ticket-fragor fran exit-ticket-mall.md. Referera tredjepersons-framing.md for exempelformuleringar.
**Warning signs:** Bedomningskriterier i lektionsplanen avviker fran bedomningskriterier.md.

### Pitfall 5: L1 saknar koppling till L6 och L8
**What goes wrong:** L1 (oppning) och L6/L8 (examination) kanns som separata lektioner utan rod trad.
**Why it happens:** De skrivs i olika planer (02-01 vs 02-02) och kopplingen gloms.
**How to avoid:** De 2-3 aterkommande casen som introduceras i L1 MASTE aterkomma i L6 (analyserat material) och L8 (seminariefragestallning). Explicit referens i varje lektionsplan.

### Pitfall 6: Generate-scriptet har styling-inkonsekvens
**What goes wrong:** Ny lektionsplan ser annorlunda ut an befintliga (lektion-2.docx etc.).
**Why it happens:** Scriptet skrivs fran grunden istallet for att ateranvanda befintliga hjalpfunktioner.
**How to avoid:** Kopiera headerCell, cell, bullet, heading2, heading3 exakt fran generate-lektion-2.js. Anvand samma fargschema (header: #2C3E50), samma cellmarginaler, samma teckensnitt och storlekar.

## Code Examples

### Generate-script-monster (fran generate-lektion-2.js)
```javascript
// Source: Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/generate-lektion-2.js
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "2C3E50", type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 22 })] })],
  });
}

function cell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 22 })] })],
  });
}
```

### 6-fas tidsplanering for L1 (80 min, ingen retrieval review)
```
| Tid        | Fas                    | Aktivitet                                    |
|------------|------------------------|----------------------------------------------|
| 0-8 min    | Forkunskapsaktivering  | Snabbundersokn.: vad ar en deepfake? troligt? |
| 8-11 min   | Malaktivering          | Presentera AI-labb, tredjepersons-framing     |
| 11-23 min  | Explicit instruktion   | Inokulering: vad ar prebunking, demo av AI    |
| 23-48 min  | Guidad ovning (AI-labb)| Eleverna skapar AI-innehall, lararstod         |
| 48-70 min  | Reflektion och analys  | Granska varandras skapelser, varfor overtygande|
| 70-80 min  | Avslut                 | Exit ticket + preview av L2                    |
```

### Kommandon for att generera och validera
```bash
# Generera lektionsplan
cd "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/"
node generate-lektion-1.js

# Validera
python resources/office-scripts/validate.py lektion-1.docx
```

## State of the Art

| Aspekt | Angreppssatt | Stod |
|--------|-------------|------|
| Inokulering/prebunking | Elever skapar sjalva overtygandeinnehall | van der Linden et al. (2020), Bad News-spelet testat pa svenska gymnasieelever |
| Aktiv inokulering > passiv | Skapande (perspektiv-byte) starkare an enbart exponering | Cambridge SDML forskning; kombinerar perspective-taking + experiential learning |
| Equity-safeguards vid konspirationsteorier | Tilldelade roller, tredjepersons-framing, privat uppfoljning | Frontiers in Education (2025): Teaching controversial issues in secondary education |
| AI-genererat innehall 2025 | 74% av nyskapat natiinnehall ar AI/bot-genererat | Microsoft-rapport 2025, Internetstiftelsen |
| Deepfakes och val | Vaxande hot mot demokratiska processer | Riksdagsmotion 2025/26:1815, Internetstiftelsen rapport infor valet 2026 |

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Manuell granskning + automatiserad filvalidering |
| Config file | Ingen - docx-validering + innehallsgranskning |
| Quick run command | `python resources/office-scripts/validate.py lektion-1.docx` |
| Full suite command | Validera alla tre .docx + manuell granskning mot success criteria |

### Phase Requirements -> Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| KONSP-02 | L1 innehaller AI-labb med inokulering | manual-only | Granska att lektion-1.docx har AI-labb-aktivitet med skapande + reflektion | N/A Wave 0 |
| KONSP-03 | Aktuella 2025-2026 svenska/internationella exempel | manual-only | Granska att alla tre lektionsplaner refererar aterkommande case | N/A Wave 0 |
| BED-02 | L8 seminarium med equity-safeguards | manual-only | Granska att lektion-8.docx har tilldelade roller, tredjepersons-framing, privat uppfoljningsplan | N/A Wave 0 |

### Sampling Rate
- **Per task commit:** `python resources/office-scripts/validate.py [fil].docx` + manuell granskning
- **Per wave merge:** Alla tre .docx validerade + inbordes koherens kontrollerad
- **Phase gate:** Alla success criteria verifierade

### Wave 0 Gaps
None - docx-generering ar etablerat i projektet med befintliga generate-scripts som mall.

## Open Questions

1. **Vilket AI-verktyg ska eleverna anvanda i L1?**
   - What we know: Claude's discretion. Eleverna behover ett verktyg som kan generera text (och garna bilder) snabbt.
   - Recommendation: ChatGPT (gratis version) - mest tillgangligt, de flesta elever har redan konto. Alternativt Copilot (inbyggt i Windows). Specificera ETT verktyg i lektionsplanen for kontrollerbar miljo. Ta med en backup-plan om verktyget ar nere.

2. **Exakt rubrikformulering for L6**
   - What we know: Claude's discretion. Rubriken ska kreva analys och tillata E/C/A-differentiering. Eleverna vet temat men inte exakt rubrik.
   - Recommendation: Formulera 2-3 rubrikforslag i lektionsplanen som lararen kan valja mellan. T.ex. "Analysera det bifogade materialet med hjalp av kallkritiska verktyg. Varfor kan detta innehall upplevas som trovärdigt, och vilka mekanismer ligger bakom?" Rubriken ska vara oppen nog for A-niva men styrd nog for E-niva.

3. **Seminariefragestallningens exakta formulering**
   - What we know: Ska baseras pa aktuellt case, inte abstrakt fragestallning. Tredjepersons-framing.
   - Recommendation: T.ex. "Med utgangspunkt i [case]: Analysera fran ert tilldelade perspektiv varfor denna typ av innehall kan upplevas som trovärdigt och vilka konsekvenser det kan fa for samhallet."

## Sources

### Primary (HIGH confidence)
- `generate-lektion-2.js` - Befintligt generate-script-monster for docx-generering
- `bedomningskriterier.md` - E/C/A-kriterier for L6 och L8 (fas 1 output)
- `exit-ticket-mall.md` - Exit ticket-fragor for alla lektioner (fas 1 output)
- `tredjepersons-framing.md` - Normdokument (fas 1 output)
- `momentplan.md` - 8-lektionsoversikt med progression (fas 1 output)
- `.claude/skills/planera-moment/SKILL.md` - 6-fas lektionsmodell, kvalitetskontroll
- `.claude/skills/planera-moment/references/lektionsplanering.md` - Lektionsmall och 6-fas tidsbudget
- `.claude/skills/docx/SKILL.md` - docx-js referens

### Secondary (MEDIUM confidence)
- [Bad News-spelet pa svenska gymnasieelever](https://www.tandfonline.com/doi/full/10.1080/15391523.2024.2338451) - Prebunking i svensk gymnasiemiljo
- [Internetstiftelsen om valet 2026](https://www.voister.se/artikel/2026/02/internetstiftelsen-sa-kan-generativ-ai-deepfakes-och-sociala-medier-paverka-valet-2026/) - Aktuella svenska deepfake-exempel
- [Riksdagsmotion om markning av AI-genererat material](https://www.riksdagen.se/sv/dokument-och-lagar/dokument/motion/markning-av-ai-genererat-material-for-att-motverka_hd021815/) - Svensk politisk kontext
- [DeepSeek och konspirationsteorier](https://arbetet.se/2025/02/05/hur-bra-ar-kinesiska-ain-deepseek-pa-att-sprida-konspirationsteorier/) - AI och konspirationsspridning
- [Google prebunking guide](https://prebunking.withgoogle.com/docs/A_Practical_Guide_to_Prebunking_Misinformation.pdf) - Praktisk guide for prebunking

### Tertiary (LOW confidence)
- Specifik procentandel "74% AI-genererat innehall" - hittad via Internetstiftelsen men ej primarkallaverifierad

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - docx-js och generate-script-monster ar etablerade i projektet med 6 befintliga filer som referens
- Architecture: HIGH - 6-fas modellen, backward design, E/C/A-progression ar val dokumenterade i pedagogik-ramverket och fas 1
- Pitfalls: HIGH - Baserat pa kanda svagheter i AI-genererad lektionsplanering (dokumenterat i SKILL.md) + specifica risker for prebunking-aktiviteter
- Aktuella exempel: MEDIUM - websokning ger relevanta svenska/internationella case, men specifika detaljer bor verifieras

**Research date:** 2026-03-12
**Valid until:** 2026-04-12 (stabil pedagogisk doman; aktuella exempel kan behova uppdateras)
