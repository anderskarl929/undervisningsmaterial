# Phase 3: Verktygslektioner (L2 + L3) - Research

**Researched:** 2026-03-12
**Domain:** Lektionsplanering - kallkritiska verktyg, SIFT/lateral reading, Harvard-referering, AI+konspirationsteorier-intersection, stationsarbete
**Confidence:** HIGH

## Summary

Fas 3 ska producera tva nya lektionsplaner som .docx (med tillhorande .pptx): L2 (kallkritiska verktyg, SIFT, lateral reading, Harvard-referering) och L3 (AI+konspirationsteori-intersection med stationsarbete). Befintliga generate-lektion-2.js och generate-lektion-3.js innehaller gammal design fran fore roadmap-processen och maste skrivas om helt. Det nya innehallet styrs av CONTEXT.md:s lasta beslut.

L2 ar en verktygsfokuserad lektion dar eleverna far de fyra kallkritiska grundfragorna, SIFT-metoden och lateral reading som hands-on ovning, plus en introduktion av Harvard-referering. L3 ar "ny pedagogisk mark" (fran STATE.md) - AI+konspirationsteori-intersektionen har fa pedagogiska modeller att folja. Stationsarbete ar beslutat format: varje station = en AI-mekanism (deepfakes, filterbubbla, AI-genererade kallor, automatiserade bottar), eleverna roterar i par/grupper och analyserar med SIFT/grundfragorna. L3 exit ticket fungerar som gateway check infor L4-seminariet.

**Primary recommendation:** Skriv om generate-lektion-2.js och generate-lektion-3.js fran grunden med nytt innehall (kopiera hjalpfunktioner exakt fran befintliga scripts). Skapa tva .pptx-presentationer med pptxgenjs. Varje lektionsplan foljer 6-fas strukturen med 80 minuters tidsbudget.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- L3 AI+konspirations-intersection: Mekanismfokus med 3-4 mekanismer (t.ex. deepfakes, filterbubbla, AI-genererade kallor, automatiserade bottar) som eleverna utforskar
- L3 Stationsarbete: varje station = en mekanism, eleverna roterar i par/grupper och analyserar exempel vid varje station med strukturerade fragor
- L3 Mix av case: nagra stationer anvander de aterkommande casen fran L1, andra har nya traffsakra exempel som battre illustrerar just den mekanismen
- L3 Strukturerad koppling till kallkritik: varje station har en explicit fraga dar eleverna anvander SIFT/grundfragorna for att granska det AI-genererade exemplet - de ovar L2-verktygen pa nytt material
- Gateway check (L3 exit ticket): Mater verktygstillampning - kan eleven sjalvstandigt tillampa kallkritiska verktyg pa ett nytt exempel utan scaffolding?
- Gateway check resultat: anvands for gruppsammansattning infor L4-seminariet - blanda starkare/svagare elever strategiskt
- Gateway check tva delar: Del 1 = gateway-fraga (kallkritisk tillampning pa nytt exempel), Del 2 = kort Harvard-referingsovning (formatera kallor korrekt)
- Enkel checklista som troskel: 2-3 konkreta kriterier (t.ex. "anvander minst 2 kallkritiska verktyg", "identifierar ratt problem") - snabbt for lararen att sortera

### Claude's Discretion
- L2 ovningsdesign: hur hands-on ovningen med SIFT/lateral reading struktureras, vilka verkliga kallor som anvands, balans mellan grundfragor och SIFT
- Harvard-referering introduktion i L2: hur mycket tid, forenklad eller fullstandig modell
- Antal och val av specifika mekanismer for L3-stationerna (inom ramen for mekanismfokus)
- Stationsarbetets praktiska logistik (tid per station, antal stationer, rotationsordning)
- Presentationernas (.pptx) design och innehall
- Exakta gateway check-kriterier (inom ramen for enkel checklista)

### Deferred Ideas (OUT OF SCOPE)
None - discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| KRIT-01 | Elev kan tillampa de fyra kallkritiska grundfragorna pa digitalt innehall | L2 introducerar grundfragorna explicit med hands-on ovning; L3 stationsarbete tillamppar dem pa AI-genererat innehall |
| KRIT-02 | Elev kan anvanda SIFT-metoden/lateral reading for att verifiera pastaenden | L2 introducerar SIFT med praktisk ovning pa verkliga exempel; L3 ateranvander SIFT vid varje station |
| KRIT-03 | Elev kan anvanda Harvard-referering i skriftliga uppgifter | L2 introducerar Harvard-referering; L3 exit ticket kraver korrekt kallhanvisning (gateway check Del 2) |
| KRIT-04 | Elev kan anvanda digitala verifieringsverktyg (omvand bildsokning, Google Lens) | L2 visar omvand bildsokning som del av lateral reading/SIFT (Investigate the source) |
| KONSP-01 | Elev kan forklara psykologiska mekanismer bakom konspirationstro | L3 stationsarbete analyserar mekanismer (t.ex. filterbubbla, confirmation bias) i konkreta exempel |
| KONSP-04 | Elev kan analysera hur AI forstarker konspirationsteoriers spridning | L3 huvudfokus: varje station ar en AI-mekanism (deepfakes, filterbubbla, AI-genererade kallor, bottar) |
| AI-01 | Elev kan kallkritiskt granska AI-genererat innehall | L3 stationer har explicit fraga dar eleverna tillamppar SIFT/grundfragorna pa AI-genererade exempel |
</phase_requirements>

## Standard Stack

### Core
| Komponent | Format | Syfte | Varfor |
|-----------|--------|-------|--------|
| Lektionsplaner | .docx (via docx-js) | Detaljerade lektionsplaner for L2 och L3 | Lasbara i Word/Google Docs; etablerat format i projektet |
| Presentationer | .pptx (via pptxgenjs) | Visuellt stod for instruktionsmoment | PPTX-skill tillganglig; presentationer behovs for instruktionsfaser |
| Generate-scripts | .js (Node.js) | Genererar .docx och .pptx programmatiskt | Befintligt monster i projektet |
| docx-js | npm-paket | Skapar Word-dokument | Redan installerat i projektkatalogen |
| pptxgenjs | npm-paket | Skapar PowerPoint-presentationer | Rekommenderat i PPTX-skill |

### Supporting
| Komponent | Syfte | Nar |
|-----------|-------|-----|
| momentplan.md | Lektionsoversikt och progressionslogik | Referens for innehall och kopplingar |
| bedomningskriterier.md | E/C/A-kriterier | Kopieras in i lektionsplanerna |
| exit-ticket-mall.md | Exit ticket-fragor for L2, L3 | Kopieras in i lektionsplanerna |
| tredjepersons-framing.md | Norm och exempelformuleringar | Refereras sarskilt i L2 och L3 (forstarkt paminnelse) |
| lektion-1.docx (L1 fran fas 2) | Aterkommande case, AI-labb-innehall | L2 retrieval review baseras pa L1 exit ticket; L3 case aterkommerl |
| validate.py | Validerar .docx-filer | Kor efter varje genererad fil |

### Outputkatalog
Alla filer sparas i: `Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/`

## Architecture Patterns

### Rekommenderad filstruktur for fas 3
```
Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/
  generate-lektion-2.js   # OMSKRIVEN - ny design med SIFT/grundfragor/Harvard
  lektion-2.docx          # OMSKRIVEN - kallkritiska verktyg-lektion
  lektion-2.pptx          # NY - presentation for L2
  generate-lektion-3.js   # OMSKRIVEN - ny design med AI+konsp stationsarbete
  lektion-3.docx          # OMSKRIVEN - AI+konspirationsteori-intersection
  lektion-3.pptx          # NY - presentation for L3
```

**VIKTIGT:** Befintliga generate-lektion-2.js och generate-lektion-3.js ar fran den gamla 7-lektionsdesignen (fore roadmap). De maste skrivas om helt med nytt innehall. Hjalpfunktionerna (headerCell, cell, bullet etc.) ska ateranvandas, men sjalva lektionsinnehallet ersatts.

### Pattern 1: Befintligt generate-script-monster
**What:** Varje lektionsplan genereras av ett Node.js-script med docx-js. Hjalpfunktioner och styling ar identiska over alla lektioner.
**When to use:** Alltid vid docx-generering i detta projekt.
**Implikation:** Kopiera headerCell, cell, bullet, heading2, heading3, bodyText, italicText, boldBodyText, boldItalicBody, spacer EXAKT fran generate-lektion-1.js (fas 2 output). A4-format (11906 x 16838 DXA), 1"-marginaler, Arial typsnitt, mork header-rad i tabeller (#2C3E50).

### Pattern 2: 6-fas lektionsstruktur (80 minuter)
**What:** Varje lektionsplan foljer Rosenshines 6-fas modell.
**L2 tidsbudget (80 min):**

| Fas | Tid | Beskrivning |
|-----|-----|-------------|
| 1. Retrieval review | 8 min | Aktiv aterkallelse fran L1 (baserat pa L1 exit ticket) |
| 2. Instruktion | 15 min | De fyra grundfragorna + SIFT-metoden |
| 3. Guidad ovning | 25 min | Hands-on SIFT/lateral reading pa verkliga kallor |
| 4. Sjalvstandig ovning | 15 min | Eleverna tillamppar pa eget exempel |
| 5. Harvard-referering intro | 9 min | Kort intro + ovning |
| 6. Avslut | 8 min | Exit ticket + preview av L3 |

**L3 tidsbudget (80 min):**

| Fas | Tid | Beskrivning |
|-----|-----|-------------|
| 1. Retrieval review | 8 min | Aktiv aterkallelse fran L2 (grundfragorna, SIFT) |
| 2. Instruktion | 10 min | AI-mekanismer - hur AI forstarker konspirationsspridning |
| 3. Stationsarbete | 35 min | 3-4 stationer, rotation i par/grupper |
| 4. Summering | 10 min | Helklassdiskussion om monster over stationerna |
| 5. Gateway exit ticket | 12 min | Del 1: tillampning + Del 2: Harvard-referering |
| 6. Framatkoppling | 5 min | Preview av L4, hur gateway-resultat anvands |

### Pattern 3: SIFT-metoden for L2
**What:** SIFT (Stop, Investigate the source, Find better coverage, Trace claims) ar Mike Caulfields metod for snabb kallkritisk bedoming. Lateral reading ar den centrala tekniken - oppna en ny flik och sok pa avsandaren istallet for att fordjupa dig i sjalva innehallet.
**How to teach:** Demo med verkligt exempel (lararen visar lateral reading steg for steg), sedan eleverna ovar pa egna kallor.
**Koppling till grundfragorna:** SIFT ar den PRAKTISKA metoden, grundfragorna ar det ANALYTISKA ramverket. L2 introducerar bada och visar hur de hanger ihop.

### Pattern 4: Stationsarbete i L3
**What:** 3-4 fysiska stationer i klassrummet, varje station fokuserar pa en AI-mekanism.
**Design:**
- **Station 1: Deepfakes** - visuellt AI-genererat innehall (bilder/video). Eleverna granskar med grundfragorna: Vem? Varfor? Hur? Nar? + SIFT.
- **Station 2: Filterbubbla/ekkammare** - hur algoritmer forstarker konspirationsteorier. Eleverna analyserar ett flodesschema over algoritmisk forstorkning.
- **Station 3: AI-genererade kallor** - fejkade nyhetssajter (2000+ identifierade av NewsGuard). Eleverna jamfor en fejkad nyhetssajt med en riktig.
- **Station 4: Automatiserade bottar** - hur botnatverk sprider narrativ i stor skala. Eleverna analyserar ett natverk av inlagg.

**Varje station innehaller:**
1. Kort bakgrundstext om mekanismen (2-3 stycken)
2. Ett konkret exempel (fran aterkommande case ELLER nytt traffsukert exempel)
3. Strukturerade fragor: (a) Beskriv mekanismen, (b) Anvand SIFT pa exemplet, (c) Hur hade de kallkritiska grundfragorna hjalpt har?
4. Reflektionsfraga: Varfor kan nagon som moter detta utan forvarning finna det overtygande? (tredjepersons-framing)

**Logistik (Claude's discretion rekommendation):**
- 4 stationer, 8-9 min per station
- Eleverna arbetar i par (inte storre grupper - varje elev maste vara aktiv)
- Rotation med signal (lararklapp eller timer)
- Skriftliga svar pa stationsblad (samlas in for formativ bedomning)

### Pattern 5: Gateway check (L3 exit ticket)
**What:** Tva-delad exit ticket som fungerar som gateway check infor L4.
**Del 1: Kallkritisk tillampning (7 min)**
- Eleverna far ett NYTT exempel (inte fran stationerna) och ska tillamppa kallkritiska verktyg UTAN scaffolding
- Exempel: en AI-genererad nyhetsartikel eller ett socialt medie-inlagg
- Ingen meningsstartare, inga ledtradare - syftet ar att se om eleven klarar sjalvstandig tillampning

**Del 2: Harvard-referering (5 min)**
- Kort ovning: formatera 2-3 kallor korrekt med Harvardsystemet
- Byggnar vidare pa L2:s introduktion

**Checklista for lararen (snabbsortering):**
1. Anvander minst 2 kallkritiska verktyg (grundfragor och/eller SIFT) ✓/✗
2. Identifierar ratt problem/vilseledning i exemplet ✓/✗
3. Harvard-referering i huvudsak korrekt (minst 1 av 2-3 kallor) ✓/✗

**Resultat -> gruppsammansattning L4:**
- Alla tre ✓: Stark - kan stotta andra i seminariegrupp
- 1-2 ✓: Medel - behover blandning med starkare elever
- 0 ✓: Behover extra stod - lararen planerar stodinsats infor L4

### Anti-Patterns to Avoid
- **Skriv INTE om fran scratch utan att kopiera hjalpfunktioner:** Hjalpfunktionerna (headerCell etc.) MASTE vara identiska med befintliga generate-scripts for konsistent utseende.
- **Anvand INTE den gamla L2/L3-designen som referens for INNEHALL:** Det ar bara hjalpfunktionerna och dokumentstrukturmonstret som ateranvands. Allt lektionsinnehall ar nytt.
- **Gor INTE stationsarbetet utan explicit SIFT/grundfrage-koppling:** CONTEXT.md kraver att varje station har en explicit fraga dar eleverna anvander L2-verktygen. Detta ar den kritiska bryggan.
- **Ge INTE gateway exit ticket med scaffolding:** Hela poangen ar att mata om eleven klarar det UTAN stodstrukturer.
- **Blanda INTE ihop .pptx-generering med docx-generering:** Anvand pptxgenjs for presentationer (separat script eller integrerat), inte docx-js.

## Don't Hand-Roll

| Problem | Bygg inte | Anvand istallet | Varfor |
|---------|-----------|-----------------|--------|
| Word-dokument | Manuell XML | docx-js via generate-script | Beprovat monster i projektet |
| PowerPoint | Manuell XML | pptxgenjs via PPTX-skill | Standardverktyg, designstod i SKILL.md |
| Exit ticket-fragor | Nya fragor | Kopiera fran exit-ticket-mall.md | Redan definierade i fas 1, koherens kraver identiska fragor |
| E/C/A-kriterier | Nya kriterier | Kopiera fran bedomningskriterier.md | Backward design - fas 1 definierade dessa |
| SIFT-innehall | Egna beskrivningar | Mike Caulfields officiella SIFT-material | Etablerad metod med specifik terminologi |
| Gateway check-design | Komplex bedomningsmatris | Enkel 3-punkts checklista (CONTEXT.md) | Maste vara snabb for lararen att sortera |

## Common Pitfalls

### Pitfall 1: L2 blir for teoritung - SIFT/grundfragorna utan hands-on
**What goes wrong:** Lektionen blir en forelesning om kallkritik istallet for praktisk ovning. Eleverna kan repetera begrepp men inte tillampa dem.
**Why it happens:** Det ar frestande att forklara allt noggrant innan eleverna ovar. Men 15 min teori + 25 min ovning ar en battre ratio an 30 min teori + 10 min ovning.
**How to avoid:** Instruktionen (fas 2) ska vara MAX 15 min. Visa grundfragorna, demonstrera SIFT med ETT verkligt exempel (modelering), ga sedan direkt till guidad ovning dar eleverna gor det sjalva.
**Warning signs:** Instruktionsfasen ar >20 min.

### Pitfall 2: L3 stationsarbete utan tydlig struktur
**What goes wrong:** Eleverna vandrar mellan stationer utan att forsta vad de ska gora. Tiden gar at till logistik istallet for larande.
**Why it happens:** Stationsarbete kraver tydligare instruktioner an vanligt grupparbete.
**How to avoid:** Skriftliga stationsblad med tydliga fragor vid varje station. Timer/signal for rotation. Kort demo av en station innan eleverna borjar.
**Warning signs:** Lektionsplanen saknar stationsblad som bilagor.

### Pitfall 3: Gateway check testar minneskunskap istallet for tillampning
**What goes wrong:** Exit ticket fragar "Vad star SIFT for?" istallet for "Tillamppa dina kallkritiska verktyg pa detta exempel."
**Why it happens:** Tillampningsfragor ar svarare att designa och bedomma.
**How to avoid:** Gateway-fragan MASTE ge ett NYTT, obekant exempel och be eleverna analysera det med sina verktyg. Checklistan bedommar om de faktiskt ANVANDER verktygen, inte om de kan namnge dem.
**Warning signs:** Gateway-fragan kan besvaras utan att anvanda nagot verktyg.

### Pitfall 4: Harvard-referering tar for mycket tid i L2
**What goes wrong:** Harvard-introduktionen expanderar och ater in i ovningstiden for SIFT/grundfragorna.
**Why it happens:** Referering ar ett eget amne med manga detaljer.
**How to avoid:** L2 introducerar Harvard pa GRUNDLAGGANDE niva (max 9 min): varfor kallhanvisa + 2-3 exempel + kort ovning. Fullstandig ovning sker i L3 exit ticket och vidare i L5. Forenklad modell: (Efternamn, ar) i texten + Efternamn (Ar). Titel. Kalla. i referenslistan. Inget mer.
**Warning signs:** Harvard-sektionen i L2 ar >12 min.

### Pitfall 5: L3 stationer har BARA aterkommande case
**What goes wrong:** Alla stationer anvander samma case fran L1. Eleverna ser inget nytt material.
**Why it happens:** Det ar enkelt att ateranvanda.
**How to avoid:** CONTEXT.md sager "Mix av case: nagra stationer anvander de aterkommande casen fran L1, andra har nya traffsakra exempel." Rekommendation: 2 stationer med aterkommande case (AI-TikTok, DeepSeek/deepfakes), 2 stationer med NYA exempel (NewsGuard-sajter, orkanen Helene-bilder fran generate-lektion-3.js bilaga 1).
**Warning signs:** Inga nya exempel i stationsarbetet.

### Pitfall 6: Inkoherens med fas 1 och fas 2 dokument
**What goes wrong:** L2/L3 refererar begrepp eller case som inte stammer overens med momentplan, exit-ticket-mall eller lektion-1.docx.
**Why it happens:** Fas 3 skrivs i separat session.
**How to avoid:** Exit ticket for L2 MASTE matcha exit-ticket-mall.md. L2 retrieval review MASTE baseras pa L1 exit ticket. L3 retrieval review MASTE baseras pa L2 exit ticket. Aterkommande case maste vara desamma som i L1.
**Warning signs:** Exit ticket-fragor avviker fran exit-ticket-mall.md.

## Code Examples

### Generate-script hjalpfunktioner (kopiera exakt)
```javascript
// Source: generate-lektion-1.js / generate-lektion-2.js (befintliga)
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
// ... alla hjalpfunktioner fran befintliga scripts
```

### PPTX-generering (fran PPTX-skill)
```bash
# Installera om inte redan globalt
npm install -g pptxgenjs

# Generera presentation
node generate-lektion-2-pptx.js
```

Se `.claude/skills/pptx/pptxgenjs.md` for fullstandig dokumentation av pptxgenjs-API:et.

### Kommandon for att generera och validera
```bash
cd "Undervisningsmaterial/Samhallskunskap/Kallkritik AI och konspirationsteorier/"

# Generera docx
node generate-lektion-2.js
node generate-lektion-3.js

# Validera docx
python "/home/anders/Second brain/resources/office-scripts/validate.py" lektion-2.docx
python "/home/anders/Second brain/resources/office-scripts/validate.py" lektion-3.docx

# Generera pptx (visual QA kraver konvertering)
# Se PPTX-skill for QA-workflow
```

## State of the Art

| Aspekt | Angreppssatt | Stod |
|--------|-------------|------|
| SIFT-metoden | 4-stegsmodell for snabb kallkritik (Stop, Investigate, Find, Trace) | Mike Caulfield, University of Washington; anvands brett i amerikanskt utbildningssystem |
| Lateral reading | Oppna ny flik, sok pa avsandaren istallet for att lasa innehallet | Stanford History Education Group (Wineburg & McGrew, 2017) |
| Stationsarbete (learning stations) | Roterande par/grupper, varje station = fokuserad uppgift | Etablerad pedagogisk metod for aktivt larande med manga verktyg/perspektiv |
| Gateway check / mastery check | Kort bedomning som avgol nasta steg | Formativ bedomning (Black & Wiliam, 1998); anvands i "mastery learning" |

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Manuell granskning + automatiserad filvalidering |
| Config file | Ingen - docx-validering + innehallsgranskning |
| Quick run command | `python resources/office-scripts/validate.py lektion-2.docx` |
| Full suite command | Validera alla .docx + .pptx + manuell granskning mot success criteria |

### Phase Requirements -> Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| KRIT-01 | L2 introducerar de fyra grundfragorna med hands-on ovning | manual-only | Granska att lektion-2.docx har grundfrage-sektion + ovningsaktivitet | N/A Wave 0 |
| KRIT-02 | L2 introducerar SIFT/lateral reading med praktisk demo | manual-only | Granska att lektion-2.docx har SIFT-introduktion + elevovning | N/A Wave 0 |
| KRIT-03 | L2 intro Harvard + L3 exit ticket ovar Harvard | manual-only | Granska bada lektionsplaner for Harvard-innehall | N/A Wave 0 |
| KRIT-04 | L2 visar omvand bildsok som del av SIFT | manual-only | Granska att lektion-2.docx namner omvand bildsok/Google Lens | N/A Wave 0 |
| KONSP-01 | L3 stationer analyserar psykologiska mekanismer | manual-only | Granska att lektion-3.docx stationsblad inkluderar mekanismer | N/A Wave 0 |
| KONSP-04 | L3 huvudfokus: AI forstarker konspirationsspridning | manual-only | Granska att lektion-3.docx har mekanismfokus i stationsarbete | N/A Wave 0 |
| AI-01 | L3 stationer har explicit SIFT/grundfrage-tillampning pa AI-innehall | manual-only | Granska stationsblad for explicit kallkritisk fraga | N/A Wave 0 |

### Sampling Rate
- **Per task commit:** `python resources/office-scripts/validate.py [fil].docx` + manuell granskning
- **Per wave merge:** Alla filer validerade + inbordes koherens kontrollerad
- **Phase gate:** Alla success criteria verifierade

### Wave 0 Gaps
None - docx-generering och pptx-generering ar etablerade i projektet med befintliga scripts och skills.

## Open Questions

1. **PPTX-design och innehall (Claude's discretion)**
   - What we know: Presentationer ska stodja instruktionsmomenten i L2 och L3. PPTX-skill rekommenderar bold fargpalett, visuella element pa varje slide, och att undvika text-only slides.
   - Recommendation: L2-pptx visar grundfragorna, SIFT-stegen, lateral reading-demo, Harvard-mall. L3-pptx visar mekanismerna (en slide per station), instruktioner for stationsarbete, timer/rotationsschema. Fargpalett: nagon av de morka/serosa alternativen i PPTX-skill (t.ex. Midnight Executive eller Ocean Gradient) for att matcha amnet.

2. **Exakta AI-mekanismer for L3-stationerna (Claude's discretion)**
   - What we know: 3-4 mekanismer, CONTEXT.md namnger deepfakes, filterbubbla, AI-genererade kallor, automatiserade bottar som exempel.
   - Recommendation: Anvand alla fyra. De tacker spektrumet av AI-hot och ger tillracklig variation. 4 stationer x 8 min = 32 min, vilket ryms inom 35-min stationsfasen.

3. **Vilka verkliga kallor for SIFT-ovningen i L2?**
   - What we know: Claude's discretion. Ovningen behover verkliga digitala kallor som eleverna kan granska.
   - Recommendation: Anvand 2-3 kallor av varierande trovardighet: (a) en riktig nyhetsartikel, (b) en vilseledande sajt, (c) ett socialt medie-inlagg. Valj kallor som ar tematiskt kopplade till momentets amne (kallkritik/AI/konspiration) men inte samma som de aterkommande casen (for att spara dem till L3). Alternativt: en av kallorna kan vara fran generate-lektion-3.js Bilaga 4-lankar (redan forskade).

## Sources

### Primary (HIGH confidence)
- `generate-lektion-2.js` (befintligt) - Hjalpfunktioner och dokumentstruktur att ateranvanda
- `generate-lektion-3.js` (befintligt) - Hjalpfunktioner och dokumentstruktur; case-kort och lankar i bilagorna ar anvandbart innehall
- `momentplan.md` - Lektionsoversikt, progressionslogik, L2 och L3 beskrivningar
- `exit-ticket-mall.md` - Exit ticket-fragor for L2 och L3 (fas 1 output)
- `tredjepersons-framing.md` - Norm med forstarkt paminnelse i L2 och L3
- `bedomningskriterier.md` - E/C/A-kriterier for moment
- `lektion-1.docx` (fas 2 output) - Aterkommande case, AI-labb-innehall, L1 exit ticket
- `.claude/skills/docx/SKILL.md` - docx-js referens
- `.claude/skills/pptx/SKILL.md` - PPTX-generering referens

### Secondary (MEDIUM confidence)
- SIFT-metoden: Mike Caulfield, University of Washington - bredd anvandning i utbildning
- Lateral reading: Wineburg & McGrew (2017), Stanford History Education Group
- NewsGuard AI-tracking (2025): 2000+ AI-genererade nyhetssajter identifierade
- Southport-kravaller och AI-bilder (2024): GNET Research
- generate-lektion-3.js Bilaga 4 lankar - redan forskade kallor for AI+desinformation

### Tertiary (LOW confidence)
- Specifik stationsarbetets tidsfordelning (8-9 min per station) - baserat pa pedagogisk erfarenhet, inte empirisk forskning om optimal stationstid

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - docx-js, pptxgenjs och generate-script-monster ar etablerade i projektet
- Architecture: HIGH - 6-fas modellen, stationsarbete, gateway check ar val dokumenterade i CONTEXT.md och pedagogisk litteratur
- Pitfalls: HIGH - Baserat pa kanda svagheter + specifica risker for verktygsintroduktion utan hands-on
- Innehall: HIGH - CONTEXT.md har detaljerade lasta beslut for L3-designen; L2 ar inom Claude's discretion men amnet (SIFT/kallkritik) ar valdefinierat

**Research date:** 2026-03-12
**Valid until:** 2026-04-12 (stabil pedagogisk doman)
