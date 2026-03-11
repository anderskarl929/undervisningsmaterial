# Ändringsplan: Tre förbättringar baserade på Devil's Advocate-granskning

**Datum:** 2026-03-10
**Berör:** Lektion 2, 3, 4, 5, 6, 7 + momentplan

---

## Ändring 1: Flytta källhänvisning från lektion 3 till lektion 2

**Problem:** Lektion 3 (75 min) har 4 lärandemål och är överbelastad. Källhänvisning (Harvard) introduceras i 10 min i slutet, försvinner i lektion 4-5, och ska tillämpas fullt i lektion 6. Tre lektioners glapp.

**Åtgärder:**

### Lektion 2 (generate-lektion-2.js)
- Korta presentationerna (55-70 min): Låt 2-3 grupper presentera muntligt, övriga ger skriftlig feedback. Frigör ~8-10 min
- Lägg till nytt block (ca 60-70 min): "Källhänvisning - kort intro till Harvard"
  - Varför källhänvisa? (2 min)
  - Hur fungerar Harvard? Visa 2-3 exempel på tavlan (4 min)
  - Dela ut kort guide (samma som nuvarande bilaga 3 i lektion 3, men förenklad)
  - Övning: "Skriv en Harvard-hänvisning till en av källorna vi använt idag" (3 min)
- Uppdatera lärandemål: Lägg till mål 4 (källhänvisning) på basic nivå
- Uppdatera framåtkoppling: Nämn att de kommer öva mer på källhänvisning

### Lektion 3 (generate-lektion-3.js)
- Ta bort blocket "Källhänvisning: Harvardsystemet" (50-60 min) helt
- Flytta den frigjorda tiden (10 min) till bearbetningsfasen - ge gruppanalysen 28 min istället för 18
- Ta bort bilaga 3 (källhänvisningsguiden) - den finns nu i lektion 2
- Behåll kravet på källhänvisning i exit ticket (eleverna har redan sett systemet)
- Uppdatera tidsplaneringen: Bearbetning 22-58 min, Summering 58-72 min
- Minska lärandemålen från 4 till 3 (ta bort mål 4 som primärt mål, behåll som övning i exit ticket)

### Momentplan (momentplan.md)
- Lektion 2: Lägg till mål 4
- Lektion 3: Ta bort mål 4 som explicit mål (det övas via exit ticket men är inte fokus)

---

## Ändring 2: Gör lektion 5 till skriftlig perspektivanalys istället för debatt

**Problem:** Lektion 4 (seminarium) och lektion 5 (debatt) är två dagars muntlig argumentation i följd. Risk för debatt-trötthet, tysta elever missgynnas, och skillnaden är otydlig. Dessutom saknas källhänvisning (mål 4) helt i lektion 4-5.

**Åtgärder:**

### Lektion 5 (generate-lektion-5.js) - STOR OMSKRIVNING
Nytt upplägg: "Perspektiven krockar - skriftlig perspektivanalys"

**Ny tidsplanering (85 min):**
- 0-8 min: Uppstart - Retrieval practice från lektion 4 (seminarium). "Skriv ner det starkaste argumentet du hörde igår - oavsett vilken sida det kom från."
- 8-15 min: Instruktion - Presentera skrivuppgiften. Gå igenom förväntningar. Påminn om Harvard-systemet (kort, 2 min - de har sett det i lektion 2 och 3)
- 15-50 min: Bearbetning - Skriftlig perspektivanalys (35 min). Eleverna väljer SAMMA fråga som debatten: "Vem bär ansvaret för att stoppa AI-driven desinformation?" Men nu ska de skriva från ALLA TRE perspektiv och landa i en egen slutsats. 300-400 ord. Minst 2 källor med Harvard-hänvisning. Rollkorten (bilaga 1) får användas som stöd.
- 50-65 min: Kamratläsning (15 min) - Byt text med en klasskamrat. Läs och ge muntlig feedback: "Vilket perspektiv var starkast? Varför?"
- 65-77 min: Helklassdiskussion (12 min) - Lyft mönster: "Vilken aktör nämndes oftast som huvudansvarig? Varför?" Koppla till demokrati.
- 77-82 min: Exit ticket (5 min) - "Formulera ETT argument som du inte höll med om i seminariet (lektion 4) men som du nu förstår bättre efter att ha skrivit om det."
- 82-85 min: Framåtkoppling till lektion 6

**Behåll från nuvarande lektion 5:**
- Rollkorten (bilaga 1) - utmärkt material, används nu som skrivstöd
- Debattreglerna tas bort (inte relevant)
- Bedömningsstödet anpassas till skriftlig form
- Källmaterial och länkar (bilaga 5) behålls

**Ny differentiering:**
- Stöd (mot E): Rollkorten som stöd, meningsstartare på tavlan ("Teknikföretagen anser att... eftersom..."), Harvard-guide framme
- Utmaning (mot A): Skriv utan rollkort. Tillägg: "Argumentera för varför kombinationen av alla tre behövs - eller varför den INTE gör det." Använd minst 3 källor.

### Momentplan (momentplan.md)
- Lektion 5: Ändra aktivitet från "Debatt med tilldelade positioner" till "Skriftlig perspektivanalys med tilldelade positioner"
- Lägg till mål 4 i lektion 5

---

## Ändring 3: Gör lektion 6 till primär summativ bedömning, lektion 7 till ren reflektion

**Problem:** Lektion 7 (65 min - kortast) ska vara summativ kulmen med en 25-min snabbskrivning utan hjälpmedel. Bedömningskriterierna presenteras först på lektion 7. Lektion 6 har en mer robust skrivuppgift (35 min, med källor) men behandlas som formativ.

**Åtgärder:**

### Lektion 6 (generate-lektion-6.js)
- Gör skrivuppgiften till den **primära summativa bedömningen**
- Lägg till bedömningsmatrisen (E/C/A per mål) - flytta den från lektion 7:s bilaga 1
- Anpassa instruktionen: "Det här är er chans att visa vad ni kan. Den här texten är ert bedömningsunderlag."
- Behåll kamratgranskning men gör den kortare (10 min istället för 13) - fokus på att eleven själv reviderar efter feedback
- Lägg till 5 min i slutet för att eleven ska kunna revidera sin text efter kamratgranskning
- Ny tidsplanering:
  - 0-8: Uppstart (oförändrad)
  - 8-20: Instruktion + Harvard-genomgång (oförändrad)
  - 20-55: Skrivarbete (oförändrad, 35 min)
  - 55-65: Kamratgranskning (10 min, kortad)
  - 65-70: Revidering efter feedback (5 min, NY)
  - 70-75: Exit ticket + framåtkoppling (5 min, kortad)

### Lektion 7 (generate-lektion-7.js)
- Ta bort snabbskrivningen som summativ bedömning
- Ersätt med kortare "syntesreflektion" (15 min, inte 25) - samma fråga men UTAN bedömningspress. Kan vara muntlig eller skriftlig.
- Utöka återblicken (bilderna från lektion 1) - ge den mer tid, det är ett starkt pedagogiskt moment
- Utöka reflektionen och självvärderingen
- Ny tidsplanering (65 min):
  - 0-12: Återblick med bilder + diskussion (utökad från 10 min)
  - 12-15: Momentets resa (kort sammanfattning)
  - 15-30: Syntesreflektion - "Skriv fritt: Vad tar du med dig? Hur tänker du annorlunda nu?" (15 min, lågstakes)
  - 30-40: Par-samtal + helklass (10 min)
  - 40-50: Självvärderingsmatris (10 min, utökad)
  - 50-58: Gemensam avslutning - "Vad tar ni med er?" (8 min)
  - 58-65: Framåtblick (7 min)

### Momentplan (momentplan.md)
- Lektion 6: Markera som "Summativ bedömning - skriftlig analys med källhänvisning"
- Lektion 7: Ändra till "Syntes, reflektion och avslutning" (ta bort summativt fokus)

### Tidigt kommunicerade bedömningskriterier
- Lägg till i lektion 1:s framåtkoppling eller förberedelse: "I slutet av momentet bedöms ni på en skriftlig analys med källhänvisning (lektion 6)"
- Påminn i lektion 3 och 5

---

## Sammanfattning av berörda filer

| Fil | Ändringens storlek |
|-----|-------------------|
| momentplan.md | Mindre uppdateringar (mål, aktivitetsbeskrivningar) |
| generate-lektion-2.js | Medium (nytt Harvard-block, kortade presentationer) |
| generate-lektion-3.js | Medium (ta bort Harvard-block, utöka analystid) |
| generate-lektion-5.js | Stor omskrivning (debatt -> skriftlig perspektivanalys) |
| generate-lektion-6.js | Medium (summativ markering, kortad kamratgranskning, ny revidering) |
| generate-lektion-7.js | Medium (ta bort summativ snabbskrivning, utöka reflektion) |

## Ordning att implementera
1. Momentplan först (den styr allt annat)
2. Lektion 2 + 3 (hänger ihop - källhänvisningsflytten)
3. Lektion 5 (stor omskrivning, oberoende av de andra)
4. Lektion 6 + 7 (hänger ihop - summativ flytten)
5. Generera nya .docx-filer för alla ändrade lektioner
