---
created: 2026-05-07
updated: 2026-05-07
created_by: claude-opus-4-7
updated_by: claude-opus-4-7
agent_version: 03.26
purpose: Självförsörjande arbetsorder för att generera allt material för momentet "Andra världskriget och Förintelsen" (Historia 1a1) i en ny session efter /clear
---

# Generationsplan: Andra världskriget och Förintelsen

## ⚠ Hur denna fil används

Detta dokument är en **fullständig arbetsorder** för en ny session med rent kontextfönster. Steg 1–4 av planera-moment-skillens 7-stegsprocess är **klara** (se `momentplan.md` i samma katalog). Återstår: generering av material (steg 5, 5a, 5b, 6, 7 och avslut).

**När du tar över denna planering:**

1. Läs hela denna fil först.
2. Läs också `momentplan.md` för bakgrund.
3. Läs kursminnet: `.claude/planera-moment/minne/historia-1a1.md` — *särskilt viktigt eftersom det avviker från skillens defaults (4-fasstruktur istället för sex-fas, eget designsystem istället för Arkiv v2.1).*
4. Aktivera NotebookLM Historia 1a1: `notebooklm use 1bc8bd88-16ca-4b2a-b8c6-c62e6dcf2bce`
5. Generera material i ordning enligt **§ Genereringsuppgifter** nedan, **en uppgift i taget**, och be läraren godkänna varje innan du fortsätter.
6. Avsluta med att uppdatera kursminnet enligt skillens avslutningsinstruktioner.

---

## § Snabbsammanfattning

| | |
|---|---|
| **Ämne/Kurs** | Historia 1a1 (50 poäng) |
| **Tema** | Andra världskriget och Förintelsen |
| **Vinkel** | Stegen mot folkmord genom **Hédi Frieds** familjs öde |
| **Pedagogisk ansats** | **Ansats A — Narrativ ankarpedagogik** (Hédis berättelse driver varje lektion; begrepp introduceras *"när Hédi möter dem"*) |
| **Omfång** | 7 lektioner × 40 min |
| **Bedömning** | Skriftligt prov efter L7 |
| **Designsystem** | Historia 1a1 — navy/guld/crème (INTE Arkiv v2.1) |
| **Differentiering** | Standard + förenklad version av varje arbetsblad/elevtext |
| **Status detta moment** | Omarbetning av tidigare "Världskrigen"-moment (2026-03-29) |
| **NotebookLM** | Aktiverad. Innehåller *A History of World Societies* + *The Penguin History of the World*. **Hédi Fried saknas** — använd Claudes inbyggda kunskap med `[VERIFIERA]`. Primo Levi finns dokumenterad och kan användas som komplement. |

---

## § Lärandemål

### Mål 1 — Redogöra för andra världskrigets ram och Förintelsens stadier *(Bloom 2)*
Eleven ska kunna redogöra för andra världskrigets huvuddrag och Förintelsens stadier — exemplifierat genom Hédi Frieds och hennes familjs öde.
- E: översiktligt redogöra
- C: utförligt redogöra
- A: utförligt och nyanserat redogöra

### Mål 2 — Använda historiska begrepp om Förintelsen *(Bloom 2–3)*
Eleven ska kunna använda 11 nyckelbegrepp: *antisemitism, rasideologi, eugenik, ghetto, deportation, koncentrationsläger, förintelseläger, "den slutliga lösningen", folkmord, bystander, vittnesmål*.
- E: använda med viss säkerhet i enkla sammanhang
- C: använda med viss säkerhet i mer komplexa sammanhang
- A: använda med säkerhet, kunna nyansera

### Mål 3 — Förklara samband: hur Förintelsen blev möjlig *(Bloom 4)*
Eleven ska kunna förklara samband mellan ideologi, politiska beslut, byråkrati, propaganda och bystander-passivitet.
- E: enkla förklaringar
- C: förklaringar
- A: komplexa förklaringar ur flera perspektiv

**Bortvalda mål:** Mål 4 (källkritik på vittnesmål) och Mål 5 (efterspel/samtidskopplingar) — medvetet strukna för att hålla momentet smalt.

---

## § Kursminnespreferenser (Historia 1a1) — viktigast att respektera

Från `.claude/planera-moment/minne/historia-1a1.md`:

- **40-minuterslektioner med 4-fasstruktur** — INTE skillens default sex-fas
  - Fas 1: Introduktion (5 min) — retrieval practice + krok
  - Fas 2: Huvudinnehåll (15 min) — lärarledd genomgång med modellering
  - Fas 3: Elevaktivitet (15 min) — eleven arbetar
  - Fas 4: Summering (5 min) — exit ticket + preview
- **Retrieval practice är prio 1** — closed-book recall efter läsning
- **Lärarledda genomgångar med tydlig modellering** prioriteras
- **Differentiering inbyggd:** alla material i två versioner (standard + förenklad)
- **Frågor taggade med Bloom-nivå** + kopplade till E/C/A explicit
- **Frågedatabaser i CSV-format**
- **Eget designsystem** (se nedan)
- **Flashcards i specifikt format** (se nedan)
- **Konstruktiv alignment** (Blooms taxonomi, bedömningskriterier per uppgift)
- **Gränsfallsanalys** vid bedömning (C/A, E/C, F/E-gränser)
- **Förenklad version:** kortare meningar, vardagligt ordförråd, tydligare rubriker

---

## § Designsystem för Historia 1a1

| Element | Värde |
|---------|-------|
| Primärfärg (navy) | `#1B3A5C` |
| Accentfärg (guld) | `#C8A951` |
| Bakgrund | crème [VERIFIERA exakt hex genom att titta på tidigare moment-material i `Brain/` eller `Undervisningsmaterial/Historia/`] |
| Brödtext | Calibri 13pt med generös radavstånd |
| Rubriker | Georgia |
| Flashcards | Landskaps-A4, 2×2-rutnät, baksidor måste byta rader (rad 1↔rad 2), kolumner behålls. Motsvarar **långsides-flip**. Titelsida: "vänd på långsidan" |

**Innan generering:** sök efter tidigare moment för Historia 1a1 (t.ex. "Industriella revolutionen", "Världskrigen") för att finna exakta värden för crème och eventuella malldokument:
```bash
find "/home/anders/Second brain" -path "*/Historia/*" -name "*.docx" 2>/dev/null | head
find "/home/anders/Second brain" -path "*/Historia/*" -name "*.html" 2>/dev/null | head
ls /tmp/undervisningsmaterial/Historia/ 2>/dev/null  # eller klona repot
```

---

## § Lektionsstruktur (godkänd)

| # | Titel | Fas i Förintelsen | Hédis kontextpunkt | Mål | Begrepp som introduceras |
|---|-------|-------------------|---------------------|-----|--------------------------|
| 1 | *Hédis vardag — och hotet hon inte ser* | Bakgrund: nazism växer (1933–39) | Hédis barndom i Sighet, normal judisk vardag | 1, 2 | antisemitism, rasideologi, eugenik |
| 2 | *Tonåringen och kriget* | Krigets utbrott + uteslutning (1939–43) | Hédi som tonåring; kriget rasar runt Sighet | 1, 2 | ockupation, uteslutning, segregation |
| 3 | *Sighet stängs in* | Ghettofiering (april 1944) | Familjen Fried tvingas till Sighetghettot | 1, 2 | ghetto, deportation |
| 4 | *Tåget och rampen* | Deportation + ankomst Auschwitz (maj 1944) | Vagonsfärden, selektionen, föräldrarnas död | 1, 2 | selektion, "den slutliga lösningen", förintelseläger |
| 5 | *Att överleva* | Lägertillvaron (Auschwitz/Bergen-Belsen) | Hédi och Livia överlever lägervardagen | 1, 2 | koncentrationsläger, dehumanisering, Sonderkommando |
| 6 | *Hur blev det möjligt?* | **Analytisk topp** | Hédis egna frågor → strukturell analys | **3**, 2 | bystander, folkmord |
| 7 | *Vita bussarna, Sverige, vittnet* | Befrielse + efterspel | Räddningen, Sverige, vittnesmålet som arv | 1, 2 | vittnesmål |

**Exit tickets:**
- L1: *"Förklara med egna ord vad antisemitism är. Ge ett exempel från Tyskland 1933–1938."*
- L2: *"Vad innebar 'uteslutning' för judar i Europa 1939–1943? Ge två konkreta exempel."*
- L3: *"Beskriv vad ett ghetto var med hjälp av Hédis upplevelse av Sighetghettot."*
- L4: *"Förklara med egna ord vad 'selektion' var. Vad avgjorde vem som skickades direkt till gaskammaren?"*
- L5: *"Ge två konkreta exempel på hur lägren systematiskt dehumaniserade fångarna — och förklara varför."*
- L6: *"Förklara med egna ord hur 'bystander-passivitet' bidrog till att Förintelsen blev möjlig. Ge ett konkret exempel."*
- L7: *"En sak du tar med dig från det här momentet."* (reflekterande)

**Retrieval-slinga** — varje lektion börjar med closed-book retrieval på begrepp från föregående lektion. Datan från exit tickets sorteras i tre högar (förstod/osäker/missade) och styr nästa lektions intro.

**Etiskt tunga lektioner:** L4 och L5. Förbered eleverna redan i L3 inför detta. Lärarledd uppläsning är säkrare än självläsning. Avsluta inte med rusha exit ticket — ge utrymme för reaktioner.

---

## § Hédi Frieds berättelse — faktaunderlag

> **OBS:** Hédi Fried nämns inte i NotebookLM:s källor. All information nedan kommer från Claudes inbyggda kunskap och behöver verifieras innan presentation till elever. Använd hennes egna böcker som primärkälla där möjligt.

### Källmaterial att utgå från

- *Skärvor av ett liv* (1992) — hennes självbiografi, primärkällan för hela momentet
- *Frågor jag fått om Förintelsen* — hennes svarsbok som driver Lektion 6
- Filmade vittnesmål (sannolikt tillgängliga via Forum för levande historia / Levande historia) [VERIFIERA tillgänglighet och rättigheter för klassrumsvisning]
- Familjefoton (om reproducerbara)

### Biografiska nyckelfakta (alla med [VERIFIERA])

- **Född** 1924 i Sighet (Sighetu Marmației), Transsylvanien (då Rumänien) [VERIFIERA år]
- Sighet annekterades av **Ungern** 1940 efter Wienskiljedomen (Vienna Award) [VERIFIERA]
- **Mars 1944:** Tyskarnas ockupation av Ungern
- **April 1944:** Sighetghettot upprättas; familjen Fried tvingas dit [VERIFIERA exakt datum]
- **Maj 1944:** Deportation till Auschwitz-Birkenau med godsvagnstransport (cirka tre dagar) [VERIFIERA]
- **Selektionen vid rampen:** Hédis föräldrar mördas direkt i gaskammare. Hédi och hennes syster **Livia** kommer in i lägret som arbetsfångar [VERIFIERA systerns namn]
- **Senare 1944/1945:** Förflyttning till **Bergen-Belsen**
- **April 1945:** Befrielse av Bergen-Belsen [VERIFIERA datum]
- **Sommar 1945:** Räddad till Sverige genom **Folke Bernadottes Vita bussar** [VERIFIERA]
- **Liv i Sverige:** Bygger nytt liv, blir psykoterapeut, författare, klassrumsföreläsare i decennier
- **Död:** 2022, 102 år gammal [VERIFIERA exakt år och ålder]

### Centrala citat/passager att finna i hennes texter

För varje lektion behöver vi 1–2 citat från hennes verk. Föreslagna teman:
- L1: hennes barndom / familjefoton i Sighet före kriget
- L2: tonårsåren när lagarna börjar verka i Ungern
- L3: dagen ghettoiseringen — vad familjen lämnade
- L4: vagonsfärden + ankomsten till rampen + ögonblicket föräldrarna försvinner
- L5: lägervardagen, hungern, kylan, hur hon och Livia höll ihop
- L6: hennes egna frågor från *Frågor jag fått om Förintelsen*
- L7: räddningen, Sverige, vad hon vill att vi ska minnas

**Generera kontentpunkter inte ordagrant återgivning** — undvik upphovsrättsproblem. Citat bör vara korta (1–3 meningar) och tydligt attribuerade.

---

## § Etiska riktlinjer (Yad Vashems princip)

- **From victim to person:** människor som de levde, inte främst hur de dog
- **Familjefoton, brev, vardagsobjekt** framför grafiska bilder
- **Inga lik-/lägerbilder utan tydlig pedagogisk avsikt och förvarning**
- **Bystander-perspektivet är pedagogiskt centralt** — *hur blev det möjligt?* är hjärtat i Mål 3
- **Identitet i klassen:** om eleverna har judisk, romsk, polsk eller annan direkt familjekoppling, krävs hänsyn. Fråga läraren innan generering om sådant gäller den specifika gruppen.
- **Tunga lektioner (L4, L5):** förbered eleverna i förväg, gör tydliga övergångar, ge tid för reaktioner

---

## § Genereringsuppgifter (i ordning)

Generera ETT material i taget. Be läraren godkänna varje innan du fortsätter.

### Uppgift 1: 7 lektionsplaner

För varje lektion enligt 4-fasstrukturen (40 min):

**Format per lektionsplan:**
- Heading 1: "Lektion N: [Titel]"
- Lärandemål för lektionen (vilka av momentets 3 mål)
- Förberedelse (vad läraren behöver i förväg)
- Retrieval review-koppling (baserat på föregående lektions exit ticket)
- Tidsplanering — tabell med Tid / Fas / Aktivitet / Beskrivning
- Lärarinstruktioner (nyckelformuleringar, modellering, kontrollfrågor)
- Elevaktiviteter (steg-för-steg)
- Differentiering (konkret stöd mot E + utmaning mot A)
- Material som behövs
- Exit ticket (fråga + hur datan informerar nästa lektion)
- Koppling till kunskapskrav (E/C/A)

**Spara två format per lektion:**
- Markdown: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/lektion-N.md`
- Word (.docx): `/home/anders/undervisningsmaterial/Historia/Andra världskriget och Förintelsen/lektion-N.docx`

**Word-specifikation** (per kursminnet och `.claude/skills/docx/SKILL.md`):
- A4 (11906 × 16838 DXA), 1″-marginaler
- Rubriker som Heading 1/Heading 2 (Georgia om möjligt, annars Calibri bold)
- Brödtext Calibri 13pt
- Tidsplaneringen som riktig tabell (DXA, ej percentage)
- Bullets med `LevelFormat.BULLET`
- Sidfot: "Historia 1a1 · Andra världskriget och Förintelsen · Lektion N"

**Kvalitetskontroll innan presentation:**
- Retrieval-koppling till föregående lektion finns?
- Elevaktiv tid > 50%?
- Differentieringen konkret (inte "stöd svagare elever")?
- [VERIFIERA]-taggar vid osäkra fakta?
- Kopplar framåt till nästa lektion?
- 4-fasstrukturen följs?

### Uppgift 2: 7 elevuppgifter (standard + förenklad)

Per kursminnet: alla material i två versioner.

För varje lektion, generera:
- `elevuppgift-lektion-N-standard.md` / `.docx`
- `elevuppgift-lektion-N-forenklad.md` / `.docx`

**Standardversion:**
- Tydlig rubrik kopplad till lektionen
- Kort kontextsättning (vad ska eleven göra och varför)
- Numrerade frågor/instruktioner
- Differentieringsstöd inbyggt (stödrutor "Tips:..." för E, fördjupningsfrågor markerade ★ för A)
- Citat/utdrag från Hédis texter direkt i dokumentet (där relevant)
- Fält för elevsvar med utrymme

**Förenklad version:**
- Kortare meningar (max 15 ord rikt schema)
- Vardagligt ordförråd (förklara facktermer i parentes)
- Tydligare rubriker
- Fler bilder/visuella stöd där möjligt
- Eventuellt punktlistor istället för löpande text
- Samma uppgift-innehåll, samma frågor — bara språkligt förenklat
- DESIGNAT FÖR ELEVER MED LÄSFÖRSTÅELSESVÅRIGHETER (inte sänkt analytisk nivå)

**Spara i båda format** (md i vault, docx utanför).

### Uppgift 3: Frågedatabas (CSV)

Generera en CSV-fil med ~50–80 frågor som täcker alla 7 lektioner.

**Format:**
```csv
fråga,svar,bloom_nivå,e_c_a,lektion,begrepp,kategori
"Vad var Sighetghettot?","Ett ghetto i Sighet (då Ungern) som upprättades april 1944 där judar tvingades bo isolerade. Familjen Fried hörde dit.",2,E,3,ghetto,faktaåtergivning
```

**Kolumner:**
- fråga
- svar (kort, faktarätt)
- bloom_nivå (1-6)
- e_c_a (E / C / A)
- lektion (1-7)
- begrepp (vilket av de 11 nyckelbegreppen som adresseras, eller tom)
- kategori (faktaåtergivning / begrepp / analys / källanalys)

**Fördelning:**
- Mål 1 (redogöra): ~30 frågor, mest E-C
- Mål 2 (begrepp): ~30 frågor, både E (definition) och A (nyans)
- Mål 3 (samband): ~10–20 frågor, främst C-A

Spara: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/fragedatabas.csv`

### Uppgift 4: 7 presentationer (Historia 1a1-designsystem)

⚠ **INTE Arkiv v2.1** — använd kursens eget designsystem (navy/guld/crème, Calibri/Georgia).

Format: HTML med reveal.js eller PowerPoint (.pptx) — välj det som matchar tidigare moment i kursen. Sök efter tidigare presentationer i Historia 1a1 för att se vilket format som används:

```bash
find "/home/anders/Second brain" -path "*/Historia/*" -name "*.pptx" 2>/dev/null | head
find "/home/anders/Second brain" -path "*/Historia/*" -name "presentation*" 2>/dev/null | head
```

**Innehåll per presentation (~10–14 slides per 40 min-lektion):**
- Titelsida med lektion N + tema
- Lektionens mål (elevvänligt språk)
- Krok från Hédis text
- Innehållsslides: faktagenomgång + Hédis citat omväxlande
- Begreppsslides (en per nyckelbegrepp som introduceras i lektionen)
- Diskussionsfråga eller EPA-slide
- Avslutsslide med exit ticket-fråga + preview nästa lektion
- Talarnoter på varje slide

**Designprinciper:**
- Navy `#1B3A5C` som primärfärg
- Guld `#C8A951` som accent
- Crème bakgrund [VERIFIERA hex]
- Calibri 13pt brödtext (storlek upp för slidesvy: minst 24pt)
- Georgia rubriker
- **INGA emojis** — typografiska tecken (▸ ● ▪ §) får användas
- Familjefoton + kartor framför grafiska bilder
- Källhänvisningar på fakta-slides

Spara: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/presentation-lektion-N.html` (eller `.pptx` om kurskonventionen är så)

### Uppgift 5: Skriftligt prov

Genereras efter lektionsplanerna men *före* momentöversikten.

**Provspecifikation:**
- **Tid:** 60 minuter (matchar Historia 1a1-traditionen om 40-min-lektioner = 1,5 lektion)
- **Format:** Pappersprov (utskrivbart Word-dokument)
- **Tre delar:**

#### Del A — Begrepp (Mål 2) — 6 frågor × 2 poäng = 12 p
Korta begreppsdefinitioner. Eleven får 6 av de 11 nyckelbegreppen att förklara med egna ord.
*Exempel:* "Förklara med egna ord vad ett *ghetto* var. Ge ett exempel från Förintelsen."

#### Del B — Faktaåtergivning (Mål 1) — 4 frågor × 3 poäng = 12 p
Korta beskrivningar. Eleven återger Förintelsens stadier eller Hédis öde.
*Exempel:* "Beskriv kortfattat vad som hände vid den så kallade *selektionen* när tågen anlände till Auschwitz."

#### Del C — Analys (Mål 3) — 1 essäfråga × 12 poäng = 12 p
**Detta är hjärtat i provet.** Eleven förklarar samband.
*Förslag på fråga:* "Förklara med stöd från Hédi Frieds berättelse hur Förintelsen blev möjlig. Behandla minst tre av följande faktorer: nazistisk rasideologi, antijudiska lagar, byråkrati, propaganda, bystander-passivitet. Skriv ungefär en A4-sida."

**Total:** 36 poäng

**Betygsgränser** [förslag, justera enligt skolans norm]:
- E: 14 p (39%)
- C: 22 p (61%)
- A: 30 p (83%)

**Bedömningsanvisningar (för läraren):**
- Per fråga: explicit E/C/A-kriterium
- För Del C: rubrikbedömning med kriterier för enkel/utförlig/utförlig och nyanserad redogörelse + enkel/komplex förklaring av samband
- Konstruktiv alignment: poängen kopplas direkt till Bloom-nivå (Del A = Bloom 2-3, Del B = Bloom 2, Del C = Bloom 4)

**Spara:**
- Provet: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/prov.md` + `.docx`
- Bedömningsanvisningar: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/prov-bedomningsanvisningar.md` + `.docx`
- Facit/exempel-svar för läraren: i bedömningsanvisningarna

### Uppgift 6: Momentöversikt för elever (HTML)

Self-contained HTML-fil enligt `.claude/skills/html-momentoversikt/SKILL.md`.

**Anpassa designen till Historia 1a1-paletten** — navy/guld/crème, Georgia rubriker, Calibri brödtext. INTE skillens default historiapalett.

**Innehåll:**
- Header: "Andra världskriget och Förintelsen — Historia 1a1"
- Kort beskrivning (2–3 meningar)
- Tidsperiod (be läraren om datum för varje lektion)
- Lärandemålen i **elevvänligt språk** (inte E/C/A-formuleringar)
- 7 lektionskort: nummer, titel, datum, kort innehållsbeskrivning, ev. förberedelser
- Provet som sista "kort" med datum
- Footer

**Innan generering, fråga läraren:**
1. Datum för varje lektion + provet
2. Förberedelser eleverna behöver göra (om någon)
3. Lärarens namn (om det ska visas)
4. Övrigt meddelande (om något)

Spara: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/momentoversikt.html`

### Uppgift 7: Flashcards

⚠ **Specifikt format per kursminne** — landskaps-A4, 2×2-rutnät, baksidor måste byta rader (rad 1↔rad 2 byter plats medan kolumner behålls). Motsvarar **långsides-flip** ("vänd på långsidan"). Titelsida ska ange "vänd på långsidan".

**Innehåll:**
- Titelsida (ange "vänd på långsidan")
- ~40 kort med nyckelbegrepp + datum/personer/platser
- Framsida: en tydlig fråga (aldrig sammansatt)
- Baksida: ett kort svar
- Designsystem: navy bakgrund eller crème-papper med navy/guld element

**Sök efter tidigare flashcards** för att se exakt format:
```bash
find "/home/anders/Second brain" -path "*/Historia/*" -iname "*flashcard*" 2>/dev/null
find "/home/anders/Second brain/resources" -iname "*flashcard*" 2>/dev/null
```

Spara: `Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/flashcards.docx` (eller `.pdf`)

### Avslutning: Uppdatera kursminnet

När alla material är godkända, analysera konversationen för:
- Vad läraren ändrade i förslagen
- Vad hen lade till eller tog bort
- Mönster i justeringar (preferenser för specifika lektionstyper, materialstilar, etc.)
- Om något nytt om gruppen kom fram

Uppdatera `.claude/planera-moment/minne/historia-1a1.md`:
- Lägg till en rad i historiken med datum, momentnamn, nyckellärdom
- Eventuellt nya pedagogiska preferenser eller bortvalda metoder

Skriv en kort sammanfattning till läraren om vad som sparades.

---

## § Filsökvägar — sammanfattning

**I vaultet** (`Undervisningsmaterial/Historia/Andra världskriget och Förintelsen/`):
- `momentplan.md` ✅ klar
- `generationsplan.md` ✅ denna fil
- `lektion-N.md` (1–7) — Uppgift 1
- `elevuppgift-lektion-N-standard.md` (1–7) — Uppgift 2
- `elevuppgift-lektion-N-forenklad.md` (1–7) — Uppgift 2
- `fragedatabas.csv` — Uppgift 3
- `presentation-lektion-N.html` (1–7) — Uppgift 4
- `prov.md` — Uppgift 5
- `prov-bedomningsanvisningar.md` — Uppgift 5
- `momentoversikt.html` — Uppgift 6
- `flashcards.docx` (eller `.pdf`) — Uppgift 7

**Utanför vaultet** (`/home/anders/undervisningsmaterial/Historia/Andra världskriget och Förintelsen/`):
- `lektion-N.docx` (1–7)
- `elevuppgift-lektion-N-standard.docx` (1–7)
- `elevuppgift-lektion-N-forenklad.docx` (1–7)
- `prov.docx`
- `prov-bedomningsanvisningar.docx`

**Båda katalogerna är redan skapade.**

---

## § NotebookLM-användning under generering

För varje lektion, hämta innehåll från notebooken med specifika frågor:

```bash
notebooklm ask --json "Ge mig fakta, nyckelbegrepp och konkreta exempel om [lektionens specifika tema]. Inkludera kallhanvisningar."
notebooklm ask --json "Vilka vanliga missforstand eller svarigheter finns kring [temat]?"
notebooklm ask --json "Finns det primorkallor eller citat om [temat] som kan anvandas i undervisning?"
```

**Använd notebookens innehåll för:**
- Fas 2 (lärarledd genomgång) — fakta och kontextuell information
- Fas 3 (elevaktivitet) — eventuella primärkällor
- Differentiering (mot A) — fördjupningsmaterial

**Hédi-specifikt material:** notebooken saknar det. Använd Claudes inbyggda kunskap + flagga `[VERIFIERA]`. Bekräfta detaljer via offentligt tillgängliga källor om möjligt.

**Primo Levi som komplement:** Levi är dokumenterad i notebooken som biografi-spalt i *A History of World Societies*. Kan användas i lektion 5 (lägertillvaron) eller lektion 6 (analysen, *"the gray zone"*) som utmaning mot A-nivå.

---

## § Status (vid sessionsslut 2026-05-07)

✅ Steg 1: Ämne, kurs, vinkel  
✅ Steg 2: Lärandemål (3 mål, mål 4-5 strukna)  
✅ Steg 3: Pedagogiskt upplägg (Ansats A — narrativ ankarpedagogik)  
✅ Steg 4: Lektionsstruktur + bedömning (prov)  
🔲 Generering — denna fil

---

## § Notering till nästa session

Skillen `planera-moment` har ett kursminne i `.claude/planera-moment/minne/historia-1a1.md`. Det innehåller många preferenser som denna generationsplan redan tagit hänsyn till. **Läs det också** — det kan finnas detaljer (t.ex. exakt designsystem-värden, tidigare materialformat) som inte hunnit destilleras hit.

Sök även efter tidigare moment i Historia 1a1 (t.ex. "Världskrigen", "Industriella revolutionen") för att se exakta layoutmallar.

Om något i denna plan är otydligt — fråga läraren innan du gissar.
