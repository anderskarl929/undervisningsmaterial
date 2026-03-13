const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak
} = require("docx");
const pptxgen = require("pptxgenjs");

// ============================================================
// HELPER FUNCTIONS (identiska med generate-lektion-1.js)
// ============================================================

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

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 24 })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true })],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true })],
  });
}

function bodyText(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24 })],
  });
}

function italicText(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24, italics: true })],
  });
}

function boldBodyText(label, text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, font: "Arial", size: 24, bold: true }),
      new TextRun({ text, font: "Arial", size: 24 }),
    ],
  });
}

function boldItalicBody(label, text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, font: "Arial", size: 24, bold: true }),
      new TextRun({ text, font: "Arial", size: 24, italics: true }),
    ],
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 60 }, children: [] });
}

// ============================================================
// DOCX - Lektion 3: AI och konspirationsteorier
// ============================================================

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1A1A2E" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Samh\u00E4llskunskap 3 / Int. relationer \u2014 K\u00E4llkritik: AI och konspirationsteorier", font: "Arial", size: 18, color: "888888" })],
        })],
      }),
    },
    children: [
      // ========== TITEL ==========
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lektion 3: AI och konspirationsteorier", font: "Arial" })],
      }),
      bodyText("N\u00E4r AI f\u00F6rst\u00E4rker konspirationsteoriers spridning \u2014 stationsarbete med k\u00E4llkritisk analys"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // ========== L\u00C4RANDEM\u00C5L ==========
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna analysera hur AI kan anv\u00E4ndas f\u00F6r att skapa och sprida konspirationsteorier, samt granska AI-genererat inneh\u00E5ll med hj\u00E4lp av k\u00E4llkritiska verktyg."),
      spacer(),
      bullet("Till\u00E4mpa k\u00E4llkritiska verktyg (grundfr\u00E5gorna och SIFT) p\u00E5 AI-genererat inneh\u00E5ll (m\u00E5l 1)"),
      bullet("Analysera hur AI f\u00F6rst\u00E4rker konspirationsteoriers spridning genom specifika mekanismer (m\u00E5l 2)"),
      bullet("Referera till k\u00E4llor med Harvard-referering (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men och f\u00F6r enkla resonemang om hur AI kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men och f\u00F6r v\u00E4lgrundade resonemang om sambandet mellan AI-mekanismer och konspirationsteoriers spridning."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men och f\u00F6r v\u00E4lgrundade och nyanserade resonemang om komplexa samband mellan AI, konspirationsteorier och demokrati."),

      // ========== CENTRALT INNEH\u00C5LL ==========
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("Hur propaganda, desinformation och konspirationsteorier fungerar som p\u00E5verkansformer i det digitala medielandskapet."),

      // ========== F\u00D6RBEREDELSE ==========
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 2 \u2014 notera vilka elever som beh\u00E4rskar grundfr\u00E5gorna och SIFT"),
      bullet("F\u00F6rbered matching-\u00F6vning f\u00F6r retrieval review (grundfr\u00E5ga \u2194 SIFT-steg)"),
      bullet("Skriv ut stationsblad (bilaga 1\u20134) \u2014 ett set per pargrupp"),
      bullet("Skriv ut gateway exit ticket (bilaga 5) \u2014 en per elev"),
      bullet("Skriv ut gateway checklista (bilaga 6) \u2014 f\u00F6r l\u00E4raren"),
      bullet("F\u00F6rbered timer f\u00F6r stationsrotation (8 min per station)"),
      bullet("Ha l\u00E4randem\u00E5len synliga p\u00E5 tavlan under hela lektionen"),

      // ========== TIDSPLANERING ==========
      heading2("Tidsplanering"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1000, 1300, 2200, 4526],
        rows: [
          new TableRow({
            children: [
              headerCell("Tid", 1000),
              headerCell("Fas", 1300),
              headerCell("Aktivitet", 2200),
              headerCell("Beskrivning", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("0\u20138 min", 1000),
              cell("Retrieval review", 1300),
              cell("Matching: grundfr\u00E5gor \u2194 SIFT", 2200),
              cell("Eleverna parar ihop grundfr\u00E5gor med r\u00E4tt steg i SIFT-metoden (t.ex. VEM? \u2194 Investigate the source). Kort genomg\u00E5ng. Baseras p\u00E5 L2 exit ticket.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201318 min", 1000),
              cell("Instruktion", 1300),
              cell("AI-mekanismer: \u00F6versikt", 2200),
              cell("Kort genomg\u00E5ng av 4 mekanismer: deepfakes, filterbubbla/ekkammare, AI-genererade k\u00E4llor, automatiserade bottar. F\u00F6r varje mekanism: principen och koppling till SIFT/grundfr\u00E5gorna. Tredjepersons-framing p\u00E5minnelse.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("18\u201353 min", 1000),
              cell("Stationsarbete", 1300),
              cell("4 stationer, 8 min/station", 2200),
              cell("Eleverna arbetar i PAR och roterar mellan 4 stationer (deepfakes, filterbubbla, AI-genererade k\u00E4llor, bottar). Varje station: bakgrundstext, case, SIFT/grundfr\u00E5ge-koppling, reflektionsfraga. 3 min f\u00F6r f\u00F6rflyttning/instruktion.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("53\u201363 min", 1000),
              cell("Summering", 1300),
              cell("Helklassdiskussion", 2200),
              cell("Diskussion: Vilka m\u00F6nster s\u00E5g ni? Vilken mekanism var sv\u00E5rast att genomsk\u00E5da? Vilket k\u00E4llkritiskt verktyg var mest anv\u00E4ndbart? L\u00E4raren sammanfattar.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("63\u201375 min", 1000),
              cell("Gateway exit ticket", 1300),
              cell("Tv\u00E5 delar: tillampning + Harvard", 2200),
              cell("Del 1 (7 min): Nytt exempel, k\u00E4llkritisk till\u00E4mpning UTAN scaffolding. Del 2 (5 min): Harvard-referering av 2\u20133 k\u00E4llor. Exit ticket-fr\u00E5ga fr\u00E5n exit-ticket-mall.md.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("75\u201380 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("Preview av L4", 2200),
              cell("N\u00E4sta lektion: seminarium. Grupperna baseras p\u00E5 dagens exit ticket. Kort f\u00F6rklaring av seminarieformatet.", 4526),
            ],
          }),
        ],
      }),

      // ========== L\u00C4RARINSTRUKTIONER ==========
      heading2("L\u00E4rarinstruktioner"),

      // --- Fas 1: Retrieval review ---
      heading3("Retrieval review (0\u20138 min)"),
      boldBodyText("Matching-\u00F6vning (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CF\u00F6rra lektionen l\u00E4rde ni er SIFT-metoden och de fyra k\u00E4llkritiska grundfr\u00E5gorna. Nu ska vi se hur de h\u00E4nger ihop. P\u00E5 tavlan ser ni grundfr\u00E5gorna i en kolumn och SIFT-stegen i en annan \u2014 men i oordning. Para ihop r\u00E4tt grundfr\u00E5ga med det SIFT-steg som b\u00E4st matchar. Ni har 2 minuter, arbeta i par.\u201D"),
      spacer(),
      bodyText("Skriv p\u00E5 tavlan (matchnings\u00F6vning):"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4513, 4513],
        rows: [
          new TableRow({
            children: [
              headerCell("Grundfr\u00E5ga", 4513),
              headerCell("SIFT-steg (i oordning)", 4513),
            ],
          }),
          new TableRow({ children: [cell("VEM? (avs\u00E4ndare)", 4513), cell("Trace claims", 4513)] }),
          new TableRow({ children: [cell("VARF\u00D6R? (syfte)", 4513), cell("Stop", 4513)] }),
          new TableRow({ children: [cell("HUR? (presentation)", 4513), cell("Find better coverage", 4513)] }),
          new TableRow({ children: [cell("N\u00C4R? (aktualitet)", 4513), cell("Investigate the source", 4513)] }),
        ],
      }),

      spacer(),
      bodyText("R\u00E4tt matchning (visa efter 2 min):"),
      bullet("VEM? \u2194 Investigate the source \u2014 b\u00E5da handlar om att ta reda p\u00E5 vem avs\u00E4ndaren \u00E4r"),
      bullet("VARF\u00D6R? \u2194 Stop \u2014 stanna och t\u00E4nk: vad \u00E4r syftet med det h\u00E4r?"),
      bullet("HUR? \u2194 Find better coverage \u2014 j\u00E4mf\u00F6r hur samma sak presenteras av andra k\u00E4llor"),
      bullet("N\u00C4R? \u2194 Trace claims \u2014 sp\u00E5ra p\u00E5st\u00E5endet tillbaka till prim\u00E4rk\u00E4llan, n\u00E4r uppstod det?"),
      spacer(),
      boldBodyText("Genomg\u00E5ng (5\u20138 min): ", "G\u00E5 igenom matchningarna. Betona att kopplingar inte \u00E4r perfekta 1:1 \u2014 men att grundfr\u00E5gorna och SIFT f\u00F6rst\u00E4rker varandra."),

      // --- Fas 2: Instruktion ---
      heading3("Instruktion: AI-mekanismer (8\u201318 min)"),
      italicText("S\u00E4g: \u201CF\u00F6rra lektionen fick ni verktygen. Idag ska ni anv\u00E4nda dem p\u00E5 riktigt \u2014 p\u00E5 fyra s\u00E4tt som AI f\u00F6rst\u00E4rker konspirationsteoriers spridning. Vid varje station kommer ni att anv\u00E4nda SIFT och grundfr\u00E5gorna f\u00F6r att granska ett konkret exempel.\u201D"),
      spacer(),
      bodyText("Presentera de fyra mekanismerna kort (2\u20133 meningar per mekanism):"),
      spacer(),

      boldBodyText("1. Deepfakes: ", "AI kan skapa realistiska bilder, videor och r\u00F6ster av verkliga personer. Deepfakes kan anv\u00E4ndas f\u00F6r att l\u00E4gga ord i politikers mun eller fabricera h\u00E4ndelser som aldrig \u00E4gt rum. Sv\u00E5rt att avsl\u00F6ja med blotta \u00F6gat."),
      boldBodyText("2. Filterbubbla/ekkammare: ", "Algoritmerna p\u00E5 sociala medier visar inneh\u00E5ll som matchar det du redan tittat p\u00E5. Om du s\u00F6ker p\u00E5 en konspirationsteori en g\u00E5ng f\u00E5r du mer liknande inneh\u00E5ll \u2014 en sj\u00E4lvf\u00F6rst\u00E4rkande loop."),
      boldBodyText("3. AI-genererade k\u00E4llor: ", "Det finns \u00F6ver 2 000 fejkade nyhetssajter som drivs helt av AI. De ser ut som riktig journalistik men har ingen redaktion och inga riktiga journalister."),
      boldBodyText("4. Automatiserade bottar: ", "N\u00E4tverk av sociala medie-konton som styrs av AI och sprider samma budskap i stor skala. Skapar en illusion av bred konsensus."),
      spacer(),

      boldItalicBody("Tredjepersons-framing \u2014 f\u00F6rst\u00E4rkt p\u00E5minnelse: ", "\u201CKom ih\u00E5g v\u00E5r analytiska norm: vi unders\u00F6ker varf\u00F6r n\u00E5gon KAN finna n\u00E5got \u00F6vertygande \u2014 vi analyserar inte om NI tror p\u00E5 det. Det \u00E4r viktigt idag, eftersom vi arbetar med inneh\u00E5ll som \u00E4r designat f\u00F6r att vara sv\u00E5rt att genomsk\u00E5da.\u201D"),

      // --- Fas 3: Stationsarbete ---
      heading3("Stationsarbete (18\u201353 min)"),
      bodyText("Eleverna arbetar i PAR (inte st\u00F6rre grupper \u2014 varje elev m\u00E5ste vara aktiv). 4 stationer, 8 minuter per station + 3 minuter f\u00F6r f\u00F6rflyttning och instruktion."),
      spacer(),
      italicText("Instruktion: \u201CNi ska rotera mellan fyra stationer. Vid varje station finns en bakgrundstext om en mekanism och ett konkret exempel. Er uppgift \u00E4r att (a) f\u00F6rst\u00E5 mekanismen, (b) till\u00E4mpa SIFT och/eller grundfr\u00E5gorna p\u00E5 exemplet, och (c) besvara reflektionsfr\u00E5gan. Skriv ner era svar kort \u2014 ni beh\u00F6ver dem i summeringen. Ni har 8 minuter per station. Jag signalerar n\u00E4r det \u00E4r dags att byta.\u201D"),
      spacer(),
      bodyText("Stationsbladen finns i bilaga 1\u20134. Dela ut ett set per pargrupp."),
      spacer(),
      bodyText("Cirkulera under stationsarbetet. St\u00E4ll f\u00F6rdjupande fr\u00E5gor:"),
      bullet("Om paret g\u00E5r f\u00F6r snabbt: \u201CVilket SIFT-steg anv\u00E4nde ni? Varf\u00F6r just det?\u201D"),
      bullet("Om paret fastnar: \u201CB\u00F6rja med att l\u00E4sa bakgrundstexten. Vilken grundfr\u00E5ga \u00E4r mest relevant h\u00E4r?\u201D"),
      bullet("Utmaning: \u201CKunde den h\u00E4r mekanismen ha fungerat lika bra utan AI? Varf\u00F6r/varf\u00F6r inte?\u201D"),

      // --- Fas 4: Summering ---
      heading3("Summering (53\u201363 min)"),
      bodyText("Helklassdiskussion. Skriv p\u00E5 tavlan:"),
      spacer(),
      bullet("\u201CVilka m\u00F6nster s\u00E5g ni \u00F6ver stationerna?\u201D"),
      bullet("\u201CVilken mekanism var sv\u00E5rast att genomsk\u00E5da? Varf\u00F6r?\u201D"),
      bullet("\u201CVilket k\u00E4llkritiskt verktyg \u2014 grundfr\u00E5gorna eller SIFT \u2014 var mest anv\u00E4ndbart? Skilde det sig mellan stationerna?\u201D"),
      spacer(),
      italicText("Sammanfatta: \u201CAlla fyra mekanismer f\u00F6rst\u00E4rks av AI men kan granskas med samma verktyg. Deepfakes avsl\u00F6jas genom lateral reading och omv\u00E4nd bilds\u00F6kning. Filterbubblor motverkas genom att medvetet s\u00F6ka andra perspektiv. AI-genererade k\u00E4llor avsl\u00F6jas genom att unders\u00F6ka avs\u00E4ndaren. Bottar k\u00E4nns igen genom att sp\u00E5ra varifrn\u00E5n p\u00E5st\u00E5enden kommer.\u201D"),

      // --- Fas 5: Gateway exit ticket ---
      heading3("Gateway exit ticket (63\u201375 min)"),
      bodyText("Exit ticket best\u00E5r av tv\u00E5 delar. Detta \u00E4r en gateway check \u2014 resultaten anv\u00E4nds f\u00F6r att best\u00E4mma gruppsammans\u00E4ttning inf\u00F6r L4-seminariet."),
      spacer(),

      boldBodyText("Del 1: K\u00E4llkritisk till\u00E4mpning (7 min)", ""),
      bodyText("Eleverna f\u00E5r ett NYTT exempel som inte anv\u00E4nts vid stationerna:"),
      spacer(),
      italicText("Exempel: Ett AI-genererat socialt medie-inl\u00E4gg p\u00E5 X (Twitter) med texten: \u201CBREAKING: L\u00E4ckt dokument visar att EU planerar att f\u00F6rbjuda kontanter senast 2028 f\u00F6r att kontrollera medborgarna. Dela innan de tar bort det! #KontantUpproret\u201D \u2014 inl\u00E4gget har 12 000 delningar, ett AI-genererat foto av ett \u201Cofficiellt\u201D EU-dokument, och kontot skapades f\u00F6r 3 veckor sedan."),
      spacer(),
      bodyText("Uppgift: Granska inl\u00E4gget med hj\u00E4lp av de k\u00E4llkritiska verktyg du l\u00E4rt dig. Inga meningsstartare, inga ledfr\u00E5gar."),
      spacer(),

      boldBodyText("Del 2: Harvard-referering (5 min)", ""),
      bodyText("Formatera f\u00F6ljande k\u00E4llor korrekt med Harvardsystemet (b\u00E5de i-text-h\u00E4nvisning och referenslista):"),
      bullet("K\u00E4lla 1: En artikel av NewsGuard fr\u00E5n 2025 med titeln \u201CTracking AI-enabled Misinformation\u201D"),
      bullet("K\u00E4lla 2: En rapport fr\u00E5n Myndigheten f\u00F6r psykologiskt f\u00F6rsvar (MPF) fr\u00E5n 2025 med titeln \u201CAI och informationsp\u00E5verkan\u201D"),
      spacer(),

      boldBodyText("Exit ticket-fr\u00E5ga: ", "\u201CVarf\u00F6r kan AI-genererade konspirationsteorier vara sv\u00E5rare att granska k\u00E4llkritiskt \u00E4n traditionella? Resonera utifr\u00E5n minst en mekanism vi diskuterat idag.\u201D"),
      spacer(),

      bodyText("Samla in. Bed\u00F6m med hj\u00E4lp av checklistan i bilaga 6."),

      // --- Fas 6: Fram\u00E5tkoppling ---
      heading3("Fram\u00E5tkoppling (75\u201380 min)"),
      italicText("S\u00E4g: \u201CN\u00E4sta lektion: seminarium d\u00E4r ni till\u00E4mpar allt ni l\u00E4rt er. Grupperna baseras p\u00E5 dagens exit ticket \u2014 s\u00E5 g\u00F6r ert b\u00E4sta. I seminariet kommer ni att diskutera en fr\u00E5ga om AI och demokrati med tilldelade perspektiv. Ni f\u00E5r veta era grupper och perspektiv i b\u00F6rjan av n\u00E4sta lektion.\u201D"),

      // ========== ELEVAKTIVITETER ==========
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval review: matching grundfr\u00E5gor \u2194 SIFT-steg (par, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Stationsarbete: 4 stationer med k\u00E4llkritisk analys (par, 35 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Helklassdiskussion: m\u00F6nster och j\u00E4mf\u00F6relser (helklass, 10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Gateway exit ticket: sj\u00E4lvst\u00E4ndig till\u00E4mpning + Harvard-referering (enskilt, 12 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 62 av 80 minuter (78%)"),

      // ========== DIFFERENTIERING ==========
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Stationsbladen har tydliga instruktioner och bakgrundstexter som ger kontext. Under stationsarbetet hj\u00E4lper l\u00E4raren par som fastnar med att identifiera relevant grundfr\u00E5ga. Tredjepersons-framing g\u00F6r det tryggare att analysera k\u00E4nsligt inneh\u00E5ll. OBS: Gateway exit ticket har INGEN scaffolding \u2014 detta \u00E4r avsiktligt f\u00F6r att m\u00E4ta sj\u00E4lvst\u00E4ndig f\u00F6rm\u00E5ga."),
      boldBodyText("Utmaning (mot A): ", "Reflektionsfr\u00E5gorna vid varje station kr\u00E4ver djupare analys av varf\u00F6r mekanismen fungerar. Under summeringen: \u201CKunde dessa mekanismer ha fungerat lika bra f\u00F6re AI? Vad har f\u00F6r\u00E4ndrats?\u201D I exit ticket: uppmana till resonemang om flera mekanismer och demokratiska konsekvenser."),

      // ========== MATERIAL ==========
      heading2("Material"),
      bullet("Exit ticket-data fr\u00E5n lektion 2 (f\u00F6r retrieval review)"),
      bullet("Matchnings\u00F6vning p\u00E5 tavlan (grundfr\u00E5gor \u2194 SIFT)"),
      bullet("Stationsblad 1\u20134 (bilagor 1\u20134) \u2014 utskrivna, ett set per pargrupp"),
      bullet("Gateway exit ticket (bilaga 5) \u2014 en per elev"),
      bullet("Gateway checklista (bilaga 6) \u2014 f\u00F6r l\u00E4raren"),
      bullet("Timer f\u00F6r stationsrotation"),

      // ========== KOPPLING TILL KUNSKAPSKRAV ==========
      heading2("Koppling till kunskapskrav"),
      bullet("Retrieval review repeterar och f\u00F6rst\u00E4rker grundfr\u00E5gor och SIFT fr\u00E5n L2 (m\u00E5l 1)"),
      bullet("Stationsarbetet tr\u00E4nar till\u00E4mpning av k\u00E4llkritiska verktyg p\u00E5 AI-specifikt inneh\u00E5ll (m\u00E5l 1, 2: E\u2013A)"),
      bullet("Reflektionsfr\u00E5gorna med tredjepersons-framing tr\u00E4nar analytisk distans (m\u00E5l 2, 3)"),
      bullet("Gateway exit ticket m\u00E4ter sj\u00E4lvst\u00E4ndig k\u00E4llkritisk f\u00F6rm\u00E5ga UTAN scaffolding (m\u00E5l 1, 2)"),
      bullet("Harvard-referering i exit ticket bygger vidare p\u00E5 L2:s introduktion (m\u00E5l 4)"),

      // ========== KOPPLINGAR ==========
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1 (\u00E5terkommande case), Lektion 2 (SIFT, grundfr\u00E5gor, lateral reading, Harvard-referering)"),
      bullet("Retrieval review baseras p\u00E5 L2 exit ticket (grundfr\u00E5gorna och SIFT)"),
      bullet("Gateway check: exit ticket avg\u00F6r elevgruppering inf\u00F6r L4-seminariet"),
      bullet("N\u00E4sta lektion: L4 \u2014 formativt seminarium med tilldelade perspektiv"),

      // ===== SIDBRYTNING F\u00D6R BILAGOR =====
      new Paragraph({ children: [new PageBreak()] }),

      // ========== BILAGA 1: Station 1 - Deepfakes ==========
      heading2("Bilaga 1: Stationsblad \u2014 Station 1: Deepfakes"),
      spacer(),
      heading3("Bakgrund"),
      bodyText("Deepfakes \u00E4r AI-genererade bilder, videor eller r\u00F6ster som ser ut och l\u00E5ter som verkliga personer. Tekniken bygger p\u00E5 maskininl\u00E4rning och kan idag skapa inneh\u00E5ll som \u00E4r n\u00E4stan om\u00F6jligt att skilja fr\u00E5n \u00E4kta material med blotta \u00F6gat."),
      bodyText("Deepfakes anv\u00E4nds f\u00F6r att fabricera h\u00E4ndelser, l\u00E4gga ord i politikers mun eller skapa falska bevis. Spridningen sker framf\u00F6r allt via sociala medier, d\u00E4r inneh\u00E5llet kan n\u00E5 miljontals m\u00E4nniskor p\u00E5 n\u00E5gra timmar."),
      spacer(),

      heading3("Konkret exempel: Deepfakes inf\u00F6r svenska valet 2026"),
      bodyText("Inf\u00F6r det svenska riksdagsvalet 2026 spreds deepfake-videor av ledande politiker p\u00E5 TikTok och X. Videorna visade politiker som p\u00E5stods g\u00F6ra kontroversiella uttalanden om invandring och ekonomi \u2014 uttalanden som aldrig gjorts."),
      bodyText("Flera videor fick hundratusentals visningar innan de avsl\u00F6jades som fabricerade. AI-tekniken hade \u00E5terskapade politikernas r\u00F6ster och ansiktsuttryck s\u00E5 realistiskt att m\u00E5nga t\u00E4nkte inte p\u00E5 att ifr\u00E5gas\u00E4tta dem."),
      spacer(),

      heading3("Fr\u00E5gor"),
      bodyText("(a) Beskriv mekanismen \u2014 hur skapas och sprids deepfakes?"),
      bodyText("(b) Till\u00E4mpa SIFT p\u00E5 exemplet \u2014 vilka steg \u00E4r mest anv\u00E4ndbara f\u00F6r att avsl\u00F6ja deepfakes? Motivera."),
      bodyText("(c) Vilken av de fyra grundfr\u00E5gorna (VEM, VARF\u00D6R, HUR, N\u00C4R) \u00E4r mest relevant n\u00E4r du granskar deepfakes? Varf\u00F6r?"),
      spacer(),

      heading3("Reflektionsfr\u00E5ga"),
      italicText("\u201CVarf\u00F6r kan n\u00E5gon som m\u00F6ter denna deepfake utan f\u00F6rvarning finna den \u00F6vertygande?\u201D"),

      // ========== BILAGA 2: Station 2 - Filterbubbla ==========
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Stationsblad \u2014 Station 2: Filterbubbla och ekkammare"),
      spacer(),
      heading3("Bakgrund"),
      bodyText("Sociala medier-plattformar anv\u00E4nder AI-drivna algoritmer f\u00F6r att best\u00E4mma vilka inl\u00E4gg, videor och nyheter du ser. Algoritmerna prioriterar inneh\u00E5ll som f\u00E5r dig att stanna kvar l\u00E4ngre \u2014 och det \u00E4r ofta inneh\u00E5ll som v\u00E4cker starka k\u00E4nslor."),
      bodyText("N\u00E4r du s\u00F6ker p\u00E5 eller interagerar med ett \u00E4mne f\u00E5r du mer liknande inneh\u00E5ll. Detta skapar en \u201Crecommendation loop\u201D \u2014 en sj\u00E4lvf\u00F6rst\u00E4rkande spiral d\u00E4r ditt fl\u00F6de gradvis domineras av ett visst perspektiv, medan andra perspektiv f\u00F6rsvinner."),
      spacer(),

      heading3("Konkret exempel: Fr\u00E5n en s\u00F6kning till en filterbubbla"),
      bodyText("F\u00F6rest\u00E4ll dig f\u00F6ljande scenario:"),
      spacer(),
      bullet("Steg 1: En person s\u00F6ker p\u00E5 YouTube: \u201C\u00C4r vaccin farligt?\u201D"),
      bullet("Steg 2: Algoritmen visar videor som diskuterar vaccinrisker (h\u00F6gt engagemang)"),
      bullet("Steg 3: Personen tittar p\u00E5 en video. Algoritmen tolkar det som intresse."),
      bullet("Steg 4: \u201CRekommenderat\u201D-fl\u00F6det fylls med allt mer extrema videor om vaccinkonspiration"),
      bullet("Steg 5: Inom en vecka domineras personens fl\u00F6de av anti-vaccininneh\u00E5ll"),
      bullet("Steg 6: Personen upplever att \u201Calla\u201D ifr\u00E5gas\u00E4tter vaccin \u2014 f\u00F6r det \u00E4r allt hen ser"),
      spacer(),
      bodyText("Varje steg f\u00F6rst\u00E4rks av AI: algoritmen l\u00E4r sig vad som h\u00E5ller anv\u00E4ndaren kvar och optimerar f\u00F6r engagemang, inte f\u00F6r korrekthet."),
      spacer(),

      heading3("Fr\u00E5gor"),
      bodyText("(a) Beskriv mekanismen \u2014 hur skapar algoritmer ekkammare?"),
      bodyText("(b) Kan du anv\u00E4nda SIFT h\u00E4r? Vilka steg \u00E4r relevanta och vilka fungerar inte lika bra? Motivera."),
      bodyText("(c) Vilken grundfr\u00E5ga \u00E4r viktigast n\u00E4r man granskar algoritmiskt valt inneh\u00E5ll? Varf\u00F6r?"),
      spacer(),

      heading3("Reflektionsfr\u00E5ga"),
      italicText("\u201CVarf\u00F6r kan n\u00E5gon i en filterbubbla uppleva sin bild som mer trov\u00E4rdig \u00E4n andras?\u201D"),

      // ========== BILAGA 3: Station 3 - AI-genererade k\u00E4llor ==========
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Stationsblad \u2014 Station 3: AI-genererade k\u00E4llor"),
      spacer(),
      heading3("Bakgrund"),
      bodyText("NewsGuard har identifierat \u00F6ver 2 000 nyhetssajter d\u00E4r allt inneh\u00E5ll \u00E4r AI-genererat. Sajtena ser ut som riktig lokaljournalistik \u2014 med logotyper, artikelrubriker och \u201Credaktioner\u201D \u2014 men ingen m\u00E4nniska har skrivit en enda artikel."),
      bodyText("Bland dessa finns 167 sajter med ryska kopplingar som systematiskt sprider falska narrativ om Ukrainakriget, EU och Nato. AI g\u00F6r det m\u00F6jligt att driva tusentals s\u00E5dana sajter samtidigt, till l\u00E5g kostnad."),
      spacer(),

      heading3("Konkret exempel: Riktig vs. fejkad nyhetssajt"),
      bodyText("J\u00E4mf\u00F6r dessa tv\u00E5 nyhetssajter:"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [4513, 4513],
        rows: [
          new TableRow({
            children: [
              headerCell("Sajt A: sverigespress.se", 4513),
              headerCell("Sajt B: svt.se/nyheter", 4513),
            ],
          }),
          new TableRow({
            children: [
              cell("Skapad: 2024. Ingen Om oss-sida. Artiklarna har ingen byline (f\u00F6rfattarnamn). Inneh\u00E5llet blandar verkliga h\u00E4ndelser med felaktiga p\u00E5st\u00E5enden. Google-s\u00F6kning p\u00E5 sajten ger inga tr\u00E4ffar fr\u00E5n etablerade medier.", 4513),
              cell("Skapad: 2000-talet. Tydlig Om oss-sida med information om SVT:s uppdrag. Namngivna journalister. K\u00E4llor anges i texten. Google-s\u00F6kning bekr\u00E4ftar att SVT \u00E4r public service-media.", 4513),
            ],
          }),
        ],
      }),

      spacer(),
      bodyText("(\u201Csverigespress.se\u201D \u00E4r ett fiktivt men realistiskt exempel p\u00E5 hur AI-genererade nyhetssajter ser ut.)"),
      spacer(),

      heading3("Fr\u00E5gor"),
      bodyText("(a) Beskriv mekanismen \u2014 varf\u00F6r skapas fejkade nyhetssajter?"),
      bodyText("(b) Till\u00E4mpa SIFT steg f\u00F6r steg \u2014 vilka signaler avsl\u00F6jar en fejkad sajt?"),
      bodyText("(c) Hur hj\u00E4lper lateral reading specifikt vid denna typ av k\u00E4lla?"),
      spacer(),

      heading3("Reflektionsfr\u00E5ga"),
      italicText("\u201CVarf\u00F6r kan n\u00E5gon lita p\u00E5 denna sajt utan att ifr\u00E5gas\u00E4tta den?\u201D"),

      // ========== BILAGA 4: Station 4 - Automatiserade bottar ==========
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: Stationsblad \u2014 Station 4: Automatiserade bottar"),
      spacer(),
      heading3("Bakgrund"),
      bodyText("Bottar \u00E4r automatiserade konton p\u00E5 sociala medier som styrs av AI. De kan publicera inl\u00E4gg, gilla, dela och kommentera \u2014 allt utan m\u00E4nsklig inblandning."),
      bodyText("Bottn\u00E4tverk anv\u00E4nds f\u00F6r att sprida narrativ i stor skala. N\u00E4r hundratals konton delar samma budskap skapas en illusion av bred samh\u00E4llelig konsensus \u2014 \u201Calla tycker s\u00E5 h\u00E4r\u201D \u2014 trots att det \u00E4r samma akt\u00F6r bakom alla inl\u00E4gg."),
      spacer(),

      heading3("Konkret exempel: Koordinerad inautentisk aktivitet"),
      bodyText("Under en politisk debatt 2025 publicerade \u00F6ver 300 konton p\u00E5 X (Twitter) n\u00E4stan identiska inl\u00E4gg inom en 20-minutersperiod:"),
      spacer(),
      italicText("\u201CRegering X har hemliga planer p\u00E5 att avskaffa kontanter. Dela vidare s\u00E5 att ALLA f\u00E5r veta! #Frihet #Kontantupproret\u201D"),
      spacer(),
      bodyText("Analys av kontona visade:"),
      bullet("De flesta skapades inom samma vecka"),
      bullet("De hade f\u00E5 eller inga egna f\u00F6ljare"),
      bullet("De delade exakt samma l\u00E4nkar och hashtags"),
      bullet("Profilbilderna var AI-genererade (m\u00E4rkbar symmetri, suddiga bakgrunder)"),
      bullet("Inl\u00E4ggen publicerades med j\u00E4mna tidsintervall (varannan minut)"),
      spacer(),
      bodyText("Denna typ av beteende kallas \u201Ccoordinated inauthentic behavior\u201D och \u00E4r ett tecken p\u00E5 att n\u00E5gon medvetet f\u00F6rs\u00F6ker manipulera opinionen."),
      spacer(),

      heading3("Fr\u00E5gor"),
      bodyText("(a) Beskriv mekanismen \u2014 hur anv\u00E4nds bottar f\u00F6r att sprida konspirationsteorier?"),
      bodyText("(b) Till\u00E4mpa grundfr\u00E5gorna: Vem st\u00E5r bakom? Varf\u00F6r sprids det? Hur sprids det?"),
      bodyText("(c) Vilka SIFT-steg hj\u00E4lper dig identifiera bot-aktivitet?"),
      spacer(),

      heading3("Reflektionsfr\u00E5ga"),
      italicText("\u201CVarf\u00F6r kan n\u00E5gon som ser m\u00E5nga liknande inl\u00E4gg tolka det som bred samh\u00E4llelig konsensus?\u201D"),

      // ========== BILAGA 5: Gateway exit ticket (elevversion) ==========
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 5: Gateway exit ticket \u2014 Lektion 3"),
      spacer(),
      heading3("Del 1: K\u00E4llkritisk till\u00E4mpning (7 minuter)"),
      bodyText("L\u00E4s f\u00F6ljande inl\u00E4gg och granska det med hj\u00E4lp av de k\u00E4llkritiska verktyg du l\u00E4rt dig."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders,
                width: { size: 9026, type: WidthType.DXA },
                margins: cellMargins,
                shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
                children: [
                  new Paragraph({
                    spacing: { after: 80 },
                    children: [new TextRun({ text: "@frihetsbansen \u2022 3 veckor sedan", font: "Arial", size: 20, color: "666666" })],
                  }),
                  new Paragraph({
                    spacing: { after: 80 },
                    children: [new TextRun({ text: "BREAKING: L\u00E4ckt dokument visar att EU planerar att f\u00F6rbjuda kontanter senast 2028 f\u00F6r att kontrollera medborgarna. Dela innan de tar bort det! #KontantUpproret", font: "Arial", size: 22, bold: true })],
                  }),
                  new Paragraph({
                    spacing: { after: 60 },
                    children: [new TextRun({ text: "[Bild: AI-genererat foto av ett \u201Cofficiellt\u201D EU-dokument]", font: "Arial", size: 20, italics: true, color: "888888" })],
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "12 400 delningar \u2022 Kontot skapat f\u00F6r 3 veckor sedan", font: "Arial", size: 20, color: "666666" })],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),

      spacer(),
      bodyText("Uppgift: Granska inl\u00E4gget. Anv\u00E4nd de k\u00E4llkritiska verktyg du l\u00E4rt dig f\u00F6r att bed\u00F6ma trov\u00E4rdigheten."),
      spacer(),

      heading3("Del 2: Harvard-referering (5 minuter)"),
      bodyText("Formatera f\u00F6ljande k\u00E4llor med Harvardsystemet. Skriv b\u00E5de i-text-h\u00E4nvisning och referenslista."),
      spacer(),
      bullet("K\u00E4lla 1: En artikel av NewsGuard fr\u00E5n 2025 med titeln \u201CTracking AI-enabled Misinformation\u201D"),
      bullet("K\u00E4lla 2: En rapport fr\u00E5n Myndigheten f\u00F6r psykologiskt f\u00F6rsvar (MPF) fr\u00E5n 2025 med titeln \u201CAI och informationsp\u00E5verkan\u201D"),
      spacer(),

      heading3("Exit ticket"),
      boldBodyText("Fr\u00E5ga: ", "\u201CVarf\u00F6r kan AI-genererade konspirationsteorier vara sv\u00E5rare att granska k\u00E4llkritiskt \u00E4n traditionella? Resonera utifr\u00E5n minst en mekanism vi diskuterat idag.\u201D"),

      // ========== BILAGA 6: Gateway checklista (l\u00E4rarversion) ==========
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 6: Gateway checklista \u2014 L\u00E4rarversion"),
      spacer(),
      bodyText("Anv\u00E4nd denna checklista f\u00F6r snabb sortering av exit tickets. M\u00E5let \u00E4r att avg\u00F6ra gruppsammans\u00E4ttning inf\u00F6r L4-seminariet."),
      spacer(),

      heading3("Checklista (3 kriterier)"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 6026, 1250, 1250],
        rows: [
          new TableRow({
            children: [
              headerCell("Nr", 500),
              headerCell("Kriterium", 6026),
              headerCell("Ja", 1250),
              headerCell("Nej", 1250),
            ],
          }),
          new TableRow({
            children: [
              cell("1", 500),
              cell("Anv\u00E4nder minst 2 k\u00E4llkritiska verktyg (grundfr\u00E5gor och/eller SIFT)", 6026),
              cell("", 1250),
              cell("", 1250),
            ],
          }),
          new TableRow({
            children: [
              cell("2", 500),
              cell("Identifierar r\u00E4tt problem/vilseledning i exemplet", 6026),
              cell("", 1250),
              cell("", 1250),
            ],
          }),
          new TableRow({
            children: [
              cell("3", 500),
              cell("Harvard-referering i huvudsak korrekt (minst 1 av 2 k\u00E4llor)", 6026),
              cell("", 1250),
              cell("", 1250),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Sorteringsguide f\u00F6r L4"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 1500, 6026],
        rows: [
          new TableRow({
            children: [
              headerCell("Resultat", 1500),
              headerCell("Niv\u00E5", 1500),
              headerCell("Konsekvens f\u00F6r L4-seminariegruppering", 6026),
            ],
          }),
          new TableRow({
            children: [
              cell("Alla tre Ja", 1500),
              cell("Stark", 1500),
              cell("Kan st\u00F6tta andra i seminariegrupp. Placera som \u201Cankare\u201D i blandade grupper.", 6026),
            ],
          }),
          new TableRow({
            children: [
              cell("1\u20132 Ja", 1500),
              cell("Medel", 1500),
              cell("Beh\u00F6ver blandning med starkare elever. Undvik att s\u00E4tta tv\u00E5 medel-elever utan st\u00F6d.", 6026),
            ],
          }),
          new TableRow({
            children: [
              cell("0 Ja", 1500),
              cell("Beh\u00F6ver st\u00F6d", 1500),
              cell("L\u00E4raren planerar st\u00F6dinsats inf\u00F6r L4 (t.ex. kort samtal, extra \u00F6vning med SIFT). Placera med stark elev i seminariegruppen.", 6026),
            ],
          }),
        ],
      }),

      spacer(),
      bodyText("Tips: Sortera exit tickets direkt efter lektionen i tre h\u00F6gar (Stark / Medel / Beh\u00F6ver st\u00F6d). Anv\u00E4nd h\u00F6garna f\u00F6r att planera seminariegrupper \u2014 varje grupp b\u00F6r ha minst en \u201CStark\u201D-elev."),
    ],
  }],
});

// ============================================================
// PPTX - Lektion 3 presentation
// ============================================================

function generatePptx() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Cornelius Agent";
  pres.title = "Lektion 3: AI och konspirationsteorier";

  // Color palette - Midnight Executive (samma som lektion-2.pptx)
  const DARK_BG = "1B2A4A";
  const MED_BG = "243656";
  const ACCENT = "4DA8DA";
  const ACCENT2 = "6BC5D2";
  const TEXT_WHITE = "FFFFFF";
  const TEXT_LIGHT = "CBD5E1";
  const TEXT_MUTED = "94A3B8";

  // --- Slide 1: Titel ---
  const s1 = pres.addSlide();
  s1.background = { color: DARK_BG };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.5, w: 10, h: 1.125, fill: { color: MED_BG } });
  s1.addText("Lektion 3", { x: 0.8, y: 1.2, w: 8.4, h: 0.8, fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, charSpacing: 6 });
  s1.addText("AI och konspirationsteorier", { x: 0.8, y: 2.0, w: 8.4, h: 1.2, fontSize: 40, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s1.addText("N\u00E4r AI f\u00F6rst\u00E4rker konspirationsteoriers spridning", { x: 0.8, y: 3.2, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });
  s1.addText("Samh\u00E4llskunskap 3 / Internationella relationer", { x: 0.8, y: 4.7, w: 8.4, h: 0.5, fontSize: 13, fontFace: "Arial", color: TEXT_MUTED });

  // --- Slide 2: Retrieval review ---
  const s2 = pres.addSlide();
  s2.background = { color: DARK_BG };
  s2.addText("\u00C5terkallelse fr\u00E5n L2", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s2.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s2.addText("Para ihop r\u00E4tt grundfr\u00E5ga med r\u00E4tt SIFT-steg:", { x: 0.8, y: 1.2, w: 8.4, h: 0.5, fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT });

  const matchItems = [
    { left: "VEM?", right: "Trace claims" },
    { left: "VARF\u00D6R?", right: "Stop" },
    { left: "HUR?", right: "Find better coverage" },
    { left: "N\u00C4R?", right: "Investigate the source" },
  ];

  matchItems.forEach((item, i) => {
    const yPos = 1.9 + i * 0.75;
    s2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 3.8, h: 0.6, fill: { color: MED_BG } });
    s2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.08, h: 0.6, fill: { color: ACCENT } });
    s2.addText(item.left, { x: 1.1, y: yPos, w: 3.3, h: 0.6, fontSize: 18, fontFace: "Arial", color: TEXT_WHITE, bold: true, valign: "middle", margin: 0 });

    s2.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yPos, w: 3.8, h: 0.6, fill: { color: MED_BG } });
    s2.addText(item.right, { x: 5.6, y: yPos, w: 3.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: ACCENT2, valign: "middle", margin: 0 });
  });

  s2.addText("Arbeta i par. 2 minuter.", { x: 0.8, y: 4.9, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 3: De fyra AI-mekanismerna ---
  const s3 = pres.addSlide();
  s3.background = { color: DARK_BG };
  s3.addText("Fyra AI-mekanismer", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s3.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  const mechanisms = [
    { name: "Deepfakes", desc: "AI-genererade bilder, video, r\u00F6ster av verkliga personer" },
    { name: "Filterbubbla", desc: "Algoritmer f\u00F6rst\u00E4rker det du redan tittar p\u00E5" },
    { name: "AI-k\u00E4llor", desc: "2 000+ fejkade nyhetssajter drivna helt av AI" },
    { name: "Bottar", desc: "N\u00E4tverk som sprider samma budskap i stor skala" },
  ];

  mechanisms.forEach((m, i) => {
    const yPos = 1.2 + i * 1.05;
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: MED_BG } });
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.08, h: 0.9, fill: { color: ACCENT } });
    s3.addText(m.name, { x: 1.1, y: yPos, w: 2.5, h: 0.9, fontSize: 20, fontFace: "Arial", color: ACCENT, bold: true, valign: "middle", margin: 0 });
    s3.addText(m.desc, { x: 3.6, y: yPos, w: 5.4, h: 0.9, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, valign: "middle", margin: 0 });
  });

  s3.addText("Vid varje station anv\u00E4nder ni SIFT och grundfr\u00E5gorna f\u00F6r att granska ett exempel.", { x: 0.8, y: 5.0, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 4: Stationsarbete - instruktioner ---
  const s4 = pres.addSlide();
  s4.background = { color: DARK_BG };
  s4.addText("Stationsarbete", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s4.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s4.addText([
    { text: "4 stationer \u2022 8 min per station \u2022 Arbeta i par", options: { fontSize: 20, fontFace: "Arial", color: ACCENT, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Vid varje station:", options: { fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(a) L\u00E4s bakgrundstexten om mekanismen", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(b) Till\u00E4mpa SIFT / grundfr\u00E5gorna p\u00E5 exemplet", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(c) Besvara reflektionsfr\u00E5gan", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Skriv ner era svar \u2014 ni beh\u00F6ver dem i summeringen.", options: { fontSize: 16, fontFace: "Arial", color: ACCENT2, italic: true } },
  ], { x: 0.8, y: 1.2, w: 8.4, h: 3.5 });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s4.addText("Tredjepersons-framing: Analysera varf\u00F6r n\u00E5gon KAN finna det \u00F6vertygande.", { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fontSize: 15, fontFace: "Arial", color: ACCENT, align: "center", valign: "middle" });

  // --- Slide 5: Station 1 - Deepfakes ---
  const s5 = pres.addSlide();
  s5.background = { color: DARK_BG };
  s5.addText("Station 1: Deepfakes", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s5.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, fill: { color: MED_BG } });
  s5.addText("AI-genererade bilder, videor och r\u00F6ster som ser verkliga ut. Anv\u00E4nds f\u00F6r att fabricera h\u00E4ndelser och l\u00E4gga ord i politikers mun.", { x: 1.1, y: 1.3, w: 7.8, h: 1.0, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });

  s5.addText([
    { text: "Case: Deepfakes inf\u00F6r svenska valet 2026", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(a) Hur skapas och sprids deepfakes?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(b) Vilka SIFT-steg \u00E4r mest anv\u00E4ndbara?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(c) Vilken grundfr\u00E5ga \u00E4r mest relevant?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 0.8, y: 2.7, w: 8.4, h: 2.0 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s5.addText("Reflektion: Varf\u00F6r kan n\u00E5gon som m\u00F6ter denna deepfake utan f\u00F6rvarning finna den \u00F6vertygande?", { x: 1.0, y: 4.7, w: 8.0, h: 0.6, fontSize: 14, fontFace: "Arial", color: ACCENT, italic: true, valign: "middle" });

  // --- Slide 6: Station 2 - Filterbubbla ---
  const s6 = pres.addSlide();
  s6.background = { color: DARK_BG };
  s6.addText("Station 2: Filterbubbla och ekkammare", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s6.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, fill: { color: MED_BG } });
  s6.addText("Algoritmer visar mer av det du redan tittar p\u00E5. En s\u00F6kning kan starta en sj\u00E4lvf\u00F6rst\u00E4rkande spiral mot allt mer extremt inneh\u00E5ll.", { x: 1.1, y: 1.3, w: 7.8, h: 1.0, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });

  s6.addText([
    { text: "Case: Fr\u00E5n en s\u00F6kning till en filterbubbla", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(a) Hur skapar algoritmer ekkammare?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(b) Kan SIFT anv\u00E4ndas h\u00E4r? Vilka steg fungerar/fungerar inte?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(c) Vilken grundfr\u00E5ga \u00E4r viktigast f\u00F6r algoritmiskt valt inneh\u00E5ll?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 0.8, y: 2.7, w: 8.4, h: 2.0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s6.addText("Reflektion: Varf\u00F6r kan n\u00E5gon i en filterbubbla uppleva sin bild som mer trov\u00E4rdig \u00E4n andras?", { x: 1.0, y: 4.7, w: 8.0, h: 0.6, fontSize: 14, fontFace: "Arial", color: ACCENT, italic: true, valign: "middle" });

  // --- Slide 7: Station 3 - AI-genererade k\u00E4llor ---
  const s7 = pres.addSlide();
  s7.background = { color: DARK_BG };
  s7.addText("Station 3: AI-genererade k\u00E4llor", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s7.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, fill: { color: MED_BG } });
  s7.addText("\u00D6ver 2 000 fejkade nyhetssajter drivna av AI. Ser ut som riktig journalistik men har ingen redaktion.", { x: 1.1, y: 1.3, w: 7.8, h: 1.0, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });

  s7.addText([
    { text: "Case: Riktig vs. fejkad nyhetssajt", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(a) Varf\u00F6r skapas fejkade nyhetssajter?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(b) Till\u00E4mpa SIFT steg f\u00F6r steg \u2014 vilka signaler avsl\u00F6jar fejk?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(c) Hur hj\u00E4lper lateral reading h\u00E4r?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 0.8, y: 2.7, w: 8.4, h: 2.0 });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s7.addText("Reflektion: Varf\u00F6r kan n\u00E5gon lita p\u00E5 denna sajt utan att ifr\u00E5gas\u00E4tta den?", { x: 1.0, y: 4.7, w: 8.0, h: 0.6, fontSize: 14, fontFace: "Arial", color: ACCENT, italic: true, valign: "middle" });

  // --- Slide 8: Station 4 - Automatiserade bottar ---
  const s8 = pres.addSlide();
  s8.background = { color: DARK_BG };
  s8.addText("Station 4: Automatiserade bottar", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s8.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s8.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, fill: { color: MED_BG } });
  s8.addText("N\u00E4tverk av AI-styrda konton som sprider samma budskap. Skapar illusion av bred konsensus.", { x: 1.1, y: 1.3, w: 7.8, h: 1.0, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });

  s8.addText([
    { text: "Case: Koordinerad inautentisk aktivitet", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(a) Hur anv\u00E4nds bottar f\u00F6r att sprida konspirationsteorier?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(b) Till\u00E4mpa grundfr\u00E5gorna: Vem? Varf\u00F6r? Hur?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "(c) Vilka SIFT-steg identifierar bot-aktivitet?", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 0.8, y: 2.7, w: 8.4, h: 2.0 });

  s8.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s8.addText("Reflektion: Varf\u00F6r kan n\u00E5gon som ser m\u00E5nga liknande inl\u00E4gg tolka det som bred konsensus?", { x: 1.0, y: 4.7, w: 8.0, h: 0.6, fontSize: 14, fontFace: "Arial", color: ACCENT, italic: true, valign: "middle" });

  // --- Slide 9: Summering ---
  const s9 = pres.addSlide();
  s9.background = { color: DARK_BG };
  s9.addText("Summering", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s9.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s9.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 8.4, h: 2.8, fill: { color: MED_BG } });
  s9.addText([
    { text: "Diskutera:", options: { fontSize: 20, fontFace: "Arial", color: ACCENT, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Vilka m\u00F6nster s\u00E5g ni \u00F6ver stationerna?", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Vilken mekanism var sv\u00E5rast att genomsk\u00E5da?", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Vilket k\u00E4llkritiskt verktyg var mest anv\u00E4ndbart?", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 1.2, y: 1.5, w: 7.6, h: 2.4 });

  s9.addText("Alla fyra mekanismer f\u00F6rst\u00E4rks av AI \u2014 men kan granskas med samma verktyg.", { x: 0.8, y: 4.6, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: ACCENT, italic: true, align: "center" });

  // --- Slide 10: Gateway exit ticket ---
  const s10 = pres.addSlide();
  s10.background = { color: DARK_BG };
  s10.addText("Gateway exit ticket", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s10.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 4.0, h: 2.5, fill: { color: MED_BG } });
  s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 0.08, h: 2.5, fill: { color: ACCENT } });
  s10.addText([
    { text: "Del 1: K\u00E4llkritisk till\u00E4mpning", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "(7 min)", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Granska ett nytt exempel med dina k\u00E4llkritiska verktyg.", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Inga meningsstartare.", options: { fontSize: 15, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "Inga ledfr\u00E5gar.", options: { fontSize: 15, fontFace: "Arial", color: TEXT_WHITE, bold: true } },
  ], { x: 1.1, y: 1.3, w: 3.5, h: 2.3 });

  s10.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.0, h: 2.5, fill: { color: MED_BG } });
  s10.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 0.08, h: 2.5, fill: { color: ACCENT } });
  s10.addText([
    { text: "Del 2: Harvard-referering", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "(5 min)", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Formatera 2 k\u00E4llor korrekt:", options: { fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "I-text-h\u00E4nvisning", options: { bullet: true, fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "Referenslista", options: { bullet: true, fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 5.5, y: 1.3, w: 3.5, h: 2.3 });

  s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.2, w: 8.4, h: 1.0, fill: { color: MED_BG } });
  s10.addText("Resultaten avg\u00F6r era grupper inf\u00F6r n\u00E4sta lektions seminarium.", { x: 0.8, y: 4.2, w: 8.4, h: 1.0, fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, align: "center", valign: "middle" });

  return pres;
}

// ============================================================
// GENERATE BOTH FILES
// ============================================================

const OUTPUT_DIR = "/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier";

async function main() {
  // Generate DOCX
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(`${OUTPUT_DIR}/lektion-3.docx`, buffer);
  console.log("lektion-3.docx skapad!");

  // Generate PPTX
  const pres = generatePptx();
  await pres.writeFile({ fileName: `${OUTPUT_DIR}/lektion-3.pptx` });
  console.log("lektion-3.pptx skapad!");
}

main().catch(err => {
  console.error("Fel vid generering:", err);
  process.exit(1);
});
