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
// DOCX - Lektion 2: K\u00E4llkritiska verktyg
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
        children: [new TextRun({ text: "Lektion 2: K\u00E4llkritiska verktyg", font: "Arial" })],
      }),
      bodyText("Grundfr\u00E5gorna, SIFT-metoden, lateral reading och Harvard-referering"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // ========== L\u00C4RANDEM\u00C5L ==========
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna anv\u00E4nda k\u00E4llkritiska verktyg (de fyra grundfr\u00E5gorna, SIFT-metoden, lateral reading) f\u00F6r att granska digitalt inneh\u00E5ll, samt introduceras till Harvard-referering p\u00E5 grundl\u00E4ggande niv\u00E5."),
      spacer(),
      bullet("Till\u00E4mpa de fyra k\u00E4llkritiska grundfr\u00E5gorna p\u00E5 digitalt inneh\u00E5ll (m\u00E5l 1)"),
      bullet("Anv\u00E4nda SIFT-metoden och lateral reading f\u00F6r att verifiera p\u00E5st\u00E5enden (m\u00E5l 1)"),
      bullet("Anv\u00E4nda omv\u00E4nd bild\u00F6kning/Google Lens som verifieringsverktyg (m\u00E5l 1)"),
      bullet("P\u00E5b\u00F6rja k\u00E4llh\u00E4nvisning med Harvard-systemet (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven anv\u00E4nder i n\u00E5gon m\u00E5n k\u00E4llkritiska verktyg och g\u00F6r en \u00F6versiktlig granskning av digitalt inneh\u00E5ll. Eleven refererar till k\u00E4llor p\u00E5 ett i huvudsak fungerande s\u00E4tt."),
      boldBodyText("C: ", "Eleven anv\u00E4nder k\u00E4llkritiska verktyg systematiskt och g\u00F6r en utf\u00F6rlig granskning av digitalt inneh\u00E5ll med visst kritiskt perspektiv. Eleven refererar till k\u00E4llor p\u00E5 ett fungerande s\u00E4tt med Harvard-referering."),
      boldBodyText("A: ", "Eleven anv\u00E4nder k\u00E4llkritiska verktyg med s\u00E4kerhet och g\u00F6r en utf\u00F6rlig och nyanserad granskning av digitalt inneh\u00E5ll ur flera perspektiv. Eleven refererar till k\u00E4llor p\u00E5 ett v\u00E4l fungerande s\u00E4tt med konsekvent Harvard-referering."),

      // ========== CENTRALT INNEH\u00C5LL ==========
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("K\u00E4llh\u00E4nvisning enligt vanliga system."),

      // ========== F\u00D6RBEREDELSE ==========
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 1 \u2014 notera vad eleverna lyfte som trov\u00E4rdighetsmark\u00F6rer"),
      bullet("F\u00F6rbered retrieval review-fr\u00E5ga baserat p\u00E5 L1 exit ticket (vad g\u00F6r AI-inneh\u00E5ll \u00F6vertygande)"),
      bullet("V\u00E4lj ett verkligt exempel f\u00F6r lateral reading-demo (t.ex. en vilseledande h\u00E4lsosajt eller sensationell nyhetsartikel)"),
      bullet("F\u00F6rbered \u00F6vningsk\u00E4llor (se bilaga 4) \u2014 skriv ut eller dela digitalt"),
      bullet("Skriv ut bilaga 1: De fyra k\u00E4llkritiska grundfr\u00E5gorna (\u00F6versiktskort, en per elev)"),
      bullet("Skriv ut bilaga 5: Exit ticket L2 (en per elev)"),
      bullet("S\u00E4kerst\u00E4ll att alla elever har tillg\u00E5ng till en dator/surfplatta f\u00F6r lateral reading-\u00F6vning"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),

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
              cell("\u00C5terkallelse fr\u00E5n L1", 2200),
              cell("Parsamtal: \u201CVad gjorde AI-inneh\u00E5llet \u00F6vertygande? N\u00E4mn trov\u00E4rdighetsmark\u00F6rer fr\u00E5n f\u00F6rra lektionen.\u201D Sedan helklasssammanfattning. Baseras p\u00E5 L1 exit ticket.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201323 min", 1000),
              cell("Instruktion", 1300),
              cell("Grundfr\u00E5gor + SIFT + lateral reading", 2200),
              cell("De fyra k\u00E4llkritiska grundfr\u00E5gorna (5 min). SIFT-metoden: Stop, Investigate, Find, Trace (5 min). Modellering: l\u00E4raren visar lateral reading steg f\u00F6r steg p\u00E5 ETT verkligt exempel (5 min). Omv\u00E4nd bild\u00F6kning/Google Lens visas som del av \u201CInvestigate the source\u201D. Tredjepersons-framing p\u00E5minnelse.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("23\u201348 min", 1000),
              cell("Guidad \u00F6vning", 1300),
              cell("Hands-on SIFT/lateral reading", 2200),
              cell("Eleverna f\u00E5r 3 k\u00E4llor av varierande trov\u00E4rdighet (se bilaga 4). Arbete i par: till\u00E4mpa SIFT p\u00E5 varje k\u00E4lla. L\u00E4raren cirkulerar och ger feedback. Avsluta med helklasssammanfattning: vilka k\u00E4llor var trov\u00E4rdiga och hur visste ni det?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("48\u201363 min", 1000),
              cell("Sj\u00E4lvst\u00E4ndig \u00F6vning", 1300),
              cell("Egen k\u00E4lla + grundfr\u00E5gor + SIFT", 2200),
              cell("Eleverna v\u00E4ljer en EGEN digital k\u00E4lla och till\u00E4mpar grundfr\u00E5gorna + SIFT. Korta skriftliga svar (1\u20132 meningar per grundfr\u00E5ga). Inkludera omv\u00E4nd bild\u00F6kning om k\u00E4llans bild \u00E4r relevant.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("63\u201372 min", 1000),
              cell("Harvard-referering", 1300),
              cell("Introduktion + \u00F6vning", 2200),
              cell("Varf\u00F6r k\u00E4llh\u00E4nvisa? (2 min). F\u00F6renklad Harvard-modell: (Efternamn, \u00E5r) i texten + Efternamn (\u00C5r). Titel. K\u00E4lla. i referenslistan (4 min). Kort \u00F6vning: formatera 2 k\u00E4llor fr\u00E5n den guidade \u00F6vningen korrekt (3 min).", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("72\u201380 min", 1000),
              cell("Avslut", 1300),
              cell("Exit ticket + preview av L3", 2200),
              cell("Exit ticket: \u201CV\u00E4lj en av de fyra k\u00E4llkritiska grundfr\u00E5gorna. F\u00F6rklara med ett eget exempel hur fr\u00E5gan hj\u00E4lper dig att granska ett digitalt inneh\u00E5ll.\u201D Preview: \u201CN\u00E4sta lektion unders\u00F6ker vi hur AI anv\u00E4nds f\u00F6r att f\u00F6rst\u00E4rka konspirationsteorier \u2014 stationsarbete.\u201D", 4526),
            ],
          }),
        ],
      }),

      // ========== L\u00C4RARINSTRUKTIONER ==========
      heading2("L\u00E4rarinstruktioner"),

      // --- Fas 1: Retrieval review ---
      heading3("Fas 1: Retrieval review (0\u20138 min)"),
      bodyText("Syftet \u00E4r att \u00E5terkalla elevernas upplevelse fr\u00E5n L1 (AI-labben) och namnge de trov\u00E4rdighetsmark\u00F6rer de identifierade. Detta bygger bryggan till dagens verktyg."),
      spacer(),
      boldBodyText("Parsamtal (0\u20134 min): ", ""),
      italicText("S\u00E4g: \u201CF\u00F6rra lektionen skapade ni AI-genererat inneh\u00E5ll. Nu vill jag att ni t\u00E4nker tillbaka: Vad var det som gjorde inneh\u00E5llet \u00F6vertygande? V\u00E4nd er till den som sitter bredvid och n\u00E4mn minst tv\u00E5 saker. Ni har 3 minuter.\u201D"),
      spacer(),
      boldBodyText("Helklasssammanfattning (4\u20138 min): ", ""),
      italicText("S\u00E4g: \u201CVilka trov\u00E4rdighetsmark\u00F6rer hittade ni?\u201D"),
      bodyText("Samla 4\u20135 svar p\u00E5 tavlan. Typiska svar fr\u00E5n L1 exit tickets:"),
      bullet("Spr\u00E5ket l\u00E5ter professionellt/journalistiskt"),
      bullet("Det refererar till verkliga h\u00E4ndelser och platser"),
      bullet("Det spelar p\u00E5 k\u00E4nslor (r\u00E4dsla, ilska, empati)"),
      bullet("Det ser ut som det kommer fr\u00E5n en trov\u00E4rdig k\u00E4lla"),
      spacer(),
      italicText("\u201CBra. Ni identifierade varf\u00F6r inneh\u00E5llet VAR \u00F6vertygande. Men ni hade inga verktyg att GRANSKA det med. Idag f\u00E5r ni verktygsl\u00E5dan.\u201D"),

      // --- Fas 2: Instruktion ---
      heading3("Fas 2: Instruktion (8\u201323 min) \u2014 MAX 15 minuter"),
      bodyText("H\u00E5ll instruktionen stram. Eleverna beh\u00F6ver \u00F6va, inte bara h\u00F6ra. Tre delar: grundfr\u00E5gor, SIFT, modellering."),
      spacer(),

      boldBodyText("Del 1: De fyra k\u00E4llkritiska grundfr\u00E5gorna (8\u201313 min)", ""),
      bodyText("Dela ut bilaga 1 (\u00F6versiktskort). Presentera fr\u00E5gorna:"),
      spacer(),
      bullet("VEM? \u2014 Vem \u00E4r avs\u00E4ndaren? Vad vet vi om dem? \u00C4r de experter, journalister, anonyma?"),
      bullet("VARF\u00D6R? \u2014 Varf\u00F6r publiceras detta? Vad \u00E4r syftet? Informera, \u00F6vertyga, s\u00E4lja, provocera?"),
      bullet("HUR? \u2014 Hur presenteras informationen? Spr\u00E5k, k\u00E4nslor, bilder, \u201Cbevis\u201D?"),
      bullet("N\u00C4R? \u2014 N\u00E4r publicerades det? \u00C4r informationen aktuell? Har kontexten f\u00F6r\u00E4ndrats?"),
      spacer(),
      italicText("S\u00E4g: \u201CDessa fyra fr\u00E5gor \u00E4r ert analytiska ramverk. De fungerar p\u00E5 ALLA digitala k\u00E4llor \u2014 nyhetssajter, sociala medier, AI-genererat inneh\u00E5ll. Skriv ner dem. De kommer att anv\u00E4ndas i varje lektion h\u00E4rifr\u00E5n.\u201D"),
      spacer(),

      boldBodyText("Del 2: SIFT-metoden (13\u201318 min)", ""),
      bodyText("Presentera SIFT som den PRAKTISKA metoden \u2014 grundfr\u00E5gorna \u00E4r det analytiska ramverket, SIFT \u00E4r handlingsplanen."),
      spacer(),
      bullet("S \u2014 Stop. Stanna upp. Reagera inte direkt. Dela inte. T\u00E4nk."),
      bullet("I \u2014 Investigate the source. Vem \u00E4r avs\u00E4ndaren? \u00D6ppna en NY flik och s\u00F6k p\u00E5 k\u00E4llan (lateral reading). Anv\u00E4nd omv\u00E4nd bild\u00F6kning/Google Lens om det finns bilder."),
      bullet("F \u2014 Find better coverage. Vad s\u00E4ger ANDRA k\u00E4llor om samma p\u00E5st\u00E5ende? S\u00F6k p\u00E5 p\u00E5st\u00E5endet i steg 1."),
      bullet("T \u2014 Trace claims. Varifr\u00E5n kommer det ursprungliga p\u00E5st\u00E5endet? Finns det en prim\u00E4rk\u00E4lla?"),
      spacer(),
      boldBodyText("Omv\u00E4nd bild\u00F6kning/Google Lens: ", "Visa kort hur man h\u00F6gerklickar p\u00E5 en bild och v\u00E4ljer \u201CS\u00F6k med Google Lens\u201D (Chrome) eller g\u00E5r till images.google.com och laddar upp bilden. Detta \u00E4r en del av steget \u201CInvestigate the source\u201D \u2014 det hj\u00E4lper att avg\u00F6ra om en bild \u00E4r \u00E4kta, manipulerad eller tagen ur sin kontext."),
      spacer(),

      boldBodyText("Del 3: Modellering \u2014 lateral reading live (18\u201323 min)", ""),
      italicText("S\u00E4g: \u201CNu visar jag hur det g\u00E5r till i praktiken. Jag g\u00F6r det steg f\u00F6r steg s\u00E5 att ni ser processen.\u201D"),
      spacer(),
      bodyText("V\u00E4lj ETT verkligt exempel (t.ex. en sensationell rubrik fr\u00E5n en ok\u00E4nd nyhetssajt). Visa p\u00E5 sk\u00E4rmen:"),
      bullet("S \u2014 \u201CJag stannar upp. Jag delar inte direkt.\u201D"),
      bullet("I \u2014 \u201CJag \u00F6ppnar en NY flik och s\u00F6ker p\u00E5 sajten. Vad s\u00E4ger andra om den h\u00E4r k\u00E4llan?\u201D (visa lateral reading live)"),
      bullet("F \u2014 \u201CJag s\u00F6ker p\u00E5 sj\u00E4lva p\u00E5st\u00E5endet. Rapporterar n\u00E5gon trov\u00E4rdig k\u00E4lla samma sak?\u201D"),
      bullet("T \u2014 \u201CHur ser den ursprungliga k\u00E4llan ut? Finns det en forskningsrapport eller myndighetsk\u00E4lla bakom?\u201D"),
      spacer(),
      boldBodyText("Tredjepersons-framing p\u00E5minnelse: ", ""),
      italicText("\u201CKom ih\u00E5g v\u00E5r princip: vi fr\u00E5gar inte om NI tror p\u00E5 det h\u00E4r. Vi fr\u00E5gar: vad g\u00F6r det h\u00E4r inneh\u00E5llet trov\u00E4rdigt f\u00F6r den som m\u00F6ter det utan f\u00F6rvarning? Grundfr\u00E5gorna och SIFT hj\u00E4lper oss att svara p\u00E5 det analytiskt.\u201D"),

      // --- Fas 3: Guidad \u00F6vning ---
      heading3("Fas 3: Guidad \u00F6vning (23\u201348 min)"),
      bodyText("Eleverna till\u00E4mpar SIFT och grundfr\u00E5gorna p\u00E5 verkliga k\u00E4llor. Detta \u00E4r lektionens k\u00E4rna \u2014 hands-on \u00F6vning \u00E4r viktigare \u00E4n teori."),
      spacer(),
      italicText("S\u00E4g: \u201CNu \u00E4r det er tur. Ni f\u00E5r tre k\u00E4llor. Er uppgift \u00E4r att till\u00E4mpa SIFT p\u00E5 var och en av dem. Arbeta i par. Anv\u00E4nd b\u00E5de grundfr\u00E5gorna och SIFT-stegen. Ni har 20 minuter.\u201D"),
      spacer(),
      bodyText("Dela ut bilaga 4 (\u00F6vningsk\u00E4llor). K\u00E4llorna \u00E4r:"),
      bullet("K\u00E4lla A: En seri\u00F6s nyhetsartikel fr\u00E5n en etablerad redaktion (trov\u00E4rdig)"),
      bullet("K\u00E4lla B: En vilseledande sajt som presenterar \u00E5sikter som nyheter (problematisk)"),
      bullet("K\u00E4lla C: Ett socialt medie-inl\u00E4gg med sensationellt p\u00E5st\u00E5ende utan k\u00E4llangivelse (otrov\u00E4rdig)"),
      spacer(),
      boldBodyText("L\u00E4raren cirkulerar: ", ""),
      bullet("Till par som fastnar vid S: \u201CVad \u00E4r er f\u00F6rsta reaktion? Noterade ni n\u00E5got som b\u00F6r f\u00E5 er att stanna upp?\u201D"),
      bullet("Till par som hoppar \u00F6ver I: \u201CHar ni \u00F6ppnat en ny flik och s\u00F6kt p\u00E5 avs\u00E4ndaren?\u201D"),
      bullet("Till par som \u00E4r klara snabbt: \u201CJag ser att ni bed\u00F6mde k\u00E4lla B. Kan ni anv\u00E4nda grundfr\u00E5gorna f\u00F6r att f\u00F6rdjupa \u2014 VEM vinner p\u00E5 att detta sprids?\u201D"),
      spacer(),
      boldBodyText("Helklasssammanfattning (43\u201348 min): ", ""),
      italicText("\u201CVilka k\u00E4llor var trov\u00E4rdiga? Hur visste ni det? Vilka steg i SIFT var mest avgn\u00E5rande?\u201D"),
      bodyText("Samla m\u00F6nster. Lyft exempel p\u00E5 lateral reading och omv\u00E4nd bild\u00F6kning som gav tydliga resultat."),

      // --- Fas 4: Sj\u00E4lvst\u00E4ndig \u00F6vning ---
      heading3("Fas 4: Sj\u00E4lvst\u00E4ndig \u00F6vning (48\u201363 min)"),
      italicText("S\u00E4g: \u201CNu ska ni v\u00E4lja en EGEN digital k\u00E4lla \u2014 n\u00E5got ni sj\u00E4lva har m\u00F6tt p\u00E5 sociala medier, i nyhetsfl\u00F6det eller p\u00E5 n\u00E4tet. Till\u00E4mpa grundfr\u00E5gorna och SIFT p\u00E5 den k\u00E4llan. Skriv korta svar \u2014 1\u20132 meningar per grundfr\u00E5ga.\u201D"),
      spacer(),
      bodyText("Eleverna arbetar enskilt. Instruktioner p\u00E5 tavlan:"),
      bullet("V\u00E4lj en digital k\u00E4lla (nyhet, inl\u00E4gg, video, bild)"),
      bullet("Skriv svar p\u00E5 de fyra grundfr\u00E5gorna (VEM, VARF\u00D6R, HUR, N\u00C4R)"),
      bullet("G\u00F6r minst ETT SIFT-steg praktiskt (lateral reading eller omv\u00E4nd bild\u00F6kning)"),
      bullet("Sammanfatta: \u00C4r k\u00E4llan trov\u00E4rdig? Varf\u00F6r/varf\u00F6r inte?"),
      spacer(),
      boldBodyText("Differentiering: ", ""),
      boldBodyText("St\u00F6d (mot E): ", "Ge f\u00F6rslag p\u00E5 k\u00E4llor att granska (t.ex. en specifik nyhetsartikel). Erbjud meningsstartare: \u201CAvs\u00E4ndaren \u00E4r...\u201D, \u201CSyftet verkar vara...\u201D"),
      boldBodyText("Utmaning (mot A): ", "V\u00E4lj en k\u00E4lla som \u00E4r sv\u00E5rbed\u00F6md (inte uppenbart fejk eller uppenbart trov\u00E4rdig). Resonera om vilka grundfr\u00E5gor som \u00E4r MEST avgn\u00E5rande f\u00F6r just den k\u00E4llan och varf\u00F6r."),

      // --- Fas 5: Harvard-referering intro ---
      heading3("Fas 5: Harvard-referering intro (63\u201372 min) \u2014 MAX 9 minuter"),
      bodyText("H\u00E5ll detta kort. Syftet \u00E4r att introducera grundprincipen \u2014 fullst\u00E4ndig \u00F6vning kommer i L3 exit ticket och L5."),
      spacer(),

      boldBodyText("Varf\u00F6r k\u00E4llh\u00E4nvisa? (63\u201365 min)", ""),
      bullet("F\u00F6r att l\u00E4saren ska kunna kontrollera dina p\u00E5st\u00E5enden"),
      bullet("F\u00F6r att visa att du bygger p\u00E5 trov\u00E4rdiga k\u00E4llor"),
      bullet("F\u00F6r att skilja dina egna tankar fr\u00E5n andras"),
      spacer(),
      italicText("S\u00E4g: \u201CN\u00E4r ni granska k\u00E4llor \u00E4r det avgn\u00E5rande att \u00E4ven VISA vilka k\u00E4llor ni anv\u00E4nder. D\u00E5 kan andra kontrollera om ert resonemang h\u00E5ller.\u201D"),
      spacer(),

      boldBodyText("F\u00F6renklad Harvard-modell (65\u201369 min)", ""),
      bodyText("Skriv p\u00E5 tavlan:"),
      spacer(),
      boldBodyText("I texten: ", "(Efternamn, \u00E5r). Exempel: \u201CAI-genererat inneh\u00E5ll har \u00F6kat kraftigt (Nygren, 2025).\u201D"),
      boldBodyText("I referenslistan: ", "Efternamn (\u00C5r). Titel. K\u00E4lla."),
      spacer(),
      bodyText("Visa 2\u20133 exempel:"),
      bullet("Nygren, T. (2025). Digital k\u00E4llkritik i praktiken. Uppsala universitet."),
      bullet("SVT Nyheter (2025). AI-genererade TikTok-konton avsl\u00F6jade. svt.se."),
      spacer(),

      boldBodyText("\u00D6vning (69\u201372 min)", ""),
      italicText("S\u00E4g: \u201CV\u00E4lj tv\u00E5 av k\u00E4llorna fr\u00E5n den guidade \u00F6vningen (bilaga 4). Skriv en Harvard-h\u00E4nvisning f\u00F6r vardera: b\u00E5de i-texten-varianten och referenslistan. Ni har 3 minuter.\u201D"),
      bodyText("Cirkulera och ge snabb \u00E5terkoppling. Vanliga fel: gl\u00F6mmer \u00E5r, gl\u00F6mmer parentes, f\u00F6rv\u00E4xlar i-texten-referens med referenslistan."),

      // --- Fas 6: Avslut ---
      heading3("Fas 6: Avslut (72\u201380 min)"),
      boldBodyText("Exit ticket (72\u201377 min): ", ""),
      bodyText("Dela ut bilaga 5 eller skriv p\u00E5 tavlan:"),
      spacer(),
      italicText("\u201CV\u00E4lj en av de fyra k\u00E4llkritiska grundfr\u00E5gorna. F\u00F6rklara med ett eget exempel hur fr\u00E5gan hj\u00E4lper dig att granska ett digitalt inneh\u00E5ll.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 meningsstartare p\u00E5 tavlan:"),
      bullet("\u201CJag v\u00E4ljer fr\u00E5gan [VEM/VARF\u00D6R/HUR/N\u00C4R] f\u00F6r att...\u201D"),
      bullet("\u201CEtt exempel \u00E4r n\u00E4r jag granskade [k\u00E4lla] och fr\u00E5gan hj\u00E4lpte mig att se att...\u201D"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r retrieval review i lektion 3."),
      spacer(),

      boldBodyText("Preview (77\u201380 min): ", ""),
      italicText("\u201CNu har ni verktygen. N\u00E4sta lektion anv\u00E4nder vi dem p\u00E5 n\u00E5got riktigt intressant: hur AI anv\u00E4nds f\u00F6r att f\u00F6rst\u00E4rka konspirationsteorier. Ni kommer att arbeta med stationsarbete d\u00E4r varje station fokuserar p\u00E5 en AI-mekanism \u2014 deepfakes, filterbubblor, AI-genererade k\u00E4llor och bottar. SIFT och grundfr\u00E5gorna kommer att beh\u00F6vas.\u201D"),

      // ========== ELEVAKTIVITETER ==========
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval review: parsamtal om trov\u00E4rdighetsmark\u00F6rer fr\u00E5n L1 (par, 4 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Guidad \u00F6vning: till\u00E4mpa SIFT p\u00E5 3 k\u00E4llor med lateral reading och omv\u00E4nd bild\u00F6kning (par, 20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Helklasssammanfattning: vilka k\u00E4llor var trov\u00E4rdiga? (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Sj\u00E4lvst\u00E4ndig \u00F6vning: granska en egen digital k\u00E4lla med grundfr\u00E5gor + SIFT (enskilt, 15 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Harvard-refererings\u00F6vning: formatera 2 k\u00E4llor (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: grundfr\u00E5ga + eget exempel (enskilt, 5 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 57 av 80 minuter (71%)"),

      // ========== DIFFERENTIERING ==========
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Grundfr\u00E5gornas \u00F6versiktskort (bilaga 1) som st\u00F6d under hela lektionen. Vid guidad \u00F6vning: b\u00F6rja med den enklaste k\u00E4llan (A). Meningsstartare vid sj\u00E4lvst\u00E4ndig \u00F6vning och exit ticket. L\u00E4raren prioriterar dessa par vid cirkulering."),
      boldBodyText("Utmaning (mot A): ", "Vid guidad \u00F6vning: resonera om vilka SIFT-steg som var mest avg\u00F6rande f\u00F6r varje k\u00E4lla och varf\u00F6r. Vid sj\u00E4lvst\u00E4ndig \u00F6vning: v\u00E4lj en sv\u00E5rbed\u00F6md k\u00E4lla. Exit ticket utan meningsstartare, med till\u00E4gget: \u201CResonera om n\u00E4r grundfr\u00E5gorna INTE r\u00E4cker \u2014 n\u00E4r beh\u00F6vs fler verktyg?\u201D"),

      // ========== MATERIAL ==========
      heading2("Material"),
      bullet("Datorer/surfplattor med internet\u00E5tkomst (en per par)"),
      bullet("Bilaga 1: De fyra k\u00E4llkritiska grundfr\u00E5gorna \u2014 \u00F6versiktskort (en per elev)"),
      bullet("Bilaga 2: SIFT-metoden steg f\u00F6r steg (en per par eller projicera)"),
      bullet("Bilaga 3: F\u00F6renklad Harvard-referering mall (en per elev)"),
      bullet("Bilaga 4: \u00D6vningsk\u00E4llor \u2014 3 k\u00E4llor f\u00F6r guidad \u00F6vning (en per par)"),
      bullet("Bilaga 5: Exit ticket L2 (en per elev)"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),
      bullet("Exit ticket-data fr\u00E5n L1 (f\u00F6r retrieval review)"),

      // ========== KOPPLING TILL KUNSKAPSKRAV ==========
      heading2("Koppling till kunskapskrav"),
      bullet("Grundfr\u00E5gorna och SIFT-metoden tr\u00E4nar f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor (m\u00E5l 1: E\u2013A)"),
      bullet("Lateral reading och omv\u00E4nd bild\u00F6kning ger konkreta digitala verifieringsverktyg (m\u00E5l 1)"),
      bullet("Den guidade \u00F6vningen med 3 k\u00E4llor av varierande trov\u00E4rdighet tr\u00E4nar systematisk granskning (m\u00E5l 1: C\u2013A)"),
      bullet("Harvard-introduktionen p\u00E5b\u00F6rjar tr\u00E4ning i k\u00E4llh\u00E4nvisning (m\u00E5l 4: E)"),
      bullet("Exit ticket m\u00E4ter f\u00F6rm\u00E5gan att koppla grundfr\u00E5ga till eget exempel (m\u00E5l 1)"),
      spacer(),
      bodyText("Den guidade \u00F6vningens scaffolding (grundfr\u00E5gekort, l\u00E4rarcirkulering) s\u00E4kerst\u00E4ller att alla elever kan n\u00E5 E-niv\u00E5. Utmaningen att resonera om n\u00E4r verktygen r\u00E4cker och inte r\u00E4cker ger A-elever m\u00F6jlighet att visa nyanserat kritiskt t\u00E4nkande."),

      // ========== KOPPLINGAR ==========
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: L1 (AI-labb, trov\u00E4rdighetsmark\u00F6rer, tredjepersons-framing)"),
      bullet("Retrieval review baseras p\u00E5 L1 exit ticket"),
      bullet("Introducerar alla verktyg som anv\u00E4nds i resten av momentet: grundfr\u00E5gor, SIFT, lateral reading, omv\u00E4nd bild\u00F6kning, Harvard-referering"),
      bullet("N\u00E4sta lektion (L3): Stationsarbete \u2014 AI + konspirationsteorier. Eleverna till\u00E4mpar L2-verktygen p\u00E5 AI-mekanismer"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r retrieval review i L3"),
      bullet("Harvard-referering byggs vidare i L3 exit ticket och L5"),

      // ========== BILAGOR ==========
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: GRUNDFR\u00C5GOR
      heading2("Bilaga 1: De fyra k\u00E4llkritiska grundfr\u00E5gorna"),
      bodyText("Anv\u00E4nd dessa fr\u00E5gor n\u00E4r du granskar ALLA typer av digitalt inneh\u00E5ll."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1200, 3913, 3913],
        rows: [
          new TableRow({
            children: [
              headerCell("Fr\u00E5ga", 1200),
              headerCell("Vad inneb\u00E4r den?", 3913),
              headerCell("Exempel", 3913),
            ],
          }),
          new TableRow({
            children: [
              cell("VEM?", 1200),
              cell("Vem \u00E4r avs\u00E4ndaren? Vad vet vi om dem? \u00C4r de experter, journalister, anonyma?", 3913),
              cell("En artikel om h\u00E4lsa \u2014 \u00E4r f\u00F6rfattaren l\u00E4kare, journalist eller anonym bloggare?", 3913),
            ],
          }),
          new TableRow({
            children: [
              cell("VARF\u00D6R?", 1200),
              cell("Varf\u00F6r publiceras detta? Vad \u00E4r syftet? Informera, \u00F6vertyga, s\u00E4lja, provocera?", 3913),
              cell("En sajt som s\u00E4ljer kosttillskott skriver om \u201Cfarliga\u201D vanliga l\u00E4kemedel.", 3913),
            ],
          }),
          new TableRow({
            children: [
              cell("HUR?", 1200),
              cell("Hur presenteras informationen? Spr\u00E5k, k\u00E4nslor, bilder, \u201Cbevis\u201D?", 3913),
              cell("Stora rubriker, dramatiskt spr\u00E5k, \u201CCHOCKERANDE\u201D \u2014 signalerar os\u00E4ker k\u00E4lla.", 3913),
            ],
          }),
          new TableRow({
            children: [
              cell("N\u00C4R?", 1200),
              cell("N\u00E4r publicerades det? \u00C4r informationen aktuell? Har kontexten f\u00F6r\u00E4ndrats?", 3913),
              cell("En artikel fr\u00E5n 2019 om \u201Ckommande pandemi\u201D \u2014 skrevs den f\u00F6re eller efter covid?", 3913),
            ],
          }),
        ],
      }),

      // BILAGA 2: SIFT
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: SIFT-metoden steg f\u00F6r steg"),
      bodyText("SIFT \u00E4r en praktisk metod f\u00F6r snabb k\u00E4llkritisk bed\u00F6mning, utvecklad av Mike Caulfield."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [800, 2000, 3113, 3113],
        rows: [
          new TableRow({
            children: [
              headerCell("Steg", 800),
              headerCell("Betydelse", 2000),
              headerCell("Vad g\u00F6r du?", 3113),
              headerCell("Tips", 3113),
            ],
          }),
          new TableRow({
            children: [
              cell("S", 800),
              cell("Stop", 2000),
              cell("Stanna upp. Reagera inte direkt. Dela inte vidare f\u00F6rr\u00E4n du granskat.", 3113),
              cell("K\u00E4nslor som ilska eller chock \u00E4r en varningssignal \u2014 det \u00E4r d\u00E5 du MEST beh\u00F6ver stanna.", 3113),
            ],
          }),
          new TableRow({
            children: [
              cell("I", 800),
              cell("Investigate the source", 2000),
              cell("\u00D6ppna en NY flik. S\u00F6k p\u00E5 avs\u00E4ndaren. Anv\u00E4nd lateral reading och omv\u00E4nd bild\u00F6kning.", 3113),
              cell("L\u00E4s INTE mer p\u00E5 sj\u00E4lva sidan \u2014 s\u00F6k p\u00E5 k\u00E4llan externt. Anv\u00E4nd Google Lens p\u00E5 bilder.", 3113),
            ],
          }),
          new TableRow({
            children: [
              cell("F", 800),
              cell("Find better coverage", 2000),
              cell("S\u00F6k p\u00E5 sj\u00E4lva p\u00E5st\u00E5endet. Vad s\u00E4ger ANDRA k\u00E4llor?", 3113),
              cell("Om bara en k\u00E4lla rapporterar n\u00E5got \u2014 var skeptisk. Trov\u00E4rdiga nyheter bekr\u00E4ftas av flera.", 3113),
            ],
          }),
          new TableRow({
            children: [
              cell("T", 800),
              cell("Trace claims", 2000),
              cell("Varifrn\u00E5n kommer p\u00E5st\u00E5endet? Finns det en prim\u00E4rk\u00E4lla?", 3113),
              cell("F\u00F6lj kedjan bak\u00E5t. Hittade du en forskningsrapport? Myndighetsk\u00E4lla? Eller bara \u201Cn\u00E5gon s\u00E4ger\u201D?", 3113),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Lateral reading: ", "Ist\u00E4llet f\u00F6r att l\u00E4sa mer p\u00E5 sj\u00E4lva sidan, \u00F6ppna en ny flik och s\u00F6k p\u00E5 avs\u00E4ndaren. Vad s\u00E4ger ANDRA om denna k\u00E4lla? Det \u00E4r det viktigaste steget i SIFT."),
      boldBodyText("Omv\u00E4nd bild\u00F6kning: ", "H\u00F6gerklicka p\u00E5 en bild \u2192 \u201CS\u00F6k med Google Lens\u201D. Eller g\u00E5 till images.google.com och ladda upp bilden. Visar om bilden anv\u00E4nds p\u00E5 andra st\u00E4llen eller \u00E4r manipulerad."),

      // BILAGA 3: HARVARD-REFERERING
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: F\u00F6renklad Harvard-referering"),
      bodyText("Anv\u00E4nd denna mall n\u00E4r du h\u00E4nvisar till k\u00E4llor i skriftliga uppgifter."),
      spacer(),

      heading3("I texten"),
      bodyText("Skriv (Efternamn, \u00E5r) efter p\u00E5st\u00E5endet."),
      spacer(),
      italicText("Exempel: \u201CAI-genererat inneh\u00E5ll har \u00F6kat kraftigt de senaste \u00E5ren (Nygren, 2025).\u201D"),
      italicText("Exempel: \u201CEnligt SVT Nyheter (2025) har fyra AI-genererade TikTok-konton avsl\u00F6jats i Sverige.\u201D"),
      spacer(),

      heading3("I referenslistan"),
      bodyText("Skriv: Efternamn (\u00C5r). Titel. K\u00E4lla."),
      spacer(),
      italicText("Nygren, T. (2025). Digital k\u00E4llkritik i praktiken. Uppsala universitet."),
      italicText("SVT Nyheter (2025). AI-genererade TikTok-konton avsl\u00F6jade. svt.se."),
      spacer(),
      bodyText("Tips:"),
      bullet("Om f\u00F6rfattaren \u00E4r en organisation, anv\u00E4nd organisationens namn"),
      bullet("Om det inte finns n\u00E5got \u00E5r, skriv (u.\u00E5.) \u2014 utan \u00E5r"),
      bullet("H\u00E5ll det enkelt \u2014 det viktigaste \u00E4r att l\u00E4saren KAN hitta k\u00E4llan"),

      // BILAGA 4: \u00D6VNINGSK\u00C4LLOR
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: \u00D6vningsk\u00E4llor f\u00F6r guidad \u00F6vning"),
      bodyText("Granska alla tre k\u00E4llor med SIFT-metoden och de fyra grundfr\u00E5gorna. Arbeta i par."),
      spacer(),

      heading3("K\u00E4lla A: \u201CAI-genererade bilder anv\u00E4nds f\u00F6r att sprida desinformation\u201D"),
      boldBodyText("Publicerad av: ", "Sveriges Radio, Ekot (2025)"),
      boldBodyText("Inneh\u00E5ll: ", "Artikeln rapporterar om hur AI-genererade bilder har anv\u00E4nts i sociala medier f\u00F6r att skapa falska nyheter. Journalisten intervjuar forskare vid G\u00F6teborgs universitet som studerat fenomenet. Artikeln inneh\u00E5ller k\u00E4llh\u00E4nvisningar till forskarrapporten och l\u00E4nkar till originalstudien."),
      boldBodyText("Bed\u00F6mning: ", "Trov\u00E4rdig \u2014 etablerad redaktion, namngiven journalist, expertk\u00E4llor, l\u00E4nk till originalforskning."),
      spacer(),

      heading3("K\u00E4lla B: \u201CSanningen om AI som media inte ber\u00E4ttar\u201D"),
      boldBodyText("Publicerad av: ", "frihetsbloggen.info (2025)"),
      boldBodyText("Inneh\u00E5ll: ", "Blogginl\u00E4gget h\u00E4vdar att \u201Cetablerade medier m\u00F6rkar\u201D hur AI redan kontrollerar nyhetsfl\u00F6det. Texten blandar verkliga h\u00E4ndelser med ol\u00E4gliga p\u00E5st\u00E5enden, anv\u00E4nder dramatiskt spr\u00E5k (\u201Cde vill inte att du ska veta\u201D) och saknar k\u00E4llh\u00E4nvisningar. Ingen f\u00F6rfattare \u00E4r namngiven."),
      boldBodyText("Bed\u00F6mning: ", "Problematisk \u2014 ok\u00E4nd avs\u00E4ndare, inga k\u00E4llor, konspirationsinspirerat spr\u00E5k, blandar fakta med \u00E5sikter."),
      spacer(),

      heading3("K\u00E4lla C: Socialt medie-inl\u00E4gg p\u00E5 X (Twitter)"),
      boldBodyText("Publicerad av: ", "@sanningsjakten99 (2025)"),
      boldBodyText("Inneh\u00E5ll: ", "\u201CBREAKING: Ny studie visar att 90% av allt nyhetsinneh\u00E5ll 2026 kommer vara AI-genererat. Medierna VILL inte att du ska veta detta. Dela innan de tar bort det!!! \ud83d\udea8\ud83d\udea8\ud83d\udea8\u201D Inl\u00E4gget har 4 500 delningar men inga k\u00E4llh\u00E4nvisningar. Ingen l\u00E4nk till studien."),
      boldBodyText("Bed\u00F6mning: ", "Otrov\u00E4rdig \u2014 anonymt konto, inga k\u00E4llor, uppmanar till snabb delning, \u00F6verdriven statistik utan bevis, konspirationsinspirerat."),
      spacer(),
      bodyText("Diskutera i paret:"),
      bullet("Vilka SIFT-steg avsl\u00F6jade mest om varje k\u00E4lla?"),
      bullet("Vilka grundfr\u00E5gor var viktigast f\u00F6r respektive k\u00E4lla?"),
      bullet("Hur skiljer sig en trov\u00E4rdig k\u00E4lla fr\u00E5n en otrov\u00E4rdig \u2014 vad \u00E4r de tydligaste tecknen?"),

      // BILAGA 5: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 5: Exit ticket \u2014 Lektion 2"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "V\u00E4lj en av de fyra k\u00E4llkritiska grundfr\u00E5gorna. F\u00F6rklara med ett eget exempel hur fr\u00E5gan hj\u00E4lper dig att granska ett digitalt inneh\u00E5ll."),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CJag v\u00E4ljer fr\u00E5gan [VEM/VARF\u00D6R/HUR/N\u00C4R] f\u00F6r att...\u201D"),
      bullet("\u201CEtt exempel \u00E4r n\u00E4r jag granskade [k\u00E4lla] och fr\u00E5gan hj\u00E4lpte mig att se att...\u201D"),
    ],
  }],
});

// ============================================================
// PPTX - Lektion 2 presentation
// ============================================================

function generatePptx() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Cornelius Agent";
  pres.title = "Lektion 2: K\u00E4llkritiska verktyg";

  // Color palette - Midnight Executive
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
  s1.addText("Lektion 2", { x: 0.8, y: 1.2, w: 8.4, h: 0.8, fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, charSpacing: 6 });
  s1.addText("K\u00E4llkritiska verktyg", { x: 0.8, y: 2.0, w: 8.4, h: 1.2, fontSize: 40, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s1.addText("Grundfr\u00E5gorna, SIFT-metoden, lateral reading och Harvard-referering", { x: 0.8, y: 3.2, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });
  s1.addText("Samh\u00E4llskunskap 3 / Internationella relationer", { x: 0.8, y: 4.7, w: 8.4, h: 0.5, fontSize: 13, fontFace: "Arial", color: TEXT_MUTED });

  // --- Slide 2: Retrieval review ---
  const s2 = pres.addSlide();
  s2.background = { color: DARK_BG };
  s2.addText("\u00C5terkallelse fr\u00E5n L1", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s2.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.4, w: 8.4, h: 2.2, fill: { color: MED_BG } });
  s2.addText([
    { text: "F\u00F6rra lektionen skapade ni AI-genererat inneh\u00E5ll.", options: { fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Vad var det som gjorde inneh\u00E5llet \u00F6vertygande?", options: { fontSize: 22, fontFace: "Arial", color: ACCENT, bold: true, breakLine: true } },
    { text: "N\u00E4mn minst tv\u00E5 trov\u00E4rdighetsmark\u00F6rer.", options: { fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 1.2, y: 1.6, w: 7.6, h: 1.8 });
  s2.addText("Diskutera med den som sitter bredvid dig. 3 minuter.", { x: 0.8, y: 4.2, w: 8.4, h: 0.5, fontSize: 16, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 3: De fyra grundfr\u00E5gorna ---
  const s3 = pres.addSlide();
  s3.background = { color: DARK_BG };
  s3.addText("De fyra k\u00E4llkritiska grundfr\u00E5gorna", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s3.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  const questions = [
    { q: "VEM?", desc: "Vem \u00E4r avs\u00E4ndaren? Expert, journalist, anonym?" },
    { q: "VARF\u00D6R?", desc: "Varf\u00F6r publiceras detta? Informera, \u00F6vertyga, s\u00E4lja?" },
    { q: "HUR?", desc: "Hur presenteras det? Spr\u00E5k, k\u00E4nslor, \u201Cbevis\u201D?" },
    { q: "N\u00C4R?", desc: "N\u00E4r publicerades det? \u00C4r det aktuellt?" },
  ];

  questions.forEach((item, i) => {
    const yPos = 1.2 + i * 1.05;
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: MED_BG } });
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.08, h: 0.9, fill: { color: ACCENT } });
    s3.addText(item.q, { x: 1.1, y: yPos, w: 1.8, h: 0.9, fontSize: 22, fontFace: "Arial", color: ACCENT, bold: true, valign: "middle", margin: 0 });
    s3.addText(item.desc, { x: 2.9, y: yPos, w: 6.0, h: 0.9, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, valign: "middle", margin: 0 });
  });

  s3.addText("Dessa fr\u00E5gor fungerar p\u00E5 ALLA digitala k\u00E4llor.", { x: 0.8, y: 5.0, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 4: SIFT-metoden ---
  const s4 = pres.addSlide();
  s4.background = { color: DARK_BG };
  s4.addText("SIFT-metoden", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s4.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  const siftSteps = [
    { letter: "S", name: "Stop", desc: "Stanna upp. Reagera inte direkt. Dela inte." },
    { letter: "I", name: "Investigate the source", desc: "S\u00F6k p\u00E5 avs\u00E4ndaren i en NY flik. Lateral reading." },
    { letter: "F", name: "Find better coverage", desc: "Vad s\u00E4ger ANDRA k\u00E4llor om samma p\u00E5st\u00E5ende?" },
    { letter: "T", name: "Trace claims", desc: "Varifrn\u00E5n kommer p\u00E5st\u00E5endet? Finns prim\u00E4rk\u00E4lla?" },
  ];

  siftSteps.forEach((step, i) => {
    const yPos = 1.2 + i * 1.05;
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: MED_BG } });
    s4.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: ACCENT } });
    s4.addText(step.letter, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: "Arial", color: DARK_BG, bold: true, align: "center", valign: "middle" });
    s4.addText(step.name, { x: 1.9, y: yPos, w: 3.2, h: 0.9, fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, valign: "middle", margin: 0 });
    s4.addText(step.desc, { x: 5.1, y: yPos, w: 3.9, h: 0.9, fontSize: 14, fontFace: "Arial", color: TEXT_LIGHT, valign: "middle", margin: 0 });
  });

  // --- Slide 5: Lateral reading ---
  const s5 = pres.addSlide();
  s5.background = { color: DARK_BG };
  s5.addText("Lateral reading \u2014 demo", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s5.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s5.addText("Det viktigaste steget i SIFT:", { x: 0.8, y: 1.2, w: 8.4, h: 0.5, fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT });

  const lrSteps = [
    "1. Du l\u00E4ser n\u00E5got p\u00E5 en ok\u00E4nd sajt",
    "2. \u00D6PPNA EN NY FLIK (l\u00E4s inte vidare p\u00E5 sajten!)",
    "3. S\u00F6k p\u00E5 avs\u00E4ndaren: vad s\u00E4ger ANDRA?",
    "4. Bed\u00F6m: Trov\u00E4rdig? Partisk? Ok\u00E4nd?",
  ];

  lrSteps.forEach((step, i) => {
    const yPos = 1.9 + i * 0.7;
    s5.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: yPos, w: 7.6, h: 0.55, fill: { color: MED_BG } });
    s5.addText(step, { x: 1.5, y: yPos, w: 7.0, h: 0.55, fontSize: 17, fontFace: "Arial", color: TEXT_WHITE, valign: "middle", margin: 0 });
  });

  s5.addText("Ist\u00E4llet f\u00F6r att l\u00E4sa mer p\u00E5 sidan \u2014 l\u00E4s vad ANDRA s\u00E4ger om sidan.", { x: 0.8, y: 4.8, w: 8.4, h: 0.5, fontSize: 15, fontFace: "Arial", color: ACCENT, italic: true });

  // --- Slide 6: Omv\u00E4nd bild\u00F6kning ---
  const s6 = pres.addSlide();
  s6.background = { color: DARK_BG };
  s6.addText("Omv\u00E4nd bild\u00F6kning / Google Lens", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s6.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s6.addText("Del av steget \u201CInvestigate the source\u201D i SIFT", { x: 0.8, y: 1.2, w: 8.4, h: 0.5, fontSize: 16, fontFace: "Arial", color: TEXT_MUTED });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.9, w: 8.4, h: 2.5, fill: { color: MED_BG } });
  s6.addText([
    { text: "S\u00E5 h\u00E4r g\u00F6r du:", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "1. H\u00F6gerklicka p\u00E5 bilden \u2192 \u201CS\u00F6k med Google Lens\u201D (Chrome)", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "2. Eller: g\u00E5 till images.google.com \u2192 klicka kameraikonen \u2192 ladda upp bilden", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Det visar:", options: { fontSize: 16, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "\u2022 Var bilden anv\u00E4nds p\u00E5 n\u00E4tet", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "\u2022 Om bilden \u00E4r tagen ur sin kontext", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "\u2022 Om bilden \u00E4r manipulerad", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 1.2, y: 2.0, w: 7.6, h: 2.2 });

  // --- Slide 7: \u00D6vningsinstruktioner ---
  const s7 = pres.addSlide();
  s7.background = { color: DARK_BG };
  s7.addText("Guidad \u00F6vning", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s7.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s7.addText([
    { text: "Ni f\u00E5r 3 k\u00E4llor av varierande trov\u00E4rdighet.", options: { fontSize: 18, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Er uppgift:", options: { fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "1. Till\u00E4mpa SIFT p\u00E5 varje k\u00E4lla", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "2. Anv\u00E4nd grundfr\u00E5gorna (VEM, VARF\u00D6R, HUR, N\u00C4R)", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "3. Praktisera lateral reading och omv\u00E4nd bild\u00F6kning", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "4. Bed\u00F6m: trov\u00E4rdig, problematisk eller otrov\u00E4rdig?", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 0.8, y: 1.2, w: 8.4, h: 3.0 });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.4, w: 8.4, h: 0.8, fill: { color: MED_BG } });
  s7.addText("Arbeta i par. Ni har 20 minuter.", { x: 0.8, y: 4.4, w: 8.4, h: 0.8, fontSize: 20, fontFace: "Arial", color: ACCENT, bold: true, align: "center", valign: "middle" });

  // --- Slide 8: Harvard-referering ---
  const s8 = pres.addSlide();
  s8.background = { color: DARK_BG };
  s8.addText("Harvard-referering \u2014 grunderna", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s8.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s8.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 4.0, h: 2.2, fill: { color: MED_BG } });
  s8.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 0.08, h: 2.2, fill: { color: ACCENT } });
  s8.addText([
    { text: "I texten:", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "(Efternamn, \u00e5r)", options: { fontSize: 20, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Exempel:", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "(Nygren, 2025)", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, italic: true } },
  ], { x: 1.1, y: 1.3, w: 3.5, h: 2.0 });

  s8.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.0, h: 2.2, fill: { color: MED_BG } });
  s8.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 0.08, h: 2.2, fill: { color: ACCENT } });
  s8.addText([
    { text: "I referenslistan:", options: { fontSize: 18, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Efternamn (\u00c5r). Titel. K\u00e4lla.", options: { fontSize: 18, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Exempel:", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "Nygren, T. (2025). Digital\nk\u00e4llkritik. Uppsala univ.", options: { fontSize: 14, fontFace: "Arial", color: TEXT_LIGHT, italic: true } },
  ], { x: 5.5, y: 1.3, w: 3.5, h: 2.0 });

  s8.addText("\u00D6vning: Formatera 2 k\u00E4llor fr\u00E5n den guidade \u00F6vningen. 3 minuter.", { x: 0.8, y: 4.2, w: 8.4, h: 0.5, fontSize: 16, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 9: Exit ticket ---
  const s9 = pres.addSlide();
  s9.background = { color: DARK_BG };
  s9.addText("Exit ticket", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s9.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s9.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: 1.5, w: 7.6, h: 2.4, fill: { color: MED_BG } });
  s9.addText([
    { text: "V\u00E4lj en av de fyra k\u00E4llkritiska grundfr\u00E5gorna.", options: { fontSize: 20, fontFace: "Arial", color: TEXT_WHITE, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "F\u00F6rklara med ett eget exempel hur fr\u00E5gan hj\u00E4lper dig att granska ett digitalt inneh\u00E5ll.", options: { fontSize: 20, fontFace: "Arial", color: ACCENT, bold: true } },
  ], { x: 1.6, y: 1.7, w: 6.8, h: 2.0 });

  s9.addText([
    { text: "\u201CJag v\u00E4ljer fr\u00E5gan [VEM/VARF\u00D6R/HUR/N\u00C4R] f\u00F6r att...\u201D", options: { fontSize: 14, fontFace: "Arial", color: TEXT_LIGHT, italic: true, breakLine: true } },
    { text: "\u201CEtt exempel \u00E4r n\u00E4r jag granskade [k\u00E4lla] och fr\u00E5gan hj\u00E4lpte mig att se att...\u201D", options: { fontSize: 14, fontFace: "Arial", color: TEXT_LIGHT, italic: true } },
  ], { x: 1.2, y: 4.2, w: 7.6, h: 1.0 });

  return pres;
}

// ============================================================
// GENERATE BOTH FILES
// ============================================================

const OUTPUT_DIR = "/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier";

async function main() {
  // Generate DOCX
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(`${OUTPUT_DIR}/lektion-2.docx`, buffer);
  console.log("lektion-2.docx skapad!");

  // Generate PPTX
  const pres = generatePptx();
  await pres.writeFile({ fileName: `${OUTPUT_DIR}/lektion-2.pptx` });
  console.log("lektion-2.pptx skapad!");
}

main().catch(err => {
  console.error("Fel vid generering:", err);
  process.exit(1);
});
