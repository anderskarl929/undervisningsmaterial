const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak
} = require("docx");

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
      // TITEL
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lektion 6: Fr\u00E5n analys till text", font: "Arial" })],
      }),
      bodyText("Summativ bed\u00F6mning \u2014 skriftlig analys med k\u00E4llh\u00E4nvisning"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "75 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna skriva en analytisk text med korrekt k\u00E4llh\u00E4nvisning d\u00E4r k\u00E4llor granskas och v\u00E4rderas. Den h\u00E4r texten \u00E4r det prim\u00E4ra summativa bed\u00F6mningsunderlaget f\u00F6r momentet."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av k\u00E4llkritiska verktyg (m\u00E5l 1)"),
      bullet("K\u00E4llh\u00E4nvisa korrekt enligt Harvard-systemet (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men och k\u00E4llh\u00E4nvisar med viss s\u00E4kerhet."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men och k\u00E4llh\u00E4nvisar korrekt i olika texttyper."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men som visar p\u00E5 komplexa samband, och k\u00E4llh\u00E4nvisar med s\u00E4kerhet i olika sammanhang."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 5 \u2014 identifiera vanliga sv\u00E5righeter med k\u00E4llh\u00E4nvisning (introducerades i lektion 3)"),
      bullet("Skriv ut skrivuppgiften med de tre fr\u00E5gealternativen (en per elev, se bilaga 1)"),
      bullet("Skriv ut Harvard-guiden (en per elev, se bilaga 2)"),
      bullet("Skriv ut kamratgranskningschecklistan (en per elev, se bilaga 3)"),
      bullet("Skriv ut k\u00E4llf\u00F6rslaget med l\u00E4nkar (en per elev eller projicera, se bilaga 4)"),
      bullet("Skriv ut bed\u00F6mningsmatrisen (en per elev, se bilaga 7)"),
      bullet("S\u00E4kerst\u00E4ll att eleverna har tillg\u00E5ng till datorer f\u00F6r skrivande och k\u00E4llh\u00E4mtning"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),

      // TIDSPLANERING
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
              cell("Uppstart", 1300),
              cell("Retrieval practice + uppgiftsintro", 2200),
              cell("Snabbskrivning (3 min): \u201CSkiv ner tre saker du m\u00E5ste t\u00E4nka p\u00E5 n\u00E4r du k\u00E4llh\u00E4nvisar.\u201D Genomg\u00E5ng. Sedan: presentera dagens uppgift \u2014 en kort analytisk text med Harvard-referenser.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201320 min", 1000),
              cell("Instruktion", 1300),
              cell("Harvard-genomg\u00E5ng + exempeltext", 2200),
              cell("G\u00E5 igenom Harvard-systemet med konkreta exempel (se bilaga 2). Visa exempeltexten (se bilaga 5) och identifiera gemensamt: Var \u00E4r h\u00E4nvisningarna? Hur ser referenslistan ut? Hur syns k\u00E4llkritik i texten?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("20\u201355 min", 1000),
              cell("Bearbetning", 1300),
              cell("Skrivarbete", 2200),
              cell("Eleverna v\u00E4ljer en av tre fr\u00E5gor (se bilaga 1) och skriver en analytisk text (300\u2013500 ord) med minst 2\u20133 k\u00E4llor och Harvard-h\u00E4nvisningar. K\u00E4llf\u00F6rslag finns i bilaga 4. L\u00E4raren cirkulerar och ger individuell handledning.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("55\u201365 min", 1000),
              cell("Bearbetning", 1300),
              cell("Kamratgranskning", 2200),
              cell("Eleverna byter text med en klasskamrat. Granskning med checklista (se bilaga 3): k\u00E4llh\u00E4nvisning korrekt? K\u00E4llkritik synlig? Argumentation tydlig? Muntlig \u00E5terkoppling i par.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("65\u201370 min", 1000),
              cell("Bearbetning", 1300),
              cell("Revidering efter feedback", 2200),
              cell("Eleverna f\u00E5r 5 minuter att revidera sin text utifr\u00E5n kamratens \u00E5terkoppling. Fokus p\u00E5 k\u00E4llh\u00E4nvisningar och k\u00E4llkritik.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("70\u201375 min", 1000),
              cell("Summering", 1300),
              cell("Exit ticket + fram\u00E5tkoppling", 2200),
              cell("Exit ticket: \u201CG\u00F6r en korrekt Harvard-h\u00E4nvisning (b\u00E5de i text och i referenslista) till en valfri k\u00E4lla du anv\u00E4nde idag.\u201D Fram\u00E5tkoppling till lektion 7.", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart (0\u20138 min)"),
      boldBodyText("Retrieval practice (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar \u2014 skriv ner tre saker du m\u00E5ste t\u00E4nka p\u00E5 n\u00E4r du k\u00E4llh\u00E4nvisar. Ni har 2 minuter.\u201D"),
      bodyText("Samla svar muntligt. F\u00F6rv\u00E4ntade svar: f\u00F6rfattare, \u00E5r, titel, var i texten man h\u00E4nvisar, referenslista. Fyll i luckor."),
      spacer(),
      boldBodyText("Uppgiftsintro (5\u20138 min): ", ""),
      italicText("S\u00E4g: \u201CDe senaste lektionerna har ni analyserat muntligt \u2014 i seminarier och debatter. Idag ska ni visa samma f\u00F6rm\u00E5ga skriftligt. Det h\u00E4r \u00E4r er chans att visa vad ni kan. Den h\u00E4r texten \u00E4r ert bed\u00F6mningsunderlag f\u00F6r momentet. Ni ska skriva en kort analytisk text d\u00E4r ni anv\u00E4nder k\u00E4llor med korrekta h\u00E4nvisningar. Ni har tillg\u00E5ng till k\u00E4llor och hj\u00E4lpmedel \u2014 anv\u00E4nd dem.\u201D"),
      bodyText("Visa l\u00E4randem\u00E5let p\u00E5 tavlan."),

      heading3("Instruktion: Harvard-genomg\u00E5ng (8\u201320 min)"),
      bodyText("Dela ut Harvard-guiden (bilaga 2). G\u00E5 igenom systemet med max 7 minuter f\u00F6rel\u00E4sning, sedan 5 minuter exempeltext."),
      spacer(),
      boldBodyText("Grundprincipen (2 min): ", "\u201CHarvard-systemet bygger p\u00E5 tv\u00E5 delar: en kort h\u00E4nvisning i texten (f\u00F6rfattare, \u00E5r) och en fullst\u00E4ndig referenslista i slutet.\u201D"),
      spacer(),
      bodyText("Visa p\u00E5 tavlan/sk\u00E4rm:"),
      bullet("H\u00E4nvisning i text: (Svensson, 2024, s. 45)"),
      bullet("I referenslistan: Svensson, A. (2024). K\u00E4llkritik i den digitala tids\u00E5ldern. Stockholm: Natur & Kultur."),
      spacer(),
      boldBodyText("Olika k\u00E4lltyper (3 min): ", "G\u00E5 igenom hur man h\u00E4nvisar till: bok, webbsida, rapport, nyhetsartikel. Anv\u00E4nd guiden (bilaga 2). Betona: det viktigaste \u00E4r att man \u00E4r konsekvent."),
      spacer(),
      boldBodyText("Direktcitat vs. omskrivning (2 min): ", "Visa skillnaden. Direktcitat: \u201CEnligt Svensson (2024, s. 45) \u00E4r \u2018k\u00E4llkritik en demokratisk grundkompetens\u2019.\u201D Omskrivning: \u201CSvensson (2024) menar att k\u00E4llkritik \u00E4r en grundl\u00E4ggande f\u00E4rdighet f\u00F6r demokratiskt deltagande.\u201D"),
      spacer(),
      boldBodyText("Exempeltext (5 min): ", "Dela ut eller projicera exempeltexten (bilaga 5). L\u00E4s h\u00F6gt. Fr\u00E5ga eleverna:"),
      bullet("\u201CVar i texten finns k\u00E4llh\u00E4nvisningarna?\u201D"),
      bullet("\u201CHur ser referenslistan ut \u2014 vad st\u00E5r med?\u201D"),
      bullet("\u201CVar syns k\u00E4llkritiken \u2014 d\u00E4r f\u00F6rfattaren v\u00E4rderar k\u00E4llorna, inte bara \u00E5terberar\u00E4ttar?\u201D"),

      heading3("Bearbetning: Skrivarbete (20\u201355 min)"),
      bodyText("Dela ut skrivuppgiften (bilaga 1) och k\u00E4llf\u00F6rslaget (bilaga 4). Dela ut bed\u00F6mningsmatrisen (bilaga 7). Eleverna v\u00E4ljer EN fr\u00E5ga och skriver p\u00E5 dator."),
      spacer(),
      italicText("Instruktion: \u201CDet h\u00E4r \u00E4r er chans att visa vad ni kan. Den h\u00E4r texten \u00E4r ert bed\u00F6mningsunderlag f\u00F6r hela momentet. V\u00E4lj en av de tre fr\u00E5gorna. Skriv en analytisk text p\u00E5 300\u2013500 ord d\u00E4r ni anv\u00E4nder minst 2\u20133 k\u00E4llor med korrekta Harvard-h\u00E4nvisningar. Ni hittar k\u00E4llf\u00F6rslag i bilaga 4, men ni f\u00E5r \u00E4ven s\u00F6ka egna k\u00E4llor. Visa k\u00E4llkritisk f\u00F6rm\u00E5ga \u2014 det r\u00E4cker inte att bara \u00E5terbera\u00E4tta vad k\u00E4llorna s\u00E4ger. Kommentera k\u00E4llornas trov\u00E4rdighet och j\u00E4mf\u00F6r dem. Titta p\u00E5 bed\u00F6mningsmatrisen s\u00E5 ni vet vad som f\u00F6rv\u00E4ntas. Ni har 35 minuter.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 skriv p\u00E5 tavlan:"),
      bullet("Inledning: Presentera din fr\u00E5ga och din huvudpo\u00E4ng (2\u20133 meningar)"),
      bullet("Analys: Anv\u00E4nd k\u00E4llor f\u00F6r att st\u00F6dja ditt resonemang. V\u00E4rdera k\u00E4llorna."),
      bullet("Avslutning: Sammanfatta och dra en slutsats (2\u20133 meningar)"),
      bullet("Referenslista: Lista alla k\u00E4llor du h\u00E4nvisar till"),
      spacer(),
      bodyText("Cirkulera under skrivarbetet. Fokusera p\u00E5:"),
      bullet("Elever som inte kommer ig\u00E5ng: \u201CVilken fr\u00E5ga valde du? Vad tycker du sj\u00E4lv om svaret? B\u00F6rja d\u00E4r.\u201D"),
      bullet("K\u00E4llh\u00E4nvisningar: \u201CJag ser att du anv\u00E4nder en k\u00E4lla h\u00E4r \u2014 har du lagt in parentesen?\u201D"),
      bullet("K\u00E4llkritik: \u201CDu \u00E5terbera\u00E4ttar k\u00E4llan bra, men vad tycker du om k\u00E4llans trov\u00E4rdighet?\u201D"),
      bullet("Snabba elever: \u201CKan du l\u00E4gga till en tredje k\u00E4lla och j\u00E4mf\u00F6ra den med de andra?\u201D"),

      heading3("Bearbetning: Kamratgranskning (55\u201365 min)"),
      italicText("S\u00E4g: \u201CStopp \u2014 pennorna ner. Nu byter ni text med personen bredvid er. Ni ska l\u00E4sa varandras texter och ge \u00E5terkoppling med hj\u00E4lp av checklistan. Det h\u00E4r \u00E4r inte r\u00E4ttning \u2014 det \u00E4r hj\u00E4lp att bli b\u00E4ttre.\u201D"),
      spacer(),
      bodyText("Dela ut kamratgranskningschecklistan (bilaga 3). Eleverna arbetar i par."),
      spacer(),
      italicText("Instruktion: \u201CL\u00E4s texten noggrant. Fyll i checklistan. N\u00E4r ni \u00E4r klara, ber\u00E4tta f\u00F6r varandra: tv\u00E5 saker som fungerar bra och en sak som kan f\u00F6rb\u00E4ttras. Ni har 10 minuter \u2014 ca 5 minuter per text.\u201D"),
      spacer(),
      bodyText("Cirkulera och lyssna p\u00E5 samtalen. Om \u00E5terkopplingen \u00E4r ytlig, st\u00E4ll f\u00F6ljdfr\u00E5gor:"),
      bullet("\u201C\u00C4r h\u00E4nvisningarna i texten korrekta? Kolla med guiden.\u201D"),
      bullet("\u201CVar i texten syns det att f\u00F6rfattaren har t\u00E4nkt kritiskt om k\u00E4llan?\u201D"),

      heading3("Bearbetning: Revidering efter feedback (65\u201370 min)"),
      italicText("S\u00E4g: \u201CNu har ni f\u00E5tt \u00E5terkoppling. Ni f\u00E5r 5 minuter att g\u00F6ra justeringar i er text utifr\u00E5n det ni h\u00F6rde. Fokusera p\u00E5 k\u00E4llh\u00E4nvisningar och k\u00E4llkritik \u2014 det \u00E4r d\u00E4r det syns mest.\u201D"),
      spacer(),
      bodyText("Eleverna \u00F6ppnar sina texter igen och reviderar. L\u00E4raren cirkulerar och st\u00F6ttar vid behov."),

      heading3("Summering: Exit ticket + fram\u00E5tkoppling (70\u201375 min)"),
      boldBodyText("Exit ticket (3 min): ", ""),
      italicText("\u201CG\u00F6r en korrekt Harvard-h\u00E4nvisning till en av k\u00E4llorna du anv\u00E4nde idag. Skriv: (1) hur h\u00E4nvisningen ser ut i l\u00F6pande text, och (2) hur den ser ut i referenslistan.\u201D"),
      spacer(),
      bodyText("Samla in texterna och exit tickets."),
      spacer(),
      boldBodyText("Fram\u00E5tkoppling (2 min): ", ""),
      italicText("\u201CNu har ni visat vad ni kan. N\u00E4sta lektion \u00E4r momentets sista \u2014 d\u00E5 knyter vi ihop allt och reflekterar \u00F6ver vad ni l\u00E4rt er. Ingen bed\u00F6mning n\u00E4sta g\u00E5ng \u2014 bara reflektion och avslutning.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: tre saker om k\u00E4llh\u00E4nvisning (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Identifiera h\u00E4nvisningar i exempeltext (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Skrivarbete: analytisk text med k\u00E4llh\u00E4nvisningar (enskilt, 35 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Kamratgranskning med checklista + muntlig \u00E5terkoppling (par, 10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Revidering av text efter kamratfeedback (enskilt, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: korrekt Harvard-h\u00E4nvisning (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 60 av 75 minuter (80%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Harvard-guiden (bilaga 2) med tydliga exempel ligger framme under hela skrivarbetet. Scaffolding p\u00E5 tavlan visar textens struktur steg f\u00F6r steg. K\u00E4llf\u00F6rslaget (bilaga 4) ger f\u00E4rdiga k\u00E4llor s\u00E5 eleverna slipper s\u00F6ka sj\u00E4lva. Exempeltexten (bilaga 5) visar hur en f\u00E4rdig text kan se ut. Kamratgranskningschecklistan har tydliga ja/nej-fr\u00E5gor. L\u00E4raren prioriterar dessa elever vid cirkulering: \u201CTitta p\u00E5 exemplet i guiden \u2014 ser din h\u00E4nvisning likadan ut?\u201D"),
      boldBodyText("Utmaning (mot A): ", "Eleverna f\u00F6rv\u00E4ntas s\u00F6ka egna k\u00E4llor ut\u00F6ver f\u00F6rslagen. V\u00E4lj fr\u00E5ga 3 (mest komplex). Utmana: \u201CJ\u00E4mf\u00F6r k\u00E4llor som s\u00E4ger olika saker \u2014 vems perspektiv \u00E4r mest trov\u00E4rdigt och varf\u00F6r?\u201D Vid kamratgranskning: ge \u00E4ven \u00E5terkoppling p\u00E5 argumentationsstruktur och nyanser. Exit ticket utan st\u00F6d, med till\u00E4gget: \u201CF\u00F6rklara ocks\u00E5 varf\u00F6r korrekt k\u00E4llh\u00E4nvisning \u00E4r viktigt i en tid av AI-genererat inneh\u00E5ll.\u201D"),

      // MATERIAL
      heading2("Material"),
      bullet("Skrivuppgift med tre fr\u00E5gealternativ (se bilaga 1)"),
      bullet("Harvard-guide med exempel f\u00F6r bok, webbsida, rapport, nyhetsartikel (se bilaga 2)"),
      bullet("Kamratgranskningschecklista (se bilaga 3)"),
      bullet("K\u00E4llf\u00F6rslag med l\u00E4nkar (se bilaga 4)"),
      bullet("Exempeltext med Harvard-h\u00E4nvisningar (se bilaga 5)"),
      bullet("Bed\u00F6mningsmatris E/C/A (se bilaga 7)"),
      bullet("Datorer f\u00F6r skrivarbete och k\u00E4lls\u00F6kning"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Skrivuppgiften \u00E4r den prim\u00E4ra summativa bed\u00F6mningen \u2014 bed\u00F6mer f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor skriftligt fr\u00E5n enkla omd\u00F6men (E) till v\u00E4lgrundade och nyanserade resonemang med komplexa samband (A) (m\u00E5l 1)"),
      bullet("Harvard-h\u00E4nvisningarna bed\u00F6mer korrekt k\u00E4llhantering \u2014 fr\u00E5n \u201Cmed viss s\u00E4kerhet\u201D (E) till \u201Cmed s\u00E4kerhet i olika sammanhang\u201D (A) (m\u00E5l 4)"),
      bullet("Kamratgranskningen ger formativ \u00E5terkoppling innan slutlig inl\u00E4mning"),
      bullet("Revideringsmomentet ger eleven chans att f\u00F6rb\u00E4ttra sin text utifr\u00E5n \u00E5terkoppling"),
      bullet("Bed\u00F6mningsmatrisen (bilaga 7) g\u00F6r bed\u00F6mningskriterierna transparenta f\u00F6r eleven"),
      spacer(),
      bodyText("Scaffolding-strukturerna (Harvard-guide, exempeltext, k\u00E4llf\u00F6rslag, steg p\u00E5 tavlan) s\u00E4kerst\u00E4ller att alla elever kan n\u00E5 E-niv\u00E5. De \u00F6ppna fr\u00E5gorna och m\u00F6jligheten att s\u00F6ka egna k\u00E4llor ger A-elever utrymme att visa nyanserad analys och s\u00E4ker k\u00E4llhantering."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1\u20133 (k\u00E4llkritiska verktyg, intro till k\u00E4llh\u00E4nvisning), lektion 4\u20135 (muntlig till\u00E4mpning)"),
      bullet("N\u00E4sta lektion: Lektion 7 (reflektion och avslutning) \u2014 ingen bed\u00F6mning, fokus p\u00E5 syntes och metakognition"),
      bullet("Skrivuppgiften \u00E4r momentets prim\u00E4ra summativa bed\u00F6mningsunderlag"),
      bullet("Exit ticket-data visar om eleverna \u00E4r redo f\u00F6r sj\u00E4lvst\u00E4ndigt k\u00E4llarbete"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: SKRIVUPPGIFT
      heading2("Bilaga 1: Skrivuppgift \u2014 Analytisk text med k\u00E4llh\u00E4nvisning"),
      spacer(),
      boldBodyText("Uppgift: ", "V\u00E4lj EN av fr\u00E5gorna nedan. Skriv en analytisk text p\u00E5 300\u2013500 ord d\u00E4r du:"),
      bullet("Besvarar fr\u00E5gan med ett tydligt resonemang"),
      bullet("Anv\u00E4nder minst 2\u20133 k\u00E4llor med korrekta Harvard-h\u00E4nvisningar"),
      bullet("Visar k\u00E4llkritisk f\u00F6rm\u00E5ga \u2014 inte bara \u00E5terbera\u00E4tta, utan \u00E4ven v\u00E4rdera k\u00E4llornas trov\u00E4rdighet"),
      bullet("Avslutar med en referenslista enligt Harvard"),
      spacer(),

      heading3("Fr\u00E5ga 1: Hur p\u00E5verkar AI-genererat inneh\u00E5ll v\u00E5r m\u00F6jlighet att granska information?"),
      bodyText("Diskutera hur AI-genererade texter, bilder och videor (deepfakes) g\u00F6r det sv\u00E5rare att avg\u00F6ra vad som \u00E4r sant. Vilka risker inneb\u00E4r det f\u00F6r samh\u00E4llet? Vilka verktyg och strategier kan hj\u00E4lpa oss?"),
      spacer(),

      heading3("Fr\u00E5ga 2: Varf\u00F6r sprids konspirationsteorier s\u00E5 effektivt p\u00E5 sociala medier?"),
      bodyText("Analysera vilka mekanismer (algoritmer, ekkammare, k\u00E4nslor) som g\u00F6r att konspirationsteorier sprids snabbt och n\u00E5r m\u00E5nga. Anv\u00E4nd konkreta exempel."),
      spacer(),

      heading3("Fr\u00E5ga 3: Hur kan AI anv\u00E4ndas f\u00F6r att b\u00E5de sprida och motverka desinformation?"),
      bodyText("Resonera om AI:s dubbla roll: som verktyg f\u00F6r att skapa och sprida desinformation, men ocks\u00E5 som verktyg f\u00F6r att uppt\u00E4cka och motverka den. J\u00E4mf\u00F6r perspektiv och dra egna slutsatser."),
      spacer(),

      bodyText("Bed\u00F6mningsfokus:"),
      bullet("K\u00E4llkritik (m\u00E5l 1): Hur v\u00E4l v\u00E4rderar du k\u00E4llorna? E = enkla omd\u00F6men, C = v\u00E4lgrundade omd\u00F6men, A = v\u00E4lgrundade och nyanserade omd\u00F6men med komplexa samband"),
      bullet("K\u00E4llh\u00E4nvisning (m\u00E5l 4): Hur korrekt h\u00E4nvisar du? E = med viss s\u00E4kerhet, C = korrekt i olika texttyper, A = med s\u00E4kerhet i olika sammanhang"),

      // BILAGA 2: HARVARD-GUIDE
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Guide till Harvard-referenssystemet"),
      spacer(),
      bodyText("Harvard-systemet anv\u00E4nds p\u00E5 de flesta svenska universitet. Det bygger p\u00E5 tv\u00E5 delar:"),
      bullet("K\u00E4llh\u00E4nvisning i texten (kort: f\u00F6rfattare, \u00E5r)"),
      bullet("Referenslista i slutet av texten (fullst\u00E4ndig information)"),
      spacer(),

      heading3("1. H\u00E4nvisning i l\u00F6pande text"),
      bodyText("N\u00E4r du h\u00E4nvisar till en k\u00E4lla i texten, ange f\u00F6rfattarens efternamn och \u00E5rtal inom parentes. L\u00E4gg till sidnummer vid direktcitat."),
      spacer(),
      boldBodyText("Omskrivning: ", "\u201CEnligt forskning har AI-genererat inneh\u00E5ll blivit allt sv\u00E5rare att skilja fr\u00E5n verkligt material (Internetstiftelsen, 2025).\u201D"),
      boldBodyText("Direktcitat: ", "\u201CInternetstiftelsen (2025, s. 12) konstaterar att \u201897 procent av svenskarna anv\u00E4nder internet\u201D.\u201D"),
      boldBodyText("Tv\u00E5 f\u00F6rfattare: ", "(Str\u00F6mb\u00E4ck & Esser, 2023)"),
      boldBodyText("Tre eller fler: ", "(Svensson m.fl., 2024)"),
      boldBodyText("Organisation som f\u00F6rfattare: ", "(Myndigheten f\u00F6r psykologiskt f\u00F6rsvar, 2025)"),
      spacer(),

      heading3("2. Referenslista \u2014 s\u00E5 skriver du"),
      bodyText("Referenslistan placeras sist i texten. K\u00E4llorna listas i bokstavsordning efter f\u00F6rfattarens efternamn. Varje k\u00E4lltyp har ett eget format:"),
      spacer(),

      boldBodyText("Bok:", ""),
      bodyText("Efternamn, Initialer. (\u00C5r). Titel i kursiv stil. Utgivningsort: F\u00F6rlag."),
      italicText("Exempel: Str\u00F6mb\u00E4ck, J. (2023). Konspirationsteorier och desinformation. Stockholm: Carlsson."),
      spacer(),

      boldBodyText("Webbsida:", ""),
      bodyText("Organisation/F\u00F6rfattare. (\u00C5r). Titel. H\u00E4mtad [datum] fr\u00E5n [URL]"),
      italicText("Exempel: Internetstiftelsen. (2025). Svenskarna och internet 2025. H\u00E4mtad 2026-03-10 fr\u00E5n https://svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),
      spacer(),

      boldBodyText("Rapport:", ""),
      bodyText("Organisation. (\u00C5r). Titel i kursiv stil. Ort: Organisation."),
      italicText("Exempel: Myndigheten f\u00F6r psykologiskt f\u00F6rsvar. (2025). AI och informationsp\u00E5verkan. Karlstad: MPF."),
      spacer(),

      boldBodyText("Nyhetsartikel (online):", ""),
      bodyText("Efternamn, Initialer. (\u00C5r, dag m\u00E5nad). Artikelns titel. Tidningens namn. H\u00E4mtad [datum] fr\u00E5n [URL]"),
      italicText("Exempel: Sj\u00F6lander, J. (2024, 15 oktober). S\u00E5 lurar deepfakes v\u00E5ra \u00F6gon. SVT Nyheter. H\u00E4mtad 2026-03-10 fr\u00E5n https://svt.se/nyheter/"),
      spacer(),

      heading3("3. Vanliga misstag att undvika"),
      bullet("Gl\u00F6mma \u00E5rtalet i parentesen"),
      bullet("Anv\u00E4nda enbart URL:er utan f\u00F6rfattare och \u00E5rtal"),
      bullet("Blanda olika referenssystem (Harvard och fotnoter)"),
      bullet("Ha k\u00E4llor i texten som saknas i referenslistan (eller tv\u00E4rtom)"),
      bullet("Gl\u00F6mma att s\u00E4tta ut citattecken vid direktcitat"),

      // BILAGA 3: KAMRATGRANSKNING
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Kamratgranskningschecklista"),
      spacer(),
      boldBodyText("Granskare: ", "_____________________ Skribent: _____________________"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [5500, 1000, 1000, 1526],
        rows: [
          new TableRow({
            children: [
              headerCell("Fr\u00E5ga", 5500),
              headerCell("Ja", 1000),
              headerCell("Nej", 1000),
              headerCell("Kommentar", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00C4LLH\u00C4NVISNING", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Finns det h\u00E4nvisningar i texten (f\u00F6rfattare, \u00E5r)?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Anv\u00E4nds minst 2\u20133 olika k\u00E4llor?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Finns en referenslista i slutet?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("St\u00E4mmer h\u00E4nvisningarna i texten \u00F6verens med referenslistan?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Ser referenserna ut enligt Harvard-mallen?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00C4LLKRITIK", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("V\u00E4rderar skribenten k\u00E4llornas trov\u00E4rdighet (inte bara \u00E5terbera\u00E4ttar)?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Anv\u00E4nds k\u00E4llkritiska begrepp (t.ex. avs\u00E4ndare, syfte, trov\u00E4rdighet)?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("J\u00E4mf\u00F6rs k\u00E4llor med varandra?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("ARGUMENTATION", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Finns en tydlig inledning med huvudpo\u00E4ng?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("H\u00E4nger resonemanget ihop logiskt?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
          new TableRow({
            children: [
              cell("Dras en slutsats i avslutningen?", 5500),
              cell("", 1000),
              cell("", 1000),
              cell("", 1526),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Tv\u00E5 saker som fungerar bra: ", ""),
      bodyText("1. _____________________________________________________________"),
      bodyText("2. _____________________________________________________________"),
      spacer(),
      boldBodyText("En sak som kan f\u00F6rb\u00E4ttras: ", ""),
      bodyText("_____________________________________________________________"),

      // BILAGA 4: KÄLLFÖRSLAG
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: K\u00E4llf\u00F6rslag med l\u00E4nkar"),
      bodyText("H\u00E4r finns k\u00E4llor du kan anv\u00E4nda i din text. Du f\u00E5r \u00E4ven s\u00F6ka egna k\u00E4llor \u2014 men kom ih\u00E5g att v\u00E4rdera dem k\u00E4llkritiskt!"),
      spacer(),

      heading3("AI, deepfakes och desinformation"),
      boldBodyText("K\u00E4lla 1: ", "Internetstiftelsen (2025). Svenskarna och internet 2025. Rapport om svenskarnas internetanv\u00E4ndning, AI och digitala utmaningar."),
      bodyText("L\u00E4nk: https://svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Internetstiftelsen \u00E4r en oberoende organisation med \u00E5rlig statistisk rapport."),
      spacer(),

      boldBodyText("K\u00E4lla 2: ", "Myndigheten f\u00F6r psykologiskt f\u00F6rsvar (2025). AI och informationsp\u00E5verkan. Temabibliotek om hur AI anv\u00E4nds i p\u00E5verkanskampanjer."),
      bodyText("L\u00E4nk: https://mpf.se/kunskap-och-stod/temabibliotek/ai-och-informationspaverkan"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Svensk myndighet med uppdrag att st\u00E4rka samh\u00E4llets motst\u00E5ndskraft mot informationsp\u00E5verkan."),
      spacer(),

      boldBodyText("K\u00E4lla 3: ", "Uppsala universitet (2026). \u201CGenerativ AI har sju tydliga roller i arbetet mot desinformation.\u201D Forskningsnyhet om AI:s dubbla roll."),
      bodyText("L\u00E4nk: https://www.uu.se/press/pressmeddelanden/2026/2026-02-16-generativ-ai-har-sju-tydliga-roller-i-arbetet-mot-desinformation"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Akademisk forskning fr\u00E5n ett svenskt universitet."),
      spacer(),

      boldBodyText("K\u00E4lla 4: ", "Riksdagen (2025). Motion 2025/26:1815: M\u00E4rkning av AI-genererat material f\u00F6r att motverka desinformation och deepfakes."),
      bodyText("L\u00E4nk: https://www.riksdagen.se/sv/dokument-och-lagar/dokument/motion/markning-av-ai-genererat-material-for-att-motverka_hd021815/"),
      italicText("Trov\u00E4rdighet: H\u00F6g som prim\u00E4rk\u00E4lla. OBS: En motion \u00E4r ett f\u00F6rslag, inte ett beslut \u2014 viktigt att skilja p\u00E5."),
      spacer(),

      boldBodyText("K\u00E4lla 5: ", "MSB (2023). Generativ AI, digital resiliens och civil beredskap: en kunskaps\u00F6versikt. Forskningsrapport."),
      bodyText("L\u00E4nk: https://www.msb.se/sv/publikationer/generativ-ai-digital-resiliens-och-civil-beredskap-en-kunskapsoversikt/"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Statlig myndighet med forskningsuppdrag."),
      spacer(),

      heading3("Konspirationsteorier och spridningsmekanismer"),
      boldBodyText("K\u00E4lla 6: ", "Forskning & Framsteg (2021). \u201CS\u00E5 k\u00E4nner du igen en konspirationsteori.\u201D Popul\u00E4rvetenskaplig artikel."),
      bodyText("L\u00E4nk: https://fof.se/artikel/2021/7/sa-kanner-du-igen-en-konspirationsteori/"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Popul\u00E4rvetenskaplig tidskrift som bygger p\u00E5 forskning."),
      spacer(),

      boldBodyText("K\u00E4lla 7: ", "G\u00F6teborgs universitet / Str\u00F6mb\u00E4ck, J. (2022). \u201CKonspiratorisk l\u00E4ggning hos svenska folket.\u201D Forskningsartikel i SOM-institutets rapport."),
      bodyText("L\u00E4nk: https://www.gu.se/sites/default/files/2022-06/073-086%20Jesper%20Str%C3%B6mb%C3%A4ck_0.pdf"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Akademisk forskning med tydlig metod och statistik."),
      spacer(),

      boldBodyText("K\u00E4lla 8: ", "Internetstiftelsen \u2014 Internetkunskap. K\u00E4llkritik och AI-genererade bilder."),
      bodyText("L\u00E4nk: https://internetkunskap.se/kallkritik/"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Pedagogiskt material fr\u00E5n Internetstiftelsen."),
      spacer(),

      heading3("K\u00E4llkritik \u2014 allm\u00E4nt"),
      boldBodyText("K\u00E4lla 9: ", "Krisinformation.se. K\u00E4llkritik. Information fr\u00E5n svenska myndigheter om hur man granskar k\u00E4llor."),
      bodyText("L\u00E4nk: https://krisinformation.se/detta-gor-samhallet/kallkritik/"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Samarbete mellan svenska myndigheter."),
      spacer(),

      boldBodyText("K\u00E4lla 10: ", "K\u00E4llkritikbyr\u00E5n. Sveriges faktagranskare \u2014 granskar p\u00E5st\u00E5enden i svensk debatt."),
      bodyText("L\u00E4nk: https://kallkritikbyran.se"),
      italicText("Trov\u00E4rdighet: H\u00F6g. Oberoende faktagranskare som f\u00F6ljer internationella principer (IFCN)."),
      spacer(),

      bodyText("Tips: Anv\u00E4nd g\u00E4rna en blandning av k\u00E4lltyper (myndighet, forskning, nyheter) \u2014 det st\u00E4rker din analys."),

      // BILAGA 5: EXEMPELTEXT
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 5: Exempeltext \u2014 Analytisk text med Harvard-h\u00E4nvisningar"),
      spacer(),
      italicText("Det h\u00E4r \u00E4r ett kort exempel p\u00E5 hur en analytisk text med k\u00E4llh\u00E4nvisningar kan se ut. Texten \u00E4r f\u00F6rkortad f\u00F6r att visa strukturen."),
      spacer(),

      heading3("Kan vi lita p\u00E5 det vi ser? AI-genererat inneh\u00E5ll och k\u00E4llkritikens nya utmaningar"),
      spacer(),
      bodyText("I takt med att AI-tekniken utvecklas blir det allt sv\u00E5rare att avg\u00F6ra vad som \u00E4r \u00E4kta och vad som \u00E4r fabricerat. Enligt Internetstiftelsen (2025) anv\u00E4nder 97 procent av svenskarna internet, och allt fler kommer i kontakt med AI-genererat inneh\u00E5ll i sin vardag. Den h\u00E4r texten unders\u00F6ker hur AI-genererat inneh\u00E5ll p\u00E5verkar v\u00E5r m\u00F6jlighet att granska information."),
      spacer(),
      bodyText("Myndigheten f\u00F6r psykologiskt f\u00F6rsvar (2025) varnar f\u00F6r att AI g\u00F6r det enklare f\u00F6r illvilliga akt\u00F6rer att skapa trov\u00E4rdigt men falskt material. Det kan handla om deepfakes \u2014 manipulerade videor d\u00E4r politiker s\u00E4ger saker de aldrig sagt \u2014 eller AI-genererade nyhetsartiklar som ser professionella ut men saknar verklig journalistisk grund. K\u00E4llan \u00E4r en svensk myndighet med specifikt uppdrag att bek\u00E4mpa informationsp\u00E5verkan, vilket g\u00F6r den s\u00E4rskilt relevant och trov\u00E4rdig i det h\u00E4r sammanhanget."),
      spacer(),
      bodyText("Samtidigt finns det ljuspunkter. Forskare vid Uppsala universitet har identifierat att AI \u00E4ven kan anv\u00E4ndas f\u00F6r att motverka desinformation, exempelvis genom automatisk faktagranskning (Uppsala universitet, 2026). Det inneb\u00E4r att samma teknik som skapar problemet ocks\u00E5 kan vara en del av l\u00F6sningen \u2014 men det kr\u00E4ver att vi som medborgare utvecklar v\u00E5r k\u00E4llkritiska f\u00F6rm\u00E5ga."),
      spacer(),
      bodyText("Sammanfattningsvis visar b\u00E5de forskning och myndighetsrapporter att AI-genererat inneh\u00E5ll utg\u00F6r en v\u00E4xande utmaning f\u00F6r v\u00E5r f\u00F6rm\u00E5ga att granska information. K\u00E4llkritik \u00E4r inte l\u00E4ngre en \u201Cskoluppgift\u201D utan en demokratisk n\u00F6dv\u00E4ndighet."),
      spacer(),

      heading3("Referenslista"),
      bodyText("Internetstiftelsen. (2025). Svenskarna och internet 2025. H\u00E4mtad 2026-03-10 fr\u00E5n https://svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),
      spacer(),
      bodyText("Myndigheten f\u00F6r psykologiskt f\u00F6rsvar. (2025). AI och informationsp\u00E5verkan. H\u00E4mtad 2026-03-10 fr\u00E5n https://mpf.se/kunskap-och-stod/temabibliotek/ai-och-informationspaverkan"),
      spacer(),
      bodyText("Uppsala universitet. (2026). Generativ AI har sju tydliga roller i arbetet mot desinformation. H\u00E4mtad 2026-03-10 fr\u00E5n https://www.uu.se/press/pressmeddelanden/2026/2026-02-16-generativ-ai-har-sju-tydliga-roller-i-arbetet-mot-desinformation"),
      spacer(),
      spacer(),
      italicText("Observera: I exempeltexten syns tre viktiga saker: (1) h\u00E4nvisningar i l\u00F6pande text, (2) k\u00E4llkritik d\u00E4r f\u00F6rfattaren v\u00E4rderar k\u00E4llan, (3) en korrekt referenslista."),

      // BILAGA 6: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 6: Exit ticket \u2014 Lektion 6"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "V\u00E4lj en av k\u00E4llorna du anv\u00E4nde i din text idag. Skriv:"),
      spacer(),
      bodyText("1. Hur ser h\u00E4nvisningen ut i l\u00F6pande text? (Skriv en exempelmening med h\u00E4nvisning i parentes.)"),
      spacer(),
      bodyText("2. Hur ser k\u00E4llan ut i referenslistan? (Skriv en fullst\u00E4ndig referens enligt Harvard.)"),
      spacer(),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CEnligt [k\u00E4lla] (\u00E5r) ...\u201D"),
      bullet("\u201CI referenslistan: [Efternamn], [Initial]. (\u00C5r). [Titel]...\u201D"),

      // BILAGA 7: BEDÖMNINGSMATRIS
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 7: Bed\u00F6mningsmatris \u2014 Summativ bed\u00F6mning"),
      spacer(),
      bodyText("Den h\u00E4r matrisen visar vad som bed\u00F6ms i din skriftliga analys. Anv\u00E4nd den som st\u00F6d medan du skriver."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 2542, 2542, 2542],
        rows: [
          new TableRow({
            children: [
              headerCell("", 1400),
              headerCell("E", 2542),
              headerCell("C", 2542),
              headerCell("A", 2542),
            ],
          }),
          new TableRow({
            children: [
              cell("M\u00E5l 1: K\u00E4llkritik", 1400),
              cell("Granskar k\u00E4llor med enkla omd\u00F6men. N\u00E4mner k\u00E4llornas trov\u00E4rdighet p\u00E5 en grundl\u00E4ggande niv\u00E5.", 2542),
              cell("Granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men. F\u00F6rklarar k\u00E4llornas styrkor och svagheter med relevanta argument.", 2542),
              cell("Granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men. Visar p\u00E5 komplexa samband och j\u00E4mf\u00F6r k\u00E4llor fr\u00E5n flera perspektiv.", 2542),
            ],
          }),
          new TableRow({
            children: [
              cell("M\u00E5l 4: K\u00E4llh\u00E4nvisning", 1400),
              cell("F\u00F6rs\u00F6ker k\u00E4llh\u00E4nvisa med viss s\u00E4kerhet. H\u00E4nvisningar finns men kan ha brister.", 2542),
              cell("K\u00E4llh\u00E4nvisar korrekt i texten. Referenslistan \u00E4r i stort sett komplett och korrekt.", 2542),
              cell("K\u00E4llh\u00E4nvisar med s\u00E4kerhet i olika sammanhang. Referenslistan \u00E4r fullst\u00E4ndig och konsekvent.", 2542),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Vad bed\u00F6ms?"),
      bullet("Din f\u00F6rm\u00E5ga att granska och v\u00E4rdera k\u00E4llor (inte bara \u00E5terbera\u00E4tta)"),
      bullet("Din f\u00F6rm\u00E5ga att k\u00E4llh\u00E4nvisa korrekt enligt Harvard-systemet"),
      bullet("Ditt resonemang och din argumentation"),
      spacer(),
      heading3("Vad bed\u00F6ms INTE?"),
      bullet("Stavning och grammatik (s\u00E5 l\u00E4nge texten g\u00E5r att f\u00F6rst\u00E5)"),
      bullet("Textens l\u00E4ngd (kvalitet \u00F6ver kvantitet)"),
      bullet("Vilken fr\u00E5ga du v\u00E4ljer (alla tre \u00E4r likv\u00E4rdiga)"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-6.docx", buffer);
  console.log("lektion-6.docx skapad!");
});
