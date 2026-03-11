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

function bodyText(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24 })],
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

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1A1A2E" },
        paragraph: { spacing: { before: 240, after: 240 } },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 200, after: 120 } },
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
          children: [new TextRun({ text: "Historia 1b \u2014 Mellankrigstiden", font: "Arial", size: 18, color: "888888" })],
        })],
      }),
    },
    children: [
      // Titel
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Lektion 4: Samma m\u00F6nster, olika l\u00E4nder?", font: "Arial" })],
      }),
      bodyText("J\u00E4mf\u00F6rande analys av demokratins fall i Tyskland, Italien och Spanien"),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Kurs: ", "Historia 1b"),
      boldBodyText("Moment: ", "Mellankrigstiden \u2014 Varf\u00F6r f\u00F6ll demokratin?"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),

      // Lärandemål
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bullet("Redog\u00F6ra f\u00F6r demokratins fall i tre l\u00E4nder och identifiera gemensamma drag (m\u00E5l 1)"),
      bullet("F\u00F6rklara komplexa samband genom att j\u00E4mf\u00F6ra akt\u00F6r- och strukturf\u00F6rklaringar mellan l\u00E4nder (m\u00E5l 2)"),
      bullet("Anv\u00E4nda historiska begrepp i j\u00E4mf\u00F6rande analys (m\u00E5l 3)"),

      // Förberedelse
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered kort presentation om spanska inb\u00F6rdeskriget och Francos makt\u00F6vertagande"),
      bullet("Skriv ut j\u00E4mf\u00F6relsematrisen (arbetsblad) \u2014 en version med st\u00F6d (E) och en utan (A)"),
      bullet("F\u00F6rbered landkort med informationspaket: Tyskland, Italien, Spanien"),
      bullet("F\u00F6rbered retrieval practice-fr\u00E5gor fr\u00E5n lektion 1\u20133"),

      // Tidsplanering
      heading2("Tidsplanering"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [900, 1400, 2200, 4526],
        rows: [
          new TableRow({
            children: [
              headerCell("Tid", 900),
              headerCell("Fas", 1400),
              headerCell("Aktivitet", 2200),
              headerCell("Beskrivning", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("0\u20138 min", 900),
              cell("Uppstart", 1400),
              cell("Retrieval practice", 2200),
              cell("Eleverna skriver enskilt (3 min): \u201CNämn tre strukturella faktorer som bidrog till nazismens framv\u00E4xt i Tyskland.\u201D Gemensam genomg\u00E5ng p\u00E5 tavlan. L\u00E4raren fr\u00E5gar: \u201CTror ni att samma faktorer fanns i andra l\u00E4nder?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201320 min", 900),
              cell("Instruktion", 1400),
              cell("Presentation: Spanien som tredje fall", 2200),
              cell("Kort genomg\u00E5ng av spanska inb\u00F6rdeskriget: politisk polarisering, ekonomiska problem, Francos milit\u00E4rkupp 1936, internationellt st\u00F6d fr\u00E5n Hitler och Mussolini. Betona: detta \u00E4r ett tredje exempel p\u00E5 demokratins fall \u2014 vad \u00E4r likt och olikt?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("20\u201355 min", 900),
              cell("Bearbetning", 1400),
              cell("J\u00E4mf\u00F6rande analys med matris", 2200),
              cell("Eleverna arbetar i grupper om 3 (en \u201Cexpert\u201D per land). Steg 1 (10 min): Varje elev l\u00E4ser sitt lands informationspaket och fyller i sin kolumn i matrisen (strukturella orsaker, centrala akt\u00F6rer, v\u00E4g till makt, roll f\u00F6r nationalism/imperialism). Steg 2 (10 min): Gruppen j\u00E4mf\u00F6r de tre l\u00E4nderna. Vad \u00E4r gemensamt? Vad \u00E4r unikt? Steg 3 (15 min): Gruppen besvarar analysfr\u00E5gan: \u201CFinns det ett m\u00F6nster n\u00E4r demokratier faller? Vilken roll spelar akt\u00F6rer respektive strukturer?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("55\u201368 min", 900),
              cell("Summering", 1400),
              cell("L\u00E4rpar med \u00E5tergivning", 2200),
              cell("Eleverna byter grupp och bildar nya par. Varje elev \u00E5terger sin grupps slutsatser: \u201CVilket m\u00F6nster hittade ni?\u201D och \u201CVad var det viktigaste \u2014 akt\u00F6rer eller strukturer?\u201D Den som lyssnar st\u00E4ller en f\u00F6ljdfr\u00E5ga. L\u00E4raren lyssnar p\u00E5 n\u00E5gra par.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("68\u201377 min", 900),
              cell("Summering", 1400),
              cell("Helklassreflektion", 2200),
              cell("L\u00E4raren sammanst\u00E4ller p\u00E5 tavlan: gemensamma strukturella faktorer (ekonomisk kris, demokratisk svaghet, nationalism) och unika akt\u00F6rer (Hitler, Mussolini, Franco). Diskussion: \u201C\u00C4r det strukturerna som skapar akt\u00F6rerna, eller akt\u00F6rerna som utnyttjar strukturerna?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("77\u201380 min", 900),
              cell("Fram\u00E5tkoppling", 1400),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CN\u00E4sta g\u00E5ng samlar vi ihop allt. Ni ska sj\u00E4lva besvara momentets stora fr\u00E5ga: Varf\u00F6r f\u00F6ll demokratin under mellankrigstiden? Det blir en skriftlig analys d\u00E4r ni anv\u00E4nder allt vi arbetat med.\u201D", 4526),
            ],
          }),
        ],
      }),

      // Lärarinstruktioner
      heading2("L\u00E4rarinstruktioner"),
      boldBodyText("Uppstart: ", "Retrieval practice speglar lektion 3. Notera om eleverna kan namnge strukturella faktorer sj\u00E4lvst\u00E4ndigt. Brygga till dagens fr\u00E5ga: \u201CVar det bara i Tyskland?\u201D"),
      boldBodyText("Instruktion: ", "Spanien beh\u00F6ver inte vara lika djupg\u00E5ende som Tyskland/Italien. Fokus \u00E4r att ge tillr\u00E4ckligt underlag f\u00F6r j\u00E4mf\u00F6relsen. Anv\u00E4nd gärna karta f\u00F6r att visa den geografiska spridningen."),
      boldBodyText("Bearbetning: ", "Expertmodellen (varje elev \u201C\u00E4ger\u201D ett land) s\u00E4kerst\u00E4ller att alla \u00E4r aktiva. Cirkulera och st\u00F6d. Till E-elever: \u201CTitta p\u00E5 matrisen \u2014 vad \u00E4r likt i kolumnen f\u00F6r strukturella orsaker?\u201D Till A-elever: \u201C\u00C4r det meningsfullt att tala om ett m\u00F6nster, eller \u00E4r varje fall unikt?\u201D"),
      boldBodyText("Summering: ", "L\u00E4rparen \u00E4r formativ avst\u00E4mning. Helklassreflektionen b\u00F6r visa att b\u00E5da perspektiven beh\u00F6vs. Betona att n\u00E4sta lektion kr\u00E4ver att de kan argumentera med b\u00E5da perspektiven."),

      // Elevaktiviteter
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: strukturella faktorer fr\u00E5n lektion 1\u20133 (8 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Expertl\u00E4sning: informationspaket om ett land (10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Grupparbete: j\u00E4mf\u00F6rande matris och analysfr\u00E5ga (25 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "L\u00E4rpar med \u00E5tergivning (13 min)", font: "Arial", size: 24 })],
      }),

      // Differentiering
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "J\u00E4mf\u00F6relsematrisen har f\u00F6rifyllda kategorier och ledande fr\u00E5gor per cell: \u201CVilken ekonomisk kris drabbade landet?\u201D \u201CVilken ledare tog makten och hur?\u201D Informationspaketen har nyckelord markerade. Analysfr\u00E5gan har meningsstartar."),
      boldBodyText("Utmaning (mot A): ", "Tom matris utan f\u00F6rdefinierade kategorier \u2014 eleven v\u00E4ljer sj\u00E4lv j\u00E4mf\u00F6relsepunkter. Till\u00E4ggsfr\u00E5ga: \u201CKan m\u00F6nstret hj\u00E4lpa oss f\u00F6rst\u00E5 hot mot demokrati idag? Vilka begr\u00E4nsningar har j\u00E4mf\u00F6relsen?\u201D"),

      // Material
      heading2("Material"),
      bullet("Kort presentation om Spanien och Franco"),
      bullet("Informationspaket: ett per land (Tyskland, Italien, Spanien)"),
      bullet("J\u00E4mf\u00F6relsematris (tv\u00E5 versioner: med/utan st\u00F6d)"),
      bullet("Retrieval practice-fr\u00E5gor fr\u00E5n lektion 1\u20133"),

      // Koppling till kunskapskrav
      heading2("Koppling till kunskapskrav"),
      bullet("Eleverna redog\u00F6r f\u00F6r demokratins fall i tre l\u00E4nder och identifierar m\u00F6nster (m\u00E5l 1: E\u2013A)"),
      bullet("J\u00E4mf\u00F6relsen tr\u00E4nar komplexa samband \u2014 s\u00E4rskilt C- och A-niv\u00E5 (m\u00E5l 2)"),
      bullet("Begreppen nationalism, imperialism, totalitarism, demokratisering anv\u00E4nds i j\u00E4mf\u00F6rande kontext (m\u00E5l 3: E\u2013A)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("Elevaktiv tid: ", "ca 56 min av 80 min (70%)"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/moment-historia-mellankrigstiden/lektion-4.docx", buffer);
  console.log("lektion-4.docx skapad!");
});
