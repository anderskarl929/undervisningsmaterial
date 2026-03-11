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
        children: [new TextRun({ text: "Lektion 2: Demokratins fiender", font: "Arial" })],
      }),
      bodyText("Fascismens och nazismens framv\u00E4xt \u2014 akt\u00F6rer, strukturer och orsak-konsekvenskedjor"),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Kurs: ", "Historia 1b"),
      boldBodyText("Moment: ", "Mellankrigstiden \u2014 Varf\u00F6r f\u00F6ll demokratin?"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),

      // Lärandemål
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bullet("Redog\u00F6ra f\u00F6r fascismens och nazismens framv\u00E4xt i Italien och Tyskland (m\u00E5l 1)"),
      bullet("F\u00F6rklara orsak-konsekvenssamband med hj\u00E4lp av akt\u00F6r- och strukturperspektiv (m\u00E5l 2)"),
      bullet("Anv\u00E4nda begreppen nationalism, imperialism, totalitarism, demokratisering (m\u00E5l 3)"),

      // Förberedelse
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered presentation om fascismens och nazismens framv\u00E4xt"),
      bullet("Skriv ut orsak-konsekvenskedjan (arbetsblad) \u2014 en version med st\u00F6dfr\u00E5gor (E) och en utan (A)"),
      bullet("F\u00F6rbered retrieval practice-fr\u00E5gorna fr\u00E5n lektion 1 p\u00E5 tavla/slide"),
      bullet("F\u00F6rbered faktakort om Mussolini och Hitler (till pararbetet)"),

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
              cell("Tre snabbfr\u00E5gor fr\u00E5n lektion 1: \u201CVad innebar krigsskuldsparagrafen?\u201D \u201CGe ett exempel p\u00E5 en strukturf\u00F6rklaring till freden.\u201D \u201CVad \u00E4r skillnaden mellan akt\u00F6r- och strukturperspektiv?\u201D Enskilt 2 min, gemensam genomg\u00E5ng 4 min.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201325 min", 900),
              cell("Instruktion", 1400),
              cell("Presentation: Fascism och nazism", 2200),
              cell("Genomg\u00E5ng av den politiska situationen i Italien och Tyskland efter f\u00F6rsta v\u00E4rldskriget. Mussolinis marsch mot Rom (1922) och Hitlers tidiga politiska karri\u00E4r. Begreppen totalitarism, nationalism och imperialism som drivkrafter. EPA-paus vid minut 18: \u201CVarf\u00F6r lockade dessa r\u00F6relser s\u00E5 m\u00E5nga m\u00E4nniskor?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("25\u201358 min", 900),
              cell("Bearbetning", 1400),
              cell("Orsak-konsekvens-kedja", 2200),
              cell("Eleverna arbetar i par med att bygga en orsak-konsekvenskedja: \u201CVarf\u00F6r fick fascismen/nazismen folkligt st\u00F6d?\u201D Steg 1 (10 min): Sortera orsaker i \u201Cakt\u00F6rf\u00F6rklaringar\u201D och \u201Cstrukturf\u00F6rklaringar\u201D med hj\u00E4lp av faktakorten. Steg 2 (10 min): Rita en orsak-konsekvenskedja som visar hur strukturer skapade utrymme f\u00F6r akt\u00F6rer. Steg 3 (13 min): Jämf\u00F6r Italien och Tyskland \u2014 vad \u00E4r likt, vad \u00E4r olikt? Gemensam genomg\u00E5ng av 2\u20133 pars kedjor p\u00E5 tavlan.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("58\u201370 min", 900),
              cell("Summering", 1400),
              cell("L\u00E4rpar med \u00E5tergivning", 2200),
              cell("Eleverna byter par. Ny partner \u00E5terger vad de l\u00E4rt sig: \u201CF\u00F6rklara f\u00F6r din nya partner varf\u00F6r fascismen v\u00E4xte i Italien ELLER varf\u00F6r nazismen v\u00E4xte i Tyskland. Anv\u00E4nd b\u00E5de akt\u00F6r- och strukturperspektiv.\u201D Den som lyssnar st\u00E4ller en f\u00F6ljdfr\u00E5ga.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("70\u201377 min", 900),
              cell("Summering", 1400),
              cell("Gemensam reflektion", 2200),
              cell("Helklasssamtal: \u201CKunde fascismen ha stoppats om andra akt\u00F6rer hade agerat annorlunda, eller var den ett resultat av strukturella krafter?\u201D L\u00E4raren sammanfattar och kopplar till n\u00E4sta lektion.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("77\u201380 min", 900),
              cell("Fram\u00E5tkoppling", 1400),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CN\u00E4sta g\u00E5ng zoomar vi in p\u00E5 Tyskland och st\u00E4ller fr\u00E5gan: Kunde Hitler ha stoppats? Vi ska arbeta med propaganda som k\u00E4llmaterial.\u201D", 4526),
            ],
          }),
        ],
      }),

      // Lärarinstruktioner
      heading2("L\u00E4rarinstruktioner"),
      boldBodyText("Uppstart: ", "Anv\u00E4nd retrieval practice-fr\u00E5gorna f\u00F6r att aktivera f\u00F6rkunskaper fr\u00E5n lektion 1. L\u00E5t eleverna skriva enskilt f\u00F6rst, g\u00E5 sedan igenom svaren muntligt. Fyll i luckor och r\u00E4tta missuppfattningar."),
      boldBodyText("Instruktion: ", "H\u00E5ll presentationen levande med konkreta exempel: Mussolinis retorik, Hitlers \u00F6lhallskupp. Pausa f\u00F6r EPA vid minut 18. Betona att det \u00E4r en genuin fr\u00E5ga \u2014 det finns inte ett \u201Cr\u00E4tt svar\u201D."),
      boldBodyText("Bearbetning: ", "Cirkulera och st\u00F6d paren. Till E-elever: \u201CTitta p\u00E5 faktakortet \u2014 vilka konkreta handlingar genomf\u00F6rde Mussolini/Hitler?\u201D Till A-elever: \u201CKan en ekonomisk kris i sig leda till fascism, eller kr\u00E4vs det specifika akt\u00F6rer?\u201D Se till att j\u00E4mf\u00F6relsen Italien\u2013Tyskland inte blir ytlig \u2014 pressa eleverna p\u00E5 specifika likheter/skillnader."),
      boldBodyText("Summering: ", "L\u00E4rparen \u00E4r den formativa avst\u00E4mningen. Lyssna p\u00E5 n\u00E5gra par och notera vilka begrepp som anv\u00E4nds. Lyft goda formuleringar i helklassamtalet."),

      // Elevaktiviteter
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: tre fr\u00E5gor fr\u00E5n lektion 1 (8 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "EPA: \u201CVarf\u00F6r lockade dessa r\u00F6relser s\u00E5 m\u00E5nga m\u00E4nniskor?\u201D (5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Pararbete: orsak-konsekvenskedja med sortering och j\u00E4mf\u00F6relse (33 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "L\u00E4rpar med \u00E5tergivning (12 min)", font: "Arial", size: 24 })],
      }),

      // Differentiering
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Orsak-konsekvenskedjan har f\u00F6rifyllda rutor med ledande fr\u00E5gor: \u201CVilken ekonomisk situation r\u00E5dde?\u201D \u201CVad lovade ledaren?\u201D \u201CVilka grupper st\u00F6dde r\u00F6relsen och varf\u00F6r?\u201D Faktakorten ger konkret information att arbeta utifr\u00E5n."),
      boldBodyText("Utmaning (mot A): ", "Kedjan saknar f\u00F6rifyllda rutor. Till\u00E4ggsuppgift: \u201CArgumentera f\u00F6r att fascismens framv\u00E4xt INTE kan f\u00F6rklaras med enbart akt\u00F6rperspektiv. Vad missar man d\u00E5?\u201D"),

      // Material
      heading2("Material"),
      bullet("Presentation om fascismens och nazismens framv\u00E4xt"),
      bullet("Arbetsblad: orsak-konsekvenskedja (tv\u00E5 versioner: med/utan st\u00F6d)"),
      bullet("Faktakort om Mussolini och Hitler"),
      bullet("Retrieval practice-fr\u00E5gor fr\u00E5n lektion 1"),

      // Koppling till kunskapskrav
      heading2("Koppling till kunskapskrav"),
      bullet("Eleverna redog\u00F6r f\u00F6r fascismens och nazismens framv\u00E4xt och kopplar till nationalism, imperialism och ekonomisk kris (m\u00E5l 1: E\u2013A)"),
      bullet("Eleverna bygger orsak-konsekvenskedjor med akt\u00F6r- och strukturperspektiv (m\u00E5l 2: E\u2013A)"),
      bullet("Begreppen totalitarism, nationalism, imperialism, demokratisering anv\u00E4nds aktivt (m\u00E5l 3: E\u2013A)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("Elevaktiv tid: ", "ca 58 min av 80 min (73%)"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/moment-historia-mellankrigstiden/lektion-2.docx", buffer);
  console.log("lektion-2.docx skapad!");
});
