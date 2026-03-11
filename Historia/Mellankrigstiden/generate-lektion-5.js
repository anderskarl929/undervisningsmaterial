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
        children: [new TextRun({ text: "Lektion 5: Varf\u00F6r f\u00F6ll demokratin?", font: "Arial" })],
      }),
      bodyText("Syntes, skriftlig analysuppgift och avslutande reflektion"),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Kurs: ", "Historia 1b"),
      boldBodyText("Moment: ", "Mellankrigstiden \u2014 Varf\u00F6r f\u00F6ll demokratin?"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),

      // Lärandemål
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bullet("Redog\u00F6ra f\u00F6r mellankrigstiden som historisk period i en sammanhängande text (m\u00E5l 1)"),
      bullet("F\u00F6rklara komplexa samband med hj\u00E4lp av akt\u00F6r- och strukturperspektiv i skriftlig analys (m\u00E5l 2)"),
      bullet("Anv\u00E4nda historiska begrepp med s\u00E4kerhet i en l\u00E4ngre text (m\u00E5l 3)"),
      bullet("Reflektera \u00F6ver k\u00E4llors anv\u00E4ndbarhet f\u00F6r att f\u00F6rst\u00E5 mellankrigstiden (m\u00E5l 4)"),

      // Förberedelse
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered analysuppgiften med tydliga instruktioner och bed\u00F6mningsmatris"),
      bullet("Skriv ut analysmall/st\u00F6dstruktur f\u00F6r E-elever"),
      bullet("F\u00F6rbered snabbskrivningsfr\u00E5ga och reflektionsfr\u00E5gor"),
      bullet("Ha elevernas utg\u00E5ngsbiljetter fr\u00E5n lektion 3 tillg\u00E4ngliga som \u00E5terkoppling"),

      // Analysuppgiften
      heading2("Analysuppgift"),
      bodyText("Momentets centrala fr\u00E5ga:"),
      new Paragraph({
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: "\u201CVarf\u00F6r f\u00F6ll demokratier under mellankrigstiden? Berodde det p\u00E5 enskilda ledare (akt\u00F6rer) eller p\u00E5 djupare samh\u00E4llskrafter (strukturer)? Argumentera med hj\u00E4lp av exempel fr\u00E5n minst tv\u00E5 l\u00E4nder.\u201D", font: "Arial", size: 24, italics: true })],
      }),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Omf\u00E5ng: ", "ca 400\u2013600 ord"),
      boldBodyText("Krav: ", "Anv\u00E4nd akt\u00F6r- och strukturperspektiv. Anv\u00E4nd minst fyra av momentets begrepp (nationalism, imperialism, totalitarism, demokratisering, akt\u00F6rperspektiv, strukturperspektiv). H\u00E4nvisa till konkreta historiska exempel."),

      // Bedömningsmatris
      heading2("Bed\u00F6mningsmatris"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 2500, 2500, 2526],
        rows: [
          new TableRow({
            children: [
              headerCell("Kriterium", 1500),
              headerCell("E", 2500),
              headerCell("C", 2500),
              headerCell("A", 2526),
            ],
          }),
          new TableRow({
            children: [
              cell("Redog\u00F6relse", 1500),
              cell("Redog\u00F6r \u00F6versiktligt f\u00F6r h\u00E4ndelser och processer", 2500),
              cell("Redog\u00F6r utf\u00F6rligt f\u00F6r h\u00E4ndelser och processer", 2500),
              cell("Redog\u00F6r utf\u00F6rligt och nyanserat f\u00F6r h\u00E4ndelser och processer", 2526),
            ],
          }),
          new TableRow({
            children: [
              cell("Samband", 1500),
              cell("F\u00F6rklarar enkla samband med akt\u00F6r- och strukturperspektiv", 2500),
              cell("F\u00F6rklarar samband med akt\u00F6r- och strukturperspektiv", 2500),
              cell("F\u00F6rklarar komplexa samband och problematiserar perspektiven", 2526),
            ],
          }),
          new TableRow({
            children: [
              cell("Begrepp", 1500),
              cell("Anv\u00E4nder begrepp med viss s\u00E4kerhet", 2500),
              cell("Anv\u00E4nder begrepp med viss s\u00E4kerhet i relevant kontext", 2500),
              cell("Anv\u00E4nder begrepp med s\u00E4kerhet och precision", 2526),
            ],
          }),
          new TableRow({
            children: [
              cell("J\u00E4mf\u00F6relse", 1500),
              cell("N\u00E4mner tv\u00E5 l\u00E4nder med \u00F6versiktlig j\u00E4mf\u00F6relse", 2500),
              cell("J\u00E4mf\u00F6r tv\u00E5 l\u00E4nder med tydliga likheter och skillnader", 2500),
              cell("J\u00E4mf\u00F6r l\u00E4nder p\u00E5 djupet och drar slutsatser om m\u00F6nster", 2526),
            ],
          }),
        ],
      }),

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
              cell("0\u201310 min", 900),
              cell("Uppstart", 1400),
              cell("Retrieval practice + \u00E5terkoppling", 2200),
              cell("Snabbskrivning (3 min): \u201CSkriv ner allt du minns om varf\u00F6r demokratier f\u00F6ll under mellankrigstiden. Anv\u00E4nd s\u00E5 m\u00E5nga begrepp du kan.\u201D L\u00E4raren \u00E5terkopplar p\u00E5 utg\u00E5ngsbiljetter fr\u00E5n lektion 3: lyfter goda exempel och vanliga missuppfattningar. Kort repetition av akt\u00F6r/struktur som analysverktyg.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("10\u201318 min", 900),
              cell("Instruktion", 1400),
              cell("Presentation av uppgiften", 2200),
              cell("L\u00E4raren presenterar analysuppgiften och bed\u00F6mningsmatrisen. G\u00E5r igenom vad som f\u00F6rv\u00E4ntas p\u00E5 E-, C- och A-niv\u00E5 med konkreta exempel. Visar modelltext (2\u20133 meningar) f\u00F6r hur man kan inleda. Eleverna f\u00E5r st\u00E4lla fr\u00E5gor.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("18\u201358 min", 900),
              cell("Bearbetning", 1400),
              cell("Skriftlig analys", 2200),
              cell("Eleverna skriver sin analys enskilt. L\u00E4raren cirkulerar och ger formativ \u00E5terkoppling under skrivandet. E-elever f\u00E5r analysmall med struktur: inledning (presentera fr\u00E5gan), akt\u00F6rargument (med exempel), strukturargument (med exempel), j\u00E4mf\u00F6relse, slutsats. A-elever skriver fritt med krav p\u00E5 att problematisera uppdelningen akt\u00F6r/struktur.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("58\u201368 min", 900),
              cell("Summering", 1400),
              cell("Kamratrespons", 2200),
              cell("Eleverna byter text med en granne och ger respons utifr\u00E5n tv\u00E5 fr\u00E5gor: \u201CAnv\u00E4nds b\u00E5de akt\u00F6r- och strukturperspektiv?\u201D och \u201CVilket begrepp saknas eller kunde anv\u00E4ndas b\u00E4ttre?\u201D Eleverna f\u00E5r 5 minuter att revidera sin text utifr\u00E5n responsen.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("68\u201377 min", 900),
              cell("Summering", 1400),
              cell("Avslutande reflektion", 2200),
              cell("Helklasssamtal: \u201CVad har vi l\u00E4rt oss om varf\u00F6r demokratier faller?\u201D L\u00E4raren lyfter koppling till nutid: \u201CKan liknande m\u00F6nster uppst\u00E5 idag?\u201D Avslutande snabbskrivning (2 min): \u201CVad tar du med dig fr\u00E5n det h\u00E4r momentet?\u201D Samlas in.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("77\u201380 min", 900),
              cell("Avslut", 1400),
              cell("Insamling och avslut", 2200),
              cell("Eleverna l\u00E4mnar in sina analystexter. L\u00E4raren tackar och ger en kort f\u00F6rhandsvisning av n\u00E4sta moment.", 4526),
            ],
          }),
        ],
      }),

      // Lärarinstruktioner
      heading2("L\u00E4rarinstruktioner"),
      boldBodyText("Uppstart: ", "\u00C5terkopplingen p\u00E5 utg\u00E5ngsbiljetter fr\u00E5n lektion 3 \u00E4r viktig \u2014 den visar eleverna att du l\u00E4st och v\u00E4rderar deras arbete. Lyft 2\u20133 goda exempel anonymt."),
      boldBodyText("Instruktion: ", "Var tydlig med vad som skiljer E, C och A. Visa konkret: \u201CEtt enkelt samband \u00E4r: \u2019Hitler tog makten.\u2019 Ett komplext samband \u00E4r: \u2019Den ekonomiska krisen skapade misn\u00F6je som Hitler utnyttjade genom sin retorik.\u2019\u201D"),
      boldBodyText("Bearbetning: ", "Cirkulera aktivt. Prim\u00E4rt fokus: s\u00E4kerst\u00E4ll att alla elever kommer ig\u00E5ng. Till elever som fastnar: \u201CB\u00F6rja med att skriva ett akt\u00F6rargument \u2014 vad gjorde Hitler/Mussolini konkret?\u201D Till elever som \u00E4r klara tidigt: \u201CLägg till en problematisering \u2014 h\u00E5ller uppdelningen akt\u00F6r/struktur egentligen?\u201D"),
      boldBodyText("Kamratrespons: ", "H\u00E5ll responsen strikt till de tv\u00E5 fr\u00E5gorna. Modellera g\u00E4rna med ett anonymt exempel f\u00F6rst."),
      boldBodyText("Avslut: ", "Kopplingen till nutid ska vara kort och \u00F6ppnande, inte moralistisk. Snabbskrivningen \u00E4r b\u00E5de formativ avst\u00E4mning och metakognitiv reflektion."),

      // Elevaktiviteter
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Snabbskrivning: allt du minns om mellankrigstiden (3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Skriftlig analysuppgift: \u201CVarf\u00F6r f\u00F6ll demokratin?\u201D (40 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Kamratrespons p\u00E5 analystext (10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Avslutande snabbskrivning: reflektion \u00F6ver momentet (2 min)", font: "Arial", size: 24 })],
      }),

      // Differentiering
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Analysmall med tydlig struktur: (1) Inledning \u2014 presentera fr\u00E5gan i en mening, (2) Akt\u00F6rargument \u2014 ge ett exempel, (3) Strukturargument \u2014 ge ett exempel, (4) J\u00E4mf\u00F6relse \u2014 hur liknar/skiljer sig tv\u00E5 l\u00E4nder?, (5) Slutsats. Meningsstartar till varje del. Begreppslista tillg\u00E4nglig."),
      boldBodyText("Utmaning (mot A): ", "Ingen mall. Krav p\u00E5 att problematisera akt\u00F6r/struktur-uppdelningen: \u201CArgumentera f\u00F6r att uppdelningen \u00E4r f\u00F6renklad \u2014 hur h\u00E4nger perspektiven ihop?\u201D Till\u00E4gg: \u201CDra en parallell till nutida hot mot demokrati.\u201D"),

      // Material
      heading2("Material"),
      bullet("Analysuppgift med bed\u00F6mningsmatris (utskriven)"),
      bullet("Analysmall/st\u00F6dstruktur f\u00F6r E-elever"),
      bullet("Begreppslista fr\u00E5n momentet"),
      bullet("Utg\u00E5ngsbiljetter fr\u00E5n lektion 3 (\u00E5terkoppling)"),
      bullet("Snabbskrivnings-lappar"),

      // Koppling till kunskapskrav
      heading2("Koppling till kunskapskrav"),
      bullet("Analysen pr\u00F6var samtliga l\u00E4randem\u00E5l i en sammanhängande uppgift (m\u00E5l 1\u20134)"),
      bullet("Skriftlig analys med akt\u00F6r/struktur tr\u00E4nar och synligg\u00F6r sambandst\u00E4nkande (m\u00E5l 2: E\u2013A)"),
      bullet("Kamratrespons fokuserar p\u00E5 begreppsanv\u00E4ndning (m\u00E5l 3)"),
      bullet("Reflektionen \u00F6ver k\u00E4llor fr\u00E5n lektion 3 kan integreras i texten (m\u00E5l 4)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("Elevaktiv tid: ", "ca 55 min av 80 min (69%)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("OBS: ", "Analysuppgiften kan \u00E4ven anv\u00E4ndas som summativ bed\u00F6mning om s\u00E5 \u00F6nskas. Justera i s\u00E5 fall tidsf\u00F6rdelningen: mer skrivtid, mindre kamratrespons."),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/moment-historia-mellankrigstiden/lektion-5.docx", buffer);
  console.log("lektion-5.docx skapad!");
});
