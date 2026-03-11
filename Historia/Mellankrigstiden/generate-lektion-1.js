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
        children: [new TextRun({ text: "Lektion 1: Freden som s\u00E5dde fr\u00F6et", font: "Arial" })],
      }),
      bodyText("F\u00F6rsta v\u00E4rldskrigets konsekvenser och introduktion av akt\u00F6r- och strukturperspektiv"),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Kurs: ", "Historia 1b"),
      boldBodyText("Moment: ", "Mellankrigstiden \u2014 Varf\u00F6r f\u00F6ll demokratin?"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),

      // Lärandemål
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bullet("Redog\u00F6ra f\u00F6r centrala f\u00F6r\u00E4ndringsprocesser och h\u00E4ndelser kring Versaillesfreden och dess konsekvenser (m\u00E5l 1)"),
      bullet("F\u00F6rklara samband med hj\u00E4lp av akt\u00F6r- och strukturperspektiv (m\u00E5l 2)"),
      bullet("Anv\u00E4nda begreppen nationalism, imperialism, akt\u00F6rperspektiv och strukturperspektiv (m\u00E5l 3)"),

      // Förberedelse
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered presentation om Versaillesfreden (se presentation-lektion-1.pptx)"),
      bullet("Skriv ut analysmallen akt\u00F6r/struktur (till E-elever)"),
      bullet("F\u00F6rbered EPA-fr\u00E5gan \u201CVar freden r\u00E4ttvis?\u201D p\u00E5 tavla eller slide"),

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
              cell("F\u00F6rkunskaps-aktivering", 2200),
              cell("\u201CSkiv ner tre saker du vet om f\u00F6rsta v\u00E4rldskriget.\u201D Enskilt 2 min, sedan gemensam sammanst\u00E4llning p\u00E5 tavlan. L\u00E4raren knyter an till momentets fr\u00E5ga: \u201CVarf\u00F6r f\u00F6ll demokratier under mellankrigstiden?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201328 min", 900),
              cell("Instruktion", 1400),
              cell("Presentation: Versailles-freden", 2200),
              cell("Genomg\u00E5ng av Versaillesfreden, krigsskuldsparagrafen, ekonomiska och territoriella konsekvenser. Introduktion av begreppen akt\u00F6rperspektiv och strukturperspektiv med konkreta exempel. EPA-paus vid minut 18: \u201CVar freden r\u00E4ttvis? T\u00E4nk 1 min, diskutera 2 min.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("28\u201360 min", 900),
              cell("Bearbetning", 1400),
              cell("Akt\u00F6r/struktur-analys", 2200),
              cell("Eleverna arbetar i par med fr\u00E5gan: \u201CVarf\u00F6r s\u00E5g freden ut som den gjorde? Anv\u00E4nd akt\u00F6r- och strukturperspektiv.\u201D E-elever f\u00E5r analysmallen med ledande fr\u00E5gor. A-elever arbetar fritt och reflekterar: \u201CKan man separera akt\u00F6r och struktur?\u201D Gemensam genomg\u00E5ng av 2\u20133 pars svar.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("60\u201372 min", 900),
              cell("Summering", 1400),
              cell("Gemensam begrepps-samling", 2200),
              cell("Helklasssamtal: vilka begrepp har vi anv\u00E4nt idag? L\u00E4raren sammanst\u00E4ller p\u00E5 tavlan: Versaillesfreden, krigsskuld, nationalism, imperialism, akt\u00F6rperspektiv, strukturperspektiv.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("72\u201377 min", 900),
              cell("Summering", 1400),
              cell("Utg\u00E5ngsbiljett", 2200),
              cell("\u201CGe ett exempel p\u00E5 en akt\u00F6rf\u00F6rklaring och en strukturf\u00F6rklaring till Tysklands situation efter 1918.\u201D Samlas in.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("77\u201380 min", 900),
              cell("Fram\u00E5t-koppling", 1400),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CN\u00E4sta g\u00E5ng tittar vi p\u00E5 vad som h\u00E4nde n\u00E4r dessa strukturer och akt\u00F6rer kolliderade \u2014 fascismens och nazismens framv\u00E4xt.\u201D", 4526),
            ],
          }),
        ],
      }),

      // Lärarinstruktioner
      heading2("L\u00E4rarinstruktioner"),
      boldBodyText("Uppstart: ", "Samla elevernas svar p\u00E5 tavlan. Kategorisera implicit i \u201Ch\u00E4ndelser\u201D och \u201Corsaker/konsekvenser\u201D. Anv\u00E4nd detta som brygga till presentationen."),
      boldBodyText("Instruktion: ", "Pausa vid minut 18 f\u00F6r EPA. St\u00E4ll fr\u00E5gan \u201CVar det rimligt att skuldbe-l\u00E4gga Tyskland ensamt?\u201D och l\u00E5t eleverna diskutera i par 2 min innan helklass."),
      boldBodyText("Bearbetning: ", "Cirkulera i klassrummet. Till E-elever: \u201CVad skulle en akt\u00F6rf\u00F6rklaring vara h\u00E4r?\u201D Till A-elever: \u201CHur h\u00E4nger akt\u00F6rernas val ihop med de strukturella f\u00F6ruts\u00E4ttningarna?\u201D"),
      boldBodyText("Summering: ", "Lyft goda exempel fr\u00E5n bearbetningen. Visa hur akt\u00F6r och struktur kompletterar varandra. Samla in utg\u00E5ngsbiljetter."),

      // Elevaktiviteter
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Enskild skrivning: tre saker om f\u00F6rsta v\u00E4rldskriget (2 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "EPA: \u201CVar freden r\u00E4ttvis?\u201D (t\u00E4nk 1 min, par 2 min, alla 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Pararbete: akt\u00F6r/struktur-analys av Versaillesfreden (30 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Utg\u00E5ngsbiljett: akt\u00F6r- och strukturf\u00F6rklaring (5 min)", font: "Arial", size: 24 })],
      }),

      // Differentiering
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Analysmall med tv\u00E5 kolumner och ledande fr\u00E5gor: \u201CVilka personer fattade besluten om freden? Vad ville de uppn\u00E5?\u201D / \u201CVilka ekonomiska f\u00F6rh\u00E5llanden p\u00E5verkade? Vilka politiska system fanns?\u201D"),
      boldBodyText("Utmaning (mot A): ", "Arbeta utan mall. Till\u00E4ggsfr\u00E5ga: \u201CKan man verkligen separera akt\u00F6r och struktur i detta fall? Motivera.\u201D"),

      // Material
      heading2("Material"),
      bullet("Presentation om Versaillesfreden (presentation-lektion-1.pptx)"),
      bullet("Analysmall akt\u00F6r/struktur (utskriven)"),
      bullet("Utg\u00E5ngsbiljett-lappar"),

      // Koppling till kunskapskrav
      heading2("Koppling till kunskapskrav"),
      bullet("Eleverna \u00F6var att redog\u00F6ra f\u00F6r Versaillesfredens konsekvenser (m\u00E5l 1: E\u2013A)"),
      bullet("Eleverna introduceras till akt\u00F6r/struktur som analysverktyg och \u00F6var att f\u00F6rklara samband (m\u00E5l 2: E\u2013A)"),
      bullet("Begreppen nationalism, imperialism, akt\u00F6r- och strukturperspektiv introduceras och anv\u00E4nds (m\u00E5l 3: E\u2013A)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("Elevaktiv tid: ", "ca 47 min av 80 min (59%)"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/moment-historia-mellankrigstiden/lektion-1.docx", buffer);
  console.log("lektion-1.docx skapad!");
});
