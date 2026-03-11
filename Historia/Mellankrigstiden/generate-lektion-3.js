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
        children: [new TextRun({ text: "Lektion 3: Kunde Hitler ha stoppats?", font: "Arial" })],
      }),
      bodyText("Weimarrepublikens fall, propaganda som k\u00E4lla och debatt om akt\u00F6r kontra struktur"),
      new Paragraph({ spacing: { after: 60 }, children: [] }),
      boldBodyText("Kurs: ", "Historia 1b"),
      boldBodyText("Moment: ", "Mellankrigstiden \u2014 Varf\u00F6r f\u00F6ll demokratin?"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),

      // Lärandemål
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bullet("Redog\u00F6ra f\u00F6r Weimarrepublikens fall och Hitlers v\u00E4g till makten (m\u00E5l 1)"),
      bullet("F\u00F6rklara komplexa samband med hj\u00E4lp av akt\u00F6r- och strukturperspektiv (m\u00E5l 2)"),
      bullet("Anv\u00E4nda historiska begrepp i muntlig och skriftlig analys (m\u00E5l 3)"),
      bullet("Granska och tolka propaganda som historisk k\u00E4lla (m\u00E5l 4)"),

      // Förberedelse
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered kort presentation om Weimarrepubliken och v\u00E4gen till 1933"),
      bullet("V\u00E4lj ut 3\u20134 propagandak\u00E4llor (affischer, tal-utdrag) fr\u00E5n Weimarrepubliken/nazismen"),
      bullet("Skriv ut k\u00E4llanalysguide med fr\u00E5gor: upphovsperson, syfte, m\u00E5lgrupp, tillf\u00F6rlitlighet"),
      bullet("F\u00F6rbered debattpositioner: \u201CAkt\u00F6rer avgjorde\u201D vs. \u201CStrukturer avgjorde\u201D"),
      bullet("F\u00F6rbered utg\u00E5ngsbiljett"),

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
              cell("Eleverna skriver enskilt (3 min): \u201CF\u00F6rklara med egna ord varf\u00F6r fascismen v\u00E4xte i Italien eller Tyskland. Anv\u00E4nd minst ett akt\u00F6r- och ett strukturargument.\u201D Kort gemensam genomg\u00E5ng (5 min) d\u00E4r l\u00E4raren lyfter goda exempel och fyller i luckor.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201322 min", 900),
              cell("Instruktion", 1400),
              cell("Presentation: Weimar och v\u00E4gen till 1933", 2200),
              cell("Kort genomg\u00E5ng: Weimarrepublikens sv\u00E5righeter (Versaillesfreden, hyperinflation, politisk instabilitet), b\u00F6rskraschen 1929, massarbetsl\u00F6shet, NSDAP:s tillv\u00E4xt, Hitlers utn\u00E4mning till rikskansler 1933. Betona sp\u00E4nningen: var det oundvikligt eller fanns det alternativ?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("22\u201342 min", 900),
              cell("Bearbetning 1", 1400),
              cell("K\u00E4llanalys: propaganda", 2200),
              cell("Eleverna arbetar i grupper om 3\u20134 med att analysera propagandak\u00E4llor. Varje grupp f\u00E5r 1\u20132 k\u00E4llor och anv\u00E4nder k\u00E4llanalysguiden: Vem skapade k\u00E4llan? I vilket syfte? Vilken bild ger den av samh\u00E4llet? Vad kan k\u00E4llan ber\u00E4tta om varf\u00F6r nazismen fick st\u00F6d? Kort redovisning: varje grupp presenterar sin k\u00E4lla p\u00E5 1 minut.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("42\u201362 min", 900),
              cell("Bearbetning 2", 1400),
              cell("Debatt: akt\u00F6r eller struktur?", 2200),
              cell("Klassen delas i tv\u00E5 halvor med tilldelade positioner. Sida A: \u201CHitler och nazisterna hade kunnat stoppas \u2014 det var akt\u00F6rernas val som avgjorde.\u201D Sida B: \u201CDet var de strukturella krafterna som gjorde nazismens makt\u00F6vertagande oundvikligt.\u201D F\u00F6rberedelse i grupp (5 min), debatt med replikr\u00E4tt (10 min), fri diskussion (5 min). L\u00E4raren modererar och s\u00E4kerst\u00E4ller att begrepp anv\u00E4nds.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("62\u201372 min", 900),
              cell("Summering", 1400),
              cell("Debriefing och syntes", 2200),
              cell("L\u00E4raren sammanfattar debattens huvudargument. Nyckelfr\u00E5ga: \u201CG\u00E5r det att v\u00E4lja en sida, eller beh\u00F6ver man b\u00E5da perspektiven?\u201D Koppla till k\u00E4llanalysen: \u201CVad avsl\u00F6jar propagandan om strukturella problem som nazisterna utnyttjade?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("72\u201377 min", 900),
              cell("Summering", 1400),
              cell("Utg\u00E5ngsbiljett", 2200),
              cell("\u201CKunde Hitler ha stoppats? Motivera med ett akt\u00F6r- och ett strukturargument. Anv\u00E4nd minst tv\u00E5 historiska begrepp.\u201D Samlas in.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("77\u201380 min", 900),
              cell("Fram\u00E5tkoppling", 1400),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CN\u00E4sta g\u00E5ng lyfter vi blicken och j\u00E4mf\u00F6r Tyskland med Italien och Spanien \u2014 fanns det ett m\u00F6nster n\u00E4r demokratier f\u00F6ll?\u201D", 4526),
            ],
          }),
        ],
      }),

      // Lärarinstruktioner
      heading2("L\u00E4rarinstruktioner"),
      boldBodyText("Uppstart: ", "Retrieval practice-uppgiften b\u00F6r visa om eleverna har f\u00F6rst\u00E5tt akt\u00F6r/struktur fr\u00E5n lektion 1\u20132. Notera vilka som beh\u00F6ver extra st\u00F6d."),
      boldBodyText("Instruktion: ", "H\u00E5ll presentationen kort (14 min). Fokusera p\u00E5 att ge eleverna tillr\u00E4cklig kontext f\u00F6r k\u00E4llanalysen och debatten. Anv\u00E4nd tidslinje p\u00E5 tavlan."),
      boldBodyText("K\u00E4llanalys: ", "V\u00E4lj k\u00E4llor med olika karakt\u00E4r (propagandaaffisch, talutdrag, tidningsklipp). St\u00F6d E-elever genom att modellera en k\u00E4llanalys p\u00E5 tavlan f\u00F6rst. Utmana A-elever med: \u201CVad ber\u00E4ttar k\u00E4llan INTE?\u201D"),
      boldBodyText("Debatt: ", "S\u00E4kerst\u00E4ll att b\u00E5da sidor f\u00E5r lika mycket talutrymme. Notera vilka begrepp som anv\u00E4nds. Om debatten stannar, st\u00E4ll provokativa motfr\u00E5gor: \u201CMen om Hitler inte hade funnits, hade n\u00E5gon annan tagit hans plats?\u201D"),
      boldBodyText("Summering: ", "Debriefingen \u00E4r viktig \u2014 f\u00E5nga upp att perspektiven kompletterar varandra snarare \u00E4n utesluter. Samla in utg\u00E5ngsbiljetter och anv\u00E4nd dem f\u00F6r att planera differentiering i lektion 4."),

      // Elevaktiviteter
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: skriftlig \u00E5tergivning fr\u00E5n lektion 1\u20132 (8 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "K\u00E4llanalys i grupp: propagandak\u00E4llor fr\u00E5n mellankrigstiden (20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Debatt med tilldelade positioner: akt\u00F6r eller struktur? (20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Utg\u00E5ngsbiljett: motiverat st\u00E4llningstagande (5 min)", font: "Arial", size: 24 })],
      }),

      // Differentiering
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "K\u00E4llanalysguiden har ledande fr\u00E5gor och exempelsvar. I debatten f\u00E5r E-elever f\u00F6rformulerade argument att utg\u00E5 fr\u00E5n. Utg\u00E5ngsbiljetten har meningsstartar: \u201CEn akt\u00F6rf\u00F6rklaring \u00E4r att...\u201D \u201CEn strukturf\u00F6rklaring \u00E4r att...\u201D"),
      boldBodyText("Utmaning (mot A): ", "K\u00E4llanalys utan ledande fr\u00E5gor, med till\u00E4ggsfr\u00E5ga: \u201CVad \u00E4r k\u00E4llans begr\u00E4nsningar som historiskt bevis?\u201D I debatten f\u00F6rv\u00E4ntas A-elever bem\u00F6ta motståndarens argument med nyanser. Utg\u00E5ngsbiljett utan meningsstartar, med till\u00E4gg: \u201CProblematisera din egen position.\u201D"),

      // Material
      heading2("Material"),
      bullet("Kort presentation om Weimarrepubliken (10\u201312 slides)"),
      bullet("3\u20134 propagandak\u00E4llor (affischer, talutdrag) \u2014 k\u00E4llf\u00F6rteckning bifogad"),
      bullet("K\u00E4llanalysguide (tv\u00E5 versioner: med/utan st\u00F6d)"),
      bullet("Debattpositioner (utskrivna, en per sida)"),
      bullet("Utg\u00E5ngsbiljett-lappar"),

      // Koppling till kunskapskrav
      heading2("Koppling till kunskapskrav"),
      bullet("Eleverna redog\u00F6r f\u00F6r Weimarrepublikens fall och kopplar till b\u00E5de v\u00E4rldskrigen (m\u00E5l 1: E\u2013A)"),
      bullet("Debatten tr\u00E4nar f\u00F6rm\u00E5gan att f\u00F6rklara komplexa samband (m\u00E5l 2: s\u00E4rskilt C\u2013A)"),
      bullet("Historiska begrepp anv\u00E4nds aktivt i k\u00E4llanalys och debatt (m\u00E5l 3: E\u2013A)"),
      bullet("K\u00E4llanalysen tr\u00E4nar granskning, tolkning och v\u00E4rdering av k\u00E4llor (m\u00E5l 4: E\u2013A)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("Elevaktiv tid: ", "ca 53 min av 80 min (66%)"),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
      boldBodyText("OBS: ", "Detta \u00E4r lektionen med h\u00F6gst kognitiv belastning i momentet (k\u00E4llkritik + debatt). \u00D6verv\u00E4g att korta presentationen om eleverna redan har god f\u00F6rf\u00F6rst\u00E5else fr\u00E5n lektion 2."),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/moment-historia-mellankrigstiden/lektion-3.docx", buffer);
  console.log("lektion-3.docx skapad!");
});
