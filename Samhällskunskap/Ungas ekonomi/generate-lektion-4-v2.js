const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require('docx');

// --- Shared config (same as generate-docx.js) ---
const FONT = "Arial";
const PAGE_WIDTH = 11906; // A4
const PAGE_HEIGHT = 16838;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN;

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const headerBg = { fill: "1B4F72", type: ShadingType.CLEAR };
const altRowBg = { fill: "F2F7FA", type: ShadingType.CLEAR };
const scenarioBg = { fill: "FFF8E1", type: ShadingType.CLEAR };
const zoomBg = { fill: "E8F5E9", type: ShadingType.CLEAR };

function styles() {
  return {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: "1B4F72" },
        paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "2C3E50" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
    ]
  };
}

function bulletConfig() {
  return {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "letters",
        levels: [{
          level: 0, format: LevelFormat.UPPER_LETTER, text: "(%1)",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  };
}

function footerSection() {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: "Samh\u00e4llskunskap 1a1 \u2014 Ungas ekonomi", font: FONT, size: 18, color: "999999" }),
        new TextRun({ text: "\tSida ", font: FONT, size: 18, color: "999999" }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "999999" }),
      ],
      tabStops: [{ type: "right", position: CONTENT_WIDTH }],
    })]
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: FONT, size: 24, ...opts })]
  });
}

function richPara(runs, spacing = { after: 120 }) {
  return new Paragraph({ spacing, children: runs });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: FONT, size: 24 })]
  });
}

function boldBullet(label, text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [
      new TextRun({ text: label, font: FONT, size: 24, bold: true }),
      new TextRun({ text: ` ${text}`, font: FONT, size: 24 }),
    ]
  });
}

function italicBullet(label, text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [
      new TextRun({ text: label, font: FONT, size: 24, italics: true }),
      new TextRun({ text: ` ${text}`, font: FONT, size: 24 }),
    ]
  });
}

function timeTable(rows) {
  const cw = [900, 1200, 1800, 5126];
  const headerTexts = ["Tid", "Fas", "Aktivitet", "Beskrivning"];

  const headerRow = new TableRow({
    children: headerTexts.map((t, i) => new TableCell({
      borders, shading: headerBg, width: { size: cw[i], type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 22, bold: true, color: "FFFFFF" })] })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders, width: { size: cw[ci], type: WidthType.DXA }, margins: cellMargins,
      shading: ri % 2 === 1 ? altRowBg : undefined,
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: FONT, size: 22 })] })]
    }))
  }));

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: cw,
    rows: [headerRow, ...dataRows]
  });
}

// Scenario box - a single-cell table with colored background
function scenarioBox(title, children) {
  const thickBorder = { style: BorderStyle.SINGLE, size: 2, color: "E0A800" };
  const boxBorders = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        children: [new TableCell({
          borders: boxBorders,
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          shading: scenarioBg,
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            new Paragraph({
              spacing: { after: 120 },
              children: [new TextRun({ text: title, font: FONT, size: 28, bold: true, color: "B8860B" })]
            }),
            ...children
          ]
        })]
      })
    ]
  });
}

// Zoom-out box
function zoomBox(text) {
  const greenBorder = { style: BorderStyle.SINGLE, size: 2, color: "4CAF50" };
  const boxBorders = { top: greenBorder, bottom: greenBorder, left: greenBorder, right: greenBorder };

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        children: [new TableCell({
          borders: boxBorders,
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          shading: zoomBg,
          margins: { top: 100, bottom: 100, left: 200, right: 200 },
          children: [
            new Paragraph({
              spacing: { after: 60 },
              children: [
                new TextRun({ text: "ZOOM UT: ", font: FONT, size: 22, bold: true, color: "2E7D32" }),
                new TextRun({ text, font: FONT, size: 22, italics: true, color: "2E7D32" }),
              ]
            })
          ]
        })]
      })
    ]
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 200 }, children: [] });
}

// ============================================================
// LEKTION 4 v2 - Kedjeformat
// ============================================================
function lektion4() {
  return new Document({
    numbering: bulletConfig(),
    styles: styles(),
    sections: [{
      properties: {
        page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } }
      },
      footers: { default: footerSection() },
      children: [
        h1('Lektion 4: "Vad h\u00e4nder om...?" \u2014 Risker, val och konsekvenser'),
        para("Kurs: Samh\u00e4llskunskap 1a1 | Moment: Ungas ekonomi | L\u00e4ngd: 60 minuter", { italics: true, color: "666666" }),

        h2("L\u00e4randem\u00e5l f\u00f6r lektionen"),
        bullet("M\u00e5l 3: Resonera om privatekonomiska val och deras konsekvenser"),
        bullet("M\u00e5l 4: Diskutera sambandet individ\u2013samh\u00e4lle \u2014 hur individens ekonomiska beslut p\u00e5verkar samh\u00e4llsekonomin"),

        h2("F\u00f6rberedelse"),
        bullet("F\u00f6rbered scenariokort i kedjeformat (utskrift, ett set per par)"),
        bullet("Kort presentation om l\u00e5neformer och r\u00e4nta"),
        bullet("Miniwhiteboard eller papper f\u00f6r utr\u00e4kningar"),
        bullet('Valfritt: skriv ut exempel p\u00e5 snabbl\u00e5ne-reklam och Klarna-reklam f\u00f6r diskussion'),
        bullet('F\u00f6rbered "Zoom ut"-fr\u00e5gorna (en per scenario) \u2014 kan visas p\u00e5 tavlan eller skrivas ut'),

        h2("Tidsplanering"),
        timeTable([
          ["0\u20135 min", "Uppstart", "Retrieval practice", '"Hur mycket hade Alex kvar varje m\u00e5nad i er budget? Vad var den sv\u00e5raste prioriteringen?"'],
          ["5\u201315 min", "Instruktion", "Risker och skuldf\u00e4llan", "Ov\u00e4ntade utgifter, l\u00e5neformer, r\u00e4nta-p\u00e5-r\u00e4nta. Konkret: Alex l\u00e5nar 10 000 kr till 35% r\u00e4nta \u2014 vad kostar l\u00e5net totalt?"],
          ["15\u201348 min", "Bearbetning", "Scenariokedja", "Scenarierna bygger p\u00e5 varandra \u2014 val i scenario 1 p\u00e5verkar scenario 2, osv. Paren hinner minst 3 av 4."],
          ["48\u201356 min", "Summering", "Helklassdiskussion", "J\u00e4mf\u00f6r kedjorna: vilka par tog snabbl\u00e5net \u2014 och hur p\u00e5verkade det resten?"],
          ["56\u201360 min", "Koppling fram\u00e5t", "N\u00e4sta g\u00e5ng", '"N\u00e4sta lektion kopplar vi ihop allt: kretsloppet, arbetsmarknaden, privatekonomin."'],
        ]),

        new Paragraph({ children: [new PageBreak()] }),

        // --- SCENARIOKORT ---
        h2("Scenariokort \u2014 Kedjeformat"),
        richPara([
          new TextRun({ text: "Instruktion till eleverna: ", font: FONT, size: 24, bold: true }),
          new TextRun({ text: "Scenarierna h\u00e4nger ihop. Arbeta i ordning \u2014 valet ni g\u00f6r i scenario 1 p\u00e5verkar Alex situation i scenario 2, och s\u00e5 vidare. Anteckna vilka val ni g\u00f6r.", font: FONT, size: 24 }),
        ]),

        spacer(),

        // --- SCENARIO 1 ---
        scenarioBox("Scenario 1: Tandl\u00e4karen + h\u00f6jd hyra (samtidigt)", [
          new Paragraph({ spacing: { after: 120 }, children: [
            new TextRun({ text: "Alex har ont i en tand. Akuttandv\u00e5rd kostar 4 200 kr. Samma vecka meddelar hyresv\u00e4rden att hyran h\u00f6js med 600 kr/m\u00e5n fr\u00e5n n\u00e4sta m\u00e5nad. Alex har 1 800 kr p\u00e5 kontot och ett buffertsparande p\u00e5 5 000 kr. N\u00e4sta l\u00f6n kommer om 18 dagar.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Alternativen:", font: FONT, size: 24, bold: true })] }),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "(A) Ta ett snabbl\u00e5n p\u00e5 5 000 kr (39% r\u00e4nta, 3 m\u00e5n avbetalning) \u2014 fixa tanden direkt, beh\u00e5ll bufferten intakt.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "(B) Anv\u00e4nd bufferten \u2014 fixa tanden, men st\u00e5 utan s\u00e4kerhetsn\u00e4t. Ingen buffert kvar.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "(C) V\u00e4nta till l\u00f6nen, st\u00e5 ut med sm\u00e4rtan \u2014 riskera att det blir v\u00e4rre (rotbehandling: 8 500 kr).", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "(D) Fr\u00e5ga f\u00f6r\u00e4ldrarna. Alex f\u00f6r\u00e4ldrar har sj\u00e4lva tight ekonomi \u2014 det skulle skapa skuld (inte ekonomisk, utan moralisk).", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Diskussionsfr\u00e5ga: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "Alla alternativ har ett pris. Vad kostar snabbl\u00e5net i kronor? Vad kostar det att v\u00e4nta i risk? Vad kostar det att fr\u00e5ga f\u00f6r\u00e4ldrarna i relation? Vilken typ av kostnad v\u00e4ger tyngst?", font: FONT, size: 24, italics: true })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Anteckna ert val \u2014 det p\u00e5verkar scenario 2.", font: FONT, size: 24, bold: true, color: "B8860B" })
          ]}),
        ]),

        spacer(),
        zoomBox("Om 30% av unga saknar buffertsparande \u2014 vad inneb\u00e4r det f\u00f6r snabbl\u00e5nebranschen? Vem tj\u00e4nar pengar p\u00e5 att unga inte har marginaler?"),

        new Paragraph({ children: [new PageBreak()] }),

        // --- SCENARIO 2 ---
        scenarioBox("Scenario 2: Kompisen som beh\u00f6ver hj\u00e4lp (konsekvenserna f\u00f6ljer med)", [
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Tv\u00e5 veckor senare. Resultatet av ert val i scenario 1:", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "Om ni valde A (snabbl\u00e5n): ", font: FONT, size: 24, italics: true }),
            new TextRun({ text: "Alex betalar 1 780 kr/m\u00e5n i 3 m\u00e5nader p\u00e5 l\u00e5net. Bufferten \u00e4r kvar, men budgeten \u00e4r pressad.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "Om ni valde B (bufferten): ", font: FONT, size: 24, italics: true }),
            new TextRun({ text: "Tanden \u00e4r fixad, men Alex har 0 kr i buffert. Ekonomin \u00e4r s\u00e5rbar.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "Om ni valde C (v\u00e4nta): ", font: FONT, size: 24, italics: true }),
            new TextRun({ text: "Tanden blev v\u00e4rre. Nu beh\u00f6ver Alex en rotbehandling f\u00f6r 8 500 kr \u2014 och har fortfarande 1 800 kr.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "Om ni valde D (f\u00f6r\u00e4ldrarna): ", font: FONT, size: 24, italics: true }),
            new TextRun({ text: 'Tanden \u00e4r fixad, bufferten kvar, men Alex mamma ringde ig\u00e5r och n\u00e4mnde att "det b\u00f6rjar bli sv\u00e5rt hemma ocks\u00e5".', font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Ny situation: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "Alex kompis Dani har blivit av med sitt jobb och kan inte betala hyran \u2014 5 000 kr. Dani fr\u00e5gar Alex om hj\u00e4lp. Samtidigt har Alex arbetsgivare meddelat att arbetstiden sk\u00e4rs ned 25% n\u00e4sta m\u00e5nad p\u00e5 grund av vikande f\u00f6rs\u00e4ljning. Alex inkomst sjunker fr\u00e5n 17 600 kr till ca 13 200 kr netto.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Diskussionsfr\u00e5ga: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: 'Hur f\u00f6r\u00e4ndrar Alex nuvarande situation (fr\u00e5n scenario 1) valm\u00f6jligheterna? \u00c4r det l\u00e4ttare att vara "en bra v\u00e4n" n\u00e4r man sj\u00e4lv har marginaler? Vad h\u00e4nder med v\u00e4nskapen om Alex s\u00e4ger nej?', font: FONT, size: 24, italics: true })
          ]}),
        ]),

        spacer(),
        zoomBox("Under en l\u00e5gkonjunktur sk\u00e4rs arbetstider ned f\u00f6r hundratusentals unga samtidigt. Vad h\u00e4nder med konsumtionen i kretsloppet? Kan det f\u00f6rv\u00e4rra l\u00e5gkonjunkturen \u2014 och i s\u00e5 fall hur?"),

        new Paragraph({ children: [new PageBreak()] }),

        // --- SCENARIO 3 ---
        scenarioBox("Scenario 3: Skuldspiralens logik (flera tryck samtidigt)", [
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "En m\u00e5nad senare. Alex har nu:", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "Eventuellt ett snabbl\u00e5n (fr\u00e5n scenario 1)", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "Eventuellt l\u00e5nat ut pengar till Dani (som inte betalat tillbaka)", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "25% mindre i inkomst", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "600 kr h\u00f6gre hyra", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "Sammanlagt: Alex budget g\u00e5r inte ihop. Det saknas ca 3 000\u20135 000 kr varje m\u00e5nad.", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Tre saker h\u00e4nder samma vecka:", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: '1. Alex jobbarkompisar planerar en weekend till K\u00f6penhamn (4 000 kr). "Kom igen, du beh\u00f6ver det, l\u00e4gg det p\u00e5 Klarna."', font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "2. Alex mobil g\u00e5r s\u00f6nder. En ny kostar 4 500 kr, en begagnad 1 800 kr. Mobilen beh\u00f6vs f\u00f6r jobb (schemat skickas via app).", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "3. Ett brev fr\u00e5n Kronofogden: Alex har en obetald r\u00e4kning p\u00e5 2 200 kr (gl\u00f6md tandv\u00e5rdsr\u00e4kning fr\u00e5n scenario 1). Betalningsanm\u00e4rkning hotar om den inte betalas inom 14 dagar.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Diskussionsfr\u00e5ga: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: 'Alex har tre akuta problem och pengar till noll av dem. Vad prioriterar ni \u2014 och vad offrar ni? Finns det ett "r\u00e4tt" svar? Hur p\u00e5verkar social press (K\u00f6penhamn) beslutet trots att det objektivt sett \u00e4r det minst akuta problemet?', font: FONT, size: 24, italics: true })
          ]}),
        ]),

        spacer(),
        zoomBox("En betalningsanm\u00e4rkning g\u00f6r att Alex inte kan teckna hyreskontrakt, mobilabonnemang eller l\u00e5n i 3 \u00e5r. Hur p\u00e5verkar det framtida valm\u00f6jligheter? \u00c4r det proportionerligt att 2 200 kr kan st\u00e4nga d\u00f6rrar i 3 \u00e5r?"),

        new Paragraph({ children: [new PageBreak()] }),

        // --- SCENARIO 4 ---
        scenarioBox("Scenario 4: Systemfr\u00e5gan \u2014 Zoom ut p\u00e5 alla tre scenarierna", [
          new Paragraph({ spacing: { after: 120 }, children: [
            new TextRun({ text: "Alex kollega Fatima har g\u00e5tt igenom exakt samma kedja \u2014 tandl\u00e4kare, nedsk\u00e4rning, skulder \u2014 men med en avg\u00f6rande skillnad: Fatima \u00e4r inhyrd via bemanningsf\u00f6retag. Hon har ingen uppsa\u0308gningstid, inga semesterdagar, ingen sjukpenning, och blev av med jobbet helt (inte bara nedsk\u00e4rning) samma vecka.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Fakta att diskutera:", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "2023 hade 18% av unga 18\u201325 en betalningsanm\u00e4rkning", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "Genomsnittlig skuld hos Kronofogden f\u00f6r unga: 89 000 kr", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 40 }, children: [
            new TextRun({ text: "Snabbl\u00e5nebranschen oms\u00e4tter ca 8 miljarder kr/\u00e5r \u2014 med 30\u201340% r\u00e4nta", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "Andelen unga med visstidsanst\u00e4llning: 55% (j\u00e4mf\u00f6rt med 15% totalt)", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Perspektiv\u00f6vning \u2014 v\u00e4lj ETT perspektiv och argumentera:", font: FONT, size: 24, bold: true })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "Individperspektivet: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "Alex borde ha sparat mer, sagt nej till K\u00f6penhamn, och inte tagit snabbl\u00e5net. Eget ansvar.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [
            new TextRun({ text: "F\u00f6retagsperspektivet: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "Klarna och snabbl\u00e5nef\u00f6retagen erbjuder en tj\u00e4nst folk efterfr\u00e5gar. Arbetsgivaren m\u00e5ste anpassa efter marknaden.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 }, children: [
            new TextRun({ text: "Samh\u00e4llsperspektivet: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "Systemet \u00e4r designat s\u00e5 att unga med l\u00e5g inkomst och os\u00e4kra anst\u00e4llningar n\u00e4stan oundvikligen hamnar i skuldf\u00e4llan. Det kr\u00e4vs politiska \u00e5tg\u00e4rder.", font: FONT, size: 24 })
          ]}),
          new Paragraph({ spacing: { after: 80 }, children: [
            new TextRun({ text: "Diskussionsfr\u00e5ga: ", font: FONT, size: 24, bold: true }),
            new TextRun({ text: "T\u00e4nk tillbaka p\u00e5 hela kedjan. Var i kedjan hade en annan struktur (h\u00f6gre ing\u00e5ngsl\u00f6n, billigare tandv\u00e5rd, f\u00f6rbud mot snabbl\u00e5n, tryggare anst\u00e4llning) kunnat bryta spiralen? Vilket ingrepp hade gett st\u00f6rst effekt?", font: FONT, size: 24, italics: true })
          ]}),
        ]),

        spacer(),
        zoomBox("Om en generation unga hamnar i skuldspiralens logik \u2014 vad h\u00e4nder med bostadsmarknaden, konsumtionen, och i f\u00f6rl\u00e4ngningen skatteint\u00e4kterna? Rita kopplingen till det ekonomiska kretsloppet fr\u00e5n lektion 1."),

        new Paragraph({ children: [new PageBreak()] }),

        // --- L\u00c4RARINSTRUKTIONER ---
        h2("L\u00e4rarinstruktioner"),
        boldBullet("Uppstart:", "Retrieval practice fr\u00e5n lektion 3 \u2014 koppla direkt till budgetarbetet. \u201CHur mycket hade Alex kvar?\u201D skapar en naturlig \u00f6verg\u00e5ng till ov\u00e4ntade utgifter."),
        boldBullet("Instruktion:", "R\u00e4kna konkret p\u00e5 tavlan: snabbl\u00e5n 10 000 kr \u00d7 35% r\u00e4nta. L\u00e5t eleverna reagera p\u00e5 totalkostnaden. H\u00e5ll det kort \u2014 den stora bearbetningen sker i scenariokedjan."),
        para("Bearbetning:", { bold: true }),
        para("Scenarierna bildar en kedja \u2014 betona att paren m\u00e5ste arbeta i ordning eftersom varje val f\u00e5r konsekvenser i n\u00e4sta scenario. Cirkulera och pusha:"),
        italicBullet("Scenario 1:", '"Vad h\u00e4nder med Alex buffertsparande om hen tar snabbl\u00e5net? Hur p\u00e5verkar det n\u00e4sta situation?" (mot C)'),
        italicBullet("Scenario 2:", '"Alex har redan en skuld ELLER ett tomt sparkonto \u2014 hur f\u00f6r\u00e4ndrar det valm\u00f6jligheterna?" (mot C/A)'),
        italicBullet("Scenario 3:", '"Nu har Alex flera ekonomiska problem samtidigt \u2014 vad g\u00f6r det med beslutskvaliteten?" (mot A)'),
        italicBullet("Scenario 4 (Zoom ut):", '"Sluta t\u00e4nka p\u00e5 Alex. T\u00e4nk p\u00e5 100 000 unga i samma situation." (mot A)'),
        boldBullet("Summering:", 'J\u00e4mf\u00f6r kedjorna \u2014 par som tog snabbl\u00e5net tidigt hade sannolikt s\u00e4mre f\u00f6ruts\u00e4ttningar i scenario 3\u20134. Po\u00e4ngen: ekonomiska beslut \u00e4r kumulativa, inte isolerade. Fr\u00e5ga: "Var det Alex d\u00e5liga val \u2014 eller d\u00e5liga f\u00f6ruts\u00e4ttningar \u2014 som ledde hit?"'),

        h2("Differentiering"),
        boldBullet("St\u00f6d (mot E):", 'St\u00f6dfr\u00e5gor p\u00e5 varje kort: "Vilka alternativ har Alex?" och "Vad h\u00e4nder om Alex v\u00e4ljer alternativ A? Alternativ B?" Eleverna fokuserar p\u00e5 diskussionsfr\u00e5gan och hoppar \u00f6ver "Zoom ut"-fr\u00e5gan. Scenario 4 kan g\u00f6ras muntligt i helklass.'),
        boldBullet("Utmaning (mot A):", 'Eleverna besvarar ALLA "Zoom ut"-fr\u00e5gor skriftligt. I scenario 4: argumentera fr\u00e5n ETT perspektiv i 5 minuter, byt sedan till ett annat perspektiv och argumentera lika \u00f6vertygande. Reflektera: "Vilket perspektiv var sv\u00e5rast att f\u00f6rsvara \u2014 och varf\u00f6r?"'),

        h2("Koppling till kunskapskrav"),
        boldBullet("E-niv\u00e5:", "Identifiera alternativ och ge enkla resonemang om konsekvenser av ekonomiska val."),
        boldBullet("C-niv\u00e5:", "F\u00f6ra v\u00e4lgrundade resonemang om kort- och l\u00e5ngsiktiga konsekvenser och visa p\u00e5 samband."),
        boldBullet("A-niv\u00e5:", "F\u00f6ra v\u00e4lgrundade och nyanserade resonemang ur flera perspektiv (individ, f\u00f6retag, samh\u00e4lle) och visa p\u00e5 komplexa samband mellan privatekonomi och samh\u00e4llsekonomi."),
      ]
    }]
  });
}

async function main() {
  const doc = lektion4();
  const buffer = await Packer.toBuffer(doc);
  const outputPath = `${__dirname}/lektion-4.docx`;
  fs.writeFileSync(outputPath, buffer);
  console.log(`Generated: ${outputPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
