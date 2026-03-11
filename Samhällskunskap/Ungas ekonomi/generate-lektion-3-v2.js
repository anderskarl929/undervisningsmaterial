const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "2E75B6" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

const FULL_WIDTH = 9360;

function heading(text, level) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function para(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text.map(t => typeof t === "string" ? new TextRun(t) : new TextRun(t))
    : [new TextRun(text)];
  return new Paragraph({ children: runs, spacing: { after: 120 }, ...opts });
}

function bold(text) {
  return { text, bold: true };
}

function bullet(textOrRuns, ref = "bullets", level = 0) {
  const children = Array.isArray(textOrRuns)
    ? textOrRuns.map(t => typeof t === "string" ? new TextRun(t) : new TextRun(t))
    : [new TextRun(textOrRuns)];
  return new Paragraph({ numbering: { reference: ref, level }, children, spacing: { after: 60 } });
}

function headerCell(text, width) {
  return new TableCell({
    borders: headerBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "2E75B6", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })]
  });
}

function cell(textOrChildren, width, opts = {}) {
  const children = Array.isArray(textOrChildren)
    ? textOrChildren
    : [new Paragraph({ children: [new TextRun({ text: textOrChildren, font: "Arial", size: 20 })], spacing: { after: 40 } })];
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    children
  });
}

function cellParagraph(runs) {
  const children = runs.map(r => typeof r === "string" ? new TextRun({ text: r, font: "Arial", size: 20 }) : new TextRun({ font: "Arial", size: 20, ...r }));
  return new Paragraph({ children, spacing: { after: 40 } });
}

// Budget table helper
function budgetTable(title, strategy, rows, taskText) {
  const colWidths = [2800, 3280, 3280];
  const content = [
    heading(title, HeadingLevel.HEADING_3),
    para([bold("Strategi: "), strategy]),
    new Table({
      width: { size: FULL_WIDTH, type: WidthType.DXA },
      columnWidths: colWidths,
      rows: [
        new TableRow({ children: [
          headerCell("Post", colWidths[0]),
          headerCell("Val A", colWidths[1]),
          headerCell("Val B", colWidths[2]),
        ]}),
        ...rows.map((r, i) => new TableRow({ children: [
          cell(r[0], colWidths[0], { shading: i % 2 === 0 ? "F5F5F5" : undefined }),
          cell(r[1], colWidths[1], { shading: i % 2 === 0 ? "F5F5F5" : undefined }),
          cell(r[2], colWidths[2], { shading: i % 2 === 0 ? "F5F5F5" : undefined }),
        ]}))
      ]
    }),
    para(""),
    para([bold("Uppgift: "), taskText]),
  ];
  return content;
}

// Händelsekort helper
function handelsekort(nummer, titel, beskrivning, fraga) {
  return [
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: "2E75B6", space: 4 } },
      spacing: { before: 240, after: 80 },
      children: [new TextRun({ text: `HÄNDELSE ${nummer}: ${titel}`, bold: true, font: "Arial", size: 22, color: "2E75B6" })]
    }),
    para(beskrivning),
    para([{ text: fraga, italics: true }]),
  ];
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1A1A1A" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "1A1A1A" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u2013", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ]},
    ]
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1273, bottom: 1440, left: 1273 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
            children: [
              new TextRun({ text: "Samhällskunskap 1a1 ", font: "Arial", size: 18, color: "666666" }),
              new TextRun({ text: " | ", font: "Arial", size: 18, color: "CCCCCC" }),
              new TextRun({ text: " Ungas ekonomi \u2014 Lektion 3 (v2)", font: "Arial", size: 18, color: "666666" }),
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Sida ", font: "Arial", size: 18, color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "999999" }),
            ]
          })]
        })
      },
      children: [
        // TITEL
        heading("Lektion 3: \u201CRäcker pengarna?\u201D", HeadingLevel.HEADING_1),
        para([bold("Privatekonomi och budget (v2)")]),
        para(""),
        para([bold("Kurs: "), "Samhällskunskap 1a1"]),
        para([bold("Moment: "), "Ungas ekonomi"]),
        para([bold("Lektionslängd: "), "60 minuter"]),

        // LÄRANDEMÅL
        heading("Lärandemål för lektionen", HeadingLevel.HEADING_2),
        bullet([bold("Mål 3: "), "Resonera om privatekonomiska val och deras konsekvenser"]),
        bullet([bold("Mål 4: "), "Diskutera sambandet individ\u2013samhälle \u2014 hur individens ekonomiska beslut påverkar och påverkas av samhällsekonomin"]),

        // FÖRBEREDELSE
        heading("Förberedelse", HeadingLevel.HEADING_2),
        bullet("Förbered presentation om privatekonomiska grunder (brutto/netto, skatt, fasta/rörliga utgifter)"),
        bullet("Skriv ut budgetblad med tre livsstilar \u2014 ett per par"),
        bullet("Skriv ut händelsekort \u2014 klipp isär, lägg i en hög"),
        bullet("Miniwhiteboard eller papper för uträkningar"),
        bullet("Valfritt: visa en budgetapp-screenshot som inspiration"),

        // TIDSPLANERING
        heading("Tidsplanering", HeadingLevel.HEADING_2),
        (() => {
          const colWidths = [1100, 1500, 2200, 4560];
          const rows = [
            ["0\u20135 min", "Uppstart", "Retrieval practice", "Vilka tre anställningsformer pratade vi om förra gången? Enskilt 2 min, par 2 min. Koppla: \u201CIdag får Alex lön \u2014 men räcker pengarna?\u201D"],
            ["5\u201318 min", "Instruktion", "Privatekonomi-grunder", "Bruttolön vs. nettolön, skatt, arbetsgivaravgift. Fasta vs. rörliga utgifter. Sparande och buffertsparande. Visa konkret: \u201CAlex får 22 000 kr i bruttolön. Hur mycket blir kvar?\u201D Räkna på tavlan."],
            ["18\u201332 min", "Bearbetning 1", "Grundbudget", "I par \u2014 välj EN av tre livsstilar och gör en fullständig månadsbudget. Fördela pengarna, räkna ut vad som blir kvar. Varje val har konsekvenser."],
            ["32\u201344 min", "Bearbetning 2", "Verklighetscheck", "Varje par drar ETT händelsekort. Justera budgeten: vad stryks? Vad prioriteras? Går det ihop \u2014 eller hamnar Alex i minus?"],
            ["44\u201350 min", "Bearbetning 3", "Jämförelse i fyrgrupp", "Två par jämför: Vilken livsstil klarade verkligheten bäst? Varför? Var det livsstilen eller händelsen som avgjorde?"],
            ["50\u201357 min", "Summering", "EPA \u2014 ansvarsfrågan", "EPA (E 2 min, P 2 min, A 3 min): \u201CÄr det Alex eget fel om hen hamnar i ekonomiska problem \u2014 eller handlar det om vilken lön arbetsmarknaden erbjuder unga?\u201D"],
            ["57\u201360 min", "Framåtkoppling", "Nästa gång", "\u201CNästa gång fördjupar vi oss i vad som händer när det verkligen går snett \u2014 skuldfällan, Klarna och konsumenträtt.\u201D"],
          ];
          return new Table({
            width: { size: FULL_WIDTH, type: WidthType.DXA },
            columnWidths: colWidths,
            rows: [
              new TableRow({ children: [
                headerCell("Tid", colWidths[0]),
                headerCell("Fas", colWidths[1]),
                headerCell("Aktivitet", colWidths[2]),
                headerCell("Beskrivning", colWidths[3]),
              ]}),
              ...rows.map((r, i) => new TableRow({ children: r.map((text, j) =>
                cell(text, colWidths[j], { shading: i % 2 === 0 ? "F0F6FB" : undefined })
              )}))
            ]
          });
        })(),

        // LÄRARINSTRUKTIONER
        heading("Lärarinstruktioner", HeadingLevel.HEADING_2),
        para([bold("Uppstart: "), "Retrieval practice från lektion 2. Koppla anställningsformer till inkomst: \u201COm Alex har timanställning \u2014 hur påverkar det budgeten jämfört med fast lön?\u201D"]),
        para([bold("Instruktion: "), "Räkna konkret på tavlan: 22 000 brutto \u2192 ca 17 600 netto. Förklara vart de 4 400 kronorna går (skatt). Koppla till kretsloppet: \u201CDet här är ett av flödena vi ritade i lektion 1!\u201D"]),
        para([bold("Fas 1: "), "Cirkulera och hjälp par som fastnar med uträkningarna. Uppmärksamma att varje livsstil har val med dolda konsekvenser \u2014 ställ frågor: \u201CNi valde ingen hemförsäkring \u2014 vad händer om det blir inbrott?\u201D Paren ska inte bara räkna utan förstå att varje post är ett aktivt val."]),
        para([bold("Fas 2: "), "Händelsekorten är designade att slå olika hårt beroende på livsstil. Pusha eleverna att vara konkreta: \u201CExakt vilken post stryker ni? Hur mycket sparar det?\u201D Uppmuntra kreativa lösningar, men kräv att budgeten går ihop."]),
        para([bold("Fas 3: "), "Den viktigaste diskussionen. Frågor att lyfta: \u201CVar det livsstilen som avgjorde \u2014 eller slumpen?\u201D (mot C). \u201CKan man säga att en livsstil är objektivt bättre, eller beror det på vilken händelse man råkar ut för?\u201D (mot A)."]),
        para([bold("Summering: "), "EPA-frågan om ansvar är medvetet polariserande. Lyft att det finns strukturella begränsningar (lön, hyresmarknad) och individuella val."]),

        // ALEX EKONOMI
        new Paragraph({ children: [new PageBreak()] }),
        heading("Alex ekonomi \u2014 utgångspunkt", HeadingLevel.HEADING_2),
        para([bold("Alex nettoinkomst: 17 600 kr/mån")]),
        para(""),
        para([bold("Fasta utgifter som alla livsstilar delar:")]),
        (() => {
          const colWidths = [5000, 4360];
          return new Table({
            width: { size: FULL_WIDTH, type: WidthType.DXA },
            columnWidths: colWidths,
            rows: [
              new TableRow({ children: [headerCell("Post", colWidths[0]), headerCell("Kostnad", colWidths[1])] }),
              new TableRow({ children: [cell("Telefon (abonnemang)", colWidths[0]), cell("399 kr", colWidths[1])] }),
              new TableRow({ children: [cell("Kollektivtrafik (månadskort)", colWidths[0], { shading: "F5F5F5" }), cell("950 kr", colWidths[1], { shading: "F5F5F5" })] }),
            ]
          });
        })(),
        para(""),
        para([bold("Totalt fast för alla: 1 349 kr"), " \u2014 resten beror på livsstilsval."]),

        // TRE LIVSSTILAR
        heading("Tre livsstilar \u2014 med val", HeadingLevel.HEADING_2),
        para("Varje livsstil har utgiftsposter där eleverna måste välja nivå. Valen påverkar hur mycket som blir kvar \u2014 och hur sårbar Alex är."),

        // Livsstil 1
        ...budgetTable(
          "Livsstil 1: \u201CSparsam Alex\u201D",
          "minimera utgifter, maximera sparande.",
          [
            ["Boende", "Dela rum i kollektiv \u2014 3 200 kr", "Dela tvåa med kompis \u2014 4 500 kr"],
            ["Mat", "Strikt matlådeplanering \u2014 2 200 kr", "Lagar hemma men flexibelt \u2014 2 800 kr"],
            ["Hemförsäkring", "Ingen \u2014 0 kr", "Grundförsäkring \u2014 150 kr"],
            ["El (andel)", "Ingår i hyran (kollektiv) / 350 kr", "Ingår i hyran / 350 kr"],
            ["Nöje & socialt", "Nästan inget \u2014 200 kr", "Fika med vänner ibland \u2014 500 kr"],
            ["Kläder", "Second hand vid behov \u2014 100 kr", "Second hand + basplagg \u2014 300 kr"],
            ["Hygien & hälsa", "Bara det nödvändigaste \u2014 200 kr", "Inkl. tandvård/friskvård \u2014 400 kr"],
            ["Sparande", "Målet: så mycket som möjligt", "Målet: så mycket som möjligt"],
          ],
          "Välj A eller B på varje post. Räkna ut totala utgifter. Hur mycket kan Sparsam Alex spara?"
        ),

        new Paragraph({ children: [new PageBreak()] }),

        // Livsstil 2
        ...budgetTable(
          "Livsstil 2: \u201CHär-och-nu-Alex\u201D",
          "leva bekvämt, njuta av livet nu.",
          [
            ["Boende", "Egen etta i förort \u2014 6 800 kr", "Egen etta centralt \u2014 8 200 kr"],
            ["Mat", "Blandar hemlagat och ute \u2014 3 800 kr", "Äter ute ofta \u2014 4 800 kr"],
            ["Hemförsäkring", "Grundförsäkring \u2014 150 kr", "Med drulleförsäkring \u2014 250 kr"],
            ["El", "450 kr", "450 kr"],
            ["Streaming & teknik", "Spotify + Netflix \u2014 250 kr", "Spotify + Netflix + gaming \u2014 450 kr"],
            ["Gym/träning", "Gymkort \u2014 399 kr", "Gymkort + PT 1 ggn/mån \u2014 799 kr"],
            ["Nöje & uteliv", "Uteliv varannan helg \u2014 1 200 kr", "Uteliv varje helg \u2014 2 000 kr"],
            ["Kläder", "Nytt varje månad \u2014 800 kr", "Märkeskläder \u2014 1 500 kr"],
            ["Sparande", "Det som blir kvar", "Det som blir kvar"],
          ],
          "Välj A eller B på varje post. Räkna ut totala utgifter. Blir det något kvar \u2014 eller hamnar Alex i minus?"
        ),

        para(""),

        // Livsstil 3
        ...budgetTable(
          "Livsstil 3: \u201CInvesterar-Alex\u201D",
          "leva stramt nu för att bygga framtiden.",
          [
            ["Boende", "Dela rum i kollektiv \u2014 3 200 kr", "Dela tvåa med kompis \u2014 4 500 kr"],
            ["Mat", "Strikt matlådeplanering \u2014 2 200 kr", "Lagar hemma men flexibelt \u2014 2 800 kr"],
            ["Hemförsäkring", "Grundförsäkring \u2014 150 kr", "Med drulleförsäkring \u2014 250 kr"],
            ["El (andel)", "Ingår i hyran / 350 kr", "Ingår i hyran / 350 kr"],
            ["Nöje & socialt", "Minimalt \u2014 300 kr", "Lite mer utrymme \u2014 600 kr"],
            ["Investering: körkort", "Intensivkurs \u2014 2 500 kr/mån (4 mån)", "Vanlig takt \u2014 1 500 kr/mån (7 mån)"],
            ["Investering: sparande", "Fondsparkonto \u2014 1 500 kr/mån", "Fondsparkonto \u2014 1 000 kr/mån"],
            ["Hygien & hälsa", "Bara det nödvändigaste \u2014 200 kr", "Inkl. friskvård \u2014 400 kr"],
          ],
          "Välj A eller B på varje post. Räkna ut totala utgifter. Går det ihop \u2014 eller måste Alex välja mellan körkort och sparande?"
        ),

        // HÄNDELSEKORT
        new Paragraph({ children: [new PageBreak()] }),
        heading("Händelsekort \u2014 \u201CVerklighetscheck\u201D", HeadingLevel.HEADING_2),
        para("Skriv ut och klipp isär. Varje par drar ett kort efter fas 1."),

        ...handelsekort(1, "Trasig mobil",
          "Alex mobil går sönder. En begagnad ersättare kostar 2 800 kr. Utan telefon fungerar varken jobb-appen, bankID eller kommunikation med chefen. Alex måste lösa det den här månaden.",
          "Hur löser Alex det? Vad stryks i budgeten? Finns det pengar på sparkontot?"),

        ...handelsekort(2, "Timmar som försvinner",
          "Alex arbetsgivare meddelar att det blir färre timmar nästa månad. Nettoinkomsten sjunker med 3 200 kr \u2014 till 14 400 kr. Alex vet inte om det är tillfälligt eller permanent.",
          "Gör om budgeten med den nya inkomsten. Vad måste bort? Vad händer om det fortsätter i tre månader?"),

        ...handelsekort(3, "Kompisen flyttar",
          "Alex kompis som hen delar boende med säger upp sig och flyttar hem till föräldrarna. Alex har en månad på sig att hitta en ny rumskompis \u2014 annars blir hyran dubbelt så hög nästa månad.",
          "Vad gör Alex om ingen ny kompis hittas? Hur påverkas budgeten? Vilka alternativ finns?"),

        ...handelsekort(4, "Tandvärk",
          "Alex har ont i en tand. Akuttandvård kostar 3 500 kr. Att vänta riskerar att göra det värre och dyrare (uppskattningsvis 8 000 kr om det blir rotfyllning).",
          "Vad väljer Alex \u2014 fixa nu eller vänta? Var kommer pengarna ifrån?"),

        ...handelsekort(5, "Oväntad möjlighet",
          "Alex blir erbjuden en kurs som kan leda till bättre jobb. Kursen kostar 4 500 kr och pågår kvällstid i sex veckor. Under kurstiden hinner Alex inte jobba lika mycket \u2014 inkomsten minskar med 1 800 kr den månaden.",
          "Total kostnad: 6 300 kr den månaden. Är det värt det? Hur finansierar Alex det?"),

        ...handelsekort(6, "Elräkningen skenar",
          "Det har varit en kall vinter. Alex elräkning blir 1 400 kr istället för de budgeterade 350\u2013450 kr. Dessutom höjer hyresvärden hyran med 300 kr/mån från och med nu.",
          "Justera budgeten. Hyreshöjningen är permanent \u2014 hur påverkar det på sikt?"),

        // ELEVAKTIVITETER
        new Paragraph({ children: [new PageBreak()] }),
        heading("Elevaktiviteter", HeadingLevel.HEADING_2),
        bullet("Enskilt: Retrieval practice från lektion 2 (2 min)"),
        bullet("Par: Dela och diskutera (2 min)"),
        bullet([bold("Par: Fas 1 \u2014 "), "välj livsstil, gör val inom den, räkna ut budget (14 min)"]),
        bullet([bold("Par: Fas 2 \u2014 "), "dra händelsekort, justera budgeten, dokumentera vad som stryks (12 min)"]),
        bullet([bold("Fyrgrupp: Fas 3 \u2014 "), "jämför livsstilar och händelser, diskutera sårbarhet (6 min)"]),
        bullet("EPA: Diskutera ansvarsfrågan \u2014 individ vs. samhälle (7 min)"),

        // DIFFERENTIERING
        heading("Differentiering", HeadingLevel.HEADING_2),
        para([bold("Stöd (mot E): "), "Budgetbladet har en exempelkolumn som visar hur man räknar. Stödfrågor: \u201CBörja med att räkna ihop de fasta utgifterna. Välj sedan A eller B på varje post \u2014 skriv ner beloppet. Räkna ihop allt. Hur mycket av 17 600 kr blir kvar?\u201D Vid händelsekortet: \u201CHur mycket kostar händelsen? Har Alex pengar på sparkontot? Om inte \u2014 vilken post kan Alex stryka eller minska?\u201D"]),
        para(""),
        para([bold("Utmaning (mot A): "), "Räkna på tre månader framåt med händelsen inräknad. Vad händer med sparandet? Reflektera: \u201CVilken livsstil bygger mest motståndskraft \u2014 och till vilket pris? Vad händer med kretsloppet om en hel generation unga lever som Sparsam Alex \u2014 sparar mycket men konsumerar minimalt?\u201D"]),

        // MATERIAL
        heading("Material", HeadingLevel.HEADING_2),
        bullet("Presentation om privatekonomiska grunder"),
        bullet("Budgetblad med tre livsstilar och valmöjligheter (utskrift, ett per par)"),
        bullet("Händelsekort (utskrift, klipp isär)"),
        bullet("Miniwhiteboard eller papper för uträkningar"),

        // KOPPLING TILL KUNSKAPSKRAV
        heading("Koppling till kunskapskrav", HeadingLevel.HEADING_2),
        para("Lektionen bygger mål 3 (resonera om privatekonomiska val) och mål 4 (individ\u2013samhälle) genom praktisk tillämpning med ökad komplexitet."),
        para(""),
        bullet([bold("E-nivå: "), "Redogöra översiktligt för privatekonomiska begrepp (inkomst, utgift, budget) och göra en enkel budget. Identifiera konsekvenser av händelsekortet."]),
        bullet([bold("C-nivå: "), "Redogöra utförligt och resonera välgrundat om konsekvenser av olika prioriteringar \u2014 både i grundbudgeten och vid oväntade händelser. Jämföra livsstilar."]),
        bullet([bold("A-nivå: "), "Resonera välgrundat och nyanserat om privatekonomiska val, analysera sårbarhet i olika livsstilar, koppla till komplexa samband i samhällsekonomin (kretsloppet, välfärd, konsumtionens roll)."]),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samhällskunskap/Ungas ekonomi/lektion-3-v2.docx", buffer);
  console.log("lektion-3-v2.docx skapad!");
});
