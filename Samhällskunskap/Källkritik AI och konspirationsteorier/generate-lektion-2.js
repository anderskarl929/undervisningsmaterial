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
        children: [new TextRun({ text: "Lektion 2: Vem tj\u00E4nar p\u00E5 att du tror det h\u00E4r?", font: "Arial" })],
      }),
      bodyText("Konspirationsteorier \u2014 hur de uppst\u00E5r och sprids"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "85 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna granska och v\u00E4rdera k\u00E4llor, resonera om hur konspirationsteorier uppst\u00E5r, sprids och p\u00E5verkar samh\u00E4llet, samt p\u00E5b\u00F6rja korrekt k\u00E4llh\u00E4nvisning."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av de fyra k\u00E4llkritiska grundfr\u00E5gorna (m\u00E5l 1)"),
      bullet("Resonera om hur konspirationsteorier uppst\u00E5r, sprids och p\u00E5verkar samh\u00E4llet (m\u00E5l 3)"),
      bullet("K\u00E4llh\u00E4nvisa korrekt med Harvardsystemet p\u00E5 grundl\u00E4ggande niv\u00E5 (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men, f\u00F6r enkla resonemang om hur konspirationsteorier sprids och k\u00E4llh\u00E4nvisar med viss s\u00E4kerhet."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men, f\u00F6r v\u00E4lgrundade resonemang om hur konspirationsteorier uppst\u00E5r och sprids, och k\u00E4llh\u00E4nvisar i texten."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men, f\u00F6r v\u00E4lgrundade och nyanserade resonemang om hur konspirationsteorier uppst\u00E5r, sprids och f\u00F6rst\u00E4rks ur flera perspektiv, och k\u00E4llh\u00E4nvisar korrekt."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 1 \u2014 notera vilka grundfr\u00E5gor eleverna beh\u00E4rskar och vilka som beh\u00F6ver repeteras"),
      bullet("F\u00F6rbered retrieval practice-fr\u00E5gor baserat p\u00E5 lektion 1:s exit tickets"),
      bullet("F\u00F6rbered presentation om konspirationsteoriers mekanismer (se inneh\u00E5ll under Instruktion)"),
      bullet("Skriv ut case-korten (4 konspirationsteorier, en per grupp, se bilaga 1)"),
      bullet("Skriv ut analysmodellen \u201CAnatomin av en konspirationsteori\u201D (en per elev, se bilaga 2)"),
      bullet("Skriv ut spridningskartan (en per grupp, se bilaga 3)"),
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
              cell("Retrieval practice + hook", 2200),
              cell("Snabbskrivning (3 min): \u201CSkiv ner de fyra k\u00E4llkritiska grundfr\u00E5gorna.\u201D Gemensam genomg\u00E5ng. Sedan hook: visa ett p\u00E5st\u00E5ende \u2014 \u201CDen h\u00E4r teorin tror 30% av amerikanerna p\u00E5.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201325 min", 1000),
              cell("Instruktion", 1300),
              cell("Vad \u00E4r en konspirationsteori?", 2200),
              cell("Miniförel\u00E4sning: Definition, tre k\u00E4nnetecken (hemlig grupp, stor plan, \u201Cbevisen\u201D bekr\u00E4ftar alltid). Spridningsmekanismer: algoritmer, ekkammare, k\u00E4nslor. EPA-stopp vid minut 18.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("25\u201355 min", 1000),
              cell("Bearbetning", 1300),
              cell("Case-analys i grupp", 2200),
              cell("Grupper om 3\u20134 f\u00E5r varsitt case (konspirationsteori). Analysera med modellen \u201CAnatomin av en konspirationsteori\u201D. Identifiera: Vad p\u00E5st\u00E5s? Vilka \u201Cbevis\u201D anv\u00E4nds? Vem tj\u00E4nar p\u00E5 den? Hur sprids den?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("55\u201362 min", 1000),
              cell("Bearbetning", 1300),
              cell("Spridningskarta + presentation", 2200),
              cell("Grupperna ritar en spridningskarta f\u00F6r sin teori (5 min). 2\u20133 grupper presenterar muntligt (2 min/grupp), \u00F6vriga ger skriftlig feedback p\u00E5 en lapp.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("62\u201371 min", 1000),
              cell("Instruktion", 1300),
              cell("K\u00E4llh\u00E4nvisning \u2014 kort intro till Harvard", 2200),
              cell("Varf\u00F6r k\u00E4llh\u00E4nvisa? (2 min). Hur fungerar Harvard? Visa 2\u20133 exempel p\u00E5 tavlan (4 min). \u00D6vning: \u201CSkiv en Harvard-h\u00E4nvisning till en av k\u00E4llorna vi anv\u00E4nt idag\u201D (3 min).", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("71\u201382 min", 1000),
              cell("Summering", 1300),
              cell("Helklassdiskussion + exit ticket", 2200),
              cell("Diskussion: Vilka m\u00F6nster ser vi? Varf\u00F6r tror m\u00E4nniskor p\u00E5 konspirationsteorier? Exit ticket: \u201CF\u00F6rklara tv\u00E5 anledningar till att konspirationsteorier sprids och ge ett exempel.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("82\u201385 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CN\u00E4sta g\u00E5ng sl\u00E5r vi ihop allt: AI + konspirationsteorier. Ni kommer ocks\u00E5 \u00F6va mer p\u00E5 k\u00E4llh\u00E4nvisning.\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart (0\u20138 min)"),
      boldBodyText("Retrieval practice (0\u20135 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar med dagens \u00E4mne \u2014 en snabb \u00E5terblick. Skriv ner de fyra k\u00E4llkritiska grundfr\u00E5gorna fr\u00E5n f\u00F6rra lektionen. Ni har 2 minuter.\u201D"),
      bodyText("Samla svar muntligt. Fyll i eventuella luckor. Om exit tickets fr\u00E5n lektion 1 visade att en specifik grundfr\u00E5ga var sv\u00E5r, l\u00E4gg extra tid h\u00E4r."),
      spacer(),
      boldBodyText("Hook (5\u20138 min): ", ""),
      bodyText("Visa p\u00E5 sk\u00E4rmen:"),
      italicText("\u201C30% av amerikanerna tror att staten medvetet d\u00F6ljer information om uf:on.\u201D (Gallup, 2021)"),
      spacer(),
      italicText("Fr\u00E5ga: \u201CVarf\u00F6r tror s\u00E5 m\u00E5nga p\u00E5 n\u00E5got som inte har vetenskapligt st\u00F6d? Det \u00E4r precis vad vi ska unders\u00F6ka idag.\u201D"),

      heading3("Instruktion (8\u201325 min)"),
      bodyText("Presentera konspirationsteoriers mekanismer. Max 17 minuter. Skriv nyckelbegreppen p\u00E5 tavlan."),
      spacer(),
      boldBodyText("1. Vad \u00E4r en konspirationsteori? (4 min)", ""),
      bodyText("Definition: En f\u00F6rklaring som h\u00E4vdar att en hemlig, m\u00E4ktig grupp i hemlighet ligger bakom viktiga h\u00E4ndelser."),
      bodyText("Tre k\u00E4nnetecken:"),
      bullet("En hemlig grupp med makt (staten, \u201Celiten\u201D, f\u00F6retag)"),
      bullet("En stor, d\u00F6ljd plan (kontrollera, skada, manipulera)"),
      bullet("Alla \u201Cbevis\u201D bekr\u00E4ftar teorin \u2014 \u00E4ven avsaknad av bevis (\u201CDet bevisar att de d\u00F6ljer det\u201D)"),
      spacer(),
      boldBodyText("2. Varf\u00F6r tror m\u00E4nniskor p\u00E5 dem? (4 min)", ""),
      bullet("Behov av f\u00F6rklaringar: sv\u00E5ra h\u00E4ndelser k\u00E4nns kaotiska \u2014 en konspirationsteori ger en \u201Clogisk\u201D f\u00F6rklaring"),
      bullet("Misstro mot auktoriteter: n\u00E4r f\u00F6rtroendet f\u00F6r myndigheter och medier \u00E4r l\u00E5gt"),
      bullet("Grupptillh\u00F6righet: att dela en teori skapar gemenskap och identitet"),
      bullet("Confirmation bias: vi s\u00F6ker information som bekr\u00E4ftar det vi redan tror"),
      spacer(),
      boldItalicBody("EPA-stopp (4 min): ", "\u201CT\u00E4nk enskilt i 1 minut: Vilken av dessa anledningar tror du \u00E4r den starkaste drivkraften? Diskutera med din granne i 2 minuter.\u201D Samla 2\u20133 svar."),
      spacer(),
      boldBodyText("3. Hur sprids de? (5 min)", ""),
      bullet("Sociala medier: algoritmer som f\u00F6rst\u00E4rker engagerande inneh\u00E5ll"),
      bullet("Ekkammare: du ser mer av det du redan interagerar med"),
      bullet("K\u00E4nslor: r\u00E4dsla, ilska och misstro sprids snabbare \u00E4n fakta"),
      bullet("Influencers och alternativmedier: trov\u00E4rdiga avs\u00E4ndare inom subkulturer"),
      spacer(),
      italicText("Koppla tillbaka till grundfr\u00E5gorna: \u201CSe hur de k\u00E4llkritiska fr\u00E5gorna fr\u00E5n f\u00F6rra lektionen hj\u00E4lper oss \u00E4ven h\u00E4r. Vem sprider? Varf\u00F6r? Hur presenteras det?\u201D"),

      heading3("Bearbetning del 1: Case-analys (25\u201355 min)"),
      bodyText("Dela in eleverna i grupper om 3\u20134. Varje grupp f\u00E5r ETT case-kort (se bilaga 1) med en konspirationsteori och analysmodellen \u201CAnatomin av en konspirationsteori\u201D (se bilaga 2)."),
      spacer(),
      italicText("Instruktion: \u201CNi f\u00E5r en konspirationsteori att analysera. Anv\u00E4nd modellen f\u00F6r att dissekera den: Vad p\u00E5st\u00E5s? Vilka \u2018bevis\u2019 anv\u00E4nds? Vem tj\u00E4nar p\u00E5 att folk tror p\u00E5 den? Hur sprids den? Ni har 20 minuter.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 skriv p\u00E5 tavlan:"),
      bullet("Steg 1: L\u00E4s igenom caset tillsammans"),
      bullet("Steg 2: Fyll i analysmodellen gemensamt"),
      bullet("Steg 3: Diskutera \u2014 varf\u00F6r tror m\u00E4nniskor p\u00E5 just den h\u00E4r teorin?"),
      spacer(),
      bodyText("Cirkulera mellan grupperna. Till grupper som fastnar: \u201CB\u00F6rja med det enklaste \u2014 vad p\u00E5st\u00E5r teorin? Vem \u00E4r den hemliga gruppen?\u201D"),
      bodyText("Till snabba grupper: \u201CKan ni se kopplingar till de psykologiska drivkrafterna vi pratade om? Vilka k\u00E4nslor spelar teorin p\u00E5?\u201D"),

      heading3("Bearbetning del 2: Spridningskarta + presentation (55\u201362 min)"),
      bodyText("Dela ut spridningskartan (se bilaga 3). Varje grupp ritar hur deras konspirationsteori sprids."),
      spacer(),
      italicText("Instruktion: \u201CRita en karta \u00F6ver hur er teori sprids. B\u00F6rja med: Vem skapade/startade den? Hur n\u00E5dde den fler m\u00E4nniskor? Vilka plattformar, grupper och mekanismer var involverade? Ni har 5 minuter.\u201D"),
      spacer(),
      bodyText("Efter 5 minuter: 2\u20133 grupper presenterar muntligt f\u00F6r klassen (max 2 minuter per grupp). \u00D6vriga grupper ger skriftlig feedback p\u00E5 en lapp: \u201CEtt m\u00F6nster vi ocks\u00E5 s\u00E5g i v\u00E5rt case:\u201D och \u201CEn skillnad mot v\u00E5rt case:\u201D."),
      spacer(),
      bodyText("Samla in lapparna. Lyft 1\u20132 gemensamma drag muntligt:"),
      italicText("\u201CVilka m\u00F6nster ser ni? Finns det gemensamma drag i hur de h\u00E4r teorierna sprids?\u201D"),

      heading3("K\u00E4llh\u00E4nvisning \u2014 kort intro till Harvard (62\u201371 min)"),
      italicText("S\u00E4g: \u201CInnan vi avslutar \u2014 en viktig sak. N\u00E4r ni skriver om k\u00E4llor beh\u00F6ver ni kunna visa VAR ni hittat informationen. Det kallas k\u00E4llh\u00E4nvisning.\u201D"),
      spacer(),
      boldBodyText("1. Varf\u00F6r k\u00E4llh\u00E4nvisa? (2 min)", ""),
      bullet("F\u00F6r att l\u00E4saren ska kunna kontrollera dina p\u00E5st\u00E5enden"),
      bullet("F\u00F6r att visa att du bygger p\u00E5 trov\u00E4rdiga k\u00E4llor"),
      bullet("F\u00F6r att skilja dina egna \u00E5sikter fr\u00E5n andras"),
      spacer(),
      boldBodyText("2. Hur fungerar Harvardsystemet? Visa 2\u20133 exempel p\u00E5 tavlan (4 min)", ""),
      bodyText("Skriv p\u00E5 tavlan:"),
      boldBodyText("I texten: ", "(Efternamn, \u00E5r). Exempel: \u201C30% av amerikanerna tror p\u00E5 ufon (Gallup, 2021).\u201D"),
      boldBodyText("I referenslistan: ", "Efternamn/Organisation (\u00C5r). Titel. K\u00E4lla. URL"),
      spacer(),
      bodyText("Visa 2\u20133 exempel:"),
      bullet("Gallup (2021). Americans\u2019 Belief in UFOs. gallup.com/poll/350096/..."),
      bullet("Kahneman, D. (2011). Thinking, Fast and Slow. Farrar, Straus and Giroux."),
      spacer(),
      boldBodyText("3. \u00D6vning (3 min)", ""),
      italicText("S\u00E4g: \u201CV\u00E4lj en av k\u00E4llorna vi anv\u00E4nt idag \u2014 t.ex. fr\u00E5n ert case-kort. Skriv en Harvard-h\u00E4nvisning: b\u00E5de i-texten-varianten och referenslistan. Ni har 3 minuter.\u201D"),
      bodyText("Cirkulera och ge snabb \u00E5terkoppling. Vanliga fel: gl\u00F6mmer \u00E5r, gl\u00F6mmer parentes."),

      heading3("Summering (71\u201382 min)"),
      bodyText("Helklassdiskussion (5 min):"),
      bullet("\u201CVarf\u00F6r \u00E4r konspirationsteorier ett problem f\u00F6r samh\u00E4llet?\u201D"),
      bullet("\u201CKan det finnas situationer d\u00E4r misstro mot myndigheter \u00E4r befogad? Hur skiljer vi det fr\u00E5n konspirationsteorier?\u201D"),
      spacer(),
      bodyText("Koppla tillbaka till l\u00E4randem\u00E5let p\u00E5 tavlan."),
      spacer(),
      boldBodyText("Exit ticket (7 min):", ""),
      italicText("\u201CF\u00F6rklara tv\u00E5 anledningar till att konspirationsteorier sprids. Ge ett konkret exempel p\u00E5 hur en av anledningarna fungerar i praktiken.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 meningsstartare p\u00E5 tavlan:"),
      bullet("\u201CEn anledning till att konspirationsteorier sprids \u00E4r att...\u201D"),
      bullet("\u201CEn annan anledning \u00E4r att...\u201D"),
      bullet("\u201CEtt exempel p\u00E5 detta \u00E4r n\u00E4r...\u201D"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r retrieval practice i lektion 3."),

      heading3("Fram\u00E5tkoppling (82\u201385 min)"),
      italicText("\u201CNu f\u00F6rst\u00E5r vi hur konspirationsteorier fungerar \u2014 och vi har verktygen fr\u00E5n f\u00F6rra lektionen f\u00F6r att granska k\u00E4llor. Ni har ocks\u00E5 sett hur k\u00E4llh\u00E4nvisning fungerar \u2014 det kommer ni \u00F6va mer p\u00E5 i kommande lektioner. N\u00E4sta g\u00E5ng sl\u00E5r vi ihop b\u00E5da delarna: Vad h\u00E4nder n\u00E4r AI anv\u00E4nds f\u00F6r att skapa och sprida konspirationsteorier? D\u00E5 blir det riktigt intressant.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: skriva ner grundfr\u00E5gorna (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "EPA: starkaste drivkraften bakom konspirationsteorier (enskilt + par, 4 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Case-analys med analysmodell (grupp 3\u20134, 20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Spridningskarta: rita spridningsm\u00F6nster (grupp, 8 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Gruppresentation f\u00F6r klassen (2\u20133 grupper muntligt, \u00F6vriga skriftlig feedback)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "K\u00E4llh\u00E4nvisnings\u00F6vning: skriv en Harvard-h\u00E4nvisning (enskilt, 3 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: spridningsorsaker + exempel (enskilt, 7 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 53 av 85 minuter (62%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Analysmodellen har ifyllda exempelf\u00E4lt. Stegen f\u00F6r case-analysen \u00E4r synliga p\u00E5 tavlan. Vid spridningskartan: ge en f\u00E4rdig startpunkt (\u201CB\u00F6rja med rutan \u2018Skapare\u2019 och dra en pil till n\u00E4sta steg\u201D). Exit ticket har meningsstartare. L\u00E4raren prioriterar dessa grupper vid cirkulering."),
      boldBodyText("Utmaning (mot A): ", "Analysera utan modell \u2014 fritt resonemang. Till\u00E4ggsfr\u00E5ga vid case-analysen: \u201CKan konspirationsteorier n\u00E5gonsin ha en demokratisk funktion? N\u00E4r blir misstro mot myndigheter befogad?\u201D Vid spridningskartan: l\u00E4gg till pilar f\u00F6r \u201Cmotspridning\u201D \u2014 vem och vad motverkar teorin? Exit ticket utan meningsstartare, med till\u00E4gget: \u201CResonera om samh\u00E4llskonsekvenser.\u201D"),

      // MATERIAL
      heading2("Material"),
      bullet("Exit ticket-data fr\u00E5n lektion 1 (f\u00F6r retrieval practice)"),
      bullet("Presentation om konspirationsteoriers mekanismer"),
      bullet("Case-kort \u2014 4 konspirationsteorier (se bilaga 1)"),
      bullet("Analysmodell \u201CAnatomin av en konspirationsteori\u201D \u2014 en per elev (se bilaga 2)"),
      bullet("Spridningskarta \u2014 en per grupp (se bilaga 3)"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Retrieval practice repeterar grundfr\u00E5gorna och f\u00F6rst\u00E4rker l\u00E5ngtidsminnet (m\u00E5l 1)"),
      bullet("Case-analysen tr\u00E4nar f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor i ett nytt sammanhang (m\u00E5l 1, 3)"),
      bullet("Spridningskartan och presentationen tr\u00E4nar resonemang om spridningsmekanismer (m\u00E5l 3: E\u2013A)"),
      bullet("Exit ticket m\u00E4ter b\u00E5de faktakunskap och analysf\u00F6rm\u00E5ga (m\u00E5l 3)"),
      bullet("K\u00E4llh\u00E4nvisnings\u00F6vningen introducerar Harvard-systemet p\u00E5 grundl\u00E4ggande niv\u00E5 (m\u00E5l 4)"),
      spacer(),
      bodyText("Case-analysens scaffolding (analysmodell, steg p\u00E5 tavlan) s\u00E4kerst\u00E4ller att alla elever kan n\u00E5 E-niv\u00E5. De \u00F6ppna fr\u00E5gorna om demokrati och \u201Cmotspridning\u201D ger A-elever m\u00F6jlighet att visa nyanserat resonemang fr\u00E5n flera perspektiv."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1 (k\u00E4llkritiska grundfr\u00E5gor)"),
      bullet("Introducerar k\u00E4llh\u00E4nvisning (Harvard) \u2014 \u00F6vas vidare i lektion 3 och fram\u00E5t"),
      bullet("N\u00E4sta lektion: \u201CVerktygsl\u00E5dan i praktiken\u201D \u2014 AI + konspirationsteorier kombineras"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r retrieval practice i lektion 3"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: CASE-KORT
      heading2("Bilaga 1: Case-kort \u2014 Konspirationsteorier"),
      bodyText("Dela ut ett case per grupp. Varje kort inneh\u00E5ller en kort beskrivning av en konspirationsteori."),
      spacer(),

      boldBodyText("Case A: \u201CMoonlandningen var fejkad\u201D", ""),
      bodyText("P\u00E5st\u00E5ende: USA landade aldrig p\u00E5 m\u00E5nen 1969. Hela Apollo-programmet var en studioinspelning regisserad av NASA och den amerikanska regeringen f\u00F6r att vinna kappl\u00F6pningen mot Sovjetunionen."),
      bodyText("\u201CBevis\u201D som anf\u00F6rs: Flaggan \u201Cvajar\u201D trots att det inte finns n\u00E5gon vind p\u00E5 m\u00E5nen. Stj\u00E4rnorna syns inte p\u00E5 bilderna. Skuggorna faller i olika riktningar (vilket skulle tyda p\u00E5 studiobelysning)."),
      bodyText("Kontext: Teorin uppstod under kalla kriget och har spridits genom b\u00F6cker, dokumentarer och sociala medier. Den lever kvar trots att varje \u201Cbevis\u201D har motbevisats vetenskapligt."),
      spacer(),

      boldBodyText("Case B: \u201C5G orsakar covid-19\u201D", ""),
      bodyText("P\u00E5st\u00E5ende: Utbyggnaden av 5G-n\u00E4tverket orsakade eller f\u00F6rst\u00E4rkte pandemin. Str\u00E5lningen f\u00F6rsvagar immunf\u00F6rsvaret eller sprider viruset direkt."),
      bodyText("\u201CBevis\u201D som anf\u00F6rs: 5G-utbyggnaden skedde samtidigt som pandemin. Wuhan, d\u00E4r pandemin b\u00F6rjade, var en tidig 5G-stad. \u201CEgna unders\u00F6kningar\u201D p\u00E5 YouTube visar \u201Ckopplingar.\u201D"),
      bodyText("Kontext: Teorin ledde till att 5G-master vandaliserarades i flera l\u00E4nder 2020. Spreds framf\u00F6r allt via Facebook, YouTube och Telegram-grupper."),
      spacer(),

      boldBodyText("Case C: \u201CL\u00E4kemedelsf\u00F6retagen d\u00F6ljer botemedel mot cancer\u201D", ""),
      bodyText("P\u00E5st\u00E5ende: Stora l\u00E4kemedelsf\u00F6retag (\u201CBig Pharma\u201D) har tillg\u00E5ng till botemedel mot cancer men d\u00F6ljer dem medvetet f\u00F6r att tj\u00E4na pengar p\u00E5 dyra behandlingar."),
      bodyText("\u201CBevis\u201D som anf\u00F6rs: L\u00E4kemedelsindustrin \u00E4r enormt l\u00F6nsam. Vissa naturliga medel \u201Ckurerar\u201D cancer men trycks ner. L\u00E4kare som f\u00F6resl\u00E5r alternativa behandlingar \u201Ctystas.\u201D"),
      bodyText("Kontext: Teorin \u00E4r farlig eftersom den kan leda till att m\u00E4nniskor v\u00E4ljer bort vetenskaplig behandling. Sprids via alternativmedicinska sajter, sociala medier och YouTubers."),
      spacer(),

      boldBodyText("Case D: \u201CValfusk i amerikanska val\u201D", ""),
      bodyText("P\u00E5st\u00E5ende: Det amerikanska valet 2020 vanns genom systematiskt fusk. R\u00F6ster \u00E4ndrades elektroniskt, d\u00F6da personer r\u00F6stade, och valurnor fylldes med falska r\u00F6stsedlar."),
      bodyText("\u201CBevis\u201D som anf\u00F6rs: Videor fr\u00E5n valkontor som \u201Cvisar\u201D fusk. Statistiska \u201Canomalier\u201D i r\u00F6str\u00E4kningen. Vittnessm\u00E5l fr\u00E5n valobservat\u00F6rer."),
      bodyText("Kontext: Trots 60+ r\u00E4ttsliga processer hittades inga bevis f\u00F6r systematiskt fusk. Teorin ledde till st\u00F6rmningen av Kapitolium den 6 januari 2021. Spreds fr\u00E4mst via politiska medier och sociala plattformar."),

      // BILAGA 2: ANALYSMODELL
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Analysmodell \u2014 Anatomin av en konspirationsteori"),
      bodyText("Fyll i modellen f\u00F6r ert case."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2800, 6226],
        rows: [
          new TableRow({
            children: [
              headerCell("Fr\u00E5ga", 2800),
              headerCell("Ert svar", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vad p\u00E5st\u00E5s?", 2800),
              cell("(Sammanfatta teorin i 1\u20132 meningar)", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vem \u00E4r den \u201Chemliga gruppen\u201D?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilka \u201Cbevis\u201D anv\u00E4nds?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Varf\u00F6r h\u00E5ller \u201Cbevisen\u201D inte? Vad s\u00E4ger forskningen?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vem tj\u00E4nar p\u00E5 att folk tror p\u00E5 teorin?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilka k\u00E4nslor spelar teorin p\u00E5?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Hur sprids den? Via vilka kanaler?", 2800),
              cell("", 6226),
            ],
          }),
          new TableRow({
            children: [
              cell("Vilka samh\u00E4llskonsekvenser har den?", 2800),
              cell("", 6226),
            ],
          }),
        ],
      }),

      // BILAGA 3: SPRIDNINGSKARTA
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Spridningskarta"),
      bodyText("Rita pilar mellan rutorna och fyll i med information fr\u00E5n ert case. L\u00E4gg till egna rutor vid behov."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2256, 2257, 2256, 2257],
        rows: [
          new TableRow({
            children: [
              headerCell("1. Skapare", 2256),
              headerCell("2. Spridare", 2257),
              headerCell("3. Plattformar", 2256),
              headerCell("4. Mottagare", 2257),
            ],
          }),
          new TableRow({
            children: [
              cell("Vem startade teorin? Varf\u00F6r?", 2256),
              cell("Vilka spred den vidare? Influencers? Medier?", 2257),
              cell("Via vilka kanaler? YouTube, Facebook, Telegram?", 2256),
              cell("Vilka p\u00E5verkas mest? Varf\u00F6r just de?", 2257),
            ],
          }),
        ],
      }),

      spacer(),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3009, 3009, 3008],
        rows: [
          new TableRow({
            children: [
              headerCell("5. Mekanismer", 3009),
              headerCell("6. F\u00F6rst\u00E4rkare", 3009),
              headerCell("7. Konsekvenser", 3008),
            ],
          }),
          new TableRow({
            children: [
              cell("Algoritmer? Ekkammare? K\u00E4nslor? Confirmation bias?", 3009),
              cell("Vad g\u00F6r att den v\u00E4xer? Kriser? Misstro? Andra teorier?", 3009),
              cell("Vilka effekter har teorin haft p\u00E5 samh\u00E4llet?", 3008),
            ],
          }),
        ],
      }),

      spacer(),
      bodyText("Rita pilar och kopplingar mellan rutorna. Skriv p\u00E5 pilarna vad som driver spridningen vid varje steg."),

      // BILAGA 4: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: Exit ticket \u2014 Lektion 2"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "F\u00F6rklara tv\u00E5 anledningar till att konspirationsteorier sprids. Ge ett konkret exempel p\u00E5 hur en av anledningarna fungerar i praktiken."),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CEn anledning till att konspirationsteorier sprids \u00E4r att...\u201D"),
      bullet("\u201CEn annan anledning \u00E4r att...\u201D"),
      bullet("\u201CEtt exempel p\u00E5 detta \u00E4r n\u00E4r...\u201D"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-2.docx", buffer);
  console.log("lektion-2.docx skapad!");
});
