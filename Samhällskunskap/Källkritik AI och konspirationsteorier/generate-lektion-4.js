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
        children: [new TextRun({ text: "Lektion 4: Informationskriget", font: "Arial" })],
      }),
      bodyText("Seminarium \u2014 AI, desinformation och demokrati"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "100 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna till\u00E4mpa k\u00E4llkritiska verktyg i en samh\u00E4llsdebattkontext, analysera hur AI-genererat inneh\u00E5ll och konspirationsteorier p\u00E5verkar demokratin, samt argumentera f\u00F6r en position med st\u00F6d i k\u00E4llor."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av de k\u00E4llkritiska grundfr\u00E5gorna i ett nytt sammanhang (m\u00E5l 1)"),
      bullet("Resonera om hur AI-genererat inneh\u00E5ll p\u00E5verkar demokrati och samh\u00E4llsdebatt (m\u00E5l 2)"),
      bullet("Analysera hur konspirationsteorier f\u00F6rst\u00E4rks av AI och p\u00E5verkar samh\u00E4llet (m\u00E5l 3)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men och f\u00F6r enkla resonemang om hur AI och konspirationsteorier p\u00E5verkar samh\u00E4llsdebatten."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men och f\u00F6r v\u00E4lgrundade resonemang om hur AI och konspirationsteorier f\u00F6r\u00E4ndrar f\u00F6ruts\u00E4ttningarna f\u00F6r demokratisk debatt."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men och f\u00F6r v\u00E4lgrundade och nyanserade resonemang om hur AI och konspirationsteorier p\u00E5verkar demokratin, med komplexa kopplingar mellan teknik, makt och informationsfrihet."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 3 \u2014 notera vilka elever som beh\u00F6ver extra st\u00F6d i seminariet"),
      bullet("F\u00F6rbered retrieval practice-fr\u00E5gor baserat p\u00E5 lektion 1\u20133 (alla tre l\u00E4randem\u00E5l)"),
      bullet("Kopiera seminarieunderlaget (bilaga 1) \u2014 en per elev, dela ut MINST 24 timmar innan lektionen"),
      bullet("Kopiera bed\u00F6mningsmatrisen/observationsschemat (bilaga 2) \u2014 en per elev f\u00F6r l\u00E4raren"),
      bullet("Kopiera positionskorten (bilaga 3) \u2014 ett per grupp, dela ut tillsammans med seminarieunderlaget"),
      bullet("F\u00F6rbered gruppindelning: 4\u20135 grupper om 4\u20136 elever, blandad f\u00F6rm\u00E5geniv\u00E5"),
      bullet("St\u00E4ll stolar i en ring eller hestesko \u2014 alla ska kunna se varandra under seminariet"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),
      bullet("OBS: Detta \u00E4r f\u00F6rsta lektionen med minskad scaffolding \u2014 inga meningsstartare p\u00E5 tavlan under seminariet"),

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
              cell("0\u201310 min", 1000),
              cell("Uppstart", 1300),
              cell("Retrieval practice + ram\u00E4ttning", 2200),
              cell("Snabbskrivning (4 min): \u201CAnv\u00E4nd tre begrepp fr\u00E5n lektion 1\u20133 f\u00F6r att f\u00F6rklara hur AI-genererat inneh\u00E5ll kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier.\u201D Kort genomg\u00E5ng. Sedan: presentera dagens seminarium och spelregler.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("10\u201325 min", 1000),
              cell("F\u00F6rberedelse", 1300),
              cell("Gruppf\u00F6rberedelse av positioner", 2200),
              cell("Grupperna samlas kring sina tilldelade positioner (se bilaga 3). De l\u00E4ser igenom sitt positionskort, diskuterar argument och f\u00F6rbereder \u00F6ppningsanf\u00F6rande (1 min per grupp). St\u00F6dfr\u00E5gor: Vad \u00E4r er huvudtes? Vilka bevis har ni? Vilka motargument kan ni f\u00F6rv\u00E4nta er?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("25\u201370 min", 1000),
              cell("Seminarium", 1300),
              cell("Seminarium: Informationskriget", 2200),
              cell("Fas 1 (25\u201340): \u00D6ppningsanf\u00F6randen (1 min per grupp). Fas 2 (40\u201360): Fri diskussion med l\u00E4raren som moderator. Seminariefr\u00E5gor driver samtalet. Fas 3 (60\u201370): Syntes \u2014 \u201CVad \u00E4r vi \u00F6verens om? Var finns de verkliga sp\u00E4nningarna?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("70\u201380 min", 1000),
              cell("L\u00E4rpar", 1300),
              cell("L\u00E4rpar-\u00E5tergivning", 2200),
              cell("Eleverna paras ihop med n\u00E5gon fr\u00E5n en ANNAN grupp. Person A \u00E5terger vad Person B:s grupp argumenterade f\u00F6r (2 min). Person B r\u00E4ttar/kompletterar (1 min). Byt. Syftet: testa om eleverna lyssnade aktivt p\u00E5 andra perspektiv.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("80\u201395 min", 1000),
              cell("Summering", 1300),
              cell("Helklassreflektion + exit ticket", 2200),
              cell("Kort diskussion: \u201CVad var det st\u00E4rkaste argumentet ni h\u00F6rde fr\u00E5n en ANNAN grupp?\u201D Exit ticket: \u201CV\u00E4lj en av seminarifr\u00E5gorna och skriv ditt eget svar. Anv\u00E4nd minst tv\u00E5 begrepp fr\u00E5n momentet och h\u00E4nvisa till minst ett argument fr\u00E5n seminariet.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("95\u2013100 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CIdag hade ni f\u00F6rberedda positioner. N\u00E4sta g\u00E5ng h\u00F6jer vi insatsen: ni f\u00E5r tilldelade positioner som ni kanske INTE h\u00E5ller med om \u2014 och ska debattera dem. D\u00E5 testas er f\u00F6rm\u00E5ga att se saken fr\u00E5n flera sidor.\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart (0\u201310 min)"),
      boldBodyText("Retrieval practice (0\u20136 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi k\u00F6r ig\u00E5ng \u2014 en snabb \u00E5terblick p\u00E5 allt vi gjort hittills. Skriv ner: Anv\u00E4nd tre begrepp fr\u00E5n lektion 1\u20133 f\u00F6r att f\u00F6rklara hur AI-genererat inneh\u00E5ll kan anv\u00E4ndas f\u00F6r att sprida konspirationsteorier. Ni har 4 minuter.\u201D"),
      bodyText("Samla 3\u20134 svar muntligt. Lyft elever som lyckas koppla begrepp fr\u00E5n b\u00E5de lektion 1 (grundfr\u00E5gorna), lektion 2 (konspirationsteorier) och lektion 3 (AI + konspirationsteorier). Fyll i luckor om n\u00E5got viktigt saknas."),
      spacer(),
      boldBodyText("Ram f\u00F6r seminariet (6\u201310 min): ", ""),
      italicText("S\u00E4g: \u201CIdag g\u00F6r vi n\u00E5got nytt. Ni ska delta i ett seminarium om informationskriget \u2014 hur AI, desinformation och konspirationsteorier p\u00E5verkar demokratin. Ni har f\u00E5tt f\u00F6rberedelseunderlag och varje grupp har en position att f\u00F6rsvara.\u201D"),
      bodyText("Presentera spelreglerna (skriv p\u00E5 tavlan):"),
      bullet("Lyssna f\u00E4rdigt \u2014 avbryt inte"),
      bullet("Reagera p\u00E5 vad andra s\u00E4ger \u2014 bygg vidare eller ifr\u00E5gas\u00E4tt"),
      bullet("Anv\u00E4nd begrepp fr\u00E5n momentet"),
      bullet("Alla ska bidra \u2014 minst ett inl\u00E4gg per person"),
      spacer(),
      italicText("S\u00E4g: \u201CDet h\u00E4r \u00E4r f\u00F6rsta g\u00E5ngen vi k\u00F6r utan meningsstartare och hj\u00E4lpfr\u00E5gor p\u00E5 tavlan. Ni har verktygen fr\u00E5n tre lektioner \u2014 nu \u00E4r det dags att anv\u00E4nda dem.\u201D"),

      heading3("Gruppf\u00F6rberedelse (10\u201325 min)"),
      bodyText("Eleverna sitter i sina grupper (4\u20136 elever per grupp, 4\u20135 grupper). Varje grupp har f\u00E5tt ett positionskort (bilaga 3) tillsammans med seminarieunderlaget minst 24 timmar i f\u00F6rv\u00E4g."),
      spacer(),
      italicText("Instruktion: \u201CNi har 15 minuter att f\u00F6rbereda er. Ni ska: 1) Formulera ert \u00F6ppningsanf\u00F6rande (max 1 minut). 2) Samla era tre starkaste argument. 3) F\u00F6rbereda minst ett motargument mot varje annan grupp. V\u00E4lj vem som h\u00E5ller \u00F6ppningsanf\u00F6randet.\u201D"),
      spacer(),
      bodyText("Cirkulera och st\u00F6tta. F\u00F6r grupper som fastnar:"),
      bullet("\"Vad \u00E4r den viktigaste konsekvensen av er position?\""),
      bullet("\"Vilka bevis fr\u00E5n seminarieunderlaget st\u00F6djer ert argument?\""),
      bullet("\"Om n\u00E5gon fr\u00E5gar 'men \u00E4r det verkligen s\u00E5?' \u2014 vad svarar ni?\""),
      spacer(),
      bodyText("OBS: Scaffolding \u00E4r riktad \u2014 ge st\u00F6d till de grupper som beh\u00F6ver det, men l\u00E5t starkare grupper arbeta sj\u00E4lvst\u00E4ndigt. Det h\u00E4r \u00E4r medveten minskning av st\u00F6dstrukturer."),

      heading3("Seminarium (25\u201370 min)"),
      bodyText("Flytta stolar till en ring eller hestesko. L\u00E4raren sitter med men agerar moderator, inte deltagare."),
      spacer(),
      boldBodyText("Fas 1 \u2014 \u00D6ppningsanf\u00F6randen (25\u201340 min): ", ""),
      bodyText("Varje grupp h\u00E5ller sitt \u00F6ppningsanf\u00F6rande (max 1 minut per grupp). \u00D6vriga lyssnar utan att avbryta. L\u00E4raren antecknar nyckelargument p\u00E5 tavlan f\u00F6r att visualisera positionerna."),
      spacer(),
      italicText("S\u00E4g efter alla \u00F6ppningsanf\u00F6randen: \u201CNu har ni h\u00F6rt alla positioner. Nu \u00F6ppnar vi f\u00F6r fri diskussion. Jag kommer att styra med fr\u00E5gor, men det \u00E4r ni som ska driva samtalet.\u201D"),
      spacer(),
      boldBodyText("Fas 2 \u2014 Fri diskussion med seminariefr\u00E5gor (40\u201360 min): ", ""),
      bodyText("L\u00E4raren anv\u00E4nder seminariefr\u00E5gorna (se nedan) f\u00F6r att driva samtalet. Anpassa ordningen utifr\u00E5n vad som kommer upp naturligt. Byt fr\u00E5ga n\u00E4r diskussionen mattas av."),
      spacer(),
      boldBodyText("Seminariefr\u00E5gor (anv\u00E4nd 3\u20134 av dessa):", ""),
      spacer(),
      boldBodyText("1. ", "\u201CN\u00E4r AI kan generera nyhetsartiklar, bilder och videor som inte g\u00E5r att skilja fr\u00E5n \u00E4kta \u2014 vad h\u00E4nder d\u00E5 med v\u00E5r f\u00F6rm\u00E5ga att fatta informerade demokratiska beslut?\u201D"),
      boldBodyText("2. ", "\u201CBorde staten reglera AI-genererat inneh\u00E5ll h\u00E5rdare, eller \u00E4r det ett hot mot yttrandefriheten? Var g\u00E5r gr\u00E4nsen?\u201D"),
      boldBodyText("3. ", "\u201CVem b\u00E4r ansvaret n\u00E4r AI-genererad desinformation p\u00E5verkar ett val \u2014 teknikf\u00F6retagen, politikerna, eller medborgarna sj\u00E4lva?\u201D"),
      boldBodyText("4. ", "\u201CKonspirationsteorier har alltid funnits. Vad \u00E4r det som g\u00F6r att AI f\u00F6r\u00E4ndrar spelplanen? \u00C4r det verkligen v\u00E4rre nu?\u201D"),
      boldBodyText("5. ", "\u201CKan k\u00E4llkritik r\u00E4cka som f\u00F6rsvar mot AI-desinformation, eller beh\u00F6vs det n\u00E5got mer?\u201D"),
      spacer(),
      bodyText("Moderatortekniker:"),
      bullet("Passiva elever: \u201C[Namn], din grupp hade ett intressant argument om X \u2014 vill du utveckla?\u201D"),
      bullet("Dominanta elever: \u201CBra poäng. Finns det n\u00E5gon som ser det annorlunda?\u201D"),
      bullet("Ytliga svar: \u201CKan du ge ett konkret exempel fr\u00E5n seminarieunderlaget?\u201D"),
      bullet("St\u00E5nde diskussion: \u201COm vi v\u00E4nder p\u00E5 det \u2014 vad s\u00E4ger de som tycker tv\u00E4rtom?\u201D"),
      bullet("H\u00F6g niv\u00E5: \u201CDet \u00E4r ett sp\u00E4nnande argument. Vilka konsekvenser f\u00E5r det om vi f\u00F6ljer den logiken?\u201D"),
      spacer(),
      boldBodyText("Fas 3 \u2014 Syntes (60\u201370 min): ", ""),
      italicText("S\u00E4g: \u201CVi b\u00F6rjar runda av. L\u00E5t oss sammanfatta: Vad \u00E4r vi faktiskt \u00F6verens om? Och var finns de verkliga sp\u00E4nningarna \u2014 de fr\u00E5gor d\u00E4r det inte finns ett enkelt svar?\u201D"),
      bodyText("Sammanfatta p\u00E5 tavlan: V\u00E4nster sida = \u201C\u00D6verens\u201D, h\u00F6ger sida = \u201CSp\u00E4nningar/ol\u00F6sta fr\u00E5gor\u201D."),
      bodyText("Lyft g\u00E4rna 1\u20132 riktigt bra argument som kom upp under seminariet: \u201CDet som [Namn] sa om... var ett starkt argument just f\u00F6r att...\u201D"),

      heading3("L\u00E4rpar-\u00E5tergivning (70\u201380 min)"),
      bodyText("Denna \u00F6vning testar om eleverna aktivt lyssnade p\u00E5 andra gruppers argument \u2014 inte bara sin egen position."),
      spacer(),
      italicText("S\u00E4g: \u201CNu ska vi testa om ni verkligen lyssnade p\u00E5 varandra. Jag parar ihop er med n\u00E5gon fr\u00E5n en ANNAN grupp.\u201D"),
      spacer(),
      bodyText("S\u00E5 h\u00E4r fungerar det:"),
      bullet("Steg 1: Para ihop elever fr\u00E5n OLIKA grupper (r\u00E4cker med att peka: \u201CDu och du\u201D)"),
      bullet("Steg 2: Person A \u00E5terger vad Person B:s grupp argumenterade f\u00F6r (2 minuter). Person A f\u00E5r INTE \u00E5terge sin egen grupps argument \u2014 bara den andras."),
      bullet("Steg 3: Person B lyssnar, r\u00E4ttar och kompletterar (1 minut). \u201CJa, men vi sa ocks\u00E5 att...\u201D eller \u201CDet st\u00E4mmer inte riktigt \u2014 v\u00E5r po\u00E4ng var snarare att...\u201D"),
      bullet("Steg 4: Byt roller. Person B \u00E5terger Person A:s grupps argument (2 min). Person A r\u00E4ttar/kompletterar (1 min)."),
      spacer(),
      boldBodyText("Varf\u00F6r denna \u00F6vning? ", "Den avsl\u00F6jar om eleverna lyssnade aktivt p\u00E5 andra perspektiv eller bara v\u00E4ntade p\u00E5 sin tur. Det tr\u00E4nar ocks\u00E5 f\u00F6rm\u00E5gan att \u00E5terge andras resonemang \u2014 en nyckelkompetens f\u00F6r A-niv\u00E5."),
      spacer(),
      bodyText("Under \u00F6vningen: Cirkulera och lyssna. Notera elever som kan \u00E5terge andra gruppers argument korrekt och nyanserat \u2014 det \u00E4r ett starkt tecken p\u00E5 h\u00F6g f\u00F6rm\u00E5ga."),

      heading3("Summering (80\u201395 min)"),
      bodyText("Helklassreflektion (5 min):"),
      italicText("\u201CVad var det starkaste argumentet ni h\u00F6rde fr\u00E5n en ANNAN grupp \u00E4n er egen?\u201D"),
      bodyText("Samla 3\u20134 svar. Lyft att f\u00F6rm\u00E5gan att se styrkan i andras argument \u00E4r central f\u00F6r nyanserade resonemang."),
      spacer(),
      bodyText("Koppla tillbaka till l\u00E4randem\u00E5let p\u00E5 tavlan."),
      spacer(),
      boldBodyText("Exit ticket (10 min):", ""),
      italicText("\u201CV\u00E4lj en av seminariefr\u00E5gorna och skriv ditt eget svar. Anv\u00E4nd minst tv\u00E5 begrepp fr\u00E5n momentet och h\u00E4nvisa till minst ett argument fr\u00E5n seminariet som du h\u00F6ll med om ELLER ville ifr\u00E5gas\u00E4tta.\u201D"),
      spacer(),
      bodyText("OBS: Inga meningsstartare. Detta \u00E4r en medveten niv\u00E5h\u00F6jning j\u00E4mf\u00F6rt med lektion 1\u20133. Eleverna ska nu kunna formulera sig sj\u00E4lvst\u00E4ndigt."),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r att bed\u00F6ma var eleverna befinner sig inf\u00F6r debatten i lektion 5."),

      heading3("Fram\u00E5tkoppling (95\u2013100 min)"),
      italicText("\u201CIdag hade ni f\u00F6rberedda positioner \u2014 ni fick argumentera f\u00F6r det ni sj\u00E4lva tror p\u00E5. N\u00E4sta g\u00E5ng h\u00F6jer vi insatsen: ni f\u00E5r tilldelade positioner som ni kanske INTE h\u00E5ller med om. D\u00E5 testas er f\u00F6rm\u00E5ga att verkligen f\u00F6rst\u00E5 och argumentera fr\u00E5n flera perspektiv. Fundera till dess: Vilken position hade du sv\u00E5rast att f\u00F6rst\u00E5 idag?\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: snabbskrivning med begrepp fr\u00E5n lektion 1\u20133 (enskilt, 4 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Gruppf\u00F6rberedelse: formulera \u00F6ppningsanf\u00F6rande och argument (grupp 4\u20136, 15 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Seminarium: \u00F6ppningsanf\u00F6randen + fri diskussion + syntes (helklass, 45 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "L\u00E4rpar-\u00E5tergivning: \u00E5terge andras argument (par, 10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: v\u00E4lj seminarifr\u00E5ga + eget resonemang utan meningsstartare (enskilt, 10 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 84 av 100 minuter (84%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Under gruppf\u00F6rberedelsen f\u00E5r dessa elever riktad l\u00E4rarst\u00F6d: \u201CVad s\u00E4ger ert positionskort om det h\u00E4r?\u201D Under seminariet anv\u00E4nds namngivna inbjudningar: \u201C[Namn], din grupp hade en intressant po\u00E4ng om X.\u201D L\u00E4rpar-\u00E5tergivningen ger extra \u00F6vning i att formulera resonemang muntligt. Exit ticket bed\u00F6ms med l\u00E4gre krav p\u00E5 begreppsanv\u00E4ndning."),
      boldBodyText("Utmaning (mot A): ", "Under seminariet uppmuntras dessa elever att problematisera: \u201CKan du se svagheter i ditt eget argument?\u201D eller \u201CVad h\u00E4nder om vi kombinerar tv\u00E5 positioner?\u201D L\u00E4rpar-\u00E5tergivningen ut\u00F6kas: \u201C\u00C5terge argumentet OCH f\u00F6rklara varf\u00F6r du h\u00E5ller med eller inte.\u201D Exit ticket utan ramar: eleven v\u00E4ljer fritt, resonerar nyanserat och kopplar till demokrati p\u00E5 samh\u00E4llsniv\u00E5."),

      // MATERIAL
      heading2("Material"),
      bullet("Exit ticket-data fr\u00E5n lektion 3 (f\u00F6r retrieval practice)"),
      bullet("Seminarieunderlag (bilaga 1) \u2014 en per elev, utdelat i f\u00F6rv\u00E4g"),
      bullet("Bed\u00F6mningsmatris/observationsschema (bilaga 2) \u2014 en per elev f\u00F6r l\u00E4raren"),
      bullet("Positionskort (bilaga 3) \u2014 ett per grupp, utdelat i f\u00F6rv\u00E4g"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),
      bullet("Stolar i ring/hestesko"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Retrieval practice repeterar och f\u00F6rst\u00E4rker samtliga l\u00E4randem\u00E5l fr\u00E5n lektion 1\u20133"),
      bullet("Seminariet tr\u00E4nar m\u00E5l 1 (granska/v\u00E4rdera k\u00E4llor i debattkontext), m\u00E5l 2 (resonera om AI och demokrati) och m\u00E5l 3 (analysera konspirationsteoriers samh\u00E4llsp\u00E5verkan)"),
      bullet("L\u00E4rpar-\u00E5tergivningen tr\u00E4nar f\u00F6rm\u00E5gan att \u00E5terge och analysera andras resonemang \u2014 centralt f\u00F6r C/A-niv\u00E5"),
      bullet("Exit ticket utan meningsstartare m\u00E4ter sj\u00E4lvst\u00E4ndig analysf\u00F6rm\u00E5ga (alla tre m\u00E5l)"),
      spacer(),
      bodyText("Minskad scaffolding \u00E4r en medveten progression. Elever som i lektion 1\u20133 visat E-niv\u00E5 med st\u00F6d f\u00E5r nu m\u00F6jlighet att visa samma niv\u00E5 med mindre st\u00F6d. Elever p\u00E5 C/A-niv\u00E5 f\u00E5r utrymme f\u00F6r sj\u00E4lvst\u00E4ndigt, nyanserat resonemang i seminarieformen."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1 (grundfr\u00E5gor), lektion 2 (konspirationsteorier), lektion 3 (AI + konspirationsteorier)"),
      bullet("V\u00C4NDPUNKT: Fr\u00E5n verktygsbygge till till\u00E4mpning \u2014 f\u00F6rsta lektionen med minskad scaffolding"),
      bullet("N\u00E4sta lektion: \u201CPerspektiven krockar\u201D \u2014 debatt med TILLDELADE positioner (ej sj\u00E4lvvalda)"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r att bed\u00F6ma elevers f\u00F6rm\u00E5geniv\u00E5 inf\u00F6r debatt i lektion 5"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: SEMINARIEUNDERLAG
      heading2("Bilaga 1: Seminarieunderlag \u2014 \u201CInformationskriget\u201D"),
      bodyText("L\u00E4s detta underlag som f\u00F6rberedelse inf\u00F6r seminariet. Du beh\u00F6ver kunna h\u00E4nvisa till inneh\u00E5llet under diskussionen."),
      spacer(),

      heading3("Text 1: Super\u00E5ret 2024 \u2014 AI-desinformation i demokratiska val"),
      bodyText("2024 kallades \u201Csuper\u00E5ret\u201D f\u00F6r val \u2014 n\u00E4stan tv\u00E5 miljarder m\u00E4nniskor gick till valurnorna i \u00F6ver 70 l\u00E4nder. Inget tidigare val\u00E5r har haft s\u00E5 mycket AI-genererat inneh\u00E5ll i omlopp. H\u00E4r \u00E4r n\u00E5gra fall:"),
      spacer(),
      boldBodyText("USA: ", "Inf\u00F6r prim\u00E4rvalet i New Hampshire i januari 2024 ringde robotsamtal med en AI-genererad kopia av president Bidens r\u00F6st till v\u00E4ljare och uppmanade dem att INTE r\u00F6sta. Samtalen n\u00E5dde tusentals v\u00E4ljare innan de stoppades. Mannen bakom samtalen d\u00F6mdes till b\u00F6ter och samh\u00E4llstj\u00E4nst."),
      boldBodyText("Rum\u00E4nien: ", "I presidentvalet 2024 ogiltigf\u00F6rklarades valresultatet efter att bevis uppt\u00E4cktes f\u00F6r AI-driven p\u00E5verkansoperation med manipulerade videor som spreds via TikTok. Det \u00E4r f\u00F6rsta g\u00E5ngen ett demokratiskt val i Europa ogiltigf\u00F6rklarades p\u00E5 grund av AI-desinformation."),
      boldBodyText("Indien: ", "Under parlaments- och premiärministervalet 2024 spreds deepfake-videor p\u00E5 WhatsApp och YouTube d\u00E4r k\u00E4ndisar verkade kritisera premi\u00E4rminister Modi och st\u00F6dja oppositionen."),
      boldBodyText("Indonesien: ", "I mars 2025 spreds en fejkad deepfake-video av president Prabowo Subianto via minst 22 olika TikTok-konton."),
      spacer(),
      bodyText("K\u00E4llor: Harvard Ash Center (2024), Rest of World (2025), Recorded Future (2024)."),

      spacer(),
      heading3("Text 2: EU:s AI-f\u00F6rordning \u2014 reglering av en ny teknologi"),
      bodyText("EU:s AI Act (AI-f\u00F6rordningen) tr\u00E4dde i kraft den 1 augusti 2024 och \u00E4r v\u00E4rldens f\u00F6rsta helt\u00E4ckande AI-lagstiftning. N\u00E5gra viktiga delar:"),
      spacer(),
      bullet("Transparenskrav: Den som skapar AI-genererat inneh\u00E5ll m\u00E5ste m\u00E4rka det s\u00E5 att mottagaren vet att det \u00E4r maskinskapat."),
      bullet("Deepfake-regler: S\u00E4rskilda regler f\u00F6r syntetiskt ljud, bild och video \u2014 obligatorisk m\u00E4rkning."),
      bullet("F\u00F6rbjudna AI-system: AI som manipulerar m\u00E4nniskors beteende genom undermedvetna tekniker \u00E4r f\u00F6rbjudet."),
      bullet("Fullst\u00E4ndig till\u00E4mpning fr\u00E5n 2 augusti 2026."),
      spacer(),
      bodyText("Kritiker menar att lagen \u00E4r f\u00F6r sv\u00E5r att genomdriva i praktiken: Hur sp\u00E5rar man vem som skapat en AI-genererad bild? Och \u00E4r m\u00E4rkning tillr\u00E4cklig n\u00E4r inneh\u00E5llet redan har spridits viralt?"),
      bodyText("F\u00F6respr\u00E5kare h\u00E4vdar att reglering \u00E4r n\u00F6dv\u00E4ndig f\u00F6r att skydda demokratiska processer och att EU visar v\u00E4gen f\u00F6r resten av v\u00E4rlden."),
      spacer(),
      bodyText("K\u00E4llor: EU-kommissionen (2024), Columbia Law CJEL (2024)."),

      spacer(),
      heading3("Text 3: Sverige \u2014 AI-desinformation som demokratihot"),
      bodyText("I Sverige v\u00E4xer oron f\u00F6r AI-driven desinformation inf\u00F6r framtida val. En riksdagsmotion fr\u00E5n 2024 (Motion 2024/25:768) lyfter AI som ett hot mot demokratins grunder och f\u00F6resl\u00E5r \u00E5tg\u00E4rder f\u00F6r att skydda det demokratiska samtalet."),
      spacer(),
      bodyText("Mediemyndigheten fick 2024 i uppdrag att st\u00E4rka befolkningens medie- och informationskunnighet, med s\u00E4rskilt fokus p\u00E5 AI. N\u00E4stan h\u00E4lften av svenska befolkningen (49%) \u00E4r mycket oroade \u00F6ver att AI-genererad desinformation kan p\u00E5verka demokratiska val. Bara en av fyra tror att de sj\u00E4lva kan k\u00E4nna igen AI-desinformation."),
      spacer(),
      bodyText("FOI (Totalf\u00F6rsvarets forskningsinstitut) unders\u00F6kte desinformation i samband med EU-valet 2024 och fann att AI-genererat inneh\u00E5ll anv\u00E4ndes i p\u00E5verkanskampanjer riktade mot europeiska v\u00E4ljare."),
      spacer(),
      bodyText("K\u00E4llor: Sveriges riksdag (2024), Mediemyndigheten (2025), FOI (2024)."),

      spacer(),
      heading3("Text 4: F\u00F6r och emot \u2014 tv\u00E5 perspektiv p\u00E5 AI och information"),
      spacer(),
      boldBodyText("Perspektiv A \u2014 \u201CAI \u00E4r det st\u00F6rsta hotet mot demokratin sedan propaganda\u201D", ""),
      bodyText("AI g\u00F6r det m\u00F6jligt att producera desinformation i industriell skala till n\u00E4stan noll kostnad. Stater som Ryssland, Iran och Venezuela experimenterar redan med AI f\u00F6r att p\u00E5verka andra l\u00E4nders val. N\u00E4r medborgare inte l\u00E4ngre kan lita p\u00E5 vad de ser och h\u00F6r undermineras sj\u00E4lva grunden f\u00F6r demokratiskt beslutsfattande. K\u00E4llkritik r\u00E4cker inte \u2014 det kr\u00E4vs reglering, tekniska l\u00F6sningar och massiv utbildningsinsats."),
      spacer(),
      boldBodyText("Perspektiv B \u2014 \u201CDesinformation \u00E4r inget nytt \u2014 och AI \u00E4r \u00F6verdramatiserat\u201D", ""),
      bodyText("Propaganda och desinformation har alltid f\u00F6ljt med ny kommunikationsteknik: boktryckarkonsten, radion, tv:n, internet. 2024 visade sig AI-hotet vara mindre \u00E4n bef\u00E4rat \u2014 billiga fejkningar (\u201Ccheap fakes\u201D) anv\u00E4ndes sju g\u00E5nger oftare \u00E4n sofistikerade deepfakes. Det verkliga problemet \u00E4r inte AI, utan bristande mediekunnighet och politisk polarisering. \u00D6verdriven reglering riskerar att begr\u00E4nsa yttrandefriheten och innovation."),

      // BILAGA 2: BEDÖMNINGSMATRIS
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Bed\u00F6mningsmatris / Observationsschema"),
      bodyText("Anv\u00E4nd detta schema under seminariet och l\u00E4rpar-\u00E5tergivningen. Markera med kryss eller kort anteckning."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1900, 1900, 1900, 1926],
        rows: [
          new TableRow({
            children: [
              headerCell("Elev", 1400),
              headerCell("Bidrar i seminariet (ja/nej + kvalitet)", 1900),
              headerCell("Anv\u00E4nder begrepp fr\u00E5n momentet", 1900),
              headerCell("Reagerar p\u00E5 andras argument", 1900),
              headerCell("Bed\u00F6mning E/C/A", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1400),
              cell("", 1900),
              cell("", 1900),
              cell("", 1900),
              cell("", 1926),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Bed\u00F6mningsindikatorer"),
      spacer(),
      boldBodyText("E-niv\u00E5: ", "Eleven bidrar med enkla argument och omd\u00F6men. Anv\u00E4nder n\u00E5got begrepp fr\u00E5n momentet. Kan \u00E5terge en annan grupps huvudargument i l\u00E4rpar-\u00F6vningen."),
      boldBodyText("C-niv\u00E5: ", "Eleven f\u00F6r v\u00E4lgrundade resonemang med tydliga kopplingar mellan AI, desinformation och demokrati. Anv\u00E4nder begrepp korrekt. Reagerar p\u00E5 andras argument och utvecklar egna. \u00C5terger andras argument korrekt i l\u00E4rpar-\u00F6vningen."),
      boldBodyText("A-niv\u00E5: ", "Eleven f\u00F6r v\u00E4lgrundade och nyanserade resonemang. Problematiserar och ser sp\u00E4nningar mellan positioner. Kopplar till demokrati p\u00E5 samh\u00E4llsniv\u00E5. Kan ifr\u00E5gas\u00E4tta \u00E4ven sin egen grupps argument. \u00C5terger andras argument nyanserat OCH v\u00E4rderar dem i l\u00E4rpar-\u00F6vningen."),

      spacer(),
      heading3("S\u00E4rskilt att observera under l\u00E4rpar-\u00E5tergivningen"),
      bullet("Kan eleven \u00E5terge en annan grupps argument utan att f\u00F6rvr\u00E4nga det?"),
      bullet("Kan eleven skilja p\u00E5 vad den andra gruppen sa och vad eleven sj\u00E4lv tycker?"),
      bullet("Kompletterar eleven med detaljer eller nyanser?"),

      // BILAGA 3: POSITIONSKORT
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Positionskort f\u00F6r seminariet"),
      bodyText("Varje grupp f\u00E5r ETT positionskort. L\u00E4s igenom det tillsammans med seminarieunderlaget. Er uppgift \u00E4r att f\u00F6rsvara denna position under seminariet."),
      spacer(),

      heading3("Position 1: \u201CRegleringskravet\u201D"),
      boldBodyText("Er st\u00E5ndpunkt: ", "AI-genererat inneh\u00E5ll m\u00E5ste regleras h\u00E5rt av staten f\u00F6r att skydda demokratin."),
      spacer(),
      bodyText("Argument att utg\u00E5 ifr\u00E5n:"),
      bullet("EU:s AI Act visar att reglering \u00E4r m\u00F6jlig \u2014 transparenskrav, m\u00E4rkning, f\u00F6rbud mot manipulation"),
      bullet("Rum\u00E4nien 2024 visar vad som h\u00E4nder utan reglering \u2014 ett helt val ogiltigf\u00F6rklarades"),
      bullet("Teknikf\u00F6retagen har visat att de inte kan reglera sig sj\u00E4lva"),
      bullet("Demokratin kr\u00E4ver att medborgare kan lita p\u00E5 information \u2014 utan reglering f\u00F6rsvinner den tilliten"),
      spacer(),
      bodyText("T\u00E4nk \u00E4ven p\u00E5: Var g\u00E5r gr\u00E4nsen f\u00F6r reglering innan den blir censur?"),

      spacer(),
      heading3("Position 2: \u201CUtbildningslinjen\u201D"),
      boldBodyText("Er st\u00E5ndpunkt: ", "L\u00F6sningen \u00E4r inte reglering utan utbildning \u2014 medborgarna m\u00E5ste l\u00E4ra sig k\u00E4llkritik f\u00F6r AI-\u00E5ldern."),
      spacer(),
      bodyText("Argument att utg\u00E5 ifr\u00E5n:"),
      bullet("Reglering kan inte h\u00E5lla jämna steg med tekniken \u2014 n\u00E4r lagen \u00E4r p\u00E5 plats har AI redan utvecklats vidare"),
      bullet("Mediemyndigheten har f\u00E5tt uppdrag att st\u00E4rka medie- och informationskunnighet \u2014 det \u00E4r r\u00E4tt v\u00E4g"),
      bullet("Historien visar att varje ny teknik (boktryckarkonst, radio, internet) medf\u00F6rde oro som l\u00F6stes genom utbildning"),
      bullet("Reglering riskerar att begr\u00E4nsa yttrandefrihet och innovation"),
      spacer(),
      bodyText("T\u00E4nk \u00E4ven p\u00E5: R\u00E4cker k\u00E4llkritik n\u00E4r AI-inneh\u00E5ll inte g\u00E5r att skilja fr\u00E5n \u00E4kta?"),

      spacer(),
      heading3("Position 3: \u201CTeknikoptimisterna\u201D"),
      boldBodyText("Er st\u00E5ndpunkt: ", "Tekniken som skapar problemet kan ocks\u00E5 l\u00F6sa det \u2014 AI mot AI."),
      spacer(),
      bodyText("Argument att utg\u00E5 ifr\u00E5n:"),
      bullet("AI-verktyg f\u00F6r att uppt\u00E4cka deepfakes och AI-genererat inneh\u00E5ll utvecklas snabbt"),
      bullet("Digitala vattenm\u00E4rken kan g\u00F6ra allt AI-inneh\u00E5ll sp\u00E5rbart"),
      bullet("Plattformar som NewsGuard och K\u00E4llkritikbyr\u00E5n anv\u00E4nder redan AI f\u00F6r faktagranskning"),
      bullet("Reglering och utbildning tar f\u00F6r l\u00E5ng tid \u2014 tekniska l\u00F6sningar kan implementeras direkt"),
      spacer(),
      bodyText("T\u00E4nk \u00E4ven p\u00E5: Kan vi verkligen lita p\u00E5 att teknikf\u00F6retagen l\u00F6ser ett problem de sj\u00E4lva skapat?"),

      spacer(),
      heading3("Position 4: \u201CMedborgarperspektivet\u201D"),
      boldBodyText("Er st\u00E5ndpunkt: ", "Ytterst \u00E4r det medborgarnas eget ansvar att vara kritiska \u2014 varken staten eller f\u00F6retagen kan skydda oss."),
      spacer(),
      bodyText("Argument att utg\u00E5 ifr\u00E5n:"),
      bullet("I en demokrati har varje medborgare ansvar f\u00F6r att vara informerad"),
      bullet("Bara 25% av svenskarna tror att de kan k\u00E4nna igen AI-desinformation \u2014 det visar att vi m\u00E5ste ta eget ansvar"),
      bullet("Att f\u00F6rlita sig p\u00E5 staten eller teknikf\u00F6retag skapar beroendestrukturer"),
      bullet("Konspirationsteorier sprids f\u00F6r att m\u00E4nniskor VILL tro p\u00E5 dem \u2014 det \u00E4r ett individuellt val"),
      spacer(),
      bodyText("T\u00E4nk \u00E4ven p\u00E5: \u00C4r det r\u00E4ttvist att l\u00E4gga ansvaret p\u00E5 individen n\u00E4r problemen \u00E4r strukturella?"),

      spacer(),
      heading3("Position 5 (valfri, vid 5 grupper): \u201CDen kritiska r\u00F6sten\u201D"),
      boldBodyText("Er st\u00E5ndpunkt: ", "Oron f\u00F6r AI-desinformation \u00E4r \u00F6verdriven och riskerar att anv\u00E4ndas f\u00F6r att motivera censur."),
      spacer(),
      bodyText("Argument att utg\u00E5 ifr\u00E5n:"),
      bullet("2024 visade att \u201Ccheap fakes\u201D anv\u00E4ndes 7 g\u00E5nger oftare \u00E4n AI-deepfakes \u2014 hotet \u00E4r \u00F6verdrivet"),
      bullet("R\u00E4dslan f\u00F6r AI-desinformation kan anv\u00E4ndas av stater f\u00F6r att motivera censur och \u00F6vervakning"),
      bullet("Historiskt har varje ny teknik (radio, TV, internet) m\u00F6tts av samma panik \u2014 samh\u00E4llet har alltid anpassat sig"),
      bullet("Det verkliga hotet \u00E4r politisk polarisering och bristande f\u00F6rtroende \u2014 inte tekniken i sig"),
      spacer(),
      bodyText("T\u00E4nk \u00E4ven p\u00E5: \u00C4r det ansvarsl\u00F6st att tona ner riskerna? Vad h\u00E4nder om ni har fel?"),

      // BILAGA 4: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: Exit ticket \u2014 Lektion 4"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "V\u00E4lj en av seminariefr\u00E5gorna nedan och skriv ditt eget svar. Anv\u00E4nd minst tv\u00E5 begrepp fr\u00E5n momentet och h\u00E4nvisa till minst ett argument fr\u00E5n seminariet."),
      spacer(),
      bodyText("Seminariefr\u00E5gor:"),
      bullet("N\u00E4r AI kan generera inneh\u00E5ll som inte g\u00E5r att skilja fr\u00E5n \u00E4kta \u2014 vad h\u00E4nder med v\u00E5r f\u00F6rm\u00E5ga att fatta informerade demokratiska beslut?"),
      bullet("Borde staten reglera AI-genererat inneh\u00E5ll h\u00E5rdare, eller \u00E4r det ett hot mot yttrandefriheten?"),
      bullet("Vem b\u00E4r ansvaret n\u00E4r AI-genererad desinformation p\u00E5verkar ett val?"),
      bullet("Kan k\u00E4llkritik r\u00E4cka som f\u00F6rsvar mot AI-desinformation, eller beh\u00F6vs det n\u00E5got mer?"),
      spacer(),
      bodyText("OBS: Inga meningsstartare. Formulera dig fritt."),

      // BILAGA 5: KÄLLMATERIAL OCH LÄNKAR
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 5: K\u00E4llmaterial och l\u00E4nkar"),
      bodyText("Samlade k\u00E4llor f\u00F6r seminarieunderlag, l\u00E4rarunderlag och f\u00F6rdjupning."),

      heading3("AI och demokrati \u2014 svenska k\u00E4llor"),
      boldBodyText("Riksdagsmotion om AI och demokrati: ", "riksdagen.se/sv/dokument-och-lagar/dokument/motion/ai-och-demokratins-grunder_hc02768/"),
      boldBodyText("Riksdagsfr\u00E5ga om AI-driven desinformation: ", "riksdagen.se/sv/dokument-och-lagar/dokument/skriftlig-fraga/ai-driven-politisk-desinformation_hc111274/"),
      boldBodyText("FOI \u2014 desinformation vid EU-valet 2024: ", "foi.se/rest-api/report/FOI%20Memo%208735"),
      boldBodyText("Mediemyndigheten \u2014 Medieutblick 2025: ", "mediemyndigheten.se/globalassets/rapporter-och-analyser/2025/medieutblick-2025---for-ett-oppet-och-tryggt-medielandskap.pdf"),
      boldBodyText("Dagens PS \u2014 AI-fejkar och svenska val: ", "dagensps.se/varlden/sa-kan-ai-fejkar-manipulera-svenska-valet/"),
      boldBodyText("SVT \u2014 Super\u00E5ret 2024: ", "svt.se/nyheter/utrikes/nu-inleds-det-historiska-supervalaret-2024"),

      heading3("AI och val \u2014 internationella k\u00E4llor"),
      boldBodyText("Harvard Ash Center \u2014 AI and the 2024 Elections: ", "ash.harvard.edu/articles/ai-and-the-2024-elections/"),
      boldBodyText("Harvard Ash Center \u2014 \u201CThe apocalypse that wasn\u2019t\u201D: ", "ash.harvard.edu/articles/the-apocalypse-that-wasnt-ai-was-everywhere-in-2024s-elections-but-deepfakes-and-misinformation-were-only-part-of-the-picture/"),
      boldBodyText("Rest of World \u2014 How AI shaped global elections: ", "restofworld.org/2025/global-elections-ai-use/"),
      boldBodyText("Recorded Future \u2014 2024 Deepfakes Report: ", "recordedfuture.com/research/targets-objectives-emerging-tactics-political-deepfakes"),
      boldBodyText("Brookings \u2014 AI and disinformation in elections: ", "brookings.edu/articles/how-do-artificial-intelligence-and-disinformation-impact-elections/"),
      boldBodyText("Carnegie \u2014 Can Democracy Survive AI?: ", "carnegieendowment.org/research/2024/12/can-democracy-survive-the-disruptive-power-of-ai"),
      boldBodyText("WEF \u2014 AI and election year 2024: ", "weforum.org/stories/2024/01/ai-democracy-election-year-2024-disinformation-misinformation/"),
      boldBodyText("Munich Security Conference \u2014 AI-pocalypse Now?: ", "securityconference.org/en/publications/analyses/ai-pocalypse-disinformation-super-election-year/"),
      boldBodyText("Knight Columbia \u2014 78 Election Deepfakes: ", "knightcolumbia.org/blog/we-looked-at-78-election-deepfakes-political-misinformation-is-not-an-ai-problem"),

      heading3("EU:s AI-f\u00F6rordning och reglering"),
      boldBodyText("EU AI Act \u2014 officiell information: ", "artificial-intelligence-act.com/"),
      boldBodyText("Columbia Law \u2014 Deepfakes and the AI Act: ", "cjel.law.columbia.edu/preliminary-reference/2024/deepfake-deep-trouble-the-european-ai-act-and-the-fight-against-ai-generated-misinformation/"),
      boldBodyText("EU-kommissionen \u2014 Countering information manipulation: ", "commission.europa.eu/topics/countering-information-manipulation_en"),
      boldBodyText("TechPolicy.Press \u2014 EU Code of Practice for Deepfakes: ", "techpolicy.press/what-the-eus-new-ai-code-of-practice-means-for-labeling-deepfakes/"),

      heading3("Deepfakes och specifika fall"),
      boldBodyText("Turing Institute \u2014 Deepfakes and Election Security 2025: ", "cetas.turing.ac.uk/publications/deepfake-scams-poisoned-chatbots"),
      boldBodyText("CIVICUS \u2014 Future-Proofing Elections Against Deepfakes: ", "civicus.org/documents/ddi/final-deepfakes-elections-report-civicus-ddi-research-2025.pdf"),
      boldBodyText("CIGI \u2014 AI Electoral Interference 2025: ", "cigionline.org/articles/then-and-now-how-does-ai-electoral-interference-compare-in-2025/"),
      boldBodyText("Global Taiwan Institute \u2014 Deepfakes i Taiwan: ", "globaltaiwan.org/2025/05/the-malicious-exploitation-of-deepfake-technology/"),

      heading3("K\u00E4llkritik och medie-/informationskunnighet"),
      boldBodyText("K\u00E4llkritikbyr\u00E5n: ", "kallkritikbyran.se"),
      boldBodyText("Internetkunskap \u2014 k\u00E4llkritik: ", "internetkunskap.se/kallkritik/"),
      boldBodyText("Krisinformation.se \u2014 k\u00E4llkritik: ", "krisinformation.se/detta-gor-samhallet/kallkritik/"),
      boldBodyText("MPF \u2014 AI och informationsp\u00E5verkan: ", "mpf.se/kunskap-och-stod/temabibliotek/ai-och-informationspaverkan"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-4.docx", buffer);
  console.log("lektion-4.docx skapad!");
});
