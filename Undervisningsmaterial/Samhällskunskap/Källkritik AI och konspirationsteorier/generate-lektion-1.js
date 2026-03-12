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
        children: [new TextRun({ text: "Lektion 1: Kan du luras av AI?", font: "Arial" })],
      }),
      bodyText("AI-labb och inokulering \u2014 s\u00E5 skapar du \u00F6vertygande desinformation"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "80 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska f\u00E5 en f\u00F6rsta upplevelse av hur l\u00E4tt det \u00E4r att skapa \u00F6vertygande AI-genererat inneh\u00E5ll, och d\u00E4rigenom bygga motst\u00E5ndskraft (inokulering) mot desinformation. Eleven introduceras till tredjepersons-framing som momentets analytiska norm."),
      spacer(),
      bullet("Anv\u00E4nda AI-verktyg f\u00F6r att skapa \u00F6vertygande inneh\u00E5ll och d\u00E4refter analysera varf\u00F6r det upplevs som trov\u00E4rdigt (m\u00E5l 1)"),
      bullet("Resonera om varf\u00F6r AI-genererat inneh\u00E5ll kan vara sv\u00E5rt att skilja fr\u00E5n \u00E4kta inneh\u00E5ll (m\u00E5l 2)"),
      spacer(),
      boldBodyText("E: ", "Eleven skapar AI-genererat inneh\u00E5ll och n\u00E4mner n\u00E5gon faktor som g\u00F6r det \u00F6vertygande. Eleven anv\u00E4nder tredjepersons-framing i n\u00E5gon m\u00E5n."),
      boldBodyText("C: ", "Eleven skapar AI-genererat inneh\u00E5ll och f\u00F6rklarar med v\u00E4lgrundade resonemang varf\u00F6r det upplevs som trov\u00E4rdigt. Eleven anv\u00E4nder tredjepersons-framing konsekvent."),
      boldBodyText("A: ", "Eleven skapar AI-genererat inneh\u00E5ll och analyserar med v\u00E4lgrundade och nyanserade resonemang vilka mekanismer som g\u00F6r det \u00F6vertygande, ur flera perspektiv. Eleven reflekterar \u00F6ver varf\u00F6r tredjepersons-framing \u00E4r viktig."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("Hur propaganda, desinformation och konspirationsteorier fungerar som p\u00E5verkansformer i det digitala medielandskapet."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("S\u00E4kerst\u00E4ll att alla elever har tillg\u00E5ng till ChatGPT (gratis version) eller Copilot (inbyggt i Windows)"),
      bullet("Testa sj\u00E4lv att AI-verktyget \u00E4r tillg\u00E4ngligt fr\u00E5n skolans n\u00E4tverk"),
      bullet("F\u00F6rbered en egen demo \u2014 visa hur du skapar ett \u00F6vertygande socialt medie-inl\u00E4gg med AI (se instruktion under fas 3)"),
      bullet("Skriv ut bilaga 2: AI-labb instruktioner (en per elev)"),
      bullet("Skriv ut bilaga 3: Exit ticket L1 (en per elev)"),
      bullet("Skriv ut bilaga 1: \u00C5terkommande case (en per grupp eller projicera)"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),
      bullet("Backup-plan: om ChatGPT \u00E4r nere, anv\u00E4nd Copilot (inbyggt i Windows). Om b\u00E5da \u00E4r nere: anv\u00E4nd f\u00F6rskrivna AI-genererade texter som eleverna analyserar ist\u00E4llet f\u00F6r att skapa egna"),

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
              cell("F\u00F6rkunskap", 1300),
              cell("Snabbunders\u00F6kning + think-pair-share", 2200),
              cell("Snabbunders\u00F6kning: \u201CVad \u00E4r en deepfake? Har ni sett n\u00E5gon?\u201D Think-pair-share (3 min). Introduktion av de tre \u00E5terkommande casen: AI-TikTok-konton i Sverige, deepfakes inf\u00F6r valet 2026, DeepSeek.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201311 min", 1000),
              cell("M\u00E5laktivering", 1300),
              cell("Presentera AI-labben + tredjepersons-framing", 2200),
              cell("Presentera dagens uppl\u00E4gg: \u201CNi ska sj\u00E4lva skapa \u00F6vertygande inneh\u00E5ll med AI.\u201D Introducera tredjepersons-framing som momentets analytiska norm: \u201CVi fr\u00E5gar inte vad du tror \u2014 vi analyserar varf\u00F6r n\u00E5gon KAN finna det \u00F6vertygande.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("11\u201323 min", 1000),
              cell("Instruktion", 1300),
              cell("Prebunking/inokulering + AI-demo", 2200),
              cell("Vad \u00E4r prebunking/inokulering? Kort f\u00F6rklaring av \u201Cmentala antikroppar.\u201D Demo: l\u00E4raren visar live hur man skapar ett \u00F6vertygande socialt medie-inl\u00E4gg med ChatGPT om ett av casen. S\u00E4kerhetsramar: allt stannar i klassrummet, analytisk distans.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("23\u201348 min", 1000),
              cell("AI-labb", 1300),
              cell("Eleverna skapar AI-inneh\u00E5ll (25 min)", 2200),
              cell("Eleverna skapar i par ett socialt medie-inl\u00E4gg med AI (ChatGPT/Copilot) som ser trov\u00E4rdigt ut men inneh\u00E5ller vilseledande information om ett av de \u00E5terkommande casen. Tredjepersons-framing: \u201CVarf\u00F6r KAN n\u00E5gon finna detta \u00F6vertygande?\u201D L\u00E4raren cirkulerar och st\u00F6djer.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("48\u201370 min", 1000),
              cell("Reflektion", 1300),
              cell("Granskning + helklassdiskussion (22 min)", 2200),
              cell("Parbyte: granska ett annat pars skapelse. Identifiera: Vad g\u00F6r det \u00F6vertygande? Vilka k\u00E4llkritiska reaktioner BORDE utl\u00F6sas? Helklassdiskussion om m\u00F6nster. Koppling till inokulering: \u201CNu har ni upplevt hur l\u00E4tt det \u00E4r.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("70\u201380 min", 1000),
              cell("Avslut", 1300),
              cell("Exit ticket + preview", 2200),
              cell("Exit ticket: \u201CDu skapade AI-genererat inneh\u00E5ll idag. N\u00E4mn tv\u00E5 saker som gjorde inneh\u00E5llet \u00F6vertygande \u2014 och en sak som en k\u00E4llkritisk granskare borde ha reagerat p\u00E5.\u201D Preview av L2: \u201CN\u00E4sta g\u00E5ng f\u00E5r ni verktygen att granska det ni just skapade.\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Fas 1: F\u00F6rkunskapsaktivering (0\u20138 min)"),
      boldBodyText("Snabbunders\u00F6kning (0\u20133 min): ", ""),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar \u2014 en snabb fr\u00E5ga. R\u00E4ck upp handen om ni vet vad en deepfake \u00E4r. Bra. Nu vill jag att ni skriver ner p\u00E5 en lapp: Vad \u00E4r en deepfake? Har ni sett n\u00E5gon? Ni har 1 minut.\u201D"),
      spacer(),
      boldBodyText("Think-pair-share (3\u20136 min): ", ""),
      italicText("S\u00E4g: \u201CJ\u00E4mf\u00F6r era svar med den som sitter bredvid er. Har ni sett samma typ av inneh\u00E5ll? Ni har 2 minuter.\u201D"),
      bodyText("Samla 2\u20133 svar muntligt. Fyll i med korta f\u00F6rtydliganden vid behov."),
      spacer(),
      boldBodyText("Introduktion av \u00E5terkommande case (6\u20138 min): ", ""),
      italicText("S\u00E4g: \u201CUnder det h\u00E4r momentet kommer vi att f\u00F6lja tre verkliga fall. Dessa kommer att \u00E5terkomma lektion efter lektion, och ni kommer se dem fr\u00E5n nya vinklar varje g\u00E5ng.\u201D"),
      spacer(),
      bodyText("Presentera kort de tre casen (projicera eller dela ut bilaga 1):"),
      bullet("Case 1: AI-genererade TikTok-konton i Sverige (2025) \u2014 fyra konton utgav sig f\u00F6r att vara svenska tjejer men var AI-genererade och spred h\u00F6gerpropaganda. En enskild person l\u00E5g bakom."),
      bullet("Case 2: Deepfakes inf\u00F6r valet 2026 \u2014 manipulerade filmer och bilder av politiska f\u00F6retr\u00E4dare syftar till att skapa misstro och p\u00E5verka opinion. 74% av nyskapat n\u00E4tinneh\u00E5ll 2025 genereras av AI/botar."),
      bullet("Case 3: DeepSeek och konspirationsspridning (2025) \u2014 kinesisk AI-modell som testats p\u00E5 konspirationsteorier visar hur AI kan generera \u00F6vertygande konspirationstexter p\u00E5 n\u00E5gra sekunder."),
      spacer(),
      italicText("\u201CDessa tre fall \u00E4r verkliga. De h\u00E4nder just nu, i Sverige och i v\u00E4rlden. Vi kommer att \u00E5terkomma till dem under hela momentet.\u201D"),

      heading3("Fas 2: M\u00E5laktivering (8\u201311 min)"),
      italicText("S\u00E4g: \u201CIdag ska ni sj\u00E4lva f\u00E5 testa hur l\u00E4tt det \u00E4r att skapa \u00F6vertygande inneh\u00E5ll med AI. Men f\u00F6rst \u2014 en viktig princip som g\u00E4ller under hela det h\u00E4r momentet.\u201D"),
      spacer(),
      boldBodyText("Introducera tredjepersons-framing: ", ""),
      bodyText("Skriv p\u00E5 tavlan:"),
      italicText("\u201CVi fr\u00E5gar inte: Tror du p\u00E5 detta?\u201D"),
      italicText("\u201CVi fr\u00E5gar: Varf\u00F6r KAN n\u00E5gon finna detta \u00F6vertygande?\u201D"),
      spacer(),
      bodyText("F\u00F6rklara varf\u00F6r:"),
      bullet("Det handlar inte om vad vi personligen tror \u2014 det handlar om att f\u00F6rst\u00E5 mekanismerna"),
      bullet("N\u00E4r vi analyserar varf\u00F6r n\u00E5gon KAN tro p\u00E5 n\u00E5got, \u00F6ppnar vi f\u00F6r djupare f\u00F6rst\u00E5else"),
      bullet("Denna analytiska h\u00E5llning g\u00E4ller i alla diskussioner, skrivuppgifter och seminarier under momentet"),
      spacer(),
      italicText("\u201CDetta \u00E4r inte en begr\u00E4nsning \u2014 det \u00E4r ett analytiskt verktyg. Det g\u00F6r att vi kan unders\u00F6ka k\u00E4nsliga \u00E4mnen utan att n\u00E5gon beh\u00F6ver k\u00E4nna sig utpekad.\u201D"),
      spacer(),
      bodyText("Visa l\u00E4randem\u00E5let p\u00E5 tavlan och koppla till AI-labben:"),
      italicText("\u201CIdag ska ni f\u00E5 uppleva hur l\u00E4tt det \u00E4r att skapa \u00F6vertygande inneh\u00E5ll \u2014 och sedan analysera varf\u00F6r det fungerar. N\u00E4sta lektion f\u00E5r ni verktygen att granska det.\u201D"),

      heading3("Fas 3: Explicit instruktion \u2014 Prebunking och AI-demo (11\u201323 min)"),
      boldBodyText("Prebunking/inokulering (11\u201315 min): ", ""),
      italicText("S\u00E4g: \u201CK\u00E4nner ni till vaccin? Hur fungerar de? Jo \u2014 kroppen f\u00E5r en liten dos av viruset s\u00E5 att den bygger antikroppar. Idag g\u00F6r vi samma sak fast f\u00F6r hj\u00E4rnan. Ni ska sj\u00E4lva skapa desinformation \u2014 s\u00E5 att ni bygger mentala antikroppar mot det.\u201D"),
      spacer(),
      bodyText("F\u00F6rklara kort:"),
      bullet("Prebunking = vaccinera mot desinformation INNAN man m\u00F6ter den"),
      bullet("Inokulering = att uppleva hur n\u00E5got fungerar s\u00E5 att man k\u00E4nner igen det"),
      bullet("Forskning visar att det fungerar \u2014 den som sj\u00E4lv har skapat vilseledande inneh\u00E5ll blir b\u00E4ttre p\u00E5 att genomsk\u00E5da det"),
      spacer(),
      boldBodyText("AI-demo (15\u201323 min): ", ""),
      italicText("S\u00E4g: \u201CNu ska jag visa er hur snabbt det g\u00E5r. Jag skapar ett socialt medie-inl\u00E4gg som ser trov\u00E4rdigt ut \u2014 p\u00E5 under en minut.\u201D"),
      spacer(),
      bodyText("Visa live p\u00E5 sk\u00E4rmen:"),
      bullet("G\u00E5 till ChatGPT (eller Copilot)"),
      bullet("Skriv en prompt som: \u201CSkriv ett Instagram-inl\u00E4gg som utger sig f\u00F6r att vara fr\u00E5n en trovärdig nyhetsk\u00E4lla. Inl\u00E4gget ska handla om deepfakes inf\u00F6r valet 2026 och antyda att demokratin \u00E4r hotad utan att ge konkreta bevis.\u201D"),
      bullet("Visa resultatet. L\u00E5t klassen reagera."),
      spacer(),
      italicText("Fr\u00E5ga klassen: \u201CVad g\u00F6r det h\u00E4r inl\u00E4gget trov\u00E4rdigt? Anv\u00E4nd v\u00E5r nya princip \u2014 varf\u00F6r KAN n\u00E5gon finna det h\u00E4r \u00F6vertygande?\u201D"),
      spacer(),
      bodyText("Samla 3\u20134 svar. Skriv nyckelord p\u00E5 tavlan (t.ex. \u201Cser professionellt ut\u201D, \u201Cspelar p\u00E5 k\u00E4nslor\u201D, \u201Cn\u00E4mner verkliga h\u00E4ndelser\u201D)."),
      spacer(),
      boldBodyText("S\u00E4kerhetsramar: ", ""),
      bullet("Allt inneh\u00E5ll ni skapar stannar i klassrummet \u2014 inget sprids"),
      bullet("Vi g\u00F6r detta f\u00F6r att f\u00F6rst\u00E5 mekanismerna, inte f\u00F6r att sprida desinformation"),
      bullet("Anv\u00E4nd alltid tredjepersons-framing: analysera varf\u00F6r det \u00E4r \u00F6vertygande, inte om det \u00E4r sant"),

      heading3("Fas 4: Guidad \u00F6vning \u2014 AI-labb (23\u201348 min)"),
      italicText("S\u00E4g: \u201CNu \u00E4r det er tur. Ni ska i par skapa ETT socialt medie-inl\u00E4gg med AI. Inl\u00E4gget ska se trov\u00E4rdigt ut men inneh\u00E5lla vilseledande information om ett av v\u00E5ra tre case.\u201D"),
      spacer(),
      bodyText("Dela ut bilaga 2 (AI-labb instruktioner) och g\u00E5 igenom stegen:"),
      bullet("Steg 1: V\u00E4lj ett av de tre \u00E5terkommande casen"),
      bullet("Steg 2: Diskutera i paret: vilken vinkel ska vi ta? Vilken plattform imiterar vi?"),
      bullet("Steg 3: Anv\u00E4nd ChatGPT eller Copilot f\u00F6r att generera texten. F\u00F6rb\u00E4ttra med egna \u00E4ndringar."),
      bullet("Steg 4: Skriv under inl\u00E4gget: \u201CVarf\u00F6r KAN n\u00E5gon finna detta \u00F6vertygande?\u201D (minst 2 anledningar)"),
      spacer(),
      boldBodyText("L\u00E4raren cirkulerar: ", ""),
      bullet("Till par som fastnar: \u201CB\u00F6rja med att v\u00E4lja case. Vilken plattform t\u00E4nker ni er att det h\u00E4r publiceras p\u00E5?\u201D"),
      bullet("Till par som \u00E4r klara snabbt: \u201CKan ni g\u00F6ra ert inl\u00E4gg \u00E4nnu mer \u00F6vertygande? Vilka retoriska strategier saknas?\u201D"),
      bullet("Om elev skriver ol\u00E4mpligt inneh\u00E5ll: \u201CKom ih\u00E5g att vi analyserar mekanismer \u2014 inte skapar verklig skada. V\u00E4lj en annan vinkel.\u201D"),
      bullet("Om AI-verktyget \u00E4r nere: \u201CAnv\u00E4nd Copilot ist\u00E4llet (inbyggt i Windows). Om inte heller det fungerar \u2014 skriv inl\u00E4gget f\u00F6r hand och analysera vad som g\u00F6r det \u00F6vertygande.\u201D"),
      spacer(),
      boldBodyText("Differentiering under AI-labben: ", ""),
      boldBodyText("St\u00F6d (mot E): ", "Ge promptf\u00F6rslag: \u201CSkriv ett Instagram-inl\u00E4gg om [case] som ser ut att komma fr\u00E5n en nyhetssida.\u201D Ge meningsstartare f\u00F6r reflektionen: \u201CDet h\u00E4r inl\u00E4gget kan vara \u00F6vertygande f\u00F6r att...\u201D"),
      boldBodyText("Utmaning (mot A): ", "Skapa inneh\u00E5llet utan promptf\u00F6rslag. Analysera dessutom vilka retoriska strategier (patos, etos, logos) som g\u00F6r inneh\u00E5llet \u00F6vertygande. Resonera om samh\u00E4llskonsekvenser: \u201CVad h\u00E4nder om tusentals s\u00E5dana inl\u00E4gg sprids?\u201D"),

      heading3("Fas 5: Reflektion och analys (48\u201370 min)"),
      italicText("S\u00E4g: \u201CStopp! Nu l\u00E4gger vi ner skapandet. Nu byter vi roll \u2014 fr\u00E5n skapare till granskare.\u201D"),
      spacer(),
      boldBodyText("Parbyte och granskning (48\u201358 min): ", ""),
      bodyText("Varje par byter inl\u00E4gg med ett annat par."),
      italicText("Instruktion: \u201CL\u00E4s det andra parets inl\u00E4gg. Diskutera och skriv ner svar p\u00E5 tv\u00E5 fr\u00E5gor: (1) Vad g\u00F6r det h\u00E4r inl\u00E4gget \u00F6vertygande? (2) Vilka k\u00E4llkritiska reaktioner BORDE utl\u00F6sas n\u00E4r man l\u00E4ser det? Ni har 10 minuter.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 skriv p\u00E5 tavlan:"),
      bullet("\u201CDet \u00E4r \u00F6vertygande f\u00F6r att...\u201D"),
      bullet("\u201CEn k\u00E4llkritisk granskare borde reagera p\u00E5 att...\u201D"),
      spacer(),
      boldBodyText("Helklassdiskussion (58\u201370 min): ", ""),
      italicText("S\u00E4g: \u201CL\u00E5t oss samla det vi l\u00E4rt oss. Vilka m\u00F6nster s\u00E5g ni? Vad var det som gjorde inl\u00E4ggen \u00F6vertygande?\u201D"),
      spacer(),
      bodyText("Samla m\u00F6nster p\u00E5 tavlan. Typiska svar:"),
      bullet("Spr\u00E5ket l\u00E5ter professionellt/journalistiskt"),
      bullet("Det refererar till verkliga h\u00E4ndelser och platser"),
      bullet("Det spelar p\u00E5 k\u00E4nslor (r\u00E4dsla, ilska, empati)"),
      bullet("Det ser ut som det kommer fr\u00E5n en trov\u00E4rdig k\u00E4lla"),
      bullet("Det \u00E4r sv\u00E5rt att skilja fr\u00E5n \u00E4kta inneh\u00E5ll"),
      spacer(),
      boldBodyText("Koppla till inokulering: ", ""),
      italicText("\u201CNu har ni upplevt hur l\u00E4tt det \u00E4r att skapa \u00F6vertygande inneh\u00E5ll med AI. Det tog n\u00E5gra minuter. T\u00E4nk p\u00E5 det \u2014 alla med en dator kan g\u00F6ra det h\u00E4r. Men nu vet ni ocks\u00E5 vad ni ska leta efter. Ni har f\u00E5tt era f\u00F6rsta mentala antikroppar.\u201D"),
      spacer(),
      italicText("\u201CN\u00E4sta lektion f\u00E5r ni de riktiga verktygen \u2014 de k\u00E4llkritiska grundfr\u00E5gorna och SIFT-metoden. D\u00E5 kan ni granska det h\u00E4r p\u00E5 riktigt.\u201D"),

      heading3("Fas 6: Exit ticket + preview (70\u201380 min)"),
      boldBodyText("Exit ticket (70\u201377 min): ", ""),
      bodyText("Dela ut bilaga 3 eller skriv p\u00E5 tavlan:"),
      spacer(),
      italicText("\u201CDu skapade AI-genererat inneh\u00E5ll idag. N\u00E4mn tv\u00E5 saker som gjorde inneh\u00E5llet \u00F6vertygande \u2014 och en sak som en k\u00E4llkritisk granskare borde ha reagerat p\u00E5.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 meningsstartare p\u00E5 tavlan:"),
      bullet("\u201CInneh\u00E5llet var \u00F6vertygande f\u00F6r att...\u201D"),
      bullet("\u201CEn annan sak som gjorde det trov\u00E4rdigt var att...\u201D"),
      bullet("\u201CEn k\u00E4llkritisk granskare borde ha reagerat p\u00E5 att...\u201D"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd svaren f\u00F6r retrieval review i lektion 2."),
      spacer(),
      boldBodyText("Preview (77\u201380 min): ", ""),
      italicText("\u201CIdag upplevde ni hur l\u00E4tt det \u00E4r att skapa \u00F6vertygande inneh\u00E5ll. Men ni hade inga verktyg att granska det med. N\u00E4sta g\u00E5ng f\u00E5r ni verktygsl\u00E5dan \u2014 de fyra k\u00E4llkritiska grundfr\u00E5gorna, SIFT-metoden och lateral reading. D\u00E5 kan ni granska det ni sj\u00E4lva skapade idag.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Snabbunders\u00F6kning + think-pair-share: vad \u00E4r en deepfake? (enskilt + par, 6 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "AI-labb: skapa ett socialt medie-inl\u00E4gg med AI i par (par, 25 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Parbyte och granskning: analysera ett annat pars inl\u00E4gg (par, 10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Helklassdiskussion: m\u00F6nster i \u00F6vertygande inneh\u00E5ll (helklass, 12 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: trov\u00E4rdighetsfaktorer + k\u00E4llkritisk reaktion (enskilt, 7 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 60 av 80 minuter (75%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Promptf\u00F6rslag f\u00F6r AI-verktyget (\u201CSkriv ett Instagram-inl\u00E4gg om [case] som ser ut att komma fr\u00E5n en nyhetssida\u201D). Meningsstartare vid reflektion (\u201CDet h\u00E4r inl\u00E4gget kan vara \u00F6vertygande f\u00F6r att...\u201D). Exit ticket med meningsstartare. L\u00E4raren prioriterar dessa par vid cirkulering."),
      boldBodyText("Utmaning (mot A): ", "Skapa inneh\u00E5ll utan promptf\u00F6rslag. Analysera vilka retoriska strategier (patos, etos, logos) som g\u00F6r inneh\u00E5llet \u00F6vertygande. Resonera om samh\u00E4llskonsekvenser: \u201CVad h\u00E4nder om tusentals s\u00E5dana inl\u00E4gg sprids inf\u00F6r ett val?\u201D Exit ticket utan meningsstartare."),

      // MATERIAL
      heading2("Material"),
      bullet("Datorer/surfplattor med tillg\u00E5ng till ChatGPT eller Copilot (en per par)"),
      bullet("Bilaga 1: \u00C5terkommande case \u2014 kort beskrivning av tre case (projicera eller en per grupp)"),
      bullet("Bilaga 2: AI-labb instruktioner \u2014 steg-f\u00F6r-steg f\u00F6r eleverna (en per par)"),
      bullet("Bilaga 3: Exit ticket L1 (en per elev)"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("AI-labben tr\u00E4nar f\u00F6rm\u00E5gan att f\u00F6rst\u00E5 hur digitalt inneh\u00E5ll kan vilseleda (m\u00E5l 1: k\u00E4llkritiska verktyg)"),
      bullet("Reflektionsfasen tr\u00E4nar resonemang om varf\u00F6r AI-genererat inneh\u00E5ll upplevs som \u00F6vertygande (m\u00E5l 2: konspirationsteorianalys)"),
      bullet("Tredjepersons-framing introduceras som analytisk norm f\u00F6r momentet (m\u00E5l 3: diskussion i analytisk form)"),
      bullet("Exit ticket m\u00E4ter f\u00F6rm\u00E5gan att identifiera trov\u00E4rdighetsfaktorer och k\u00E4llkritiska reaktioner (m\u00E5l 1, 2)"),
      spacer(),
      bodyText("AI-labbens differentiering (promptf\u00F6rslag, meningsstartare) s\u00E4kerst\u00E4ller att alla elever kan n\u00E5 E-niv\u00E5. Den \u00F6ppna reflektionsdiskussionen och retorisk analys ger A-elever m\u00F6jlighet att visa nyanserat resonemang fr\u00E5n flera perspektiv."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("F\u00F6rsta lektionen i momentet \u2014 ingen f\u00F6reg\u00E5ende lektion att bygga p\u00E5"),
      bullet("Introducerar de tre \u00E5terkommande casen (AI-TikTok, deepfakes valet 2026, DeepSeek) som f\u00F6ljer med genom hela momentet"),
      bullet("Introducerar tredjepersons-framing som g\u00E4ller i alla \u00E5tta lektioner"),
      bullet("Bygger grund f\u00F6r L2 (k\u00E4llkritiska verktyg \u2014 verktygen som saknades under AI-labben)"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r retrieval review i L2"),
      bullet("AI-labb-upplevelsen \u00E5terkopplas i L6 (skrivuppgift) och L8 (seminarium)"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: ÅTERKOMMANDE CASE
      heading2("Bilaga 1: \u00C5terkommande case"),
      bodyText("Dessa tre fall f\u00F6ljer med genom hela momentet. I varje lektion ser vi dem fr\u00E5n nya vinklar."),
      spacer(),

      boldBodyText("Case 1: AI-genererade TikTok-konton i Sverige (2025)", ""),
      bodyText("Under 2025 uppt\u00E4cktes fyra TikTok-konton som utgav sig f\u00F6r att vara svenska tjejer. Kontona anv\u00E4nde AI-genererade profilbilder och r\u00F6ster, och publicerade inneh\u00E5ll som s\u00E5g ut som vanliga ungdomsinl\u00E4gg \u2014 men som gradvis blandade in h\u00F6gerextrem propaganda. Kontona fick tusentals f\u00F6ljare innan de avsl\u00F6jades. En enskild person erk\u00E4nde att han l\u00E5g bakom alla fyra konton."),
      spacer(),
      bodyText("Centrala fr\u00E5gor:"),
      bullet("Hur kunde AI-genererade profiler passera som \u00E4kta?"),
      bullet("Varf\u00F6r valde skaparen att imitera unga svenska tjejer?"),
      bullet("Vilka k\u00E4llkritiska verktyg hade kunnat avsl\u00F6ja kontona?"),
      spacer(),

      boldBodyText("Case 2: Deepfakes inf\u00F6r valet 2026", ""),
      bodyText("Inf\u00F6r det svenska valet 2026 har manipulerade filmer och bilder av politiska f\u00F6retr\u00E4dare b\u00F6rjat spridas p\u00E5 sociala medier. Deepfake-teknologi g\u00F6r det m\u00F6jligt att skapa realistiska videoklipp d\u00E4r politiker s\u00E4ger saker de aldrig sagt. Enligt rapporter genereras 74% av nyskapat n\u00E4tinneh\u00E5ll 2025 av AI eller botar. Riksdagsmotioner har lagts f\u00F6r att kr\u00E4va m\u00E4rkning av AI-genererat material."),
      spacer(),
      bodyText("Centrala fr\u00E5gor:"),
      bullet("Hur p\u00E5verkas demokratin n\u00E4r v\u00E4ljare inte kan lita p\u00E5 vad de ser?"),
      bullet("Vilka tekniska och samh\u00E4lleliga motmedel finns?"),
      bullet("Hur sprids deepfakes \u2014 och varf\u00F6r \u00E4r de sv\u00E5ra att stoppa?"),
      spacer(),

      boldBodyText("Case 3: DeepSeek och konspirationsspridning (2025)", ""),
      bodyText("DeepSeek \u00E4r en kinesisk AI-modell som under 2025 testades p\u00E5 att generera konspirationsteorier. Testerna visade att AI-modellen kunde producera l\u00E5nga, detaljerade och \u00F6vertygande konspirationstexter p\u00E5 n\u00E5gra sekunder \u2014 inklusive falska k\u00E4llh\u00E4nvisningar och vetenskapligt klingande argumentation. Resultaten v\u00E4ckte debatt om AI-modellers ansvar f\u00F6r desinformation."),
      spacer(),
      bodyText("Centrala fr\u00E5gor:"),
      bullet("Vad skiljer en AI-genererad konspirationsteori fr\u00E5n en m\u00E4nniskoskapad?"),
      bullet("Hur p\u00E5verkas trov\u00E4rdigheten n\u00E4r AI kan skapa falska k\u00E4llh\u00E4nvisningar?"),
      bullet("Vem b\u00E4r ansvaret n\u00E4r AI anv\u00E4nds f\u00F6r att sprida desinformation?"),

      // BILAGA 2: AI-LABB INSTRUKTIONER
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: AI-labb \u2014 Instruktioner"),
      bodyText("F\u00F6lj stegen nedan. Ni arbetar i par."),
      spacer(),

      boldBodyText("S\u00E4kerhetsregler:", ""),
      bullet("Allt inneh\u00E5ll ni skapar stannar i klassrummet \u2014 INGET publiceras eller sprids"),
      bullet("Syftet \u00E4r att f\u00F6rst\u00E5 mekanismerna \u2014 inte att skapa verklig desinformation"),
      bullet("Anv\u00E4nd tredjepersons-framing: analysera varf\u00F6r n\u00E5gon KAN tycka det \u00E4r \u00F6vertygande"),
      spacer(),

      boldBodyText("Steg 1: V\u00E4lj case", ""),
      bodyText("V\u00E4lj ETT av de tre \u00E5terkommande casen:"),
      bullet("AI-genererade TikTok-konton i Sverige"),
      bullet("Deepfakes inf\u00F6r valet 2026"),
      bullet("DeepSeek och konspirationsspridning"),
      spacer(),

      boldBodyText("Steg 2: Planera ert inl\u00E4gg", ""),
      bodyText("Diskutera i paret:"),
      bullet("Vilken plattform imiterar ni? (Instagram, TikTok, Twitter/X, nyhetssida)"),
      bullet("Vilken vinkel tar ni? (Nyhetsartikel, personligt inl\u00E4gg, \u201Cavsl\u00F6jande\u201D, debattinl\u00E4gg)"),
      bullet("Vem \u00E4r avs\u00E4ndaren? (Journalist, expert, vanlig person, anonym)"),
      spacer(),

      boldBodyText("Steg 3: Skapa med AI", ""),
      bodyText("Anv\u00E4nd ChatGPT eller Copilot. Exempel p\u00E5 prompt:"),
      italicText("\u201CSkriv ett kort Instagram-inl\u00E4gg som utger sig f\u00F6r att vara fr\u00E5n en svensk nyhetskanal. Inl\u00E4gget ska handla om [ert case] och antyda att [vinkel] utan att ge k\u00E4llor.\u201D"),
      spacer(),
      bodyText("L\u00E4s igenom resultatet. F\u00F6rb\u00E4ttra med egna \u00E4ndringar om ni vill."),
      spacer(),

      boldBodyText("Steg 4: Reflektera", ""),
      bodyText("Skriv under ert inl\u00E4gg:"),
      italicText("\u201CVarf\u00F6r KAN n\u00E5gon finna detta \u00F6vertygande?\u201D"),
      bodyText("Ge minst tv\u00E5 anledningar."),

      // BILAGA 3: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Exit ticket \u2014 Lektion 1"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "Du skapade AI-genererat inneh\u00E5ll idag. N\u00E4mn tv\u00E5 saker som gjorde inneh\u00E5llet \u00F6vertygande \u2014 och en sak som en k\u00E4llkritisk granskare borde ha reagerat p\u00E5."),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CInneh\u00E5llet var \u00F6vertygande f\u00F6r att...\u201D"),
      bullet("\u201CEn annan sak som gjorde det trov\u00E4rdigt var att...\u201D"),
      bullet("\u201CEn k\u00E4llkritisk granskare borde ha reagerat p\u00E5 att...\u201D"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-1.docx", buffer);
  console.log("lektion-1.docx skapad!");
});
