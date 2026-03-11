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
        children: [new TextRun({ text: "Lektion 1: Kan du lita p\u00E5 det du ser?", font: "Arial" })],
      }),
      bodyText("Introduktion av k\u00E4llkritiska grundfr\u00E5gor och AI-genererat inneh\u00E5ll"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "100 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska kunna granska och v\u00E4rdera information fr\u00E5n olika k\u00E4llor samt resonera om hur AI-genererat inneh\u00E5ll p\u00E5verkar informationslandskapet."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av de fyra k\u00E4llkritiska grundfr\u00E5gorna (m\u00E5l 1)"),
      bullet("Resonera om hur AI-genererat inneh\u00E5ll p\u00E5verkar m\u00F6jligheten att granska information (m\u00E5l 2)"),
      spacer(),
      boldBodyText("E: ", "Eleven granskar k\u00E4llor med enkla omd\u00F6men och f\u00F6r enkla resonemang om AI-genererat inneh\u00E5ll."),
      boldBodyText("C: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade omd\u00F6men och f\u00F6r v\u00E4lgrundade resonemang om hur AI f\u00F6r\u00E4ndrar informationslandskapet."),
      boldBodyText("A: ", "Eleven granskar k\u00E4llor med v\u00E4lgrundade och nyanserade omd\u00F6men och f\u00F6r v\u00E4lgrundade och nyanserade resonemang om AI:s p\u00E5verkan, med kopplingar till demokrati och samh\u00E4llsdebatt."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered \u00C4kta-eller-fejk-bildspel med 8\u201310 exempel: AI-genererade bilder, manipulerade nyhetsartiklar, \u00E4kta nyheter, deepfake-video (se bilaga)"),
      bullet("Skriv ut analysmatrisen \u201CK\u00E4llkritiska grundfr\u00E5gor\u201D (en per elev, se bilaga)"),
      bullet("Skriv ut sorteringskorten \u201C\u00C4kta, AI eller manipulerat?\u201D (en upps\u00E4ttning per grupp om 3\u20134 elever, se bilaga)"),
      bullet("Testa att alla digitala exempel fungerar (video, l\u00E4nkar)"),
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
              cell("\u201CKan du lita p\u00E5 det du ser?\u201D", 2200),
              cell("Visa tre bilder (en \u00E4kta nyhetsbild, en AI-genererad, en manipulerad). Eleverna r\u00F6star: \u00E4kta eller fejk? Kort genomg\u00E5ng av r\u00E4tt svar. Landa i fr\u00E5gan: Hur kan vi avg\u00F6ra vad som \u00E4r sant?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201325 min", 1000),
              cell("Instruktion", 1300),
              cell("De fyra k\u00E4llkritiska grundfr\u00E5gorna + AI-intro", 2200),
              cell("Presentera de fyra grundfr\u00E5gorna: Vem? Varf\u00F6r? Hur? N\u00E4r? Koppla varje fr\u00E5ga till ett konkret AI-exempel. EPA-stopp vid minut 18: \u201CVilken av grundfr\u00E5gorna \u00E4r sv\u00E5rast att besvara n\u00E4r inneh\u00E5llet \u00E4r AI-genererat?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("25\u201360 min", 1000),
              cell("Bearbetning", 1300),
              cell("Sortering + analysmatris", 2200),
              cell("Del 1 (25\u201340 min): Sortera kort i grupper: \u00E4kta / AI-genererat / manipulerat. Del 2 (40\u201360 min): Analysera 2\u20133 av korten med analysmatrisen (de fyra grundfr\u00E5gorna). Helklassgenomg\u00E5ng av sv\u00E5ra fall.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("60\u201380 min", 1000),
              cell("F\u00F6rdjupning", 1300),
              cell("AI-labb: \u201CKan du lura klassen?\u201D", 2200),
              cell("Eleverna f\u00E5r i par anv\u00E4nda ett AI-verktyg f\u00F6r att skapa en kort text eller bild. M\u00E5let: skapa n\u00E5got som ser \u00E4kta ut. Sedan granskar ett annat par deras produkt med grundfr\u00E5gorna.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("80\u201395 min", 1000),
              cell("Summering", 1300),
              cell("Gemensam reflektion + exit ticket", 2200),
              cell("Helklassdiskussion: Vad var sv\u00E5rast att avsl\u00F6ja? Varf\u00F6r? Koppla till l\u00E4randem\u00E5let. Exit ticket: \u201CNamnge de fyra k\u00E4llkritiska grundfr\u00E5gorna och f\u00F6rklara varf\u00F6r EN av dem blir extra viktig n\u00E4r inneh\u00E5ll \u00E4r AI-genererat.\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("95\u2013100 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("N\u00E4sta lektion", 2200),
              cell("\u201CNu har vi verktyg f\u00F6r att granska k\u00E4llor. N\u00E4sta g\u00E5ng tittar vi p\u00E5 n\u00E5got som g\u00F6r det \u00E4nnu sv\u00E5rare \u2014 konspirationsteorier. Vem tj\u00E4nar p\u00E5 att du tror p\u00E5 dem?\u201D", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart (0\u20138 min)"),
      italicText("S\u00E4g: \u201CInnan vi b\u00F6rjar \u2014 jag ska visa er tre bilder. Er uppgift \u00E4r enkel: best\u00E4m om bilden \u00E4r \u00E4kta eller fejk. R\u00E4ck upp handen f\u00F6r \u00E4kta, h\u00E5ll ner f\u00F6r fejk.\u201D"),
      bodyText("Visa tre bilder i snabb f\u00F6ljd. F\u00F6r varje bild: l\u00E5t eleverna r\u00F6sta, avsl\u00F6ja svaret, kort kommentar."),
      bodyText("F\u00F6rslag p\u00E5 bilder:"),
      bullet("Bild 1: En \u00E4kta nyhetsbild fr\u00E5n en k\u00E4nd h\u00E4ndelse (\u00E4kta) \u2014 t.ex. en SVT-bild fr\u00E5n en valr\u00F6relse eller klimatdemonstration"),
      bullet("Bild 2: P\u00E5ven i vit dunjacka \u2014 den virala AI-bilden fr\u00E5n mars 2023, skapad i Midjourney (fejk). S\u00F6k p\u00E5 \u201CPope puffer jacket AI\u201D"),
      bullet("Bild 3: AI-genererad bild av en explosion vid Pentagon (maj 2023) som fick b\u00F6rsen att falla (fejk). S\u00F6k p\u00E5 \u201CPentagon explosion AI image 2023\u201D"),
      spacer(),
      italicText("Landa i: \u201CDet var inte s\u00E5 l\u00E4tt, eller hur? Idag f\u00E5r ni verktyg f\u00F6r att bli b\u00E4ttre p\u00E5 att avg\u00F6ra vad som \u00E4r sant \u2014 s\u00E4rskilt n\u00E4r AI kan skapa inneh\u00E5ll som ser helt \u00E4kta ut.\u201D"),
      bodyText("Visa l\u00E4randem\u00E5let p\u00E5 tavlan."),

      heading3("Instruktion (8\u201325 min)"),
      bodyText("Presentera de fyra k\u00E4llkritiska grundfr\u00E5gorna. Max 17 minuter. Skriv fr\u00E5gorna p\u00E5 tavlan under genomg\u00E5ngen."),
      spacer(),
      boldBodyText("1. Vem? (3 min) ", "\u2014 Vem ligger bakom informationen? \u00C4r det en journalist, en myndighet, en anonym anv\u00E4ndare, eller en AI? Koppla till AI: \u201COm en text \u00E4r skriven av ChatGPT \u2014 vem \u00E4r d\u00E5 avs\u00E4ndaren?\u201D"),
      boldBodyText("2. Varf\u00F6r? (3 min) ", "\u2014 Vad \u00E4r syftet? Informera, \u00F6vertyga, s\u00E4lja, underh\u00E5lla, manipulera? Koppla till AI: \u201CWarum kan AI-genererade nyheter vara sv\u00E5ra att bed\u00F6ma \u2014 de har inget eget syfte.\u201D"),
      spacer(),
      boldItalicBody("EPA-stopp (4 min): ", "\u201CT\u00E4nk enskilt i 1 minut: Vilken av grundfr\u00E5gorna \u00E4r sv\u00E5rast att besvara n\u00E4r inneh\u00E5llet \u00E4r AI-genererat? Diskutera med din granne i 2 minuter.\u201D Samla 2\u20133 svar."),
      spacer(),
      boldBodyText("3. Hur? (3 min) ", "\u2014 Hur presenteras informationen? Vilka k\u00E4nslor v\u00E4cks? Anv\u00E4nds laddade ord? Koppla till AI: Visa ett exempel p\u00E5 AI-genererad text som l\u00E5ter trov\u00E4rdig men inneh\u00E5ller felaktigheter."),
      boldBodyText("4. N\u00E4r? (2 min) ", "\u2014 N\u00E4r publicerades det? \u00C4r informationen aktuell? Koppla till AI: \u201CAI-modeller har en tr\u00E4ningsgr\u00E4ns \u2014 de kan presentera gammal information som ny.\u201D"),

      heading3("Bearbetning (25\u201360 min)"),
      boldBodyText("Del 1 \u2014 Sorterings\u00F6vning: \u201C\u00C4kta, AI eller manipulerat?\u201D (25\u201340 min)", ""),
      spacer(),
      bodyText("Dela in eleverna i grupper om 3\u20134. Dela ut sorteringskorten (se bilaga). Varje kort har en text, bild eller sk\u00E4rmdump."),
      spacer(),
      italicText("Instruktion: \u201CNi har 10 kort. Sortera dem i tre h\u00F6gar: \u00C4KTA, AI-GENERERAT och MANIPULERAT. Anv\u00E4nd grundfr\u00E5gorna p\u00E5 tavlan som hj\u00E4lp. Diskutera och motivera \u2014 det finns inga sj\u00E4lvklara svar. Ni har 12 minuter.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 skriv p\u00E5 tavlan:"),
      bullet("Vem ligger bakom det h\u00E4r?"),
      bullet("Varf\u00F6r har det skapats?"),
      bullet("Hur presenteras informationen?"),
      bullet("Finns det n\u00E5got som k\u00E4nns \u201Coff\u201D?"),
      spacer(),
      bodyText("Cirkulera mellan grupperna. Fr\u00E5ga: \u201CVarf\u00F6r la ni det kortet d\u00E4r?\u201D, \u201CVad var det som avgjorde?\u201D"),
      bodyText("Avsluta med 3 minuters helklassgenomg\u00E5ng: Lyft 2\u20133 kort som var sv\u00E5ra. Fr\u00E5ga: \u201CVarf\u00F6r var just de h\u00E4r sv\u00E5ra att bed\u00F6ma?\u201D"),
      spacer(),
      boldBodyText("Del 2 \u2014 Analysmatris (40\u201360 min)", ""),
      spacer(),
      bodyText("Dela ut analysmatrisen (se bilaga). Eleverna arbetar i par."),
      spacer(),
      italicText("Instruktion: \u201CV\u00E4lj 2\u20133 av korten som ni tyckte var sv\u00E5ra. Analysera dem med hj\u00E4lp av matrisen \u2014 fyll i en rad per kort. F\u00F6rsta raden \u00E4r ifylld som exempel.\u201D"),
      spacer(),
      bodyText("Cirkulera och ge st\u00F6d. Uppmuntra elever som \u00E4r klara tidigt: \u201CKan ni hitta ytterligare tecken p\u00E5 att inneh\u00E5llet \u00E4r AI-genererat?\u201D"),

      heading3("F\u00F6rdjupning: AI-labb (60\u201380 min)"),
      italicText("S\u00E4g: \u201CNu har ni \u00F6vat p\u00E5 att avsl\u00F6ja fejkat inneh\u00E5ll. Nu ska ni byta sida \u2014 er uppgift \u00E4r att skapa n\u00E5got som ser \u00E4kta ut.\u201D"),
      spacer(),
      bodyText("Eleverna arbetar i par. Varje par f\u00E5r i uppgift att:"),
      bullet("Anv\u00E4nda ett AI-verktyg (t.ex. ChatGPT, Copilot) f\u00F6r att skapa en kort nyhetstext ELLER en AI-genererad bild"),
      bullet("M\u00E5let: g\u00F6ra det s\u00E5 trov\u00E4rdigt som m\u00F6jligt"),
      bullet("Tidsgr\u00E4ns: 12 minuter"),
      spacer(),
      bodyText("Efter 12 minuter: Paren byter produkter med ett annat par. Det mottagande paret granskar med grundfr\u00E5gorna och ska identifiera vad som \u00E4r AI-genererat och varf\u00F6r."),
      spacer(),
      bodyText("Helklassgenomg\u00E5ng (5 min): Lyft de b\u00E4sta exemplen. \u201CVad gjorde det sv\u00E5rt att avsl\u00F6ja?\u201D"),

      heading3("Summering (80\u201395 min)"),
      bodyText("Helklassdiskussion (8 min):"),
      bullet("\u201CVad var sv\u00E5rast att avsl\u00F6ja idag?\u201D"),
      bullet("\u201CVilken grundfr\u00E5ga var mest anv\u00E4ndbar?\u201D"),
      bullet("\u201CHur f\u00F6r\u00E4ndrar AI v\u00E5r m\u00F6jlighet att granska information?\u201D"),
      spacer(),
      bodyText("Koppla tillbaka till l\u00E4randem\u00E5let p\u00E5 tavlan."),
      spacer(),
      boldBodyText("Exit ticket (7 min):", ""),
      italicText("\u201CNamnge de fyra k\u00E4llkritiska grundfr\u00E5gorna. V\u00E4lj sedan EN av dem och f\u00F6rklara varf\u00F6r just den blir extra viktig n\u00E4r inneh\u00E5ll \u00E4r AI-genererat.\u201D"),
      spacer(),
      bodyText("Scaffolding \u2014 meningsstartare p\u00E5 tavlan:"),
      bullet("\u201CDe fyra grundfr\u00E5gorna \u00E4r...\u201D"),
      bullet("\u201CJag tycker att fr\u00E5gan om... \u00E4r viktigast n\u00E4r det g\u00E4ller AI, eftersom...\u201D"),
      spacer(),
      bodyText("Samla in (digitalt eller p\u00E5 papper). Anv\u00E4nd svaren f\u00F6r att planera retrieval practice i lektion 2."),

      heading3("Fram\u00E5tkoppling (95\u2013100 min)"),
      italicText("\u201CNu har vi verktyg f\u00F6r att granska k\u00E4llor \u2014 vi vet hur vi st\u00E4ller r\u00E4tt fr\u00E5gor. N\u00E4sta g\u00E5ng tittar vi p\u00E5 n\u00E5got som g\u00F6r det \u00E4nnu sv\u00E5rare: konspirationsteorier. Vem tj\u00E4nar p\u00E5 att du tror p\u00E5 dem? Och varf\u00F6r \u00E4r de s\u00E5 effektiva?\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "R\u00F6stning: \u00E4kta eller fejk? (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "EPA: vilken grundfr\u00E5ga \u00E4r sv\u00E5rast f\u00F6r AI-inneh\u00E5ll? (enskilt + par, 4 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Sorterings\u00F6vning i grupp: \u00E4kta / AI / manipulerat (grupp 3\u20134, 15 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Analysmatris: granska 2\u20133 k\u00E4llor med grundfr\u00E5gorna (par, 20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "AI-labb: skapa trov\u00E4rdigt AI-inneh\u00E5ll + granska andras (par, 20 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Exit ticket: grundfr\u00E5gorna + resonemang om AI (enskilt, 7 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 71 av 100 minuter (71%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Analysmatrisen har en ifylld exempelrad. Grundfr\u00E5gorna \u00E4r synliga p\u00E5 tavlan under hela lektionen. Sorterings\u00F6vningen har hj\u00E4lpfr\u00E5gor. Exit ticket har meningsstartare. L\u00E4raren prioriterar dessa elever vid cirkulering och ger konkreta ledtr\u00E5dar: \u201CTitta p\u00E5 vem som publicerat det h\u00E4r \u2014 vad s\u00E4ger det dig?\u201D"),
      boldBodyText("Utmaning (mot A): ", "Under AI-labben: \u201CKan ni skapa n\u00E5got som klarar alla fyra grundfr\u00E5gorna \u2014 som fortfarande ser trov\u00E4rdigt ut trots granskning?\u201D Vid analysmatrisen: l\u00E4gg till kolumnen \u201CVad saknas?\u201D \u2014 vilken information beh\u00F6ver du som inte finns? Exit ticket utan meningsstartare, med till\u00E4gget: \u201CResonera om vad detta inneb\u00E4r f\u00F6r demokratin.\u201D"),

      // MATERIAL
      heading2("Material"),
      bullet("\u00C4kta-eller-fejk-bildspel med 8\u201310 exempel (bilder, texter, sk\u00E4rmdumpar)"),
      bullet("Sorteringskort \u201C\u00C4kta, AI eller manipulerat?\u201D \u2014 10 kort per grupp (se bilaga)"),
      bullet("Analysmatris \u201CK\u00E4llkritiska grundfr\u00E5gor\u201D \u2014 en per elev (se bilaga)"),
      bullet("Tillg\u00E5ng till AI-verktyg f\u00F6r AI-labben (elevernas datorer/telefoner)"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Sorterings\u00F6vningen och analysmatrisen tr\u00E4nar f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor (m\u00E5l 1: E\u2013A)"),
      bullet("EPA-diskussionen och AI-labben tr\u00E4nar resonemang om AI-genererat inneh\u00E5lls p\u00E5verkan p\u00E5 informationslandskapet (m\u00E5l 2: E\u2013A)"),
      bullet("Exit ticket ger formativ information om b\u00E5de faktakunskap (grundfr\u00E5gorna) och analysf\u00F6rm\u00E5ga (resonemang om AI)"),
      spacer(),
      bodyText("Scaffolding-strukturerna (hj\u00E4lpfr\u00E5gor, ifyllt exempel, meningsstartare) ger alla elever m\u00F6jlighet att n\u00E5 minst E-niv\u00E5. AI-labben och de \u00F6ppnare uppgifterna ger A-elever utrymme att visa nyanserat resonemang om AI:s samh\u00E4llsp\u00E5verkan."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("N\u00E4sta lektion: Konspirationsteorier \u2014 \u201CVem tj\u00E4nar p\u00E5 att du tror det h\u00E4r?\u201D"),
      bullet("Exit ticket-data anv\u00E4nds f\u00F6r retrieval practice i lektion 2"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: SORTERINGSKORT
      heading2("Bilaga 1: Sorteringskort \u2014 \u201C\u00C4kta, AI eller manipulerat?\u201D"),
      bodyText("Klipp ut korten. Ett set per grupp. Varje kort beskriver en k\u00E4lla."),
      spacer(),
      boldBodyText("Kort 1 (\u00C4KTA): ", "SMHI:s sida om FN:s klimatpanel IPCC. Publicerad av myndigheten SMHI med namngivna forskare och h\u00E4nvisningar till IPCC:s rapporter. L\u00E4nk: smhi.se/klimat/klimatarbetet-pa-smhi/fns-klimatpanel-ipcc"),
      boldBodyText("Kort 2 (AI): ", "Bilden p\u00E5 \u201Cp\u00E5ven i en vit dunjacka\u201D som spreds viralt i mars 2023. Skapad med AI-verktyget Midjourney av en privatperson i Chicago. Ingen k\u00E4lla angiven, delad av anonyma konton. Se Snopes faktagranskning: snopes.com/fact-check/not-real-photo-pope-in-puffy-coat/"),
      boldBodyText("Kort 3 (MANIPULERAT): ", "En blogg som h\u00E4vdar att kosttillskott botar cancer. Inga k\u00E4llh\u00E4nvisningar. Sidan s\u00E4ljer \u201Cnaturliga kosttillskott\u201D. J\u00E4mf\u00F6r med Cancerfondens varning: cancerfonden.se/nyhet/kosttillskott-kan-fa-cancerceller-att-spridas"),
      boldBodyText("Kort 4 (\u00C4KTA): ", "Ett klipp fr\u00E5n SVT Nyheter d\u00E4r en minister uttalar sig om ett politiskt beslut. Publicerat p\u00E5 SVT:s officiella sida med datum och kontext. L\u00E4nk: svt.se/nyheter/inrikes/"),
      boldBodyText("Kort 5 (MANIPULERAT): ", "En fabricerad sk\u00E4rmdump fr\u00E5n X/Twitter d\u00E4r NFL-spelaren Travis Kelce p\u00E5st\u00E5s ha hotat Elon Musk (september 2024). Skaparen erk\u00E4nde att det var p\u00E5hittat. Faktagranskning: snopes.com/fact-check/kelce-musk-twist-pretzel/"),
      boldBodyText("Kort 6 (AI): ", "En \u201Cnyhetsartikel\u201D med perfekt spr\u00E5k och en bild av en journalist som inte finns p\u00E5 n\u00E5gon redaktion. Publicerad p\u00E5 en sajt som startades f\u00F6r 2 veckor sedan. NewsGuard har kartlagt \u00F6ver 2 000 s\u00E5dana AI-genererade nyhetssajter: newsguardtech.com/special-reports/ai-tracking-center/"),
      boldBodyText("Kort 7 (\u00C4KTA): ", "Svenskspr\u00E5kig Wikipedia-artikel om Europeiska unionen. K\u00E4llh\u00E4nvisningar till akademisk forskning och officiella EU-dokument. Senast redigerad f\u00F6r 3 dagar sedan. L\u00E4nk: sv.wikipedia.org/wiki/Europeiska_unionen"),
      boldBodyText("Kort 8 (AI): ", "En AI-genererad bild av en \u00F6versv\u00E4mmad stad delad p\u00E5 Facebook med texten \u201CTITTA! S\u00E5 ser det ut just nu!\u201D J\u00E4mf\u00F6r med AI-genererade bilder fr\u00E5n orkanen Helene 2024 som spreds som \u00E4kta: npr.org/2024/10/18/nx-s1-5153741/ai-images-hurricanes-disasters-propaganda"),
      boldBodyText("Kort 9 (MANIPULERAT): ", "Ett TikTok-klipp d\u00E4r en person \u201Cavsl\u00F6jar\u201D att mobilstr\u00E5lning orsakar cancer. 2 miljoner visningar men inga k\u00E4llor. J\u00E4mf\u00F6r med Str\u00E5ls\u00E4kerhetsmyndighetens bed\u00F6mning: stralsakerhetsmyndigheten.se/omraden/magnetfalt-och-tradlos-teknik/mobiltelefoni/"),
      boldBodyText("Kort 10 (\u00C4KTA): ", "Rapporten \u201CSvenskarna och internet 2025\u201D fr\u00E5n Internetstiftelsen. Publicerad med metodbeskrivning, statistik och fullst\u00E4ndig PDF. L\u00E4nk: svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),

      spacer(),
      bodyText("Facit (f\u00F6r l\u00E4raren):"),
      bullet("\u00C4kta: Kort 1, 4, 7, 10"),
      bullet("AI-genererat: Kort 2, 6, 8"),
      bullet("Manipulerat/vilseledande: Kort 3, 5, 9"),

      // BILAGA 2: ANALYSMATRIS
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Analysmatris \u2014 K\u00E4llkritiska grundfr\u00E5gor"),
      bodyText("Fyll i en rad per k\u00E4lla du analyserar."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1200, 1560, 1560, 1560, 1560, 1586],
        rows: [
          new TableRow({
            children: [
              headerCell("K\u00E4lla", 1200),
              headerCell("Vem?", 1560),
              headerCell("Varf\u00F6r?", 1560),
              headerCell("Hur?", 1560),
              headerCell("N\u00E4r?", 1560),
              headerCell("Bed\u00F6mning", 1586),
            ],
          }),
          new TableRow({
            children: [
              cell("Kort 6 (exempel)", 1200),
              cell("Ok\u00E4nd \u201Cjournalist\u201D p\u00E5 ny sajt", 1560),
              cell("Ser ut som nyheter men syftet \u00E4r oklart", 1560),
              cell("Perfekt spr\u00E5k, ser proffsigt ut", 1560),
              cell("Sajten \u00E4r 2 veckor gammal", 1560),
              cell("Troligen AI-genererat: \u201Cjournalisten\u201D finns inte, sajten \u00E4r ny", 1586),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1200),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1586),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1200),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1586),
            ],
          }),
          new TableRow({
            children: [
              cell("", 1200),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1560),
              cell("", 1586),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Ordbank: ", "avs\u00E4ndare, k\u00E4lla, syfte, trov\u00E4rdighet, vinkling, AI-genererat, deepfake, manipulation, verifiering, faktagranskning, byline, anonym, algorithm, informationsp\u00E5verkan"),

      // BILAGA 3: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Exit ticket \u2014 Lektion 1"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "Namnge de fyra k\u00E4llkritiska grundfr\u00E5gorna. V\u00E4lj sedan EN av dem och f\u00F6rklara varf\u00F6r just den blir extra viktig n\u00E4r inneh\u00E5ll \u00E4r AI-genererat."),
      spacer(),
      bodyText("Meningsstartare:"),
      bullet("\u201CDe fyra k\u00E4llkritiska grundfr\u00E5gorna \u00E4r...\u201D"),
      bullet("\u201CJag tycker att fr\u00E5gan om... \u00E4r viktigast n\u00E4r det g\u00E4ller AI, eftersom...\u201D"),
      bullet("\u201CEtt exempel p\u00E5 n\u00E4r den fr\u00E5gan \u00E4r avg\u00F6rande \u00E4r...\u201D"),

      // BILAGA 4: KÄLLMATERIAL OCH LÄNKAR
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: K\u00E4llmaterial och l\u00E4nkar"),
      bodyText("Samlade l\u00E4nkar f\u00F6r sorteringskort, uppstarts\u00F6vning och f\u00F6rdjupning."),

      heading3("Uppstarts\u00F6vning \u2014 bilder f\u00F6r \u201C\u00C4kta eller fejk?\u201D"),
      bullet("P\u00E5ven i dunjacka (AI, mars 2023): snopes.com/fact-check/not-real-photo-pope-in-puffy-coat/"),
      bullet("Pentagon-explosion (AI, maj 2023): npr.org/2023/05/22/1177590231/fake-viral-images-of-an-explosion-at-the-pentagon-were-probably-created-by-ai"),
      bullet("Katy Perry p\u00E5 Met Gala (AI, maj 2024): npr.org/2024/05/07/1249570785/katy-perry-met-gala-deepfake"),

      heading3("K\u00E4llor f\u00F6r sorteringskorten"),
      boldBodyText("Kort 1 \u2014 SMHI om IPCC: ", "smhi.se/klimat/klimatarbetet-pa-smhi/fns-klimatpanel-ipcc"),
      boldBodyText("Kort 2 \u2014 P\u00E5ven i dunjacka (Snopes): ", "snopes.com/fact-check/not-real-photo-pope-in-puffy-coat/"),
      boldBodyText("Kort 2 \u2014 Skaparen ber\u00E4ttar (BuzzFeed): ", "buzzfeednews.com/article/chrisstokelwalker/pope-puffy-jacket-ai-midjourney-image-creator-interview"),
      boldBodyText("Kort 3 \u2014 Cancerfonden varnar: ", "cancerfonden.se/nyhet/kosttillskott-kan-fa-cancerceller-att-spridas"),
      boldBodyText("Kort 3 \u2014 Livsmedelsverket om risker: ", "livsmedelsverket.se/livsmedel-och-innehall/kosttillskott/risker-med-kosttillskott/"),
      boldBodyText("Kort 4 \u2014 SVT Nyheter: ", "svt.se/nyheter/inrikes/"),
      boldBodyText("Kort 5 \u2014 Kelce/Musk fejktweet (Snopes): ", "snopes.com/fact-check/kelce-musk-twist-pretzel/"),
      boldBodyText("Kort 5 \u2014 Guide: verifiera tweets: ", "gijn.org/resource/simple-tips-for-verifying-if-a-tweet-screenshot-is-real-or-fake/"),
      boldBodyText("Kort 6 \u2014 AI-nyhetssajter (NewsGuard): ", "newsguardtech.com/special-reports/ai-tracking-center/"),
      boldBodyText("Kort 7 \u2014 Wikipedia om EU: ", "sv.wikipedia.org/wiki/Europeiska_unionen"),
      boldBodyText("Kort 8 \u2014 AI-katastrofbilder (NPR): ", "npr.org/2024/10/18/nx-s1-5153741/ai-images-hurricanes-disasters-propaganda"),
      boldBodyText("Kort 8 \u2014 Fejkbilder p\u00E5 Facebook (sv): ", "digitalfotoforalla.se/fotografering/falska-bilder-invaderar-facebook"),
      boldBodyText("Kort 9 \u2014 Str\u00E5ls\u00E4kerhetsmyndigheten: ", "stralsakerhetsmyndigheten.se/omraden/magnetfalt-och-tradlos-teknik/mobiltelefoni/"),
      boldBodyText("Kort 9 \u2014 SVT om 5G-oro: ", "svt.se/nyheter/vetenskap/5g-ar-det-nya-mobilnatet-en-halsofara"),
      boldBodyText("Kort 10 \u2014 Svenskarna och internet 2025: ", "svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),
      boldBodyText("Kort 10 \u2014 Sammanfattning (PDF): ", "svenskarnaochinternet.se/app/uploads/2025/09/svenskarna-och-internet-2025-sammanfattning.pdf"),

      heading3("Faktagranskning och k\u00E4llkritik \u2014 resurser"),
      bullet("K\u00E4llkritikbyr\u00E5n (Sveriges faktagranskare): kallkritikbyran.se"),
      bullet("Internetkunskap \u2014 k\u00E4llkritik: internetkunskap.se/kallkritik/"),
      bullet("Internetkunskap \u2014 avsl\u00F6ja fejkbilder: internetkunskap.se/artiklar/grundkurs-i-ai/blotta-ogat-racker-inte-sa-avslojar-du-falska-bilder-i-en-tid-av-ai/"),
      bullet("Krisinformation.se \u2014 k\u00E4llkritik: krisinformation.se/detta-gor-samhallet/kallkritik/"),
      bullet("MPF \u2014 AI och informationsp\u00E5verkan: mpf.se/kunskap-och-stod/temabibliotek/ai-och-informationspaverkan"),
      bullet("SO-rummet \u2014 k\u00E4llkritik (de fyra kriterierna): so-rummet.se/content/kallkritik-fyra-viktiga-kriterier"),
      bullet("Digitala lektioner \u2014 k\u00E4llkritik: digitalalektioner.se/amnesomrade/digital-kallkritik/"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-1.docx", buffer);
  console.log("lektion-1.docx skapad!");
});
