const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak
} = require("docx");
const pptxgen = require("pptxgenjs");

// ============================================================
// HELPER FUNCTIONS
// ============================================================

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

// ============================================================
// DOCX - Lektion 5: Perspektiven krockar (reviderad)
// ============================================================

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
        children: [new TextRun({ text: "Lektion 5: Perspektiven krockar", font: "Arial" })],
      }),
      bodyText("Skriftlig perspektivanalys \u2014 dress rehearsal inf\u00F6r L6"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "85 minuter"),
      boldBodyText("Grupp: ", "MSA23"),
      spacer(),
      bodyText("Lektion 5 \u00E4r en f\u00F6rberedande \u00F6vning inf\u00F6r den summativa skrivuppgiften i lektion 6. Eleverna skriver en perspektivanalys i samma format som L6 men kortare (300\u2013400 ord ist\u00E4llet f\u00F6r l\u00E4ngre). Temat \u00E4r bekant fr\u00E5n L4-seminariet, vilket ger en naturlig \u00F6verg\u00E5ng fr\u00E5n muntligt till skriftligt. L\u00E4raren samlar in texterna och ger kort formativ feedback inf\u00F6r L6."),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven ska genom en skriftlig perspektivanalys kunna analysera olika samh\u00E4llsakt\u00F6rers ansvar f\u00F6r AI-driven desinformation, argumentera fr\u00E5n flera perspektiv och h\u00E4nvisa till k\u00E4llor med Harvard-systemet."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor och argument fr\u00E5n olika samh\u00E4llsakt\u00F6rer (m\u00E5l 1)"),
      bullet("Analysera hur AI-genererat inneh\u00E5ll p\u00E5verkar informationslandskapet ur olika perspektiv (m\u00E5l 2)"),
      bullet("Analysera konspirationsteoriers spridning genom att v\u00E4ga in teknikf\u00F6retagens, statens och medborgarnas roller (m\u00E5l 3)"),
      bullet("K\u00E4llh\u00E4nvisa korrekt enligt Harvard-systemet (m\u00E5l 4)"),
      spacer(),
      boldBodyText("E: ", "Eleven skriver en text som presenterar alla tre perspektiv p\u00E5 ett enkelt s\u00E4tt och g\u00F6r enkla omd\u00F6men om akt\u00F6rernas ansvar. K\u00E4llh\u00E4nvisar med viss s\u00E4kerhet."),
      boldBodyText("C: ", "Eleven skriver en text med v\u00E4lgrundade argument fr\u00E5n alla tre perspektiv, underbyggda med konkreta exempel och k\u00E4llor. K\u00E4llh\u00E4nvisar korrekt och landar i en egen motiverad slutsats."),
      boldBodyText("A: ", "Eleven skriver en text med v\u00E4lgrundade och nyanserade argument som visar p\u00E5 komplexa samband mellan akt\u00F6rerna. K\u00E4llh\u00E4nvisar med s\u00E4kerhet, anv\u00E4nder minst 3 k\u00E4llor och resonerar om varf\u00F6r kombinationen av perspektiv beh\u00F6vs \u2014 eller inte g\u00F6r det."),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),
      bullet("Demokrati och m\u00E4nskliga r\u00E4ttigheter i en digital kontext."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("Analysera exit tickets fr\u00E5n lektion 4 (seminariet) \u2014 vilka perspektiv och argument lyftes?"),
      bullet("Skriv ut rollkorten (en upps\u00E4ttning per elev eller projicera, se bilaga 1) \u2014 anv\u00E4nds som skrivst\u00F6d"),
      bullet("F\u00F6rbered tavlan med skrivfr\u00E5gan: \u201CVem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation?\u201D"),
      bullet("Ha Harvard-guiden (fr\u00E5n lektion 2/3) tillg\u00E4nglig \u2014 projicera eller skriv ut"),
      bullet("Ha l\u00E4randem\u00E5let synligt p\u00E5 tavlan under hela lektionen"),
      bullet("S\u00E4kerst\u00E4ll att eleverna har tillg\u00E5ng till datorer f\u00F6r skrivande"),
      bullet("F\u00F6rbered scaffolding-material (bilaga 6) \u2014 anv\u00E4nd ENBART om du bed\u00F6mer att det beh\u00F6vs"),

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
              cell("Retrieval practice", 2200),
              cell("Retrieval practice fr\u00E5n lektion 4 (seminarium): \u201CSkriv ner det starkaste argumentet du h\u00F6rde ig\u00E5r \u2014 oavsett vilken sida det kom ifr\u00E5n.\u201D Enskilt skrivande (3 min), sedan delar 3\u20134 elever muntligt.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("8\u201315 min", 1000),
              cell("Instruktion", 1300),
              cell("Presentera skrivuppgiften", 2200),
              cell("Presentera uppgiften: skriftlig perspektivanalys. G\u00E5 igenom f\u00F6rv\u00E4ntningar (300\u2013400 ord, alla tre perspektiv, egen slutsats, minst 2 k\u00E4llor med Harvard). Kort p\u00E5minnelse om Harvard-systemet (2 min). Dela ut rollkorten som skrivst\u00F6d.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("15\u201360 min", 1000),
              cell("Bearbetning", 1300),
              cell("Skriftlig perspektivanalys", 2200),
              cell("Eleverna skriver individuellt (45 min). Fr\u00E5ga: \u201CVem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation?\u201D Skriv fr\u00E5n ALLA TRE perspektiv (teknikf\u00F6retag, stat/EU, medborgare) och landa i en egen slutsats. 300\u2013400 ord, minst 2 k\u00E4llor med Harvard-h\u00E4nvisning. Rollkorten (bilaga 1) f\u00E5r anv\u00E4ndas som st\u00F6d.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("60\u201365 min", 1000),
              cell("Revision", 1300),
              cell("Individuell genoml\u00E4sning", 2200),
              cell("Eleverna l\u00E4ser igenom sin text och g\u00F6r f\u00F6rb\u00E4ttringar. Fokus: Finns alla tre perspektiv? \u00C4r k\u00E4llh\u00E4nvisningarna korrekta? \u00C4r slutsatsen tydlig?", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("65\u201375 min", 1000),
              cell("Summering", 1300),
              cell("Helklassdiskussion", 2200),
              cell("Lyft m\u00F6nster: \u201CVilken akt\u00F6r n\u00E4mndes oftast som huvudansvarig? Varf\u00F6r?\u201D Koppla till demokrati. Diskutera: \u201CKan en akt\u00F6r l\u00F6sa det h\u00E4r ensam?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("75\u201382 min", 1000),
              cell("Exit ticket", 1300),
              cell("Metakognitiv reflektion", 2200),
              cell("Exit ticket: \u201CL\u00E4s igenom din perspektivanalys. Markera en mening d\u00E4r du anv\u00E4nde ett k\u00E4llkritiskt verktyg och en mening d\u00E4r du kunde ha st\u00E4rkt din argumentation. Vad skulle du g\u00F6ra annorlunda?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("82\u201385 min", 1000),
              cell("Fram\u00E5tkoppling", 1300),
              cell("N\u00E4sta lektion + insamling", 2200),
              cell("Fram\u00E5tkoppling till L6: \u201CN\u00E4sta lektion skriver ni en liknande analys \u2014 den \u00E4r ert bed\u00F6mningsunderlag. Jag samlar in era texter nu och ger er kort feedback innan dess. Anv\u00E4nd feedbacken f\u00F6r att f\u00F6rbereda er.\u201D Samla in texterna.", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Uppstart: Retrieval practice (0\u20138 min)"),
      italicText("S\u00E4g: \u201CIg\u00E5r hade ni seminarium om AI-driven desinformation. Innan vi g\u00E5r vidare vill jag att ni t\u00E4nker tillbaka. Skriv ner det starkaste argumentet ni h\u00F6rde ig\u00E5r \u2014 oavsett vilken sida det kom ifr\u00E5n. Ni har 3 minuter.\u201D"),
      spacer(),
      bodyText("L\u00E5t eleverna skriva enskilt. Samla sedan 3\u20134 svar muntligt. Anv\u00E4nd svaren f\u00F6r att skapa en bro till dagens uppgift:"),
      italicText("\u201CBra \u2014 ni har redan starka argument i huvudet. Idag ska ni anv\u00E4nda dem, men p\u00E5 ett nytt s\u00E4tt.\u201D"),

      heading3("Instruktion: Presentera skrivuppgiften (8\u201315 min)"),
      italicText("S\u00E4g: \u201CIdag ska ni skriva, inte prata. Ni f\u00E5r samma fr\u00E5ga som i seminariet: Vem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation? Men den h\u00E4r g\u00E5ngen ska ni skriva fr\u00E5n ALLA TRE perspektiv \u2014 teknikf\u00F6retagen, staten och medborgarna \u2014 och sedan landa i en egen slutsats.\u201D"),
      spacer(),
      italicText("S\u00E4g: \u201CDet h\u00E4r \u00E4r en \u00F6vning inf\u00F6r n\u00E4sta lektions bed\u00F6mningsuppgift. Ni skriver i exakt samma format som ni kommer anv\u00E4nda d\u00E5 \u2014 s\u00E5 se det som en general\u00F6vning. Jag samlar in era texter efterr\u00E5t och ger er kort feedback.\u201D"),
      spacer(),
      bodyText("G\u00E5 igenom f\u00F6rv\u00E4ntningarna:"),
      bullet("300\u2013400 ord"),
      bullet("Alla tre perspektiv ska finnas med"),
      bullet("En egen slutsats d\u00E4r du tar st\u00E4llning"),
      bullet("Minst 2 k\u00E4llor med Harvard-h\u00E4nvisning"),
      bullet("Rollkorten (bilaga 1) f\u00E5r anv\u00E4ndas som st\u00F6d"),
      spacer(),
      boldBodyText("Harvard-p\u00E5minnelse (2 min): ", ""),
      italicText("S\u00E4g: \u201CNi har \u00F6vat p\u00E5 Harvard-systemet i lektion 2 och 3. Kort p\u00E5minnelse: h\u00E4nvisning i texten med (F\u00F6rfattare, \u00E5r) och en referenslista i slutet. Harvard-guiden finns framme om ni beh\u00F6ver den.\u201D"),
      spacer(),
      bodyText("Dela ut rollkorten (bilaga 1)."),
      spacer(),
      boldBodyText("Scaffolding p\u00E5 beg\u00E4ran: ", "Om du bed\u00F6mer att klassen eller enskilda elever beh\u00F6ver st\u00F6d, anv\u00E4nd meningsstartarna och textstrukturen i bilaga 6. Skriv upp dem p\u00E5 tavlan eller ge dem till enskilda elever. Dela INTE ut scaffolding som standard \u2014 l\u00E5t eleverna f\u00F6rs\u00F6ka sj\u00E4lva f\u00F6rst."),

      heading3("Bearbetning: Skriftlig perspektivanalys (15\u201360 min)"),
      bodyText("Eleverna skriver individuellt p\u00E5 dator. Skriv fr\u00E5gan p\u00E5 tavlan:"),
      boldBodyText("", "\u201CVem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation?\u201D"),
      spacer(),
      italicText("Instruktion: \u201CNi har 45 minuter. Skriv fr\u00E5n alla tre perspektiv och avsluta med er egen slutsats. Anv\u00E4nd rollkorten som st\u00F6d, men skriv med egna ord. Minst 2 k\u00E4llor med Harvard-h\u00E4nvisning \u2014 k\u00E4llorna finns p\u00E5 rollkorten och i bilaga 5.\u201D"),
      spacer(),
      bodyText("Cirkulera under skrivarbetet. Fokusera p\u00E5:"),
      bullet("Elever som inte kommer ig\u00E5ng: \u201CTitta p\u00E5 rollkortet f\u00F6r teknikf\u00F6retagen \u2014 vad \u00E4r deras st\u00F6rsta argument? B\u00F6rja d\u00E4r.\u201D"),
      bullet("K\u00E4llh\u00E4nvisningar: \u201CJag ser att du h\u00E4nvisar till EU:s AI Act \u2014 har du lagt in parentesen (EU, 2024)?\u201D"),
      bullet("Perspektivbredd: \u201CDu har skrivit bra om teknikf\u00F6retagen. Vad s\u00E4ger staten? Och medborgarna?\u201D"),
      bullet("Snabba elever: \u201CKan du l\u00E4gga till en tredje k\u00E4lla och j\u00E4mf\u00F6ra perspektiven mer?\u201D"),
      spacer(),
      boldBodyText("Vid behov av scaffolding: ", "Om en elev \u00E4r fast efter 5\u201310 minuter, ge individuellt st\u00F6d med hj\u00E4lp av bilaga 6 (meningsstartare, textstruktur). Skriv inte upp dem p\u00E5 tavlan f\u00F6r alla om det inte beh\u00F6vs."),

      heading3("Individuell revision (60\u201365 min)"),
      italicText("S\u00E4g: \u201CStopp \u2014 l\u00E4gg undan rollkorten. L\u00E4s igenom er text en sista g\u00E5ng. Kontrollera tre saker: Finns alla tre perspektiv? \u00C4r k\u00E4llh\u00E4nvisningarna korrekta? Har ni en tydlig slutsats? Ni har 5 minuter.\u201D"),
      spacer(),
      bodyText("Syftet \u00E4r att tr\u00E4na eleverna i sj\u00E4lvgranskning inf\u00F6r L6 d\u00E4r de inte f\u00E5r feedback under skrivandet."),

      heading3("Helklassdiskussion (65\u201375 min)"),
      italicText("S\u00E4g: \u201CJag \u00E4r nyfiken \u2014 vilken akt\u00F6r n\u00E4mndes oftast som huvudansvarig? R\u00E4ck upp handen: Hur m\u00E5nga landade i teknikf\u00F6retagen? Staten? Medborgarna? Kombination?\u201D"),
      spacer(),
      bodyText("Notera f\u00F6rdelningen p\u00E5 tavlan."),
      spacer(),
      bodyText("Diskussionsfr\u00E5gor:"),
      bullet("\u201CVilken akt\u00F6r n\u00E4mndes oftast? Varf\u00F6r tror ni det?\u201D"),
      bullet("\u201CKan en akt\u00F6r l\u00F6sa det h\u00E4r ensam? Vad beh\u00F6vs i kombination?\u201D"),
      bullet("\u201CHur kopplar den h\u00E4r fr\u00E5gan till demokrati? Vad h\u00E4nder med demokratin om ingen tar ansvar?\u201D"),
      bullet("\u201C\u00C4ndrade n\u00E5gon av er uppfattning under skrivandet? Vad var det som \u00E4ndrade er?\u201D"),
      spacer(),
      bodyText("Koppla tillbaka till l\u00E4randem\u00E5len. Betona att skrivandet tvingar fram djupare analys \u00E4n muntlig diskussion \u2014 man m\u00E5ste formulera argumenten tydligt och v\u00E4lja k\u00E4llor medvetet."),

      heading3("Exit ticket: Metakognitiv reflektion (75\u201382 min)"),
      boldBodyText("Exit ticket: ", ""),
      italicText("\u201CL\u00E4s igenom din perspektivanalys. Markera en mening d\u00E4r du anv\u00E4nde ett k\u00E4llkritiskt verktyg och en mening d\u00E4r du kunde ha st\u00E4rkt din argumentation. Vad skulle du g\u00F6ra annorlunda?\u201D"),
      spacer(),
      bodyText("Denna exit ticket \u00E4r metakognitiv \u2014 eleven reflekterar \u00F6ver sin egen text och identifierar f\u00F6rb\u00E4ttringsomr\u00E5den. Detta tr\u00E4nar sj\u00E4lvbed\u00F6mning inf\u00F6r L6."),
      spacer(),
      bodyText("Samla in exit tickets separat fr\u00E5n texterna. Exit ticket-svaren visar om eleverna utvecklat metakognitiv f\u00F6rm\u00E5ga \u2014 kan de identifiera styrkor och svagheter i sitt eget resonemang?"),

      heading3("Fram\u00E5tkoppling och insamling (82\u201385 min)"),
      italicText("S\u00E4g: \u201CN\u00E4sta lektion skriver ni en liknande analys \u2014 den \u00E4r ert bed\u00F6mningsunderlag. Jag samlar in era texter nu och ger er kort feedback innan dess. Anv\u00E4nd feedbacken f\u00F6r att f\u00F6rbereda er. Ni har redan \u00F6vat p\u00E5 argumenten, Harvard-h\u00E4nvisningarna och textstrukturen \u2014 det \u00E4r precis det ni g\u00F6r i L6, bara lite l\u00E4ngre.\u201D"),
      spacer(),
      boldBodyText("Samla in texterna. ", "Ge kort formativ feedback p\u00E5 varje elevtext innan L6: 2\u20133 styrkor och 1 f\u00F6rb\u00E4ttringsomr\u00E5de. Feedbacken \u00E4r formativ och ska hj\u00E4lpa eleven f\u00F6rbereda sig f\u00F6r den summativa uppgiften i L6. Returnera texterna med feedback inf\u00F6r n\u00E4sta lektion."),

      // LÄRARFEEDBACK-INSTRUKTION
      heading2("L\u00E4rarfeedback efter lektionen"),
      bodyText("L\u00E4rarfeedback p\u00E5 L5-texterna \u00E4r en central komponent i kopplingen L5\u2192L6. Feedbacken ska vara kort, konkret och framf\u00F6r\u00E5triktad."),
      spacer(),
      boldBodyText("Format: ", "2\u20133 styrkor + 1 f\u00F6rb\u00E4ttringsomr\u00E5de per elevtext."),
      spacer(),
      bodyText("Exempel p\u00E5 feedback:"),
      bullet("Styrka: \u201CDu presenterar alla tre perspektiv tydligt och anv\u00E4nder relevanta k\u00E4llor.\u201D"),
      bullet("Styrka: \u201CDin slutsats \u00E4r v\u00E4lmotiverad \u2014 du v\u00E4ger perspektiven mot varandra.\u201D"),
      bullet("F\u00F6rb\u00E4ttring: \u201CF\u00F6rs\u00F6k koppla perspektiven till varandra mer i L6. Hur p\u00E5verkar teknikf\u00F6retagens agerande statens m\u00F6jligheter?\u201D"),
      spacer(),
      boldBodyText("Tidsram: ", "Returnera texterna med feedback senast dagen f\u00F6re L6."),
      boldBodyText("Syfte: ", "Feedbacken hj\u00E4lper eleverna identifiera konkreta f\u00F6rb\u00E4ttringsomr\u00E5den inf\u00F6r den summativa uppgiften. Den \u00E4r formativ \u2014 inget betyg s\u00E4tts p\u00E5 L5-texten."),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Retrieval practice: starkaste argumentet fr\u00E5n seminariet (enskilt + helklass, 8 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Skriftlig perspektivanalys med k\u00E4llh\u00E4nvisning (enskilt, 45 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Individuell revision och sj\u00E4lvgranskning (enskilt, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Helklassdiskussion: m\u00F6nster och demokratikoppling (helklass, 10 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Metakognitiv exit ticket (enskilt, 7 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 75 av 85 minuter (88%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Rollkorten (bilaga 1) inneh\u00E5ller f\u00E4rdiga argument med k\u00E4llor \u2014 eleverna kan anv\u00E4nda dem som skrivst\u00F6d. Harvard-guiden fr\u00E5n lektion 2/3 ligger framme. Scaffolding (bilaga 6) finns tillg\u00E4nglig p\u00E5 beg\u00E4ran: meningsstartare och textstruktur. L\u00E4raren prioriterar dessa elever vid cirkulering."),
      boldBodyText("Utmaning (mot A): ", "Skriv utan rollkort \u2014 anv\u00E4nd egna argument och k\u00E4llor. Till\u00E4ggsuppgift: \u201CArgumentera f\u00F6r varf\u00F6r kombinationen av alla tre akt\u00F6rer beh\u00F6vs \u2014 eller varf\u00F6r den INTE g\u00F6r det.\u201D Anv\u00E4nd minst 3 k\u00E4llor med Harvard-h\u00E4nvisning."),

      // MATERIAL
      heading2("Material"),
      bullet("Rollkort f\u00F6r tre positioner (se bilaga 1) \u2014 anv\u00E4nds som skrivst\u00F6d"),
      bullet("Harvard-guide (fr\u00E5n lektion 2/3) \u2014 projicera eller skriv ut"),
      bullet("L\u00E4randem\u00E5let utskrivet/projicerat"),
      bullet("K\u00E4llmaterial och l\u00E4nkar (se bilaga 5)"),
      bullet("Datorer f\u00F6r skrivarbete"),
      bullet("Anteckningspapper f\u00F6r exit ticket"),
      bullet("Scaffolding-material (bilaga 6) \u2014 f\u00F6r l\u00E4raren att anv\u00E4nda vid behov"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("Perspektivanalysen tr\u00E4nar f\u00F6rm\u00E5gan att granska och v\u00E4rdera k\u00E4llor och argument fr\u00E5n olika akt\u00F6rer (m\u00E5l 1: E\u2013A)"),
      bullet("Skrivandet fr\u00E5n tre perspektiv tr\u00E4nar resonemang om AI-genererat inneh\u00E5lls p\u00E5verkan p\u00E5 informationslandskapet (m\u00E5l 2: E\u2013A)"),
      bullet("Att v\u00E4ga in teknikf\u00F6retag, stat och medborgare tr\u00E4nar f\u00F6rm\u00E5gan att analysera konspirationsteoriers spridning ur flera perspektiv (m\u00E5l 3: E\u2013A)"),
      bullet("Harvard-h\u00E4nvisningarna i texten tr\u00E4nar korrekt k\u00E4llhantering (m\u00E5l 4: E\u2013A)"),
      bullet("Metakognitiv exit ticket m\u00E4ter elevens f\u00F6rm\u00E5ga att granska och f\u00F6rb\u00E4ttra sitt eget resonemang (AI-02)"),
      spacer(),
      bodyText("Skrivformatet s\u00E4kerst\u00E4ller att alla elever m\u00E5ste bearbeta alla tre perspektiv. Rollkorten med f\u00E4rdiga argument ger E-elever en grund att st\u00E5 p\u00E5. Kravet p\u00E5 egen slutsats och k\u00E4llh\u00E4nvisning ger A-elever utrymme att visa nyanserat resonemang och s\u00E4ker k\u00E4llhantering."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Lektion 1\u20134 (alla verktyg, begrepp och muntlig argumentation anv\u00E4nds nu skriftligt)"),
      bullet("Lektion 4 (seminariet) \u2014 retrieval practice kopplar direkt till g\u00E5rdagens diskussion; samma tema (AI-driven desinformation)"),
      bullet("Lektion 2\u20133 (Harvard-systemet) \u2014 till\u00E4mpas nu i en l\u00E4ngre text"),
      bullet("N\u00E4sta lektion (6): Summativ skrivuppgift i SAMMA FORMAT \u2014 L5 \u00E4r dress rehearsal"),
      bullet("L\u00E4rarens feedback p\u00E5 L5-texterna hj\u00E4lper eleverna f\u00F6rbereda sig inf\u00F6r L6"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: ROLLKORT
      heading2("Bilaga 1: Rollkort \u2014 Tre positioner (skrivst\u00F6d)"),
      bodyText("Anv\u00E4nd rollkorten som st\u00F6d n\u00E4r du skriver. Varje kort inneh\u00E5ller position, argument och k\u00E4llor som du kan h\u00E4nvisa till i din text."),
      spacer(),

      // ROLLKORT A: TEKNIKFÖRETAGEN
      heading3("Rollkort A: Teknikf\u00F6retagen (Meta, Google, OpenAI)"),
      boldBodyText("Position: ", "Teknikf\u00F6retagen kan b\u00E4st l\u00F6sa problemet genom sj\u00E4lvreglering, innovation och bransch\u00F6verenskommelser. Lagstiftning \u00E4r f\u00F6r l\u00E5ngsam och riskerar att h\u00E4mma innovation."),
      spacer(),
      boldBodyText("Argument 1 \u2014 F\u00F6retagen har tekniken: ", "Bara teknikf\u00F6retagen har den tekniska kompetensen att bygga verktyg som k\u00E4nner igen och m\u00E4rker AI-genererat inneh\u00E5ll. OpenAI har inf\u00F6rt policyer mot deepfakes och st\u00F6djer NO FAKES Act. Meta och Google utvecklar AI-drivna system f\u00F6r att identifiera desinformation."),
      boldBodyText("K\u00E4lla: ", "OpenAI Usage Policies \u2014 openai.com/policies/usage-policies/"),
      spacer(),
      boldBodyText("Argument 2 \u2014 Lagstiftning \u00E4r f\u00F6r l\u00E5ngsam: ", "AI utvecklas s\u00E5 snabbt att lagar \u00E4r f\u00F6r\u00E5ldrade innan de tr\u00E4der i kraft. EU:s AI Act antogs 2024 men tr\u00E4der inte i full kraft f\u00F6rr\u00E4n 2026\u20132027. Under tiden har tekniken redan f\u00F6r\u00E4ndrats."),
      boldBodyText("K\u00E4lla: ", "EU AI Act \u2014 artificial-intelligence-act.com"),
      spacer(),
      boldBodyText("Argument 3 \u2014 Bransch\u00F6verenskommelser fungerar: ", "F\u00F6retagen har frivilligt anslutit sig till EU:s Code of Practice on Disinformation. Branschsamarbeten som C2PA (Coalition for Content Provenance and Authenticity) utvecklar standarder f\u00F6r inneh\u00E5llsm\u00E4rkning som g\u00E5r snabbare \u00E4n lagstiftning."),
      boldBodyText("K\u00E4lla: ", "EU Code of Practice on Disinformation \u2014 digital-strategy.ec.europa.eu/en/policies/code-practice-disinformation"),
      spacer(),
      boldBodyText("Argument 4 \u2014 Reglering hotar innovation: ", "Googles och Metas chefer har varnat f\u00F6r att Europas strikta reglering g\u00F6r att produkter f\u00F6rsenas eller utsl\u00E4ttas, vilket drabbar anv\u00E4ndarna. \u00D6verreglering riskerar att pressa AI-utvecklingen till l\u00E4nder med f\u00E4rre regler."),
      boldBodyText("K\u00E4lla: ", "CNBC \u2014 cnbc.com/2025/02/21/google-meta-execs-blast-europe-over-strict-ai-regulation.html"),

      new Paragraph({ children: [new PageBreak()] }),

      // ROLLKORT B: STATEN/EU
      heading3("Rollkort B: Staten/EU"),
      boldBodyText("Position: ", "Lagstiftning och statlig reglering \u00E4r den enda l\u00F6sningen som kan g\u00E4lla alla akt\u00F6rer lika. Utan bindande regler kan f\u00F6retagen v\u00E4lja att prioritera vinst \u00F6ver samh\u00E4llsansvar."),
      spacer(),
      boldBodyText("Argument 1 \u2014 Bindande regler g\u00E4ller alla: ", "EU:s AI Act kr\u00E4ver att leverant\u00F6rer av generativ AI m\u00E4rker AI-genererat inneh\u00E5ll och f\u00F6ljer transparenskrav. Digital Services Act (DSA) \u00E5l\u00E4gger stora plattformar att g\u00F6ra riskbed\u00F6mningar och motverka desinformation. Lagstiftning \u00E4r den enda mekanismen som \u00E4r juridiskt bindande."),
      boldBodyText("K\u00E4lla: ", "EU:s AI Act \u2014 digital-strategy.ec.europa.eu/en/policies/regulatory-framework-ai"),
      spacer(),
      boldBodyText("Argument 2 \u2014 Sj\u00E4lvreglering r\u00E4cker inte: ", "Meta avskaffade sina tredjepartsfaktagranskare i januari 2025 och ers\u00E4tte dem med \u201Ccommunity notes\u201D. Metas Responsible AI-enhet l\u00F6stes upp redan 2023. N\u00E4r f\u00F6retagen sj\u00E4lva best\u00E4mmer, v\u00E4ljer de det som \u00E4r billigast \u2014 inte det som skyddar demokratin."),
      boldBodyText("K\u00E4lla: ", "Meta\u2019s Fact-Checking Rollback \u2014 btlj.org/2025/05/metas-fact-checking-rollback-governance-free-speech-and-user-safety/"),
      spacer(),
      boldBodyText("Argument 3 \u2014 Demokrati kr\u00E4ver skydd: ", "EU:s European Democracy Shield (november 2025) \u00E4r en strategi f\u00F6r att skydda demokratin mot desinformation och utl\u00E4ndsk informationsp\u00E5verkan. AI-genererad desinformation anv\u00E4ndes i Ecuadors val 2025 och v\u00E4ntas \u00F6ka kraftigt i amerikanska mellanval 2026. Demokratier beh\u00F6ver aktiva skyddsmekanismer."),
      boldBodyText("K\u00E4lla: ", "European Democracy Shield \u2014 disinfo.eu/disinfo-update-13-01-2026"),
      spacer(),
      boldBodyText("Argument 4 \u2014 Ansvar kan utk\u00E4vas: ", "Med lagstiftning kan myndigheter utd\u00F6ma sanktioner mot f\u00F6retag som bryter mot reglerna. DSA kr\u00E4ver att stora plattformar g\u00F6r systematiska riskbed\u00F6mningar och revisioner. Utan lag finns inget straff."),
      boldBodyText("K\u00E4lla: ", "Digital Services Act \u2014 digital-strategy.ec.europa.eu/en/policies/digital-services-act"),

      new Paragraph({ children: [new PageBreak()] }),

      // ROLLKORT C: MEDBORGARNA
      heading3("Rollkort C: Medborgarna/civilsamh\u00E4llet"),
      boldBodyText("Position: ", "Varken lagstiftning eller sj\u00E4lvreglering r\u00E4cker om m\u00E4nniskor inte kan granska information sj\u00E4lva. Utbildning, mediekunnighet och ett starkt civilsamh\u00E4lle \u00E4r den enda h\u00E5llbara l\u00F6sningen."),
      spacer(),
      boldBodyText("Argument 1 \u2014 Utbildning \u00E4r h\u00E5llbart: ", "Lagar och tekniska filter kan kringg\u00E5s. En medborgare med mediekunnighet b\u00E4r med sig verktyget \u00F6verallt, i alla sammanhang, p\u00E5 alla plattformar. World Economic Forum framh\u00E5ller att mediekunnighet m\u00E5ste integreras i utbildningssystem, arbetsplatser och offentliga kampanjer."),
      boldBodyText("K\u00E4lla: ", "WEF \u2014 Rethinking Media Literacy 2025 \u2014 reports.weforum.org/docs/WEF_Rethinking_Media_Literacy_2025.pdf"),
      spacer(),
      boldBodyText("Argument 2 \u2014 Civilsamh\u00E4llet \u00E4r oberoende: ", "Organisationer som K\u00E4llkritikbyr\u00E5n, Internetstiftelsen och faktagranskningsn\u00E4tverk arbetar oberoende fr\u00E5n b\u00E5de stat och f\u00F6retag. De har ingen vinst att skydda och ingen politisk agenda \u2014 bara medborgarnas intresse."),
      boldBodyText("K\u00E4lla: ", "K\u00E4llkritikbyr\u00E5n \u2014 kallkritikbyran.se"),
      spacer(),
      boldBodyText("Argument 3 \u2014 B\u00E5de lag och teknik har brister: ", "EU:s AI Act tar \u00E5r att implementera. Teknikf\u00F6retagen har ekonomiska incitament att h\u00E5lla anv\u00E4ndare engagerade, inte kritiskt t\u00E4nkande. Bara 17% av UNESCO:s medlemsl\u00E4nder har en sj\u00E4lvst\u00E4ndig policy f\u00F6r mediekunnighet \u2014 h\u00E4r finns enormt utrymme f\u00F6r f\u00F6rb\u00E4ttring."),
      boldBodyText("K\u00E4lla: ", "UNESCO \u2014 weforum.org/stories/2025/07/disinformation-media-and-information-literacy/"),
      spacer(),
      boldBodyText("Argument 4 \u2014 Hela samh\u00E4llet m\u00E5ste engageras: ", "Forskning visar att en \u201Cwhole-of-society approach\u201D d\u00E4r skolor, bibliotek, religi\u00F6sa grupper, lokala medier och civilsamh\u00E4lle samverkar \u00E4r mest effektiv mot desinformation. Det r\u00E4cker inte med en lag eller ett filter \u2014 det kr\u00E4vs en kulturf\u00F6r\u00E4ndring."),
      boldBodyText("K\u00E4lla: ", "Frontiers \u2014 frontiersin.org/journals/communication/articles/10.3389/fcomm.2026.1771055/full"),

      // BILAGA 2: BEDÖMNINGSSTÖD
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Bed\u00F6mningsst\u00F6d f\u00F6r l\u00E4raren"),
      bodyText("Anv\u00E4nd detta st\u00F6d f\u00F6r att ge formativ feedback p\u00E5 elevernas skriftliga perspektivanalys. OBS: L5 \u00E4r formativ \u2014 inget betyg s\u00E4tts. Fokusera p\u00E5 framf\u00F6r\u00E5triktad feedback inf\u00F6r L6."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 2509, 2509, 2508],
        rows: [
          new TableRow({
            children: [
              headerCell("Kriterium", 1500),
              headerCell("E-niv\u00E5", 2509),
              headerCell("C-niv\u00E5", 2509),
              headerCell("A-niv\u00E5", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("Perspektiv-bredd", 1500),
              cell("N\u00E4mner alla tre perspektiv p\u00E5 ett enkelt s\u00E4tt", 2509),
              cell("Beskriver alla tre perspektiv med v\u00E4lgrundade argument och konkreta exempel", 2509),
              cell("Analyserar alla tre perspektiv med nyanserade argument som visar p\u00E5 samband och sp\u00E4nningar mellan dem", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("Argumentation", 1500),
              cell("Framf\u00F6r enkla argument som \u00E5terger rollkortens inneh\u00E5ll", 2509),
              cell("Framf\u00F6r v\u00E4lgrundade argument med egna formuleringar och konkreta h\u00E4nvisningar", 2509),
              cell("Framf\u00F6r v\u00E4lgrundade och nyanserade argument som visar p\u00E5 komplexa samband", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("Egen slutsats", 1500),
              cell("Tar st\u00E4llning p\u00E5 ett enkelt s\u00E4tt", 2509),
              cell("Tar motiverad st\u00E4llning med st\u00F6d i argumenten", 2509),
              cell("Tar v\u00E4lmotiverad st\u00E4llning som visar p\u00E5 komplexiteten och v\u00E4ger perspektiven mot varandra", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00E4llanv\u00E4ndning", 1500),
              cell("H\u00E4nvisar till minst 2 k\u00E4llor p\u00E5 ett enkelt s\u00E4tt", 2509),
              cell("Anv\u00E4nder minst 2 k\u00E4llor f\u00F6r att underbygga argument p\u00E5 ett relevant s\u00E4tt", 2509),
              cell("Anv\u00E4nder minst 3 k\u00E4llor kritiskt och v\u00E4ger dem mot varandra", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("K\u00E4llh\u00E4nvisning (Harvard)", 1500),
              cell("K\u00E4llh\u00E4nvisar med viss s\u00E4kerhet \u2014 formatet \u00E4r igenk\u00E4nnbart men inneh\u00E5ller brister", 2509),
              cell("K\u00E4llh\u00E4nvisar korrekt med parenteser i text och referenslista", 2509),
              cell("K\u00E4llh\u00E4nvisar med s\u00E4kerhet, konsekvent och korrekt genom hela texten", 2508),
            ],
          }),
          new TableRow({
            children: [
              cell("Demokratikoppling", 1500),
              cell("N\u00E4mner demokrati eller samh\u00E4llsp\u00E5verkan i allm\u00E4nna termer", 2509),
              cell("Resonerar om konkreta konsekvenser f\u00F6r demokratin", 2509),
              cell("Analyserar komplexa samband mellan desinformation, makt och demokrati", 2508),
            ],
          }),
        ],
      }),

      spacer(),
      heading3("Observationspunkter under skrivarbetet"),
      bullet("Vilka elever anv\u00E4nder rollkorten som st\u00F6d? Vilka skriver fritt?"),
      bullet("Anv\u00E4nder eleverna k\u00E4llor i sin text eller bara \u00E5sikter?"),
      bullet("Lyckas eleverna skriva fr\u00E5n alla tre perspektiv eller fastnar de i ett?"),
      bullet("Hur ser k\u00E4llh\u00E4nvisningarna ut? K\u00E4nner de igen Harvard-formatet?"),
      bullet("G\u00F6r n\u00E5gon kopplingar till tidigare lektioner (k\u00E4llkritiska grundfr\u00E5gor, konspirationsteoriers mekanismer)?"),

      // BILAGA 3: EXIT TICKET
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 3: Exit ticket \u2014 Lektion 5 (metakognitiv)"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "L\u00E4s igenom din perspektivanalys. Markera en mening d\u00E4r du anv\u00E4nde ett k\u00E4llkritiskt verktyg och en mening d\u00E4r du kunde ha st\u00E4rkt din argumentation. Vad skulle du g\u00F6ra annorlunda?"),
      spacer(),
      bodyText("Denna exit ticket \u00E4r metakognitiv: eleven reflekterar \u00F6ver sin egen text och identifierar b\u00E5de styrkor och f\u00F6rb\u00E4ttringsomr\u00E5den."),
      spacer(),
      bodyText("F\u00F6r h\u00F6gre betyg:"),
      bullet("Visa p\u00E5 hur du anv\u00E4nde ett specifikt k\u00E4llkritiskt verktyg (SIFT, grundfr\u00E5gor, lateral reading)"),
      bullet("F\u00F6rklara KONKRET vad du skulle \u00E4ndra och HUR det skulle st\u00E4rka argumentationen"),
      bullet("Koppla till k\u00E4llor: Skulle en annan k\u00E4lla ha gjort argumentet starkare?"),
      spacer(),
      boldBodyText("L\u00E4randem\u00E5l som m\u00E4ts: ", "M\u00E5l 1 (k\u00E4llkritiska verktyg) och m\u00E5l 4 (k\u00E4llh\u00E4nvisning)"),
      boldBodyText("Retrieval review i L6: ", "Eleverna f\u00E5r se ett kort textexempel och identifierar styrkor och svagheter i argumentationen \u2014 som uppv\u00E4rmning inf\u00F6r den summativa skrivuppgiften."),

      // BILAGA 4: SKRIVUPPGIFT (sammanfattning för eleverna)
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 4: Skrivuppgift \u2014 Perspektivanalys"),
      spacer(),
      boldBodyText("Fr\u00E5ga: ", "Vem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation?"),
      spacer(),
      bodyText("Skriv en text p\u00E5 300\u2013400 ord d\u00E4r du:"),
      bullet("Presenterar alla tre perspektiv: teknikf\u00F6retagen, staten/EU och medborgarna/civilsamh\u00E4llet"),
      bullet("Anv\u00E4nder minst 2 k\u00E4llor med korrekta Harvard-h\u00E4nvisningar"),
      bullet("Landar i en egen slutsats d\u00E4r du tar st\u00E4llning"),
      spacer(),
      bodyText("F\u00F6rslag p\u00E5 textstruktur:"),
      bullet("Inledning: Presentera fr\u00E5gan kort (1\u20132 meningar)"),
      bullet("Perspektiv 1 \u2014 Teknikf\u00F6retagen: Vad \u00E4r deras argument?"),
      bullet("Perspektiv 2 \u2014 Staten/EU: Vad \u00E4r deras argument?"),
      bullet("Perspektiv 3 \u2014 Medborgarna: Vad \u00E4r deras argument?"),
      bullet("Slutsats: Vem b\u00E4r st\u00F6rst ansvar? Varf\u00F6r? Eller beh\u00F6vs alla tre?"),
      bullet("Referenslista"),
      spacer(),
      boldBodyText("OBS: ", "Detta \u00E4r en \u00F6vning inf\u00F6r lektion 6. Du f\u00E5r feedback av l\u00E4raren p\u00E5 din text \u2014 anv\u00E4nd den f\u00F6r att f\u00F6rbereda dig."),
      spacer(),
      boldBodyText("Bed\u00F6mningsfokus:", ""),
      bullet("Perspektivbredd (m\u00E5l 2\u20133): Hur v\u00E4l presenterar och analyserar du alla tre perspektiv?"),
      bullet("Argumentation (m\u00E5l 1): Hur v\u00E4lgrundade \u00E4r dina argument?"),
      bullet("K\u00E4llh\u00E4nvisning (m\u00E5l 4): Hur korrekt h\u00E4nvisar du enligt Harvard?"),
      spacer(),
      bodyText("Tips: Anv\u00E4nd rollkorten som st\u00F6d. K\u00E4llorna p\u00E5 rollkorten kan du h\u00E4nvisa till i din text."),

      // BILAGA 5: KÄLLMATERIAL OCH LÄNKAR
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 5: K\u00E4llmaterial och l\u00E4nkar"),
      bodyText("Samlade l\u00E4nkar f\u00F6r rollkorten och f\u00F6rdjupning."),
      spacer(),

      heading3("K\u00E4llor f\u00F6r Rollkort A \u2014 Teknikf\u00F6retagen"),
      boldBodyText("OpenAI Usage Policies: ", "openai.com/policies/usage-policies/"),
      boldBodyText("OpenAI mot deepfakes (CNBC): ", "cnbc.com/2025/10/20/open-ai-sora-bryan-cranston-sag-aftra.html"),
      boldBodyText("EU AI Act \u00F6versikt: ", "artificial-intelligence-act.com"),
      boldBodyText("EU Code of Practice on Disinformation: ", "digital-strategy.ec.europa.eu/en/policies/code-practice-disinformation"),
      boldBodyText("Google och Meta kritiserar EU-reglering (CNBC): ", "cnbc.com/2025/02/21/google-meta-execs-blast-europe-over-strict-ai-regulation.html"),
      spacer(),

      heading3("K\u00E4llor f\u00F6r Rollkort B \u2014 Staten/EU"),
      boldBodyText("EU:s AI Act: ", "digital-strategy.ec.europa.eu/en/policies/regulatory-framework-ai"),
      boldBodyText("Digital Services Act (DSA): ", "digital-strategy.ec.europa.eu/en/policies/digital-services-act"),
      boldBodyText("Meta\u2019s Fact-Checking Rollback: ", "btlj.org/2025/05/metas-fact-checking-rollback-governance-free-speech-and-user-safety/"),
      boldBodyText("Meta bort fr\u00E5n faktagranskning (Northeastern): ", "news.northeastern.edu/2025/01/07/meta-fact-checking-policy-change/"),
      boldBodyText("European Democracy Shield: ", "disinfo.eu/disinfo-update-13-01-2026"),
      boldBodyText("AI-desinformation i Ecuadors val 2025: ", "frontiersin.org/journals/political-science/articles/10.3389/fpos.2025.1624206/full"),
      boldBodyText("AI-desinformation inf\u00F6r USA:s mellanval 2026: ", "witf.org/2025/12/20/voters-to-face-unprecedented-levels-of-ai-generated-misinformation-in-2026/"),
      boldBodyText("Deepfakes och EU AI Act (Columbia): ", "cjel.law.columbia.edu/preliminary-reference/2024/deepfake-deep-trouble-the-european-ai-act-and-the-fight-against-ai-generated-misinformation/"),
      spacer(),

      heading3("K\u00E4llor f\u00F6r Rollkort C \u2014 Medborgarna/civilsamh\u00E4llet"),
      boldBodyText("WEF Rethinking Media Literacy 2025: ", "reports.weforum.org/docs/WEF_Rethinking_Media_Literacy_2025.pdf"),
      boldBodyText("WEF: Mediekunnighet i AI-eran: ", "weforum.org/stories/2025/10/media-information-literacy-ai/"),
      boldBodyText("WEF: Mediekunnighet i desinformations\u00E5ldern: ", "weforum.org/stories/2025/07/disinformation-media-and-information-literacy/"),
      boldBodyText("K\u00E4llkritikbyr\u00E5n: ", "kallkritikbyran.se"),
      boldBodyText("Internetstiftelsen \u2014 Svenskarna och internet 2025: ", "svenskarnaochinternet.se/rapporter/svenskarna-och-internet-2025/"),
      boldBodyText("Internetkunskap \u2014 k\u00E4llkritik: ", "internetkunskap.se/kallkritik/"),
      boldBodyText("Kritisk mediekunnighet och AI (Frontiers): ", "frontiersin.org/journals/communication/articles/10.3389/fcomm.2026.1771055/full"),
      boldBodyText("Mediekunnighet i utbildning (Frontiers): ", "frontiersin.org/journals/human-dynamics/articles/10.3389/fhumd.2025.1608911/full"),
      spacer(),

      heading3("\u00D6vriga resurser"),
      boldBodyText("AI-driven desinformation \u2014 policyrekommendationer (Frontiers): ", "frontiersin.org/journals/artificial-intelligence/articles/10.3389/frai.2025.1569115/full"),
      boldBodyText("EU DisinfoLab \u2014 AI Disinfo Hub: ", "disinfo.eu/ai-disinfo-hub/"),
      boldBodyText("DSA och desinformation f\u00F6rklarat: ", "sirenassociates.com/content/tackling-disinformation-the-eu-digital-services-act-explained"),
      boldBodyText("Scientific American \u2014 lagar mot deepfakes: ", "scientificamerican.com/article/we-need-laws-to-stop-ai-generated-deepfakes/"),

      // BILAGA 6: SCAFFOLDING PÅ BEGÄRAN
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 6: Scaffolding p\u00E5 beg\u00E4ran"),
      bodyText("Anv\u00E4nd detta material ENBART om du bed\u00F6mer att elever beh\u00F6ver st\u00F6d. Dela inte ut som standard. Skriv upp p\u00E5 tavlan eller ge individuellt till elever som \u00E4r fast."),
      spacer(),

      heading3("Meningsstartare f\u00F6r skrivande"),
      bullet("\u201CTeknikf\u00F6retagen anser att... eftersom...\u201D"),
      bullet("\u201CStaten/EU argumenterar f\u00F6r att... d\u00E5...\u201D"),
      bullet("\u201CMedborgarna/civilsamh\u00E4llet menar att... d\u00E4rf\u00F6r att...\u201D"),
      bullet("\u201CMin slutsats \u00E4r att... eftersom...\u201D"),
      bullet("\u201C\u00C5 ena sidan h\u00E4vdar [akt\u00F6r] att... \u00E5 andra sidan menar [akt\u00F6r] att...\u201D"),
      bullet("\u201CDet st\u00E4rkaste argumentet f\u00F6r att [akt\u00F6r] b\u00E4r st\u00F6rst ansvar \u00E4r...\u201D"),
      spacer(),

      heading3("Textstruktur (steg f\u00F6r steg)"),
      bullet("Inledning: Presentera fr\u00E5gan kort (1\u20132 meningar)"),
      bullet("Perspektiv 1 \u2014 Teknikf\u00F6retagen: Deras argument och k\u00E4lla"),
      bullet("Perspektiv 2 \u2014 Staten/EU: Deras argument och k\u00E4lla"),
      bullet("Perspektiv 3 \u2014 Medborgarna: Deras argument och k\u00E4lla"),
      bullet("Slutsats: Vilken akt\u00F6r b\u00E4r st\u00F6rst ansvar? Varf\u00F6r? Eller beh\u00F6vs alla tre?"),
      bullet("Referenslista"),
      spacer(),

      heading3("Harvard-p\u00E5minnelse"),
      bodyText("I texten: (F\u00F6rfattare, \u00E5r). Exempel: (OpenAI, 2024)"),
      bodyText("I referenslistan: F\u00F6rfattare (\u00E5r). Titel. K\u00E4lla. URL."),
      bodyText("Exempel: OpenAI (2024). Usage Policies. openai.com/policies/usage-policies/"),
    ],
  }],
});

// ============================================================
// PPTX - Lektion 5 presentation
// ============================================================

function generatePptx() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Cornelius Agent";
  pres.title = "Lektion 5: Perspektiven krockar";

  // Color palette - Midnight Executive (samma som L2/L3)
  const DARK_BG = "1B2A4A";
  const MED_BG = "243656";
  const ACCENT = "4DA8DA";
  const ACCENT2 = "6BC5D2";
  const TEXT_WHITE = "FFFFFF";
  const TEXT_LIGHT = "CBD5E1";
  const TEXT_MUTED = "94A3B8";

  // --- Slide 1: Titel ---
  const s1 = pres.addSlide();
  s1.background = { color: DARK_BG };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.5, w: 10, h: 1.125, fill: { color: MED_BG } });
  s1.addText("Lektion 5", { x: 0.8, y: 1.2, w: 8.4, h: 0.8, fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, charSpacing: 6 });
  s1.addText("Perspektiven krockar", { x: 0.8, y: 2.0, w: 8.4, h: 1.2, fontSize: 40, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s1.addText("Skriftlig perspektivanalys - dress rehearsal inf\u00F6r L6", { x: 0.8, y: 3.2, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT });
  s1.addText("Samh\u00E4llskunskap 3 / Internationella relationer", { x: 0.8, y: 4.7, w: 8.4, h: 0.5, fontSize: 13, fontFace: "Arial", color: TEXT_MUTED });

  // --- Slide 2: Retrieval practice ---
  const s2 = pres.addSlide();
  s2.background = { color: DARK_BG };
  s2.addText("\u00C5terkallelse fr\u00E5n L4", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s2.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 8.4, h: 2.2, fill: { color: MED_BG } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 0.08, h: 2.2, fill: { color: ACCENT } });
  s2.addText([
    { text: "Skriv ner det starkaste argumentet du h\u00F6rde ig\u00E5r", options: { fontSize: 22, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "- oavsett vilken sida det kom ifr\u00E5n.", options: { fontSize: 22, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Enskilt skrivande. 3 minuter.", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT } },
  ], { x: 1.1, y: 1.4, w: 7.9, h: 2.0 });

  s2.addText("Sedan delar 3-4 elever muntligt.", { x: 0.8, y: 4.9, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true });

  // --- Slide 3: Skrivuppgiften ---
  const s3 = pres.addSlide();
  s3.background = { color: DARK_BG };
  s3.addText("Skrivuppgift: Perspektivanalys", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s3.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, fill: { color: MED_BG } });
  s3.addText("Vem b\u00E4r ansvaret f\u00F6r att stoppa AI-driven desinformation?", { x: 1.1, y: 1.2, w: 7.8, h: 1.0, fontSize: 20, fontFace: "Arial", color: ACCENT2, bold: true, valign: "middle" });

  const expectations = [
    { label: "Omfattning", desc: "300-400 ord" },
    { label: "Perspektiv", desc: "Alla tre: teknikf\u00F6retag, stat/EU, medborgare" },
    { label: "Slutsats", desc: "Din egen st\u00E4llning - vem b\u00E4r st\u00F6rst ansvar?" },
    { label: "K\u00E4llor", desc: "Minst 2 k\u00E4llor med Harvard-h\u00E4nvisning" },
  ];

  expectations.forEach((item, i) => {
    const yPos = 2.5 + i * 0.7;
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.55, fill: { color: MED_BG } });
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.08, h: 0.55, fill: { color: ACCENT } });
    s3.addText(item.label, { x: 1.1, y: yPos, w: 2.2, h: 0.55, fontSize: 16, fontFace: "Arial", color: ACCENT, bold: true, valign: "middle", margin: 0 });
    s3.addText(item.desc, { x: 3.3, y: yPos, w: 5.7, h: 0.55, fontSize: 15, fontFace: "Arial", color: TEXT_LIGHT, valign: "middle", margin: 0 });
  });

  s3.addText("Dress rehearsal inf\u00F6r L6 - samma format, l\u00E4gre insats!", { x: 0.8, y: 5.0, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: ACCENT, italic: true, align: "center" });

  // --- Slide 4: Harvard-påminnelse ---
  const s4 = pres.addSlide();
  s4.background = { color: DARK_BG };
  s4.addText("Harvard-p\u00E5minnelse", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s4.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  // I texten
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 4.0, h: 2.5, fill: { color: MED_BG } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 0.08, h: 2.5, fill: { color: ACCENT } });
  s4.addText([
    { text: "I texten", options: { fontSize: 20, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "(F\u00F6rfattare, \u00E5r)", options: { fontSize: 18, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Exempel:", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "(OpenAI, 2024)", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, italic: true } },
  ], { x: 1.1, y: 1.4, w: 3.5, h: 2.3 });

  // I referenslistan
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.0, h: 2.5, fill: { color: MED_BG } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 0.08, h: 2.5, fill: { color: ACCENT } });
  s4.addText([
    { text: "I referenslistan", options: { fontSize: 20, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "F\u00F6rfattare (\u00E5r). Titel. URL.", options: { fontSize: 18, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Exempel:", options: { fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, breakLine: true } },
    { text: "OpenAI (2024). Usage Policies. openai.com/...", options: { fontSize: 14, fontFace: "Arial", color: TEXT_LIGHT, italic: true } },
  ], { x: 5.5, y: 1.4, w: 3.5, h: 2.3 });

  s4.addText("Harvard-guiden fr\u00E5n L2/L3 finns tillg\u00E4nglig.", { x: 0.8, y: 4.9, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true, align: "center" });

  // --- Slide 5: Textstruktur ---
  const s5 = pres.addSlide();
  s5.background = { color: DARK_BG };
  s5.addText("Textstruktur", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s5.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  const struktur = [
    { nr: "1", text: "Inledning - presentera fr\u00E5gan kort" },
    { nr: "2", text: "Teknikf\u00F6retagen - deras argument och k\u00E4lla" },
    { nr: "3", text: "Staten/EU - deras argument och k\u00E4lla" },
    { nr: "4", text: "Medborgarna - deras argument och k\u00E4lla" },
    { nr: "5", text: "Slutsats - vem b\u00E4r st\u00F6rst ansvar?" },
    { nr: "6", text: "Referenslista" },
  ];

  struktur.forEach((item, i) => {
    const yPos = 1.2 + i * 0.65;
    s5.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.5, fill: { color: MED_BG } });
    s5.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.08, h: 0.5, fill: { color: ACCENT } });
    s5.addText(item.nr, { x: 1.1, y: yPos, w: 0.5, h: 0.5, fontSize: 18, fontFace: "Arial", color: ACCENT, bold: true, valign: "middle", margin: 0 });
    s5.addText(item.text, { x: 1.7, y: yPos, w: 7.3, h: 0.5, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, valign: "middle", margin: 0 });
  });

  s5.addText("Anv\u00E4nd rollkorten som st\u00F6d. K\u00E4llorna finns p\u00E5 rollkorten och i bilaga 5.", { x: 0.8, y: 5.0, w: 8.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: TEXT_MUTED, italic: true, align: "center" });

  // --- Slide 6: Exit ticket ---
  const s6 = pres.addSlide();
  s6.background = { color: DARK_BG };
  s6.addText("Exit ticket - metakognitiv reflektion", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s6.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 8.4, h: 2.8, fill: { color: MED_BG } });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 0.08, h: 2.8, fill: { color: ACCENT } });
  s6.addText([
    { text: "L\u00E4s igenom din perspektivanalys.", options: { fontSize: 20, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Markera en mening d\u00E4r du anv\u00E4nde ett k\u00E4llkritiskt verktyg", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Markera en mening d\u00E4r du kunde ha st\u00E4rkt din argumentation", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "Vad skulle du g\u00F6ra annorlunda?", options: { bullet: true, fontSize: 17, fontFace: "Arial", color: ACCENT2, bold: true } },
  ], { x: 1.2, y: 1.5, w: 7.6, h: 2.4 });

  s6.addText("Reflektera \u00F6ver ditt eget resonemang - tr\u00E4ning inf\u00F6r L6.", { x: 0.8, y: 4.6, w: 8.4, h: 0.6, fontSize: 16, fontFace: "Arial", color: ACCENT, italic: true, align: "center" });

  // --- Slide 7: Framåtkoppling till L6 ---
  const s7 = pres.addSlide();
  s7.background = { color: DARK_BG };
  s7.addText("N\u00E4sta lektion: L6", { x: 0.8, y: 0.3, w: 8.4, h: 0.6, fontSize: 28, fontFace: "Arial", color: TEXT_WHITE, bold: true });
  s7.addShape(pres.shapes.LINE, { x: 0.8, y: 0.95, w: 2, h: 0, line: { color: ACCENT, width: 3 } });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 8.4, h: 3.2, fill: { color: MED_BG } });
  s7.addText([
    { text: "Summativ skrivuppgift", options: { fontSize: 22, fontFace: "Arial", color: ACCENT2, bold: true, breakLine: true } },
    { text: "Samma format som idag - perspektivanalys med k\u00E4llh\u00E4nvisning", options: { fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Vad \u00E4r annorlunda i L6?", options: { fontSize: 18, fontFace: "Arial", color: TEXT_WHITE, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Texten \u00E4r ert bed\u00F6mningsunderlag", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "Ni f\u00E5r feedback p\u00E5 dagens text innan dess", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: TEXT_LIGHT, breakLine: true } },
    { text: "Anv\u00E4nd feedbacken f\u00F6r att f\u00F6rbereda er", options: { bullet: true, fontSize: 16, fontFace: "Arial", color: ACCENT2, bold: true } },
  ], { x: 1.2, y: 1.5, w: 7.6, h: 2.8 });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fill: { color: MED_BG } });
  s7.addText("Ni har redan \u00F6vat p\u00E5 argumenten, Harvard och textstrukturen - det \u00E4r precis det ni g\u00F6r i L6!", { x: 0.8, y: 4.7, w: 8.4, h: 0.6, fontSize: 15, fontFace: "Arial", color: ACCENT, align: "center", valign: "middle" });

  return pres;
}

// ============================================================
// GENERATE BOTH FILES
// ============================================================

const OUTPUT_DIR = "/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier";

async function main() {
  // Generate DOCX
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(`${OUTPUT_DIR}/lektion-5.docx`, buffer);
  console.log("lektion-5.docx skapad!");

  // Generate PPTX
  const pres = generatePptx();
  await pres.writeFile({ fileName: `${OUTPUT_DIR}/lektion-5.pptx` });
  console.log("lektion-5.pptx skapad!");
}

main().catch(err => {
  console.error("Fel vid generering:", err);
  process.exit(1);
});
