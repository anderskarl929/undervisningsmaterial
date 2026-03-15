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
        children: [new TextRun({ text: "Lektion 7: Metakognitiv reflektion och sj\u00E4lvbed\u00F6mning", font: "Arial" })],
      }),
      bodyText("Syntes, metakognitiv reflektion och sj\u00E4lvbed\u00F6mning"),
      spacer(),
      boldBodyText("Kurs: ", "Samh\u00E4llskunskap 3 / Internationella relationer"),
      boldBodyText("Moment: ", "K\u00E4llkritik \u2014 AI-genererat inneh\u00E5ll och konspirationsteorier"),
      boldBodyText("Lektionsl\u00E4ngd: ", "65 minuter"),
      boldBodyText("Grupp: ", "MSA23"),

      // LÄRANDEMÅL
      heading2("L\u00E4randem\u00E5l f\u00F6r lektionen"),
      bodyText("Eleven reflekterar \u00F6ver sitt l\u00E4rande och visar f\u00F6rst\u00E5else f\u00F6r momentets centrala inneh\u00E5ll. Den h\u00E4r lektionen \u00E4r inte summativ \u2014 bed\u00F6mningen gjordes i lektion 6. Fokus ligger p\u00E5 syntes, metakognition och framtida till\u00E4mpning."),
      spacer(),
      bullet("Granska och v\u00E4rdera k\u00E4llor med hj\u00E4lp av de fyra k\u00E4llkritiska grundfr\u00E5gorna (m\u00E5l 1)"),
      bullet("Resonera om hur AI-genererat inneh\u00E5ll p\u00E5verkar m\u00F6jligheten att granska information (m\u00E5l 2)"),
      bullet("Resonera om hur konspirationsteorier uppst\u00E5r, sprids och p\u00E5verkar samh\u00E4llet (m\u00E5l 3)"),

      // CENTRALT INNEHÅLL
      heading2("Centralt inneh\u00E5ll (Gy11)"),
      bullet("K\u00E4llkritisk granskning, tolkning och v\u00E4rdering av information fr\u00E5n olika k\u00E4llor och medier i digital och annan form i arbetet med komplexa samh\u00E4llsfr\u00E5gor."),

      // FÖRBEREDELSE
      heading2("F\u00F6rberedelse"),
      bullet("F\u00F6rbered \u201C\u00C5terblick\u201D-bildspelet: samma tre bilder fr\u00E5n lektion 1 (p\u00E5ven i dunjacka, Pentagon-explosionen, \u00E4kta nyhetsbild)"),
      bullet("Skriv ut eller dela digitalt: reflektionsformul\u00E4ret (se bilaga 1)"),
      bullet("Skriv ut eller dela digitalt: sj\u00E4lvv\u00E4rderingsmatrisen (se bilaga 2)"),
      bullet("Ha l\u00E4randem\u00E5len synliga p\u00E5 tavlan under hela lektionen"),
      bullet("F\u00F6rbered en kort sammanfattning av momentets resa (lektion 1\u20138) f\u00F6r den muntliga \u00E5terblicken"),

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
              cell("0\u201312 min", 1000),
              cell("Retrieval review", 1300),
              cell("\u00C5terblick: \u201CKan du lita p\u00E5 det du ser?\u201D", 2200),
              cell("Visa samma tre bilder fr\u00E5n lektion 1. Eleverna r\u00F6star igen: \u00E4kta eller fejk? J\u00E4mf\u00F6r med f\u00F6rsta g\u00E5ngen. Muntlig diskussion: \u201CVad g\u00F6r ni annorlunda nu? Vilka verktyg har ni?\u201D Landa i: \u201CFr\u00E5n gissningar till systematisk granskning.\u201D \u00C5terkoppla till L6:s skrivuppgift.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("12\u201315 min", 1000),
              cell("Instruktion", 1300),
              cell("Momentets resa", 2200),
              cell("Kort muntlig sammanfattning av momentet (lektion 1\u20137). Visa progressionen: verktyg \u2192 till\u00E4mpning \u2192 analys. Koppla fram\u00E5t till L8 seminarie-examination.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("15\u201330 min", 1000),
              cell("Guidad \u00F6vning", 1300),
              cell("Syntesreflektion", 2200),
              cell("L\u00E5gstakes skrivande eller muntlig reflektion: \u201CSkiv fritt: Vad tar du med dig? Hur t\u00E4nker du annorlunda nu j\u00E4mf\u00F6rt med f\u00F6re momentet?\u201D Ingen bed\u00F6mning \u2014 fokus p\u00E5 \u00E4rlig reflektion. P\u00E5minnelse: anv\u00E4nd tredjepersons-framing n\u00E4r ni resonerar om konspirationsteorier.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("30\u201340 min", 1000),
              cell("Sj\u00E4lvst\u00E4ndig \u00F6vning", 1300),
              cell("Par-samtal + helklass", 2200),
              cell("Par-samtal (5 min): dela det viktigaste fr\u00E5n syntesreflektionen. Helklassdiskussion (5 min): lyft m\u00F6nster och \u00F6verraskningar. \u201CVad har ni l\u00E4rt er som ni inte visste f\u00F6re det h\u00E4r momentet?\u201D", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("40\u201350 min", 1000),
              cell("Sj\u00E4lvst\u00E4ndig \u00F6vning", 1300),
              cell("Sj\u00E4lvv\u00E4rderingsmatris", 2200),
              cell("Eleverna fyller i sj\u00E4lvv\u00E4rderingsmatrisen (E/C/A per m\u00E5l). Ingen bed\u00F6mning \u2014 \u00E4rlig sj\u00E4lvuppskattning. L\u00E4raren samlar in.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("50\u201358 min", 1000),
              cell("Exit ticket", 1300),
              cell("Gemensam avslutning", 2200),
              cell("\u201CVad tar ni med er fr\u00E5n det h\u00E4r momentet?\u201D Samla r\u00F6ster fr\u00E5n klassen. Sammanfatta p\u00E5 tavlan. Knyt ihop med l\u00E4randem\u00E5len.", 4526),
            ],
          }),
          new TableRow({
            children: [
              cell("58\u201365 min", 1000),
              cell("F\u00F6rhandsvisning", 1300),
              cell("Fram\u00E5t: era nya verktyg", 2200),
              cell("\u201CNi har nu en verktygsl\u00E5da f\u00F6r att navigera informationslandskapet. Anv\u00E4nd den \u2014 i skolan, i nyhetsfl\u00F6det, i samtal. Var den som st\u00E4ller fr\u00E5gorna.\u201D Koppla till det avslutande arbetet. Tack och avslut.", 4526),
            ],
          }),
        ],
      }),

      // LÄRARINSTRUKTIONER
      heading2("L\u00E4rarinstruktioner"),

      heading3("Retrieval review (0\u201312 min)"),
      italicText("S\u00E4g: \u201CI lektion 6 skrev ni en perspektivanalys. Idag blickar vi tillbaka p\u00E5 hela momentet. Vi b\u00F6rjade med tre bilder och en fr\u00E5ga: Kan du lita p\u00E5 det du ser? Nu, sju lektioner senare, ska vi testa samma sak igen.\u201D"),
      spacer(),
      bodyText("Visa samma tre bilder fr\u00E5n lektion 1:"),
      bullet("Bild 1: \u00C4kta nyhetsbild (t.ex. SVT-bild fr\u00E5n en valr\u00F6relse)"),
      bullet("Bild 2: P\u00E5ven i vit dunjacka (AI-genererad, Midjourney, mars 2023)"),
      bullet("Bild 3: AI-genererad bild av explosion vid Pentagon (maj 2023)"),
      spacer(),
      bodyText("L\u00E5t eleverna r\u00F6sta igen. F\u00F6rv\u00E4nta att de nu \u00E4r mer tr\u00E4ffs\u00E4kra \u2014 och viktigare: att de kan f\u00F6rklara VARF\u00D6R."),
      spacer(),
      italicText("Fr\u00E5ga: \u201CF\u00F6rra g\u00E5ngen gissade ni. Vad g\u00F6r ni annorlunda nu? Vilka verktyg har ni som ni inte hade d\u00E5?\u201D"),
      bodyText("Samla 4\u20135 svar muntligt. Skriv nyckelord p\u00E5 tavlan (t.ex. \u201Cgrundfrp\u00E5gorna\u201D, \u201Cspridningsmekanismer\u201D, \u201Ccui bono\u201D). Ge den h\u00E4r diskussionen tid \u2014 det \u00E4r h\u00E4r progressionen syns."),
      spacer(),
      italicText("Fr\u00E5ga vidare: \u201CHur k\u00E4nns det att se bilderna nu j\u00E4mf\u00F6rt med f\u00F6rsta g\u00E5ngen? Vad har f\u00F6r\u00E4ndrats?\u201D"),
      spacer(),
      italicText("Landa i: \u201CP\u00E5 sju lektioner har ni g\u00E5tt fr\u00E5n gissningar till ett systematiskt s\u00E4tt att granska information. Idag handlar det inte om bed\u00F6mning \u2014 den gjorde vi ig\u00E5r. Idag handlar det om att reflektera \u00F6ver vad ni l\u00E4rt er och vad ni tar med er.\u201D"),

      heading3("Momentets resa (12\u201315 min)"),
      bodyText("Kort muntlig sammanfattning (3 min max). Koppla fram\u00E5t till L8 seminarie-examination:"),
      bullet("Lektion 1: \u00D6ppning med AI-labb och inokulering"),
      bullet("Lektion 2: K\u00E4llkritiska verktyg"),
      bullet("Lektion 3: AI och konspirationsteorier"),
      bullet("Lektion 4: Formativt seminarium"),
      bullet("Lektion 5: Skriftlig perspektivanalys"),
      bullet("Lektion 6: Summativ skrivuppgift"),
      bullet("Lektion 7 (idag): Metakognitiv reflektion och sj\u00E4lvbed\u00F6mning"),
      bullet("Lektion 8: Seminarie-examination"),

      heading3("Syntesreflektion (15\u201330 min)"),
      italicText("S\u00E4g: \u201CDet h\u00E4r \u00E4r ingen bed\u00F6mning. Det h\u00E4r handlar om er \u2014 vad tar ni med er? Jag vill att ni skriver fritt i 15 minuter. Ni kan ocks\u00E5 prata med en kompis om ni hellre vill det, men f\u00F6rs\u00F6k att skriva ner \u00E5tminstone n\u00E5gra meningar.\u201D"),
      spacer(),
      bodyText("Skriv fr\u00E5gorna p\u00E5 tavlan:"),
      bullet("\u201CVad tar du med dig fr\u00E5n det h\u00E4r momentet?\u201D"),
      bullet("\u201CHur t\u00E4nker du annorlunda nu j\u00E4mf\u00F6rt med f\u00F6re momentet?\u201D"),
      bullet("\u201CN\u00E4r tror du att du kommer anv\u00E4nda det du l\u00E4rt dig?\u201D"),
      spacer(),
      bodyText("Cirkulera tyst. Det h\u00E4r \u00E4r inget som bed\u00F6ms \u2014 uppmuntra \u00E4rlighet och personliga reflektioner. P\u00E5minn om tredjepersons-framing n\u00E4r eleverna resonerar om konspirationsteorier \u2014 \u201Cforskare pekar p\u00E5\u201D ist\u00E4llet f\u00F6r \u201Cfolk tror p\u00E5\u201D. Om elever fastnar:"),
      bullet("\u201CT\u00E4nk p\u00E5 en specifik lektion \u2014 vad \u00F6verraskade dig?\u201D"),
      bullet("\u201CHar du m\u00E4rkt n\u00E5got i vardagen sedan vi b\u00F6rjade momentet?\u201D"),

      heading3("Par-samtal + helklass (30\u201340 min)"),
      italicText("S\u00E4g: \u201CBeskriv f\u00F6r din partner: det viktigaste du tar med dig. Vad har f\u00F6r\u00E4ndrats i hur du t\u00E4nker?\u201D"),
      spacer(),
      bodyText("Par-samtal (5 min). Cirkulera och lyssna."),
      spacer(),
      bodyText("Helklassdiskussion (5 min):"),
      italicText("\u201CVem vill dela med sig? Vad har ni l\u00E4rt er som ni inte visste f\u00F6re det h\u00E4r momentet?\u201D"),
      bodyText("Lyft 3\u20134 r\u00F6ster. Bekr\u00E4fta och koppla till l\u00E4randem\u00E5len."),
      spacer(),
      italicText("Fr\u00E5ga: \u201CSer ni n\u00E5gra m\u00F6nster i vad vi tar med oss? N\u00E5got som flera av er n\u00E4mner?\u201D"),
      bodyText("Sammanfatta p\u00E5 tavlan."),

      heading3("Sj\u00E4lvv\u00E4rdering (40\u201350 min)"),
      bodyText("Dela ut sj\u00E4lvv\u00E4rderingsmatrisen (se bilaga 2)."),
      spacer(),
      italicText("S\u00E4g: \u201CDet h\u00E4r \u00E4r ingen bed\u00F6mning \u2014 det \u00E4r er egen \u00E4rliga uppskattning av var ni ligger. L\u00E4s varje m\u00E5l och markera den niv\u00E5 som b\u00E4st beskriver vad ni kan just nu. Det finns inga r\u00E4tta svar h\u00E4r \u2014 bara \u00E4rlighet.\u201D"),
      spacer(),
      bodyText("L\u00E5t eleverna fylla i (7\u20138 minuter). Uppmuntra dem att skriva kommentarer \u2014 var har de utvecklats mest? Vad vill de f\u00F6rb\u00E4ttra?"),
      spacer(),
      bodyText("Samla in. Anv\u00E4nd tillsammans med den summativa bed\u00F6mningen fr\u00E5n lektion 6 f\u00F6r en helhetsbild av elevens l\u00E4rande."),

      heading3("Gemensam avslutning (50\u201358 min)"),
      italicText("\u201CL\u00E5t oss avsluta. Hela det h\u00E4r momentet handlade om en sak: att inte bli lurad. Inte av AI, inte av konspirationsteorier, inte av n\u00E5gon som vill att ni ska sluta t\u00E4nka kritiskt.\u201D"),
      spacer(),
      italicText("\u201CVad tar ni med er?\u201D"),
      spacer(),
      bodyText("Samla 4\u20135 r\u00F6ster fr\u00E5n klassen. Sammanfatta p\u00E5 tavlan. Knyt ihop med l\u00E4randem\u00E5len:"),
      bullet("\u201CNi n\u00E4mner grundfr\u00E5gorna \u2014 det \u00E4r m\u00E5l 1."),
      bullet("\u201CNi pratar om AI och deepfakes \u2014 det \u00E4r m\u00E5l 2."),
      bullet("\u201CNi f\u00F6rst\u00E5r varf\u00F6r konspirationsteorier sprids \u2014 det \u00E4r m\u00E5l 3."),
      spacer(),
      italicText("\u201CSe vad ni kan nu. Det \u00E4r imponerande.\u201D"),

      heading3("Fram\u00E5tblick (58\u201365 min)"),
      italicText("S\u00E4g: \u201CNi har nu en verktygsl\u00E5da som de flesta vuxna inte har. Ni vet hur man st\u00E4ller grundfr\u00E5gorna. Ni vet hur AI-genererat inneh\u00E5ll fungerar. Ni vet varf\u00F6r konspirationsteorier sprids och hur man genomsk\u00E5dar dem.\u201D"),
      spacer(),
      italicText("\u201CDet h\u00E4r slutar inte h\u00E4r. Varje g\u00E5ng ni \u00F6ppnar sociala medier, l\u00E4ser en nyhet eller h\u00F6r n\u00E5got som l\u00E5ter f\u00F6r bra \u2014 eller f\u00F6r skr\u00E4mmande \u2014 f\u00F6r att vara sant: st\u00E4ll fr\u00E5gorna. Vem? Varf\u00F6r? Hur? N\u00E4r?\u201D"),
      spacer(),
      italicText("\u201CI det avslutande arbetet kommer ni att beh\u00F6va allt det h\u00E4r \u2014 k\u00E4llkritik, k\u00E4llh\u00E4nvisning, f\u00F6rm\u00E5gan att granska och v\u00E4rdera. Ni \u00E4r redo.\u201D"),
      spacer(),
      italicText("\u201CVar den personen i rummet som st\u00E4ller fr\u00E5gorna som andra inte g\u00F6r. Det \u00E4r vad k\u00E4llkritik handlar om. Tack f\u00F6r ett riktigt bra moment.\u201D"),

      // ELEVAKTIVITETER
      heading2("Elevaktiviteter"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "R\u00F6stning: \u00E4kta eller fejk? \u2014 \u00E5terblick fr\u00E5n lektion 1 (helklass, 7 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Muntlig diskussion: vilka verktyg har vi nu? (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Syntesreflektion: fritt skrivande om l\u00E4rande (enskilt, 15 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Par-samtal: det viktigaste jag tar med mig (par, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Helklassdiskussion: m\u00F6nster och \u00F6verraskningar (helklass, 5 min)", font: "Arial", size: 24 })],
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: "Sj\u00E4lvv\u00E4rderingsmatris (enskilt, 8 min)", font: "Arial", size: 24 })],
      }),
      spacer(),
      boldBodyText("Elevaktiv tid: ", "ca 52 av 65 minuter (80%)"),

      // DIFFERENTIERING
      heading2("Differentiering"),
      boldBodyText("St\u00F6d (mot E): ", "Syntesreflektionen har st\u00F6dfr\u00E5gor p\u00E5 tavlan. Eleverna kan v\u00E4lja att prata med en kompis ist\u00E4llet f\u00F6r att skriva. Sj\u00E4lvv\u00E4rderingsmatrisen har konkreta beskrivningar per niv\u00E5. L\u00E4raren kan hj\u00E4lpa med \u201CT\u00E4nk p\u00E5 en specifik lektion \u2014 vad \u00F6verraskade dig?\u201D"),
      boldBodyText("Utmaning (mot A): ", "Uppmuntra djupare reflektion: \u201CResonera om vad detta inneb\u00E4r f\u00F6r demokratin och samh\u00E4llsdebatten.\u201D Sj\u00E4lvv\u00E4rderingsmatrisen har en kolumn \u201CBevis/exempel\u201D d\u00E4r eleven motiverar sin bed\u00F6mning. Syntesreflektionen kan koppla till egna erfarenheter utanf\u00F6r skolan."),

      // MATERIAL
      heading2("Material"),
      bullet("\u00C5terblicksbildspel: samma tre bilder fr\u00E5n lektion 1 (\u00E4kta nyhetsbild, p\u00E5ven i dunjacka, Pentagon)"),
      bullet("Reflektionsformul\u00E4r \u2014 en per elev (se bilaga 1)"),
      bullet("Sj\u00E4lvv\u00E4rderingsmatris \u2014 en per elev (se bilaga 2)"),
      bullet("L\u00E4randem\u00E5len utskrivna/projicerade"),

      // KOPPLING TILL KUNSKAPSKRAV
      heading2("Koppling till kunskapskrav"),
      bullet("\u00C5terblicken visar utveckling fr\u00E5n lektion 1 \u2014 eleverna kan nu motivera sina bed\u00F6mningar med verktyg (m\u00E5l 1, 2)"),
      bullet("Syntesreflektionen st\u00E4rker metakognition och transferf\u00F6rm\u00E5ga (m\u00E5l 1\u20133)"),
      bullet("Sj\u00E4lvv\u00E4rderingen ger l\u00E4raren kompletterande information om elevens egen upplevelse av m\u00E5luppfyllelse"),
      spacer(),
      bodyText("Den summativa bed\u00F6mningen gjordes i lektion 6 (skriftlig analys med k\u00E4llh\u00E4nvisning). Den h\u00E4r lektionen ger kompletterande information genom sj\u00E4lvv\u00E4rdering och fri reflektion, men \u00E4r inte bed\u00F6mningsgrundande."),

      // KOPPLINGAR
      heading2("Kopplingar"),
      bullet("Bygger p\u00E5: Samtliga tidigare lektioner (L1\u2013L6)"),
      bullet("Kopplar tillbaka till L1: samma bilder, samma grundfr\u00E5ga \u2014 visar progression"),
      bullet("Summativ bed\u00F6mning gjordes i L6 (summativ skrivuppgift) \u2014 den h\u00E4r lektionen \u00E4r reflekterande"),
      bullet("Sj\u00E4lvv\u00E4rderingen kan anv\u00E4ndas f\u00F6r utvecklingssamtal"),
      bullet("Fram\u00E5tblicken kopplar till L8 seminarie-examination"),

      // SIDBRYTNING FÖR BILAGOR
      new Paragraph({ children: [new PageBreak()] }),

      // BILAGA 1: REFLEKTIONSFORMULÄR (tidigare bilaga 2)
      heading2("Bilaga 1: Reflektionsformul\u00E4r \u2014 Momentets avslutning"),
      spacer(),
      bodyText("Fyll i enskilt. Det finns inga r\u00E4tta svar \u2014 det h\u00E4r handlar om din upplevelse."),
      spacer(),

      boldBodyText("1. ", "Vad \u00E4r det viktigaste du l\u00E4rt dig under det h\u00E4r momentet?"),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      spacer(),

      boldBodyText("2. ", "Beskriv en situation d\u00E4r du redan har anv\u00E4nt (eller kunde ha anv\u00E4nt) det du l\u00E4rt dig \u2014 utanf\u00F6r klassrummet."),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      spacer(),

      boldBodyText("3. ", "Vad \u00F6verraskade dig mest under momentet?"),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      spacer(),

      boldBodyText("4. ", "Om du fick ge ETT r\u00E5d till n\u00E5gon som inte vet n\u00E5got om k\u00E4llkritik, AI-inneh\u00E5ll och konspirationsteorier \u2014 vad skulle det vara?"),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      spacer(),

      boldBodyText("5. ", "Vilken fr\u00E5ga vill du forts\u00E4tta utforska efter det h\u00E4r momentet?"),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),

      // BILAGA 2: SJÄLVVÄRDERINGSMATRIS (tidigare bilaga 3)
      new Paragraph({ children: [new PageBreak()] }),
      heading2("Bilaga 2: Sj\u00E4lvv\u00E4rderingsmatris"),
      spacer(),
      bodyText("Markera den niv\u00E5 som b\u00E4st beskriver var du befinner dig just nu. Var \u00E4rlig \u2014 det h\u00E4r \u00E4r ingen bed\u00F6mning, utan en hj\u00E4lp f\u00F6r dig sj\u00E4lv."),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 2408, 2410, 2408],
        rows: [
          new TableRow({
            children: [
              headerCell("L\u00E4randem\u00E5l", 1800),
              headerCell("E \u2014 Jag kan grunderna", 2408),
              headerCell("C \u2014 Jag kan f\u00F6rklara och till\u00E4mpa", 2410),
              headerCell("A \u2014 Jag kan analysera p\u00E5 djupet", 2408),
            ],
          }),
          new TableRow({
            children: [
              cell("1. Granska och v\u00E4rdera k\u00E4llor", 1800),
              cell("Jag kan namnge grundfr\u00E5gorna och anv\u00E4nda dem p\u00E5 enkla k\u00E4llor.", 2408),
              cell("Jag kan anv\u00E4nda grundfr\u00E5gorna p\u00E5 olika typer av k\u00E4llor och f\u00F6rklara min bed\u00F6mning.", 2410),
              cell("Jag kan granska komplexa k\u00E4llor, se nyanser och f\u00F6rklara varf\u00F6r en k\u00E4lla kan vara b\u00E5de trov\u00E4rdig och vilseledande.", 2408),
            ],
          }),
          new TableRow({
            children: [
              cell("2. AI-genererat inneh\u00E5ll", 1800),
              cell("Jag vet att AI kan skapa inneh\u00E5ll och att det \u00E4r sv\u00E5rt att avsl\u00F6ja.", 2408),
              cell("Jag kan f\u00F6rklara hur AI-inneh\u00E5ll p\u00E5verkar informationslandskapet och ge konkreta exempel.", 2410),
              cell("Jag kan analysera AI:s p\u00E5verkan fr\u00E5n flera perspektiv, inklusive demokrati och samh\u00E4llsdebatt.", 2408),
            ],
          }),
          new TableRow({
            children: [
              cell("3. Konspirationsteorier", 1800),
              cell("Jag kan f\u00F6rklara vad en konspirationsteori \u00E4r och n\u00E4mna n\u00E5gra orsaker till att de sprids.", 2408),
              cell("Jag kan f\u00F6rklara spridningsmekanismer och koppla till k\u00E4llkritik.", 2410),
              cell("Jag kan analysera konspirationsteorier fr\u00E5n flera perspektiv och resonera om samh\u00E4llskonsekvenser.", 2408),
            ],
          }),
        ],
      }),

      spacer(),
      boldBodyText("Min markering: ", "Ringa in eller markera den ruta som st\u00E4mmer b\u00E4st f\u00F6r varje rad."),
      spacer(),
      boldBodyText("Kommentar (valfritt): ", "Finns det n\u00E5got m\u00E5l d\u00E4r du k\u00E4nner att du utvecklats extra mycket? N\u00E5got du vill f\u00F6rb\u00E4ttra?"),
      spacer(),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
      bodyText("__________________________________________________________"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/anders/Second brain/Undervisningsmaterial/Samh\u00E4llskunskap/K\u00E4llkritik AI och konspirationsteorier/lektion-7.docx", buffer);
  console.log("lektion-7.docx skapad!");
});
